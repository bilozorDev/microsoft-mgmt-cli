import { resolve, dirname, join } from "path";
import { mkdirSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generateReport } from "../report-template.ts";
import { appDir, escapePS } from "../utils.ts";

interface ForwardingMailbox {
  DisplayName: string;
  UserPrincipalName: string;
  ForwardingSmtpAddress: string | null;
  ForwardingAddress: string | null;
  DeliverToMailboxAndForward: boolean;
}

interface InboxRule {
  Name: string;
  ForwardTo: string[] | string | null;
  ForwardAsAttachmentTo: string[] | string | null;
  RedirectTo: string[] | string | null;
  Enabled: boolean;
}

interface UserMailbox {
  UserPrincipalName: string;
  DisplayName: string;
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

function forwardTarget(m: ForwardingMailbox): string {
  return (m.ForwardingSmtpAddress ?? m.ForwardingAddress ?? "").replace(/^smtp:/i, "");
}

/**
 * Parse raw inbox rule target string like:
 *   "Andrew Isnetto" [EX:/o=ExchangeLabs/.../cn=abc-aisnetto]
 *   "someone@ext.com" [SMTP:someone@ext.com]
 * Returns { displayName, address } where address is the bracket contents.
 */
function parseRuleTarget(raw: string): { displayName: string; address: string } {
  const smtpMatch = raw.match(/\[SMTP:(.+?)\]/i);
  if (smtpMatch) return { displayName: "", address: smtpMatch[1]! };
  const exMatch = raw.match(/\[EX:(.+?)\]/i);
  if (exMatch) return { displayName: raw.replace(/\s*\[.*\]/, "").replace(/^"|"$/g, ""), address: exMatch[1]! };
  return { displayName: "", address: raw };
}

/** Resolve EX: distinguished names to SMTP addresses via Get-Recipient, with caching. */
async function resolveTargets(
  ps: PowerShellSession,
  ruleRows: { mailbox: UserMailbox; rule: InboxRule }[],
): Promise<Map<string, string>> {
  // Collect all unique EX: DNs
  const exDNs = new Set<string>();
  for (const { rule } of ruleRows) {
    for (const field of [rule.ForwardTo, rule.ForwardAsAttachmentTo, rule.RedirectTo]) {
      if (!field) continue;
      const arr = Array.isArray(field) ? field : [field];
      for (const t of arr) {
        const parsed = parseRuleTarget(t);
        if (/\[EX:/i.test(t)) exDNs.add(parsed.address);
      }
    }
  }

  // Resolve each unique DN → PrimarySmtpAddress
  const cache = new Map<string, string>();
  for (const dn of exDNs) {
    try {
      const { output } = await ps.runCommand(
        `Get-Recipient -Identity '${escapePS(dn)}' | Select-Object -ExpandProperty PrimarySmtpAddress`,
      );
      if (output?.trim()) cache.set(dn, output.trim());
    } catch {
      // Leave unresolved — will fall back to display name
    }
  }
  return cache;
}

/** Collect all raw target strings from an inbox rule. */
function collectRawTargets(rule: InboxRule): string[] {
  const targets: string[] = [];
  for (const field of [rule.ForwardTo, rule.ForwardAsAttachmentTo, rule.RedirectTo]) {
    if (!field) continue;
    const arr = Array.isArray(field) ? field : [field];
    targets.push(...arr);
  }
  return targets;
}

/** Format a list of raw rule targets into clean email addresses. */
function formatRuleTargets(rawTargets: string[], dnCache: Map<string, string>, sep = "; "): string {
  return rawTargets.map((t) => {
    const parsed = parseRuleTarget(t);
    if (/\[SMTP:/i.test(t)) return parsed.address;
    if (/\[EX:/i.test(t)) return dnCache.get(parsed.address) ?? parsed.displayName;
    return t;
  }).join(sep);
}

export async function run(ps: PowerShellSession): Promise<void> {
  const spin = p.spinner();

  // Step 1: Fetch mailbox-level forwarding
  spin.start("Fetching mailbox forwarding configurations…");
  const raw = await ps.runCommandJson<ForwardingMailbox | ForwardingMailbox[]>(
    `Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Where-Object { $_.ForwardingSmtpAddress -ne $null -or $_.ForwardingAddress -ne $null } | Select-Object DisplayName,UserPrincipalName,ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward`,
  );
  const forwarding = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
  spin.stop(`Found ${forwarding.length} forwarding rule(s) set up by admin.`);

  // Step 2: Fetch all mailboxes and scan inbox rules
  spin.start("Fetching all user mailboxes…");
  const allRaw = await ps.runCommandJson<UserMailbox | UserMailbox[]>(
    `Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object UserPrincipalName,DisplayName`,
  );
  const allMailboxes = allRaw ? (Array.isArray(allRaw) ? allRaw : [allRaw]) : [];
  spin.stop(`Found ${allMailboxes.length} user mailbox(es).`);

  spin.start(`Scanning inbox rules (1/${allMailboxes.length})…`);
  const inboxRuleRows: { mailbox: UserMailbox; rule: InboxRule }[] = [];
  for (let i = 0; i < allMailboxes.length; i++) {
    spin.message(`Scanning inbox rules (${i + 1}/${allMailboxes.length})…`);
    const mb = allMailboxes[i]!;
    const escaped = escapePS(mb.UserPrincipalName);

    let rules: InboxRule[] = [];
    try {
      const rulesRaw = await ps.runCommandJson<InboxRule | InboxRule[]>(
        `Get-InboxRule -Mailbox '${escaped}' | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo } | Select-Object Name,ForwardTo,ForwardAsAttachmentTo,RedirectTo,Enabled`,
      );
      rules = rulesRaw ? (Array.isArray(rulesRaw) ? rulesRaw : [rulesRaw]) : [];
    } catch {
      // Skip mailboxes where Get-InboxRule fails (corrupted rules, permissions, etc.)
    }

    for (const rule of rules) {
      inboxRuleRows.push({ mailbox: mb, rule });
    }
  }
  spin.stop(`Scan complete. Found ${inboxRuleRows.length} inbox rule(s) with forwarding.`);

  // Resolve EX: distinguished names to SMTP addresses
  let dnCache = new Map<string, string>();
  if (inboxRuleRows.length > 0) {
    spin.start("Resolving recipient addresses…");
    dnCache = await resolveTargets(ps, inboxRuleRows);
    spin.stop("Resolved recipient addresses.");
  }

  if (forwarding.length === 0 && inboxRuleRows.length === 0) {
    p.log.info("No forwarding found (mailbox-level or inbox rules).");
    return;
  }

  // Display terminal table
  const tableLines: string[] = [];

  if (forwarding.length > 0) {
    const header = `${"User".padEnd(25)} ${"Email".padEnd(35)} ${"Forwards To".padEnd(35)} Keep Copy`;
    const separator = "─".repeat(header.length + 5);
    const displayRows = forwarding.slice(0, 50);
    tableLines.push(header, separator);
    for (const m of displayRows) {
      const name = truncate(m.DisplayName ?? "", 24).padEnd(25);
      const email = truncate(m.UserPrincipalName ?? "", 34).padEnd(35);
      const target = truncate(forwardTarget(m), 34).padEnd(35);
      const keepCopy = m.DeliverToMailboxAndForward ? "Yes" : "No";
      tableLines.push(`${name} ${email} ${target} ${keepCopy}`);
    }
    if (forwarding.length > 50) {
      tableLines.push(`… and ${forwarding.length - 50} more`);
    }
  }

  if (inboxRuleRows.length > 0) {
    if (tableLines.length > 0) tableLines.push("");
    const ruleHeader = `${"User".padEnd(25)} ${"Rule Name".padEnd(25)} ${"Forwards To".padEnd(35)} Enabled`;
    const ruleSep = "─".repeat(ruleHeader.length + 5);
    const displayRules = inboxRuleRows.slice(0, 50);
    tableLines.push(ruleHeader, ruleSep);
    for (const { mailbox, rule } of displayRules) {
      const name = truncate(mailbox.DisplayName ?? "", 24).padEnd(25);
      const ruleName = truncate(rule.Name ?? "", 24).padEnd(25);
      const rawTargets = collectRawTargets(rule);
      const target = truncate(formatRuleTargets(rawTargets, dnCache), 34).padEnd(35);
      const enabled = rule.Enabled ? "Yes" : "No";
      tableLines.push(`${name} ${ruleName} ${target} ${enabled}`);
    }
    if (inboxRuleRows.length > 50) {
      tableLines.push(`… and ${inboxRuleRows.length - 50} more`);
    }
  }

  const total = forwarding.length + inboxRuleRows.length;
  p.note(tableLines.join("\n"), `Forwarding audit (${total} result${total === 1 ? "" : "s"})`);

  // Excel export
  const exportXlsx = await p.confirm({
    message: "Export to Excel?",
    initialValue: false,
  });
  if (p.isCancel(exportXlsx)) return;

  if (exportXlsx) {
    const excelRows: string[][] = [];

    for (const m of forwarding) {
      const target = forwardTarget(m);
      const type = m.ForwardingSmtpAddress ? "SMTP Forwarding" : "Internal Forwarding";
      excelRows.push([
        m.DisplayName ?? "",
        m.UserPrincipalName,
        type,
        target,
        m.DeliverToMailboxAndForward ? "Yes" : "No",
        "",
        "",
      ]);
    }

    for (const { mailbox, rule } of inboxRuleRows) {
      const rawTargets = collectRawTargets(rule);
      excelRows.push([
        mailbox.DisplayName ?? "",
        mailbox.UserPrincipalName,
        "Inbox Rule",
        formatRuleTargets(rawTargets, dnCache, "\n"),
        "",
        rule.Name ?? "",
        rule.Enabled ? "Yes" : "No",
      ]);
    }

    spin.start("Generating Excel report…");

    const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
    const dateSlug = new Date().toISOString().slice(0, 10);
    const outputDir = join(appDir(), "reports output");
    const fullPath = resolve(join(outputDir, `${tenantSlug}-forwarding-audit-${dateSlug}.xlsx`));
    mkdirSync(dirname(fullPath), { recursive: true });

    const buffer = await generateReport({
      sheetName: "Forwarding Audit",
      title: "Mailbox Forwarding Audit",
      tenant: ps.tenantDomain ?? "Unknown",
      summary: `${forwarding.length} mailbox-level forwarding, ${inboxRuleRows.length} inbox rule(s)`,
      columns: [
        { header: "User", width: 25 },
        { header: "Email", width: 35 },
        { header: "Type", width: 20 },
        { header: "Forwards To", width: 40, wrapText: true },
        { header: "Keep Copy", width: 12 },
        { header: "Rule Name", width: 25 },
        { header: "Rule Enabled", width: 14 },
      ],
      rows: excelRows,
    });

    await Bun.write(fullPath, buffer);
    spin.stop(`Exported ${excelRows.length} rows to ${fullPath}`);

    const folder = dirname(fullPath);
    try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
  }
}
