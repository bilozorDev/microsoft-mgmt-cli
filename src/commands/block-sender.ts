import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { escapePS, handleOrgCustomizationError } from "../utils.ts";

function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isValidDomain(domain: string): boolean {
  return /^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)*\.[a-zA-Z]{2,}$/.test(domain);
}

async function fetchBlockedLists(ps: PowerShellSession): Promise<{ domains: string[]; senders: string[] }> {
  const { output: domainOut } = await ps.runCommand(
    '(Get-HostedContentFilterPolicy -Identity "Default").BlockedSenderDomains | ForEach-Object { $_.Domain }'
  );
  const { output: senderOut } = await ps.runCommand(
    '(Get-HostedContentFilterPolicy -Identity "Default").BlockedSenders | ForEach-Object { $_.Sender }'
  );
  const domains = (domainOut || "").split("\n").map((l) => l.trim().toLowerCase()).filter(Boolean);
  const senders = (senderOut || "").split("\n").map((l) => l.trim().toLowerCase()).filter(Boolean);
  return { domains, senders };
}

async function viewBlocked(ps: PowerShellSession): Promise<void> {
  const spin = p.spinner();
  spin.start("Fetching blocked senders and domains...");
  const { domains, senders } = await fetchBlockedLists(ps);
  spin.stop("Fetched blocked lists.");

  if (domains.length === 0 && senders.length === 0) {
    p.log.info("No blocked senders or domains configured.");
    return;
  }

  const lines: string[] = [];
  if (domains.length > 0) {
    lines.push("Blocked Domains:");
    for (const d of domains) lines.push(`  ${d}`);
  }
  if (senders.length > 0) {
    if (lines.length > 0) lines.push("");
    lines.push("Blocked Senders:");
    for (const s of senders) lines.push(`  ${s}`);
  }
  p.note(lines.join("\n"), `Blocked list (${domains.length} domain(s), ${senders.length} sender(s))`);
}

async function addBlocked(ps: PowerShellSession): Promise<void> {
  const type = await p.select({
    message: "What would you like to block?",
    options: [
      { value: "domains", label: "Domain(s)", hint: "e.g. spammer.com" },
      { value: "senders", label: "Sender email(s)", hint: "e.g. bad@spammer.com" },
    ],
  });
  if (p.isCancel(type)) return;

  const isDomain = type === "domains";
  const input = await p.text({
    message: isDomain
      ? "Enter domain(s) to block (comma-separated)"
      : "Enter sender email(s) to block (comma-separated)",
    placeholder: isDomain ? "spammer.com, evil.org" : "bad@spammer.com, junk@evil.org",
    validate(value = "") {
      if (!value.trim()) return "Please enter at least one value";
      const items = value.split(",").map((v) => v.trim()).filter(Boolean);
      if (isDomain) {
        const invalid = items.filter((d) => !isValidDomain(d));
        if (invalid.length > 0) return `Invalid domain(s): ${invalid.join(", ")}`;
      } else {
        const invalid = items.filter((e) => !isValidEmail(e));
        if (invalid.length > 0) return `Invalid email(s): ${invalid.join(", ")}`;
      }
    },
  });
  if (p.isCancel(input)) return;

  const items = input.split(",").map((v) => v.trim()).filter(Boolean);

  // Duplicate detection
  const spin = p.spinner();
  spin.start("Checking existing blocked list...");
  const { domains: existingDomains, senders: existingSenders } = await fetchBlockedLists(ps);
  spin.stop("Checked existing blocked list.");

  const existing = new Set(isDomain ? existingDomains : existingSenders);
  const alreadyPresent = items.filter((v) => existing.has(v.toLowerCase()));
  const toAdd = items.filter((v) => !existing.has(v.toLowerCase()));

  for (const v of alreadyPresent) {
    p.log.info(`${v} is already in blocked list`);
  }

  if (toAdd.length === 0) return;

  p.note(toAdd.map((v) => `  ${v}`).join("\n"), isDomain ? "Domains to block" : "Senders to block");

  const confirm = await p.confirm({
    message: `Block ${toAdd.length} ${isDomain ? "domain(s)" : "sender(s)"}?`,
  });
  if (p.isCancel(confirm) || !confirm) {
    p.log.info("Cancelled.");
    return;
  }

  spin.start(`Blocking ${isDomain ? "domain(s)" : "sender(s)"}...`);

  const param = isDomain ? "BlockedSenderDomains" : "BlockedSenders";
  const results: { item: string; success: boolean; error?: string }[] = [];
  let orgCustomizationHandled = false;

  for (const item of toAdd) {
    const cmd = `Set-HostedContentFilterPolicy -Identity 'Default' -${param} @{Add='${escapePS(item)}'}`;
    const { error } = await ps.runCommand(cmd);

    if (error && !orgCustomizationHandled && error.includes("Enable-OrganizationCustomization")) {
      orgCustomizationHandled = true;
      spin.stop("Organization customization required.");

      const fixed = await handleOrgCustomizationError(ps, error);
      if (fixed) {
        spin.start(`Blocking ${isDomain ? "domain(s)" : "sender(s)"}...`);
        let retry = await ps.runCommand(cmd);
        for (let attempt = 1; retry.error && attempt < 3; attempt++) {
          await Bun.sleep(10_000);
          retry = await ps.runCommand(cmd);
        }
        results.push({ item, success: !retry.error, error: retry.error || undefined });
      } else {
        const msg = "Organization customization required";
        results.push({ item, success: false, error: msg });
        for (const remaining of toAdd.slice(toAdd.indexOf(item) + 1)) {
          results.push({ item: remaining, success: false, error: msg });
        }
        break;
      }
    } else {
      results.push({ item, success: !error, error: error || undefined });
    }
  }

  spin.stop("Done.");

  // Verify
  spin.start("Verifying blocked list...");
  const { domains: newDomains, senders: newSenders } = await fetchBlockedLists(ps);
  spin.stop("Verification complete.");

  for (const r of results) {
    if (r.success) {
      p.log.success(`${r.item} — blocked`);
    } else {
      p.log.error(`${r.item} — failed: ${r.error}`);
    }
  }

  const allBlocked = isDomain ? newDomains : newSenders;
  if (allBlocked.length > 0) {
    p.note(
      allBlocked.map((v) => `  ${v}`).join("\n"),
      isDomain ? "Current blocked domains" : "Current blocked senders"
    );
  }
}

async function removeBlocked(ps: PowerShellSession): Promise<void> {
  const spin = p.spinner();
  spin.start("Fetching blocked lists...");
  const { domains, senders } = await fetchBlockedLists(ps);
  spin.stop("Fetched blocked lists.");

  if (domains.length === 0 && senders.length === 0) {
    p.log.info("No blocked senders or domains to remove.");
    return;
  }

  // Build combined options
  const options: { value: string; label: string; hint: string }[] = [];
  for (const d of domains) options.push({ value: `domain:${d}`, label: d, hint: "domain" });
  for (const s of senders) options.push({ value: `sender:${s}`, label: s, hint: "sender" });

  const selected = await p.multiselect({
    message: "Select entries to unblock",
    options,
    required: true,
  });
  if (p.isCancel(selected)) return;

  const confirm = await p.confirm({
    message: `Remove ${selected.length} entry/entries from blocked list?`,
  });
  if (p.isCancel(confirm) || !confirm) {
    p.log.info("Cancelled.");
    return;
  }

  spin.start("Removing entries...");

  for (const entry of selected) {
    const [type, value] = entry.split(":", 2) as [string, string];
    const param = type === "domain" ? "BlockedSenderDomains" : "BlockedSenders";
    const cmd = `Set-HostedContentFilterPolicy -Identity 'Default' -${param} @{Remove='${escapePS(value)}'}`;
    const { error } = await ps.runCommand(cmd);
    if (error) {
      p.log.error(`Failed to remove ${value}: ${error}`);
    } else {
      p.log.success(`${value} — removed`);
    }
  }

  spin.stop("Done removing entries.");
}

export async function run(ps: PowerShellSession): Promise<void> {
  while (true) {
    const action = await p.select({
      message: "Block Sender/Domain",
      options: [
        { value: "view", label: "View blocked senders & domains" },
        { value: "add", label: "Add to blocked list" },
        { value: "remove", label: "Remove from blocked list" },
        { value: "back", label: "Back" },
      ],
    });

    if (p.isCancel(action) || action === "back") return;

    switch (action) {
      case "view":
        await viewBlocked(ps);
        break;
      case "add":
        await addBlocked(ps);
        break;
      case "remove":
        await removeBlocked(ps);
        break;
    }
  }
}
