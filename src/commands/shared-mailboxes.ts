import { resolve, dirname, join } from "path";
import { mkdirSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";
import { generateReport } from "../report-template.ts";
import { appDir } from "../utils.ts";

interface SharedMailbox {
  DisplayName: string;
  PrimarySmtpAddress: string;
  Alias: string;
  WhenCreated: string;
}

interface SubscribedSku {
  SkuId: string;
  SkuPartNumber: string;
}

interface GraphUser {
  UserPrincipalName: string;
  LicenseSkuIds: string[];
}

interface MailboxPermission {
  User: string;
  AccessRights: string[] | string;
}

interface RecipientPermission {
  Trustee: string;
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

export async function run(ps: PowerShellSession): Promise<void> {
  const spin = p.spinner();
  spin.start("Fetching shared mailboxes…");

  const raw = await ps.runCommandJson<SharedMailbox | SharedMailbox[]>(
    `Get-Mailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName,PrimarySmtpAddress,Alias,WhenCreated`,
  );

  const mailboxes = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
  spin.stop(`Found ${mailboxes.length} shared mailbox(es).`);

  if (mailboxes.length === 0) {
    p.log.info("No shared mailboxes found.");
    return;
  }

  // Check for licenses via Graph
  spin.start("Checking for unnecessary licenses…");
  let licenseMap = new Map<string, string[]>(); // UPN → friendly license names
  try {
    await ps.ensureGraphConnected();

    // Build filter for shared mailbox UPNs
    const upns = mailboxes.map((m) => m.PrimarySmtpAddress);
    const filterParts = upns.map((u) => `UserPrincipalName eq '${u.replace(/'/g, "''")}'`);
    // Graph $filter has a URL length limit — batch in groups of 15
    const chunks: string[][] = [];
    for (let i = 0; i < filterParts.length; i += 15) {
      chunks.push(filterParts.slice(i, i + 15));
    }

    const graphUsers: GraphUser[] = [];
    for (const chunk of chunks) {
      const filter = chunk.join(" or ");
      const chunkRaw = await ps.runCommandJson<GraphUser | GraphUser[]>(
        `Get-MgUser -Filter "${filter}" -Property UserPrincipalName,AssignedLicenses | ForEach-Object { [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName; LicenseSkuIds = @($_.AssignedLicenses.SkuId) } }`,
      );
      if (chunkRaw) {
        const arr = Array.isArray(chunkRaw) ? chunkRaw : [chunkRaw];
        graphUsers.push(...arr);
      }
    }

    // Fetch SKU names for friendly display
    const skuRaw = await ps.runCommandJson<SubscribedSku | SubscribedSku[]>(
      `Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber`,
    );
    const skuList = skuRaw ? (Array.isArray(skuRaw) ? skuRaw : [skuRaw]) : [];
    const skuMap = new Map(skuList.map((s) => [s.SkuId, s.SkuPartNumber]));

    for (const gu of graphUsers) {
      const ids = (gu.LicenseSkuIds ?? []).filter(Boolean);
      if (ids.length > 0) {
        const names = ids.map((id) => {
          const part = skuMap.get(id);
          return part ? friendlySkuName(part) : id;
        });
        licenseMap.set(gu.UserPrincipalName.toLowerCase(), names);
      }
    }

    const licensedCount = licenseMap.size;
    spin.stop(
      licensedCount > 0
        ? `Warning: ${licensedCount} shared mailbox(es) have licenses assigned.`
        : "No unnecessary licenses found.",
    );

    if (licensedCount > 0) {
      p.log.warn(
        "Shared mailboxes don't need licenses. Removing them can save costs.",
      );
    }
  } catch {
    spin.stop("Could not check licenses (Graph connection failed).");
  }

  // Display table
  const header = `${"Name".padEnd(30)} ${"Email".padEnd(40)} ${"Created".padEnd(12)} Licenses`;
  const separator = "─".repeat(header.length + 20);
  const displayRows = mailboxes.slice(0, 50);
  const rows = displayRows.map((m) => {
    const name = truncate(m.DisplayName ?? "", 29).padEnd(30);
    const email = truncate(m.PrimarySmtpAddress ?? "", 39).padEnd(40);
    const created = m.WhenCreated
      ? new Date(m.WhenCreated).toISOString().slice(0, 10)
      : "";
    const licenses = licenseMap.get(m.PrimarySmtpAddress.toLowerCase());
    const licStr = licenses ? `⚠ ${licenses.join(", ")}` : "—";
    return `${name} ${email} ${created.padEnd(12)} ${licStr}`;
  });

  const lines = [header, separator, ...rows];
  if (mailboxes.length > 50) {
    lines.push(`… and ${mailboxes.length - 50} more`);
  }
  p.note(lines.join("\n"), `Shared mailboxes (${mailboxes.length})`);

  // Excel export
  const exportXlsx = await p.confirm({
    message: "Export full results to Excel?",
    initialValue: false,
  });
  if (p.isCancel(exportXlsx)) return;

  if (exportXlsx) {
    const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
    const dateSlug = new Date().toISOString().slice(0, 10);
    const outputDir = join(appDir(), "reports output");
    const defaultName = join(outputDir, `${tenantSlug}-shared-mailboxes-${dateSlug}.xlsx`);

    const xlsxPath = await p.text({
      message: "File path",
      placeholder: defaultName,
      defaultValue: defaultName,
    });
    if (p.isCancel(xlsxPath)) return;

    const fullPath = resolve((xlsxPath as string).trim());
    mkdirSync(dirname(fullPath), { recursive: true });

    spin.start("Generating Excel report…");

    const buffer = await generateReport({
      sheetName: "Shared Mailboxes",
      title: "Shared Mailboxes Report",
      tenant: ps.tenantDomain ?? "Unknown",
      summary: `${mailboxes.length} shared mailbox(es) · ${licenseMap.size} with licenses`,
      columns: [
        { header: "Display Name", width: 30 },
        { header: "Email", width: 38 },
        { header: "Alias", width: 20 },
        { header: "Created", width: 16 },
        { header: "Licenses", width: 40 },
      ],
      rows: mailboxes.map((m) => {
        const licenses = licenseMap.get(m.PrimarySmtpAddress.toLowerCase());
        return [
          m.DisplayName ?? "",
          m.PrimarySmtpAddress,
          m.Alias,
          m.WhenCreated ? new Date(m.WhenCreated).toISOString().slice(0, 10) : "",
          licenses ? licenses.join("; ") : "",
        ];
      }),
    });

    await Bun.write(fullPath, buffer);
    spin.stop(`Exported ${mailboxes.length} rows to ${fullPath}`);

    const folder = dirname(fullPath);
    try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
  }

  // Permission drill-down
  const viewPerms = await p.confirm({
    message: "View permissions for a mailbox?",
    initialValue: false,
  });
  if (p.isCancel(viewPerms) || !viewPerms) return;

  while (true) {
    const options = mailboxes.map((m) => ({
      value: m.PrimarySmtpAddress,
      label: m.DisplayName,
      hint: m.PrimarySmtpAddress,
    }));
    options.push({ value: "back", label: "Back", hint: "" });

    const selected = await p.select({
      message: "Select a mailbox",
      options,
    });
    if (p.isCancel(selected) || selected === "back") break;

    const email = selected as string;
    const escaped = email.replace(/'/g, "''");

    spin.start(`Fetching permissions for ${email}…`);

    // FullAccess permissions
    const fullAccessRaw = await ps.runCommandJson<
      MailboxPermission | MailboxPermission[]
    >(
      `Get-MailboxPermission -Identity '${escaped}' | Where-Object { $_.User -ne 'NT AUTHORITY\\SELF' -and $_.IsInherited -eq $false } | Select-Object User,AccessRights`,
    );
    const fullAccess = fullAccessRaw
      ? Array.isArray(fullAccessRaw) ? fullAccessRaw : [fullAccessRaw]
      : [];

    // SendAs permissions
    const sendAsRaw = await ps.runCommandJson<
      RecipientPermission | RecipientPermission[]
    >(
      `Get-RecipientPermission -Identity '${escaped}' | Where-Object { $_.Trustee -ne 'NT AUTHORITY\\SELF' } | Select-Object Trustee`,
    );
    const sendAs = sendAsRaw
      ? Array.isArray(sendAsRaw) ? sendAsRaw : [sendAsRaw]
      : [];

    spin.stop(`Permissions for ${email}`);

    const permLines: string[] = [];

    if (fullAccess.length > 0) {
      permLines.push("Full Access:");
      for (const perm of fullAccess) {
        const rights = Array.isArray(perm.AccessRights)
          ? perm.AccessRights.join(", ")
          : perm.AccessRights;
        permLines.push(`  ${perm.User} (${rights})`);
      }
    } else {
      permLines.push("Full Access: none");
    }

    permLines.push("");

    if (sendAs.length > 0) {
      permLines.push("Send As:");
      for (const perm of sendAs) {
        permLines.push(`  ${perm.Trustee}`);
      }
    } else {
      permLines.push("Send As: none");
    }

    p.note(permLines.join("\n"), email);
  }
}
