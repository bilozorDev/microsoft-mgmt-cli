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
    `Get-Mailbox -RecipientTypeDetails SharedMailbox | Select-Object DisplayName,PrimarySmtpAddress`,
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
  const header = `${"Name".padEnd(30)} ${"Email".padEnd(40)} Licenses`;
  const separator = "─".repeat(header.length + 20);
  const displayRows = mailboxes.slice(0, 50);
  const rows = displayRows.map((m) => {
    const name = truncate(m.DisplayName ?? "", 29).padEnd(30);
    const email = truncate(m.PrimarySmtpAddress ?? "", 39).padEnd(40);
    const licenses = licenseMap.get(m.PrimarySmtpAddress.toLowerCase());
    const licStr = licenses ? `⚠ ${licenses.join(", ")}` : "—";
    return `${name} ${email} ${licStr}`;
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
    const fullPath = resolve(join(outputDir, `${tenantSlug}-shared-mailboxes-${dateSlug}.xlsx`));
    mkdirSync(dirname(fullPath), { recursive: true });

    // Fetch members (Full Access + Send As) for each mailbox
    const membersMap = new Map<string, string[]>();
    spin.start(`Fetching members (1/${mailboxes.length})…`);
    for (let i = 0; i < mailboxes.length; i++) {
      spin.message(`Fetching members (${i + 1}/${mailboxes.length})…`);
      const m = mailboxes[i]!;
      const escaped = m.PrimarySmtpAddress.replace(/'/g, "''");
      const members = new Set<string>();

      const fullAccessRaw = await ps.runCommandJson<
        MailboxPermission | MailboxPermission[]
      >(
        `Get-MailboxPermission -Identity '${escaped}' | Where-Object { $_.User -ne 'NT AUTHORITY\\SELF' -and $_.IsInherited -eq $false } | Select-Object User,AccessRights`,
      );
      const fullAccess = fullAccessRaw
        ? Array.isArray(fullAccessRaw) ? fullAccessRaw : [fullAccessRaw]
        : [];
      for (const perm of fullAccess) {
        if (!perm.User.startsWith("S-1-5-") && !perm.User.match(/^[0-9a-f-]{36}$/)) members.add(perm.User);
      }

      const sendAsRaw = await ps.runCommandJson<
        RecipientPermission | RecipientPermission[]
      >(
        `Get-RecipientPermission -Identity '${escaped}' | Where-Object { $_.Trustee -ne 'NT AUTHORITY\\SELF' } | Select-Object Trustee`,
      );
      const sendAs = sendAsRaw
        ? Array.isArray(sendAsRaw) ? sendAsRaw : [sendAsRaw]
        : [];
      for (const perm of sendAs) {
        if (!perm.Trustee.startsWith("S-1-5-") && !perm.Trustee.match(/^[0-9a-f-]{36}$/)) members.add(perm.Trustee);
      }

      if (members.size > 0) {
        membersMap.set(m.PrimarySmtpAddress.toLowerCase(), [...members]);
      }
    }
    spin.stop(`Fetched members for ${mailboxes.length} mailbox(es).`);

    spin.start("Generating Excel report…");

    const buffer = await generateReport({
      sheetName: "Shared Mailboxes",
      title: "Shared Mailboxes Report",
      tenant: ps.tenantDomain ?? "Unknown",
      summary: `${mailboxes.length} shared mailbox(es)`,
      columns: [
        { header: "Display Name", width: 30 },
        { header: "Email", width: 38 },
        { header: "Members", width: 40, wrapText: true },
        { header: "Notes", width: 40 },
      ],
      rows: mailboxes.map((m) => {
        const members = membersMap.get(m.PrimarySmtpAddress.toLowerCase()) ?? [];
        return [
          m.DisplayName ?? "",
          m.PrimarySmtpAddress,
          members.join("\n"),
          "",
        ];
      }),
    });

    await Bun.write(fullPath, buffer);
    spin.stop(`Exported ${mailboxes.length} rows to ${fullPath}`);

    const folder = dirname(fullPath);
    try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
  }

}
