import { resolve } from "path";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";

interface SharedMailbox {
  DisplayName: string;
  PrimarySmtpAddress: string;
  Alias: string;
  WhenCreated: string;
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

  // Display table
  const header = `${"Name".padEnd(30)} ${"Email".padEnd(40)} ${"Created".padEnd(12)}`;
  const separator = "─".repeat(header.length);
  const displayRows = mailboxes.slice(0, 50);
  const rows = displayRows.map((m) => {
    const name = truncate(m.DisplayName ?? "", 29).padEnd(30);
    const email = truncate(m.PrimarySmtpAddress ?? "", 39).padEnd(40);
    const created = m.WhenCreated
      ? new Date(m.WhenCreated).toISOString().slice(0, 10)
      : "";
    return `${name} ${email} ${created.padEnd(12)}`;
  });

  const lines = [header, separator, ...rows];
  if (mailboxes.length > 50) {
    lines.push(`… and ${mailboxes.length - 50} more`);
  }
  p.note(lines.join("\n"), `Shared mailboxes (${mailboxes.length})`);

  // CSV export
  const exportCsv = await p.confirm({
    message: "Export to CSV?",
    initialValue: false,
  });
  if (p.isCancel(exportCsv)) return;

  if (exportCsv) {
    const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
    const dateSlug = new Date().toISOString().slice(0, 10);
    const defaultName = `${tenantSlug}-shared-mailboxes-${dateSlug}.csv`;

    const csvPath = await p.text({
      message: "File path",
      placeholder: defaultName,
      defaultValue: defaultName,
    });
    if (p.isCancel(csvPath)) return;

    const fullPath = resolve((csvPath as string).trim());
    const csvLines = [
      "DisplayName,PrimarySmtpAddress,Alias,WhenCreated",
      ...mailboxes.map((m) => {
        const created = m.WhenCreated
          ? new Date(m.WhenCreated).toISOString().slice(0, 10)
          : "";
        return `"${(m.DisplayName ?? "").replace(/"/g, '""')}","${m.PrimarySmtpAddress}","${m.Alias}","${created}"`;
      }),
    ];
    await Bun.write(fullPath, csvLines.join("\n"));
    p.log.success(`Exported to ${fullPath}`);
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
