import { resolve, dirname, join } from "path";
import { mkdirSync, chmodSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";
import { generateReport } from "../report-template.ts";
import { appDir, escapePS } from "../utils.ts";

interface SubscribedSku {
  SkuId: string;
  SkuPartNumber: string;
}

interface LicensedUser {
  Id: string;
  DisplayName: string;
  UserPrincipalName: string;
  LicenseSkuIds: string[];
}

interface DirectoryRole {
  Id: string;
  DisplayName: string;
}

interface RoleMember {
  Id: string;
  ODataType: string | null;
}

interface AuthMethod {
  ODataType: string | null;
}

const MFA_METHOD_NAMES: Record<string, string> = {
  microsoftAuthenticatorAuthenticationMethod: "Authenticator App",
  phoneAuthenticationMethod: "Phone",
  fido2AuthenticationMethod: "FIDO2 Security Key",
  emailAuthenticationMethod: "Email",
  softwareOathAuthenticationMethod: "Software Token",
  windowsHelloForBusinessAuthenticationMethod: "Windows Hello",
  temporaryAccessPassAuthenticationMethod: "Temporary Access Pass",
  platformCredentialAuthenticationMethod: "Platform Credential",
};

function elapsedTimer(
  spin: { message(msg?: string): void },
  baseMsg: string,
): () => void {
  const start = Date.now();
  const interval = setInterval(() => {
    const secs = Math.floor((Date.now() - start) / 1000);
    const mins = Math.floor(secs / 60);
    const elapsed = mins > 0 ? `${mins}m ${secs % 60}s` : `${secs}s`;
    spin.message(`${baseMsg} (${elapsed})`);
  }, 1000);
  return () => clearInterval(interval);
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

function friendlyMfaMethod(odataType: string): string | null {
  // e.g. "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"
  const lastSegment = odataType.split(".").pop() ?? "";
  if (lastSegment === "passwordAuthenticationMethod") return null;
  return MFA_METHOD_NAMES[lastSegment] ?? lastSegment;
}

function getUserLicenses(
  skuIds: string[],
  skuMap: Map<string, string>,
): string[] {
  return skuIds.map((id) => {
    const partNumber = skuMap.get(id);
    return partNumber ? friendlySkuName(partNumber) : id;
  });
}

export async function run(ps: PowerShellSession): Promise<void> {
  await ps.ensureExchangeConnected();

  // 1. Connect to Graph
  const spin = p.spinner();
  spin.start("Connecting to Microsoft Graph…");
  try {
    await ps.ensureGraphConnected(["User.Read.All", "Organization.Read.All", "RoleManagement.Read.Directory", "AuditLog.Read.All", "UserAuthenticationMethod.Read.All"]);
  } catch (e: any) {
    spin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(e.message);
    return;
  }
  spin.stop("Connected to Microsoft Graph.");

  // 2. Ask which optional columns to include
  const extras = await p.multiselect({
    message: "Include optional columns? (space to toggle, enter to confirm)",
    options: [
      { value: "mailbox", label: "Mailbox Size", hint: "per-user, can be slow" },
      { value: "mfa", label: "MFA Methods", hint: "per-user, can be slow" },
      { value: "roles", label: "Admin Roles", hint: "bulk fetch" },
    ],
    required: false,
  });
  if (p.isCancel(extras)) return;

  const includeMailbox = extras.includes("mailbox");
  const includeMfa = extras.includes("mfa");
  const includeRoles = extras.includes("roles");

  // 3. Fetch all licensed users
  spin.start("Fetching licensed users…");
  const stopTimer = elapsedTimer(spin, "Fetching licensed users");

  const usersRaw = await ps.runCommandJson<LicensedUser | LicensedUser[]>(
    [
      `Get-MgUser -All -Property 'Id','DisplayName','UserPrincipalName','AssignedLicenses'`,
      `| Where-Object { $_.AssignedLicenses.Count -gt 0 }`,
      `| ForEach-Object { [PSCustomObject]@{`,
      `  Id = $_.Id;`,
      `  DisplayName = $_.DisplayName;`,
      `  UserPrincipalName = $_.UserPrincipalName;`,
      `  LicenseSkuIds = @($_.AssignedLicenses.SkuId)`,
      `} }`,
    ].join(" "),
  );
  stopTimer();

  const users = usersRaw ? (Array.isArray(usersRaw) ? usersRaw : [usersRaw]) : [];

  // Fetch SKU map
  const skuRaw = await ps.runCommandJson<SubscribedSku | SubscribedSku[]>(
    `Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber`,
  );
  const skuList = skuRaw ? (Array.isArray(skuRaw) ? skuRaw : [skuRaw]) : [];
  const skuMap = new Map(skuList.map((s) => [s.SkuId, s.SkuPartNumber]));

  spin.stop(`Found ${users.length} licensed user(s).`);

  if (users.length === 0) {
    p.log.info("No licensed users found.");
    return;
  }

  // 4. Admin Roles (bulk)
  const roleMap = new Map<string, string[]>(); // userId → role names
  if (includeRoles) {
    spin.start("Fetching admin roles…");
    const rolesRaw = await ps.runCommandJson<DirectoryRole | DirectoryRole[]>(
      `Get-MgDirectoryRole -All | Select-Object Id, DisplayName`,
    );
    const roles = rolesRaw ? (Array.isArray(rolesRaw) ? rolesRaw : [rolesRaw]) : [];

    for (let i = 0; i < roles.length; i++) {
      const role = roles[i]!;
      spin.message(`Fetching role members (${i + 1}/${roles.length}) — ${role.DisplayName}…`);
      try {
        const membersRaw = await ps.runCommandJson<RoleMember | RoleMember[]>(
          [
            `Get-MgDirectoryRoleMember -DirectoryRoleId '${escapePS(role.Id)}' -All | ForEach-Object {`,
            `[PSCustomObject]@{`,
            `Id = $_.Id;`,
            `ODataType = $_.AdditionalProperties['@odata.type']`,
            `} }`,
          ].join(" "),
        );
        const members = membersRaw ? (Array.isArray(membersRaw) ? membersRaw : [membersRaw]) : [];
        for (const m of members) {
          if (!m.ODataType?.includes("user") || !m.Id) continue;
          const existing = roleMap.get(m.Id);
          if (existing) {
            existing.push(role.DisplayName);
          } else {
            roleMap.set(m.Id, [role.DisplayName]);
          }
        }
      } catch {
        // Skip roles where member listing fails
      }
    }
    spin.stop(`Fetched admin roles (${roleMap.size} admin(s) found).`);
  }

  // 5. Mailbox Size (per-user)
  const mailboxMap = new Map<string, string>(); // upn → size string
  if (includeMailbox) {
    spin.start(`Fetching mailbox sizes (0/${users.length})…`);
    const stopMbTimer = elapsedTimer(spin, "Fetching mailbox sizes");

    for (let i = 0; i < users.length; i++) {
      const user = users[i]!;
      spin.message(`Fetching mailbox sizes (${i + 1}/${users.length})…`);

      if (!user.UserPrincipalName) continue;
      try {
        const { output, error } = await ps.runCommand(
          `try { $stats = Get-EXOMailboxStatistics -Identity '${escapePS(user.UserPrincipalName)}' -ErrorAction SilentlyContinue; if ($stats) { $bytes = $stats.TotalItemSize.Value.ToBytes(); if ($bytes -ge 1GB) { "{0:N2} GB" -f ($bytes / 1GB) } elseif ($bytes -ge 1MB) { "{0:N2} MB" -f ($bytes / 1MB) } else { "{0:N0} KB" -f ($bytes / 1KB) } } else { "No mailbox" } } catch { "No mailbox" }`,
        );
        mailboxMap.set(user.UserPrincipalName, error ? "No mailbox" : (output || "No mailbox"));
      } catch {
        mailboxMap.set(user.UserPrincipalName, "No mailbox");
      }
    }

    stopMbTimer();
    spin.stop(`Fetched mailbox sizes for ${users.length} user(s).`);
  }

  // 6. MFA Methods (per-user)
  const mfaMap = new Map<string, string[]>(); // userId → method names
  if (includeMfa) {
    spin.start(`Fetching MFA methods (0/${users.length})…`);
    const stopMfaTimer = elapsedTimer(spin, "Fetching MFA methods");

    for (let i = 0; i < users.length; i++) {
      const user = users[i]!;
      spin.message(`Fetching MFA methods (${i + 1}/${users.length})…`);

      try {
        const methodsRaw = await ps.runCommandJson<AuthMethod | AuthMethod[]>(
          [
            `Get-MgUserAuthenticationMethod -UserId '${escapePS(user.Id)}'`,
            `| ForEach-Object { [PSCustomObject]@{ ODataType = $_.AdditionalProperties['@odata.type'] } }`,
          ].join(" "),
        );
        const methods = methodsRaw ? (Array.isArray(methodsRaw) ? methodsRaw : [methodsRaw]) : [];
        const names: string[] = [];
        for (const m of methods) {
          if (!m.ODataType) continue;
          const name = friendlyMfaMethod(m.ODataType);
          if (name) names.push(name);
        }
        mfaMap.set(user.Id, names.length > 0 ? names : ["Not configured"]);
      } catch {
        mfaMap.set(user.Id, ["Error fetching"]);
      }
    }

    stopMfaTimer();
    spin.stop(`Fetched MFA methods for ${users.length} user(s).`);
  }

  // 7. Sort alphabetically by display name
  users.sort((a, b) => (a.DisplayName ?? "").localeCompare(b.DisplayName ?? ""));

  // 8. Terminal preview
  const displayRows = users.slice(0, 50);
  const headerParts = [`${"Name".padEnd(25)} ${"Email".padEnd(35)} ${"Licenses".padEnd(30)}`];
  if (includeRoles) headerParts.push(`${"Admin Roles".padEnd(25)}`);
  if (includeMailbox) headerParts.push(`${"Mailbox Size".padEnd(22)}`);
  if (includeMfa) headerParts.push(`${"MFA Methods".padEnd(25)}`);
  const header = headerParts.join(" ");
  const separator = "─".repeat(header.length);

  const rows = displayRows.map((u) => {
    const licenses = getUserLicenses(u.LicenseSkuIds ?? [], skuMap);
    const parts = [
      truncate(u.DisplayName ?? "", 24).padEnd(25),
      truncate(u.UserPrincipalName, 34).padEnd(35),
      truncate(licenses.join(", "), 29).padEnd(30),
    ];
    if (includeRoles) {
      const roles = roleMap.get(u.Id) ?? [];
      parts.push(truncate(roles.length > 0 ? roles.join(", ") : "—", 24).padEnd(25));
    }
    if (includeMailbox) {
      parts.push(truncate(mailboxMap.get(u.UserPrincipalName) ?? "—", 21).padEnd(22));
    }
    if (includeMfa) {
      const methods = mfaMap.get(u.Id) ?? [];
      parts.push(truncate(methods.length > 0 ? methods.join(", ") : "—", 24).padEnd(25));
    }
    return parts.join(" ");
  });

  const lines = [header, separator, ...rows];
  if (users.length > 50) {
    lines.push(`… and ${users.length - 50} more (export to Excel for full list)`);
  }
  p.note(lines.join("\n"), `Licensed Users (${users.length})`);

  // 9. Excel export
  const exportXlsx = await p.confirm({
    message: "Export to Excel?",
    initialValue: false,
  });
  if (p.isCancel(exportXlsx) || !exportXlsx) return;

  const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
  const dateSlug = new Date().toISOString().slice(0, 10);
  const outputDir = join(appDir(), "reports output");
  const fullPath = resolve(join(outputDir, `${tenantSlug}-licensed-users-${dateSlug}.xlsx`));
  mkdirSync(dirname(fullPath), { recursive: true });
  try { chmodSync(dirname(fullPath), 0o700); } catch {}

  spin.start("Generating Excel report…");

  // Build dynamic columns
  const columns: { header: string; width: number; wrapText?: boolean }[] = [
    { header: "Display Name", width: 30 },
    { header: "Email", width: 38 },
    { header: "Licenses", width: 45, wrapText: true },
  ];
  if (includeRoles) columns.push({ header: "Admin Roles", width: 45, wrapText: true });
  if (includeMailbox) columns.push({ header: "Mailbox Size", width: 28 });
  if (includeMfa) columns.push({ header: "MFA Methods", width: 40, wrapText: true });
  columns.push({ header: "Notes", width: 20 });

  const excelRows = users.map((u) => {
    const licenses = getUserLicenses(u.LicenseSkuIds ?? [], skuMap);
    const row: string[] = [
      u.DisplayName ?? "",
      u.UserPrincipalName,
      licenses.join("\n"),
    ];
    if (includeRoles) {
      const roles = roleMap.get(u.Id) ?? [];
      row.push(roles.join("\n"));
    }
    if (includeMailbox) {
      row.push(mailboxMap.get(u.UserPrincipalName) ?? "");
    }
    if (includeMfa) {
      const methods = mfaMap.get(u.Id) ?? [];
      row.push(methods.join("\n"));
    }
    row.push("");
    return row;
  });

  // Build summary
  const summaryParts = [`${users.length} licensed users`];
  if (includeRoles) summaryParts.push(`${roleMap.size} with admin roles`);
  if (includeMailbox) summaryParts.push("mailbox sizes included");
  if (includeMfa) summaryParts.push("MFA methods included");

  const buffer = await generateReport({
    sheetName: "Licensed Users",
    title: "Licensed Users Report",
    tenant: "",
    summary: summaryParts.join(" · "),
    columns,
    rows: excelRows,
  });

  await Bun.write(fullPath, buffer);
  try { chmodSync(fullPath, 0o600); } catch {}
  spin.stop(`Exported ${users.length} rows to ${fullPath}`);

  const folder = dirname(fullPath);
  try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
}
