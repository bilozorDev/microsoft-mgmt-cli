import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";
import { escapePS } from "../utils.ts";

interface MgUser {
  DisplayName: string;
  UserPrincipalName: string;
  Id: string;
  LicenseCount: number;
}

interface UserDetails {
  DisplayName: string;
  GivenName: string | null;
  Surname: string | null;
  UserPrincipalName: string;
  Mail: string | null;
  JobTitle: string | null;
  Department: string | null;
  AccountEnabled: boolean;
  CreatedDateTime: string | null;
  LastSignInDateTime: string | null;
}

interface LicenseDetail {
  SkuPartNumber: string;
}

interface DirectoryRole {
  DisplayName: string;
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

function friendlyMfaMethod(odataType: string): string | null {
  const lastSegment = odataType.split(".").pop() ?? "";
  if (lastSegment === "passwordAuthenticationMethod") return null;
  return MFA_METHOD_NAMES[lastSegment] ?? lastSegment;
}

async function fetchUsers(ps: PowerShellSession): Promise<MgUser[]> {
  const spin = p.spinner();
  spin.start("Loading users...");

  const { output: countOutput } = await ps.runCommand(
    "Get-MgUser -Top 1 -CountVariable ct -ConsistencyLevel eventual | Out-Null; $ct",
  );
  const count = parseInt(countOutput.trim(), 10);

  let users: MgUser[];

  if (count <= 50) {
    const raw = await ps.runCommandJson<MgUser | MgUser[]>(
      "Get-MgUser -All -Property DisplayName,UserPrincipalName,Id,AssignedLicenses | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName; Id = $_.Id; LicenseCount = $_.AssignedLicenses.Count } }",
    );
    users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
    spin.stop(`Found ${users.length} user(s).`);
  } else {
    spin.stop(`${count} users in tenant — search to find a user.`);
    while (true) {
      const query = await p.text({
        message: "Search for user by name",
        placeholder: "e.g. Jane Doe",
        validate: (v = "") => (!v.trim() ? "Enter a search term" : undefined),
      });
      if (p.isCancel(query)) return [];

      const searchSpin = p.spinner();
      searchSpin.start("Searching users...");
      try {
        const raw = await ps.runCommandJson<MgUser | MgUser[]>(
          `Get-MgUser -Search '"displayName:${escapePS(query)}"' -ConsistencyLevel eventual -Property DisplayName,UserPrincipalName,Id,AssignedLicenses | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName; Id = $_.Id; LicenseCount = $_.AssignedLicenses.Count } }`,
        );
        users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        searchSpin.stop(`Found ${users.length} user(s).`);
      } catch {
        searchSpin.stop("Search returned no results.");
        users = [];
      }

      if (users.length === 0) {
        p.log.warn("No users found. Try a different search term.");
        continue;
      }
      break;
    }

    return users.sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));
  }

  return users.sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));
}

export async function run(ps: PowerShellSession): Promise<void> {
  // 1. Connect to Graph
  const spin = p.spinner();
  spin.start("Connecting to Microsoft Graph…");
  try {
    await ps.ensureGraphConnected([
      "User.Read.All",
      "Directory.Read.All",
      "RoleManagement.Read.Directory",
      "UserAuthenticationMethod.Read.All",
      "AuditLog.Read.All",
    ]);
  } catch (e: any) {
    spin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(e.message);
    return;
  }
  spin.stop("Connected to Microsoft Graph.");

  // 2. Pick a user
  const users = await fetchUsers(ps);
  if (users.length === 0) return;

  const userId = await p.select({
    message: "Select a user",
    options: users.map((u) => ({
      value: u.Id,
      label: u.DisplayName,
      hint: u.LicenseCount > 0 ? u.UserPrincipalName : `${u.UserPrincipalName} (not licensed)`,
    })),
  });
  if (p.isCancel(userId)) return;

  const selectedUser = users.find((u) => u.Id === userId)!;

  // 3. Fetch all info sequentially (single PS stdout stream)
  spin.start("Fetching user details…");

  // User details
  const details = await ps.runCommandJson<UserDetails>(
    [
      `Get-MgUser -UserId '${escapePS(userId)}'`,
      `-Property DisplayName,GivenName,Surname,UserPrincipalName,Mail,JobTitle,Department,AccountEnabled,CreatedDateTime,SignInActivity`,
      `| ForEach-Object { [PSCustomObject]@{`,
      `DisplayName = $_.DisplayName;`,
      `GivenName = $_.GivenName;`,
      `Surname = $_.Surname;`,
      `UserPrincipalName = $_.UserPrincipalName;`,
      `Mail = $_.Mail;`,
      `JobTitle = $_.JobTitle;`,
      `Department = $_.Department;`,
      `AccountEnabled = $_.AccountEnabled;`,
      `CreatedDateTime = if ($_.CreatedDateTime) { $_.CreatedDateTime.ToString('yyyy-MM-dd') } else { $null };`,
      `LastSignInDateTime = if ($_.SignInActivity.LastSignInDateTime) { $_.SignInActivity.LastSignInDateTime.ToString('yyyy-MM-dd HH:mm') } else { $null }`,
      `} }`,
    ].join(" "),
  );

  // Licenses
  spin.message("Fetching licenses…");
  const licensesRaw = await ps.runCommandJson<LicenseDetail | LicenseDetail[]>(
    `Get-MgUserLicenseDetail -UserId '${escapePS(userId)}' | Select-Object SkuPartNumber`,
  );

  // Roles
  spin.message("Fetching admin roles…");
  const rolesRaw = await ps.runCommandJson<DirectoryRole | DirectoryRole[]>(
    [
      `Get-MgUserMemberOf -UserId '${escapePS(userId)}' -All`,
      `| Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' }`,
      `| ForEach-Object { [PSCustomObject]@{ DisplayName = $_.AdditionalProperties['displayName'] } }`,
    ].join(" "),
  );

  // Mailbox size
  spin.message("Fetching mailbox size…");
  let mailboxSize = "No mailbox";
  try {
    const { output, error } = await ps.runCommand(
      `try { $stats = Get-EXOMailboxStatistics -Identity '${escapePS(selectedUser.UserPrincipalName)}' -ErrorAction SilentlyContinue; if ($stats) { $bytes = $stats.TotalItemSize.Value.ToBytes(); if ($bytes -ge 1GB) { "{0:N2} GB" -f ($bytes / 1GB) } elseif ($bytes -ge 1MB) { "{0:N2} MB" -f ($bytes / 1MB) } else { "{0:N0} KB" -f ($bytes / 1KB) } } else { "No mailbox" } } catch { "No mailbox" }`,
    );
    mailboxSize = error ? "No mailbox" : (output || "No mailbox");
  } catch {
    // keep default
  }

  // MFA methods
  spin.message("Fetching 2FA methods…");
  let mfaMethods: string[] = [];
  try {
    const methodsRaw = await ps.runCommandJson<AuthMethod | AuthMethod[]>(
      [
        `Get-MgUserAuthenticationMethod -UserId '${escapePS(userId)}'`,
        `| ForEach-Object { [PSCustomObject]@{ ODataType = $_.AdditionalProperties['@odata.type'] } }`,
      ].join(" "),
    );
    const methods = methodsRaw ? (Array.isArray(methodsRaw) ? methodsRaw : [methodsRaw]) : [];
    for (const m of methods) {
      if (!m.ODataType) continue;
      const name = friendlyMfaMethod(m.ODataType);
      if (name) mfaMethods.push(name);
    }
  } catch {
    mfaMethods = ["Error fetching"];
  }

  spin.stop("Done.");

  // 4. Format and display
  const licenses = licensesRaw
    ? (Array.isArray(licensesRaw) ? licensesRaw : [licensesRaw])
    : [];
  const roles = rolesRaw
    ? (Array.isArray(rolesRaw) ? rolesRaw : [rolesRaw])
    : [];

  const lines: string[] = [];

  // Basic info
  if (details) {
    lines.push(`  Name:           ${details.DisplayName}`);
    lines.push(`  Email:          ${details.UserPrincipalName}`);
    if (details.Mail && details.Mail !== details.UserPrincipalName) {
      lines.push(`  Mail:           ${details.Mail}`);
    }
    if (details.JobTitle) lines.push(`  Job Title:      ${details.JobTitle}`);
    if (details.Department) lines.push(`  Department:     ${details.Department}`);
    lines.push(`  Account:        ${details.AccountEnabled ? "Enabled" : "Disabled"}`);
    if (details.CreatedDateTime) lines.push(`  Created:        ${details.CreatedDateTime}`);
    if (details.LastSignInDateTime) lines.push(`  Last Sign-In:   ${details.LastSignInDateTime}`);
  }

  lines.push("");

  // Licenses
  lines.push(`  Licenses (${licenses.length}):`);
  if (licenses.length === 0) {
    lines.push("    None");
  } else {
    for (const lic of licenses) {
      lines.push(`    - ${friendlySkuName(lic.SkuPartNumber)}`);
    }
  }

  lines.push("");

  // Admin roles
  lines.push(`  Admin Roles (${roles.length}):`);
  if (roles.length === 0) {
    lines.push("    None");
  } else {
    for (const role of roles) {
      lines.push(`    - ${role.DisplayName}`);
    }
  }

  lines.push("");

  // Mailbox
  lines.push(`  Mailbox Size:   ${mailboxSize}`);

  lines.push("");

  // MFA
  lines.push(`  2FA Methods (${mfaMethods.length}):`);
  if (mfaMethods.length === 0) {
    lines.push("    Not configured");
  } else {
    for (const method of mfaMethods) {
      lines.push(`    - ${method}`);
    }
  }

  p.note(lines.join("\n"), selectedUser.DisplayName);
}
