import { resolve, dirname, join } from "path";
import { mkdirSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generateReport } from "../report-template.ts";
import { appDir, escapePS } from "../utils.ts";

interface DirectoryRole {
  Id: string;
  DisplayName: string;
}

interface RoleMember {
  Id: string;
  ODataType: string | null;
  DisplayName: string | null;
  UserPrincipalName: string | null;
}

interface AdminUserDetails {
  DisplayName: string;
  UserPrincipalName: string;
  AccountEnabled: boolean;
  CreatedDateTime: string | null;
  LastSignIn: string | null;
}

interface AdminEntry {
  id: string;
  displayName: string;
  upn: string;
  roles: string[];
  accountEnabled: boolean;
  createdDateTime: string | null;
  lastSignIn: string | null;
}

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

function formatDate(dateStr: string | null, fallback = "N/A"): string {
  if (!dateStr) return fallback;
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? fallback : d.toISOString().slice(0, 10);
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

export async function run(ps: PowerShellSession): Promise<void> {
  // 1. Connect to Graph
  const spin = p.spinner();
  spin.start("Connecting to Microsoft Graph…");
  try {
    await ps.ensureGraphConnected();
  } catch (e: any) {
    spin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(e.message);
    return;
  }
  spin.stop("Connected to Microsoft Graph.");

  p.log.warn(
    "Note: This report only includes permanently assigned roles. PIM-eligible (just-in-time) role assignments are not captured.",
  );

  // 2. Fetch all activated directory roles
  spin.start("Fetching directory roles…");
  const stopTimer = elapsedTimer(spin, "Fetching directory roles");

  const rolesRaw = await ps.runCommandJson<DirectoryRole | DirectoryRole[]>(
    `Get-MgDirectoryRole -All | Select-Object Id, DisplayName`,
  );
  const roles = rolesRaw ? (Array.isArray(rolesRaw) ? rolesRaw : [rolesRaw]) : [];

  stopTimer();
  spin.stop(`Found ${roles.length} activated role(s).`);

  if (roles.length === 0) {
    p.log.warn("No directory roles found.");
    return;
  }

  // 3. Fetch members for each role
  spin.start(`Fetching role members (0/${roles.length})…`);
  const admins = new Map<string, AdminEntry>();
  let totalAssignments = 0;

  for (let i = 0; i < roles.length; i++) {
    const role = roles[i]!;
    spin.message(`Fetching role members (${i + 1}/${roles.length}) — ${role.DisplayName}…`);

    let members: RoleMember[] = [];
    try {
      const membersRaw = await ps.runCommandJson<RoleMember | RoleMember[]>(
        [
          `Get-MgDirectoryRoleMember -DirectoryRoleId '${escapePS(role.Id)}' -All | ForEach-Object {`,
          `[PSCustomObject]@{`,
          `Id = $_.Id;`,
          `ODataType = $_.AdditionalProperties['@odata.type'];`,
          `DisplayName = $_.AdditionalProperties['displayName'];`,
          `UserPrincipalName = $_.AdditionalProperties['userPrincipalName']`,
          `} }`,
        ].join(" "),
      );
      members = membersRaw ? (Array.isArray(membersRaw) ? membersRaw : [membersRaw]) : [];
    } catch {
      // Skip roles where member listing fails
    }

    for (const m of members) {
      // Only include user objects
      if (!m.ODataType || !m.ODataType.includes("user")) continue;
      if (!m.Id) continue;

      totalAssignments++;
      const existing = admins.get(m.Id);
      if (existing) {
        existing.roles.push(role.DisplayName);
      } else {
        admins.set(m.Id, {
          id: m.Id,
          displayName: m.DisplayName ?? "(unknown)",
          upn: m.UserPrincipalName ?? "",
          roles: [role.DisplayName],
          accountEnabled: true,
          createdDateTime: null,
          lastSignIn: null,
        });
      }
    }
  }
  spin.stop(`Found ${admins.size} admin(s) across ${roles.length} roles (${totalAssignments} total assignments).`);

  if (admins.size === 0) {
    p.log.warn("No admin users found in any directory role.");
    return;
  }

  // 4. Fetch sign-in activity and created date for each admin
  const adminList = Array.from(admins.values());
  spin.start(`Fetching admin details (0/${adminList.length})…`);

  for (let i = 0; i < adminList.length; i++) {
    const admin = adminList[i]!;
    spin.message(`Fetching admin details (${i + 1}/${adminList.length})…`);

    try {
      const detailRaw = await ps.runCommandJson<AdminUserDetails>(
        [
          `Get-MgUser -UserId '${escapePS(admin.id)}' -Property 'DisplayName','UserPrincipalName','AccountEnabled','CreatedDateTime','SignInActivity'`,
          `| ForEach-Object { [PSCustomObject]@{`,
          `DisplayName = $_.DisplayName;`,
          `UserPrincipalName = $_.UserPrincipalName;`,
          `AccountEnabled = $_.AccountEnabled;`,
          `CreatedDateTime = $_.CreatedDateTime;`,
          `LastSignIn = $_.SignInActivity.LastSuccessfulSignInDateTime`,
          `} }`,
        ].join(" "),
      );
      if (detailRaw) {
        admin.displayName = detailRaw.DisplayName ?? admin.displayName;
        admin.upn = detailRaw.UserPrincipalName ?? admin.upn;
        admin.accountEnabled = detailRaw.AccountEnabled;
        admin.createdDateTime = detailRaw.CreatedDateTime;
        admin.lastSignIn = detailRaw.LastSignIn;
      }
    } catch {
      // Keep defaults if user lookup fails
    }
  }
  spin.stop(`Fetched details for ${adminList.length} admin(s).`);

  // 5. Detect missing P1/P2
  const hasAnySignIn = adminList.some((a) => a.lastSignIn !== null);
  if (!hasAnySignIn) {
    p.log.warn(
      "No sign-in data found for any admin. This usually means the tenant lacks an Entra ID P1/P2 license.",
    );
  }

  // 6. Sort: most roles first, then alphabetical
  adminList.sort((a, b) => {
    if (b.roles.length !== a.roles.length) return b.roles.length - a.roles.length;
    return a.displayName.localeCompare(b.displayName);
  });

  // 7. Display
  const disabledCount = adminList.filter((a) => !a.accountEnabled).length;
  const neverSignedIn = adminList.filter((a) => !a.lastSignIn).length;
  const uniqueRoles = new Set(adminList.flatMap((a) => a.roles)).size;

  const displayRows = adminList.slice(0, 50);
  const header = `${"Name".padEnd(22)} ${"Email".padEnd(32)} ${"Roles".padEnd(30)} ${"Last Sign-In".padEnd(20)} Enabled`;
  const separator = "─".repeat(header.length);
  const rows = displayRows.map((a) => {
    const name = truncate(a.displayName, 21).padEnd(22);
    const email = truncate(a.upn, 31).padEnd(32);
    const roleStr = truncate(a.roles.join(", "), 29).padEnd(30);
    const lastSign = formatDate(a.lastSignIn, "Info not available").padEnd(20);
    const enabled = a.accountEnabled ? "Yes" : "No";
    return `${name} ${email} ${roleStr} ${lastSign} ${enabled}`;
  });

  const lines = [header, separator, ...rows];
  if (adminList.length > 50) {
    lines.push(`… and ${adminList.length - 50} more (export to Excel for full list)`);
  }
  p.note(lines.join("\n"), "Admin Role Report");

  p.log.info(
    `${admins.size} admins across ${uniqueRoles} roles · ${totalAssignments} total role assignments · ${disabledCount} disabled · ${neverSignedIn} never signed in`,
  );

  // 8. Excel export
  const exportXlsx = await p.confirm({
    message: "Export to Excel?",
    initialValue: false,
  });
  if (p.isCancel(exportXlsx) || !exportXlsx) return;

  const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
  const dateSlug = new Date().toISOString().slice(0, 10);
  const outputDir = join(appDir(), "reports output");
  const fullPath = resolve(join(outputDir, `${tenantSlug}-admin-roles-${dateSlug}.xlsx`));
  mkdirSync(dirname(fullPath), { recursive: true });

  spin.start("Generating Excel report…");

  const buffer = await generateReport({
    sheetName: "Admin Roles",
    title: "Admin Role Report",
    tenant: "",
    summary: `${admins.size} admins across ${uniqueRoles} roles · ${totalAssignments} total assignments · ${disabledCount} disabled · ${neverSignedIn} never signed in`,
    columns: [
      { header: "User", width: 28 },
      { header: "Email", width: 38 },
      { header: "Admin Roles", width: 50, wrapText: true },
      { header: "Last Sign-In", width: 16 },
      { header: "Account Created", width: 16 },
      { header: "Account Enabled", width: 16 },
    ],
    rows: adminList.map((a) => [
      a.displayName,
      a.upn,
      a.roles.join("\n"),
      formatDate(a.lastSignIn, "Info not available"),
      formatDate(a.createdDateTime),
      a.accountEnabled ? "Yes" : "No",
    ]),
  });

  await Bun.write(fullPath, buffer);
  spin.stop(`Exported ${adminList.length} rows to ${fullPath}`);

  const folder = dirname(fullPath);
  try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
}
