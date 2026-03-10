import { resolve, dirname, join } from "path";
import { mkdirSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";
import { generateReport } from "../report-template.ts";
import { appDir } from "../utils.ts";

interface SubscribedSku {
  SkuId: string;
  SkuPartNumber: string;
}

interface FlatUser {
  DisplayName: string;
  UserPrincipalName: string;
  AccountEnabled: boolean;
  CreatedDateTime: string | null;
  UserType: string | null;
  LicenseSkuIds: string[] | null;
  LastSuccessfulSignIn: string | null;
  LastSignIn: string | null;
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

function getLastSignIn(user: FlatUser): Date | null {
  const dateStr = user.LastSuccessfulSignIn ?? user.LastSignIn;
  if (!dateStr) return null;
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
}

function formatDate(d: Date | null): string {
  if (!d) return "Never";
  return d.toISOString().slice(0, 10);
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

function getUserLicenses(
  user: FlatUser,
  skuMap: Map<string, string>,
): string[] {
  if (!user.LicenseSkuIds || user.LicenseSkuIds.length === 0) return [];
  return user.LicenseSkuIds.map((id) => {
    const partNumber = skuMap.get(id);
    return partNumber ? friendlySkuName(partNumber) : id;
  });
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

  // 2. Prompt threshold
  const threshold = await p.select({
    message: "How many days of inactivity?",
    options: [
      { value: 30, label: "30 days" },
      { value: 60, label: "60 days" },
      { value: 90, label: "90 days" },
      { value: -1, label: "Custom" },
    ],
  });
  if (p.isCancel(threshold)) return;

  let days = threshold as number;
  if (days === -1) {
    const custom = await p.text({
      message: "Enter number of days",
      placeholder: "45",
      validate: (v = "") => {
        const n = parseInt(v.trim(), 10);
        if (isNaN(n) || n < 1) return "Enter a positive number";
        return undefined;
      },
    });
    if (p.isCancel(custom)) return;
    days = parseInt((custom as string).trim(), 10);
  }

  // 3. Fetch users
  spin.start("Fetching users with sign-in activity…");
  const stopTimer = elapsedTimer(spin, "Fetching users with sign-in activity");

  const raw = await ps.runCommandJson<FlatUser | FlatUser[]>(
    [
      `Get-MgUser -All -Property 'UserPrincipalName','DisplayName','SignInActivity','AccountEnabled','CreatedDateTime','UserType','AssignedLicenses'`,
      `| ForEach-Object { [PSCustomObject]@{`,
      `  UserPrincipalName = $_.UserPrincipalName;`,
      `  DisplayName = $_.DisplayName;`,
      `  AccountEnabled = $_.AccountEnabled;`,
      `  CreatedDateTime = $_.CreatedDateTime;`,
      `  UserType = $_.UserType;`,
      `  LicenseSkuIds = @($_.AssignedLicenses.SkuId);`,
      `  LastSuccessfulSignIn = $_.SignInActivity.LastSuccessfulSignInDateTime;`,
      `  LastSignIn = $_.SignInActivity.LastSignInDateTime`,
      `} }`,
    ].join(" "),
  );
  stopTimer();

  const allUsers = raw ? (Array.isArray(raw) ? raw : [raw]) : [];

  // Fetch tenant SKU list to resolve license names
  const skuRaw = await ps.runCommandJson<SubscribedSku | SubscribedSku[]>(
    `Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber`,
  );
  const skuList = skuRaw ? (Array.isArray(skuRaw) ? skuRaw : [skuRaw]) : [];
  const skuMap = new Map(skuList.map((s) => [s.SkuId, s.SkuPartNumber]));

  spin.stop(`Fetched ${allUsers.length} user(s).`);

  if (allUsers.length === 0) {
    p.log.warn("No users found.");
    return;
  }

  // 5. Detect missing P1/P2
  const hasAnySignIn = allUsers.some(
    (u) => u.LastSuccessfulSignIn !== null || u.LastSignIn !== null,
  );
  if (!hasAnySignIn) {
    p.log.warn(
      "All users have empty sign-in activity. This usually means the tenant lacks an Entra ID P1/P2 license.",
    );
    return;
  }

  // 6. Filter
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);

  const inactive = allUsers.filter((u) => {
    // Exclude guests
    if (u.UserType !== "Member" || u.UserPrincipalName.includes("#EXT#"))
      return false;
    // Exclude unlicensed users (shared mailboxes, room mailboxes, etc.)
    if (!u.LicenseSkuIds || u.LicenseSkuIds.filter(Boolean).length === 0)
      return false;
    // Keep if last sign-in older than threshold or never signed in
    const lastSignIn = getLastSignIn(u);
    return lastSignIn === null || lastSignIn < cutoff;
  });

  if (inactive.length === 0) {
    p.log.info(`No inactive users found (threshold: ${days} days).`);
    return;
  }

  // 7. Sort — never-signed-in first (by creation date), then by last sign-in ascending
  inactive.sort((a, b) => {
    const aSign = getLastSignIn(a);
    const bSign = getLastSignIn(b);
    // Never signed in sorts first
    if (!aSign && bSign) return -1;
    if (aSign && !bSign) return 1;
    if (!aSign && !bSign) {
      // Sort by creation date ascending
      const aCreated = a.CreatedDateTime
        ? new Date(a.CreatedDateTime).getTime()
        : 0;
      const bCreated = b.CreatedDateTime
        ? new Date(b.CreatedDateTime).getTime()
        : 0;
      return aCreated - bCreated;
    }
    return aSign!.getTime() - bSign!.getTime();
  });

  // 8. Display
  const neverCount = inactive.filter((u) => !getLastSignIn(u)).length;
  const displayRows = inactive.slice(0, 50);

  const header = `${"Name".padEnd(25)} ${"UPN".padEnd(35)} ${"Last Sign-In".padEnd(12)} Notes`;
  const separator = "─".repeat(header.length);
  const rows = displayRows.map((u) => {
    const name = truncate(u.DisplayName ?? "", 24).padEnd(25);
    const upn = truncate(u.UserPrincipalName, 34).padEnd(35);
    const lastSign = formatDate(getLastSignIn(u)).padEnd(12);
    const notes = u.AccountEnabled ? "" : "Sign-in blocked";
    return `${name} ${upn} ${lastSign} ${notes}`;
  });

  const lines = [header, separator, ...rows];
  if (inactive.length > 50) {
    lines.push(`… and ${inactive.length - 50} more (export to Excel for full list)`);
  }
  p.note(lines.join("\n"), `Inactive users (>${days} days)`);

  // 9. Summary
  p.log.info(
    `Total: ${inactive.length} inactive · ${neverCount} never signed in`,
  );

  // 10. Excel export
  const exportXlsx = await p.confirm({
    message: "Export full results to Excel?",
    initialValue: false,
  });
  if (p.isCancel(exportXlsx) || !exportXlsx) return;

  const companyName = await p.text({
    message: "Company name",
    placeholder: ps.tenantDomain ?? "Acme Corp",
  });
  if (p.isCancel(companyName)) return;

  const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
  const dateSlug = new Date().toISOString().slice(0, 10);
  const outputDir = join(appDir(), "reports output");
  const fullPath = resolve(join(outputDir, `${tenantSlug}-users-report-${dateSlug}.xlsx`));
  mkdirSync(dirname(fullPath), { recursive: true });

  spin.start("Generating Excel report…");

  const buffer = await generateReport({
    sheetName: "Inactive Users",
    title: "Inactive Users Report",
    tenant: (companyName as string).trim(),
    summary: `${inactive.length} inactive (>${days} days) · ${neverCount} never signed in`,
    columns: [
      { header: "Display Name", width: 30 },
      { header: "UPN", width: 38 },
      { header: "Last Sign-In", width: 16 },
      { header: "Created", width: 16 },
      { header: "Licenses", width: 45 },
      { header: "Notes", width: 20 },
    ],
    rows: inactive.map((u) => {
      const lastSign = getLastSignIn(u);
      return [
        u.DisplayName ?? "",
        u.UserPrincipalName,
        formatDate(lastSign),
        u.CreatedDateTime
          ? new Date(u.CreatedDateTime).toISOString().slice(0, 10)
          : "",
        getUserLicenses(u, skuMap).join("; "),
        u.AccountEnabled ? "" : "Sign-in blocked",
      ];
    }),
  });

  await Bun.write(fullPath, buffer);
  spin.stop(`Exported ${inactive.length} rows to ${fullPath}`);

  const folder = dirname(fullPath);
  try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
}
