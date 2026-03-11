import { resolve, dirname, join } from "path";
import { mkdirSync, chmodSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";
import { generateReport } from "../report-template.ts";
import { appDir } from "../utils.ts";

interface SubscribedSku {
  SkuId: string;
  SkuPartNumber: string;
  ConsumedUnits: number;
  Enabled: number;
  Warning: number;
  Suspended: number;
}

interface Subscription {
  Id: string;
  SkuId: string;
  SkuPartNumber: string;
  TotalLicenses: number;
  Status: string;
  IsTrial: boolean;
  CreatedDateTime: string | null;
  NextLifecycleDateTime: string | null;
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

function daysUntil(dateStr: string | null): number | null {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return null;
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);
  return Math.round((d.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
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

  // 2. Fetch SKUs, subscriptions, and tenant ID in sequence
  spin.start("Fetching license and subscription data…");
  const stopTimer = elapsedTimer(spin, "Fetching license and subscription data");

  const skuRaw = await ps.runCommandJson<SubscribedSku | SubscribedSku[]>(
    [
      `Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits,`,
      `@{N='Enabled';E={$_.PrepaidUnits.Enabled}},`,
      `@{N='Warning';E={$_.PrepaidUnits.Warning}},`,
      `@{N='Suspended';E={$_.PrepaidUnits.Suspended}}`,
    ].join(" "),
  );
  const skuList = skuRaw ? (Array.isArray(skuRaw) ? skuRaw : [skuRaw]) : [];
  const skuMap = new Map(skuList.map((s) => [s.SkuId, s]));

  const subsRaw = await ps.runCommandJson<Subscription | Subscription[]>(
    [
      `Get-MgDirectorySubscription -All | Select-Object Id, SkuId, SkuPartNumber,`,
      `TotalLicenses, Status, IsTrial, CreatedDateTime, NextLifecycleDateTime`,
    ].join(" "),
  );
  const subscriptions = subsRaw ? (Array.isArray(subsRaw) ? subsRaw : [subsRaw]) : [];

  stopTimer();
  spin.stop(`Found ${subscriptions.length} subscription(s) across ${skuList.length} SKU(s).`);

  if (subscriptions.length === 0) {
    p.log.warn("No subscriptions found.");
    return;
  }

  // 3. Sort: NextLifecycleDateTime ascending, nulls last
  subscriptions.sort((a, b) => {
    const aDate = a.NextLifecycleDateTime;
    const bDate = b.NextLifecycleDateTime;
    if (!aDate && !bDate) return 0;
    if (!aDate) return 1;
    if (!bDate) return -1;
    return new Date(aDate).getTime() - new Date(bDate).getTime();
  });

  // 4. Build display rows
  const rows = subscriptions.map((sub) => {
    const sku = skuMap.get(sub.SkuId);
    const assigned = sku ? sku.ConsumedUnits : 0;
    const available = sub.TotalLicenses - assigned;

    return {
      name: friendlySkuName(sub.SkuPartNumber),
      status: sub.Status,
      totalSeats: sub.TotalLicenses,
      assigned,
      available,
      trial: sub.IsTrial ? "Yes" : "No",
      renewalDate: formatDate(sub.NextLifecycleDateTime),
      daysUntilRenewal: daysUntil(sub.NextLifecycleDateTime),
    };
  });

  // 5. Terminal preview
  const displayRows = rows.slice(0, 50);
  const header = `${"License".padEnd(32)} ${"Status".padEnd(12)} ${"Seats".padEnd(8)} ${"Used".padEnd(8)} ${"Avail".padEnd(8)} ${"Trial".padEnd(6)} ${"Renewal".padEnd(12)} ${"Days".padEnd(6)}`;
  const separator = "─".repeat(header.length);

  const terminalRows = displayRows.map((r) => {
    return [
      truncate(r.name, 31).padEnd(32),
      truncate(r.status, 11).padEnd(12),
      String(r.totalSeats).padEnd(8),
      String(r.assigned).padEnd(8),
      String(r.available).padEnd(8),
      r.trial.padEnd(6),
      r.renewalDate.padEnd(12),
      (r.daysUntilRenewal !== null ? String(r.daysUntilRenewal) : "N/A").padEnd(6),
    ].join(" ");
  });

  const lines = [header, separator, ...terminalRows];
  if (rows.length > 50) {
    lines.push(`… and ${rows.length - 50} more (export to Excel for full list)`);
  }
  p.note(lines.join("\n"), `License & Subscription Report (${rows.length})`);

  // 6. Excel export
  const exportXlsx = await p.confirm({
    message: "Export to Excel?",
    initialValue: false,
  });
  if (p.isCancel(exportXlsx) || !exportXlsx) return;

  const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
  const dateSlug = new Date().toISOString().slice(0, 10);
  const outputDir = join(appDir(), "reports output");
  const fullPath = resolve(join(outputDir, `${tenantSlug}-license-report-${dateSlug}.xlsx`));
  mkdirSync(dirname(fullPath), { recursive: true });
  try { chmodSync(dirname(fullPath), 0o700); } catch {}

  spin.start("Generating Excel report…");

  const buffer = await generateReport({
    sheetName: "Licenses & Subscriptions",
    title: "License & Subscription Report",
    tenant: "",
    summary: `${rows.length} subscriptions`,
    columns: [
      { header: "License Name", width: 35 },
      { header: "Status", width: 14 },
      { header: "Total Seats", width: 12 },
      { header: "Assigned", width: 12 },
      { header: "Available", width: 12 },
      { header: "Trial", width: 8 },
      { header: "Renewal / Expiry Date", width: 18 },
      { header: "Days Until Renewal", width: 18 },
      { header: "Notes", width: 20 },
    ],
    rows: rows.map((r) => [
      r.name,
      r.status,
      r.totalSeats,
      r.assigned,
      r.available,
      r.trial,
      r.renewalDate,
      r.daysUntilRenewal !== null ? r.daysUntilRenewal : "N/A",
      "",
    ]),
  });

  await Bun.write(fullPath, buffer);
  try { chmodSync(fullPath, 0o600); } catch {}
  spin.stop(`Exported ${rows.length} rows to ${fullPath}`);

  const folder = dirname(fullPath);
  try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
}
