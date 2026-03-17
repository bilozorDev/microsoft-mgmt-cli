import { resolve, dirname, join } from "path";
import { mkdirSync, chmodSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";
import { generateReport } from "../report-template.ts";
import { appDir } from "../utils.ts";
import { GraphClient } from "../graph-client.ts";

interface SubscribedSku {
  skuId: string;
  skuPartNumber: string;
  consumedUnits: number;
  prepaidUnits: { enabled: number; warning: number; suspended: number };
}

interface Subscription {
  id: string;
  skuId: string;
  skuPartNumber: string;
  totalLicenses: number;
  status: string;
  isTrial: boolean;
  createdDateTime: string | null;
  nextLifecycleDateTime: string | null;
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
    await ps.ensureGraphConnected(["Organization.Read.All", "Directory.Read.All"]);
  } catch (e: any) {
    spin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(e.message);
    return;
  }
  spin.stop("Connected to Microsoft Graph.");

  // 2. Fetch SKUs and subscriptions via Graph REST API
  spin.start("Fetching license and subscription data…");
  const stopTimer = elapsedTimer(spin, "Fetching license and subscription data");

  const graph = new GraphClient(ps);

  let skuList: SubscribedSku[];
  let subscriptions: Subscription[];
  try {
    skuList = await graph.getAll<SubscribedSku>("/subscribedSkus", {
      params: { $select: "skuId,skuPartNumber,consumedUnits,prepaidUnits" },
    });

    subscriptions = await graph.getAll<Subscription>(
      "/directory/subscriptions",
      {
        params: {
          $select:
            "id,skuId,skuPartNumber,totalLicenses,status,isTrial,createdDateTime,nextLifecycleDateTime",
        },
      },
    );
  } catch (e: any) {
    spin.stop("Failed to fetch license and subscription data.");
    p.log.error(e.message ?? String(e));
    return;
  } finally {
    stopTimer();
  }

  const skuMap = new Map(skuList.map((s) => [s.skuId, s]));
  spin.stop(`Found ${subscriptions.length} subscription(s) across ${skuList.length} SKU(s).`);

  if (subscriptions.length === 0) {
    p.log.warn("No subscriptions found.");
    return;
  }

  // 3. Sort: nextLifecycleDateTime ascending, nulls last
  subscriptions.sort((a, b) => {
    const aDate = a.nextLifecycleDateTime;
    const bDate = b.nextLifecycleDateTime;
    if (!aDate && !bDate) return 0;
    if (!aDate) return 1;
    if (!bDate) return -1;
    return new Date(aDate).getTime() - new Date(bDate).getTime();
  });

  // 4. Build display rows
  const rows = subscriptions.map((sub) => {
    const sku = skuMap.get(sub.skuId);
    const assigned = sku ? sku.consumedUnits : 0;
    const available = sub.totalLicenses - assigned;

    return {
      subscriptionId: sub.id,
      skuPartNumber: sub.skuPartNumber,
      name: friendlySkuName(sub.skuPartNumber),
      status: sub.status,
      totalSeats: sub.totalLicenses,
      assigned,
      available,
      trial: sub.isTrial ? "Yes" : "No",
      renewalDate: formatDate(sub.nextLifecycleDateTime),
      rawRenewalDate: sub.nextLifecycleDateTime,
      daysUntilRenewal: daysUntil(sub.nextLifecycleDateTime),
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

  // 7. Calendar reminders
  const renewableRows = rows.filter((r) => r.rawRenewalDate && r.daysUntilRenewal !== null);
  if (renewableRows.length === 0) return;

  const createReminders = await p.confirm({
    message: "Create calendar reminders for upcoming renewals?",
    initialValue: false,
  });
  if (p.isCancel(createReminders) || !createReminders) return;

  const selected = await p.multiselect({
    message: "Select subscriptions to create reminders for:",
    options: renewableRows.map((r) => ({
      value: r,
      label: r.name,
      hint: `renews ${r.renewalDate} (${r.daysUntilRenewal} days)`,
    })),
  });
  if (p.isCancel(selected)) return;

  const tenantDomain = ps.tenantDomain ?? "Unknown";
  const icsDir = join(outputDir, "calendar reminders");
  mkdirSync(icsDir, { recursive: true });
  try { chmodSync(icsDir, 0o700); } catch {}

  for (const row of selected) {
    const startDate = new Date(row.rawRenewalDate!);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 1);

    const dtstart = startDate.toISOString().slice(0, 10).replace(/-/g, "");
    const dtend = endDate.toISOString().slice(0, 10).replace(/-/g, "");
    const uid = `${row.subscriptionId}-${tenantSlug}@m365-admin-cli`;

    const ics = [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//M365 Admin CLI//EN",
      "BEGIN:VEVENT",
      `UID:${uid}`,
      `DTSTART;VALUE=DATE:${dtstart}`,
      `DTEND;VALUE=DATE:${dtend}`,
      `SUMMARY:License Renewal: ${row.name} — ${tenantDomain}`,
      `DESCRIPTION:Subscription renewal/expiry for ${row.name}\\nTenant: ${tenantDomain}\\nSeats: ${row.totalSeats} (${row.assigned} assigned)\\nStatus: ${row.status}`,
      "BEGIN:VALARM",
      "TRIGGER:-P7D",
      "ACTION:DISPLAY",
      "DESCRIPTION:License renewal in 7 days",
      "END:VALARM",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\r\n");

    const icsPath = join(icsDir, `${tenantSlug}-renewal-${row.skuPartNumber}.ics`);
    await Bun.write(icsPath, ics);
    try { chmodSync(icsPath, 0o600); } catch {}
  }

  p.log.success(`Created ${selected.length} calendar reminder(s) in ${icsDir}`);
}
