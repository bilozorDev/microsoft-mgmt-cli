import { resolve, dirname, join } from "path";
import { mkdirSync, chmodSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generateReport } from "../report-template.ts";
import { appDir, escapePS } from "../utils.ts";

interface TraceMessage {
  MessageId: string;
  MessageTraceId: string;
  SenderAddress: string;
  RecipientAddress: string;
  Subject: string;
  Status: string;
  Received: string;
  Size: number;
}

interface TraceDetail {
  Event: string;
  Action: string;
  Detail: string;
  Date: string;
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function dateOffset(days: number): string {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d.toISOString().slice(0, 19);
}

export async function run(ps: PowerShellSession): Promise<void> {
  await ps.ensureExchangeConnected();

  // Sender
  const sender = await p.text({
    message: "Sender email (leave blank for any)",
    placeholder: "user@example.com",
    defaultValue: "",
  });
  if (p.isCancel(sender)) return;

  // Recipient
  const recipient = await p.text({
    message: "Recipient email (leave blank for any)",
    placeholder: "user@example.com",
    defaultValue: "",
  });
  if (p.isCancel(recipient)) return;

  if (!sender.trim() && !recipient.trim()) {
    p.log.error("At least one of sender or recipient is required.");
    return;
  }

  // Date range
  const dateRange = await p.select({
    message: "Date range",
    options: [
      { value: "1", label: "Last 24 hours" },
      { value: "2", label: "Last 48 hours" },
      { value: "7", label: "Last 7 days" },
      { value: "10", label: "Last 10 days" },
      { value: "custom", label: "Custom range" },
    ],
  });
  if (p.isCancel(dateRange)) return;

  let startDate: string;
  let endDate: string = new Date().toISOString().slice(0, 19);

  if (dateRange === "custom") {
    const customStart = await p.text({
      message: "Start date (YYYY-MM-DD)",
      placeholder: "2025-01-01",
      validate(value = "") {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(value.trim())) return "Use YYYY-MM-DD format";
        const d = new Date(value.trim());
        if (isNaN(d.getTime())) return "Invalid date";
        const daysDiff = (Date.now() - d.getTime()) / (1000 * 60 * 60 * 24);
        if (daysDiff > 10) return "Message trace only supports up to 10 days";
      },
    });
    if (p.isCancel(customStart)) return;

    const customEnd = await p.text({
      message: "End date (YYYY-MM-DD)",
      placeholder: new Date().toISOString().slice(0, 10),
      defaultValue: new Date().toISOString().slice(0, 10),
      validate(value = "") {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(value.trim())) return "Use YYYY-MM-DD format";
        if (isNaN(new Date(value.trim()).getTime())) return "Invalid date";
      },
    });
    if (p.isCancel(customEnd)) return;

    startDate = `${customStart.trim()}T00:00:00`;
    endDate = `${customEnd.trim()}T23:59:59`;
  } else {
    startDate = dateOffset(parseInt(dateRange));
  }

  // Status filter
  const statusFilter = await p.select({
    message: "Status filter",
    options: [
      { value: "all", label: "All" },
      { value: "Delivered", label: "Delivered" },
      { value: "Failed", label: "Failed" },
      { value: "Quarantined", label: "Quarantined" },
      { value: "Pending", label: "Pending" },
      { value: "FilteredAsSpam", label: "Filtered as spam" },
    ],
  });
  if (p.isCancel(statusFilter)) return;

  // Build command
  let cmd = "Get-MessageTraceV2";
  if (sender.trim()) cmd += ` -SenderAddress '${escapePS(sender.trim())}'`;
  if (recipient.trim()) cmd += ` -RecipientAddress '${escapePS(recipient.trim())}'`;
  cmd += ` -StartDate '${startDate}' -EndDate '${endDate}'`;
  if (statusFilter !== "all") cmd += ` -Status '${statusFilter}'`;
  cmd += " | Select-Object MessageId,MessageTraceId,SenderAddress,RecipientAddress,Subject,Status,Received,Size";

  const spin = p.spinner();
  spin.start("Running message trace...");

  let useV1 = false;
  let raw = await ps.runCommandJson<TraceMessage | TraceMessage[]>(cmd);

  // Check for V2 not recognized — fallback to V1
  if (raw === null) {
    const { error } = await ps.runCommand("Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue");
    if (error || !(await ps.runCommand("Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue")).output?.trim()) {
      useV1 = true;
      spin.message("Falling back to Get-MessageTrace (V1)...");
      const v1Cmd = cmd.replace("Get-MessageTraceV2", "Get-MessageTrace");
      raw = await ps.runCommandJson<TraceMessage | TraceMessage[]>(v1Cmd);
    }
  }

  const messages = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
  spin.stop(`Found ${messages.length} message(s).${useV1 ? " (using V1 — V2 not available)" : ""}`);

  if (useV1) {
    p.log.warn("Get-MessageTraceV2 is not available. Using deprecated Get-MessageTrace (V1).");
  }

  if (messages.length === 0) {
    p.log.info("No messages found matching your criteria.");
    return;
  }

  // Display table
  const header = `${"Sender".padEnd(28)} ${"Recipient".padEnd(28)} ${"Subject".padEnd(30)} ${"Status".padEnd(16)} Received`;
  const separator = "─".repeat(header.length + 5);
  const displayRows = messages.slice(0, 50);
  const lines = [header, separator];

  for (const m of displayRows) {
    const sender = truncate(m.SenderAddress ?? "", 27).padEnd(28);
    const recipient = truncate(m.RecipientAddress ?? "", 27).padEnd(28);
    const subject = truncate(m.Subject ?? "", 29).padEnd(30);
    const status = (m.Status ?? "").padEnd(16);
    const received = (m.Received ?? "").slice(0, 19);
    lines.push(`${sender} ${recipient} ${subject} ${status} ${received}`);
  }
  if (messages.length > 50) {
    lines.push(`… and ${messages.length - 50} more`);
  }

  p.note(lines.join("\n"), `Message Trace (${messages.length} result(s))`);

  // Drill into details or export
  while (true) {
    const nextAction = await p.select({
      message: "What next?",
      options: [
        { value: "details", label: "View message details", hint: "drill into a specific message" },
        { value: "export", label: "Export to Excel" },
        { value: "back", label: "Back" },
      ],
    });
    if (p.isCancel(nextAction) || nextAction === "back") break;

    if (nextAction === "details") {
      const msgOptions = messages.slice(0, 50).map((m, i) => ({
        value: String(i),
        label: truncate(m.Subject ?? "(no subject)", 50),
        hint: `${m.SenderAddress} → ${m.RecipientAddress}`,
      }));

      const selected = await p.select({
        message: "Select a message",
        options: [...msgOptions, { value: "back", label: "Back", hint: "" }],
      });
      if (p.isCancel(selected) || selected === "back") continue;

      const msg = messages[parseInt(selected)]!;

      const detailSpin = p.spinner();
      detailSpin.start("Fetching message details...");

      const detailCmdBase = useV1 ? "Get-MessageTraceDetail" : "Get-MessageTraceDetailV2";
      const detailCmd = `${detailCmdBase} -MessageTraceId '${escapePS(msg.MessageTraceId)}' -RecipientAddress '${escapePS(msg.RecipientAddress)}' | Select-Object Event,Action,Detail,Date`;

      const detailRaw = await ps.runCommandJson<TraceDetail | TraceDetail[]>(detailCmd);
      const details = detailRaw ? (Array.isArray(detailRaw) ? detailRaw : [detailRaw]) : [];
      detailSpin.stop(`Found ${details.length} event(s).`);

      if (details.length === 0) {
        p.log.info("No detail events found for this message.");
        continue;
      }

      const detailHeader = `${"Event".padEnd(22)} ${"Action".padEnd(22)} ${"Date".padEnd(20)} Detail`;
      const detailSep = "─".repeat(detailHeader.length + 30);
      const detailLines = [detailHeader, detailSep];

      for (const d of details) {
        const event = truncate(d.Event ?? "", 21).padEnd(22);
        const action = truncate(d.Action ?? "", 21).padEnd(22);
        const date = (d.Date ?? "").slice(0, 19).padEnd(20);
        const detail = truncate(d.Detail ?? "", 60);
        detailLines.push(`${event} ${action} ${date} ${detail}`);
      }

      p.note(detailLines.join("\n"), `Message Details: ${truncate(msg.Subject ?? "", 40)}`);
    }

    if (nextAction === "export") {
      const exportSpin = p.spinner();
      exportSpin.start("Generating Excel report...");

      const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
      const dateSlug = new Date().toISOString().slice(0, 10);
      const outputDir = join(appDir(), "reports output");
      const fullPath = resolve(join(outputDir, `${tenantSlug}-message-trace-${dateSlug}.xlsx`));
      mkdirSync(dirname(fullPath), { recursive: true });
      try { chmodSync(dirname(fullPath), 0o700); } catch {}

      const buffer = await generateReport({
        sheetName: "Message Trace",
        title: "Message Trace Report",
        tenant: ps.tenantDomain ?? "Unknown",
        summary: `${messages.length} message(s) found`,
        columns: [
          { header: "Sender", width: 30 },
          { header: "Recipient", width: 30 },
          { header: "Subject", width: 40 },
          { header: "Status", width: 16 },
          { header: "Received", width: 20 },
          { header: "Size", width: 12 },
        ],
        rows: messages.map((m) => [
          m.SenderAddress ?? "",
          m.RecipientAddress ?? "",
          m.Subject ?? "",
          m.Status ?? "",
          (m.Received ?? "").slice(0, 19),
          formatSize(m.Size ?? 0),
        ]),
      });

      await Bun.write(fullPath, buffer);
      try { chmodSync(fullPath, 0o600); } catch {}
      exportSpin.stop(`Exported ${messages.length} rows to ${fullPath}`);

      const folder = dirname(fullPath);
      try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
    }
  }
}
