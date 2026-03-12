import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { escapePS } from "../utils.ts";

interface QuarantineMessage {
  Identity: string;
  MessageId: string;
  SenderAddress: string;
  RecipientAddress: string;
  Subject: string;
  ReceivedTime: string;
  Type: string;
  ReleaseStatus: string;
}

interface PreviewResult {
  Subject: string;
  Body: string;
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function dateOffset(days: number): string {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d.toISOString().slice(0, 19);
}

async function listMessages(ps: PowerShellSession): Promise<QuarantineMessage[]> {
  // Type filter
  const typeFilter = await p.select({
    message: "Filter by type",
    options: [
      { value: "all", label: "All types" },
      { value: "Spam", label: "Spam" },
      { value: "HighConfidenceSpam", label: "High confidence spam" },
      { value: "Phish", label: "Phishing" },
      { value: "HighConfidencePhish", label: "High confidence phishing" },
      { value: "Malware", label: "Malware" },
      { value: "Bulk", label: "Bulk" },
    ],
  });
  if (p.isCancel(typeFilter)) return [];

  // Date range
  const dateRange = await p.select({
    message: "Date range",
    options: [
      { value: "1", label: "Last 24 hours" },
      { value: "3", label: "Last 3 days" },
      { value: "7", label: "Last 7 days" },
      { value: "14", label: "Last 14 days" },
      { value: "30", label: "Last 30 days" },
    ],
  });
  if (p.isCancel(dateRange)) return [];

  // Recipient filter
  const recipientInput = await p.text({
    message: "Filter by recipient email (leave blank for all)",
    placeholder: "user@example.com",
    defaultValue: "",
  });
  if (p.isCancel(recipientInput)) return [];

  // Build command
  let cmd = "Get-QuarantineMessage";
  if (typeFilter !== "all") cmd += ` -Type '${typeFilter}'`;
  cmd += ` -StartExpiresDate '${dateOffset(parseInt(dateRange))}'`;
  cmd += ` -EndExpiresDate '${new Date().toISOString().slice(0, 19)}'`;
  if (recipientInput.trim()) cmd += ` -RecipientAddress '${escapePS(recipientInput.trim())}'`;
  cmd += " -PageSize 100";
  cmd += " | Select-Object Identity,MessageId,SenderAddress,RecipientAddress,Subject,ReceivedTime,Type,ReleaseStatus";

  const spin = p.spinner();
  spin.start("Fetching quarantined messages...");

  const raw = await ps.runCommandJson<QuarantineMessage | QuarantineMessage[]>(cmd);
  const messages = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
  spin.stop(`Found ${messages.length} quarantined message(s).`);

  if (messages.length === 0) {
    p.log.info("No quarantined messages found matching your criteria.");
    return [];
  }

  // Display table
  const header = `${"Sender".padEnd(28)} ${"Recipient".padEnd(28)} ${"Subject".padEnd(30)} ${"Type".padEnd(12)} ${"Status".padEnd(14)} Received`;
  const separator = "─".repeat(header.length + 5);
  const displayRows = messages.slice(0, 50);
  const lines = [header, separator];

  for (const m of displayRows) {
    const sender = truncate(m.SenderAddress ?? "", 27).padEnd(28);
    const recipient = truncate(m.RecipientAddress ?? "", 27).padEnd(28);
    const subject = truncate(m.Subject ?? "", 29).padEnd(30);
    const type = truncate(m.Type ?? "", 11).padEnd(12);
    const status = (m.ReleaseStatus ?? "").padEnd(14);
    const received = (m.ReceivedTime ?? "").slice(0, 19);
    lines.push(`${sender} ${recipient} ${subject} ${type} ${status} ${received}`);
  }
  if (messages.length > 50) {
    lines.push(`… and ${messages.length - 50} more`);
  }

  p.note(lines.join("\n"), `Quarantine (${messages.length} message(s))`);

  return messages;
}

async function previewMessage(ps: PowerShellSession, messages: QuarantineMessage[]): Promise<void> {
  if (messages.length === 0) {
    p.log.info("No messages to preview. List messages first.");
    return;
  }

  const msgOptions = messages.slice(0, 50).map((m, i) => ({
    value: String(i),
    label: truncate(m.Subject ?? "(no subject)", 50),
    hint: `${m.SenderAddress} → ${m.RecipientAddress}`,
  }));

  const selected = await p.select({
    message: "Select a message to preview",
    options: [...msgOptions, { value: "back", label: "Back", hint: "" }],
  });
  if (p.isCancel(selected) || selected === "back") return;

  const msg = messages[parseInt(selected)]!;
  const spin = p.spinner();
  spin.start("Fetching message preview...");

  const previewRaw = await ps.runCommandJson<PreviewResult>(
    `Preview-QuarantineMessage -Identity '${escapePS(msg.Identity)}' | Select-Object Subject,Body`
  );
  spin.stop("Preview loaded.");

  if (!previewRaw) {
    p.log.error("Could not preview this message.");
    return;
  }

  const body = stripHtml(previewRaw.Body ?? "");
  const previewText = body.length > 2000 ? body.slice(0, 2000) + "\n\n… (truncated)" : body;

  p.note(previewText, `Preview: ${truncate(previewRaw.Subject ?? msg.Subject ?? "", 50)}`);
}

async function releaseMessages(ps: PowerShellSession, messages: QuarantineMessage[]): Promise<void> {
  const unreleased = messages.filter((m) =>
    m.ReleaseStatus !== "Released" && m.ReleaseStatus !== "Approved"
  );

  if (unreleased.length === 0) {
    p.log.info("No unreleased messages available. List messages first.");
    return;
  }

  const options = unreleased.slice(0, 50).map((m, i) => ({
    value: String(i),
    label: truncate(m.Subject ?? "(no subject)", 50),
    hint: `${m.SenderAddress} → ${m.RecipientAddress} [${m.Type}]`,
  }));

  const selected = await p.multiselect({
    message: "Select messages to release",
    options,
    required: true,
  });
  if (p.isCancel(selected)) return;

  const toRelease = selected.map((i) => unreleased[parseInt(i)]!);

  const confirm = await p.confirm({
    message: `Release ${toRelease.length} message(s) to all original recipients?`,
  });
  if (p.isCancel(confirm) || !confirm) {
    p.log.info("Cancelled.");
    return;
  }

  const spin = p.spinner();
  spin.start("Releasing messages...");

  let successCount = 0;
  for (let i = 0; i < toRelease.length; i++) {
    const msg = toRelease[i]!;
    spin.message(`Releasing message ${i + 1}/${toRelease.length}...`);

    const { error } = await ps.runCommand(
      `Release-QuarantineMessage -Identity '${escapePS(msg.Identity)}' -ReleaseToAll`
    );
    if (error) {
      p.log.error(`Failed to release "${truncate(msg.Subject ?? "", 40)}": ${error}`);
    } else {
      successCount++;
    }

    // Throttle between operations
    if (i < toRelease.length - 1) await Bun.sleep(500);
  }

  spin.stop(`Released ${successCount}/${toRelease.length} message(s).`);
}

async function deleteMessages(ps: PowerShellSession, messages: QuarantineMessage[]): Promise<void> {
  if (messages.length === 0) {
    p.log.info("No messages to delete. List messages first.");
    return;
  }

  const options = messages.slice(0, 50).map((m, i) => ({
    value: String(i),
    label: truncate(m.Subject ?? "(no subject)", 50),
    hint: `${m.SenderAddress} [${m.Type}]`,
  }));

  const selected = await p.multiselect({
    message: "Select messages to delete",
    options,
    required: true,
  });
  if (p.isCancel(selected)) return;

  const toDelete = selected.map((i) => messages[parseInt(i)]!);

  // Double confirm
  const confirm1 = await p.confirm({
    message: `Delete ${toDelete.length} quarantined message(s)?`,
  });
  if (p.isCancel(confirm1) || !confirm1) {
    p.log.info("Cancelled.");
    return;
  }

  const confirm2 = await p.confirm({
    message: `Are you sure? This action cannot be undone.`,
    initialValue: false,
  });
  if (p.isCancel(confirm2) || !confirm2) {
    p.log.info("Cancelled.");
    return;
  }

  const spin = p.spinner();
  spin.start("Deleting messages...");

  let successCount = 0;
  for (let i = 0; i < toDelete.length; i++) {
    const msg = toDelete[i]!;
    spin.message(`Deleting message ${i + 1}/${toDelete.length}...`);

    const { error } = await ps.runCommand(
      `Delete-QuarantineMessage -Identity '${escapePS(msg.Identity)}' -Confirm:$false`
    );
    if (error) {
      p.log.error(`Failed to delete "${truncate(msg.Subject ?? "", 40)}": ${error}`);
    } else {
      successCount++;
    }

    if (i < toDelete.length - 1) await Bun.sleep(500);
  }

  spin.stop(`Deleted ${successCount}/${toDelete.length} message(s).`);
}

export async function run(ps: PowerShellSession): Promise<void> {
  let cachedMessages: QuarantineMessage[] = [];

  while (true) {
    const action = await p.select({
      message: "Quarantine Management",
      options: [
        { value: "list", label: "List quarantined messages" },
        { value: "preview", label: "Preview a message", hint: cachedMessages.length > 0 ? `${cachedMessages.length} loaded` : "list first" },
        { value: "release", label: "Release message(s)", hint: cachedMessages.length > 0 ? `${cachedMessages.length} loaded` : "list first" },
        { value: "delete", label: "Delete message(s)", hint: cachedMessages.length > 0 ? `${cachedMessages.length} loaded` : "list first" },
        { value: "back", label: "Back" },
      ],
    });

    if (p.isCancel(action) || action === "back") return;

    switch (action) {
      case "list":
        cachedMessages = await listMessages(ps);
        break;
      case "preview":
        await previewMessage(ps, cachedMessages);
        break;
      case "release":
        await releaseMessages(ps, cachedMessages);
        break;
      case "delete":
        await deleteMessages(ps, cachedMessages);
        break;
    }
  }
}
