import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { escapePS } from "../utils.ts";

interface DistributionGroup {
  DisplayName: string;
  PrimarySmtpAddress: string;
}

interface SecurityGroup {
  DisplayName: string;
  Id: string;
}

interface SharedMailbox {
  DisplayName: string;
  PrimarySmtpAddress: string;
}

interface DistMember {
  DisplayName: string;
  PrimarySmtpAddress: string;
}

interface GroupMember {
  DisplayName: string;
  UserPrincipalName: string;
}

async function deleteDistributionGroup(ps: PowerShellSession): Promise<void> {
  const fetchSpin = p.spinner();
  fetchSpin.start("Fetching distribution groups...");
  let groups: DistributionGroup[];
  try {
    const raw = await ps.runCommandJson<DistributionGroup | DistributionGroup[]>(
      "Get-DistributionGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress",
    );
    groups = (raw ? (Array.isArray(raw) ? raw : [raw]) : []).sort((a, b) =>
      a.DisplayName.localeCompare(b.DisplayName),
    );
    fetchSpin.stop(`Found ${groups.length} distribution group(s).`);
  } catch (e) {
    fetchSpin.stop("Failed to fetch distribution groups.");
    p.log.error(`${e}`);
    return;
  }

  if (groups.length === 0) {
    p.log.warn("No distribution groups found.");
    return;
  }

  const selected = await p.select({
    message: "Select distribution group to delete",
    options: groups.map((g) => ({
      value: g.PrimarySmtpAddress,
      label: g.DisplayName,
      hint: g.PrimarySmtpAddress,
    })),
  });
  if (p.isCancel(selected)) return;

  const group = groups.find((g) => g.PrimarySmtpAddress === selected)!;

  // Fetch details
  const detailSpin = p.spinner();
  detailSpin.start("Fetching group details...");

  let members: DistMember[] = [];
  let owners: string[] = [];

  try {
    const rawMembers = await ps.runCommandJson<DistMember | DistMember[]>(
      `Get-DistributionGroupMember -Identity '${escapePS(selected)}' -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress`,
    );
    members = rawMembers ? (Array.isArray(rawMembers) ? rawMembers : [rawMembers]) : [];
  } catch { /* empty group */ }

  try {
    const { output } = await ps.runCommand(
      `Get-DistributionGroup -Identity '${escapePS(selected)}' | Select-Object -ExpandProperty ManagedBy`,
    );
    if (output.trim()) {
      const dns = output.trim().split("\n").map((l) => l.trim()).filter(Boolean);
      for (const dn of dns) {
        const { output: recipientOutput } = await ps.runCommand(
          `Get-EXORecipient '${escapePS(dn)}' -Properties DisplayName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName`,
        );
        owners.push(recipientOutput.trim() || dn);
      }
    }
  } catch { /* no owners */ }

  detailSpin.stop("Group details loaded.");

  const details = [
    `Name:    ${group.DisplayName}`,
    `Email:   ${group.PrimarySmtpAddress}`,
    `Members: ${members.length > 0 ? members.map((m) => m.DisplayName).join(", ") : "(none)"}`,
    `Owners:  ${owners.length > 0 ? owners.join(", ") : "(none)"}`,
  ];
  p.note(details.join("\n"), "Distribution group details");

  const ok = await p.confirm({
    message: `Delete distribution group "${group.DisplayName}"? This cannot be undone.`,
  });
  if (p.isCancel(ok) || !ok) {
    p.log.info("Cancelled.");
    return;
  }

  const delSpin = p.spinner();
  delSpin.start("Deleting distribution group...");
  const { error } = await ps.runCommand(
    `Remove-DistributionGroup -Identity '${escapePS(selected)}' -Confirm:$false`,
  );
  if (error) {
    delSpin.stop("Failed to delete distribution group.");
    p.log.error(error);
    return;
  }
  delSpin.stop("Distribution group deleted.");
  p.log.success(`Deleted distribution group "${group.DisplayName}".`);
}

async function deleteSecurityGroup(ps: PowerShellSession): Promise<void> {
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph (check your browser)...");
  try {
    await ps.ensureGraphConnected();
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return;
  }

  const fetchSpin = p.spinner();
  fetchSpin.start("Fetching security groups...");
  let groups: SecurityGroup[];
  try {
    const raw = await ps.runCommandJson<SecurityGroup | SecurityGroup[]>(
      'Get-MgGroup -Filter "securityEnabled eq true and mailEnabled eq false" -All | Select-Object DisplayName, Id',
    );
    groups = (raw ? (Array.isArray(raw) ? raw : [raw]) : []).sort((a, b) =>
      a.DisplayName.localeCompare(b.DisplayName),
    );
    fetchSpin.stop(`Found ${groups.length} security group(s).`);
  } catch (e) {
    fetchSpin.stop("Failed to fetch security groups.");
    p.log.error(`${e}`);
    return;
  }

  if (groups.length === 0) {
    p.log.warn("No security groups found.");
    return;
  }

  const selectedId = await p.select({
    message: "Select security group to delete",
    options: groups.map((g) => ({
      value: g.Id,
      label: g.DisplayName,
    })),
  });
  if (p.isCancel(selectedId)) return;

  const group = groups.find((g) => g.Id === selectedId)!;

  // Fetch details
  const detailSpin = p.spinner();
  detailSpin.start("Fetching group details...");

  let members: GroupMember[] = [];
  let owners: GroupMember[] = [];

  try {
    const rawMembers = await ps.runCommandJson<Record<string, unknown> | Record<string, unknown>[]>(
      `Get-MgGroupMember -GroupId '${escapePS(selectedId)}' -All`,
    );
    const rawArr = rawMembers ? (Array.isArray(rawMembers) ? rawMembers : [rawMembers]) : [];
    members = rawArr.map((m) => ({
      DisplayName: (m.AdditionalProperties as Record<string, string>)?.displayName ?? "(unknown)",
      UserPrincipalName: (m.AdditionalProperties as Record<string, string>)?.userPrincipalName ?? "",
    }));
  } catch { /* no members */ }

  try {
    const rawOwners = await ps.runCommandJson<Record<string, unknown> | Record<string, unknown>[]>(
      `Get-MgGroupOwner -GroupId '${escapePS(selectedId)}' -All`,
    );
    const rawArr = rawOwners ? (Array.isArray(rawOwners) ? rawOwners : [rawOwners]) : [];
    owners = rawArr.map((m) => ({
      DisplayName: (m.AdditionalProperties as Record<string, string>)?.displayName ?? "(unknown)",
      UserPrincipalName: (m.AdditionalProperties as Record<string, string>)?.userPrincipalName ?? "",
    }));
  } catch { /* no owners */ }

  detailSpin.stop("Group details loaded.");

  const details = [
    `Name:    ${group.DisplayName}`,
    `Members: ${members.length > 0 ? members.map((m) => m.DisplayName).join(", ") : "(none)"}`,
    `Owners:  ${owners.length > 0 ? owners.map((m) => m.DisplayName).join(", ") : "(none)"}`,
  ];
  p.note(details.join("\n"), "Security group details");

  const ok = await p.confirm({
    message: `Delete security group "${group.DisplayName}"? This cannot be undone.`,
  });
  if (p.isCancel(ok) || !ok) {
    p.log.info("Cancelled.");
    return;
  }

  const delSpin = p.spinner();
  delSpin.start("Deleting security group...");
  const { error } = await ps.runCommand(
    `Remove-MgGroup -GroupId '${escapePS(selectedId)}' -Confirm:$false`,
  );
  if (error) {
    delSpin.stop("Failed to delete security group.");
    p.log.error(error);
    return;
  }
  delSpin.stop("Security group deleted.");
  p.log.success(`Deleted security group "${group.DisplayName}".`);
}

async function deleteSharedMailbox(ps: PowerShellSession): Promise<void> {
  const fetchSpin = p.spinner();
  fetchSpin.start("Fetching shared mailboxes...");
  let mailboxes: SharedMailbox[];
  try {
    const raw = await ps.runCommandJson<SharedMailbox | SharedMailbox[]>(
      "Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress",
    );
    mailboxes = (raw ? (Array.isArray(raw) ? raw : [raw]) : []).sort((a, b) =>
      a.DisplayName.localeCompare(b.DisplayName),
    );
    fetchSpin.stop(`Found ${mailboxes.length} shared mailbox(es).`);
  } catch (e) {
    fetchSpin.stop("Failed to fetch shared mailboxes.");
    p.log.error(`${e}`);
    return;
  }

  if (mailboxes.length === 0) {
    p.log.warn("No shared mailboxes found.");
    return;
  }

  const selected = await p.select({
    message: "Select shared mailbox to delete",
    options: mailboxes.map((m) => ({
      value: m.PrimarySmtpAddress,
      label: m.DisplayName,
      hint: m.PrimarySmtpAddress,
    })),
  });
  if (p.isCancel(selected)) return;

  const mailbox = mailboxes.find((m) => m.PrimarySmtpAddress === selected)!;

  // Fetch details
  const detailSpin = p.spinner();
  detailSpin.start("Fetching mailbox details...");

  const permHolders: string[] = [];

  try {
    const { output: faOutput } = await ps.runCommand(
      `Get-EXOMailboxPermission -Identity '${escapePS(selected)}' | Where-Object { $_.User -ne 'NT AUTHORITY\\SELF' -and $_.User -notmatch '^S-1-' -and $_.AccessRights -contains 'FullAccess' } | Select-Object -ExpandProperty User`,
    );
    if (faOutput.trim()) {
      for (const user of faOutput.trim().split("\n").map((l) => l.trim()).filter(Boolean)) {
        if (!permHolders.includes(user)) permHolders.push(user);
      }
    }
  } catch { /* no perms */ }

  try {
    const { output: saOutput } = await ps.runCommand(
      `Get-EXORecipientPermission -Identity '${escapePS(selected)}' | Where-Object { $_.Trustee -ne 'NT AUTHORITY\\SELF' -and $_.Trustee -notmatch '^S-1-' } | Select-Object -ExpandProperty Trustee`,
    );
    if (saOutput.trim()) {
      for (const user of saOutput.trim().split("\n").map((l) => l.trim()).filter(Boolean)) {
        if (!permHolders.includes(user)) permHolders.push(user);
      }
    }
  } catch { /* no perms */ }

  try {
    const { output: sobOutput } = await ps.runCommand(
      `Get-Mailbox -Identity '${escapePS(selected)}' | Select-Object -ExpandProperty GrantSendOnBehalfTo`,
    );
    if (sobOutput.trim()) {
      const dns = sobOutput.trim().split("\n").map((l) => l.trim()).filter(Boolean);
      for (const dn of dns) {
        const { output: recipientOutput } = await ps.runCommand(
          `Get-EXORecipient '${escapePS(dn)}' -Properties DisplayName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName`,
        );
        const name = recipientOutput.trim() || dn;
        if (!permHolders.includes(name)) permHolders.push(name);
      }
    }
  } catch { /* no perms */ }

  detailSpin.stop("Mailbox details loaded.");

  const details = [
    `Name:              ${mailbox.DisplayName}`,
    `Email:             ${mailbox.PrimarySmtpAddress}`,
    `Permission holders: ${permHolders.length > 0 ? permHolders.join(", ") : "(none)"}`,
  ];
  p.note(details.join("\n"), "Shared mailbox details");

  const ok = await p.confirm({
    message: `Delete shared mailbox "${mailbox.DisplayName}"? This cannot be undone.`,
  });
  if (p.isCancel(ok) || !ok) {
    p.log.info("Cancelled.");
    return;
  }

  const delSpin = p.spinner();
  delSpin.start("Deleting shared mailbox...");
  const { error } = await ps.runCommand(
    `Remove-Mailbox -Identity '${escapePS(selected)}' -Confirm:$false`,
  );
  if (error) {
    delSpin.stop("Failed to delete shared mailbox.");
    p.log.error(error);
    return;
  }
  delSpin.stop("Shared mailbox deleted.");
  p.log.success(`Deleted shared mailbox "${mailbox.DisplayName}".`);
}

export async function run(ps: PowerShellSession): Promise<void> {
  const type = await p.select({
    message: "What would you like to delete?",
    options: [
      { value: "distribution", label: "Distribution group" },
      { value: "security", label: "Security group" },
      { value: "shared-mailbox", label: "Shared mailbox" },
    ],
  });
  if (p.isCancel(type)) return;

  switch (type) {
    case "distribution":
      await deleteDistributionGroup(ps);
      break;
    case "security":
      await deleteSecurityGroup(ps);
      break;
    case "shared-mailbox":
      await deleteSharedMailbox(ps);
      break;
  }
}
