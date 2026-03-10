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

interface MgUser {
  DisplayName: string;
  UserPrincipalName: string;
  Id: string;
}

async function fetchUsers(ps: PowerShellSession): Promise<MgUser[]> {
  const { output: countOutput } = await ps.runCommand(
    "Get-MgUser -Top 1 -CountVariable ct -ConsistencyLevel eventual | Out-Null; $ct",
  );
  const count = parseInt(countOutput.trim(), 10);

  let users: MgUser[];

  if (count <= 50) {
    const raw = await ps.runCommandJson<MgUser | MgUser[]>(
      "Get-MgUser -Filter 'accountEnabled eq true' -All -Property DisplayName,UserPrincipalName,Id,AssignedLicenses -ConsistencyLevel eventual | Where-Object { $_.AssignedLicenses.Count -gt 0 } | Select-Object DisplayName,UserPrincipalName,Id",
    );
    users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
  } else {
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
          `Get-MgUser -Search '"displayName:${escapePS(query)}"' -ConsistencyLevel eventual -Property DisplayName,UserPrincipalName,Id,AssignedLicenses | Where-Object { $_.AssignedLicenses.Count -gt 0 } | Select-Object DisplayName,UserPrincipalName,Id`,
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

async function selectMultipleUsers(
  ps: PowerShellSession,
  message: string,
): Promise<MgUser[]> {
  while (true) {
    const users = await fetchUsers(ps);

    if (users.length === 0) {
      const retry = await p.confirm({ message: "No users found. Search again?" });
      if (p.isCancel(retry) || !retry) return [];
      continue;
    }

    const selected = await p.multiselect({
      message: `${message} (space to select, enter to confirm)`,
      options: users.map((u) => ({
        value: u.Id,
        label: u.DisplayName,
        hint: u.UserPrincipalName,
      })),
      required: true,
    });
    if (p.isCancel(selected)) return [];

    return users.filter((u) => selected.includes(u.Id));
  }
}

// ── Distribution Group ──────────────────────────────────────────────────

interface DistOwner {
  displayName: string;
  identity: string;
}

async function fetchDistGroupDetails(
  ps: PowerShellSession,
  email: string,
): Promise<{ members: DistMember[]; owners: DistOwner[] }> {
  let members: DistMember[] = [];
  let owners: DistOwner[] = [];

  try {
    const rawMembers = await ps.runCommandJson<DistMember | DistMember[]>(
      `Get-DistributionGroupMember -Identity '${escapePS(email)}' -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress`,
    );
    members = rawMembers ? (Array.isArray(rawMembers) ? rawMembers : [rawMembers]) : [];
  } catch { /* empty */ }

  try {
    const { output } = await ps.runCommand(
      `Get-DistributionGroup -Identity '${escapePS(email)}' | Select-Object -ExpandProperty ManagedBy`,
    );
    if (output.trim()) {
      const dns = output.trim().split("\n").map((l) => l.trim()).filter(Boolean);
      for (const dn of dns) {
        const { output: recipientOutput } = await ps.runCommand(
          `Get-EXORecipient '${escapePS(dn)}' -Properties DisplayName, PrimarySmtpAddress -ErrorAction SilentlyContinue | Select-Object DisplayName, PrimarySmtpAddress | ForEach-Object { "$($_.DisplayName)|$($_.PrimarySmtpAddress)" }`,
        );
        const parts = recipientOutput.trim().split("|");
        const displayName = parts[0] || dn;
        const identity = parts[1] || dn;
        owners.push({ displayName, identity });
      }
    }
  } catch { /* no owners */ }

  return { members, owners };
}

function showDistGroupDetails(
  displayName: string,
  email: string,
  members: DistMember[],
  owners: DistOwner[],
): void {
  p.note(
    [
      `Name:    ${displayName}`,
      `Email:   ${email}`,
      `Members: ${members.length > 0 ? members.map((m) => m.DisplayName).join(", ") : "(none)"}`,
      `Owners:  ${owners.length > 0 ? owners.map((o) => o.displayName).join(", ") : "(none)"}`,
    ].join("\n"),
    "Distribution group details",
  );
}

async function editDistributionGroup(ps: PowerShellSession): Promise<void> {
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

  const selectedEmail = await p.select({
    message: "Select distribution group to edit",
    options: groups.map((g) => ({
      value: g.PrimarySmtpAddress,
      label: g.DisplayName,
      hint: g.PrimarySmtpAddress,
    })),
  });
  if (p.isCancel(selectedEmail)) return;

  let currentName = groups.find((g) => g.PrimarySmtpAddress === selectedEmail)!.DisplayName;

  // Ensure Graph for user fetching
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph...");
  try {
    await ps.ensureGraphConnected();
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return;
  }

  // Show current details
  const detailSpin = p.spinner();
  detailSpin.start("Fetching group details...");
  let { members, owners } = await fetchDistGroupDetails(ps, selectedEmail);
  detailSpin.stop("Group details loaded.");
  showDistGroupDetails(currentName, selectedEmail, members, owners);

  // Edit loop
  while (true) {
    const action = await p.select({
      message: "Edit action",
      options: [
        { value: "rename", label: "Rename" },
        { value: "add-members", label: "Add members" },
        { value: "remove-members", label: "Remove members" },
        { value: "add-owners", label: "Add owners" },
        { value: "remove-owners", label: "Remove owners" },
        { value: "done", label: "Done" },
      ],
    });
    if (p.isCancel(action) || action === "done") break;

    if (action === "rename") {
      const newName = await p.text({
        message: "New display name",
        initialValue: currentName,
        validate: (v = "") => (!v.trim() ? "Display name is required" : undefined),
      });
      if (p.isCancel(newName)) continue;

      const spin = p.spinner();
      spin.start("Renaming...");
      const { error } = await ps.runCommand(
        `Set-DistributionGroup -Identity '${escapePS(selectedEmail)}' -DisplayName '${escapePS(newName)}'`,
      );
      if (error) {
        spin.stop("Failed to rename.");
        p.log.error(error);
      } else {
        spin.stop(`Renamed to "${newName}".`);
        currentName = newName;
      }
    }

    if (action === "add-members") {
      const newMembers = await selectMultipleUsers(ps, "Select members to add");
      for (const member of newMembers) {
        const spin = p.spinner();
        spin.start(`Adding ${member.DisplayName}...`);
        const { error } = await ps.runCommand(
          `Add-DistributionGroupMember -Identity '${escapePS(selectedEmail)}' -Member '${escapePS(member.UserPrincipalName)}'`,
        );
        if (error) {
          spin.stop(`Failed to add ${member.DisplayName}.`);
          p.log.error(error);
        } else {
          spin.stop(`Added ${member.DisplayName}.`);
        }
      }
    }

    if (action === "remove-members") {
      if (members.length === 0) {
        p.log.warn("No members to remove.");
        continue;
      }
      const toRemove = await p.multiselect({
        message: "Select members to remove (space to select, enter to confirm)",
        options: members.map((m) => ({
          value: m.PrimarySmtpAddress,
          label: m.DisplayName,
          hint: m.PrimarySmtpAddress,
        })),
        required: true,
      });
      if (p.isCancel(toRemove)) continue;

      for (const email of toRemove) {
        const name = members.find((m) => m.PrimarySmtpAddress === email)?.DisplayName ?? email;
        const spin = p.spinner();
        spin.start(`Removing ${name}...`);
        const { error } = await ps.runCommand(
          `Remove-DistributionGroupMember -Identity '${escapePS(selectedEmail)}' -Member '${escapePS(email)}' -Confirm:$false`,
        );
        if (error) {
          spin.stop(`Failed to remove ${name}.`);
          p.log.error(error);
        } else {
          spin.stop(`Removed ${name}.`);
        }
      }
    }

    if (action === "add-owners") {
      const newOwners = await selectMultipleUsers(ps, "Select owners to add");
      for (const owner of newOwners) {
        const spin = p.spinner();
        spin.start(`Adding ${owner.DisplayName} as owner...`);
        const { error } = await ps.runCommand(
          `Set-DistributionGroup -Identity '${escapePS(selectedEmail)}' -ManagedBy @{Add='${escapePS(owner.UserPrincipalName)}'}`,
        );
        if (error) {
          spin.stop(`Failed to add ${owner.DisplayName} as owner.`);
          p.log.error(error);
        } else {
          spin.stop(`Added ${owner.DisplayName} as owner.`);
        }
      }
    }

    if (action === "remove-owners") {
      if (owners.length === 0) {
        p.log.warn("No owners to remove.");
        continue;
      }
      const toRemove = await p.multiselect({
        message: "Select owners to remove (space to select, enter to confirm)",
        options: owners.map((o) => ({
          value: o.identity,
          label: o.displayName,
          hint: o.identity,
        })),
        required: true,
      });
      if (p.isCancel(toRemove)) continue;

      for (const ownerIdentity of toRemove) {
        const ownerName = owners.find((o) => o.identity === ownerIdentity)?.displayName ?? ownerIdentity;
        const spin = p.spinner();
        spin.start(`Removing ${ownerName} as owner...`);
        const { error } = await ps.runCommand(
          `Set-DistributionGroup -Identity '${escapePS(selectedEmail)}' -ManagedBy @{Remove='${escapePS(ownerIdentity)}'}`,
        );
        if (error) {
          spin.stop(`Failed to remove ${ownerName} as owner.`);
          p.log.error(error);
        } else {
          spin.stop(`Removed ${ownerName} as owner.`);
        }
      }
    }

    // Refresh details after each action
    const refreshSpin = p.spinner();
    refreshSpin.start("Refreshing details...");
    ({ members, owners } = await fetchDistGroupDetails(ps, selectedEmail));
    refreshSpin.stop("Details refreshed.");
    showDistGroupDetails(currentName, selectedEmail, members, owners);
  }
}

// ── Security Group ──────────────────────────────────────────────────────

interface GraphMember {
  Id: string;
  DisplayName: string;
  UserPrincipalName: string;
}

async function fetchSecGroupDetails(
  ps: PowerShellSession,
  groupId: string,
): Promise<{ members: GraphMember[]; owners: GraphMember[] }> {
  let members: GraphMember[] = [];
  let owners: GraphMember[] = [];

  try {
    const rawMembers = await ps.runCommandJson<Record<string, unknown> | Record<string, unknown>[]>(
      `Get-MgGroupMember -GroupId '${escapePS(groupId)}' -All`,
    );
    const rawArr = rawMembers ? (Array.isArray(rawMembers) ? rawMembers : [rawMembers]) : [];
    members = rawArr.map((m) => ({
      Id: m.Id as string,
      DisplayName: (m.AdditionalProperties as Record<string, string>)?.displayName ?? "(unknown)",
      UserPrincipalName: (m.AdditionalProperties as Record<string, string>)?.userPrincipalName ?? "",
    }));
  } catch { /* no members */ }

  try {
    const rawOwners = await ps.runCommandJson<Record<string, unknown> | Record<string, unknown>[]>(
      `Get-MgGroupOwner -GroupId '${escapePS(groupId)}' -All`,
    );
    const rawArr = rawOwners ? (Array.isArray(rawOwners) ? rawOwners : [rawOwners]) : [];
    owners = rawArr.map((m) => ({
      Id: m.Id as string,
      DisplayName: (m.AdditionalProperties as Record<string, string>)?.displayName ?? "(unknown)",
      UserPrincipalName: (m.AdditionalProperties as Record<string, string>)?.userPrincipalName ?? "",
    }));
  } catch { /* no owners */ }

  return { members, owners };
}

function showSecGroupDetails(
  displayName: string,
  members: GraphMember[],
  owners: GraphMember[],
): void {
  p.note(
    [
      `Name:    ${displayName}`,
      `Members: ${members.length > 0 ? members.map((m) => m.DisplayName).join(", ") : "(none)"}`,
      `Owners:  ${owners.length > 0 ? owners.map((m) => m.DisplayName).join(", ") : "(none)"}`,
    ].join("\n"),
    "Security group details",
  );
}

async function editSecurityGroup(ps: PowerShellSession): Promise<void> {
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
    message: "Select security group to edit",
    options: groups.map((g) => ({
      value: g.Id,
      label: g.DisplayName,
    })),
  });
  if (p.isCancel(selectedId)) return;

  let currentName = groups.find((g) => g.Id === selectedId)!.DisplayName;

  // Show current details
  const detailSpin = p.spinner();
  detailSpin.start("Fetching group details...");
  let { members, owners } = await fetchSecGroupDetails(ps, selectedId);
  detailSpin.stop("Group details loaded.");
  showSecGroupDetails(currentName, members, owners);

  // Edit loop
  while (true) {
    const action = await p.select({
      message: "Edit action",
      options: [
        { value: "rename", label: "Rename" },
        { value: "add-members", label: "Add members" },
        { value: "remove-members", label: "Remove members" },
        { value: "add-owners", label: "Add owners" },
        { value: "remove-owners", label: "Remove owners" },
        { value: "done", label: "Done" },
      ],
    });
    if (p.isCancel(action) || action === "done") break;

    if (action === "rename") {
      const newName = await p.text({
        message: "New display name",
        initialValue: currentName,
        validate: (v = "") => (!v.trim() ? "Display name is required" : undefined),
      });
      if (p.isCancel(newName)) continue;

      const spin = p.spinner();
      spin.start("Renaming...");
      const { error } = await ps.runCommand(
        `Update-MgGroup -GroupId '${escapePS(selectedId)}' -DisplayName '${escapePS(newName)}'`,
      );
      if (error) {
        spin.stop("Failed to rename.");
        p.log.error(error);
      } else {
        spin.stop(`Renamed to "${newName}".`);
        currentName = newName;
      }
    }

    if (action === "add-members") {
      const newMembers = await selectMultipleUsers(ps, "Select members to add");
      for (const member of newMembers) {
        const spin = p.spinner();
        spin.start(`Adding ${member.DisplayName}...`);
        const { error } = await ps.runCommand(
          `New-MgGroupMemberByRef -GroupId '${escapePS(selectedId)}' -OdataId 'https://graph.microsoft.com/v1.0/directoryObjects/${escapePS(member.Id)}'`,
        );
        if (error) {
          spin.stop(`Failed to add ${member.DisplayName}.`);
          p.log.error(error);
        } else {
          spin.stop(`Added ${member.DisplayName}.`);
        }
      }
    }

    if (action === "remove-members") {
      if (members.length === 0) {
        p.log.warn("No members to remove.");
        continue;
      }
      const toRemove = await p.multiselect({
        message: "Select members to remove (space to select, enter to confirm)",
        options: members.map((m) => ({
          value: m.Id,
          label: m.DisplayName,
          hint: m.UserPrincipalName,
        })),
        required: true,
      });
      if (p.isCancel(toRemove)) continue;

      for (const memberId of toRemove) {
        const name = members.find((m) => m.Id === memberId)?.DisplayName ?? memberId;
        const spin = p.spinner();
        spin.start(`Removing ${name}...`);
        const { error } = await ps.runCommand(
          `Remove-MgGroupMemberByRef -GroupId '${escapePS(selectedId)}' -DirectoryObjectId '${escapePS(memberId)}'`,
        );
        if (error) {
          spin.stop(`Failed to remove ${name}.`);
          p.log.error(error);
        } else {
          spin.stop(`Removed ${name}.`);
        }
      }
    }

    if (action === "add-owners") {
      const newOwners = await selectMultipleUsers(ps, "Select owners to add");
      for (const owner of newOwners) {
        const spin = p.spinner();
        spin.start(`Adding ${owner.DisplayName} as owner...`);
        const { error } = await ps.runCommand(
          `New-MgGroupOwnerByRef -GroupId '${escapePS(selectedId)}' -OdataId 'https://graph.microsoft.com/v1.0/directoryObjects/${escapePS(owner.Id)}'`,
        );
        if (error) {
          spin.stop(`Failed to add ${owner.DisplayName} as owner.`);
          p.log.error(error);
        } else {
          spin.stop(`Added ${owner.DisplayName} as owner.`);
        }
      }
    }

    if (action === "remove-owners") {
      if (owners.length === 0) {
        p.log.warn("No owners to remove.");
        continue;
      }
      if (owners.length === 1) {
        p.log.warn("Cannot remove the last owner.");
        continue;
      }
      const toRemove = await p.multiselect({
        message: "Select owners to remove (space to select, enter to confirm)",
        options: owners.map((o) => ({
          value: o.Id,
          label: o.DisplayName,
          hint: o.UserPrincipalName,
        })),
        required: true,
      });
      if (p.isCancel(toRemove)) continue;

      if (toRemove.length >= owners.length) {
        p.log.warn("Cannot remove all owners. At least one must remain.");
        continue;
      }

      for (const ownerId of toRemove) {
        const name = owners.find((o) => o.Id === ownerId)?.DisplayName ?? ownerId;
        const spin = p.spinner();
        spin.start(`Removing ${name} as owner...`);
        const { error } = await ps.runCommand(
          `Remove-MgGroupOwnerByRef -GroupId '${escapePS(selectedId)}' -DirectoryObjectId '${escapePS(ownerId)}'`,
        );
        if (error) {
          spin.stop(`Failed to remove ${name} as owner.`);
          p.log.error(error);
        } else {
          spin.stop(`Removed ${name} as owner.`);
        }
      }
    }

    // Refresh details after each action
    const refreshSpin = p.spinner();
    refreshSpin.start("Refreshing details...");
    ({ members, owners } = await fetchSecGroupDetails(ps, selectedId));
    refreshSpin.stop("Details refreshed.");
    showSecGroupDetails(currentName, members, owners);
  }
}

// ── Shared Mailbox ──────────────────────────────────────────────────────

interface MailboxPermHolder {
  user: string;
  permissions: string[];
}

async function fetchSharedMailboxDetails(
  ps: PowerShellSession,
  email: string,
): Promise<{ fullAccess: string[]; sendAs: string[]; sendOnBehalf: string[] }> {
  let fullAccess: string[] = [];
  let sendAs: string[] = [];
  let sendOnBehalf: string[] = [];

  try {
    const { output } = await ps.runCommand(
      `Get-EXOMailboxPermission -Identity '${escapePS(email)}' | Where-Object { $_.User -ne 'NT AUTHORITY\\SELF' -and $_.User -notmatch '^S-1-' -and $_.AccessRights -contains 'FullAccess' } | Select-Object -ExpandProperty User`,
    );
    if (output.trim()) {
      fullAccess = output.trim().split("\n").map((l) => l.trim()).filter(Boolean);
    }
  } catch { /* no perms */ }

  try {
    const { output } = await ps.runCommand(
      `Get-EXORecipientPermission -Identity '${escapePS(email)}' | Where-Object { $_.Trustee -ne 'NT AUTHORITY\\SELF' -and $_.Trustee -notmatch '^S-1-' } | Select-Object -ExpandProperty Trustee`,
    );
    if (output.trim()) {
      sendAs = output.trim().split("\n").map((l) => l.trim()).filter(Boolean);
    }
  } catch { /* no perms */ }

  try {
    const { output } = await ps.runCommand(
      `Get-Mailbox -Identity '${escapePS(email)}' | Select-Object -ExpandProperty GrantSendOnBehalfTo`,
    );
    if (output.trim()) {
      const dns = output.trim().split("\n").map((l) => l.trim()).filter(Boolean);
      for (const dn of dns) {
        const { output: emailOut } = await ps.runCommand(
          `Get-EXORecipient '${escapePS(dn)}' -Properties PrimarySmtpAddress -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress`,
        );
        sendOnBehalf.push(emailOut.trim() || dn);
      }
    }
  } catch { /* no perms */ }

  return { fullAccess, sendAs, sendOnBehalf };
}

function showSharedMailboxDetails(
  displayName: string,
  email: string,
  details: { fullAccess: string[]; sendAs: string[]; sendOnBehalf: string[] },
): void {
  const lines = [
    `Name:           ${displayName}`,
    `Email:          ${email}`,
    `Full Access:    ${details.fullAccess.length > 0 ? details.fullAccess.join(", ") : "(none)"}`,
    `Send As:        ${details.sendAs.length > 0 ? details.sendAs.join(", ") : "(none)"}`,
    `Send on Behalf: ${details.sendOnBehalf.length > 0 ? details.sendOnBehalf.join(", ") : "(none)"}`,
  ];
  p.note(lines.join("\n"), "Shared mailbox details");
}

function getUniquePermHolders(
  details: { fullAccess: string[]; sendAs: string[]; sendOnBehalf: string[] },
): MailboxPermHolder[] {
  const allUsers = new Set([...details.fullAccess, ...details.sendAs, ...details.sendOnBehalf]);
  const holders: MailboxPermHolder[] = [];
  for (const user of allUsers) {
    const perms: string[] = [];
    if (details.fullAccess.includes(user)) perms.push("FullAccess");
    if (details.sendAs.includes(user)) perms.push("SendAs");
    if (details.sendOnBehalf.includes(user)) perms.push("SendOnBehalf");
    holders.push({ user, permissions: perms });
  }
  return holders;
}

async function editSharedMailbox(ps: PowerShellSession): Promise<void> {
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

  const selectedEmail = await p.select({
    message: "Select shared mailbox to edit",
    options: mailboxes.map((m) => ({
      value: m.PrimarySmtpAddress,
      label: m.DisplayName,
      hint: m.PrimarySmtpAddress,
    })),
  });
  if (p.isCancel(selectedEmail)) return;

  let currentName = mailboxes.find((m) => m.PrimarySmtpAddress === selectedEmail)!.DisplayName;

  // Ensure Graph for user fetching
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph...");
  try {
    await ps.ensureGraphConnected();
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return;
  }

  // Show current details
  const detailSpin = p.spinner();
  detailSpin.start("Fetching mailbox details...");
  let details = await fetchSharedMailboxDetails(ps, selectedEmail);
  detailSpin.stop("Mailbox details loaded.");
  showSharedMailboxDetails(currentName, selectedEmail, details);

  // Edit loop
  while (true) {
    const action = await p.select({
      message: "Edit action",
      options: [
        { value: "rename", label: "Rename" },
        { value: "add-permissions", label: "Add user permissions" },
        { value: "remove-permissions", label: "Remove user permissions" },
        { value: "done", label: "Done" },
      ],
    });
    if (p.isCancel(action) || action === "done") break;

    if (action === "rename") {
      const newName = await p.text({
        message: "New display name",
        initialValue: currentName,
        validate: (v = "") => (!v.trim() ? "Display name is required" : undefined),
      });
      if (p.isCancel(newName)) continue;

      const spin = p.spinner();
      spin.start("Renaming...");
      const { error } = await ps.runCommand(
        `Set-Mailbox -Identity '${escapePS(selectedEmail)}' -DisplayName '${escapePS(newName)}'`,
      );
      if (error) {
        spin.stop("Failed to rename.");
        p.log.error(error);
      } else {
        spin.stop(`Renamed to "${newName}".`);
        currentName = newName;
      }
    }

    if (action === "add-permissions") {
      const users = await selectMultipleUsers(ps, "Select users to grant permissions");
      if (users.length > 0) {
        const permissions = await p.multiselect({
          message: "Select permissions to grant (space to select, enter to confirm)",
          options: [
            { value: "read-manage", label: "Read and Manage (Full Access)" },
            { value: "send-as", label: "Send As" },
            { value: "send-on-behalf", label: "Send on Behalf" },
          ],
          required: true,
        });
        if (!p.isCancel(permissions)) {
          const permLabels: Record<string, string> = {
            "read-manage": "Read and Manage",
            "send-as": "Send As",
            "send-on-behalf": "Send on Behalf",
          };

          for (const user of users) {
            const spin = p.spinner();
            spin.start(`Granting permissions to ${user.DisplayName}...`);
            const errors: string[] = [];

            for (const perm of permissions) {
              let result: { error: string };

              if (perm === "read-manage") {
                result = await ps.runCommand(
                  `Add-MailboxPermission -Identity '${escapePS(selectedEmail)}' -User '${escapePS(user.UserPrincipalName)}' -AccessRights FullAccess -InheritanceType All`,
                );
              } else if (perm === "send-as") {
                result = await ps.runCommand(
                  `Add-RecipientPermission -Identity '${escapePS(selectedEmail)}' -Trustee '${escapePS(user.UserPrincipalName)}' -AccessRights SendAs -Confirm:$false`,
                );
              } else {
                result = await ps.runCommand(
                  `Set-Mailbox -Identity '${escapePS(selectedEmail)}' -GrantSendOnBehalfTo @{Add='${escapePS(user.UserPrincipalName)}'}`,
                );
              }

              if (result.error) {
                errors.push(`${permLabels[perm]}: ${result.error}`);
              }
            }

            if (errors.length === 0) {
              spin.stop(`Granted permissions to ${user.DisplayName}.`);
            } else if (errors.length < permissions.length) {
              spin.stop(`Some permissions failed for ${user.DisplayName}.`);
              for (const err of errors) p.log.error(err);
            } else {
              spin.stop(`Failed to grant permissions to ${user.DisplayName}.`);
              for (const err of errors) p.log.error(err);
            }
          }
        }
      }
    }

    if (action === "remove-permissions") {
      const holders = getUniquePermHolders(details);
      if (holders.length === 0) {
        p.log.warn("No permission holders to remove.");
        continue;
      }

      const toRemove = await p.multiselect({
        message: "Select users to remove permissions from (space to select, enter to confirm)",
        options: holders.map((h) => ({
          value: h.user,
          label: h.user,
          hint: h.permissions.join(", "),
        })),
        required: true,
      });
      if (p.isCancel(toRemove)) continue;

      for (const user of toRemove) {
        const holder = holders.find((h) => h.user === user)!;
        const spin = p.spinner();
        spin.start(`Removing permissions from ${user}...`);
        const errors: string[] = [];

        if (holder.permissions.includes("FullAccess")) {
          const { error } = await ps.runCommand(
            `Remove-MailboxPermission -Identity '${escapePS(selectedEmail)}' -User '${escapePS(user)}' -AccessRights FullAccess -Confirm:$false`,
          );
          if (error) errors.push(`FullAccess: ${error}`);
        }

        if (holder.permissions.includes("SendAs")) {
          const { error } = await ps.runCommand(
            `Remove-RecipientPermission -Identity '${escapePS(selectedEmail)}' -Trustee '${escapePS(user)}' -AccessRights SendAs -Confirm:$false`,
          );
          if (error) errors.push(`SendAs: ${error}`);
        }

        if (holder.permissions.includes("SendOnBehalf")) {
          const { error } = await ps.runCommand(
            `Set-Mailbox -Identity '${escapePS(selectedEmail)}' -GrantSendOnBehalfTo @{Remove='${escapePS(user)}'}`,
          );
          if (error) errors.push(`SendOnBehalf: ${error}`);
        }

        if (errors.length === 0) {
          spin.stop(`Removed all permissions from ${user}.`);
        } else {
          spin.stop(`Some removals failed for ${user}.`);
          for (const err of errors) p.log.error(err);
        }
      }
    }

    // Refresh details after each action
    const refreshSpin = p.spinner();
    refreshSpin.start("Refreshing details...");
    details = await fetchSharedMailboxDetails(ps, selectedEmail);
    refreshSpin.stop("Details refreshed.");
    showSharedMailboxDetails(currentName, selectedEmail, details);
  }
}

export async function run(ps: PowerShellSession): Promise<void> {
  const type = await p.select({
    message: "What would you like to edit?",
    options: [
      { value: "distribution", label: "Distribution group" },
      { value: "security", label: "Security group" },
      { value: "shared-mailbox", label: "Shared mailbox" },
    ],
  });
  if (p.isCancel(type)) return;

  switch (type) {
    case "distribution":
      await editDistributionGroup(ps);
      break;
    case "security":
      await editSecurityGroup(ps);
      break;
    case "shared-mailbox":
      await editSharedMailbox(ps);
      break;
  }
}
