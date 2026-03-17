import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { escapePS } from "../utils.ts";

interface AcceptedDomain {
  DomainName: string;
  Default: boolean;
}

interface MgUser {
  DisplayName: string;
  UserPrincipalName: string;
  Id: string;
}

interface CreatedGroup {
  Id: string;
  DisplayName: string;
}

async function copyToClipboard(ps: PowerShellSession, text: string): Promise<boolean> {
  try {
    if (process.platform === "win32") {
      await ps.runCommand(`Set-Clipboard -Value '${escapePS(text)}'`);
    } else {
      const proc = Bun.spawn(["pbcopy"], { stdin: new Blob([text]) });
      await proc.exited;
    }
    return true;
  } catch {
    return false;
  }
}

async function fetchAcceptedDomains(ps: PowerShellSession): Promise<AcceptedDomain[]> {
  const raw = await ps.runCommandJson<AcceptedDomain | AcceptedDomain[]>(
    "Get-AcceptedDomain | Select-Object DomainName, Default",
  );
  return raw ? (Array.isArray(raw) ? raw : [raw]) : [];
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

async function promptEmail(
  ps: PowerShellSession,
): Promise<{ alias: string; domain: string; email: string } | null> {
  const domainSpin = p.spinner();
  domainSpin.start("Fetching accepted domains...");
  let domains: AcceptedDomain[];
  try {
    domains = await fetchAcceptedDomains(ps);
    domainSpin.stop(`Found ${domains.length} domain(s).`);
  } catch (e) {
    domainSpin.stop("Failed to fetch domains.");
    p.log.error(`${e}`);
    return null;
  }

  if (domains.length === 0) {
    p.log.error("No accepted domains found.");
    return null;
  }

  const alias = await p.text({
    message: "Email alias (before @)",
    validate: (v = "") => {
      if (!v.trim()) return "Alias is required";
      if (/[^a-zA-Z0-9._-]/.test(v)) return "Invalid characters in alias";
    },
  });
  if (p.isCancel(alias)) return null;

  const defaultDomain = domains.find((d) => d.Default)?.DomainName ?? domains[0]?.DomainName;

  const domain = await p.select({
    message: "Domain",
    options: domains.map((d) => ({
      value: d.DomainName,
      label: d.DomainName,
      hint: d.Default ? "default" : undefined,
    })),
    initialValue: defaultDomain,
  });
  if (p.isCancel(domain)) return null;

  return { alias, domain, email: `${alias}@${domain}` };
}

async function checkRecipientExists(ps: PowerShellSession, email: string): Promise<boolean> {
  const checkSpin = p.spinner();
  checkSpin.start(`Checking if ${email} is available...`);
  const { output } = await ps.runCommand(
    `Get-EXORecipient -Identity '${escapePS(email)}' -Properties PrimarySmtpAddress -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress`,
  );
  if (output.trim()) {
    checkSpin.stop(`${email} is already taken.`);
    return true;
  }
  checkSpin.stop(`${email} is available.`);
  return false;
}

async function createDistributionGroup(ps: PowerShellSession): Promise<void> {
  const displayName = await p.text({
    message: "Display name",
    validate: (v = "") => (!v.trim() ? "Display name is required" : undefined),
  });
  if (p.isCancel(displayName)) return;

  const emailResult = await promptEmail(ps);
  if (!emailResult) return;

  if (await checkRecipientExists(ps, emailResult.email)) {
    p.log.error("Please choose a different email address.");
    return;
  }

  const createSpin = p.spinner();
  createSpin.start("Creating distribution group...");
  const { error } = await ps.runCommand(
    `New-DistributionGroup -Name '${escapePS(displayName)}' -PrimarySmtpAddress '${escapePS(emailResult.email)}'`,
  );
  if (error) {
    createSpin.stop("Failed to create distribution group.");
    p.log.error(error);
    return;
  }
  createSpin.stop("Distribution group created.");

  // Ensure Graph for user fetching
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph...");
  try {
    await ps.ensureGraphConnected(["User.Read.All"]);
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    p.note(
      [`Name:  ${displayName}`, `Email: ${emailResult.email}`].join("\n"),
      "Distribution group created",
    );
    return;
  }

  // Add members
  const addedMembers: { name: string; upn: string }[] = [];
  const addMembers = await p.confirm({ message: "Add members?" });
  if (!p.isCancel(addMembers) && addMembers) {
    const members = await selectMultipleUsers(ps, "Select members");
    for (const member of members) {
      const spin = p.spinner();
      spin.start(`Adding ${member.DisplayName}...`);
      const { error } = await ps.runCommand(
        `Add-DistributionGroupMember -Identity '${escapePS(emailResult.email)}' -Member '${escapePS(member.UserPrincipalName)}'`,
      );
      if (error) {
        spin.stop(`Failed to add ${member.DisplayName}.`);
        p.log.error(error);
      } else {
        spin.stop(`Added ${member.DisplayName}.`);
        addedMembers.push({ name: member.DisplayName, upn: member.UserPrincipalName });
      }
    }
  }

  // Add owners
  const addedOwners: { name: string; upn: string }[] = [];
  const addOwners = await p.confirm({ message: "Add owners?" });
  if (!p.isCancel(addOwners) && addOwners) {
    const owners = await selectMultipleUsers(ps, "Select owners");
    for (const owner of owners) {
      const spin = p.spinner();
      spin.start(`Adding ${owner.DisplayName} as owner...`);
      const { error } = await ps.runCommand(
        `Set-DistributionGroup -Identity '${escapePS(emailResult.email)}' -ManagedBy @{Add='${escapePS(owner.UserPrincipalName)}'}`,
      );
      if (error) {
        spin.stop(`Failed to add ${owner.DisplayName} as owner.`);
        p.log.error(error);
      } else {
        spin.stop(`Added ${owner.DisplayName} as owner.`);
        addedOwners.push({ name: owner.DisplayName, upn: owner.UserPrincipalName });
      }
    }
  }

  p.note(
    [`Name:  ${displayName}`, `Email: ${emailResult.email}`].join("\n"),
    "Distribution group created",
  );

  // Ticket note clipboard option
  let ticketNote = `Created distribution group ${displayName} (${emailResult.email}).`;
  if (addedMembers.length > 0) {
    ticketNote += `\nAdded members:\n${addedMembers.map((m) => `  - ${m.name} (${m.upn})`).join("\n")}`;
  }
  if (addedOwners.length > 0) {
    ticketNote += `\nAdded owners:\n${addedOwners.map((o) => `  - ${o.name} (${o.upn})`).join("\n")}`;
  }

  const distAction = await p.select({
    message: "Next",
    options: [
      { value: "copy-ticket", label: "Copy ticket update note to clipboard" },
      { value: "done", label: "Done" },
    ],
  });
  if (!p.isCancel(distAction) && distAction === "copy-ticket") {
    if (await copyToClipboard(ps, ticketNote)) {
      p.log.success("Ticket update note copied to clipboard.");
    } else {
      p.log.info(ticketNote);
    }
  }
}

async function createSecurityGroup(ps: PowerShellSession): Promise<void> {
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph (check your browser)...");
  try {
    await ps.ensureGraphConnected(["User.Read.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All"]);
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return;
  }

  const displayName = await p.text({
    message: "Display name",
    validate: (v = "") => (!v.trim() ? "Display name is required" : undefined),
  });
  if (p.isCancel(displayName)) return;

  const suggestedNickname = displayName.trim().toLowerCase().replace(/\s+/g, "-").replace(/[^a-zA-Z0-9._-]/g, "");

  const mailNickname = await p.text({
    message: "Mail nickname",
    initialValue: suggestedNickname,
    validate: (v = "") => {
      if (!v.trim()) return "Mail nickname is required";
      if (/[^a-zA-Z0-9._-]/.test(v)) return "Invalid characters (use letters, numbers, . _ -)";
    },
  });
  if (p.isCancel(mailNickname)) return;

  const createSpin = p.spinner();
  createSpin.start("Creating security group...");
  let group: CreatedGroup | null;
  try {
    group = await ps.runCommandJson<CreatedGroup>(
      `New-MgGroup -DisplayName '${escapePS(displayName)}' -MailNickname '${escapePS(mailNickname)}' -MailEnabled:$false -SecurityEnabled:$true -GroupTypes @() | Select-Object Id, DisplayName`,
    );
    createSpin.stop("Security group created.");
  } catch (e) {
    createSpin.stop("Failed to create security group.");
    p.log.error(`${e}`);
    return;
  }

  if (!group?.Id) {
    p.log.error("Failed to retrieve group ID.");
    return;
  }

  // Add members
  const addedMembers: { name: string; upn: string }[] = [];
  const addMembers = await p.confirm({ message: "Add members?" });
  if (!p.isCancel(addMembers) && addMembers) {
    const members = await selectMultipleUsers(ps, "Select members");
    for (const member of members) {
      const spin = p.spinner();
      spin.start(`Adding ${member.DisplayName}...`);
      const { error } = await ps.runCommand(
        `New-MgGroupMemberByRef -GroupId '${escapePS(group.Id)}' -OdataId 'https://graph.microsoft.com/v1.0/directoryObjects/${escapePS(member.Id)}'`,
      );
      if (error) {
        spin.stop(`Failed to add ${member.DisplayName}.`);
        p.log.error(error);
      } else {
        spin.stop(`Added ${member.DisplayName}.`);
        addedMembers.push({ name: member.DisplayName, upn: member.UserPrincipalName });
      }
    }
  }

  // Add owners
  const addedOwners: { name: string; upn: string }[] = [];
  const addOwners = await p.confirm({ message: "Add owners?" });
  if (!p.isCancel(addOwners) && addOwners) {
    const owners = await selectMultipleUsers(ps, "Select owners");
    for (const owner of owners) {
      const spin = p.spinner();
      spin.start(`Adding ${owner.DisplayName} as owner...`);
      const { error } = await ps.runCommand(
        `New-MgGroupOwnerByRef -GroupId '${escapePS(group.Id)}' -OdataId 'https://graph.microsoft.com/v1.0/directoryObjects/${escapePS(owner.Id)}'`,
      );
      if (error) {
        spin.stop(`Failed to add ${owner.DisplayName} as owner.`);
        p.log.error(error);
      } else {
        spin.stop(`Added ${owner.DisplayName} as owner.`);
        addedOwners.push({ name: owner.DisplayName, upn: owner.UserPrincipalName });
      }
    }
  }

  p.note(
    [`Name:     ${displayName}`, `Nickname: ${mailNickname}`].join("\n"),
    "Security group created",
  );

  // Ticket note clipboard option
  let ticketNote = `Created security group ${displayName} (${mailNickname}).`;
  if (addedMembers.length > 0) {
    ticketNote += `\nAdded members:\n${addedMembers.map((m) => `  - ${m.name} (${m.upn})`).join("\n")}`;
  }
  if (addedOwners.length > 0) {
    ticketNote += `\nAdded owners:\n${addedOwners.map((o) => `  - ${o.name} (${o.upn})`).join("\n")}`;
  }

  const secAction = await p.select({
    message: "Next",
    options: [
      { value: "copy-ticket", label: "Copy ticket update note to clipboard" },
      { value: "done", label: "Done" },
    ],
  });
  if (!p.isCancel(secAction) && secAction === "copy-ticket") {
    if (await copyToClipboard(ps, ticketNote)) {
      p.log.success("Ticket update note copied to clipboard.");
    } else {
      p.log.info(ticketNote);
    }
  }
}

async function createSharedMailbox(ps: PowerShellSession): Promise<void> {
  const displayName = await p.text({
    message: "Display name",
    validate: (v = "") => (!v.trim() ? "Display name is required" : undefined),
  });
  if (p.isCancel(displayName)) return;

  const emailResult = await promptEmail(ps);
  if (!emailResult) return;

  if (await checkRecipientExists(ps, emailResult.email)) {
    p.log.error("Please choose a different email address.");
    return;
  }

  const createSpin = p.spinner();
  createSpin.start("Creating shared mailbox...");
  const { error } = await ps.runCommand(
    `New-Mailbox -Name '${escapePS(displayName)}' -Shared -PrimarySmtpAddress '${escapePS(emailResult.email)}'`,
  );
  if (error) {
    createSpin.stop("Failed to create shared mailbox.");
    p.log.error(error);
    return;
  }
  createSpin.stop("Shared mailbox created.");

  // Add user permissions
  const addedUsers: { name: string; upn: string; perms: string[] }[] = [];
  const addPerms = await p.confirm({ message: "Add user permissions?" });
  if (!p.isCancel(addPerms) && addPerms) {
    const graphSpin = p.spinner();
    graphSpin.start("Connecting to Microsoft Graph...");
    try {
      await ps.ensureGraphConnected(["User.Read.All"]);
      graphSpin.stop("Connected to Microsoft Graph.");
    } catch (e) {
      graphSpin.stop("Failed to connect to Microsoft Graph.");
      p.log.error(`${e}`);
      p.note(
        [`Name:  ${displayName}`, `Email: ${emailResult.email}`].join("\n"),
        "Shared mailbox created",
      );
      return;
    }

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
          const grantedPerms: string[] = [];

          for (const perm of permissions) {
            let result: { error: string };

            if (perm === "read-manage") {
              result = await ps.runCommand(
                `Add-MailboxPermission -Identity '${escapePS(emailResult.email)}' -User '${escapePS(user.UserPrincipalName)}' -AccessRights FullAccess -InheritanceType All`,
              );
            } else if (perm === "send-as") {
              result = await ps.runCommand(
                `Add-RecipientPermission -Identity '${escapePS(emailResult.email)}' -Trustee '${escapePS(user.UserPrincipalName)}' -AccessRights SendAs -Confirm:$false`,
              );
            } else {
              result = await ps.runCommand(
                `Set-Mailbox -Identity '${escapePS(emailResult.email)}' -GrantSendOnBehalfTo @{Add='${escapePS(user.UserPrincipalName)}'}`,
              );
            }

            if (result.error) {
              errors.push(`${permLabels[perm]}: ${result.error}`);
            } else {
              grantedPerms.push(permLabels[perm] ?? perm);
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

          if (grantedPerms.length > 0) {
            addedUsers.push({ name: user.DisplayName, upn: user.UserPrincipalName, perms: grantedPerms });
          }
        }
      }
    }
  }

  p.note(
    [`Name:  ${displayName}`, `Email: ${emailResult.email}`].join("\n"),
    "Shared mailbox created",
  );

  // Ticket note clipboard option
  let ticketNote = `Created shared mailbox ${displayName} (${emailResult.email}).`;
  if (addedUsers.length > 0) {
    ticketNote += `\nAdded users with permissions:`;
    for (const u of addedUsers) {
      ticketNote += `\n  - ${u.name} (${u.upn}) — ${u.perms.join(", ")}`;
    }
  }

  const mbAction = await p.select({
    message: "Next",
    options: [
      { value: "copy-ticket", label: "Copy ticket update note to clipboard" },
      { value: "done", label: "Done" },
    ],
  });
  if (!p.isCancel(mbAction) && mbAction === "copy-ticket") {
    if (await copyToClipboard(ps, ticketNote)) {
      p.log.success("Ticket update note copied to clipboard.");
    } else {
      p.log.info(ticketNote);
    }
  }
}

export async function run(ps: PowerShellSession): Promise<void> {
  const type = await p.select({
    message: "What would you like to create?",
    options: [
      { value: "distribution", label: "Distribution group" },
      { value: "security", label: "Security group" },
      { value: "shared-mailbox", label: "Shared mailbox" },
    ],
  });
  if (p.isCancel(type)) return;

  switch (type) {
    case "distribution":
      await ps.ensureExchangeConnected();
      await createDistributionGroup(ps);
      break;
    case "security":
      await createSecurityGroup(ps);
      break;
    case "shared-mailbox":
      await ps.ensureExchangeConnected();
      await createSharedMailbox(ps);
      break;
  }
}
