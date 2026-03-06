import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";
import { escapePS } from "../utils.ts";

interface MgUser {
  DisplayName: string;
  UserPrincipalName: string;
  Id: string;
}

interface LicenseDetail {
  SkuPartNumber: string;
  SkuId: string;
}

interface SharedMailboxPerm {
  DisplayName: string;
  PrimarySmtpAddress: string;
  HasFullAccess: boolean;
  HasSendAs: boolean;
  HasSendOnBehalf: boolean;
  OtherFullAccessCount: number;
}

interface DistGroupMembership {
  DisplayName: string;
  PrimarySmtpAddress: string;
  MemberCount: number;
}

interface SecGroupMembership {
  DisplayName: string;
  Id: string;
  MemberCount: number;
}

function elapsedTimer(spin: { message(msg?: string): void }, baseMsg: string): () => void {
  const start = Date.now();
  const interval = setInterval(() => {
    const secs = Math.floor((Date.now() - start) / 1000);
    const mins = Math.floor(secs / 60);
    const elapsed = mins > 0 ? `${mins}m ${secs % 60}s` : `${secs}s`;
    spin.message(`${baseMsg} (${elapsed})`);
  }, 1000);
  return () => clearInterval(interval);
}

async function fetchUsers(
  ps: PowerShellSession,
  excludeUpn?: string,
): Promise<MgUser[]> {
  // Get user count
  const { output: countOutput } = await ps.runCommand(
    "Get-MgUser -Top 1 -CountVariable ct -ConsistencyLevel eventual | Out-Null; $ct",
  );
  const count = parseInt(countOutput.trim(), 10);

  let users: MgUser[];

  if (count <= 50) {
    const raw = await ps.runCommandJson<MgUser | MgUser[]>(
      "Get-MgUser -All -Property DisplayName,UserPrincipalName,Id | Select-Object DisplayName,UserPrincipalName,Id",
    );
    users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
  } else {
    // Search mode for large tenants
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
          `Get-MgUser -Search '"displayName:${escapePS(query)}"' -ConsistencyLevel eventual -Property DisplayName,UserPrincipalName,Id | Select-Object DisplayName,UserPrincipalName,Id`,
        );
        users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        searchSpin.stop(`Found ${users.length} user(s).`);
      } catch {
        searchSpin.stop("Search returned no results.");
        users = [];
      }

      // Filter before checking if empty, so excluded user doesn't count
      if (excludeUpn) {
        users = users.filter(
          (u) => u.UserPrincipalName.toLowerCase() !== excludeUpn.toLowerCase(),
        );
      }

      if (users.length === 0) {
        p.log.warn("No users found. Try a different search term.");
        continue;
      }
      break;
    }

    return users.sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));
  }

  // Filter for the non-search path
  if (excludeUpn) {
    users = users.filter(
      (u) => u.UserPrincipalName.toLowerCase() !== excludeUpn.toLowerCase(),
    );
  }

  return users.sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));
}

async function selectUser(
  ps: PowerShellSession,
  message: string,
  excludeUpn?: string,
): Promise<MgUser | null> {
  const users = await fetchUsers(ps, excludeUpn);
  if (users.length === 0) return null;

  const userId = await p.select({
    message,
    options: users.map((u) => ({
      value: u.Id,
      label: u.DisplayName,
      hint: u.UserPrincipalName,
    })),
  });
  if (p.isCancel(userId)) return null;

  return users.find((u) => u.Id === userId) ?? null;
}

async function selectMultipleUsers(
  ps: PowerShellSession,
  message: string,
  excludeUpn?: string,
): Promise<MgUser[]> {
  while (true) {
    const users = await fetchUsers(ps, excludeUpn);

    if (users.length === 0) {
      const retry = await p.confirm({ message: "No other users found. Search again?" });
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

export async function run(ps: PowerShellSession): Promise<void> {
  // Ensure Graph connected
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

  // ── GATHER PHASE ──────────────────────────────────────────────────────

  // 1. Select user
  const user = await selectUser(ps, "Select user to delete");
  if (!user) return;

  const upn = user.UserPrincipalName;
  const userId = user.Id;

  // Fetch licenses
  const licSpin = p.spinner();
  licSpin.start("Fetching user details...");
  let licenses: LicenseDetail[] = [];
  try {
    const raw = await ps.runCommandJson<LicenseDetail | LicenseDetail[]>(
      `Get-MgUserLicenseDetail -UserId '${escapePS(userId)}' | Select-Object SkuPartNumber, SkuId`,
    );
    licenses = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
    licSpin.stop("User details loaded.");
  } catch {
    licSpin.stop("User details loaded.");
  }

  p.note(
    [
      `Name:     ${user.DisplayName}`,
      `UPN:      ${upn}`,
      `Licenses: ${licenses.length > 0 ? licenses.map((l) => friendlySkuName(l.SkuPartNumber)).join(", ") : "(none)"}`,
    ].join("\n"),
    "User to delete",
  );

  // 2. Check mailbox & ask about conversion
  const mbSpin = p.spinner();
  mbSpin.start("Checking mailbox...");
  const { output: mbOutput } = await ps.runCommand(
    `Get-Mailbox -Identity '${escapePS(upn)}' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty RecipientTypeDetails`,
  );
  const mbType = mbOutput.trim();
  mbSpin.stop(mbType ? `Mailbox type: ${mbType}` : "No Exchange mailbox found.");

  let willConvertToShared = false;
  let mailboxDelegates: MgUser[] = [];

  if (mbType === "UserMailbox") {
    const convert = await p.confirm({
      message: "Convert mailbox to shared before deletion?",
    });
    if (p.isCancel(convert)) return;

    if (convert) {
      willConvertToShared = true;

      const grantAccess = await p.confirm({
        message: "Grant other users access to this shared mailbox?",
      });
      if (p.isCancel(grantAccess)) return;

      if (grantAccess) {
        mailboxDelegates = await selectMultipleUsers(
          ps,
          "Select delegate(s) for mailbox access",
          upn,
        );
      }
    }
  } else if (mbType === "SharedMailbox") {
    p.log.info("Mailbox is already shared — skipping conversion.");
  }

  // 3. Ask about OneDrive
  const shareDrive = await p.confirm({
    message: "Grant another user access to this user's OneDrive?",
  });
  if (p.isCancel(shareDrive)) return;

  if (shareDrive) {
    p.log.warn("Sharing OneDrive from terminal is not currently supported.");
    p.log.info("You can grant OneDrive access from the Microsoft 365 admin center instead.");
    const cont = await p.confirm({ message: "Continue without OneDrive sharing?" });
    if (p.isCancel(cont) || !cont) {
      p.log.info("Cancelled.");
      return;
    }
  }

  // ── CONFIRMATION ──────────────────────────────────────────────────────

  const planLines: string[] = [
    willConvertToShared
      ? `Convert to shared mailbox: ${user.DisplayName} (${upn})`
      : `Delete user: ${user.DisplayName} (${upn})`,
  ];

  if (licenses.length > 0) {
    planLines.push(
      `\nLicenses to ${willConvertToShared ? "remove" : "release"}:\n${licenses.map((l) => `  - ${friendlySkuName(l.SkuPartNumber)}`).join("\n")}`,
    );
  }

  if (willConvertToShared) {
    planLines.push("\nMailbox: Convert to shared");
    planLines.push("Account: Convert to shared mailbox (licenses removed)");
    if (mailboxDelegates.length > 0) {
      planLines.push(
        `Mailbox delegates: ${mailboxDelegates.map((d) => d.DisplayName).join(", ")}`,
      );
    }
  }

  if (willConvertToShared) {
    planLines.push("\nWill scan & clean up memberships (shared mailboxes, distribution groups, security groups)");
  }

  p.note(planLines.join("\n"), "Planned actions");

  const confirm = await p.confirm({
    message: willConvertToShared ? "Proceed?" : "Proceed with deletion?",
  });
  if (p.isCancel(confirm) || !confirm) {
    p.log.info("Cancelled.");
    return;
  }

  // ── EXECUTE PHASE ─────────────────────────────────────────────────────

  const releasedLicenses = licenses.map((l) => friendlySkuName(l.SkuPartNumber));
  let convertedToShared = false;
  const mailboxDelegateNames: string[] = [];
  const removedMailboxes: string[] = [];
  const removedDistGroups: string[] = [];
  const removedSecGroups: string[] = [];
  const orphans: { type: string; name: string }[] = [];

  // Execute: Convert mailbox
  if (willConvertToShared) {
    const convertSpin = p.spinner();
    convertSpin.start("Converting to shared mailbox...");
    const { error } = await ps.runCommand(
      `Set-Mailbox -Identity '${escapePS(upn)}' -Type Shared`,
    );
    if (error) {
      convertSpin.stop("Failed to convert mailbox.");
      p.log.error(error);
      p.log.error("Cannot proceed — mailbox must be converted before deletion.");
      return;
    }
    convertSpin.stop("Mailbox converted to shared.");
    convertedToShared = true;
  }

  // Grant mailbox delegates (only if conversion succeeded)
  if (convertedToShared) {
    for (const delegate of mailboxDelegates) {
      const dSpin = p.spinner();
      dSpin.start(`Granting mailbox access to ${delegate.DisplayName}...`);
      const errors: string[] = [];

      const { error: faErr } = await ps.runCommand(
        `Add-MailboxPermission -Identity '${escapePS(upn)}' -User '${escapePS(delegate.UserPrincipalName)}' -AccessRights FullAccess -InheritanceType All -AutoMapping $true`,
      );
      if (faErr) errors.push(`FullAccess: ${faErr}`);

      const { error: saErr } = await ps.runCommand(
        `Add-RecipientPermission -Identity '${escapePS(upn)}' -Trustee '${escapePS(delegate.UserPrincipalName)}' -AccessRights SendAs -Confirm:$false`,
      );
      if (saErr) errors.push(`SendAs: ${saErr}`);

      if (errors.length === 0) {
        dSpin.stop(`Granted access to ${delegate.DisplayName}.`);
        mailboxDelegateNames.push(delegate.DisplayName);
      } else if (errors.length < 2) {
        dSpin.stop(`Partially granted access to ${delegate.DisplayName}.`);
        mailboxDelegateNames.push(delegate.DisplayName);
        for (const err of errors) p.log.error(err);
      } else {
        dSpin.stop(`Failed to grant access to ${delegate.DisplayName}.`);
        for (const err of errors) p.log.error(err);
      }
    }
  }

  // Execute: Scan & clean up memberships (only when mailbox is preserved as shared)
  if (convertedToShared) {
    p.log.info("Scanning memberships to clean up — this can take a few minutes.");

    let sharedMailboxPerms: SharedMailboxPerm[] = [];
    let distGroupMemberships: DistGroupMembership[] = [];
    let secGroupMemberships: SecGroupMembership[] = [];

    // Scan shared mailbox permissions
    {
      const scanMbSpin = p.spinner();
      scanMbSpin.start("Scanning shared mailbox permissions...");
      const stopTimer = elapsedTimer(scanMbSpin, "Scanning shared mailbox permissions");
      try {
        const raw = await ps.runCommandJson<SharedMailboxPerm | SharedMailboxPerm[]>(
          `$targetUpn = '${escapePS(upn)}'
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | ForEach-Object {
  $mb = $_
  $fa = $null; $sa = $null; $sob = $false
  try { $fa = Get-MailboxPermission -Identity $mb.Identity -User $targetUpn -ErrorAction Stop } catch {}
  try { $sa = Get-RecipientPermission -Identity $mb.Identity -Trustee $targetUpn -ErrorAction Stop } catch {}
  if ($mb.GrantSendOnBehalfTo -contains $targetUpn) { $sob = $true }
  if ($fa -or $sa -or $sob) {
    $otherFa = @(Get-MailboxPermission -Identity $mb.Identity | Where-Object { $_.User -ne $targetUpn -and $_.User -ne 'NT AUTHORITY\\SELF' -and $_.AccessRights -contains 'FullAccess' }).Count
    [PSCustomObject]@{
      DisplayName = $mb.DisplayName
      PrimarySmtpAddress = [string]$mb.PrimarySmtpAddress
      HasFullAccess = ($null -ne $fa)
      HasSendAs = ($null -ne $sa)
      HasSendOnBehalf = $sob
      OtherFullAccessCount = $otherFa
    }
  }
}`,
        );
        stopTimer();
        sharedMailboxPerms = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        scanMbSpin.stop(
          `Found permissions on ${sharedMailboxPerms.length} shared mailbox(es).`,
        );
      } catch {
        stopTimer();
        scanMbSpin.stop("No shared mailbox permissions found.");
      }
    }

    // Scan distribution group memberships
    {
      const scanDgSpin = p.spinner();
      scanDgSpin.start("Scanning distribution group memberships...");
      const stopTimer = elapsedTimer(scanDgSpin, "Scanning distribution group memberships");
      try {
        const raw = await ps.runCommandJson<DistGroupMembership | DistGroupMembership[]>(
          `Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
  $members = Get-DistributionGroupMember -Identity $_.PrimarySmtpAddress -ResultSize Unlimited
  if ($members.PrimarySmtpAddress -contains '${escapePS(upn)}') {
    [PSCustomObject]@{
      DisplayName = $_.DisplayName
      PrimarySmtpAddress = [string]$_.PrimarySmtpAddress
      MemberCount = $members.Count
    }
  }
}`,
        );
        stopTimer();
        distGroupMemberships = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        scanDgSpin.stop(
          `Found ${distGroupMemberships.length} distribution group membership(s).`,
        );
      } catch {
        stopTimer();
        scanDgSpin.stop("No distribution group memberships found.");
      }
    }

    // Scan security group memberships
    {
      const scanSgSpin = p.spinner();
      scanSgSpin.start("Scanning security group memberships...");
      const stopTimer = elapsedTimer(scanSgSpin, "Scanning security group memberships");
      try {
        const raw = await ps.runCommandJson<SecGroupMembership | SecGroupMembership[]>(
          `Get-MgUserMemberOf -UserId '${escapePS(userId)}' -All | ForEach-Object {
  $grp = Get-MgGroup -GroupId $_.Id -ErrorAction SilentlyContinue
  if ($grp -and $grp.SecurityEnabled) {
    $count = (Get-MgGroupMember -GroupId $grp.Id -All).Count
    [PSCustomObject]@{ DisplayName = $grp.DisplayName; Id = $grp.Id; MemberCount = $count }
  }
}`,
        );
        stopTimer();
        secGroupMemberships = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        scanSgSpin.stop(
          `Found ${secGroupMemberships.length} security group membership(s).`,
        );
      } catch {
        stopTimer();
        scanSgSpin.stop("No security group memberships found.");
      }
    }

    // Detect orphans
    for (const mb of sharedMailboxPerms) {
      if (mb.HasFullAccess && mb.OtherFullAccessCount === 0) {
        orphans.push({ type: "Shared mailbox", name: mb.DisplayName });
      }
    }
    for (const dg of distGroupMemberships) {
      if (dg.MemberCount === 1) {
        orphans.push({ type: "Distribution group", name: dg.DisplayName });
      }
    }
    for (const sg of secGroupMemberships) {
      if (sg.MemberCount === 1) {
        orphans.push({ type: "Security group", name: sg.DisplayName });
      }
    }

    if (orphans.length > 0) {
      p.log.warn(
        `Orphaned resources (will have no remaining members):\n${orphans.map((o) => `  - [${o.type}] ${o.name}`).join("\n")}`,
      );
    }

    // Remove shared mailbox permissions
    if (sharedMailboxPerms.length > 0) {
      const rmMbSpin = p.spinner();
      rmMbSpin.start("Removing shared mailbox permissions...");
      for (const mb of sharedMailboxPerms) {
        const errors: string[] = [];

        if (mb.HasFullAccess) {
          const { error } = await ps.runCommand(
            `Remove-MailboxPermission -Identity '${escapePS(mb.PrimarySmtpAddress)}' -User '${escapePS(upn)}' -AccessRights FullAccess -Confirm:$false`,
          );
          if (error) errors.push(`FullAccess: ${error}`);
        }
        if (mb.HasSendAs) {
          const { error } = await ps.runCommand(
            `Remove-RecipientPermission -Identity '${escapePS(mb.PrimarySmtpAddress)}' -Trustee '${escapePS(upn)}' -AccessRights SendAs -Confirm:$false`,
          );
          if (error) errors.push(`SendAs: ${error}`);
        }
        if (mb.HasSendOnBehalf) {
          const { error } = await ps.runCommand(
            `Set-Mailbox -Identity '${escapePS(mb.PrimarySmtpAddress)}' -GrantSendOnBehalfTo @{Remove='${escapePS(upn)}'}`,
          );
          if (error) errors.push(`SendOnBehalf: ${error}`);
        }

        if (errors.length > 0) {
          for (const err of errors) p.log.error(`${mb.DisplayName}: ${err}`);
        }
        removedMailboxes.push(mb.DisplayName);
      }
      rmMbSpin.stop(`Removed permissions from ${removedMailboxes.length} shared mailbox(es).`);
    }

    // Remove distribution group memberships
    if (distGroupMemberships.length > 0) {
      const rmDgSpin = p.spinner();
      rmDgSpin.start("Removing distribution group memberships...");
      for (const dg of distGroupMemberships) {
        const { error } = await ps.runCommand(
          `Remove-DistributionGroupMember -Identity '${escapePS(dg.PrimarySmtpAddress)}' -Member '${escapePS(upn)}' -Confirm:$false`,
        );
        if (error) {
          p.log.error(`${dg.DisplayName}: ${error}`);
        } else {
          removedDistGroups.push(dg.DisplayName);
        }
      }
      rmDgSpin.stop(`Removed from ${removedDistGroups.length} distribution group(s).`);
    }

    // Remove security group memberships
    if (secGroupMemberships.length > 0) {
      const rmSgSpin = p.spinner();
      rmSgSpin.start("Removing security group memberships...");
      for (const sg of secGroupMemberships) {
        const { error } = await ps.runCommand(
          `Remove-MgGroupMemberByRef -GroupId '${escapePS(sg.Id)}' -DirectoryObjectId '${escapePS(userId)}'`,
        );
        if (error) {
          p.log.error(`${sg.DisplayName}: ${error}`);
        } else {
          removedSecGroups.push(sg.DisplayName);
        }
      }
      rmSgSpin.stop(`Removed from ${removedSecGroups.length} security group(s).`);
    }
  }

  // Execute: Delete or disable user
  let userDeleted = false;
  let userDisabled = false;

  if (convertedToShared) {
    // Remove licenses to free them up
    if (licenses.length > 0) {
      const licRemSpin = p.spinner();
      licRemSpin.start("Removing licenses...");
      const skuIds = licenses.map((l) => l.SkuId);
      const { error } = await ps.runCommand(
        `Set-MgUserLicense -UserId '${escapePS(userId)}' -RemoveLicenses @(${skuIds.map((id) => `'${escapePS(id)}'`).join(",")}) -AddLicenses @{}`,
      );
      if (error) {
        licRemSpin.stop("Failed to remove licenses.");
        p.log.error(error);
      } else {
        licRemSpin.stop("Licenses removed.");
      }
    }

    // Block sign-in
    const blockSpin = p.spinner();
    blockSpin.start("Blocking sign-in...");
    const { error: blockErr } = await ps.runCommand(
      `Update-MgUser -UserId '${escapePS(userId)}' -AccountEnabled:$false`,
    );
    if (blockErr) {
      blockSpin.stop("Failed to block sign-in.");
      p.log.error(blockErr);
    } else {
      blockSpin.stop("Sign-in blocked.");
      userDisabled = true;
    }
  } else {
    // No shared mailbox — fully delete the user
    const delSpin = p.spinner();
    delSpin.start("Deleting user...");
    const { error: delError } = await ps.runCommand(
      `Remove-MgUser -UserId '${escapePS(userId)}'`,
    );
    if (delError) {
      delSpin.stop("Failed to delete user.");
      p.log.error(delError);
      p.log.warn("The user was NOT deleted. Other actions above may have already been applied.");
    } else {
      delSpin.stop("User deleted.");
      userDeleted = true;
    }
  }

  // ── RESULTS SUMMARY ───────────────────────────────────────────────────

  const summaryParts: string[] = [];

  if (userDisabled) {
    summaryParts.push(`Converted to shared mailbox: ${user.DisplayName} (${upn})`);
  } else if (userDeleted) {
    summaryParts.push(`Deleted user: ${user.DisplayName} (${upn})`);
  } else {
    summaryParts.push(`FAILED to ${convertedToShared ? "convert" : "delete"} user: ${user.DisplayName} (${upn})`);
  }

  if (!convertedToShared && releasedLicenses.length > 0) {
    summaryParts.push(
      `\nLicenses released:\n${releasedLicenses.map((l) => `  - ${l}`).join("\n")}`,
    );
  }

  if (convertedToShared) {
    summaryParts.push("\nMailbox: Converted to shared");
  }
  if (mailboxDelegateNames.length > 0) {
    summaryParts.push(`Mailbox delegates: ${mailboxDelegateNames.join(", ")}`);
  }

  if (removedMailboxes.length > 0) {
    summaryParts.push(
      `\nRemoved from shared mailbox(es):\n${removedMailboxes.map((n) => `  - ${n}`).join("\n")}`,
    );
  }

  if (removedDistGroups.length > 0) {
    summaryParts.push(
      `\nRemoved from distribution group(s):\n${removedDistGroups.map((n) => `  - ${n}`).join("\n")}`,
    );
  }

  if (removedSecGroups.length > 0) {
    summaryParts.push(
      `\nRemoved from security group(s):\n${removedSecGroups.map((n) => `  - ${n}`).join("\n")}`,
    );
  }

  if (orphans.length > 0) {
    summaryParts.push(
      `\nOrphaned resources (no remaining members):\n${orphans.map((o) => `  - [${o.type}] ${o.name}`).join("\n")}`,
    );
  }

  p.note(summaryParts.join("\n"), userDisabled ? "Conversion summary" : "Deletion summary");

  if (userDisabled) {
    p.log.success("Account converted to shared mailbox.");
  } else if (userDeleted) {
    p.log.success("User deleted successfully.");
    p.log.info(
      "The user can be restored from the Entra ID recycle bin within 30 days.",
    );
  } else {
    p.log.error(`User ${convertedToShared ? "conversion" : "deletion"} failed. Review the summary above for actions already taken.`);
  }
}
