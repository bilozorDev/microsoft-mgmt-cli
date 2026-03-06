import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { friendlySkuName } from "../sku-names.ts";

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

function escapePS(value: string): string {
  return value.replace(/'/g, "''");
}

function formatStorageSize(mb: number): string {
  if (mb >= 1024) return `${(mb / 1024).toFixed(1)} GB`;
  return `${mb} MB`;
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
    users = Array.isArray(raw) ? raw : [raw];
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
        users = Array.isArray(raw) ? raw : [raw];
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
  }

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
  const users = await fetchUsers(ps, excludeUpn);
  if (users.length === 0) return [];

  const selected = await p.multiselect({
    message,
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

  // Step 1: Select user to delete
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
    licenses = Array.isArray(raw) ? raw : [raw];
    licSpin.stop("User details loaded.");
  } catch {
    licSpin.stop("User details loaded.");
    // No licenses or error — continue
  }

  const licenseLines =
    licenses.length > 0
      ? licenses.map((l) => `  - ${friendlySkuName(l.SkuPartNumber)}`).join("\n")
      : "  (none)";

  p.note(
    [
      `Name:     ${user.DisplayName}`,
      `UPN:      ${upn}`,
      `ID:       ${userId}`,
      `Licenses:\n${licenseLines}`,
    ].join("\n"),
    "User to delete",
  );

  const proceed = await p.confirm({ message: "Proceed with deleting this user?" });
  if (p.isCancel(proceed) || !proceed) {
    p.log.info("Cancelled.");
    return;
  }

  // Tracking arrays for summary
  const releasedLicenses = licenses.map((l) => friendlySkuName(l.SkuPartNumber));
  let convertedToShared = false;
  let mailboxDelegateNames: string[] = [];
  let oneDriveDelegateNames: string[] = [];
  let oneDriveSize: string | null = null;
  const removedMailboxes: string[] = [];
  const removedDistGroups: string[] = [];
  const removedSecGroups: string[] = [];
  const orphans: { type: string; name: string }[] = [];

  // Step 2: Convert to shared mailbox (optional)
  const mbSpin = p.spinner();
  mbSpin.start("Checking mailbox...");
  let hasExchangeMailbox = false;

  const { output: mbOutput } = await ps.runCommand(
    `Get-Mailbox -Identity '${escapePS(upn)}' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty RecipientTypeDetails`,
  );
  const mbType = mbOutput.trim();
  mbSpin.stop(mbType ? `Mailbox type: ${mbType}` : "No Exchange mailbox found.");

  if (mbType === "UserMailbox") {
    hasExchangeMailbox = true;

    const convert = await p.confirm({
      message: "Convert mailbox to shared before deletion?",
    });
    if (p.isCancel(convert)) return;

    if (convert) {
      const convertSpin = p.spinner();
      convertSpin.start("Converting to shared mailbox...");
      const { error } = await ps.runCommand(
        `Set-Mailbox -Identity '${escapePS(upn)}' -Type Shared`,
      );
      if (error) {
        convertSpin.stop("Failed to convert mailbox.");
        p.log.error(error);
      } else {
        convertSpin.stop("Mailbox converted to shared.");
        convertedToShared = true;
      }

      const grantAccess = await p.confirm({
        message: "Grant other users access to this shared mailbox?",
      });
      if (p.isCancel(grantAccess)) return;

      if (grantAccess) {
        const delegates = await selectMultipleUsers(
          ps,
          "Select delegate(s) for mailbox access",
          upn,
        );

        for (const delegate of delegates) {
          const dSpin = p.spinner();
          dSpin.start(`Granting access to ${delegate.DisplayName}...`);
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
          } else if (errors.length === 1) {
            dSpin.stop(`Partially granted access to ${delegate.DisplayName}.`);
            mailboxDelegateNames.push(delegate.DisplayName);
            for (const err of errors) p.log.error(err);
          } else {
            dSpin.stop(`Failed to grant access to ${delegate.DisplayName}.`);
            for (const err of errors) p.log.error(err);
          }
        }
      }
    }
  } else if (mbType === "SharedMailbox") {
    hasExchangeMailbox = true;
    p.log.info("Mailbox is already shared — skipping conversion.");
  }

  // Step 2b: Share OneDrive (optional)
  const shareDrive = await p.confirm({
    message: "Grant another user access to this user's OneDrive?",
  });
  if (p.isCancel(shareDrive)) return;

  if (shareDrive) {
    const spoSpin = p.spinner();
    spoSpin.start("Connecting to SharePoint Online (check your browser)...");
    let spoConnected = false;
    try {
      await ps.ensureSPOConnected();
      spoSpin.stop("Connected to SharePoint Online.");
      spoConnected = true;
    } catch (e) {
      spoSpin.stop("Failed to connect to SharePoint Online.");
      p.log.error(`${e}`);
      p.log.warn("Skipping OneDrive sharing.");
    }

    if (spoConnected) {
      // Construct OneDrive URL
      const { output: tenantDomain } = await ps.runCommand(
        "Get-AcceptedDomain | Where-Object { $_.DomainName -like '*.onmicrosoft.com' -and $_.DomainName -notlike '*.mail.onmicrosoft.com' } | Select-Object -ExpandProperty DomainName",
      );
      const tenantName = tenantDomain.trim().replace(".onmicrosoft.com", "");
      const personalPath = upn.replace(/[^a-zA-Z0-9]/g, "_");
      const oneDriveUrl = `https://${tenantName}-my.sharepoint.com/personal/${personalPath}`;

      const sizeSpin = p.spinner();
      sizeSpin.start("Checking OneDrive size...");
      const { output: sizeOutput, error: sizeError } = await ps.runCommand(
        `Get-SPOSite -Identity '${escapePS(oneDriveUrl)}' | Select-Object -ExpandProperty StorageUsageCurrent`,
      );

      if (sizeError) {
        sizeSpin.stop("OneDrive not found or not provisioned.");
        p.log.info("Skipping OneDrive sharing.");
      } else {
        const sizeMB = parseInt(sizeOutput.trim(), 10);
        const sizeFormatted = isNaN(sizeMB) ? "unknown size" : formatStorageSize(sizeMB);
        oneDriveSize = sizeFormatted;
        sizeSpin.stop(`OneDrive is using ${sizeFormatted}.`);

        const delegates = await selectMultipleUsers(
          ps,
          "Select delegate(s) for OneDrive access",
          upn,
        );

        for (const delegate of delegates) {
          const dSpin = p.spinner();
          dSpin.start(`Granting OneDrive access to ${delegate.DisplayName}...`);
          const { error } = await ps.runCommand(
            `Set-SPOUser -Site '${escapePS(oneDriveUrl)}' -LoginName '${escapePS(delegate.UserPrincipalName)}' -IsSiteCollectionAdmin $true`,
          );
          if (error) {
            dSpin.stop(`Failed to grant OneDrive access to ${delegate.DisplayName}.`);
            p.log.error(error);
          } else {
            dSpin.stop(`Granted OneDrive access to ${delegate.DisplayName}.`);
            oneDriveDelegateNames.push(delegate.DisplayName);
          }
        }
      }
    }
  }

  // Step 3: Scan phase
  let sharedMailboxPerms: SharedMailboxPerm[] = [];
  let distGroupMemberships: DistGroupMembership[] = [];
  let secGroupMemberships: SecGroupMembership[] = [];

  // 3a. Shared mailbox permissions
  if (hasExchangeMailbox) {
    const scanMbSpin = p.spinner();
    scanMbSpin.start("Scanning shared mailbox permissions...");
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
      sharedMailboxPerms = Array.isArray(raw) ? raw : [raw];
      scanMbSpin.stop(
        `Found permissions on ${sharedMailboxPerms.length} shared mailbox(es).`,
      );
    } catch {
      scanMbSpin.stop("No shared mailbox permissions found.");
    }
  }

  // 3b. Distribution group memberships
  {
    const scanDgSpin = p.spinner();
    scanDgSpin.start("Scanning distribution group memberships...");
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
      distGroupMemberships = Array.isArray(raw) ? raw : [raw];
      scanDgSpin.stop(
        `Found ${distGroupMemberships.length} distribution group membership(s).`,
      );
    } catch {
      scanDgSpin.stop("No distribution group memberships found.");
    }
  }

  // 3c. Security group memberships
  {
    const scanSgSpin = p.spinner();
    scanSgSpin.start("Scanning security group memberships...");
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
      secGroupMemberships = Array.isArray(raw) ? raw : [raw];
      scanSgSpin.stop(
        `Found ${secGroupMemberships.length} security group membership(s).`,
      );
    } catch {
      scanSgSpin.stop("No security group memberships found.");
    }
  }

  // Step 4: Display findings & orphan warnings
  const totalFindings =
    sharedMailboxPerms.length +
    distGroupMemberships.length +
    secGroupMemberships.length;

  if (totalFindings > 0) {
    const lines: string[] = [];
    if (sharedMailboxPerms.length > 0)
      lines.push(
        `Shared mailbox permissions: ${sharedMailboxPerms.length}`,
      );
    if (distGroupMemberships.length > 0)
      lines.push(
        `Distribution group memberships: ${distGroupMemberships.length}`,
      );
    if (secGroupMemberships.length > 0)
      lines.push(
        `Security group memberships: ${secGroupMemberships.length}`,
      );
    p.note(lines.join("\n"), "Memberships to remove");
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

  // Step 5: Removal phase
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
        // May fail for dynamic groups
        p.log.error(`${sg.DisplayName}: ${error}`);
      } else {
        removedSecGroups.push(sg.DisplayName);
      }
    }
    rmSgSpin.stop(`Removed from ${removedSecGroups.length} security group(s).`);
  }

  // Step 6: Delete user
  const delSpin = p.spinner();
  delSpin.start("Deleting user...");
  const { error: delError } = await ps.runCommand(
    `Remove-MgUser -UserId '${escapePS(userId)}'`,
  );
  if (delError) {
    delSpin.stop("Failed to delete user.");
    p.log.error(delError);
    return;
  }
  delSpin.stop("User deleted.");

  // Step 7: Final summary
  const summaryParts: string[] = [
    `Deleted user: ${user.DisplayName} (${upn})`,
  ];

  if (releasedLicenses.length > 0) {
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

  if (oneDriveDelegateNames.length > 0) {
    const sizeLabel = oneDriveSize ? ` (${oneDriveSize})` : "";
    summaryParts.push(`\nOneDrive${sizeLabel}: Access granted`);
    summaryParts.push(`OneDrive delegates: ${oneDriveDelegateNames.join(", ")}`);
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

  p.note(summaryParts.join("\n"), "Deletion summary");

  p.log.success("User deleted successfully.");
  p.log.info(
    "The user can be restored from the Entra ID recycle bin within 30 days.",
  );
}
