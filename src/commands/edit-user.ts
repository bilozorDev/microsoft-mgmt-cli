import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generatePassword, validatePassword } from "../password.ts";
import { friendlySkuName } from "../sku-names.ts";
import { escapePS, createSecretLink } from "../utils.ts";
import {
  type AuthMethod,
  MFA_DETAIL_CMDS,
  MFA_REMOVE_CMDLETS,
  friendlyMfaMethod,
  mfaTypeKey,
} from "../mfa-utils.ts";
import { run as addToDistributionGroup } from "./add-to-distribution-group.ts";
import { run as addToSecurityGroup } from "./add-to-security-group.ts";
import { run as addToSharedMailbox } from "./add-to-shared-mailbox.ts";

interface MgUser {
  DisplayName: string;
  UserPrincipalName: string;
  Id: string;
  LicenseCount: number;
}

interface MgUserDetails {
  DisplayName: string;
  GivenName: string | null;
  Surname: string | null;
  UserPrincipalName: string;
  AccountEnabled: boolean;
  Id: string;
}

interface LicenseDetail {
  SkuPartNumber: string;
  SkuId: string;
}

interface AcceptedDomain {
  DomainName: string;
  Default: boolean;
}

interface RoleDefinition {
  Id: string;
  DisplayName: string;
  Description: string | null;
}

interface RoleAssignment {
  Id: string;
  RoleDefinitionId: string;
  PrincipalId: string;
}

interface SubscribedSku {
  SkuId: string;
  SkuPartNumber: string;
  ConsumedUnits: number;
  PrepaidUnits: { Enabled: number; [key: string]: unknown };
}

interface ForwardingConfig {
  ForwardingSmtpAddress: string | null;
  ForwardingAddress: string | null;
  DeliverToMailboxAndForward: boolean;
}

interface MailboxPermission {
  User: string;
  AccessRights: string | string[];
}

interface RecipientPermission {
  Trustee: string;
  AccessRights: string | string[];
}

async function fetchUsers(ps: PowerShellSession): Promise<MgUser[]> {
  const spin = p.spinner();
  spin.start("Loading users...");

  const { output: countOutput } = await ps.runCommand(
    "Get-MgUser -Top 1 -CountVariable ct -ConsistencyLevel eventual | Out-Null; $ct",
  );
  const count = parseInt(countOutput.trim(), 10);

  let users: MgUser[];

  if (count <= 50) {
    const raw = await ps.runCommandJson<MgUser | MgUser[]>(
      "Get-MgUser -All -Property DisplayName,UserPrincipalName,Id,AssignedLicenses | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName; Id = $_.Id; LicenseCount = $_.AssignedLicenses.Count } }",
    );
    users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
    spin.stop(`Found ${users.length} user(s).`);
  } else {
    spin.stop(`${count} users in tenant — search to find a user.`);
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
          `Get-MgUser -Search '"displayName:${escapePS(query)}"' -ConsistencyLevel eventual -Property DisplayName,UserPrincipalName,Id,AssignedLicenses | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName; Id = $_.Id; LicenseCount = $_.AssignedLicenses.Count } }`,
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

async function selectUser(
  ps: PowerShellSession,
  message: string,
): Promise<MgUser | null> {
  const users = await fetchUsers(ps);
  if (users.length === 0) return null;

  const userId = await p.select({
    message,
    options: users.map((u) => ({
      value: u.Id,
      label: u.DisplayName,
      hint: u.LicenseCount > 0 ? u.UserPrincipalName : `${u.UserPrincipalName} (not licensed)`,
    })),
  });
  if (p.isCancel(userId)) return null;

  return users.find((u) => u.Id === userId) ?? null;
}

async function fetchUserDetails(
  ps: PowerShellSession,
  userId: string,
): Promise<{ details: MgUserDetails; licenses: LicenseDetail[] } | null> {
  const spin = p.spinner();
  spin.start("Fetching user details...");

  let details: MgUserDetails;
  try {
    const raw = await ps.runCommandJson<MgUserDetails>(
      `Get-MgUser -UserId '${escapePS(userId)}' -Property DisplayName,GivenName,Surname,UserPrincipalName,AccountEnabled,Id | Select-Object DisplayName,GivenName,Surname,UserPrincipalName,AccountEnabled,Id`,
    );
    if (!raw) {
      spin.stop("Failed to fetch user details.");
      p.log.error("User not found.");
      return null;
    }
    details = raw;
  } catch (e) {
    spin.stop("Failed to fetch user details.");
    p.log.error(`${e}`);
    return null;
  }

  let licenses: LicenseDetail[] = [];
  try {
    const raw = await ps.runCommandJson<LicenseDetail | LicenseDetail[]>(
      `Get-MgUserLicenseDetail -UserId '${escapePS(userId)}' | Select-Object SkuPartNumber, SkuId`,
    );
    licenses = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
  } catch {
    // no licenses
  }

  spin.stop("User details loaded.");
  return { details, licenses };
}

function displayUserDetails(details: MgUserDetails, licenses: LicenseDetail[]) {
  p.note(
    [
      `Name:      ${details.DisplayName}`,
      `First:     ${details.GivenName ?? "(not set)"}`,
      `Last:      ${details.Surname ?? "(not set)"}`,
      `UPN:       ${details.UserPrincipalName}`,
      `Licenses:  ${licenses.length > 0 ? licenses.map((l) => friendlySkuName(l.SkuPartNumber)).join(", ") : "(none)"}`,
      `Status:    ${details.AccountEnabled ? "Enabled" : "Disabled"}`,
    ].join("\n"),
    "User details",
  );
}

export async function run(ps: PowerShellSession): Promise<void> {
  // 1. Ensure Graph connected
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph (check your browser)...");
  try {
    await ps.ensureGraphConnected(["User.ReadWrite.All", "Organization.Read.All", "GroupMember.ReadWrite.All", "User-PasswordProfile.ReadWrite.All", "RoleManagement.ReadWrite.Directory", "UserAuthenticationMethod.ReadWrite.All"]);
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return;
  }

  // 2. Select user
  const user = await selectUser(ps, "Select user to edit");
  if (!user) return;

  const userId = user.Id;
  let upn = user.UserPrincipalName;

  // 3. Show current details
  let current = await fetchUserDetails(ps, userId);
  if (!current) return;

  displayUserDetails(current.details, current.licenses);

  // 4. Edit menu loop
  let passwordWasReset = false;
  let nameChangedTo: string | null = null;
  const licensesAdded: string[] = [];
  const licensesRemoved: string[] = [];
  const addedDistGroups: { name: string; email: string }[] = [];
  const addedSecGroups: { name: string; email: string }[] = [];
  const addedMailboxes: { name: string; email: string }[] = [];
  let otsUrl: string | null = null;
  let domainChangedFrom: string | null = null;
  const aliasesAdded: string[] = [];
  const rolesAdded: string[] = [];
  const rolesRemoved: string[] = [];
  let signInToggled: "blocked" | "unblocked" | null = null;
  const delegationsAdded: string[] = [];
  const delegationsRemoved: string[] = [];
  let forwardingSet: string | null = null;
  let forwardingRemoved = false;
  const mfaMethodsRemoved: string[] = [];
  let acceptedDomains: AcceptedDomain[] | null = null;

  while (true) {
    const menuOptions: { value: string; label: string; hint?: string }[] = [
      { value: "edit-name", label: "Edit name" },
      { value: "change-domain", label: "Change primary domain", hint: upn.split("@")[1] },
      { value: "add-alias", label: "Add email alias" },
      { value: "reset-password", label: "Reset password" },
      { value: "toggle-signin", label: current.details.AccountEnabled ? "Block sign-in" : "Unblock sign-in" },
      { value: "manage-licenses", label: "Manage licenses" },
      { value: "manage-roles", label: "Manage admin roles" },
      { value: "reset-mfa", label: "Reset MFA methods" },
      { value: "manage-delegation", label: "Manage mailbox delegation" },
      { value: "email-forwarding", label: "Email forwarding" },
      { value: "add-distribution-group", label: "Add to distribution group" },
      { value: "add-security-group", label: "Add to security group" },
      { value: "add-shared-mailbox", label: "Add to shared mailbox" },
    ];
    if (otsUrl) {
      menuOptions.push({ value: "copy-ots", label: "Copy OTS link to clipboard" });
    }
    menuOptions.push({ value: "copy-ticket", label: "Copy ticket update note to clipboard" });
    menuOptions.push({ value: "done", label: "Done" });

    const action = await p.select({
      message: "Edit action",
      options: menuOptions,
    });
    if (p.isCancel(action) || action === "done") break;

    switch (action) {
      case "edit-name": {
        const firstName = await p.text({
          message: "First name",
          initialValue: current.details.GivenName ?? "",
          validate: (v = "") => (!v.trim() ? "First name is required" : undefined),
        });
        if (p.isCancel(firstName)) break;

        const lastName = await p.text({
          message: "Last name",
          initialValue: current.details.Surname ?? "",
          validate: (v = "") => (!v.trim() ? "Last name is required" : undefined),
        });
        if (p.isCancel(lastName)) break;

        const displayName = await p.text({
          message: "Display name",
          initialValue: `${firstName} ${lastName}`,
          validate: (v = "") => (!v.trim() ? "Display name is required" : undefined),
        });
        if (p.isCancel(displayName)) break;

        const spin = p.spinner();
        spin.start("Updating name...");
        const { error } = await ps.runCommand(
          `Update-MgUser -UserId '${escapePS(userId)}' -DisplayName '${escapePS(displayName)}' -GivenName '${escapePS(firstName)}' -Surname '${escapePS(lastName)}'`,
        );
        if (error) {
          spin.stop("Failed to update name.");
          p.log.error(error);
        } else {
          spin.stop("Name updated.");
          nameChangedTo = displayName;
          current.details.GivenName = firstName;
          current.details.Surname = lastName;
          current.details.DisplayName = displayName;
        }
        break;
      }

      case "reset-password": {
        const pwMethod = await p.select({
          message: "Password",
          options: [
            { value: "auto", label: "Auto-generate", hint: "16 chars, crypto-random" },
            { value: "manual", label: "Enter manually" },
          ],
        });
        if (p.isCancel(pwMethod)) break;

        let password: string;
        if (pwMethod === "auto") {
          password = generatePassword();
        } else {
          const manualPw = await p.password({
            message: "Enter password",
            validate: (v) => validatePassword(v ?? ""),
          });
          if (p.isCancel(manualPw)) break;
          password = manualPw;
        }

        const spin = p.spinner();
        spin.start("Resetting password...");
        const { error } = await ps.runCommand(
          `Update-MgUser -UserId '${escapePS(userId)}' -PasswordProfile @{ Password = '${escapePS(password)}'; ForceChangePasswordNextSignIn = $true }`,
        );
        if (error) {
          spin.stop("Failed to reset password.");
          p.log.error(error);
        } else {
          spin.stop("Password reset.");
          p.note(
            [`UPN:      ${upn}`, `Password: ${password}`].join("\n"),
            "New credentials (user must change password at next sign-in)",
          );
          passwordWasReset = true;

          const otsSpin = p.spinner();
          otsSpin.start("Creating one-time secret link...");
          const otsResult = await createSecretLink(`${password}`);
          if ("url" in otsResult) {
            otsUrl = otsResult.url;
            otsSpin.stop("One-time secret link created.");
            p.log.info(`Secret link: ${otsUrl}`);
          } else {
            otsSpin.stop("Failed to create one-time secret link.");
            p.log.error("error" in otsResult ? otsResult.error : "Unknown error");
          }
        }
        break;
      }

      case "manage-licenses": {
        const skuSpin = p.spinner();
        skuSpin.start("Fetching available licenses...");

        let skus: SubscribedSku[];
        try {
          const raw = await ps.runCommandJson<SubscribedSku | SubscribedSku[]>(
            "Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, PrepaidUnits",
          );
          skus = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
          skuSpin.stop(`Found ${skus.length} license type(s).`);
        } catch (e) {
          skuSpin.stop("Failed to fetch licenses.");
          p.log.error(`${e}`);
          break;
        }

        if (skus.length === 0) {
          p.log.warn("No license types found in tenant.");
          break;
        }

        const currentSkuIds = new Set(current.licenses.map((l) => l.SkuId));

        const choices = await p.multiselect({
          message: "Toggle licenses (space to toggle, enter to confirm)",
          options: skus.map((s) => {
            const avail = s.PrepaidUnits.Enabled - s.ConsumedUnits;
            const assigned = currentSkuIds.has(s.SkuId);
            let hint: string;
            if (assigned && avail <= 0) {
              hint = "assigned — 0 seats remaining";
            } else {
              hint = `${avail} of ${s.PrepaidUnits.Enabled} available`;
            }
            return {
              value: s.SkuId,
              label: friendlySkuName(s.SkuPartNumber),
              hint,
              disabled: avail <= 0 && !assigned,
            };
          }),
          initialValues: skus.filter((s) => currentSkuIds.has(s.SkuId)).map((s) => s.SkuId),
          required: false,
        });
        if (p.isCancel(choices)) break;

        const selectedSet = new Set(choices);
        const toAdd = skus.filter((s) => selectedSet.has(s.SkuId) && !currentSkuIds.has(s.SkuId));
        const toRemove = skus.filter((s) => !selectedSet.has(s.SkuId) && currentSkuIds.has(s.SkuId));

        if (toAdd.length === 0 && toRemove.length === 0) {
          p.log.info("No changes.");
          break;
        }

        // Show summary and confirm
        const summaryLines: string[] = [];
        if (toAdd.length > 0) {
          summaryLines.push(`Add:    ${toAdd.map((s) => friendlySkuName(s.SkuPartNumber)).join(", ")}`);
        }
        if (toRemove.length > 0) {
          summaryLines.push(`Remove: ${toRemove.map((s) => friendlySkuName(s.SkuPartNumber)).join(", ")}`);
        }
        p.note(summaryLines.join("\n"), "License changes");

        // Extra warning if removing ALL licenses
        if (toRemove.length > 0 && toRemove.length === current.licenses.length && toAdd.length === 0) {
          const warnConfirm = await p.confirm({
            message: "This will remove ALL licenses. User may lose access to data. Are you sure?",
          });
          if (p.isCancel(warnConfirm) || !warnConfirm) break;
        } else {
          const confirm = await p.confirm({ message: "Apply these changes?" });
          if (p.isCancel(confirm) || !confirm) break;
        }

        const spin = p.spinner();
        spin.start("Updating licenses...");

        const addEntries = toAdd.length > 0
          ? `@(${toAdd.map((s) => `@{SkuId = '${escapePS(s.SkuId)}'}`).join(", ")})`
          : "@()";
        const removeEntries = toRemove.length > 0
          ? `@(${toRemove.map((s) => `'${escapePS(s.SkuId)}'`).join(",")})`
          : "@()";
        const { error } = await ps.runCommand(
          `Set-MgUserLicense -UserId '${escapePS(userId)}' -AddLicenses ${addEntries} -RemoveLicenses ${removeEntries}`,
        );
        if (error) {
          spin.stop("Failed to update licenses.");
          p.log.error(error);
        } else {
          const parts: string[] = [];
          if (toAdd.length > 0) {
            const addedNames = toAdd.map((s) => friendlySkuName(s.SkuPartNumber));
            parts.push(`Added: ${addedNames.join(", ")}`);
            licensesAdded.push(...addedNames);
            for (const s of toAdd) {
              current.licenses.push({ SkuId: s.SkuId, SkuPartNumber: s.SkuPartNumber });
            }
          }
          if (toRemove.length > 0) {
            const removedNames = toRemove.map((s) => friendlySkuName(s.SkuPartNumber));
            parts.push(`Removed: ${removedNames.join(", ")}`);
            licensesRemoved.push(...removedNames);
            const removeIds = new Set(toRemove.map((s) => s.SkuId));
            current.licenses = current.licenses.filter((l) => !removeIds.has(l.SkuId));
          }
          spin.stop(parts.join(". ") + ".");
        }
        break;
      }

      case "manage-roles": {
        const roleSpin = p.spinner();
        roleSpin.start("Fetching admin roles…");

        let roleDefs: RoleDefinition[];
        try {
          const raw = await ps.runCommandJson<RoleDefinition | RoleDefinition[]>(
            `Get-MgRoleManagementDirectoryRoleDefinition -All | Select-Object Id, DisplayName, Description`,
          );
          roleDefs = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        } catch (e) {
          roleSpin.stop("Failed to fetch roles.");
          p.log.error(`${e}`);
          break;
        }

        // Get current assignments for this user
        let currentAssignments: RoleAssignment[];
        try {
          const raw = await ps.runCommandJson<RoleAssignment | RoleAssignment[]>(
            `Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '${escapePS(userId)}'" -All | Select-Object Id, RoleDefinitionId, PrincipalId`,
          );
          currentAssignments = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        } catch {
          currentAssignments = [];
        }

        roleSpin.stop(`Found ${roleDefs.length} role(s), user has ${currentAssignments.length} assigned.`);

        const currentRoleIds = new Set(currentAssignments.map((a) => a.RoleDefinitionId));

        // Sort: assigned first, then alphabetical
        roleDefs.sort((a, b) => {
          const aAssigned = currentRoleIds.has(a.Id) ? 0 : 1;
          const bAssigned = currentRoleIds.has(b.Id) ? 0 : 1;
          if (aAssigned !== bAssigned) return aAssigned - bAssigned;
          return a.DisplayName.localeCompare(b.DisplayName);
        });

        const choices = await p.multiselect({
          message: "Toggle admin roles (space to toggle, enter to confirm)",
          options: roleDefs.map((r) => ({
            value: r.Id,
            label: r.DisplayName,
            hint: r.Description ? r.Description.slice(0, 60) : undefined,
          })),
          initialValues: roleDefs.filter((r) => currentRoleIds.has(r.Id)).map((r) => r.Id),
          required: false,
        });
        if (p.isCancel(choices)) break;

        const selectedSet = new Set(choices);
        const toAdd = roleDefs.filter((r) => selectedSet.has(r.Id) && !currentRoleIds.has(r.Id));
        const toRemove = roleDefs.filter((r) => !selectedSet.has(r.Id) && currentRoleIds.has(r.Id));

        if (toAdd.length === 0 && toRemove.length === 0) {
          p.log.info("No changes.");
          break;
        }

        const summaryLines: string[] = [];
        if (toAdd.length > 0) summaryLines.push(`Add:    ${toAdd.map((r) => r.DisplayName).join(", ")}`);
        if (toRemove.length > 0) summaryLines.push(`Remove: ${toRemove.map((r) => r.DisplayName).join(", ")}`);
        p.note(summaryLines.join("\n"), "Role changes");

        const confirm = await p.confirm({ message: "Apply these changes?" });
        if (p.isCancel(confirm) || !confirm) break;

        const spin = p.spinner();
        spin.start("Updating roles…");

        const runRolesAdded: string[] = [];
        const runRolesRemoved: string[] = [];

        // Add roles
        for (const r of toAdd) {
          spin.message(`Adding ${r.DisplayName}…`);
          const { error } = await ps.runCommand(
            `New-MgRoleManagementDirectoryRoleAssignment -RoleDefinitionId '${escapePS(r.Id)}' -PrincipalId '${escapePS(userId)}' -DirectoryScopeId '/'`,
          );
          if (error) {
            p.log.error(`Failed to add ${r.DisplayName}: ${error}`);
          } else {
            runRolesAdded.push(r.DisplayName);
            rolesAdded.push(r.DisplayName);
          }
        }

        // Remove roles
        for (const r of toRemove) {
          spin.message(`Removing ${r.DisplayName}…`);
          const assignment = currentAssignments.find((a) => a.RoleDefinitionId === r.Id);
          if (!assignment) continue;
          const { error } = await ps.runCommand(
            `Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId '${escapePS(assignment.Id)}'`,
          );
          if (error) {
            p.log.error(`Failed to remove ${r.DisplayName}: ${error}`);
          } else {
            runRolesRemoved.push(r.DisplayName);
            rolesRemoved.push(r.DisplayName);
          }
        }

        const parts: string[] = [];
        if (runRolesAdded.length > 0) parts.push(`Added: ${runRolesAdded.join(", ")}`);
        if (runRolesRemoved.length > 0) parts.push(`Removed: ${runRolesRemoved.join(", ")}`);
        spin.stop(parts.length > 0 ? parts.join(". ") + "." : "No changes applied.");
        break;
      }

      case "toggle-signin": {
        const newState = current.details.AccountEnabled ? false : true;
        const label = newState ? "Unblocking" : "Blocking";
        const spin = p.spinner();
        spin.start(`${label} sign-in...`);
        const { error } = await ps.runCommand(
          `Update-MgUser -UserId '${escapePS(userId)}' -AccountEnabled:$${newState}`,
        );
        if (error) {
          spin.stop(`Failed to ${label.toLowerCase()} sign-in.`);
          p.log.error(error);
        } else {
          current.details.AccountEnabled = newState;
          signInToggled = newState ? "unblocked" : "blocked";
          spin.stop(`Sign-in ${newState ? "unblocked" : "blocked"}.`);
        }
        break;
      }

      case "reset-mfa": {
        const mfaSpin = p.spinner();
        mfaSpin.start("Fetching MFA methods...");

        const raw = await ps.runCommandJson<AuthMethod | AuthMethod[]>(
          `Get-MgUserAuthenticationMethod -UserId '${escapePS(userId)}' | ForEach-Object { [PSCustomObject]@{ Id = $_.Id; ODataType = $_.AdditionalProperties['@odata.type'] } }`,
        );
        const methods = raw ? (Array.isArray(raw) ? raw : [raw]) : [];

        const removable = methods.filter((m) => {
          if (!m.ODataType) return false;
          const key = mfaTypeKey(m.ODataType);
          return key !== "passwordAuthenticationMethod" && key in MFA_REMOVE_CMDLETS;
        });

        if (removable.length === 0) {
          mfaSpin.stop("No removable MFA methods found (only password).");
          break;
        }

        // Fetch details for each method
        mfaSpin.message("Fetching method details...");
        const methodDetails: { method: AuthMethod; friendly: string; detail: string }[] = [];

        for (const m of removable) {
          const key = mfaTypeKey(m.ODataType!);
          const friendly = friendlyMfaMethod(m.ODataType!) ?? key;
          let detail = "";

          const detailFetcher = MFA_DETAIL_CMDS[key];
          if (detailFetcher) {
            try {
              const detailRaw = await ps.runCommandJson<Record<string, unknown>>(
                detailFetcher.cmd(userId, m.Id),
              );
              if (detailRaw) detail = detailFetcher.format(detailRaw);
            } catch {
              // detail stays empty
            }
          }

          methodDetails.push({ method: m, friendly, detail });
        }

        mfaSpin.stop(`Found ${removable.length} MFA method(s).`);

        const choices = await p.multiselect({
          message: "Select MFA methods to remove (space to toggle, enter to confirm)",
          options: methodDetails.map((d) => ({
            value: d.method.Id,
            label: d.friendly,
            hint: d.detail || undefined,
          })),
          required: false,
        });
        if (p.isCancel(choices) || choices.length === 0) break;

        const removeSpin = p.spinner();
        removeSpin.start("Removing selected MFA methods...");

        const runMfaRemoved: string[] = [];
        for (const methodId of choices) {
          const detail = methodDetails.find((d) => d.method.Id === methodId)!;
          const key = mfaTypeKey(detail.method.ODataType!);
          const info = MFA_REMOVE_CMDLETS[key]!;

          removeSpin.message(`Removing ${detail.friendly}...`);
          const { error } = await ps.runCommand(
            `${info.cmdlet} -UserId '${escapePS(userId)}' ${info.param} '${escapePS(methodId)}'`,
          );
          if (error) {
            p.log.error(`Failed to remove ${detail.friendly}: ${error}`);
          } else {
            runMfaRemoved.push(detail.friendly);
            mfaMethodsRemoved.push(detail.friendly);
          }
        }

        removeSpin.stop(
          runMfaRemoved.length > 0
            ? `Removed ${runMfaRemoved.length} MFA method(s).`
            : "No methods removed.",
        );
        break;
      }

      case "manage-delegation": {
        const delegAction = await p.select({
          message: "Mailbox delegation",
          options: [
            { value: "add", label: "Add delegation" },
            { value: "remove", label: "Remove delegation" },
            { value: "back", label: "Back" },
          ],
        });
        if (p.isCancel(delegAction) || delegAction === "back") break;

        if (delegAction === "add") {
          const delegate = await selectUser(ps, "Select delegate user");
          if (!delegate) break;

          const permTypes = await p.multiselect({
            message: "Select permission types to grant",
            options: [
              { value: "FullAccess", label: "Full Access", hint: "read/manage mailbox" },
              { value: "SendAs", label: "Send As", hint: "send email as this user" },
              { value: "SendOnBehalf", label: "Send on Behalf", hint: "send on behalf of this user" },
            ],
            required: true,
          });
          if (p.isCancel(permTypes)) break;

          const delegSpin = p.spinner();
          delegSpin.start("Adding delegation...");
          const addedCount = delegationsAdded.length;

          for (const perm of permTypes) {
            delegSpin.message(`Adding ${perm}...`);
            let error: string;
            if (perm === "FullAccess") {
              ({ error } = await ps.runCommand(
                `Add-MailboxPermission -Identity '${escapePS(upn)}' -User '${escapePS(delegate.UserPrincipalName)}' -AccessRights FullAccess -InheritanceType All -Confirm:$false`,
              ));
            } else if (perm === "SendAs") {
              ({ error } = await ps.runCommand(
                `Add-RecipientPermission -Identity '${escapePS(upn)}' -Trustee '${escapePS(delegate.UserPrincipalName)}' -AccessRights SendAs -Confirm:$false`,
              ));
            } else {
              ({ error } = await ps.runCommand(
                `Set-Mailbox -Identity '${escapePS(upn)}' -GrantSendOnBehalfTo @{Add='${escapePS(delegate.UserPrincipalName)}'}`,
              ));
            }
            if (error) {
              p.log.error(`Failed to add ${perm}: ${error}`);
            } else {
              delegationsAdded.push(`${perm} → ${delegate.DisplayName}`);
            }
          }

          const newlyAdded = delegationsAdded.length - addedCount;
          delegSpin.stop(
            newlyAdded > 0
              ? `Added ${newlyAdded} delegation(s).`
              : "No delegations added.",
          );
        } else {
          // Remove delegation
          const fetchSpin = p.spinner();
          fetchSpin.start("Fetching current delegations...");

          const allDelegations: { value: string; label: string; hint: string }[] = [];

          // Full Access
          try {
            const faRaw = await ps.runCommandJson<MailboxPermission | MailboxPermission[]>(
              `Get-MailboxPermission -Identity '${escapePS(upn)}' | Where-Object { $_.User -ne 'NT AUTHORITY\\SELF' -and $_.IsInherited -eq $false } | Select-Object User,AccessRights`,
            );
            const fa = faRaw ? (Array.isArray(faRaw) ? faRaw : [faRaw]) : [];
            for (const perm of fa) {
              const rights = Array.isArray(perm.AccessRights)
                ? perm.AccessRights
                : [String(perm.AccessRights)];
              if (!rights.some((r) => r.includes("FullAccess"))) continue;
              allDelegations.push({
                value: `FullAccess:${perm.User}`,
                label: perm.User,
                hint: "Full Access",
              });
            }
          } catch { /* no mailbox */ }

          // Send As
          try {
            const saRaw = await ps.runCommandJson<RecipientPermission | RecipientPermission[]>(
              `Get-RecipientPermission -Identity '${escapePS(upn)}' | Where-Object { $_.Trustee -ne 'NT AUTHORITY\\SELF' } | Select-Object Trustee,AccessRights`,
            );
            const sa = saRaw ? (Array.isArray(saRaw) ? saRaw : [saRaw]) : [];
            for (const perm of sa) {
              allDelegations.push({
                value: `SendAs:${perm.Trustee}`,
                label: perm.Trustee,
                hint: "Send As",
              });
            }
          } catch { /* no mailbox */ }

          // Send on Behalf
          try {
            const { output, error: sobError } = await ps.runCommand(
              `$mbx = Get-Mailbox -Identity '${escapePS(upn)}' -ErrorAction SilentlyContinue; if ($mbx -and $mbx.GrantSendOnBehalfTo.Count -gt 0) { ($mbx.GrantSendOnBehalfTo | ForEach-Object { (Get-Recipient $_ -ErrorAction SilentlyContinue).PrimarySmtpAddress }) -join ',' } else { '' }`,
            );
            if (!sobError && output.trim()) {
              for (const u of output.trim().split(",").filter(Boolean)) {
                allDelegations.push({
                  value: `SendOnBehalf:${u}`,
                  label: u,
                  hint: "Send on Behalf",
                });
              }
            }
          } catch { /* no mailbox */ }

          fetchSpin.stop(`Found ${allDelegations.length} delegation(s).`);

          if (allDelegations.length === 0) {
            p.log.info("No delegations to remove.");
            break;
          }

          const toRemove = await p.multiselect({
            message: "Select delegations to remove",
            options: allDelegations,
            required: false,
          });
          if (p.isCancel(toRemove) || toRemove.length === 0) break;

          const removeSpin = p.spinner();
          removeSpin.start("Removing delegations...");

          const runDelegRemoved: string[] = [];
          for (const entry of toRemove) {
            const [type, ...userParts] = entry.split(":");
            const delegateUser = userParts.join(":");
            removeSpin.message(`Removing ${type} for ${delegateUser}...`);

            let error: string;
            if (type === "FullAccess") {
              ({ error } = await ps.runCommand(
                `Remove-MailboxPermission -Identity '${escapePS(upn)}' -User '${escapePS(delegateUser)}' -AccessRights FullAccess -InheritanceType All -Confirm:$false`,
              ));
            } else if (type === "SendAs") {
              ({ error } = await ps.runCommand(
                `Remove-RecipientPermission -Identity '${escapePS(upn)}' -Trustee '${escapePS(delegateUser)}' -AccessRights SendAs -Confirm:$false`,
              ));
            } else {
              ({ error } = await ps.runCommand(
                `Set-Mailbox -Identity '${escapePS(upn)}' -GrantSendOnBehalfTo @{Remove='${escapePS(delegateUser)}'}`,
              ));
            }
            if (error) {
              p.log.error(`Failed to remove ${type} for ${delegateUser}: ${error}`);
            } else {
              runDelegRemoved.push(`${type} ← ${delegateUser}`);
              delegationsRemoved.push(`${type} ← ${delegateUser}`);
            }
          }

          removeSpin.stop(
            runDelegRemoved.length > 0
              ? `Removed ${runDelegRemoved.length} delegation(s).`
              : "No delegations removed.",
          );
        }
        break;
      }

      case "email-forwarding": {
        const fwdSpin = p.spinner();
        fwdSpin.start("Checking current forwarding...");

        const fwdRaw = await ps.runCommandJson<ForwardingConfig>(
          `Get-Mailbox -Identity '${escapePS(upn)}' | Select-Object ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward`,
        );

        if (!fwdRaw) {
          fwdSpin.stop("Could not read mailbox forwarding.");
          break;
        }

        const hasForwarding =
          (fwdRaw.ForwardingSmtpAddress && fwdRaw.ForwardingSmtpAddress !== "") ||
          (fwdRaw.ForwardingAddress && fwdRaw.ForwardingAddress !== "");

        if (hasForwarding) {
          fwdSpin.stop("Forwarding is active.");
          p.note(
            [
              `SMTP Forward:    ${fwdRaw.ForwardingSmtpAddress ?? "(none)"}`,
              `Forward Address: ${fwdRaw.ForwardingAddress ?? "(none)"}`,
              `Keep copy:       ${fwdRaw.DeliverToMailboxAndForward}`,
            ].join("\n"),
            "Current Forwarding",
          );

          const removeConfirm = await p.confirm({
            message: "Remove forwarding?",
          });
          if (p.isCancel(removeConfirm) || !removeConfirm) break;

          const removeSpin = p.spinner();
          removeSpin.start("Removing forwarding...");
          const { error } = await ps.runCommand(
            `Set-Mailbox -Identity '${escapePS(upn)}' -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false`,
          );
          if (error) {
            removeSpin.stop("Failed to remove forwarding.");
            p.log.error(error);
          } else {
            removeSpin.stop("Forwarding removed.");
            forwardingRemoved = true;
          }
        } else {
          fwdSpin.stop("No forwarding configured.");

          const setConfirm = await p.confirm({
            message: "Set up email forwarding?",
          });
          if (p.isCancel(setConfirm) || !setConfirm) break;

          const targetEmail = await p.text({
            message: "Forward to email address",
            validate: (v = "") => {
              if (!v.trim()) return "Email is required";
              if (!v.includes("@")) return "Enter a valid email address";
            },
          });
          if (p.isCancel(targetEmail)) break;

          const keepCopy = await p.confirm({
            message: "Keep a copy in the original mailbox?",
            initialValue: true,
          });
          if (p.isCancel(keepCopy)) break;

          const setSpin = p.spinner();
          setSpin.start("Setting forwarding...");
          const { error } = await ps.runCommand(
            `Set-Mailbox -Identity '${escapePS(upn)}' -ForwardingSmtpAddress 'smtp:${escapePS(targetEmail)}' -DeliverToMailboxAndForward $${keepCopy}`,
          );
          if (error) {
            setSpin.stop("Failed to set forwarding.");
            p.log.error(error);
          } else {
            setSpin.stop(`Forwarding set to ${targetEmail}${keepCopy ? " (keeping copy)" : ""}.`);
            forwardingSet = targetEmail;
          }
        }
        break;
      }

      case "add-distribution-group": {
        const items = await addToDistributionGroup(ps, upn);
        addedDistGroups.push(...items);
        break;
      }

      case "add-security-group": {
        const item = await addToSecurityGroup(ps, upn);
        if (item) addedSecGroups.push(item);
        break;
      }

      case "add-shared-mailbox": {
        const items = await addToSharedMailbox(ps, upn);
        addedMailboxes.push(...items);
        break;
      }

      case "change-domain": {
        const currentUsername = upn.split("@")[0]!;
        const currentDomain = upn.split("@")[1]!;

        if (!acceptedDomains) {
          const domainSpin = p.spinner();
          domainSpin.start("Fetching accepted domains...");
          try {
            const raw = await ps.runCommandJson<AcceptedDomain | AcceptedDomain[]>(
              "Get-AcceptedDomain | Select-Object DomainName, Default",
            );
            acceptedDomains = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
            domainSpin.stop(`Found ${acceptedDomains.length} domain(s).`);
          } catch (e) {
            domainSpin.stop("Failed to fetch domains.");
            p.log.error(`${e}`);
            break;
          }
        }

        const otherDomains = acceptedDomains.filter(
          (d) => d.DomainName.toLowerCase() !== currentDomain.toLowerCase(),
        );
        if (otherDomains.length === 0) {
          p.log.warn("No other domains available in this tenant.");
          break;
        }

        const newDomain = await p.select({
          message: "Select new domain",
          options: otherDomains.map((d) => ({
            value: d.DomainName,
            label: d.DomainName,
            hint: d.Default ? "default" : undefined,
          })),
        });
        if (p.isCancel(newDomain)) break;

        const newUpn = `${currentUsername}@${newDomain}`;
        const confirm = await p.confirm({
          message: `Change ${upn} → ${newUpn}? This will affect sign-in.`,
        });
        if (p.isCancel(confirm) || !confirm) break;

        const oldUpn = upn;
        const spin = p.spinner();
        spin.start("Updating UPN...");
        const { error } = await ps.runCommand(
          `Update-MgUser -UserId '${escapePS(userId)}' -UserPrincipalName '${escapePS(newUpn)}'`,
        );
        if (error) {
          spin.stop("Failed to update UPN.");
          p.log.error(error);
          break;
        }
        spin.stop(`UPN changed to ${newUpn}.`);

        // Add old address as alias (use old UPN as identity since Exchange may not recognize new UPN yet)
        // Graph dual-write already updates the primary SMTP when UPN changes
        const aliasSpin = p.spinner();
        aliasSpin.start("Adding old address as alias...");
        const { error: aliasError } = await ps.runCommand(
          `Set-Mailbox -Identity '${escapePS(oldUpn)}' -EmailAddresses @{Add='smtp:${escapePS(oldUpn)}'}`,
        );
        if (aliasError) {
          aliasSpin.stop("Could not add old address as alias.");
          p.log.warn(`Old address may not be preserved as alias. You can add it manually.\n${aliasError}`);
        } else {
          aliasSpin.stop(`Old address ${oldUpn} preserved as alias.`);
        }

        upn = newUpn;
        current.details.UserPrincipalName = newUpn;
        if (!domainChangedFrom) domainChangedFrom = oldUpn;
        break;
      }

      case "add-alias": {
        const aliasUsername = await p.text({
          message: "Alias username (before @)",
          validate: (v = "") => {
            if (!v.trim()) return "Username is required";
            if (/[^a-zA-Z0-9._-]/.test(v)) return "Invalid characters in username";
          },
        });
        if (p.isCancel(aliasUsername)) break;

        if (!acceptedDomains) {
          const domainSpin = p.spinner();
          domainSpin.start("Fetching accepted domains...");
          try {
            const raw = await ps.runCommandJson<AcceptedDomain | AcceptedDomain[]>(
              "Get-AcceptedDomain | Select-Object DomainName, Default",
            );
            acceptedDomains = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
            domainSpin.stop(`Found ${acceptedDomains.length} domain(s).`);
          } catch (e) {
            domainSpin.stop("Failed to fetch domains.");
            p.log.error(`${e}`);
            break;
          }
        }

        const aliasDomain = await p.select({
          message: "Domain for alias",
          options: acceptedDomains.map((d) => ({
            value: d.DomainName,
            label: d.DomainName,
            hint: d.Default ? "default" : undefined,
          })),
        });
        if (p.isCancel(aliasDomain)) break;

        const aliasAddress = `${aliasUsername}@${aliasDomain}`;
        const spin = p.spinner();
        spin.start(`Adding alias ${aliasAddress}...`);
        const { error } = await ps.runCommand(
          `Set-Mailbox -Identity '${escapePS(upn)}' -EmailAddresses @{Add='smtp:${escapePS(aliasAddress)}'}`,
        );
        if (error) {
          spin.stop("Failed to add alias.");
          p.log.error(error);
        } else {
          spin.stop(`Alias ${aliasAddress} added.`);
          aliasesAdded.push(aliasAddress);
        }
        break;
      }

      case "copy-ots": {
        try {
          if (process.platform === "win32") {
            await ps.runCommand(`Set-Clipboard -Value '${escapePS(otsUrl!)}'`);
          } else {
            const proc = Bun.spawn(["pbcopy"], { stdin: new Blob([otsUrl!]) });
            await proc.exited;
          }
          p.log.success(`Copied to clipboard: ${otsUrl}`);
        } catch {
          p.log.info(`Secret link: ${otsUrl}`);
        }
        break;
      }

      case "copy-ticket": {
        const formatItem = (g: { name: string; email: string }) =>
          g.email ? `${g.name} - ${g.email}` : g.name;

        const lines: string[] = [];

        if (passwordWasReset) {
          lines.push(`Reset password for ${current.details.DisplayName} (${upn}).`);
        }
        if (nameChangedTo) {
          lines.push(`Changed display name to ${nameChangedTo}.`);
        }
        if (domainChangedFrom) {
          lines.push(`Changed primary domain from ${domainChangedFrom} to ${upn}.`);
        }
        if (aliasesAdded.length > 0) {
          lines.push(`Added email alias(es): ${aliasesAdded.join(", ")}.`);
        }
        if (licensesAdded.length > 0) {
          lines.push(`Added license(s): ${licensesAdded.join(", ")}.`);
        }
        if (licensesRemoved.length > 0) {
          lines.push(`Removed license(s): ${licensesRemoved.join(", ")}.`);
        }
        if (rolesAdded.length > 0) {
          lines.push(`Added admin role(s): ${rolesAdded.join(", ")}.`);
        }
        if (rolesRemoved.length > 0) {
          lines.push(`Removed admin role(s): ${rolesRemoved.join(", ")}.`);
        }
        if (signInToggled) {
          lines.push(`${signInToggled === "blocked" ? "Blocked" : "Unblocked"} sign-in for ${current.details.DisplayName}.`);
        }
        if (delegationsAdded.length > 0) {
          lines.push(`Added mailbox delegation(s): ${delegationsAdded.join(", ")}.`);
        }
        if (delegationsRemoved.length > 0) {
          lines.push(`Removed mailbox delegation(s): ${delegationsRemoved.join(", ")}.`);
        }
        if (forwardingSet) {
          lines.push(`Set email forwarding to ${forwardingSet}.`);
        }
        if (forwardingRemoved) {
          lines.push("Removed email forwarding.");
        }
        if (mfaMethodsRemoved.length > 0) {
          lines.push(`Removed MFA method(s): ${mfaMethodsRemoved.join(", ")}.`);
        }

        const allGroups: string[] = [];
        for (const g of addedDistGroups) allGroups.push(formatItem(g));
        for (const g of addedSecGroups) allGroups.push(formatItem(g));
        for (const g of addedMailboxes) allGroups.push(formatItem(g));
        if (allGroups.length > 0) {
          lines.push(`Added to group(s):\n${allGroups.map((g) => `  - ${g}`).join("\n")}`);
        }

        if (otsUrl) {
          lines.push(`\nYou can retrieve new credentials via this link: ${otsUrl}\nMake sure to save it since it's a one-time link and it will expire in 7 days.`);
        }

        if (lines.length === 0) {
          p.log.warn("No changes to copy.");
          break;
        }

        const ticketNote = lines.join("\n");

        try {
          if (process.platform === "win32") {
            await ps.runCommand(`Set-Clipboard -Value '${escapePS(ticketNote)}'`);
          } else {
            const proc = Bun.spawn(["pbcopy"], { stdin: new Blob([ticketNote]) });
            await proc.exited;
          }
          p.log.success("Ticket update note copied to clipboard.");
        } catch {
          p.log.info(ticketNote);
        }
        break;
      }
    }
  }

  // Re-fetch and display final details
  const final = await fetchUserDetails(ps, userId);
  if (final) {
    displayUserDetails(final.details, final.licenses);
  }

  p.log.success("Done editing user.");
}
