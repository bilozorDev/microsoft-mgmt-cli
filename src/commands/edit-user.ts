import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generatePassword, validatePassword } from "../password.ts";
import { friendlySkuName } from "../sku-names.ts";
import { escapePS, createSecretLink } from "../utils.ts";
import { run as addToDistributionGroup } from "./add-to-distribution-group.ts";
import { run as addToSecurityGroup } from "./add-to-security-group.ts";
import { run as addToSharedMailbox } from "./add-to-shared-mailbox.ts";

interface MgUser {
  DisplayName: string;
  UserPrincipalName: string;
  Id: string;
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

interface SubscribedSku {
  SkuId: string;
  SkuPartNumber: string;
  ConsumedUnits: number;
  PrepaidUnits: { Enabled: number; [key: string]: unknown };
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
      "Get-MgUser -All -Property DisplayName,UserPrincipalName,Id | Select-Object DisplayName,UserPrincipalName,Id",
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
          `Get-MgUser -Search '"displayName:${escapePS(query)}"' -ConsistencyLevel eventual -Property DisplayName,UserPrincipalName,Id | Select-Object DisplayName,UserPrincipalName,Id`,
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
      hint: u.UserPrincipalName,
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
    await ps.ensureGraphConnected(true);
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
  const upn = user.UserPrincipalName;

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

  while (true) {
    const menuOptions: { value: string; label: string }[] = [
      { value: "edit-name", label: "Edit name" },
      { value: "reset-password", label: "Reset password" },
      { value: "manage-licenses", label: "Manage licenses" },
      { value: "add-distribution-group", label: "Add to distribution group" },
      { value: "add-security-group", label: "Add to security group" },
      { value: "add-shared-mailbox", label: "Add to shared mailbox" },
    ];
    if (otsUrl) {
      menuOptions.push({ value: "copy-ots", label: "Copy OTS link to clipboard" });
      menuOptions.push({ value: "copy-ticket", label: "Copy ticket update note to clipboard" });
    }
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
        const licAction = await p.select({
          message: "License action",
          options: [
            { value: "add", label: "Add license" },
            { value: "remove", label: "Remove license" },
          ],
        });
        if (p.isCancel(licAction)) break;

        if (licAction === "add") {
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

          // Exclude already-assigned licenses
          const currentSkuIds = new Set(current.licenses.map((l) => l.SkuId));
          const available = skus.filter(
            (s) => !currentSkuIds.has(s.SkuId) && s.PrepaidUnits.Enabled - s.ConsumedUnits > 0,
          );

          if (available.length === 0) {
            p.log.warn("No additional licenses with available seats.");
            break;
          }

          const choices = await p.multiselect({
            message: "Select license(s) to add (space to toggle, enter to confirm)",
            options: available.map((s) => {
              const avail = s.PrepaidUnits.Enabled - s.ConsumedUnits;
              return {
                value: s.SkuId,
                label: friendlySkuName(s.SkuPartNumber),
                hint: `${avail} of ${s.PrepaidUnits.Enabled} available`,
              };
            }),
            required: true,
          });
          if (p.isCancel(choices)) break;

          const spin = p.spinner();
          spin.start("Adding license(s)...");
          const skuEntries = choices.map((id) => `@{SkuId = '${escapePS(id)}'}`).join(", ");
          const { error } = await ps.runCommand(
            `Set-MgUserLicense -UserId '${escapePS(userId)}' -AddLicenses @(${skuEntries}) -RemoveLicenses @()`,
          );
          if (error) {
            spin.stop("Failed to add license(s).");
            p.log.error(error);
          } else {
            const addedNames = choices
              .map((id) => skus.find((s) => s.SkuId === id))
              .filter(Boolean)
              .map((s) => friendlySkuName(s!.SkuPartNumber));
            spin.stop(`Added: ${addedNames.join(", ")}`);
            licensesAdded.push(...addedNames);
            // Update local state
            for (const id of choices) {
              const sku = skus.find((s) => s.SkuId === id);
              if (sku) current.licenses.push({ SkuId: sku.SkuId, SkuPartNumber: sku.SkuPartNumber });
            }
          }
        } else {
          // Remove license
          if (current.licenses.length === 0) {
            p.log.warn("User has no licenses to remove.");
            break;
          }

          const choices = await p.multiselect({
            message: "Select license(s) to remove (space to toggle, enter to confirm)",
            options: current.licenses.map((l) => ({
              value: l.SkuId,
              label: friendlySkuName(l.SkuPartNumber),
            })),
            required: true,
          });
          if (p.isCancel(choices)) break;

          // Warn if removing ALL licenses
          if (choices.length === current.licenses.length) {
            const confirm = await p.confirm({
              message: "This will remove ALL licenses. User may lose access to data. Are you sure?",
            });
            if (p.isCancel(confirm) || !confirm) break;
          }

          const spin = p.spinner();
          spin.start("Removing license(s)...");
          const skuIds = choices.map((id) => `'${escapePS(id)}'`).join(",");
          const { error } = await ps.runCommand(
            `Set-MgUserLicense -UserId '${escapePS(userId)}' -AddLicenses @() -RemoveLicenses @(${skuIds})`,
          );
          if (error) {
            spin.stop("Failed to remove license(s).");
            p.log.error(error);
          } else {
            const removedNames = choices
              .map((id) => current!.licenses.find((l) => l.SkuId === id))
              .filter(Boolean)
              .map((l) => friendlySkuName(l!.SkuPartNumber));
            spin.stop(`Removed: ${removedNames.join(", ")}`);
            licensesRemoved.push(...removedNames);
            current.licenses = current.licenses.filter((l) => !choices.includes(l.SkuId));
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
        if (licensesAdded.length > 0) {
          lines.push(`Added license(s): ${licensesAdded.join(", ")}.`);
        }
        if (licensesRemoved.length > 0) {
          lines.push(`Removed license(s): ${licensesRemoved.join(", ")}.`);
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
