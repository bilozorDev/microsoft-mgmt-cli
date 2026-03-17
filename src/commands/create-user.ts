import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generatePassword, validatePassword } from "../password.ts";
import { friendlySkuName } from "../sku-names.ts";
import { escapePS, createSecretLink } from "../utils.ts";
import { run as addToDistributionGroup } from "./add-to-distribution-group.ts";
import { run as addToSecurityGroup } from "./add-to-security-group.ts";
import { run as addToSharedMailbox } from "./add-to-shared-mailbox.ts";

interface AcceptedDomain {
  DomainName: string;
  Default: boolean;
}

interface SubscribedSku {
  SkuId: string;
  SkuPartNumber: string;
  ConsumedUnits: number;
  PrepaidUnits: { Enabled: number; [key: string]: unknown };
}

function suggestUsername(displayName: string): string {
  const parts = displayName.trim().split(/\s+/);
  if (parts.length < 2) return parts[0]!.toLowerCase();
  return `${parts[0]!}.${parts[parts.length - 1]!}`.toLowerCase();
}

export async function run(ps: PowerShellSession): Promise<void> {
  // 1. Ensure Graph + Exchange connected
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph (check your browser)...");
  try {
    await ps.ensureGraphConnected(["User.ReadWrite.All", "Organization.Read.All", "GroupMember.ReadWrite.All"]);
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return;
  }

  const exoSpin = p.spinner();
  exoSpin.start("Connecting to Exchange Online (check your browser)...");
  try {
    await ps.ensureExchangeConnected();
    exoSpin.stop("Connected to Exchange Online.");
  } catch (e) {
    exoSpin.stop("Failed to connect to Exchange Online.");
    p.log.error(`${e}`);
    return;
  }

  // 2. First name, last name, display name
  const firstName = await p.text({
    message: "First name",
    placeholder: "Jane",
    validate: (v = "") => !v.trim() ? "First name is required" : undefined,
  });
  if (p.isCancel(firstName)) return;

  const lastName = await p.text({
    message: "Last name",
    placeholder: "Doe",
    validate: (v = "") => !v.trim() ? "Last name is required" : undefined,
  });
  if (p.isCancel(lastName)) return;

  const displayName = await p.text({
    message: "Display name",
    initialValue: `${firstName} ${lastName}`,
    validate: (v = "") => !v.trim() ? "Display name is required" : undefined,
  });
  if (p.isCancel(displayName)) return;

  // 3–4. Username + domain selection (with UPN availability check)
  let upn: string;
  let lastUsername = suggestUsername(`${firstName} ${lastName}`);

  let domains: AcceptedDomain[] | null = null;
  let defaultDomain: string | undefined;

  upnLoop: while (true) {
    const username = await p.text({
      message: "Username (before @)",
      initialValue: lastUsername,
      validate: (v = "") => {
        if (!v.trim()) return "Username is required";
        if (/[^a-zA-Z0-9._-]/.test(v)) return "Invalid characters in username";
      },
    });
    if (p.isCancel(username)) return;
    lastUsername = username;

    if (!domains) {
      const domainSpin = p.spinner();
      domainSpin.start("Fetching accepted domains...");
      try {
        const raw = await ps.runCommandJson<AcceptedDomain | AcceptedDomain[]>(
          "Get-AcceptedDomain | Select-Object DomainName, Default",
        );
        domains = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        domainSpin.stop(`Found ${domains.length} domain(s).`);
      } catch (e) {
        domainSpin.stop("Failed to fetch domains.");
        p.log.error(`${e}`);
        return;
      }
      defaultDomain = domains.find((d) => d.Default)?.DomainName ?? domains[0]?.DomainName;
    }

    const domain = await p.select({
      message: "Domain",
      options: domains.map((d) => ({
        value: d.DomainName,
        label: d.DomainName,
        hint: d.Default ? "default" : undefined,
      })),
      initialValue: defaultDomain,
    });
    if (p.isCancel(domain)) return;

    upn = `${username}@${domain}`;

    const checkSpin = p.spinner();
    checkSpin.start(`Checking if ${upn} is available...`);
    const { output } = await ps.runCommand(
      `Get-MgUser -Filter "userPrincipalName eq '${escapePS(upn)}'" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id`,
    );
    if (output.trim()) {
      checkSpin.stop(`${upn} is already taken.`);
      p.log.warn("Please choose a different username.");
      continue upnLoop;
    }
    checkSpin.stop(`${upn} is available.`);
    break;
  }

  // 5. License selection
  let selectedSkus: SubscribedSku[] = [];

  licenseLoop: while (true) {
    const licSpin = p.spinner();
    licSpin.start("Fetching available licenses...");

    let skus: SubscribedSku[];
    try {
      const raw = await ps.runCommandJson<SubscribedSku | SubscribedSku[]>(
        "Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, PrepaidUnits",
      );
      skus = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
      licSpin.stop(`Found ${skus.length} license type(s).`);
    } catch (e) {
      licSpin.stop("Failed to fetch licenses.");
      p.log.error(`${e}`);
      return;
    }

    if (skus.length > 0) {
      const choices = await p.multiselect({
        message: "Assign licenses (space to toggle, enter to confirm)",
        options: skus.map((s) => {
          const available = s.PrepaidUnits.Enabled - s.ConsumedUnits;
          return {
            value: s.SkuId,
            label: friendlySkuName(s.SkuPartNumber),
            hint: `${available} of ${s.PrepaidUnits.Enabled} available`,
            disabled: available <= 0,
          };
        }),
        required: false,
      });
      if (p.isCancel(choices)) return;

      if (choices.length === 0) {
        const confirm = await p.select({
          message: "No licenses selected. Create user without a license?",
          options: [
            { value: "back", label: "Go back to licenses" },
            { value: "none", label: "Yes, continue without license" },
          ],
        });
        if (p.isCancel(confirm)) return;
        if (confirm === "back") continue licenseLoop;
      }

      selectedSkus = skus.filter((s) => choices.includes(s.SkuId));
      break;
    } else {
      // No seats available — show counts and offer options
      p.log.warn("No licenses with available seats.");
      for (const s of skus) {
        p.log.info(
          `${friendlySkuName(s.SkuPartNumber)}: ${s.ConsumedUnits}/${s.PrepaidUnits.Enabled} used`,
        );
      }

      const noSeatsAction = await p.select({
        message: "What would you like to do?",
        options: [
          { value: "refresh", label: "Refresh", hint: "re-query licenses" },
          { value: "none", label: "Create without license" },
          { value: "cancel", label: "Cancel" },
        ],
      });
      if (p.isCancel(noSeatsAction) || noSeatsAction === "cancel") return;
      if (noSeatsAction === "none") break;
      // "refresh" loops back
      continue licenseLoop;
    }
  }

  // 7. Password
  let password: string;

  const pwMethod = await p.select({
    message: "Password",
    options: [
      { value: "auto", label: "Auto-generate", hint: "16 chars, crypto-random" },
      { value: "manual", label: "Enter manually" },
    ],
  });
  if (p.isCancel(pwMethod)) return;

  if (pwMethod === "auto") {
    password = generatePassword();
  } else {
    const manualPw = await p.password({
      message: "Enter password",
      validate: (v) => validatePassword(v ?? ""),
    });
    if (p.isCancel(manualPw)) return;
    password = manualPw;
  }

  // 8. Confirmation
  const licenseLine = selectedSkus.length > 0
    ? selectedSkus.map((s) => friendlySkuName(s.SkuPartNumber)).join(", ")
    : "(none)";

  p.note(
    [
      `Name:     ${displayName}`,
      `UPN:      ${upn}`,
      `Licenses: ${licenseLine}`,
    ].join("\n"),
    "New user summary",
  );

  const ok = await p.confirm({ message: "Create this user?" });
  if (p.isCancel(ok) || !ok) {
    p.log.info("Cancelled.");
    return;
  }

  // 9. Create user
  const createSpin = p.spinner();
  createSpin.start("Creating user...");

  const createCmd = [
    "New-MgUser -BodyParameter @{",
    `  DisplayName = '${escapePS(displayName)}'`,
    `  GivenName = '${escapePS(firstName)}'`,
    `  Surname = '${escapePS(lastName)}'`,
    `  UserPrincipalName = '${escapePS(upn)}'`,
    `  MailNickname = '${escapePS(upn.split("@")[0]!)}'`,
    "  AccountEnabled = $true",
    "  UsageLocation = 'US'",
    "  PasswordProfile = @{",
    `    Password = '${escapePS(password)}'`,
    "    ForceChangePasswordNextSignIn = $true",
    "  }",
    "}",
  ].join("\n");

  const { error: createError } = await ps.runCommand(createCmd);
  if (createError) {
    createSpin.stop("Failed to create user.");
    p.log.error(createError);
    return;
  }
  createSpin.stop("User created.");

  // 10. Assign licenses
  if (selectedSkus.length > 0) {
    const licSpin = p.spinner();
    licSpin.start(`Assigning ${selectedSkus.length} license(s)...`);

    const skuEntries = selectedSkus
      .map((s) => `@{SkuId = '${s.SkuId}'}`)
      .join(", ");
    const licCmd = `Set-MgUserLicense -UserId '${escapePS(upn)}' -AddLicenses @(${skuEntries}) -RemoveLicenses @()`;

    const { error: licError } = await ps.runCommand(licCmd);
    if (licError) {
      licSpin.stop("Failed to assign licenses.");
      p.log.error(licError);
      p.log.warn("User was created but licenses were not assigned.");
    } else {
      licSpin.stop(`${selectedSkus.length} license(s) assigned.`);
    }
  }

  // 11. Show credentials & create one-time secret link
  p.note(
    [`UPN:      ${upn}`, `Password: ${password}`].join("\n"),
    "Credentials (user must change password at first sign-in)",
  );

  const otsSpin = p.spinner();
  otsSpin.start("Creating one-time secret link...");
  const otsResult = await createSecretLink(
    `${password}`,
  );
  const otsUrl = "url" in otsResult ? otsResult.url : null;
  if (otsUrl) {
    otsSpin.stop("One-time secret link created.");
    p.log.info(`Secret link: ${otsUrl}`);
  } else {
    otsSpin.stop("Failed to create one-time secret link.");
    p.log.error("error" in otsResult ? otsResult.error : "Unknown error");
  }

  p.log.success("User created successfully.");

  // 12. Wait for Exchange Online to recognize the new user
  {
    const waitSpin = p.spinner();
    waitSpin.start("Waiting for Exchange Online to provision user...");
    let exchangeReady = false;
    for (let i = 0; i < 30; i++) {
      const { output } = await ps.runCommand(
        `Get-Recipient -Identity '${escapePS(upn)}' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PrimarySmtpAddress`,
      );
      if (output.trim()) {
        exchangeReady = true;
        break;
      }
      await new Promise((r) => setTimeout(r, 10_000));
    }
    if (exchangeReady) {
      waitSpin.stop("User is visible in Exchange Online.");
    } else {
      waitSpin.stop("User not yet visible in Exchange Online.");
      p.log.warn(
        "Exchange provisioning is still in progress. Shared mailbox operations may fail — try again from the main menu later.",
      );
    }
  }

  // 13. Post-creation membership setup
  const addedDistGroups: { name: string; email: string }[] = [];
  const addedSecGroups: { name: string; email: string }[] = [];
  const addedMailboxes: { name: string; email: string }[] = [];

  while (true) {
    const menuOptions: { value: string; label: string }[] = [
      { value: "distribution-group", label: "Distribution group" },
      { value: "security-group", label: "Security group" },
      { value: "shared-mailbox", label: "Shared mailbox" },
    ];
    if (otsUrl) {
      menuOptions.push({ value: "copy-ots", label: "Copy OTS link to clipboard" });
      menuOptions.push({ value: "copy-ticket", label: "Copy ticket update note to clipboard" });
    }
    menuOptions.push({ value: "done", label: "Done" });

    const action = await p.select({
      message: "Add user to...",
      options: menuOptions,
    });
    if (p.isCancel(action) || action === "done") break;

    switch (action) {
      case "distribution-group": {
        const items = await addToDistributionGroup(ps, upn);
        addedDistGroups.push(...items);
        break;
      }
      case "security-group": {
        const item = await addToSecurityGroup(ps, upn);
        if (item) addedSecGroups.push(item);
        break;
      }
      case "shared-mailbox": {
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

        let ticketNote = `Created mailbox for ${displayName} (${upn}).`;

        const allGroups: string[] = [];
        for (const g of addedDistGroups) allGroups.push(formatItem(g));
        for (const g of addedSecGroups) allGroups.push(formatItem(g));
        for (const g of addedMailboxes) allGroups.push(formatItem(g));
        if (allGroups.length > 0) {
          ticketNote += `\nAdded to group(s):\n${allGroups.map((g) => `  - ${g}`).join("\n")}`;
        }

        ticketNote += `\n\nYou can retrieve credentials via this link: ${otsUrl}\nMake sure to save it since it's a one-time link and it will expire in 7 days.`;

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

  // 14. Summary
  const parts = [`Created user ${upn}`];
  if (addedDistGroups.length > 0) {
    parts.push(`Distribution group(s): ${addedDistGroups.map((g) => g.name).join(", ")}`);
  }
  if (addedSecGroups.length > 0) {
    parts.push(`Security group(s): ${addedSecGroups.map((g) => g.name).join(", ")}`);
  }
  if (addedMailboxes.length > 0) {
    parts.push(`Shared mailbox(es): ${addedMailboxes.map((g) => g.name).join(", ")}`);
  }

  p.note(parts.join("\n"), "Summary");
}
