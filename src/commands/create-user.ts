import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generatePassword, validatePassword } from "../password.ts";
import { friendlySkuName } from "../sku-names.ts";

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

function escapePS(value: string): string {
  return value.replace(/'/g, "''");
}

export async function run(ps: PowerShellSession): Promise<void> {
  // 1. Ensure Graph connected
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

  // 2. Full name
  const displayName = await p.text({
    message: "Full name (display name)",
    placeholder: "Jane Doe",
    validate: (v = "") => !v.trim() ? "Name is required" : undefined,
  });
  if (p.isCancel(displayName)) return;

  // 3. Username
  const username = await p.text({
    message: "Username (before @)",
    initialValue: suggestUsername(displayName),
    validate: (v = "") => {
      if (!v.trim()) return "Username is required";
      if (/[^a-zA-Z0-9._-]/.test(v)) return "Invalid characters in username";
    },
  });
  if (p.isCancel(username)) return;

  // 4. Domain selection
  const domainSpin = p.spinner();
  domainSpin.start("Fetching accepted domains...");

  let domains: AcceptedDomain[];
  try {
    const raw = await ps.runCommandJson<AcceptedDomain | AcceptedDomain[]>(
      "Get-AcceptedDomain | Select-Object DomainName, Default",
    );
    domains = Array.isArray(raw) ? raw : [raw];
    domainSpin.stop(`Found ${domains.length} domain(s).`);
  } catch (e) {
    domainSpin.stop("Failed to fetch domains.");
    p.log.error(`${e}`);
    return;
  }

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
  if (p.isCancel(domain)) return;

  const upn = `${username}@${domain}`;

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
      skus = Array.isArray(raw) ? raw : [raw];
      licSpin.stop(`Found ${skus.length} license type(s).`);
    } catch (e) {
      licSpin.stop("Failed to fetch licenses.");
      p.log.error(`${e}`);
      return;
    }

    const withSeats = skus.filter(
      (s) => s.PrepaidUnits.Enabled - s.ConsumedUnits > 0,
    );

    if (withSeats.length > 0) {
      const choices = await p.multiselect({
        message: "Assign licenses (space to toggle, enter to confirm)",
        options: withSeats.map((s) => {
          const available = s.PrepaidUnits.Enabled - s.ConsumedUnits;
          return {
            value: s.SkuId,
            label: friendlySkuName(s.SkuPartNumber),
            hint: `${available} of ${s.PrepaidUnits.Enabled} available`,
          };
        }),
        required: false,
      });
      if (p.isCancel(choices)) return;

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
    `  UserPrincipalName = '${escapePS(upn)}'`,
    `  MailNickname = '${escapePS(username)}'`,
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

  // 11. Show credentials
  p.note(
    [`UPN:      ${upn}`, `Password: ${password}`].join("\n"),
    "Credentials (user must change password at first sign-in)",
  );

  p.log.success("User created successfully.");
}
