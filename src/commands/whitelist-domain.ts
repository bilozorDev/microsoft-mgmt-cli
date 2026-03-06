import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { escapePS, handleOrgCustomizationError } from "../utils.ts";

function isValidDomain(domain: string): boolean {
  return /^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)*\.[a-zA-Z]{2,}$/.test(domain);
}

export async function run(ps: PowerShellSession): Promise<void> {
  const input = await p.text({
    message: "Enter domain(s) to whitelist (comma-separated)",
    placeholder: "example.com, contoso.com",
    validate(value = "") {
      if (!value.trim()) return "Please enter at least one domain";

      const domains = value.split(",").map((d) => d.trim()).filter(Boolean);
      const invalid = domains.filter((d) => !isValidDomain(d));
      if (invalid.length > 0) {
        return `Invalid domain(s): ${invalid.join(", ")}`;
      }
    },
  });

  if (p.isCancel(input)) return;

  const domains = input.split(",").map((d) => d.trim()).filter(Boolean);

  // Fetch current allowed sender domains to skip duplicates
  const spin = p.spinner();
  spin.start("Checking existing safe sender list...");

  let existingDomains = new Set<string>();
  const { output: existingList, error: fetchError } = await ps.runCommand(
    '(Get-HostedContentFilterPolicy -Identity "Default").AllowedSenderDomains | ForEach-Object { $_.Domain }'
  );

  if (fetchError) {
    spin.stop("Could not fetch existing list — will attempt to add all domains.");
    p.log.warn(fetchError);
  } else {
    spin.stop("Checked existing safe sender list.");
    existingDomains = new Set(
      (existingList || "").split("\n").map((l) => l.trim().toLowerCase()).filter(Boolean)
    );
  }

  const alreadyPresent = domains.filter((d) => existingDomains.has(d.toLowerCase()));
  const toAdd = domains.filter((d) => !existingDomains.has(d.toLowerCase()));

  if (toAdd.length === 0) {
    for (const d of alreadyPresent) {
      p.log.info(`${d} is already in safe sender list`);
    }
    return;
  }

  for (const d of alreadyPresent) {
    p.log.info(`${d} is already in safe sender list`);
  }

  p.note(toAdd.map((d) => `  ${d}`).join("\n"), "Domains to whitelist");

  const confirm = await p.confirm({
    message: `Add ${toAdd.length} domain(s) to the safe sender list?`,
  });

  if (p.isCancel(confirm) || !confirm) {
    p.log.info("Cancelled.");
    return;
  }

  // Add each domain
  spin.start("Adding domain(s) to safe sender list...");

  const results: { domain: string; success: boolean; error?: string }[] = [];
  let orgCustomizationHandled = false;

  for (const domain of toAdd) {
    const cmd = `Set-HostedContentFilterPolicy -Identity 'Default' -AllowedSenderDomains @{Add='${escapePS(domain)}'}`;
    const { error } = await ps.runCommand(cmd);

    if (
      error &&
      !orgCustomizationHandled &&
      error.includes("Enable-OrganizationCustomization")
    ) {
      orgCustomizationHandled = true;
      spin.stop("Organization customization required.");

      const fixed = await handleOrgCustomizationError(ps, error);

      if (fixed) {
        // Retry this domain (with up to 3 attempts for transient errors)
        spin.start("Adding domain(s) to safe sender list...");
        let retry = await ps.runCommand(cmd);
        for (let attempt = 1; retry.error && attempt < 3; attempt++) {
          await Bun.sleep(10_000);
          retry = await ps.runCommand(cmd);
        }
        results.push({
          domain,
          success: !retry.error,
          error: retry.error || undefined,
        });
      } else {
        // Mark this and all remaining domains as failed
        const msg = "Organization customization required";
        results.push({ domain, success: false, error: msg });
        for (const remaining of toAdd.slice(toAdd.indexOf(domain) + 1)) {
          results.push({ domain: remaining, success: false, error: msg });
        }
        break;
      }
    } else {
      results.push({ domain, success: !error, error: error || undefined });
    }
  }

  spin.stop("Done adding domains.");

  // Verify by fetching current list
  spin.start("Verifying safe sender list...");

  const { output: currentList, error: listError } = await ps.runCommand(
    '(Get-HostedContentFilterPolicy -Identity "Default").AllowedSenderDomains | ForEach-Object { $_.Domain }'
  );

  spin.stop("Verification complete.");

  // Show results
  for (const r of results) {
    if (r.success) {
      p.log.success(`${r.domain} — added`);
    } else {
      p.log.error(`${r.domain} — failed: ${r.error}`);
    }
  }

  if (!listError && currentList) {
    const domainList = currentList.split("\n").map((l) => l.trim()).filter(Boolean);
    p.note(domainList.map((d) => `  ${d}`).join("\n"), "Current allowed sender domains");
  }
}
