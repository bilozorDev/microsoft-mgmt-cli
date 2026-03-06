import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { escapePS } from "../utils.ts";

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

  for (const domain of toAdd) {
    const { error } = await ps.runCommand(
      `Set-HostedContentFilterPolicy -Identity 'Default' -AllowedSenderDomains @{Add='${escapePS(domain)}'}`
    );
    results.push({ domain, success: !error, error: error || undefined });
  }

  spin.stop("Done adding domains.");

  // Verify by fetching current list
  spin.start("Verifying safe sender list...");

  const { output: currentList, error: listError } = await ps.runCommand(
    'Get-HostedContentFilterPolicy -Identity "Default" | Select-Object -ExpandProperty AllowedSenderDomains'
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
    p.note(currentList, "Current allowed sender domains");
  }
}
