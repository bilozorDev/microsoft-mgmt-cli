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

  p.note(domains.map((d) => `  ${d}`).join("\n"), "Domains to whitelist");

  const confirm = await p.confirm({
    message: `Add ${domains.length} domain(s) to the safe sender list?`,
  });

  if (p.isCancel(confirm) || !confirm) {
    p.log.info("Cancelled.");
    return;
  }

  const spin = p.spinner();

  // Add each domain
  spin.start("Adding domain(s) to safe sender list...");

  const results: { domain: string; success: boolean; error?: string }[] = [];

  for (const domain of domains) {
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
