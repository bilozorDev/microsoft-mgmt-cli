import { resolve, dirname, join } from "path";
import { mkdirSync, chmodSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generateReport } from "../report-template.ts";
import { appDir, escapePS, handleOrgCustomizationError } from "../utils.ts";

interface AcceptedDomain {
  DomainName: string;
  DomainType: string;
  Default: boolean;
}

interface DkimConfig {
  Domain: string;
  Enabled: boolean;
  Status: string;
  Selector1CNAME: string | null;
  Selector2CNAME: string | null;
}

interface MergedDomain {
  domain: string;
  domainType: string;
  isDefault: boolean;
  dkimEnabled: boolean;
  dkimStatus: string;
  selector1CNAME: string;
  selector2CNAME: string;
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

async function fetchDomainData(ps: PowerShellSession): Promise<MergedDomain[]> {
  const spin = p.spinner();
  spin.start("Fetching accepted domains...");

  const domainsRaw = await ps.runCommandJson<AcceptedDomain | AcceptedDomain[]>(
    "Get-AcceptedDomain | Select-Object DomainName,DomainType,Default"
  );
  const domains = domainsRaw ? (Array.isArray(domainsRaw) ? domainsRaw : [domainsRaw]) : [];

  spin.message("Fetching DKIM signing configurations...");

  let dkimConfigs: DkimConfig[] = [];
  try {
    const dkimRaw = await ps.runCommandJson<DkimConfig | DkimConfig[]>(
      "Get-DkimSigningConfig | Select-Object Domain,Enabled,Status,Selector1CNAME,Selector2CNAME"
    );
    dkimConfigs = dkimRaw ? (Array.isArray(dkimRaw) ? dkimRaw : [dkimRaw]) : [];
  } catch {
    // Tenant may never have configured DKIM
  }

  spin.stop(`Found ${domains.length} domain(s), ${dkimConfigs.length} DKIM config(s).`);

  // Merge datasets
  const dkimMap = new Map<string, DkimConfig>();
  for (const c of dkimConfigs) {
    dkimMap.set(c.Domain.toLowerCase(), c);
  }

  return domains.map((d) => {
    const dkim = dkimMap.get(d.DomainName.toLowerCase());
    return {
      domain: d.DomainName,
      domainType: d.DomainType,
      isDefault: d.Default,
      dkimEnabled: dkim?.Enabled ?? false,
      dkimStatus: dkim?.Status ?? "Not configured",
      selector1CNAME: dkim?.Selector1CNAME ?? "",
      selector2CNAME: dkim?.Selector2CNAME ?? "",
    };
  });
}

async function viewDomains(ps: PowerShellSession): Promise<MergedDomain[]> {
  const merged = await fetchDomainData(ps);

  if (merged.length === 0) {
    p.log.info("No accepted domains found.");
    return [];
  }

  const header = `${"Domain".padEnd(30)} ${"Type".padEnd(14)} ${"DKIM".padEnd(8)} ${"Status".padEnd(18)} Default`;
  const separator = "─".repeat(header.length + 5);
  const lines = [header, separator];

  for (const d of merged.slice(0, 50)) {
    const domain = truncate(d.domain, 29).padEnd(30);
    const type = d.domainType.padEnd(14);
    const dkim = (d.dkimEnabled ? "Yes" : "No").padEnd(8);
    const status = truncate(d.dkimStatus, 17).padEnd(18);
    const def = d.isDefault ? "Yes" : "";
    lines.push(`${domain} ${type} ${dkim} ${status} ${def}`);
  }
  if (merged.length > 50) {
    lines.push(`… and ${merged.length - 50} more`);
  }

  p.note(lines.join("\n"), `DKIM Audit (${merged.length} domain(s))`);

  // Show CNAME records for unconfigured domains
  const unconfigured = merged.filter((d) => !d.dkimEnabled && d.selector1CNAME);
  if (unconfigured.length > 0) {
    const cnameLines: string[] = [];
    for (const d of unconfigured) {
      cnameLines.push(`${d.domain}:`);
      if (d.selector1CNAME) cnameLines.push(`  selector1._domainkey → ${d.selector1CNAME}`);
      if (d.selector2CNAME) cnameLines.push(`  selector2._domainkey → ${d.selector2CNAME}`);
    }
    p.note(cnameLines.join("\n"), "Required CNAME records for disabled domains");
  }

  return merged;
}

async function enableDkim(ps: PowerShellSession): Promise<void> {
  const merged = await fetchDomainData(ps);
  const disabled = merged.filter((d) => !d.dkimEnabled);

  if (disabled.length === 0) {
    p.log.success("DKIM is already enabled for all domains.");
    return;
  }

  const selected = await p.select({
    message: "Select a domain to enable DKIM",
    options: [
      ...disabled.map((d) => ({
        value: d.domain,
        label: d.domain,
        hint: d.dkimStatus === "Not configured" ? "will create config" : d.dkimStatus,
      })),
      { value: "back", label: "Back" },
    ],
  });
  if (p.isCancel(selected) || selected === "back") return;

  const domain = selected;
  const spin = p.spinner();
  spin.start(`Enabling DKIM for ${domain}...`);

  // Try Set first (exists but disabled)
  let { error } = await ps.runCommand(
    `Set-DkimSigningConfig -Identity '${escapePS(domain)}' -Enabled $true`
  );

  // If config doesn't exist, create it
  if (error && /not found|couldn't find/i.test(error)) {
    const createResult = await ps.runCommand(
      `New-DkimSigningConfig -DomainName '${escapePS(domain)}' -Enabled $true`
    );
    error = createResult.error;
  }

  if (error && error.includes("Enable-OrganizationCustomization")) {
    spin.stop("Organization customization required.");
    const fixed = await handleOrgCustomizationError(ps, error);
    if (!fixed) return;

    spin.start(`Enabling DKIM for ${domain}...`);
    const retry = await ps.runCommand(
      `Set-DkimSigningConfig -Identity '${escapePS(domain)}' -Enabled $true`
    );
    error = retry.error;
    if (error && /not found|couldn't find/i.test(error)) {
      const createRetry = await ps.runCommand(
        `New-DkimSigningConfig -DomainName '${escapePS(domain)}' -Enabled $true`
      );
      error = createRetry.error;
    }
  }

  if (error) {
    spin.stop("Failed to enable DKIM.");
    // CNAME not published — parse and display required records
    if (/CNAME/i.test(error) || /DNS/i.test(error)) {
      p.log.error("DKIM cannot be enabled because DNS CNAME records are not published.");
      p.log.info("Add the following CNAME records to your DNS, then try again:");

      // Try to fetch the required CNAME values
      const { output: cnameOut } = await ps.runCommand(
        `Get-DkimSigningConfig -Identity '${escapePS(domain)}' | Select-Object Selector1CNAME,Selector2CNAME | ConvertTo-Json -Compress`
      );
      try {
        const cnames = JSON.parse(cnameOut || "{}");
        if (cnames.Selector1CNAME) {
          p.log.info(`  selector1._domainkey.${domain} → ${cnames.Selector1CNAME}`);
        }
        if (cnames.Selector2CNAME) {
          p.log.info(`  selector2._domainkey.${domain} → ${cnames.Selector2CNAME}`);
        }
      } catch {
        p.log.warn(error);
      }
    } else {
      p.log.error(error);
    }
    return;
  }

  spin.stop(`DKIM enabled for ${domain}.`);

  // Verify
  const verifySpin = p.spinner();
  verifySpin.start("Verifying...");
  const verifyRaw = await ps.runCommandJson<DkimConfig>(
    `Get-DkimSigningConfig -Identity '${escapePS(domain)}' | Select-Object Domain,Enabled,Status`
  );
  verifySpin.stop("Verified.");

  if (verifyRaw) {
    p.log.success(`${verifyRaw.Domain}: Enabled=${verifyRaw.Enabled}, Status=${verifyRaw.Status}`);
  }
}

async function exportExcel(ps: PowerShellSession): Promise<void> {
  const merged = await fetchDomainData(ps);

  if (merged.length === 0) {
    p.log.info("No domains to export.");
    return;
  }

  const spin = p.spinner();
  spin.start("Generating Excel report...");

  const tenantSlug = (ps.tenantDomain ?? "tenant").replace(/\./g, "-");
  const dateSlug = new Date().toISOString().slice(0, 10);
  const outputDir = join(appDir(), "reports output");
  const fullPath = resolve(join(outputDir, `${tenantSlug}-dkim-audit-${dateSlug}.xlsx`));
  mkdirSync(dirname(fullPath), { recursive: true });
  try { chmodSync(dirname(fullPath), 0o700); } catch {}

  const buffer = await generateReport({
    sheetName: "DKIM Audit",
    title: "DKIM / Email Authentication Audit",
    tenant: ps.tenantDomain ?? "Unknown",
    summary: `${merged.length} domain(s), ${merged.filter((d) => d.dkimEnabled).length} with DKIM enabled`,
    columns: [
      { header: "Domain", width: 30 },
      { header: "Domain Type", width: 18 },
      { header: "DKIM Enabled", width: 14 },
      { header: "Status", width: 18 },
      { header: "Selector1 CNAME", width: 55 },
      { header: "Selector2 CNAME", width: 55 },
    ],
    rows: merged.map((d) => [
      d.domain,
      d.domainType,
      d.dkimEnabled ? "Yes" : "No",
      d.dkimStatus,
      d.selector1CNAME,
      d.selector2CNAME,
    ]),
  });

  await Bun.write(fullPath, buffer);
  try { chmodSync(fullPath, 0o600); } catch {}
  spin.stop(`Exported ${merged.length} rows to ${fullPath}`);

  const folder = dirname(fullPath);
  try { Bun.spawn(process.platform === "win32" ? ["explorer", folder] : ["open", folder]); } catch {}
}

export async function run(ps: PowerShellSession): Promise<void> {
  await ps.ensureExchangeConnected();

  while (true) {
    const action = await p.select({
      message: "DKIM / Email Authentication",
      options: [
        { value: "view", label: "View all domains & DKIM status" },
        { value: "enable", label: "Enable DKIM for a domain" },
        { value: "export", label: "Export to Excel" },
        { value: "back", label: "Back" },
      ],
    });

    if (p.isCancel(action) || action === "back") return;

    switch (action) {
      case "view":
        await viewDomains(ps);
        break;
      case "enable":
        await enableDkim(ps);
        break;
      case "export":
        await exportExcel(ps);
        break;
    }
  }
}
