import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { escapePS } from "../utils.ts";

interface DistributionGroup {
  DisplayName: string;
  PrimarySmtpAddress: string;
}

export async function run(ps: PowerShellSession, upn: string): Promise<{ name: string; email: string }[]> {
  await ps.ensureExchangeConnected();

  const spin = p.spinner();
  spin.start("Fetching distribution groups...");

  let groups: DistributionGroup[];
  try {
    const raw = await ps.runCommandJson<DistributionGroup | DistributionGroup[]>(
      "Get-DistributionGroup -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress",
    );
    groups = (raw ? (Array.isArray(raw) ? raw : [raw]) : []).sort((a, b) =>
      a.DisplayName.localeCompare(b.DisplayName),
    );
    spin.stop(`Found ${groups.length} distribution group(s).`);
  } catch (e) {
    spin.stop("Failed to fetch distribution groups.");
    p.log.error(`${e}`);
    return [];
  }

  if (groups.length === 0) {
    p.log.warn("No distribution groups found.");
    return [];
  }

  const selectedAddresses = await p.multiselect({
    message: "Select distribution group(s) (space to select, esc to go back)",
    options: groups.map((g) => ({
      value: g.PrimarySmtpAddress,
      label: g.DisplayName,
      hint: g.PrimarySmtpAddress,
    })),
    required: true,
  });
  if (p.isCancel(selectedAddresses)) return [];

  const added: { name: string; email: string }[] = [];

  for (const address of selectedAddresses) {
    const name = groups.find((g) => g.PrimarySmtpAddress === address)?.DisplayName ?? address;

    const addSpin = p.spinner();
    addSpin.start(`Adding ${upn} to ${name}...`);

    const { error } = await ps.runCommand(
      `Add-DistributionGroupMember -Identity '${escapePS(address)}' -Member '${escapePS(upn)}'`,
    );

    if (error) {
      addSpin.stop(`Failed to add to ${name}.`);
      p.log.error(error);
    } else {
      addSpin.stop(`Added ${upn} to ${name}.`);
      added.push({ name, email: address });
    }
  }

  return added;
}
