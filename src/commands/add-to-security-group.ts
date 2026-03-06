import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";

interface SecurityGroup {
  DisplayName: string;
  Id: string;
  Mail: string | null;
}

function escapePS(value: string): string {
  return value.replace(/'/g, "''");
}

export async function run(ps: PowerShellSession, upn: string): Promise<string | null> {
  // Ensure Graph connected
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph...");
  try {
    await ps.ensureGraphConnected();
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return null;
  }

  const spin = p.spinner();
  spin.start("Fetching security groups...");

  let groups: SecurityGroup[];
  try {
    const raw = await ps.runCommandJson<SecurityGroup | SecurityGroup[]>(
      'Get-MgGroup -Filter "securityEnabled eq true" -All | Select-Object DisplayName, Id, Mail',
    );
    groups = (Array.isArray(raw) ? raw : [raw]).sort((a, b) =>
      a.DisplayName.localeCompare(b.DisplayName),
    );
    spin.stop(`Found ${groups.length} security group(s).`);
  } catch (e) {
    spin.stop("Failed to fetch security groups.");
    p.log.error(`${e}`);
    return null;
  }

  if (groups.length === 0) {
    p.log.warn("No security groups found.");
    return null;
  }

  const groupId = await p.select({
    message: "Select security group (esc to go back)",
    options: groups.map((g) => ({
      value: g.Id,
      label: g.DisplayName,
      hint: g.Mail ?? undefined,
    })),
  });
  if (p.isCancel(groupId)) return null;

  const addSpin = p.spinner();
  const groupName = groups.find((g) => g.Id === groupId)?.DisplayName ?? groupId;
  addSpin.start(`Adding ${upn} to ${groupName}...`);

  const { error } = await ps.runCommand(
    `New-MgGroupMember -GroupId '${escapePS(groupId)}' -DirectoryObjectId (Get-MgUser -Filter "userPrincipalName eq '${escapePS(upn)}'").Id`,
  );

  if (error) {
    addSpin.stop("Failed to add member.");
    p.log.error(error);
    return null;
  }

  addSpin.stop(`Added ${upn} to ${groupName}.`);
  return groupName;
}
