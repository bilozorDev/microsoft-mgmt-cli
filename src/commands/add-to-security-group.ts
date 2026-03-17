import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { GraphClient } from "../graph-client.ts";

interface SecurityGroup {
  displayName: string;
  id: string;
  mail: string | null;
}

export async function run(ps: PowerShellSession, upn: string): Promise<{ name: string; email: string } | null> {
  // Ensure Graph connected
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph...");
  try {
    await ps.ensureGraphConnected(["User.Read.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All"]);
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return null;
  }

  const graph = new GraphClient(ps);

  const spin = p.spinner();
  spin.start("Fetching security groups...");

  let groups: SecurityGroup[];
  try {
    groups = (
      await graph.getAll<SecurityGroup>("/groups", {
        params: {
          $filter: "securityEnabled eq true",
          $select: "displayName,id,mail",
          $count: "true",
          $top: "999",
        },
        headers: { ConsistencyLevel: "eventual" },
      })
    ).sort((a, b) => a.displayName.localeCompare(b.displayName));
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
      value: g.id,
      label: g.displayName,
      hint: g.mail ?? undefined,
    })),
  });
  if (p.isCancel(groupId)) return null;

  const addSpin = p.spinner();
  const group = groups.find((g) => g.id === groupId);
  const groupName = group?.displayName ?? groupId;
  addSpin.start(`Adding ${upn} to ${groupName}...`);

  try {
    // Resolve user ID from UPN
    const user = await graph.request<{ id: string }>(
      `/users/${encodeURIComponent(upn)}`,
      { params: { $select: "id" } },
    );

    // Add member via Graph REST API
    await graph.post(`/groups/${groupId}/members/$ref`, {
      "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${user.id}`,
    });
  } catch (e: any) {
    addSpin.stop("Failed to add member.");
    p.log.error(e.message);
    return null;
  }

  addSpin.stop(`Added ${upn} to ${groupName}.`);
  return { name: groupName, email: group?.mail ?? "" };
}
