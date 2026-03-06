import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";

interface SharedMailbox {
  DisplayName: string;
  PrimarySmtpAddress: string;
}

function escapePS(value: string): string {
  return value.replace(/'/g, "''");
}

export async function run(ps: PowerShellSession, upn: string): Promise<string[]> {
  const spin = p.spinner();
  spin.start("Fetching shared mailboxes...");

  let mailboxes: SharedMailbox[];
  try {
    const raw = await ps.runCommandJson<SharedMailbox | SharedMailbox[]>(
      "Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress",
    );
    mailboxes = (Array.isArray(raw) ? raw : [raw]).sort((a, b) =>
      a.DisplayName.localeCompare(b.DisplayName),
    );
    spin.stop(`Found ${mailboxes.length} shared mailbox(es).`);
  } catch (e) {
    spin.stop("Failed to fetch shared mailboxes.");
    p.log.error(`${e}`);
    return [];
  }

  if (mailboxes.length === 0) {
    p.log.warn("No shared mailboxes found.");
    return [];
  }

  const selectedAddresses = await p.multiselect({
    message: "Select shared mailbox(es) (space to select, esc to go back)",
    options: mailboxes.map((m) => ({
      value: m.PrimarySmtpAddress,
      label: m.DisplayName,
      hint: m.PrimarySmtpAddress,
    })),
    required: true,
  });
  if (p.isCancel(selectedAddresses)) return [];

  let permissions: string[];
  while (true) {
    const permsChoice = await p.multiselect({
      message: "Select permissions to grant (space to select, esc to go back)",
      options: [
        { value: "read-manage", label: "Read and Manage" },
        { value: "send-as", label: "Send As" },
        { value: "send-on-behalf", label: "Send on Behalf" },
      ],
      required: true,
    });

    if (p.isCancel(permsChoice)) {
      const confirm = await p.select({
        message: "Go back and discard mailbox selection?",
        options: [
          { value: "back-perms", label: "No, return to permissions" },
          { value: "back-menu", label: "Yes, go back" },
        ],
      });
      if (p.isCancel(confirm) || confirm === "back-menu") return [];
      continue;
    }

    permissions = permsChoice;
    break;
  }

  const permLabels: Record<string, string> = {
    "read-manage": "Read and Manage",
    "send-as": "Send As",
    "send-on-behalf": "Send on Behalf",
  };
  const permsSummary = permissions.map((p) => permLabels[p]).join(", ");
  const added: string[] = [];

  for (const mailbox of selectedAddresses) {
    const name = mailboxes.find((m) => m.PrimarySmtpAddress === mailbox)?.DisplayName ?? mailbox;
    const spin = p.spinner();
    spin.start(`Granting permissions on ${name}...`);

    const errors: string[] = [];

    for (const perm of permissions) {
      let result: { error: string };

      if (perm === "read-manage") {
        result = await ps.runCommand(
          `Add-MailboxPermission -Identity '${escapePS(mailbox)}' -User '${escapePS(upn)}' -AccessRights FullAccess -InheritanceType All -AutoMapping $true`,
        );
      } else if (perm === "send-as") {
        result = await ps.runCommand(
          `Add-RecipientPermission -Identity '${escapePS(mailbox)}' -Trustee '${escapePS(upn)}' -AccessRights SendAs -Confirm:$false`,
        );
      } else {
        result = await ps.runCommand(
          `Set-Mailbox -Identity '${escapePS(mailbox)}' -GrantSendOnBehalfTo @{Add='${escapePS(upn)}'}`,
        );
      }

      if (result.error) {
        errors.push(`${permLabels[perm]}: ${result.error}`);
      }
    }

    if (errors.length === 0) {
      spin.stop(`${name}: granted ${permsSummary}.`);
      added.push(name);
    } else if (errors.length < permissions.length) {
      spin.stop(`${name}: some permissions failed.`);
      added.push(name);
      for (const err of errors) p.log.error(err);
    } else {
      spin.stop(`${name}: failed to grant permissions.`);
      for (const err of errors) p.log.error(err);
    }
  }

  return added;
}
