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

  const added: string[] = [];

  for (const mailbox of selectedAddresses) {
    const name = mailboxes.find((m) => m.PrimarySmtpAddress === mailbox)?.DisplayName ?? mailbox;
    let anySuccess = false;

    for (const perm of permissions) {
      const permSpin = p.spinner();

      if (perm === "read-manage") {
        permSpin.start(`Granting Read and Manage on ${name}...`);
        const { error } = await ps.runCommand(
          `Add-MailboxPermission -Identity '${escapePS(mailbox)}' -User '${escapePS(upn)}' -AccessRights FullAccess -InheritanceType All -AutoMapping $true`,
        );
        if (error) {
          permSpin.stop(`Failed to grant Read and Manage on ${name}.`);
          p.log.error(error);
        } else {
          permSpin.stop(`Read and Manage granted on ${name}.`);
          anySuccess = true;
        }
      }

      if (perm === "send-as") {
        permSpin.start(`Granting Send As on ${name}...`);
        const { error } = await ps.runCommand(
          `Add-RecipientPermission -Identity '${escapePS(mailbox)}' -Trustee '${escapePS(upn)}' -AccessRights SendAs -Confirm:$false`,
        );
        if (error) {
          permSpin.stop(`Failed to grant Send As on ${name}.`);
          p.log.error(error);
        } else {
          permSpin.stop(`Send As granted on ${name}.`);
          anySuccess = true;
        }
      }

      if (perm === "send-on-behalf") {
        permSpin.start(`Granting Send on Behalf on ${name}...`);
        const { error } = await ps.runCommand(
          `Set-Mailbox -Identity '${escapePS(mailbox)}' -GrantSendOnBehalfTo @{Add='${escapePS(upn)}'}`,
        );
        if (error) {
          permSpin.stop(`Failed to grant Send on Behalf on ${name}.`);
          p.log.error(error);
        } else {
          permSpin.stop(`Send on Behalf granted on ${name}.`);
          anySuccess = true;
        }
      }
    }

    if (anySuccess) added.push(name);
  }

  return added;
}
