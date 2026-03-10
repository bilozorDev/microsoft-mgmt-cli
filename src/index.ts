import * as p from "@clack/prompts";
import { PowerShellSession } from "./powershell.ts";
import { checkRequirements } from "./requirements.ts";
import { run as whitelistDomain } from "./commands/whitelist-domain.ts";
import { run as createUser } from "./commands/create-user.ts";
import { run as editUser } from "./commands/edit-user.ts";
import { run as deleteUser } from "./commands/delete-user.ts";
import { run as inactiveUsers } from "./commands/inactive-users.ts";
import { run as sharedMailboxes } from "./commands/shared-mailboxes.ts";

const ps = new PowerShellSession();

async function cleanup() {
  p.log.info("Disconnecting from Exchange Online...");
  await ps.end();
  p.outro("Goodbye!");
  process.exit(0);
}

process.on("SIGINT", () => void cleanup());
process.on("SIGTERM", () => void cleanup());

async function main() {
  p.intro("Profulgent — Exchange Online Admin CLI");

  await checkRequirements();

  // Start PowerShell session
  const connectSpin = p.spinner();
  connectSpin.start("Starting PowerShell session...");
  try {
    await ps.start();
    connectSpin.stop("PowerShell session ready.");
  } catch (e) {
    connectSpin.stop("Failed to start PowerShell.");
    p.log.error(`Could not start pwsh. Is PowerShell Core installed?\n${e}`);
    process.exit(1);
  }

  // Connect to Exchange Online (opens browser for auth)
  const authSpin = p.spinner();
  authSpin.start("Connecting to Exchange Online (check your browser)...");
  try {
    await ps.connectExchangeOnline();
    authSpin.stop("Connected to Exchange Online.");
  } catch (e) {
    authSpin.stop("Connection failed.");
    p.log.error(`${e}`);
    await ps.end();
    process.exit(1);
  }

  // Fetch tenant domain
  try {
    const domain = await ps.getTenantDomain();
    p.log.success(`Connected to: ${domain}`);
  } catch {
    p.log.warn("Could not determine tenant domain.");
  }

  // Main menu loop
  while (true) {
    const category = await p.select({
      message: "What would you like to do?",
      options: [
        { value: "user-management", label: "User Management" },
        { value: "spam-management", label: "Spam Management" },
        { value: "reports", label: "Reports" },
        { value: "switch-tenant", label: "Switch tenant", hint: ps.tenantDomain ? `connected to ${ps.tenantDomain}` : undefined },
        { value: "exit", label: "Exit" },
      ],
    });

    if (p.isCancel(category) || category === "exit") {
      break;
    }

    if (category === "switch-tenant") {
      const switchSpin = p.spinner();
      switchSpin.start("Disconnecting...");
      try {
        await ps.switchTenant();
        switchSpin.stop(`Switched to: ${ps.tenantDomain}`);
      } catch (e) {
        switchSpin.stop("Failed to switch tenant.");
        p.log.error(`${e}`);
        await ps.end();
        process.exit(1);
      }
      continue;
    }

    if (category === "user-management") {
      const action = await p.select({
        message: "User Management",
        options: [
          { value: "create-user", label: "Create user", hint: "will prompt to login" },
          { value: "edit-user", label: "Edit user", hint: "will prompt to login" },
          { value: "delete-user", label: "Delete user", hint: "will prompt to login" },
          { value: "back", label: "Back" },
        ],
      });

      if (p.isCancel(action) || action === "back") continue;

      switch (action) {
        case "create-user":
          await createUser(ps);
          break;
        case "edit-user":
          await editUser(ps);
          break;
        case "delete-user":
          await deleteUser(ps);
          break;
      }
    }

    if (category === "reports") {
      const action = await p.select({
        message: "Reports",
        options: [
          { value: "inactive-users", label: "Inactive users" },
          { value: "shared-mailboxes", label: "Shared mailboxes", hint: "will prompt to login" },
          { value: "back", label: "Back" },
        ],
      });

      if (p.isCancel(action) || action === "back") continue;

      switch (action) {
        case "inactive-users":
          await inactiveUsers(ps);
          break;
        case "shared-mailboxes":
          await sharedMailboxes(ps);
          break;
      }
    }

    if (category === "spam-management") {
      const action = await p.select({
        message: "Spam Management",
        options: [
          { value: "whitelist-domain", label: "Whitelist domain(s)" },
          { value: "back", label: "Back" },
        ],
      });

      if (p.isCancel(action) || action === "back") continue;

      switch (action) {
        case "whitelist-domain":
          await whitelistDomain(ps);
          break;
      }
    }
  }

  await cleanup();
}

main().catch(async (e) => {
  p.log.error(`Unexpected error: ${e}`);
  await ps.end();
  process.exit(1);
});
