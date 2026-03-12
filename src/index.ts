import * as p from "@clack/prompts";
import { PowerShellSession } from "./powershell.ts";
import { checkRequirements } from "./requirements.ts";
import { run as whitelistDomain } from "./commands/whitelist-domain.ts";
import { run as blockSender } from "./commands/block-sender.ts";
import { run as quarantineManagement } from "./commands/quarantine-management.ts";
import { run as messageTrace } from "./commands/message-trace.ts";
import { run as dkimAudit } from "./commands/dkim-audit.ts";
import { run as createUser } from "./commands/create-user.ts";
import { run as editUser } from "./commands/edit-user.ts";
import { run as deleteUser } from "./commands/delete-user.ts";
import { run as inactiveUsers } from "./commands/inactive-users.ts";
import { run as sharedMailboxes } from "./commands/shared-mailboxes.ts";
import { run as forwardingAudit } from "./commands/forwarding-audit.ts";
import { run as adminRoles } from "./commands/admin-roles.ts";
import { run as usersReport } from "./commands/users-report.ts";
import { run as licenseReport } from "./commands/license-report.ts";
import { run as createGroup } from "./commands/create-group.ts";
import { run as editGroup } from "./commands/edit-group.ts";
import { run as deleteGroup } from "./commands/delete-group.ts";
import { run as emergencyResponse } from "./commands/emergency-response.ts";
import { checkForUpdates } from "./auto-update.ts";
import pkg from "../package.json";

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
  p.intro(`Microsoft 365 Admin CLI v${pkg.version}`);

  await checkRequirements();
  await checkForUpdates();

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

  // Extract tenant ID for Graph isolation
  await ps.extractTenantId();

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
        { value: "group-management", label: "Groups & Shared Mailbox Management" },
        { value: "email-security", label: "Email Security" },
        { value: "reports", label: "Reports" },
        { value: "emergency-response", label: "Emergency Response", hint: "compromised account" },
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
          { value: "forwarding-audit", label: "Forwarding audit", hint: "security & compliance" },
          { value: "admin-roles", label: "Admin role report", hint: "security audit" },
          { value: "users-report", label: "Licensed users report", hint: "licenses, MFA, mailbox size" },
          { value: "license-report", label: "License & subscription report", hint: "expiration, renewal, usage" },
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
        case "forwarding-audit":
          await forwardingAudit(ps);
          break;
        case "admin-roles":
          await adminRoles(ps);
          break;
        case "users-report":
          await usersReport(ps);
          break;
        case "license-report":
          await licenseReport(ps);
          break;
      }
    }

    if (category === "group-management") {
      const action = await p.select({
        message: "Groups & Shared Mailbox Management",
        options: [
          { value: "create-group", label: "Create group or shared mailbox" },
          { value: "edit-group", label: "Edit group or shared mailbox" },
          { value: "delete-group", label: "Delete group or shared mailbox" },
          { value: "back", label: "Back" },
        ],
      });

      if (p.isCancel(action) || action === "back") continue;

      switch (action) {
        case "create-group":
          await createGroup(ps);
          break;
        case "edit-group":
          await editGroup(ps);
          break;
        case "delete-group":
          await deleteGroup(ps);
          break;
      }
    }

    if (category === "email-security") {
      const action = await p.select({
        message: "Email Security",
        options: [
          { value: "whitelist-domain", label: "Whitelist domain(s)" },
          { value: "block-sender", label: "Block sender/domain(s)" },
          { value: "quarantine", label: "Quarantine management" },
          { value: "message-trace", label: "Message trace" },
          { value: "dkim-audit", label: "DKIM / Email authentication" },
          { value: "back", label: "Back" },
        ],
      });

      if (p.isCancel(action) || action === "back") continue;

      switch (action) {
        case "whitelist-domain":
          await whitelistDomain(ps);
          break;
        case "block-sender":
          await blockSender(ps);
          break;
        case "quarantine":
          await quarantineManagement(ps);
          break;
        case "message-trace":
          await messageTrace(ps);
          break;
        case "dkim-audit":
          await dkimAudit(ps);
          break;
      }
    }

    if (category === "emergency-response") {
      await emergencyResponse(ps);
    }
  }

  await cleanup();
}

main().catch(async (e) => {
  p.log.error(`Unexpected error: ${e}`);
  await ps.end();
  process.exit(1);
});
