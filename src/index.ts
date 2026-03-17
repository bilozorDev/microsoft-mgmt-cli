import * as p from "@clack/prompts";
import * as Sentry from "@sentry/bun";
import { PowerShellSession } from "./powershell.ts";
import { checkRequirements } from "./requirements.ts";
import {
  initTelemetry,
  trackCommand,
  setTenantContext,
  flushTelemetry,
} from "./telemetry.ts";
import {
  initAnalytics,
  trackAppLaunch,
  identifyTenant,
  trackCommandEvent,
  shutdownAnalytics,
} from "./analytics.ts";
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
import { run as userInfo } from "./commands/user-info.ts";
import { run as emergencyResponse } from "./commands/emergency-response.ts";
import { checkForUpdates } from "./auto-update.ts";
import pkg from "../package.json";

const ps = new PowerShellSession();

async function runCmd(name: string, fn: () => Promise<void>): Promise<void> {
  trackCommandEvent(name);
  try {
    await trackCommand(name, fn);
  } catch (e) {
    p.log.error(`Unexpected error in ${name}: ${e}`);
  }
}

async function cleanup() {
  if (ps.isExchangeConnected) {
    p.log.info("Disconnecting...");
  }
  await ps.end();
  await shutdownAnalytics();
  await flushTelemetry();
  p.outro("Goodbye!");
  process.exit(0);
}

process.on("SIGINT", () => void cleanup());
process.on("SIGTERM", () => void cleanup());
process.on("uncaughtException", (error) => Sentry.captureException(error));
process.on("unhandledRejection", (reason) => Sentry.captureException(reason));

async function main() {
  initTelemetry();
  initAnalytics();
  trackAppLaunch(pkg.version);
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

  // Set up callback for when tenant is resolved (on first Exchange connection)
  ps.onTenantResolved = (domain) => {
    setTenantContext(domain);
    identifyTenant(domain);
    p.log.success(`Connected to: ${domain}`);
  };

  // Connect to Exchange Online at startup (triggers tenant resolution callback)
  try {
    await ps.ensureExchangeConnected();
  } catch (e) {
    p.log.error(`Failed to connect to Exchange Online: ${e}`);
    p.log.warn("Some commands may not work without an Exchange connection.");
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
        { value: "switch-tenant", label: "Switch tenant", hint: ps.tenantDomain ? `connected to ${ps.tenantDomain}` : "connect to a tenant first", disabled: !ps.isExchangeConnected },
        { value: "exit", label: "Exit" },
      ],
    });

    if (p.isCancel(category) || category === "exit") {
      break;
    }

    if (category === "switch-tenant") {
      const oldDomain = ps.tenantDomain;
      const switchSpin = p.spinner();
      switchSpin.start("Disconnecting...");
      await ps.disconnectTenant();
      switchSpin.stop(`Successfully disconnected from ${oldDomain}`);
      continue;
    }

    if (category === "user-management") {
      const action = await p.select({
        message: "User Management",
        options: [
          { value: "create-user", label: "Create user" },
          { value: "edit-user", label: "Edit user" },
          { value: "delete-user", label: "Delete user" },
          { value: "user-info", label: "User info", hint: "roles, licenses, mailbox, 2FA" },
          { value: "back", label: "Back" },
        ],
      });

      if (p.isCancel(action) || action === "back") continue;

      switch (action) {
        case "user-info":
          await runCmd("user-info", () => userInfo(ps));
          break;
        case "create-user":
          await runCmd("create-user", () => createUser(ps));
          break;
        case "edit-user":
          await runCmd("edit-user", () => editUser(ps));
          break;
        case "delete-user":
          await runCmd("delete-user", () => deleteUser(ps));
          break;
      }
    }

    if (category === "reports") {
      const action = await p.select({
        message: "Reports",
        options: [
          { value: "inactive-users", label: "Inactive users" },
          { value: "shared-mailboxes", label: "Shared mailboxes" },
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
          await runCmd("inactive-users", () => inactiveUsers(ps));
          break;
        case "shared-mailboxes":
          await runCmd("shared-mailboxes", () => sharedMailboxes(ps));
          break;
        case "forwarding-audit":
          await runCmd("forwarding-audit", () => forwardingAudit(ps));
          break;
        case "admin-roles":
          await runCmd("admin-roles", () => adminRoles(ps));
          break;
        case "users-report":
          await runCmd("users-report", () => usersReport(ps));
          break;
        case "license-report":
          await runCmd("license-report", () => licenseReport(ps));
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
          await runCmd("create-group", () => createGroup(ps));
          break;
        case "edit-group":
          await runCmd("edit-group", () => editGroup(ps));
          break;
        case "delete-group":
          await runCmd("delete-group", () => deleteGroup(ps));
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
          await runCmd("whitelist-domain", () => whitelistDomain(ps));
          break;
        case "block-sender":
          await runCmd("block-sender", () => blockSender(ps));
          break;
        case "quarantine":
          await runCmd("quarantine", () => quarantineManagement(ps));
          break;
        case "message-trace":
          await runCmd("message-trace", () => messageTrace(ps));
          break;
        case "dkim-audit":
          await runCmd("dkim-audit", () => dkimAudit(ps));
          break;
      }
    }

    if (category === "emergency-response") {
      await runCmd("emergency-response", () => emergencyResponse(ps));
    }
  }

  await cleanup();
}

main().catch(async (e) => {
  Sentry.captureException(e);
  p.log.error(`Unexpected error: ${e}`);
  await ps.end();
  await shutdownAnalytics();
  await flushTelemetry();
  process.exit(1);
});
