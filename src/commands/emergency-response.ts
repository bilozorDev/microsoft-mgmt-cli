import { resolve, dirname, join } from "path";
import { mkdirSync, chmodSync } from "fs";
import * as p from "@clack/prompts";
import type { PowerShellSession } from "../powershell.ts";
import { generatePassword } from "../password.ts";
import { escapePS, createSecretLink, appDir } from "../utils.ts";
import {
  type AuthMethod,
  MFA_DETAIL_CMDS,
  MFA_REMOVE_CMDLETS,
  friendlyMfaMethod,
  mfaTypeKey,
} from "../mfa-utils.ts";

interface MgUser {
  DisplayName: string;
  UserPrincipalName: string;
  Id: string;
  LicenseCount: number;
}

interface ForwardingConfig {
  ForwardingSmtpAddress: string | null;
  ForwardingAddress: string | null;
  DeliverToMailboxAndForward: boolean;
}

interface InboxRule {
  Identity: string;
  Name: string;
  Enabled: boolean;
  ForwardTo: string | null;
  ForwardAsAttachmentTo: string | null;
  RedirectTo: string | null;
  DeleteMessage: boolean;
  MoveToFolder: string | null;
  Description: string | null;
}

interface MailboxPermission {
  User: string;
  AccessRights: string | string[];
}

interface RecipientPermission {
  Trustee: string;
  AccessRights: string | string[];
}

interface TraceMessage {
  MessageId: string;
  MessageTraceId: string;
  SenderAddress: string;
  RecipientAddress: string;
  Subject: string;
  Status: string;
  Received: string;
  Size: number;
}

interface IncidentFindings {
  user: { displayName: string; upn: string; id: string };
  timestamp: string;
  passwordReset: boolean;
  newPassword: string | null;
  sessionsRevoked: boolean;
  mfaMethodsBefore: { name: string; detail: string }[];
  mfaMethodsRemoved: string[];
  forwardingFound: { smtp: string | null; address: string | null; deliverToBoth: boolean } | null;
  forwardingRemoved: boolean;
  forwardingChecked: boolean;
  inboxRules: { name: string; enabled: boolean; forwardTo: string; redirectTo: string; deleteMessage: boolean; moveToFolder: string; description: string }[];
  inboxRulesChecked: boolean;
  rulesRemoved: string[];
  permissions: { user: string; type: string; rights: string }[];
  permissionsChecked: boolean;
  messageTrace: { recipient: string; subject: string; status: string; received: string }[];
  messageTraceChecked: boolean;
}

function truncate(s: string, len: number): string {
  return s.length > len ? s.slice(0, len - 1) + "…" : s;
}

function dateOffset(days: number): string {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d.toISOString().slice(0, 19);
}

async function copyToClipboard(ps: PowerShellSession, text: string): Promise<boolean> {
  try {
    if (process.platform === "win32") {
      await ps.runCommand(`Set-Clipboard -Value '${escapePS(text)}'`);
    } else {
      const proc = Bun.spawn(["pbcopy"], { stdin: new Blob([text]) });
      await proc.exited;
    }
    return true;
  } catch {
    return false;
  }
}

async function fetchUsers(ps: PowerShellSession): Promise<MgUser[]> {
  const spin = p.spinner();
  spin.start("Loading users...");

  const { output: countOutput } = await ps.runCommand(
    "Get-MgUser -Top 1 -CountVariable ct -ConsistencyLevel eventual | Out-Null; $ct",
  );
  const count = parseInt(countOutput.trim(), 10);

  let users: MgUser[];

  if (count <= 50) {
    const raw = await ps.runCommandJson<MgUser | MgUser[]>(
      "Get-MgUser -All -Property DisplayName,UserPrincipalName,Id,AssignedLicenses | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName; Id = $_.Id; LicenseCount = $_.AssignedLicenses.Count } }",
    );
    users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
    spin.stop(`Found ${users.length} user(s).`);
  } else {
    spin.stop(`${count} users in tenant — search to find a user.`);
    while (true) {
      const query = await p.text({
        message: "Search for user by name",
        placeholder: "e.g. Jane Doe",
        validate: (v = "") => (!v.trim() ? "Enter a search term" : undefined),
      });
      if (p.isCancel(query)) return [];

      const searchSpin = p.spinner();
      searchSpin.start("Searching users...");
      try {
        const raw = await ps.runCommandJson<MgUser | MgUser[]>(
          `Get-MgUser -Search '"displayName:${escapePS(query)}"' -ConsistencyLevel eventual -Property DisplayName,UserPrincipalName,Id,AssignedLicenses | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName; Id = $_.Id; LicenseCount = $_.AssignedLicenses.Count } }`,
        );
        users = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        searchSpin.stop(`Found ${users.length} user(s).`);
      } catch {
        searchSpin.stop("Search returned no results.");
        users = [];
      }

      if (users.length === 0) {
        p.log.warn("No users found. Try a different search term.");
        continue;
      }
      break;
    }

    return users.sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));
  }

  return users.sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));
}

async function selectUser(ps: PowerShellSession): Promise<MgUser | null> {
  const users = await fetchUsers(ps);
  if (users.length === 0) return null;

  const userId = await p.select({
    message: "Select compromised user",
    options: users.map((u) => ({
      value: u.Id,
      label: u.DisplayName,
      hint: u.UserPrincipalName,
    })),
  });
  if (p.isCancel(userId)) return null;

  return users.find((u) => u.Id === userId) ?? null;
}

export async function run(ps: PowerShellSession): Promise<void> {
  // Ensure Graph connected with write access
  const graphSpin = p.spinner();
  graphSpin.start("Connecting to Microsoft Graph (check your browser)...");
  try {
    await ps.ensureGraphConnected(["User.ReadWrite.All", "User-PasswordProfile.ReadWrite.All", "UserAuthenticationMethod.ReadWrite.All"]);
    graphSpin.stop("Connected to Microsoft Graph.");
  } catch (e) {
    graphSpin.stop("Failed to connect to Microsoft Graph.");
    p.log.error(`${e}`);
    return;
  }

  // Select compromised user
  const user = await selectUser(ps);
  if (!user) return;

  const findings: IncidentFindings = {
    user: { displayName: user.DisplayName, upn: user.UserPrincipalName, id: user.Id },
    timestamp: new Date().toISOString(),
    passwordReset: false,
    newPassword: null,
    sessionsRevoked: false,
    mfaMethodsBefore: [] as { name: string; detail: string }[],
    mfaMethodsRemoved: [],
    forwardingFound: null,
    forwardingRemoved: false,
    forwardingChecked: false,
    inboxRules: [],
    inboxRulesChecked: false,
    rulesRemoved: [],
    permissions: [],
    permissionsChecked: false,
    messageTrace: [],
    messageTraceChecked: false,
  };

  const upn = user.UserPrincipalName;
  const userId = user.Id;

  // Sub-menu loop
  while (true) {
    const action = await p.select({
      message: `Emergency Response — ${upn}`,
      options: [
        { value: "revoke-sessions", label: "Force sign-out", hint: "revoke all sessions" },
        { value: "remove-mfa", label: "Remove all MFA methods" },
        { value: "reset-password", label: "Reset password" },
        { value: "check-forwarding", label: "Check & remove forwarding" },
        { value: "check-inbox-rules", label: "Check & remove inbox rules" },
        { value: "audit-permissions", label: "Audit mailbox permissions" },
        { value: "message-trace", label: "Message trace", hint: "last 10 days" },
        { value: "ticket-notes", label: "Generate ticket notes" },
        { value: "back", label: "Back" },
      ],
    });

    if (p.isCancel(action) || action === "back") break;

    switch (action) {
      // ── 1. Force Sign-Out ──────────────────────────────────────────
      case "revoke-sessions": {
        const spin = p.spinner();
        spin.start("Revoking all sessions...");
        const { error } = await ps.runCommand(
          `Revoke-MgUserSignInSession -UserId '${escapePS(userId)}'`,
        );
        if (error) {
          spin.stop("Failed to revoke sessions.");
          p.log.error(error);
        } else {
          spin.stop("All sessions revoked — user is signed out everywhere.");
          findings.sessionsRevoked = true;
        }
        break;
      }

      // ── 2. Remove All MFA Methods ─────────────────────────────────
      case "remove-mfa": {
        const spin = p.spinner();
        spin.start("Fetching MFA methods...");

        const raw = await ps.runCommandJson<AuthMethod | AuthMethod[]>(
          `Get-MgUserAuthenticationMethod -UserId '${escapePS(userId)}' | ForEach-Object { [PSCustomObject]@{ Id = $_.Id; ODataType = $_.AdditionalProperties['@odata.type'] } }`,
        );
        const methods = raw ? (Array.isArray(raw) ? raw : [raw]) : [];

        // Filter to removable methods (exclude password)
        const removable = methods.filter((m) => {
          if (!m.ODataType) return false;
          const key = mfaTypeKey(m.ODataType);
          return key !== "passwordAuthenticationMethod" && key in MFA_REMOVE_CMDLETS;
        });

        if (removable.length === 0) {
          spin.stop("No removable MFA methods found (only password).");
          break;
        }

        // Fetch details for each method
        spin.message("Fetching method details...");
        const methodDetails: { method: AuthMethod; friendly: string; detail: string }[] = [];

        for (const m of removable) {
          const key = mfaTypeKey(m.ODataType!);
          const friendly = friendlyMfaMethod(m.ODataType!) ?? key;
          let detail = "";

          const detailFetcher = MFA_DETAIL_CMDS[key];
          if (detailFetcher) {
            try {
              const raw = await ps.runCommandJson<Record<string, unknown>>(
                detailFetcher.cmd(userId, m.Id),
              );
              if (raw) detail = detailFetcher.format(raw);
            } catch {
              // detail stays empty
            }
          }

          methodDetails.push({ method: m, friendly, detail });
        }

        // Record all methods before removal (with details)
        findings.mfaMethodsBefore = methodDetails.map((d) => ({
          name: d.friendly,
          detail: d.detail,
        }));

        spin.stop(`Found ${removable.length} MFA method(s).`);

        // Display methods with details
        const mfaLines = methodDetails.map((d) =>
          d.detail ? `${d.friendly}: ${d.detail}` : d.friendly,
        );
        p.note(mfaLines.join("\n"), "MFA Methods");

        const confirm = await p.confirm({
          message: `Remove all ${removable.length} MFA method(s)? This cannot be undone.`,
        });
        if (p.isCancel(confirm) || !confirm) break;

        const removeSpin = p.spinner();
        removeSpin.start("Removing MFA methods...");
        const removed: string[] = [];
        const failed: string[] = [];

        // Step 1: First pass — attempt all removals
        let retryMethods: typeof removable = [];
        for (const method of removable) {
          const key = mfaTypeKey(method.ODataType!);
          const info = MFA_REMOVE_CMDLETS[key]!;
          const friendly = friendlyMfaMethod(method.ODataType!) ?? key;

          removeSpin.message(`Removing ${friendly}...`);
          const { error } = await ps.runCommand(
            `${info.cmdlet} -UserId '${escapePS(userId)}' ${info.param} '${escapePS(method.Id)}'`,
          );
          if (error) {
            retryMethods.push(method);
          } else {
            removed.push(friendly);
          }
        }

        // Step 2: Simple retry — default may have auto-shifted after step 1
        if (retryMethods.length > 0) {
          const stillFailing: typeof removable = [];
          for (const method of retryMethods) {
            const key = mfaTypeKey(method.ODataType!);
            const info = MFA_REMOVE_CMDLETS[key]!;
            const friendly = friendlyMfaMethod(method.ODataType!) ?? key;

            removeSpin.message(`Retrying ${friendly}...`);
            const { error } = await ps.runCommand(
              `${info.cmdlet} -UserId '${escapePS(userId)}' ${info.param} '${escapePS(method.Id)}'`,
            );
            if (error) {
              stillFailing.push(method);
            } else {
              removed.push(friendly);
            }
          }
          retryMethods = stillFailing;
        }

        // Step 3: Beta API fallback — change default method, then retry
        if (retryMethods.length > 0) {
          // Determine which method types we're NOT deleting to use as new default
          const deletingKeys = new Set(retryMethods.map((m) => mfaTypeKey(m.ODataType!)));
          const preferenceOptions: { pref: string; key: string }[] = [
            { pref: "push", key: "microsoftAuthenticatorAuthenticationMethod" },
            { pref: "oath", key: "softwareOathAuthenticationMethod" },
            { pref: "sms", key: "phoneAuthenticationMethod" },
            { pref: "voiceMobile", key: "phoneAuthenticationMethod" },
            { pref: "voiceAlternateMobile", key: "phoneAuthenticationMethod" },
            { pref: "voiceOffice", key: "phoneAuthenticationMethod" },
          ];

          // Try each preference that doesn't correspond to a method being deleted
          let defaultChanged = false;
          let newDefaultKey: string | null = null;
          for (const opt of preferenceOptions) {
            if (deletingKeys.has(opt.key)) continue;

            removeSpin.message("Changing default MFA method via beta API...");
            const { error } = await ps.runCommand(
              `Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/users/${escapePS(userId)}/authentication/signInPreferences' -Method PATCH -Body '{"userPreferredMethodForSecondaryAuthentication":"${opt.pref}"}' -ContentType 'application/json'`,
            );
            if (!error) {
              defaultChanged = true;
              newDefaultKey = opt.key;
              break;
            }
          }

          // If we're deleting all method types, try each preference anyway
          if (!defaultChanged) {
            for (const opt of preferenceOptions) {
              removeSpin.message("Changing default MFA method via beta API...");
              const { error } = await ps.runCommand(
                `Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/users/${escapePS(userId)}/authentication/signInPreferences' -Method PATCH -Body '{"userPreferredMethodForSecondaryAuthentication":"${opt.pref}"}' -ContentType 'application/json'`,
              );
              if (!error) {
                defaultChanged = true;
                newDefaultKey = opt.key;
                break;
              }
            }
          }

          // Retry failed methods after changing default
          // Delete non-default methods first, then the new default last
          if (defaultChanged) {
            const nonDefault = retryMethods.filter((m) => mfaTypeKey(m.ODataType!) !== newDefaultKey);
            const isDefault = retryMethods.filter((m) => mfaTypeKey(m.ODataType!) === newDefaultKey);
            const ordered = [...nonDefault, ...isDefault];

            const finalFailing: typeof removable = [];
            for (const method of ordered) {
              const key = mfaTypeKey(method.ODataType!);
              const info = MFA_REMOVE_CMDLETS[key]!;
              const friendly = friendlyMfaMethod(method.ODataType!) ?? key;

              removeSpin.message(`Retrying ${friendly}...`);
              const { error } = await ps.runCommand(
                `${info.cmdlet} -UserId '${escapePS(userId)}' ${info.param} '${escapePS(method.Id)}'`,
              );
              if (error) {
                finalFailing.push(method);
              } else {
                removed.push(friendly);
              }
            }
            retryMethods = finalFailing;
          }
        }

        // Build final failed list from any remaining methods
        for (const method of retryMethods) {
          const key = mfaTypeKey(method.ODataType!);
          const friendly = friendlyMfaMethod(method.ODataType!) ?? key;
          failed.push(`${friendly}: could not remove (may be default method)`);
        }

        findings.mfaMethodsRemoved = removed;

        if (failed.length > 0) {
          removeSpin.stop(`Removed ${removed.length}, failed ${failed.length}.`);
          for (const f of failed) p.log.error(f);
        } else {
          removeSpin.stop(`Removed all ${removed.length} MFA method(s).`);
        }
        break;
      }

      // ── 3. Reset Password ──────────────────────────────────────────
      case "reset-password": {
        const password = generatePassword();

        const spin = p.spinner();
        spin.start("Resetting password...");
        const { error } = await ps.runCommand(
          `Update-MgUser -UserId '${escapePS(userId)}' -PasswordProfile @{ Password = '${escapePS(password)}'; ForceChangePasswordNextSignIn = $true }`,
        );
        if (error) {
          spin.stop("Failed to reset password.");
          p.log.error(error);
          break;
        }
        spin.stop("Password reset (user must change at next sign-in).");
        findings.passwordReset = true;
        findings.newPassword = password;

        p.note(
          [`UPN:      ${upn}`, `Password: ${password}`].join("\n"),
          "New credentials",
        );

        const otsSpin = p.spinner();
        otsSpin.start("Creating one-time secret link...");
        const otsResult = await createSecretLink(password);
        if ("url" in otsResult) {
          otsSpin.stop("One-time secret link created.");
          p.log.info(`Secret link: ${otsResult.url}`);
          const copied = await copyToClipboard(ps, otsResult.url);
          if (copied) p.log.success("Link copied to clipboard.");
        } else {
          otsSpin.stop("Failed to create one-time secret link.");
          p.log.error("error" in otsResult ? otsResult.error : "Unknown error");
          // Fallback: copy password directly
          const copied = await copyToClipboard(ps, password);
          if (copied) p.log.success("Password copied to clipboard.");
        }
        break;
      }

      // ── 4. Check & Remove Forwarding ───────────────────────────────
      case "check-forwarding": {
        const spin = p.spinner();
        spin.start("Checking mailbox forwarding...");

        const raw = await ps.runCommandJson<ForwardingConfig>(
          `Get-Mailbox -Identity '${escapePS(upn)}' | Select-Object ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward`,
        );

        if (!raw) {
          spin.stop("Could not read mailbox forwarding.");
          break;
        }

        findings.forwardingChecked = true;

        const hasForwarding =
          (raw.ForwardingSmtpAddress && raw.ForwardingSmtpAddress !== "") ||
          (raw.ForwardingAddress && raw.ForwardingAddress !== "");

        if (!hasForwarding) {
          spin.stop("No forwarding configured.");
          findings.forwardingFound = { smtp: null, address: null, deliverToBoth: false };
          break;
        }

        findings.forwardingFound = {
          smtp: raw.ForwardingSmtpAddress,
          address: raw.ForwardingAddress,
          deliverToBoth: raw.DeliverToMailboxAndForward,
        };

        spin.stop("Forwarding found!");
        p.note(
          [
            `SMTP Forward:       ${raw.ForwardingSmtpAddress ?? "(none)"}`,
            `Forward Address:    ${raw.ForwardingAddress ?? "(none)"}`,
            `Deliver to both:    ${raw.DeliverToMailboxAndForward}`,
          ].join("\n"),
          "Mailbox Forwarding",
        );

        const confirm = await p.confirm({
          message: "Remove all forwarding from this mailbox?",
        });
        if (p.isCancel(confirm) || !confirm) break;

        const removeSpin = p.spinner();
        removeSpin.start("Removing forwarding...");
        const { error } = await ps.runCommand(
          `Set-Mailbox -Identity '${escapePS(upn)}' -ForwardingSmtpAddress $null -ForwardingAddress $null -DeliverToMailboxAndForward $false`,
        );
        if (error) {
          removeSpin.stop("Failed to remove forwarding.");
          p.log.error(error);
        } else {
          removeSpin.stop("Forwarding removed.");
          findings.forwardingRemoved = true;
        }
        break;
      }

      // ── 5. Check & Remove Inbox Rules ──────────────────────────────
      case "check-inbox-rules": {
        const spin = p.spinner();
        spin.start("Fetching inbox rules...");

        const raw = await ps.runCommandJson<InboxRule | InboxRule[]>(
          `Get-InboxRule -Mailbox '${escapePS(upn)}' | Select-Object Identity,Name,Enabled,ForwardTo,ForwardAsAttachmentTo,RedirectTo,DeleteMessage,MoveToFolder,Description`,
        );
        const rules = raw ? (Array.isArray(raw) ? raw : [raw]) : [];

        findings.inboxRulesChecked = true;
        findings.inboxRules = rules.map((r) => ({
          name: r.Name,
          enabled: r.Enabled,
          forwardTo: [r.ForwardTo, r.ForwardAsAttachmentTo].filter(Boolean).join(", ") || "",
          redirectTo: r.RedirectTo ?? "",
          deleteMessage: r.DeleteMessage,
          moveToFolder: r.MoveToFolder ?? "",
          description: r.Description ?? "",
        }));

        if (rules.length === 0) {
          spin.stop("No inbox rules found.");
          break;
        }

        spin.stop(`Found ${rules.length} inbox rule(s).`);

        // Display rules with full detail
        const lines: string[] = [];
        for (const r of rules) {
          const suspicious = r.ForwardTo || r.ForwardAsAttachmentTo || r.RedirectTo || r.DeleteMessage;
          const marker = suspicious ? " [SUSPICIOUS]" : "";
          lines.push(`${r.Enabled ? "[ON] " : "[OFF]"} ${r.Name}${marker}`);

          // Show actions
          const actions: string[] = [];
          if (r.ForwardTo) actions.push(`Forward to: ${r.ForwardTo}`);
          if (r.ForwardAsAttachmentTo) actions.push(`Forward as attachment to: ${r.ForwardAsAttachmentTo}`);
          if (r.RedirectTo) actions.push(`Redirect to: ${r.RedirectTo}`);
          if (r.DeleteMessage) actions.push("Delete message");
          if (r.MoveToFolder) actions.push(`Move to: ${r.MoveToFolder}`);
          for (const a of actions) lines.push(`      Action: ${a}`);

          // Show description (contains conditions)
          if (r.Description) {
            for (const descLine of r.Description.split("\n").filter((l) => l.trim())) {
              lines.push(`      ${descLine.trim()}`);
            }
          }
          lines.push("");
        }
        p.note(lines.join("\n").trimEnd(), "Inbox Rules");

        // Build multiselect with action hints
        const ruleOptions = rules.map((r) => {
          const parts: string[] = [];
          if (r.ForwardTo) parts.push(`fwd: ${r.ForwardTo}`);
          if (r.ForwardAsAttachmentTo) parts.push(`fwd-attach: ${r.ForwardAsAttachmentTo}`);
          if (r.RedirectTo) parts.push(`redirect: ${r.RedirectTo}`);
          if (r.DeleteMessage) parts.push("deletes messages");
          if (r.MoveToFolder) parts.push(`move: ${r.MoveToFolder}`);
          const isSuspicious = r.ForwardTo || r.ForwardAsAttachmentTo || r.RedirectTo || r.DeleteMessage;
          const hint = [isSuspicious ? "SUSPICIOUS" : null, ...parts].filter(Boolean).join(" — ");
          return { value: r.Identity, label: r.Name, hint: hint || undefined };
        });

        const rulesToRemove = await p.multiselect({
          message: "Select rules to remove (space to toggle, enter to confirm)",
          options: ruleOptions,
          required: false,
        });

        if (p.isCancel(rulesToRemove) || rulesToRemove.length === 0) break;

        const removeSpin = p.spinner();
        removeSpin.start("Removing selected rules...");
        for (const identity of rulesToRemove) {
          const { error } = await ps.runCommand(
            `Remove-InboxRule -Identity '${escapePS(identity)}' -Confirm:$false`,
          );
          if (error) {
            p.log.error(`Failed to remove rule "${identity}": ${error}`);
          } else {
            const ruleName = rules.find((r) => r.Identity === identity)?.Name ?? identity;
            findings.rulesRemoved.push(ruleName);
          }
        }
        removeSpin.stop(`Removed ${findings.rulesRemoved.length} rule(s).`);
        break;
      }

      // ── 6. Audit Mailbox Permissions ───────────────────────────────
      case "audit-permissions": {
        const spin = p.spinner();
        spin.start("Auditing mailbox permissions...");

        // Full Access
        const fullAccessRaw = await ps.runCommandJson<MailboxPermission | MailboxPermission[]>(
          `Get-MailboxPermission -Identity '${escapePS(upn)}' | Where-Object { $_.User -ne 'NT AUTHORITY\\SELF' -and $_.IsInherited -eq $false } | Select-Object User,AccessRights`,
        );
        const fullAccess = fullAccessRaw ? (Array.isArray(fullAccessRaw) ? fullAccessRaw : [fullAccessRaw]) : [];

        // Send As
        const sendAsRaw = await ps.runCommandJson<RecipientPermission | RecipientPermission[]>(
          `Get-RecipientPermission -Identity '${escapePS(upn)}' | Where-Object { $_.Trustee -ne 'NT AUTHORITY\\SELF' } | Select-Object Trustee,AccessRights`,
        );
        const sendAs = sendAsRaw ? (Array.isArray(sendAsRaw) ? sendAsRaw : [sendAsRaw]) : [];

        // Send on Behalf
        const sobRaw = await ps.runCommandJson<string | string[]>(
          `Get-Mailbox -Identity '${escapePS(upn)}' | Select-Object -ExpandProperty GrantSendOnBehalfTo`,
        );
        const sendOnBehalf = sobRaw ? (Array.isArray(sobRaw) ? sobRaw : [sobRaw]) : [];

        findings.permissionsChecked = true;

        const allPerms: { user: string; type: string; rights: string }[] = [];
        for (const fa of fullAccess) {
          const rights = Array.isArray(fa.AccessRights) ? fa.AccessRights.join(", ") : String(fa.AccessRights);
          allPerms.push({ user: fa.User, type: "FullAccess", rights });
        }
        for (const sa of sendAs) {
          const rights = Array.isArray(sa.AccessRights) ? sa.AccessRights.join(", ") : String(sa.AccessRights);
          allPerms.push({ user: sa.Trustee, type: "SendAs", rights });
        }
        for (const sob of sendOnBehalf) {
          allPerms.push({ user: sob, type: "SendOnBehalf", rights: "SendOnBehalf" });
        }

        findings.permissions = allPerms;

        if (allPerms.length === 0) {
          spin.stop("No non-inherited mailbox permissions found.");
          break;
        }

        spin.stop(`Found ${allPerms.length} permission(s).`);

        const header = `${"User".padEnd(38)} ${"Type".padEnd(14)} Rights`;
        const separator = "─".repeat(header.length + 10);
        const permLines = [header, separator];
        for (const perm of allPerms) {
          permLines.push(
            `${truncate(perm.user, 37).padEnd(38)} ${perm.type.padEnd(14)} ${perm.rights}`,
          );
        }
        p.note(permLines.join("\n"), "Mailbox Permissions");
        break;
      }

      // ── 7. Message Trace (Last 10 Days) ────────────────────────────
      case "message-trace": {
        const spin = p.spinner();
        spin.start("Running message trace (last 10 days)...");

        const startDate = dateOffset(10);
        const endDate = new Date().toISOString().slice(0, 19);
        let cmd = `Get-MessageTraceV2 -SenderAddress '${escapePS(upn)}' -StartDate '${startDate}' -EndDate '${endDate}' | Select-Object MessageId,MessageTraceId,SenderAddress,RecipientAddress,Subject,Status,Received,Size`;

        let raw: TraceMessage | TraceMessage[] | null = null;
        try {
          raw = await ps.runCommandJson<TraceMessage | TraceMessage[]>(cmd);
        } catch (e) {
          spin.stop("Message trace failed.");
          p.log.error(`${e instanceof Error ? e.message : e}`);
          break;
        }

        const messages = raw ? (Array.isArray(raw) ? raw : [raw]) : [];
        spin.stop(`Found ${messages.length} message(s) sent in last 10 days.`);

        findings.messageTraceChecked = true;
        findings.messageTrace = messages.map((m) => ({
          recipient: m.RecipientAddress ?? "",
          subject: m.Subject ?? "",
          status: m.Status ?? "",
          received: (m.Received ?? "").slice(0, 19),
        }));

        if (messages.length === 0) {
          p.log.info("No messages sent by this user in the last 10 days.");
          break;
        }

        // Display summary table (max 50 rows)
        const header = `${"Recipient".padEnd(32)} ${"Subject".padEnd(30)} ${"Status".padEnd(14)} Received`;
        const separator = "─".repeat(header.length + 5);
        const displayRows = messages.slice(0, 50);
        const lines = [header, separator];
        for (const m of displayRows) {
          lines.push(
            `${truncate(m.RecipientAddress ?? "", 31).padEnd(32)} ${truncate(m.Subject ?? "", 29).padEnd(30)} ${(m.Status ?? "").padEnd(14)} ${(m.Received ?? "").slice(0, 19)}`,
          );
        }
        if (messages.length > 50) lines.push(`… and ${messages.length - 50} more`);
        p.note(lines.join("\n"), `Message Trace (${messages.length} result(s))`);
        break;
      }

      // ── 8. Generate Ticket Notes ───────────────────────────────────
      case "ticket-notes": {
        const now = new Date();
        const dateStr = now.toISOString().slice(0, 16).replace("T", " ") + " UTC";

        const lines: string[] = [
          "=== INCIDENT RESPONSE REPORT ===",
          `Date: ${dateStr}`,
          `Tenant: ${ps.tenantDomain ?? "Unknown"}`,
          `Affected User: ${findings.user.displayName} (${findings.user.upn})`,
          "",
          "--- ACTIONS TAKEN ---",
        ];

        // Sessions
        if (findings.sessionsRevoked) {
          lines.push("[x] Signed out of all sessions");
        } else {
          lines.push("[ ] Sign out of all sessions: Not done");
        }

        // MFA
        if (findings.mfaMethodsRemoved.length > 0) {
          lines.push(`[x] MFA methods removed: ${findings.mfaMethodsRemoved.join(", ")}`);
          if (findings.mfaMethodsBefore.length > 0) {
            lines.push("    Previous MFA methods:");
            for (const m of findings.mfaMethodsBefore) {
              lines.push(`      - ${m.name}${m.detail ? `: ${m.detail}` : ""}`);
            }
          }
        } else if (findings.mfaMethodsBefore.length > 0) {
          lines.push(`[ ] MFA methods: ${findings.mfaMethodsBefore.length} found, none removed`);
          for (const m of findings.mfaMethodsBefore) {
            lines.push(`      - ${m.name}${m.detail ? `: ${m.detail}` : ""}`);
          }
        } else {
          lines.push("[ ] MFA methods: Not checked");
        }

        // Password
        if (findings.passwordReset) {
          lines.push("[x] Password reset (force change at next sign-in)");
          lines.push("    New credentials shared via one-time link");
        } else {
          lines.push("[ ] Password reset: Not done");
        }

        // Forwarding
        if (findings.forwardingChecked) {
          if (findings.forwardingFound && (findings.forwardingFound.smtp || findings.forwardingFound.address)) {
            const target = findings.forwardingFound.smtp ?? findings.forwardingFound.address ?? "";
            if (findings.forwardingRemoved) {
              lines.push(`[x] Forwarding: ${target} (REMOVED)`);
            } else {
              lines.push(`[!] Forwarding: ${target} (NOT REMOVED)`);
            }
          } else {
            lines.push("[x] Forwarding: None configured");
          }
        } else {
          lines.push("[ ] Forwarding: Not checked");
        }

        // Inbox rules
        if (findings.inboxRulesChecked) {
          if (findings.inboxRules.length === 0) {
            lines.push("[x] Inbox rules: None found");
          } else {
            const removedSet = new Set(findings.rulesRemoved);
            if (findings.rulesRemoved.length > 0) {
              lines.push(`[x] Inbox rules: ${findings.inboxRules.length} found, ${findings.rulesRemoved.length} removed`);
            } else {
              lines.push(`[x] Inbox rules: ${findings.inboxRules.length} found, none removed`);
            }
            for (const r of findings.inboxRules) {
              const wasRemoved = removedSet.has(r.name);
              const status = wasRemoved ? "REMOVED" : r.enabled ? "active" : "disabled";
              lines.push(`      - "${r.name}" (${status})`);
              const actions: string[] = [];
              if (r.forwardTo) actions.push(`Forward to: ${r.forwardTo}`);
              if (r.redirectTo) actions.push(`Redirect to: ${r.redirectTo}`);
              if (r.deleteMessage) actions.push("Delete message");
              if (r.moveToFolder) actions.push(`Move to: ${r.moveToFolder}`);
              if (actions.length > 0) lines.push(`        Actions: ${actions.join(", ")}`);
            }
          }
        } else {
          lines.push("[ ] Inbox rules: Not checked");
        }

        // Permissions
        if (findings.permissionsChecked) {
          if (findings.permissions.length === 0) {
            lines.push("[x] Mailbox permissions: No non-inherited permissions");
          } else {
            lines.push(`[x] Mailbox permissions: ${findings.permissions.length} found`);
            for (const perm of findings.permissions) {
              lines.push(`    ${perm.user} — ${perm.type}`);
            }
          }
        } else {
          lines.push("[ ] Mailbox permissions: Not checked");
        }

        // Message trace
        if (findings.messageTraceChecked) {
          lines.push(`[x] Message trace (last 10 days): ${findings.messageTrace.length} messages sent`);
          if (findings.messageTrace.length > 0) {
            // Top recipients
            const recipientCounts = new Map<string, number>();
            for (const m of findings.messageTrace) {
              recipientCounts.set(m.recipient, (recipientCounts.get(m.recipient) ?? 0) + 1);
            }
            const topRecipients = [...recipientCounts.entries()]
              .sort((a, b) => b[1] - a[1])
              .slice(0, 5)
              .map(([addr, count]) => `${addr} (${count})`)
              .join(", ");
            lines.push(`    Top recipients: ${topRecipients}`);
          }
        } else {
          lines.push("[ ] Message trace: Not checked");
        }

        lines.push("================================");

        const report = lines.join("\n");
        p.note(report, "Incident Response Report");

        const copied = await copyToClipboard(ps, report);
        if (copied) {
          p.log.success("Report copied to clipboard.");
        } else {
          p.log.info("Could not copy to clipboard — text displayed above.");
        }
        break;
      }
    }
  }
}
