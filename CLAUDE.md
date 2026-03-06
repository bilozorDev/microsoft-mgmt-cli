# Profulgent — Exchange Online Admin CLI

## Tech Stack

- **Runtime:** Bun (TypeScript)
- **Interactive UI:** `@clack/prompts`
- **Exchange Online:** PowerShell subprocess (`ExchangeOnlineManagement` module)

## Architecture

Persistent PowerShell session (`src/powershell.ts`) spawned via `Bun.spawn`. A command loop script using `[Console]::In.ReadLine()` processes commands one at a time (PowerShell's `-Command -` mode buffers all stdin until EOF, so we can't use it for a persistent session).

### When to use PowerShell (ExchangeOnlineManagement module)

Microsoft Graph API does **not** support organization-wide mail flow rules or anti-spam policies. PowerShell is **required** for:

- Anti-spam policy management (`Set-HostedContentFilterPolicy`, `Get-HostedContentFilterPolicy`)
- Mail flow / transport rules (`New-TransportRule`, `Set-TransportRule`)
- Connection filter policies (`Set-HostedConnectionFilterPolicy`)
- Outbound spam policies
- Any Exchange Admin Center operation under **Mail flow** or **Anti-spam**

### When to use Microsoft Graph API

Graph API should be used for operations it supports:

- Mailbox settings (signatures, auto-replies, forwarding)
- Mail folder management
- Calendar operations
- User/group management
- Distribution list membership
- Message tracking (limited)

### Auth

`Connect-ExchangeOnline` handles authentication via interactive browser login. The subprocess can open a browser directly — do NOT use `-UseDeviceAuthentication` (not available in all module versions).

## Adding Commands

1. Create `src/commands/<name>.ts` exporting `async run(ps: PowerShellSession)`
2. Add option to the `select()` menu in `src/index.ts`
3. Use `ps.runCommand()` to execute PowerShell commands
