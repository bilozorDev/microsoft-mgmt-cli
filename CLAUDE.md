# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

- `bun install` — install dependencies
- `bun run start` — run the CLI (`bun run src/index.ts`)

No test framework or build step configured — Bun runs TypeScript directly.

## Tech Stack

- **Runtime:** Bun (TypeScript, strict mode, `noEmit`)
- **Interactive UI:** `@clack/prompts`
- **Exchange Online:** PowerShell subprocess (`ExchangeOnlineManagement` module)
- **Microsoft Graph:** PowerShell `Microsoft.Graph` module (lazily connected)

## Architecture

### PowerShell session (`src/powershell.ts`)

Persistent `pwsh` process spawned via `Bun.spawn`. A custom loop script reads stdin line-by-line using `[Console]::In.ReadLine()` (PowerShell's `-Command -` mode buffers all stdin until EOF, so we can't use it). Commands are accumulated in a `StringBuilder` until an exec marker, then run via `Invoke-Expression`. Output uses end/error markers written through `[Console]::Out` to bypass pipeline buffering.

Key methods:
- `runCommand(cmd)` → `{ output, error }` — raw string output
- `runCommandJson<T>(cmd)` — appends `| ConvertTo-Json -Depth 5 -Compress` and parses
- `ensureGraphConnected()` — lazy Graph connection (only when a command needs it)
- `connectExchangeOnline()` — called at startup

### Startup flow (`src/index.ts`)

1. `checkRequirements()` — verifies `pwsh`, ExchangeOnlineManagement, Microsoft.Graph are installed; offers auto-install
2. Start persistent PowerShell session
3. `Connect-ExchangeOnline` (opens browser for auth)
4. Menu loop dispatching to commands

### Commands (`src/commands/`)

Each file exports `async function run(ps: PowerShellSession): Promise<void | string[]>`. To add a new command:

1. Create `src/commands/<name>.ts` exporting `run(ps)`
2. Add option to the `select()` menu in `src/index.ts`
3. Use `ps.runCommand()` / `ps.runCommandJson()` for PowerShell execution

### When to use PowerShell vs Graph API

**PowerShell required** (Graph doesn't support these):
- Anti-spam policies (`*-HostedContentFilterPolicy`)
- Mail flow / transport rules (`*-TransportRule`)
- Connection filter policies (`*-HostedConnectionFilterPolicy`)
- Outbound spam policies

**Graph API** (via `ensureGraphConnected()`):
- User/group management, license assignment
- Distribution list and security group membership
- Mailbox settings, calendar operations

### Auth

`Connect-ExchangeOnline` uses interactive browser login. Do NOT use `-UseDeviceAuthentication`. Graph connects lazily via `Connect-MgGraph` with specific scopes when first needed.

## Patterns

- **Single quotes in PowerShell:** escape with `value.replace(/'/g, "''")`
- **Array normalization:** Graph may return a single object or array — use `Array.isArray(raw) ? raw : [raw]`
- **Cancel handling:** all `@clack/prompts` calls must check `p.isCancel()` and bail
- **Spinners:** wrap long operations with `p.spinner().start()` / `.stop()`
- **Error flow:** check `error` property from `runCommand()`; display with `p.log.error()` but generally continue
