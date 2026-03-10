# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

- `bun install` — install dependencies
- `bun run start` — run the CLI (`bun run src/index.ts`)
- `bun run build:windows` — compile Windows exe to `dist/Microsoft 365 Admin CLI/m365-admin.exe`

No test framework or build step configured — Bun runs TypeScript directly.

## Tech Stack

- **Runtime:** Bun (TypeScript, strict mode, `noEmit`)
- **Interactive UI:** `@clack/prompts`
- **Exchange Online:** PowerShell subprocess (`ExchangeOnlineManagement` module)
- **Microsoft Graph:** PowerShell `Microsoft.Graph` module (lazily connected)
- **Excel reports:** `exceljs` via reusable template (`src/report-template.ts`)

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
2. `checkForUpdates()` — auto-update from GitHub Releases (compiled `.exe` only; skipped in dev)
3. Start persistent PowerShell session
4. `Connect-ExchangeOnline` (opens browser for auth)
5. Menu loop with category submenus (User Management, Spam Management, Reports)

### Commands (`src/commands/`)

Each file exports `async function run(ps: PowerShellSession): Promise<void | string[]>`. To add a new command:

1. Create `src/commands/<name>.ts` exporting `run(ps)`
2. Import in `src/index.ts`, add option to the relevant submenu `select()`, add `case` in the switch
3. Use `ps.runCommand()` / `ps.runCommandJson()` for PowerShell execution

### Excel report template (`src/report-template.ts`)

All Excel reports use `generateReport(opts)` which produces a branded workbook buffer. To add a new report with Excel export:

```ts
import { generateReport } from "../report-template.ts";

const buffer = await generateReport({
  sheetName: "My Report",
  title: "My Report Title",
  tenant: ps.tenantDomain ?? "Unknown",
  summary: "10 items found",
  columns: [
    { header: "Name", width: 30 },
    { header: "Email", width: 38 },
  ],
  rows: data.map((d) => [d.name, d.email]),
});
await Bun.write(fullPath, buffer);
```

The template handles: title/tenant/date/summary header rows starting at column A, blue divider, styled table headers (#2B5797 background), and alternating row fills (#E8EDF2). To change styling, edit `src/report-template.ts` once — all reports inherit the change.

### When to use PowerShell vs Graph API

**PowerShell required** (Graph doesn't support these):
- Anti-spam policies (`*-HostedContentFilterPolicy`)
- Mail flow / transport rules (`*-TransportRule`)
- Connection filter policies (`*-HostedConnectionFilterPolicy`)
- Outbound spam policies
- Shared mailbox management (`Get-Mailbox`, `Get-MailboxPermission`, `Get-RecipientPermission`)

**Graph API** (via `ensureGraphConnected()`):
- User/group management, license assignment
- Distribution list and security group membership
- Mailbox settings, calendar operations

### Auth

`Connect-ExchangeOnline` uses interactive browser login. Do NOT use `-UseDeviceAuthentication`. Graph connects lazily via `Connect-MgGraph` with specific scopes when first needed.

### One-time secret links (`src/utils.ts`)

`createSecretLink(secret, ttl)` calls the onetimesecret.com API (anonymous, no auth) to create a self-destructing link. Used after user creation to share credentials securely. Default TTL is 7 days.

### Auto-update (`src/auto-update.ts`)

Checks GitHub Releases for `bilozorDev/microsoft-mgmt-cli` on startup. Only runs when `process.execPath` ends with `.exe` (compiled binary). Downloads the `.zip` asset, extracts via `pwsh Expand-Archive`, renames the running exe to `.old` (Windows allows rename but not overwrite of a running exe), copies the new `m365-admin.exe`, then exits for restart. Network/API errors are caught silently so the app always starts.

### Versioning and releases

The `version` field in `package.json` is the source of truth — Bun inlines it at compile time via `import pkg from "../package.json"`. To ship a new release:

1. Bump `version` in `package.json`
2. Commit and push the version bump
3. `bun run build:windows` — produces `dist/Microsoft 365 Admin CLI.zip`
4. `gh release create v<version> --title "v<version>" --notes "..."` — create the GitHub release
5. `gh release upload v<version> "dist/Microsoft 365 Admin CLI.zip"` — attach the binary

All five steps are required — the auto-updater downloads the zip asset from the GitHub release, so a release without the binary is not usable. The auto-updater compares `tag_name` (stripped of leading `v`) against the compiled-in version.

### Helper commands (`src/commands/add-to-*.ts`)

Reusable membership helpers called from create-user, edit-user, and delete-user. They accept an optional `upn` parameter — when provided, they skip the user-selection prompt and operate on that UPN directly.

## Cross-Platform (Windows / macOS)

The app runs on macOS via `bun run start` and ships as a Windows exe via `bun run build:windows`. Key platform differences to handle:

- **Clipboard:** use `Set-Clipboard` via the PowerShell session on Windows, `pbcopy` via `Bun.spawn` on macOS
- **Open folder:** `explorer` on Windows, `open` on macOS — always wrap in `try/catch`
- **File paths:** use Node's `path.join()`/`path.resolve()` (platform-aware); never hardcode `/` or `\`
- **Asset embedding:** `import.meta.dir` and `import.meta.url` don't resolve to filesystem paths in compiled Bun binaries. Use `fileURLToPath(new URL("...", import.meta.url))` and read assets into buffers at module load time. (Currently no bundled assets.)
- **Report saving:** always call `mkdirSync(dirname(fullPath), { recursive: true })` before `Bun.write()` to ensure parent directories exist
- **Signals:** `SIGTERM` is not supported on Windows; `SIGINT` (Ctrl+C) works on both
- **PowerShell:** requires `pwsh` (PowerShell Core 7+), not the built-in Windows PowerShell 5.1 (`powershell.exe`)

## Patterns

- **Single quotes in PowerShell:** escape with `value.replace(/'/g, "''")`
- **Array normalization:** Graph may return a single object or array — use `Array.isArray(raw) ? raw : [raw]`
- **Cancel handling:** all `@clack/prompts` calls must check `p.isCancel()` and bail
- **Spinners:** wrap long operations with `p.spinner().start()` / `.stop()`
- **Error flow:** check `error` property from `runCommand()`; display with `p.log.error()` but generally continue
