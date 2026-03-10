# Microsoft 365 Admin CLI

An interactive command-line tool for managing Microsoft 365 tenants. Built with Bun and PowerShell, it provides a friendly terminal UI for common admin tasks — user management, spam policies, and reporting — with Excel export.

## Features

- **User Management** — Create, edit, and delete users with license assignment, group membership, and shared mailbox access
- **Spam Management** — Whitelist domains across anti-spam policies
- **Reports** — Inactive users and shared mailboxes with Excel export
- **Multi-tenant** — Switch between tenants without restarting
- **Auto-update** — Self-updating Windows binary via GitHub Releases
- **Cross-platform** — Runs on macOS (dev) and Windows (compiled exe)

## Prerequisites

- [Bun](https://bun.sh) (for development) or the compiled Windows exe
- [PowerShell Core 7+](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell) (`pwsh`)
- PowerShell modules (auto-installed on first run if missing):
  - `ExchangeOnlineManagement`
  - `Microsoft.Graph`

## Getting Started

```bash
# Install dependencies
bun install

# Run the CLI
bun run start
```

On first launch the CLI checks for `pwsh` and required modules, offering to install anything missing. It then opens your browser for Microsoft 365 authentication via `Connect-ExchangeOnline`.

## Building for Windows

```bash
bun run build:windows
```

Produces `dist/Microsoft 365 Admin CLI/m365-admin.exe` and a zip archive for distribution.

## Usage

The CLI presents an interactive menu:

```
◆  What would you like to do?
│  User Management
│  Spam Management
│  Reports
│  Switch tenant
│  Exit
```

### User Management

- **Create user** — Set display name, username, domain, password, and optionally assign licenses, add to distribution/security groups, and grant shared mailbox access. Generates a one-time secret link for sharing credentials.
- **Edit user** — Modify display name, username, licenses, and group/mailbox memberships for an existing user.
- **Delete user** — Remove a user and optionally revoke their group and mailbox access first.

### Spam Management

- **Whitelist domain(s)** — Add sender domains to the allowed list across all anti-spam policies.

### Reports

- **Inactive users** — Lists users who haven't signed in within a configurable number of days. Exports to Excel.
- **Shared mailboxes** — Lists all shared mailboxes with their members (Full Access, Send As, Send on Behalf). Exports to Excel.

Reports are saved to a `reports output/` directory and can be opened directly from the CLI.

## License

Private — not published to npm.
