# Testing Notes

## Philosophy

No mocked responses. Mocks don't represent real Exchange/Graph behavior and give false confidence. All testing should run against real Microsoft 365 infrastructure.

## Test Tenant

### Microsoft 365 Developer Program

- Free E5 tenant with 25 users, 90-day renewable
- Includes Exchange Online, Graph API, anti-spam policies, transport rules
- Sign up: https://developer.microsoft.com/en-us/microsoft-365/dev-program
- **Caveat:** Since Jan 2024, free sandbox provisioning is restricted. May need a Visual Studio Professional/Enterprise subscription to qualify.

### Setting Up Test Data

```powershell
Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All"

$passwordProfile = @{
    Password = "P@ssw0rd123!"
    ForceChangePasswordNextSignIn = $false
}

New-MgUser -DisplayName "Test User 1" `
    -MailNickname "testuser1" `
    -UserPrincipalName "testuser1@yourdevtenant.onmicrosoft.com" `
    -PasswordProfile $passwordProfile `
    -AccountEnabled

New-MgGroup -DisplayName "Test Security Group" `
    -MailEnabled:$false `
    -MailNickname "testsecgroup" `
    -SecurityEnabled
```

## Unattended Auth (for CI/CD)

Interactive browser login won't work in CI. Use certificate-based authentication:

### Exchange Online

```powershell
Connect-ExchangeOnline `
    -CertificateThumbprint "AABBCCDD..." `
    -AppId "your-app-id-guid" `
    -Organization "contoso.onmicrosoft.com" `
    -ShowBanner:$false
```

### Microsoft Graph

```powershell
Connect-MgGraph `
    -CertificateThumbprint "AABBCCDD..." `
    -ClientId "your-app-id-guid" `
    -TenantId "your-tenant-id"
```

### Setup Steps

1. Register an app in Entra ID (Azure AD)
2. Grant Exchange admin role + Graph API permissions
3. Generate a self-signed certificate and upload to the app registration
4. Store thumbprint, app ID, and tenant as environment variables or CI secrets

## Microsoft Graph Dev Proxy

Intercepts real HTTP calls at the system level — useful for simulating Graph API errors (throttling, 429s, 500s) without mocking response data.

```bash
brew tap microsoft/dev-proxy && brew install dev-proxy
```

### Error Simulation Config (`devproxyrc.json`)

```json
{
  "plugins": [
    {
      "name": "GraphRandomErrorPlugin",
      "enabled": true,
      "pluginPath": "~appFolder/plugins/DevProxy.Plugins.dll",
      "configSection": "graphRandomErrorPlugin"
    }
  ],
  "urlsToWatch": [
    "https://graph.microsoft.com/v1.0/*",
    "https://graph.microsoft.com/beta/*"
  ],
  "graphRandomErrorPlugin": {
    "rate": 50
  }
}
```

This simulates throttling and server errors at 50% rate to stress-test error handling paths.

**Limitation:** Exchange Online cmdlets use a different protocol than HTTP/REST, so Dev Proxy cannot intercept them. Only useful for Graph API calls.

## Pester (PowerShell-Native Testing)

If we ever extract reusable `.ps1` scripts, Pester is the standard PowerShell test framework.

```powershell
Install-Module Pester -Force -SkipPublisherCheck
Invoke-Pester tests/pester/ -Output Detailed
```

Syntax: `Describe` / `Context` / `It` / `Should` blocks. Useful for testing PowerShell logic in isolation.

## What to Actually Test

### Manual test checklist (against dev tenant)

- [ ] Connect-ExchangeOnline succeeds
- [ ] Connect-MgGraph succeeds (lazy, on first Graph command)
- [ ] Create user end-to-end (user appears in tenant)
- [ ] Edit user (display name, department, etc.)
- [ ] Delete user (removes from tenant + groups)
- [ ] Add/remove from distribution groups
- [ ] Add/remove from security groups
- [ ] Shared mailbox listing + permissions
- [ ] Whitelist/blacklist domain in anti-spam policy
- [ ] View/manage transport rules
- [ ] Inactive users report generates valid Excel
- [ ] OTS secret link creation works
- [ ] Cancel handling (Ctrl+C / Escape) doesn't crash

### Integration test ideas (automated, against dev tenant)

- `PowerShellSession` starts, runs a command, returns output
- JSON round-trip: PowerShell object → `ConvertTo-Json` → TypeScript parse
- Error capture: invalid command returns error string
- Sequential commands don't bleed output across markers
- `ensureGraphConnected()` connects only once
- Real `Get-MgUser` returns expected test user
- Real `Get-HostedContentFilterPolicy` returns Default policy

## Resources

- Microsoft 365 Developer Program: https://developer.microsoft.com/en-us/microsoft-365/dev-program
- App-only auth for Exchange Online: https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2
- Dev Proxy docs: https://learn.microsoft.com/en-us/microsoft-cloud/dev/dev-proxy/overview
- Pester framework: https://pester.dev/
- Bun test runner: https://bun.sh/docs/cli/test
