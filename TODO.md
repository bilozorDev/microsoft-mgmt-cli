# TODO — New Reports

## High Value, Easy to Build

- [x] **License Usage Report** — All SKUs with total/consumed/available seats, cost per license, users assigned to each. Helps identify wasted spend. (`Get-MgSubscribedSku` + `Get-MgUser` with license details)

- [ ] **Distribution Group Audit** — All distribution groups with member counts, owners, email addresses. Flags empty groups and groups with only one member. (`Get-DistributionGroup` + `Get-DistributionGroupMember`)

- [x] **Mailbox Forwarding Audit** — Users with forwarding rules, SMTP forwarding, or inbox rules that forward externally. Major security/compliance concern. (`Get-Mailbox -Properties ForwardingSmtpAddress,ForwardingAddress,DeliverToMailboxAndForward` + `Get-InboxRule`)

- [ ] **Guest User Audit** — All external/guest users (#EXT#), when created, last sign-in, which groups they belong to. Common compliance requirement. (Graph `Get-MgUser -Filter "userType eq 'Guest'"` + sign-in activity)

- [x] **Admin Role Report** — Users with admin roles (Global Admin, Exchange Admin, etc.) and whether they have MFA enabled. Security audit essential. (`Get-MgDirectoryRole` + `Get-MgDirectoryRoleMember`)

## Medium Effort, Very Useful

- [ ] **Mailbox Size / Quota Report** — Mailbox sizes, quota status, items in deleted items, archive mailbox status. (`Get-MailboxStatistics` + `Get-Mailbox` quota properties)

- [ ] **Security Group Audit** — All security groups with member counts, type (M365 vs security vs mail-enabled), and owners. (`Get-MgGroup` with filtering)

- [ ] **Transport Rules Report** — All mail flow rules with priority, conditions, actions, enabled/disabled status. Important for compliance reviews. (`Get-TransportRule`)

- [ ] **Anti-Spam Policy Report** — Current spam filter settings, allowed/blocked senders and domains, quarantine policies. Snapshot of tenant security posture. (`Get-HostedContentFilterPolicy` + `Get-HostedConnectionFilterPolicy`)

## Stretch Goals

- [x] **Mailbox Permissions Matrix** — Cross-reference of who has access to which shared mailboxes (full access + send-as).

- [ ] **Password & MFA Status Report** — Users with/without MFA, password age, password never expires flag. Requires `AuthenticationMethod.Read.All` scope addition.

## UX Improvements

- [x] **Clipboard Copy for Ticket Notes** — After actions (user created, deleted, license changed, etc.), auto-copy a summary note to clipboard so it can be pasted into the ticketing system. Use `pbcopy` on macOS, `clip` on Windows.
