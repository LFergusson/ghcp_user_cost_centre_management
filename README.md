# GitHub Enterprise Cost Centre Membership

PowerShell scripts that assigns GitHub Enterprise Copilot users to cost centres
from a CSV, plus a companion script that generates that CSV from an Entra ID group.

The cost‑centre API logic mirrors the [`github/cost-center-automation`](https://github.com/github/cost-center-automation)
project, reimplemented in PowerShell for CSV‑driven assignment in environments with a preference for PowerShell.

## Contents

| File | Description |
|------|-------------|
| `Set-CostCentreMembership.ps1` | Reads a CSV and assigns users to cost centres in GitHub Enterprise. |
| `Export-EntraGroupToCsv.ps1` | Queries an Entra ID group and writes a compatible CSV (name, email, configurable cost‑centre source attribute — defaults to department). |
| `example.env` | Template for the `.env` file consumed by both scripts. Copy to `.env` and fill in your values. |
| `outputs/sample-assignments.csv` | Example input CSV. |
| `outputs/entra-group-export.csv` | Example output from `Export-EntraGroupToCsv.ps1`. |

## End-to-end workflow

```
Entra group ──Export-EntraGroupToCsv.ps1──▶ CSV ──Set-CostCentreMembership.ps1──▶ GitHub cost centres
```

1. **Generate** a CSV from an Entra group (department becomes the cost centre by
   default — pass `-CostCentreProperty` to use a different attribute):
   ```powershell
   .\Export-EntraGroupToCsv.ps1 -GroupName "Engineering" -OutputPath .\outputs\users.csv `
       -TenantId $env:AZURE_TENANT_ID -ClientId $env:AZURE_CLIENT_ID -ClientSecret $env:AZURE_CLIENT_SECRET
   ```
2. **Review / edit** the CSV (remap source values to cost-centre names if needed).
3. **Assign** the users to cost centres:
   ```powershell
   .\Set-CostCentreMembership.ps1 -CsvPath .\outputs\users.csv -Enterprise contoso -CreateMissingCostCentres
   ```

---

# Set-CostCentreMembership.ps1

Reads a CSV mapping **user email addresses** to **cost centre names** and assigns
each user to the named cost centre in a GitHub Enterprise (Cloud / GHE.com) account.

## How it works

1. **Load & validate the CSV.** Header matching is case‑insensitive. Blank rows
   are skipped.
2. **Resolve emails → GitHub usernames.** The GitHub cost‑centre API operates on
   GitHub usernames (logins), not email addresses. Each email is resolved by:
   1. Listing every assigned Copilot seat via
      `GET /enterprises/{enterprise}/copilot/billing/seats` to collect the
      canonical logins.
   2. Calling `GET /users/{login}` for each login and indexing the public
      profile `email` field.

   Both the login and the public email are indexed, so a CSV that already lists
   GitHub usernames still resolves. Only logins whose public profile email is
   set can be matched by email — for tenants where most users don't expose a
   public email, supply logins directly in the CSV and pass `-SkipUserLookup`.
3. **Load cost centres.** Active cost centres are read from
   `GET /enterprises/{enterprise}/settings/billing/cost-centers`. Missing ones are
   created (`POST`) when `-CreateMissingCostCentres` is supplied.
4. **Assign users.** Users are grouped by cost centre, anyone already a member is
   skipped, and the remainder are added in **batches of up to 50** (the API limit)
   via `POST /enterprises/{enterprise}/settings/billing/cost-centers/{id}/resource`.
5. **Summarise.** Counts of assigned / skipped / failed users are printed, along
   with any unresolved emails or unknown cost centres. The script exits with code
   `1` if any assignment failed or any email could not be resolved.

## Prerequisites

- **PowerShell 7.0+**
- **GitHub Enterprise Cloud** with admin access.
- A **GitHub token** with:
  - `manage_billing:enterprise` — required for all cost‑centre operations.
  - `manage_billing:copilot` (or `read:enterprise`) — required to list Copilot
    seat assignees for the email→login lookup (unless `-SkipUserLookup` is used).
  - `read:user` / `user:email` — required to read each user's public profile
    email via `GET /users/{login}` (unless `-SkipUserLookup` is used).

## CSV format

A header row is required. Default column names are `email` and `cost_centre`
(override with `-EmailColumn` / `-CostCentreColumn`).

```csv
email,cost_centre
alice@contoso.com,Engineering
bob@contoso.com,Marketing
carol@contoso.com,Engineering
```

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-CsvPath` *(required)* | Path to the CSV file. |
| `-Enterprise` | Enterprise slug. Defaults to `$env:GITHUB_ENTERPRISE`. |
| `-Token` | GitHub token. Defaults to `$env:GITHUB_TOKEN`. |
| `-ApiBaseUrl` | API base URL. Defaults to `$env:GITHUB_API_BASE_URL` or `https://api.github.com`. For GHE.com use `https://api.SUBDOMAIN.ghe.com`. |
| `-EmailColumn` | CSV email column name. Default: `email`. |
| `-CostCentreColumn` | CSV cost centre column name. Default: `cost_centre`. |
| `-CreateMissingCostCentres` | Create cost centres that don't already exist. Otherwise unknown cost centres are reported as failures. |
| `-SkipUserLookup` | Treat the email column as GitHub logins directly and skip the Copilot seats + `/users/{login}` lookup. Recommended when your CSV already contains GitHub usernames (e.g. for tenants where users don't expose a public profile email). |
| `-LogPath` | Optional path to a log file; console output is also written there. |
| `-WhatIf` / `-Confirm` | Standard PowerShell dry‑run / confirmation support. |

## Usage

```powershell
# Authenticate (token needs manage_billing:enterprise plus manage_billing:copilot
# and read:user / user:email for the email→login lookup)
$env:GITHUB_TOKEN = '...'
$env:GITHUB_ENTERPRISE = 'contoso'

# Preview changes without applying them
.\Set-CostCentreMembership.ps1 -CsvPath .\outputs\sample-assignments.csv -WhatIf

# Apply, creating any missing cost centres
.\Set-CostCentreMembership.ps1 -CsvPath .\outputs\sample-assignments.csv -CreateMissingCostCentres

# GHE.com data‑resident enterprise
.\Set-CostCentreMembership.ps1 -CsvPath .\assignments.csv -Enterprise contoso `
    -ApiBaseUrl 'https://api.contoso.ghe.com'

# CSV already contains GitHub usernames - skip the Copilot seats + /users lookup
.\Set-CostCentreMembership.ps1 -CsvPath .\handles.csv -SkipUserLookup
```

## Behaviour notes

- **Idempotent.** Users already in a target cost centre are skipped, so the script
  can be re‑run safely.
- **Batching.** A maximum of 50 users are sent per API request, matching the
  GitHub limit; larger groups are split automatically.
- **Resilience.** Requests retry with exponential backoff on rate limiting (429)
  and transient server errors (5xx).
- **Exit codes.** `0` on full success; `1` if any assignment failed or any email
  could not be resolved (useful for CI/automation).
- **Public profile emails.** The email→login map is only populated for logins
  whose GitHub public profile email is set. If most of your users hide that
  field, put GitHub logins in the CSV and pass `-SkipUserLookup` instead.
- **`.env` loading.** The script reads `.env` from the script directory
  (`$PSScriptRoot/.env`) when present. Precedence is:
  `explicit parameter > process environment variable > .env value > built-in default`.
  Quoted values are supported and blank/comment lines are ignored.

---

# Export-EntraGroupToCsv.ps1

Queries Microsoft Graph for the members of an Entra ID (Azure AD) group and writes
their display name, email and a configurable source attribute to a CSV. By default
the **department** is written to the cost‑centre column, producing a file that feeds
straight into `Set-CostCentreMembership.ps1`. Pass `-CostCentreProperty` to use any
other Entra user property (including `extensionAttributeN` and directory extensions)
as the cost‑centre source.

## How it works

1. **Authenticate** to Microsoft Graph — either with a supplied `-AccessToken` or
   via the app‑only client‑credentials flow (`-TenantId`, `-ClientId`,
   `-ClientSecret`).
2. **Resolve the group.** `-GroupId` is used directly; `-GroupName` is resolved via
   `GET /groups?$filter=displayName eq '...'` (aborts if the name is ambiguous).
3. **Fetch members.** Reads `members` (or `transitiveMembers` with `-Transitive`),
   cast to `microsoft.graph.user` so non‑user objects are excluded, selecting
   `displayName`, `mail`, `userPrincipalName` and the configured cost‑centre source
   property (default `department`). Pages through all results via `@odata.nextLink`.
4. **Write the CSV** with columns `name`, `email`, `cost_centre`. Email uses `mail`,
   falling back to `userPrincipalName`. The value of `-CostCentreProperty` (default
   `department`) is written to `cost_centre`.
   When `-GitHubUsernameProperty` is supplied, that Entra property's value (the
   user's GitHub username) is written to the email column instead, so the file can
   be consumed with `Set-CostCentreMembership.ps1 -SkipUserLookup` to look users up
   directly by their GitHub handle.

## Prerequisites

- **PowerShell 7.0+**
- An Entra **app registration** (for client‑credentials auth) with the Microsoft
  Graph **application** permissions and admin consent:
  - `GroupMember.Read.All` (or `Group.Read.All`)
  - `User.Read.All`
- Alternatively, a pre‑acquired Graph bearer token passed via `-AccessToken`.

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-GroupId` | Object id (GUID) of the group. Mutually exclusive with `-GroupName`. |
| `-GroupName` | Display name of the group; resolved to an id via Graph. |
| `-OutputPath` | CSV file to create. Default: `.\entra-group-export.csv`. |
| `-TenantId` | Entra tenant id/domain. Default: `$env:AZURE_TENANT_ID`. |
| `-ClientId` | App registration (client) id. Default: `$env:AZURE_CLIENT_ID`. |
| `-ClientSecret` | App client secret. Default: `$env:AZURE_CLIENT_SECRET`. |
| `-AccessToken` | Pre‑acquired Graph token; if set, client‑credential params are ignored. |
| `-GraphBaseUrl` | Graph base URL. Default: `https://graph.microsoft.com/v1.0`. Sovereign clouds supported. |
| `-Transitive` | Include nested‑group members (transitive membership). |
| `-GitHubUsernameProperty` | Entra/Graph user property holding the GitHub username. When set, its value is written to the email column instead of email/UPN. Accepts top‑level properties, `extensionAttributeN` short names (resolved under `onPremisesExtensionAttributes`), or dot‑paths / directory‑extension names. |
| `-EmailColumn` | Email header name. Default: `email`. |
| `-CostCentreColumn` | Cost‑centre header name. Default: `cost_centre`. |
| `-CostCentreProperty` | Entra/Graph user property whose value is written to the cost‑centre column. Default: `department`. Accepts the same forms as `-GitHubUsernameProperty` (top‑level properties, `extensionAttributeN` short names, dot‑paths / directory‑extension names). |
| `-DefaultCostCentre` | Value used when a member has no value for the cost‑centre source property. |
| `-LogPath` | Optional log file. |

## Usage

```powershell
# Client-credentials auth
.\Export-EntraGroupToCsv.ps1 -GroupName "Engineering" -OutputPath .\eng.csv `
    -TenantId $env:AZURE_TENANT_ID -ClientId $env:AZURE_CLIENT_ID -ClientSecret $env:AZURE_CLIENT_SECRET

# Pre-acquired token, by group id, including nested groups
.\Export-EntraGroupToCsv.ps1 -GroupId 00000000-0000-0000-0000-000000000000 `
    -AccessToken $token -Transitive -OutputPath .\users.csv

# Fill blank departments with a default cost centre
.\Export-EntraGroupToCsv.ps1 -GroupName "Contractors" -DefaultCostCentre "Unassigned" -OutputPath .\contractors.csv

# Use a custom Entra attribute (e.g. extensionAttribute2) as the cost-centre source
.\Export-EntraGroupToCsv.ps1 -GroupName "Engineering" -CostCentreProperty extensionAttribute2 -OutputPath .\eng.csv

# Export the GitHub username (e.g. stored in extensionAttribute1) into the email
# column, then assign cost centres by looking users up directly (no Copilot seat lookup)
.\Export-EntraGroupToCsv.ps1 -GroupName "Engineering" -GitHubUsernameProperty extensionAttribute1 -OutputPath .\handles.csv
.\Set-CostCentreMembership.ps1 -CsvPath .\handles.csv -Enterprise contoso -SkipUserLookup -WhatIf
```

## Behaviour notes

- **Output columns** are `name,email,cost_centre`. The `name` column is ignored by
  `Set-CostCentreMembership.ps1` but kept for human review.
- **Email fallback.** If a member has no `mail`, their `userPrincipalName` is used.
- **Cost‑centre source.** Defaults to the user's `department`. Override with
  `-CostCentreProperty` to use any other Entra user property, including
  `extensionAttributeN` (resolved under `onPremisesExtensionAttributes`) or a
  dot‑path / directory‑extension name.
- **GitHub username mode.** With `-GitHubUsernameProperty`, the named Entra property
  is written to the email column instead of email/UPN. Members lacking a value for
  that property are skipped. Feed the result to `Set-CostCentreMembership.ps1` with
  `-SkipUserLookup`.
- **Missing department.** Logged as a warning; the row is written with a blank cost
  centre unless `-DefaultCostCentre` is supplied. Blank cost‑centre rows are skipped
  by the downstream script, so edit the CSV or set a default first.
- **User‑only.** Nested groups, devices and service principals are filtered out.
- **Resilience.** Graph requests retry with backoff on 429/5xx, honouring the
  `Retry-After` header.
