#Requires -Version 7.0

<#
.SYNOPSIS
    Assigns GitHub Enterprise Copilot users to cost centres from a CSV file.

.DESCRIPTION
    Reads a CSV that maps user email addresses to cost centre names, then iterates
    through the file and assigns each user to the named cost centre in a GitHub
    Enterprise (Cloud / GHE.com) account.

    Because the GitHub cost-centre API operates on GitHub usernames (logins) rather
    than email addresses, each email is resolved to its GitHub login by pairing two
    APIs:
      - /enterprises/{enterprise}/copilot/billing/seats lists every assigned
        Copilot seat (and therefore the canonical login for each user we care about).
      - GET /users/{login} returns the public profile, whose `email` field is used
        as the lookup key.

    Only logins whose public profile email is populated can be matched by email.
    For environments where most users do not expose a public email supply logins directly in the CSV and pass -SkipUserLookup.

    Workflow:
      1. Validate inputs and load the CSV.
      2. List Copilot seat logins, then call /users/{login} for each to build an
         email -> login map.
      3. Load existing cost centres (name -> id). Optionally create missing ones.
      4. Group CSV rows by cost centre, resolve each email to a login, skip users
         already assigned, then add the remainder in batches of up to 50 (API limit).
      5. Print a summary of successes, skips and failures.

    Supports -WhatIf / -Confirm for safe dry runs.

.PARAMETER CsvPath
    Path to the CSV file. Must contain an email column and a cost centre column
    (see -EmailColumn and -CostCentreColumn). Header matching is case-insensitive.

.PARAMETER Enterprise
    The enterprise slug. Defaults to the GITHUB_ENTERPRISE environment variable.

.PARAMETER Token
    A GitHub token. Defaults to the GITHUB_TOKEN environment variable.
    Requires 'manage_billing:enterprise' (cost centres) and 'manage_billing:copilot'
    plus 'read:user'/'user:email' scopes for the seat + user lookups.

.PARAMETER ApiBaseUrl
    API base URL. Defaults to GITHUB_API_BASE_URL or https://api.github.com.
    For GHE.com data-resident enterprises use https://api.SUBDOMAIN.ghe.com.

.PARAMETER EmailColumn
    Name of the CSV column holding the email address. Default: 'email'.

.PARAMETER CostCentreColumn
    Name of the CSV column holding the cost centre name. Default: 'cost_centre'.

.PARAMETER CreateMissingCostCentres
    Create cost centres that do not already exist. Without this switch, rows that
    reference an unknown cost centre are reported as failures.

.PARAMETER SkipUserLookup
    Treat the value in the email column as the GitHub login directly and skip the
    Copilot seats + /users lookup entirely. Useful if your CSV already contains
    logins (recommended for EMU tenants where public profile emails are rarely set).

.PARAMETER LogPath
    Optional path to a log file. All console output is also written here.

.EXAMPLE
    .\Set-CostCentreMembership.ps1 -CsvPath .\assignments.csv -Enterprise contoso -WhatIf

.EXAMPLE
    $env:GITHUB_TOKEN = '...'
    .\Set-CostCentreMembership.ps1 -CsvPath .\assignments.csv -Enterprise contoso -CreateMissingCostCentres

.NOTES
    CSV format (header row required):
        email,cost_centre
        alice@contoso.com,Engineering
        bob@contoso.com,Marketing
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ if (Test-Path -LiteralPath $_ -PathType Leaf) { $true } else { throw "CSV file not found: $_" } })]
    [string]$CsvPath,

    [Parameter()]
    [string]$Enterprise,

    [Parameter()]
    [string]$Token,

    [Parameter()]
    [string]$ApiBaseUrl,

    [Parameter()]
    [string]$EmailColumn = 'email',

    [Parameter()]
    [string]$CostCentreColumn = 'cost_centre',

    [Parameter()]
    [switch]$CreateMissingCostCentres,

    [Parameter()]
    [switch]$SkipUserLookup,

    [Parameter()]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# GitHub cost-centre API accepts a maximum of 50 users per request.
$script:MaxUsersPerRequest = 50
$script:ApiVersion = '2022-11-28'

#region Helpers ----------------------------------------------------------------

function Get-DotEnvValues {
    <#
        Reads KEY=VALUE pairs from a .env file and returns a hashtable.
        Supports quoted values, ignores blank lines and full-line comments.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $values = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $values
    }

    $lineNumber = 0
    foreach ($rawLine in (Get-Content -LiteralPath $Path)) {
        $lineNumber++
        $line = "$rawLine".Trim()

        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) {
            continue
        }

        $eqIndex = $line.IndexOf('=')
        if ($eqIndex -lt 1) {
            Write-Warning ".env line $lineNumber in '$Path' is invalid and will be ignored."
            continue
        }

        $name = $line.Substring(0, $eqIndex).Trim()
        $value = $line.Substring($eqIndex + 1).Trim()

        if ($name.StartsWith('export ')) {
            $name = $name.Substring(7).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($name)) {
            Write-Warning ".env line $lineNumber in '$Path' has an empty key and will be ignored."
            continue
        }

        if (
            ($value.StartsWith('"') -and $value.EndsWith('"') -and $value.Length -ge 2) -or
            ($value.StartsWith("'") -and $value.EndsWith("'") -and $value.Length -ge 2)
        ) {
            $value = $value.Substring(1, $value.Length - 2)
        }
        elseif ($value -match '^([^#]*?)\s+#') {
            # Support inline comments for unquoted values: KEY=value # comment
            $value = $Matches[1].TrimEnd()
        }

        $values[$name] = $value
    }

    return $values
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')][string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$Level] $Message"

    switch ($Level) {
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'DEBUG'   { Write-Verbose $line }
        default   { Write-Host $line }
    }

    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line
    }
}

function Invoke-GitHubApi {
    <#
        Wrapper around Invoke-RestMethod that adds the appropriate headers,
        retries on rate limiting (429) and transient 5xx errors, and returns a
        result object: @{ Success; StatusCode; Content }.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [ValidateSet('GET', 'POST', 'DELETE', 'PATCH')][string]$Method = 'GET',
        [object]$Body,
        [int]$MaxRetries = 3
    )

    $headers = @{
        'Authorization'        = "Bearer $script:Token"
        'X-GitHub-Api-Version' = $script:ApiVersion
        'User-Agent'           = 'ghcp-costcentre-mgmt'
        'Accept'               = 'application/vnd.github+json'
    }

    $params = @{
        Uri                = $Uri
        Method             = $Method
        Headers            = $headers
        SkipHttpErrorCheck = $true
        UseBasicParsing    = $true
    }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $params['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
        $params['ContentType'] = 'application/json'
    }

    $attempt = 0
    while ($true) {
        $attempt++
        $statusCode = 0
        $responseHeaders = @{}
        try {
            $response = Invoke-WebRequest @params
            $statusCode = [int]$response.StatusCode
            $responseHeaders = $response.Headers
            $rawContent = $response.Content
            if ([string]::IsNullOrEmpty($rawContent)) {
                $content = $null
            }
            else {
                try { $content = $rawContent | ConvertFrom-Json } catch { $content = $rawContent }
            }
        }
        catch {
            # Network-level failure (DNS, TLS, connection reset, etc.)
            if ($attempt -le $MaxRetries) {
                $wait = [math]::Pow(2, $attempt)
                Write-Log "Request to $Uri failed ($($_.Exception.Message)). Retrying in $wait s..." 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
            return @{ Success = $false; StatusCode = 0; Content = $_.Exception.Message; Headers = @{} }
        }

        # Rate limited - honour Retry-After / X-RateLimit-Reset where available, else fall back to 60s.
        if ($statusCode -eq 429 -or ($statusCode -eq 403 -and "$content" -match 'rate limit')) {
            if ($attempt -le $MaxRetries) {
                $wait = 60
                $retryAfter = Get-HeaderValue -Headers $responseHeaders -Name 'Retry-After'
                if ($retryAfter -and ($retryAfter -as [int])) {
                    $wait = [int]$retryAfter
                }
                else {
                    $reset = Get-HeaderValue -Headers $responseHeaders -Name 'X-RateLimit-Reset'
                    if ($reset -and ($reset -as [long])) {
                        $delta = [long]$reset - [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
                        if ($delta -gt 0) { $wait = [int][math]::Min($delta, 600) }
                    }
                }
                if ($wait -lt 1) { $wait = 1 }
                Write-Log "Rate limited on $Uri. Waiting $wait s (attempt $attempt/$MaxRetries)..." 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
        }

        # Transient server errors - back off and retry.
        if ($statusCode -ge 500 -and $statusCode -lt 600 -and $attempt -le $MaxRetries) {
            $wait = [math]::Pow(2, $attempt)
            Write-Log "Server error $statusCode on $Uri. Retrying in $wait s..." 'WARN'
            Start-Sleep -Seconds $wait
            continue
        }

        $success = ($statusCode -ge 200 -and $statusCode -lt 300)
        return @{ Success = $success; StatusCode = $statusCode; Content = $content; Headers = $responseHeaders }
    }
}

#endregion Helpers -------------------------------------------------------------

#region User lookup -----------------------------------------------------------

function Get-HeaderValue {
    <# Returns the first value of a header by name, tolerating dictionary/object shapes. #>
    param($Headers, [Parameter(Mandatory)][string]$Name)
    if (-not $Headers) { return $null }
    try { $raw = $Headers[$Name] } catch { $raw = $null }
    if (-not $raw -and ($Headers.PSObject.Properties[$Name])) {
        $raw = $Headers.PSObject.Properties[$Name].Value
    }
    if (-not $raw) { return $null }
    return (@($raw))[0]
}

function Get-NextLinkFromHeaders {
    <# Parses an RFC 5988 Link header collection and returns the rel="next" URL, or $null. #>
    param($Headers)
    $link = Get-HeaderValue -Headers $Headers -Name 'Link'
    if (-not $link) { return $null }
    foreach ($part in $link -split ',') {
        if ($part -match '<([^>]+)>\s*;\s*rel="next"') { return $Matches[1] }
    }
    return $null
}

function Get-CopilotSeatLogins {
    <#
        Pages through /enterprises/{enterprise}/copilot/billing/seats and returns
        a string[] of assignee logins (the real GitHub logins as the cost-centre
        API expects them).
    #>
    param([Parameter(Mandatory = $true)][string]$Enterprise)

    $logins = [System.Collections.Generic.List[string]]::new()
    $perPage = 100
    $uri = "$script:ApiBaseUrl/enterprises/$Enterprise/copilot/billing/seats?per_page=$perPage"
    $pageCount = 0

    Write-Log "Listing Copilot seat assignments from /enterprises/$Enterprise/copilot/billing/seats..."

    while ($uri) {
        $result = Invoke-GitHubApi -Uri $uri
        if (-not $result.Success) {
            throw "Failed to list Copilot seats (HTTP $($result.StatusCode)): $($result.Content)"
        }

        $seats = @()
        if ($result.Content -and $result.Content.PSObject.Properties.Name -contains 'seats' -and $result.Content.seats) {
            $seats = @($result.Content.seats)
        }
        $pageCount++

        foreach ($seat in $seats) {
            if ($seat.assignee -and $seat.assignee.login -and $seat.assignee.type -eq 'User') {
                $logins.Add($seat.assignee.login)
            }
        }

        $uri = Get-NextLinkFromHeaders -Headers $result.Headers
    }

    $unique = $logins | Sort-Object -Unique
    Write-Log "Found $($unique.Count) Copilot seat login(s) across $pageCount page(s)." 'INFO'
    return @($unique)
}

function Get-UserEmail {
    <#
        Fetches the public profile email for a given GitHub login via
        GET /users/{username}. Returns the email string, or $null.
    #>
    param([Parameter(Mandatory = $true)][string]$Login)

    $uri = "$script:ApiBaseUrl/users/$Login"
    $result = Invoke-GitHubApi -Uri $uri

    if (-not $result.Success) {
        Write-Log "  Could not read user '$Login' (HTTP $($result.StatusCode))." 'DEBUG'
        return $null
    }

    if ($result.Content -and $result.Content.PSObject.Properties.Name -contains 'email' -and $result.Content.email) {
        return $result.Content.email
    }
    return $null
}

function Get-UserEmailMap {
    <#
        Builds the email -> login map by:
          1. Listing all Copilot seat assignees (real GitHub logins).
          2. For each login, calling GET /users/{login} and reading the `email`
             field from the user object.
        Indexes the email (lower-cased) and the login itself.
    #>
    param([Parameter(Mandatory = $true)][string]$Enterprise)

    $map = @{}
    $logins = Get-CopilotSeatLogins -Enterprise $Enterprise

    Write-Log "Resolving emails for $($logins.Count) login(s) via /users/{login}..."
    $withEmail = 0

    foreach ($login in $logins) {
        $map[$login.ToLower()] = $login
        $email = Get-UserEmail -Login $login
        if ($email) {
            $withEmail++
            $map[$email.ToLower()] = $login
        }
    }

    Write-Log "Resolved emails for $withEmail/$($logins.Count) login(s); $($map.Count) lookup entries." 'INFO'
    return $map
}

#endregion User lookup --------------------------------------------------------

#region Cost centres -----------------------------------------------------------

function Get-CostCentreMap {
    <# Returns a hashtable mapping active cost-centre name -> id. #>
    param([Parameter(Mandatory = $true)][string]$Enterprise)

    $uri = "$script:ApiBaseUrl/enterprises/$Enterprise/settings/billing/cost-centers"
    $result = Invoke-GitHubApi -Uri $uri

    if (-not $result.Success) {
        throw "Failed to list cost centres (HTTP $($result.StatusCode)): $($result.Content)"
    }

    $map = @{}
    $costCenters = @()
    if ($result.Content.PSObject.Properties.Name -contains 'costCenters' -and $result.Content.costCenters) {
        $costCenters = @($result.Content.costCenters)
    }

    foreach ($cc in $costCenters) {
        $state = if ($cc.PSObject.Properties.Name -contains 'state') { "$($cc.state)" } else { 'active' }
        if ($state.ToLower() -eq 'active' -and $cc.name -and $cc.id) {
            $map[$cc.name] = $cc.id
        }
    }

    Write-Log "Loaded $($map.Count) active cost centre(s)." 'INFO'
    return $map
}

function New-CostCentre {
    <# Creates a cost centre and returns its id, or $null on failure. #>
    param(
        [Parameter(Mandatory = $true)][string]$Enterprise,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $uri = "$script:ApiBaseUrl/enterprises/$Enterprise/settings/billing/cost-centers"
    $result = Invoke-GitHubApi -Uri $uri -Method 'POST' -Body @{ name = $Name }

    if ($result.Success -and $result.Content.id) {
        Write-Log "Created cost centre '$Name' (id: $($result.Content.id))." 'SUCCESS'
        return $result.Content.id
    }

    # 409: already exists - try to recover the existing id.
    if ($result.StatusCode -eq 409) {
        $existing = Get-CostCentreMap -Enterprise $Enterprise
        if ($existing.ContainsKey($Name)) {
            Write-Log "Cost centre '$Name' already exists (id: $($existing[$Name]))." 'INFO'
            return $existing[$Name]
        }
    }

    Write-Log "Failed to create cost centre '$Name' (HTTP $($result.StatusCode)): $($result.Content)" 'ERROR'
    return $null
}

function Get-CostCentreMembers {
    <# Returns a string[] of usernames currently assigned to a cost centre. #>
    param(
        [Parameter(Mandatory = $true)][string]$Enterprise,
        [Parameter(Mandatory = $true)][string]$CostCentreId
    )

    $uri = "$script:ApiBaseUrl/enterprises/$Enterprise/settings/billing/cost-centers/$CostCentreId"
    $result = Invoke-GitHubApi -Uri $uri

    if (-not $result.Success) {
        Write-Log "Could not read members of cost centre $CostCentreId (HTTP $($result.StatusCode)). Proceeding without skip list." 'WARN'
        return @()
    }

    $members = [System.Collections.Generic.List[string]]::new()
    if ($result.Content.PSObject.Properties.Name -contains 'resources' -and $result.Content.resources) {
        foreach ($resource in $result.Content.resources) {
            if ($resource.type -eq 'User' -and $resource.name) {
                $members.Add($resource.name)
            }
        }
    }
    return $members.ToArray()
}

function Add-UsersToCostCentre {
    <# Adds up to 50 users to a cost centre. Returns the API result object. #>
    param(
        [Parameter(Mandatory = $true)][string]$Enterprise,
        [Parameter(Mandatory = $true)][string]$CostCentreId,
        [Parameter(Mandatory = $true)][string[]]$Usernames
    )

    $uri = "$script:ApiBaseUrl/enterprises/$Enterprise/settings/billing/cost-centers/$CostCentreId/resource"
    return Invoke-GitHubApi -Uri $uri -Method 'POST' -Body @{ users = $Usernames }
}

#endregion Cost centres --------------------------------------------------------

#region Main -------------------------------------------------------------------

# Initialise log file.
$script:LogFile = $null
if ($LogPath) {
    $script:LogFile = $LogPath
    $logDir = Split-Path -Parent $LogPath
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
}

# Resolve configuration from parameters, environment variables, then .env.
$dotEnvPath = Join-Path -Path $PSScriptRoot -ChildPath '.env'
$dotEnvValues = Get-DotEnvValues -Path $dotEnvPath

function Resolve-ConfigValue {
    param(
        [Parameter(Mandatory)][string]$ParamName,
        [string]$Current,
        [Parameter(Mandatory)][string]$EnvName,
        [hashtable]$DotEnv,
        [string]$Default = ''
    )
    if ($script:CallerBoundParameters.ContainsKey($ParamName) -and -not [string]::IsNullOrWhiteSpace($Current)) {
        return $Current
    }
    $envValue = [Environment]::GetEnvironmentVariable($EnvName)
    if (-not [string]::IsNullOrWhiteSpace($envValue)) { return $envValue }
    if ($DotEnv -and $DotEnv.ContainsKey($EnvName) -and -not [string]::IsNullOrWhiteSpace($DotEnv[$EnvName])) {
        return $DotEnv[$EnvName]
    }
    return $Default
}

$script:CallerBoundParameters = $PSBoundParameters
$Enterprise = Resolve-ConfigValue -ParamName 'Enterprise' -Current $Enterprise -EnvName 'GITHUB_ENTERPRISE' -DotEnv $dotEnvValues
$Token      = Resolve-ConfigValue -ParamName 'Token'      -Current $Token      -EnvName 'GITHUB_TOKEN'      -DotEnv $dotEnvValues
$ApiBaseUrl = Resolve-ConfigValue -ParamName 'ApiBaseUrl' -Current $ApiBaseUrl -EnvName 'GITHUB_API_BASE_URL' -DotEnv $dotEnvValues -Default 'https://api.github.com'

if ($dotEnvValues.Count -gt 0) {
    Write-Log "Loaded $($dotEnvValues.Count) value(s) from .env at '$dotEnvPath'." 'DEBUG'
}

$script:Token = $Token
$script:ApiBaseUrl = $ApiBaseUrl.TrimEnd('/')

# Validate required inputs.
if ([string]::IsNullOrWhiteSpace($Enterprise)) {
    throw "Enterprise is required. Pass -Enterprise or set the GITHUB_ENTERPRISE environment variable."
}
if ([string]::IsNullOrWhiteSpace($script:Token)) {
    throw "Token is required. Pass -Token or set the GITHUB_TOKEN environment variable."
}

Write-Log "===== GitHub Enterprise cost centre assignment ====="
Write-Log "Enterprise : $Enterprise"
Write-Log "API base   : $script:ApiBaseUrl"
Write-Log "CSV        : $CsvPath"
if ($WhatIfPreference) { Write-Log "Mode       : DRY RUN (-WhatIf) - no changes will be made" 'WARN' }

# Load and validate the CSV.
$rows = @(Import-Csv -LiteralPath $CsvPath)
if ($rows.Count -eq 0) {
    Write-Log "CSV contains no data rows. Nothing to do." 'WARN'
    return
}

$headers = $rows[0].PSObject.Properties.Name
$emailHeader = $headers | Where-Object { $_ -ieq $EmailColumn } | Select-Object -First 1
$ccHeader = $headers | Where-Object { $_ -ieq $CostCentreColumn } | Select-Object -First 1

if (-not $emailHeader) {
    throw "CSV is missing the email column '$EmailColumn'. Found columns: $($headers -join ', ')"
}
if (-not $ccHeader) {
    throw "CSV is missing the cost centre column '$CostCentreColumn'. Found columns: $($headers -join ', ')"
}

Write-Log "Loaded $($rows.Count) row(s) from CSV."

# Build identity lookup.
$emailMap = @{}
if (-not $SkipUserLookup) {
    $emailMap = Get-UserEmailMap -Enterprise $Enterprise
}
else {
    Write-Log "Skipping user lookup - treating the '$emailHeader' column as GitHub usernames." 'WARN'
}

# Load existing cost centres.
$costCentreMap = Get-CostCentreMap -Enterprise $Enterprise

# Resolve rows: group usernames by cost-centre id.
$assignments = @{}                                   # costCentreId -> HashSet[string] usernames (case-insensitive)
$ccNameById = @{}                                    # costCentreId -> name (for logging)
$unresolvedEmails  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$unknownCostCentres = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$rowNumber = 1

foreach ($row in $rows) {
    $rowNumber++
    $email = "$($row.$emailHeader)".Trim()
    $ccName = "$($row.$ccHeader)".Trim()

    if ([string]::IsNullOrWhiteSpace($email) -or [string]::IsNullOrWhiteSpace($ccName)) {
        Write-Log "Row ${rowNumber}: blank email or cost centre - skipping." 'WARN'
        continue
    }

    # Resolve username.
    if ($SkipUserLookup) {
        $username = $email
    }
    elseif ($emailMap.ContainsKey($email.ToLower())) {
        $username = $emailMap[$email.ToLower()]
    }
    else {
        Write-Log "Row ${rowNumber}: could not resolve '$email' to a GitHub login via Copilot seats + /users." 'ERROR'
        [void]$unresolvedEmails.Add($email)
        continue
    }

    # Resolve cost centre id (creating if requested).
    if (-not $costCentreMap.ContainsKey($ccName)) {
        if ($CreateMissingCostCentres) {
            if ($PSCmdlet.ShouldProcess($ccName, "Create cost centre")) {
                $newId = New-CostCentre -Enterprise $Enterprise -Name $ccName
                if ($newId) {
                    $costCentreMap[$ccName] = $newId
                }
                else {
                    [void]$unknownCostCentres.Add($ccName)
                    continue
                }
            }
            else {
                # -WhatIf: pretend the cost centre exists with a placeholder so we can
                # still report the intended assignment without mutating anything.
                $costCentreMap[$ccName] = "(would-create:$ccName)"
            }
        }
        else {
            Write-Log "Row ${rowNumber}: cost centre '$ccName' does not exist (use -CreateMissingCostCentres to create it)." 'ERROR'
            [void]$unknownCostCentres.Add($ccName)
            continue
        }
    }

    $ccId = $costCentreMap[$ccName]
    $ccNameById[$ccId] = $ccName
    if (-not $assignments.ContainsKey($ccId)) {
        $assignments[$ccId] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    [void]$assignments[$ccId].Add($username)
}

# Apply assignments per cost centre, in batches of 50, skipping existing members.
$totalAssigned = 0
$totalSkipped = 0
$totalFailed = 0

foreach ($ccId in $assignments.Keys) {
    $ccName = $ccNameById[$ccId]
    $requested = $assignments[$ccId]

    $isPlaceholder = "$ccId".StartsWith('(would-create:')

    # Determine who is already a member (skip them) - not possible for not-yet-created CCs.
    $existing = @()
    if (-not $isPlaceholder) {
        $existing = Get-CostCentreMembers -Enterprise $Enterprise -CostCentreId $ccId
    }
    $existingSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$existing, [System.StringComparer]::OrdinalIgnoreCase)

    $toAdd = [System.Collections.Generic.List[string]]::new()
    foreach ($u in $requested) {
        if ($existingSet.Contains($u)) {
            Write-Log "  $u already in '$ccName' - skipping." 'INFO'
            $totalSkipped++
        }
        else {
            $toAdd.Add($u)
        }
    }

    if ($toAdd.Count -eq 0) {
        Write-Log "Cost centre '$ccName': nothing to add." 'INFO'
        continue
    }

    Write-Log "Cost centre '$ccName' (id: $ccId): adding $($toAdd.Count) user(s)."

    for ($i = 0; $i -lt $toAdd.Count; $i += $script:MaxUsersPerRequest) {
        $batch = @($toAdd[$i..([math]::Min($i + $script:MaxUsersPerRequest - 1, $toAdd.Count - 1))])
        $target = "$($batch.Count) user(s) -> cost centre '$ccName'"

        if (-not $PSCmdlet.ShouldProcess($target, "Add to cost centre")) {
            Write-Log "  [WhatIf] would add: $($batch -join ', ')" 'INFO'
            $totalAssigned += $batch.Count
            continue
        }

        $result = Add-UsersToCostCentre -Enterprise $Enterprise -CostCentreId $ccId -Usernames $batch
        if ($result.Success) {
            Write-Log "  Added $($batch.Count) user(s): $($batch -join ', ')" 'SUCCESS'
            $totalAssigned += $batch.Count
        }
        else {
            Write-Log "  Failed to add batch to '$ccName' (HTTP $($result.StatusCode)): $($result.Content)" 'ERROR'
            $totalFailed += $batch.Count
        }
    }
}

# Summary.
Write-Log "===== Summary ====="
Write-Log "Assigned : $totalAssigned" 'SUCCESS'
Write-Log "Skipped  : $totalSkipped (already assigned)"
Write-Log "Failed   : $totalFailed" $(if ($totalFailed -gt 0) { 'ERROR' } else { 'INFO' })
if ($unresolvedEmails.Count -gt 0) {
    Write-Log "Unresolved emails ($($unresolvedEmails.Count)): $(($unresolvedEmails) -join ', ')" 'WARN'
}
if ($unknownCostCentres.Count -gt 0) {
    Write-Log "Unknown/uncreatable cost centres ($($unknownCostCentres.Count)): $(($unknownCostCentres) -join ', ')" 'WARN'
}

if ($totalFailed -gt 0 -or $unresolvedEmails.Count -gt 0) {
    exit 1
}

#endregion Main ----------------------------------------------------------------
