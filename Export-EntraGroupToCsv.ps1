#Requires -Version 7.0

<#
.SYNOPSIS
    Exports members of an Entra ID (Azure AD) group to a CSV compatible with
    Set-CostCentreMembership.ps1.

.DESCRIPTION
    Queries Microsoft Graph for the members of an Entra ID group and writes their
    display name, email address and department to a CSV file. The department is
    written to the cost-centre column so the resulting file can be fed straight
    into Set-CostCentreMembership.ps1.

    End-to-end workflow:
      1. Run this script with a group name or id -> produces a CSV.
      2. Review the CSV (optionally remap departments to cost-centre names).
      3. Run Set-CostCentreMembership.ps1 with that CSV to assign users.

    Authentication (one of):
      * -AccessToken : a pre-acquired Graph bearer token.
      * Client credentials : -TenantId, -ClientId and -ClientSecret (app-only).
        Requires the app registration to have the Graph application permissions
        'GroupMember.Read.All' (or 'Group.Read.All') and 'User.Read.All' with
        admin consent granted.

.PARAMETER GroupId
    The object id (GUID) of the Entra group. Mutually exclusive with -GroupName.

.PARAMETER GroupName
    The display name of the Entra group. Resolved to an id via Graph. If multiple
    groups share the name the script aborts and lists the matches.

.PARAMETER OutputPath
    Path to the CSV file to create. Default: .\outputs\entra-group-export.csv.

.PARAMETER TenantId
    Entra tenant id (GUID or domain). Resolved in order: parameter > AZURE_TENANT_ID
    environment variable > AZURE_TENANT_ID in $PSScriptRoot/.env.

.PARAMETER ClientId
    App registration (client) id. Resolved in order: parameter > AZURE_CLIENT_ID env
    var > AZURE_CLIENT_ID in $PSScriptRoot/.env.

.PARAMETER ClientSecret
    App registration client secret. Resolved in order: parameter > AZURE_CLIENT_SECRET
    env var > AZURE_CLIENT_SECRET in $PSScriptRoot/.env.

.PARAMETER AccessToken
    A pre-acquired Microsoft Graph access token. If supplied, the client-credential
    parameters are ignored. Resolved in order: parameter > AZURE_ACCESS_TOKEN env var
    > AZURE_ACCESS_TOKEN in $PSScriptRoot/.env.

.PARAMETER GraphBaseUrl
    Microsoft Graph base URL. Resolved in order: parameter > GRAPH_BASE_URL env var >
    GRAPH_BASE_URL in $PSScriptRoot/.env > https://graph.microsoft.com/v1.0.
    Use https://graph.microsoft.us/v1.0 for US Government, etc.

.PARAMETER Transitive
    Include members of nested groups (transitive membership) rather than only
    direct members.

.PARAMETER GitHubUsernameProperty
    Name of the Entra/Graph user property that holds the user's GitHub username.
    When supplied, this property's value is written to the email column instead of
    the user's email/UPN, so the resulting CSV can be fed to
    Set-CostCentreMembership.ps1 with -SkipUserLookup to look up users directly by
    their GitHub handle (no Copilot seats / /users lookup required).

    Supports:
      * Top-level user properties (e.g. 'userPrincipalName').
      * Schema/on-premises extension attributes by short name, e.g.
        'extensionAttribute1' (resolved under onPremisesExtensionAttributes).
      * Dot-notation paths for nested or directory-extension properties, e.g.
        'onPremisesExtensionAttributes.extensionAttribute1' or
        'extension_<appId>_githubUsername'.

.PARAMETER EmailColumn
    Header name for the email column. Default: 'email' (matches the cost-centre script).
    When -GitHubUsernameProperty is used, the GitHub username is written to this column.

.PARAMETER CostCentreColumn
    Header name for the cost-centre column. Default: 'cost_centre' (matches the
    cost-centre script). The cost-centre value is written here.

.PARAMETER CostCentreProperty
    Name of the Entra/Graph user property whose value is written to the cost-centre
    column. Default: 'department'.

    Supports the same forms as -GitHubUsernameProperty:
      * Top-level user properties (e.g. 'jobTitle', 'companyName').
      * Schema/on-premises extension attributes by short name, e.g.
        'extensionAttribute2' (resolved under onPremisesExtensionAttributes).
      * Dot-notation paths for nested or directory-extension properties, e.g.
        'onPremisesExtensionAttributes.extensionAttribute2' or
        'extension_<appId>_costCentre'.

.PARAMETER DefaultCostCentre
    Value to use when a member has no value for the cost-centre source property.
    Without this, members lacking a value are written with a blank cost-centre and
    a warning is logged.

.PARAMETER LogPath
    Optional path to a log file. Console output is also written there.

.EXAMPLE
    .\Export-EntraGroupToCsv.ps1 -GroupName "Engineering" -OutputPath .\eng.csv `
        -TenantId $env:AZURE_TENANT_ID -ClientId $env:AZURE_CLIENT_ID -ClientSecret $env:AZURE_CLIENT_SECRET

.EXAMPLE
    # Using an already-acquired token, then feeding the CSV to the cost-centre script
    .\Export-EntraGroupToCsv.ps1 -GroupId 00000000-0000-0000-0000-000000000000 -AccessToken $token -OutputPath .\users.csv
    .\Set-CostCentreMembership.ps1 -CsvPath .\users.csv -Enterprise contoso -WhatIf

.EXAMPLE
    # Export the GitHub username (stored in extensionAttribute1) instead of email,
    # then assign cost centres by looking users up directly by their GitHub handle.
    .\Export-EntraGroupToCsv.ps1 -GroupName "Engineering" -GitHubUsernameProperty extensionAttribute1 `
        -OutputPath .\handles.csv
    .\Set-CostCentreMembership.ps1 -CsvPath .\handles.csv -Enterprise contoso -SkipUserLookup -WhatIf

.NOTES
    Output CSV columns: name,email,cost_centre
    The 'name' column is ignored by Set-CostCentreMembership.ps1 but kept for review.
#>
[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ById')]
    [string]$GroupId,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByName')]
    [string]$GroupName,

    [Parameter()]
    [string]$OutputPath = '.\outputs\entra-group-export.csv',

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$ClientSecret,

    [Parameter()]
    [string]$AccessToken,

    [Parameter()]
    [string]$GraphBaseUrl,

    [Parameter()]
    [switch]$Transitive,

    [Parameter()]
    [string]$GitHubUsernameProperty,

    [Parameter()]
    [string]$EmailColumn = 'email',

    [Parameter()]
    [string]$CostCentreColumn = 'cost_centre',

    [Parameter()]
    [string]$CostCentreProperty = 'department',

    [Parameter()]
    [string]$DefaultCostCentre,

    [Parameter()]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function Get-GraphToken {
    <#
        Acquires an app-only Microsoft Graph token via the client-credentials flow.
        Returns the access token string.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$ClientSecret,
        [Parameter(Mandatory = $true)][string]$GraphBaseUrl
    )

    # Derive the login authority and scope host from the Graph base URL so the
    # script also works against sovereign clouds (US Gov, China, etc.).
    $scopeHost = ([System.Uri]$GraphBaseUrl).GetLeftPart([System.UriPartial]::Authority)
    $loginHost = switch -Wildcard ($scopeHost) {
        '*graph.microsoft.us'        { 'https://login.microsoftonline.us' }
        '*microsoftgraph.chinacloudapi.cn' { 'https://login.chinacloudapi.cn' }
        default                       { 'https://login.microsoftonline.com' }
    }

    $tokenUri = "$loginHost/$TenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "$scopeHost/.default"
        grant_type    = 'client_credentials'
    }

    Write-Log "Acquiring Graph token via client credentials (tenant: $TenantId)..."
    try {
        $resp = Invoke-RestMethod -Uri $tokenUri -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded'
    }
    catch {
        throw "Failed to acquire Graph token: $($_.Exception.Message)"
    }

    if (-not $resp.access_token) {
        throw "Token endpoint did not return an access_token."
    }
    Write-Log "Token acquired." 'SUCCESS'
    return $resp.access_token
}

function Invoke-GraphApi {
    <#
        GET wrapper for Microsoft Graph with retry on 429/5xx. Returns the parsed
        response object. Honours the Retry-After header on throttling.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [int]$MaxRetries = 4
    )

    $headers = @{
        'Authorization' = "Bearer $script:Token"
        'Accept'        = 'application/json'
        'ConsistencyLevel' = 'eventual'
    }

    $attempt = 0
    while ($true) {
        $attempt++
        $statusCode = 0
        try {
            $content = Invoke-RestMethod -Uri $Uri -Method Get -Headers $headers `
                -SkipHttpErrorCheck -StatusCodeVariable statusCode -ResponseHeadersVariable respHeaders
        }
        catch {
            if ($attempt -le $MaxRetries) {
                $wait = [math]::Pow(2, $attempt)
                Write-Log "Graph request failed ($($_.Exception.Message)). Retrying in $wait s..." 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
            throw "Graph request to $Uri failed: $($_.Exception.Message)"
        }

        if ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)) {
            if ($attempt -le $MaxRetries) {
                $wait = [math]::Pow(2, $attempt)
                if ($respHeaders -and $respHeaders['Retry-After']) {
                    $retryAfter = $respHeaders['Retry-After']
                    if ($retryAfter -is [array]) { $retryAfter = $retryAfter[0] }
                    [int]$parsed = 0
                    if ([int]::TryParse("$retryAfter", [ref]$parsed) -and $parsed -gt 0) { $wait = $parsed }
                }
                Write-Log "Graph returned $statusCode on $Uri. Waiting $wait s (attempt $attempt/$MaxRetries)..." 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
        }

        if ($statusCode -ge 200 -and $statusCode -lt 300) {
            return $content
        }

        throw "Graph request to $Uri failed (HTTP $statusCode): $($content | ConvertTo-Json -Depth 6 -Compress)"
    }
}

function Resolve-GroupId {
    <# Resolves a group display name to its object id. #>
    param([Parameter(Mandatory = $true)][string]$Name)

    $encoded = [System.Uri]::EscapeDataString($Name)
    $uri = "$script:GraphBaseUrl/groups?`$filter=displayName eq '$encoded'&`$select=id,displayName"
    $resp = Invoke-GraphApi -Uri $uri

    $groups = @()
    if ($resp.PSObject.Properties.Name -contains 'value' -and $resp.value) {
        $groups = @($resp.value)
    }

    if ($groups.Count -eq 0) {
        throw "No Entra group found with display name '$Name'."
    }
    if ($groups.Count -gt 1) {
        $ids = ($groups | ForEach-Object { $_.id }) -join ', '
        throw "Multiple groups named '$Name' found ($ids). Re-run with -GroupId to disambiguate."
    }

    Write-Log "Resolved group '$Name' to id $($groups[0].id)." 'INFO'
    return $groups[0].id
}

function Get-GroupMemberIds {
    <#
        Returns the object ids of all user members of a group (direct or transitive).
        The /members endpoint reliably returns user ids even though it does not
        always propagate extended user properties like 'department'.
    #>
    param([Parameter(Mandatory = $true)][string]$Id)

    $membersPath = if ($Transitive) { 'transitiveMembers' } else { 'members' }
    # Request only the id — this is all we need from the members endpoint.
    $uri = "$script:GraphBaseUrl/groups/$Id/$membersPath/microsoft.graph.user?`$select=id&`$top=999"

    $ids = [System.Collections.Generic.List[string]]::new()
    while ($uri) {
        $resp = Invoke-GraphApi -Uri $uri
        if ($resp.PSObject.Properties.Name -contains 'value' -and $resp.value) {
            foreach ($u in $resp.value) { if ($u.id) { $ids.Add($u.id) } }
        }
        $uri = if ($resp.PSObject.Properties.Name -contains '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
    }
    return $ids
}

function Get-UserPropertySelect {
    <#
        Returns the top-level Graph $select segment required to retrieve the
        given user property. Short extensionAttributeN names are resolved under
        onPremisesExtensionAttributes; dot-paths use their first segment.
    #>
    param([Parameter(Mandatory = $true)][string]$Property)

    if ($Property -match '^extensionAttribute([1-9]|1[0-5])$') {
        return 'onPremisesExtensionAttributes'
    }
    return ($Property -split '\.')[0]
}

function Get-UserPropertyValue {
    <#
        Resolves a value from a user object for the given property name. Supports
        top-level properties, extensionAttributeN short names (resolved under
        onPremisesExtensionAttributes) and dot-notation paths.
    #>
    param(
        [Parameter(Mandatory = $true)]$User,
        [Parameter(Mandatory = $true)][string]$Property
    )

    if ($Property -match '^extensionAttribute([1-9]|1[0-5])$') {
        $path = @('onPremisesExtensionAttributes', $Property)
    }
    else {
        $path = $Property -split '\.'
    }

    $value = $User
    foreach ($segment in $path) {
        if ($null -eq $value) { return $null }
        $prop = $value.PSObject.Properties[$segment]
        if (-not $prop) { return $null }
        $value = $prop.Value
    }
    return $value
}

function Get-UsersByIds {
    <#
        Fetches full user objects (including 'department') for a list of user ids by
        querying the /users endpoint with an 'id in (...)' filter. Graph supports up
        to ~15 ids per filter clause, so we batch in groups of 15.
    #>
    param([Parameter(Mandatory = $true)][string[]]$Ids)

    $selectParts = [System.Collections.Generic.List[string]]::new()
    foreach ($p in @('id', 'displayName', 'mail', 'userPrincipalName')) { $selectParts.Add($p) }

    foreach ($prop in @($CostCentreProperty, $GitHubUsernameProperty)) {
        if ([string]::IsNullOrWhiteSpace($prop)) { continue }
        $segment = Get-UserPropertySelect -Property $prop
        if (-not $selectParts.Contains($segment)) { $selectParts.Add($segment) }
    }
    $select = $selectParts -join ','
    $batchSize = 15
    $users = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $Ids.Count; $i += $batchSize) {
        $batch = $Ids[$i..([math]::Min($i + $batchSize - 1, $Ids.Count - 1))]
        $inClause = ($batch | ForEach-Object { "'$_'" }) -join ','
        $uri = "$script:GraphBaseUrl/users?`$filter=id in ($inClause)&`$select=$select&`$top=999"

        while ($uri) {
            $resp = Invoke-GraphApi -Uri $uri
            if ($resp.PSObject.Properties.Name -contains 'value' -and $resp.value) {
                foreach ($u in $resp.value) { $users.Add($u) }
            }
            $uri = if ($resp.PSObject.Properties.Name -contains '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
        }
    }
    return $users
}

function Get-GroupMembers {
    <#
        Returns the user members of a group, each with displayName, mail,
        userPrincipalName and the configured cost-centre source property.

        Strategy:
          Step 1 - get member ids from the group members endpoint (reliably returns
                   ids but not always extended properties like 'department').
          Step 2 - fetch full user objects in batches from /users?$filter=id in (...)
                   which reliably returns all selected properties.
    #>
    param([Parameter(Mandatory = $true)][string]$Id)

    Write-Log "Fetching member ids from group..."
    $ids = @(Get-GroupMemberIds -Id $Id)
    Write-Log "Found $($ids.Count) user member id(s). Fetching full user profiles..."

    if ($ids.Count -eq 0) { return @() }

    $users = @(Get-UsersByIds -Ids $ids)
    Write-Log "Retrieved $($users.Count) full user profile(s)." 'INFO'
    return $users
}

#endregion Helpers -------------------------------------------------------------

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
$TenantId     = Resolve-ConfigValue -ParamName 'TenantId'     -Current $TenantId     -EnvName 'AZURE_TENANT_ID'     -DotEnv $dotEnvValues
$ClientId     = Resolve-ConfigValue -ParamName 'ClientId'     -Current $ClientId     -EnvName 'AZURE_CLIENT_ID'     -DotEnv $dotEnvValues
$ClientSecret = Resolve-ConfigValue -ParamName 'ClientSecret' -Current $ClientSecret -EnvName 'AZURE_CLIENT_SECRET' -DotEnv $dotEnvValues
$AccessToken  = Resolve-ConfigValue -ParamName 'AccessToken'  -Current $AccessToken  -EnvName 'AZURE_ACCESS_TOKEN'  -DotEnv $dotEnvValues
$GraphBaseUrl = Resolve-ConfigValue -ParamName 'GraphBaseUrl' -Current $GraphBaseUrl -EnvName 'GRAPH_BASE_URL'      -DotEnv $dotEnvValues -Default 'https://graph.microsoft.com/v1.0'

if ($dotEnvValues.Count -gt 0) {
    Write-Log "Loaded $($dotEnvValues.Count) value(s) from .env at '$dotEnvPath'." 'DEBUG'
}

$script:GraphBaseUrl = $GraphBaseUrl.TrimEnd('/')

Write-Log "===== Entra group -> CSV export ====="
Write-Log "Graph base : $script:GraphBaseUrl"
Write-Log "Output     : $OutputPath"
Write-Log "Cost centre: property '$CostCentreProperty' (written to '$CostCentreColumn' column)."
if (-not [string]::IsNullOrWhiteSpace($GitHubUsernameProperty)) {
    Write-Log "Identifier : GitHub username from property '$GitHubUsernameProperty' (written to '$EmailColumn' column)."
}

# Acquire / set the token.
if ($AccessToken) {
    Write-Log "Using supplied access token."
    $script:Token = $AccessToken
}
else {
    if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($ClientSecret)) {
        throw "Provide either -AccessToken, or -TenantId/-ClientId/-ClientSecret (or set AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET via environment variables or .env)."
    }
    $script:Token = Get-GraphToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -GraphBaseUrl $script:GraphBaseUrl
}

# Resolve the group id.
$resolvedId = if ($PSCmdlet.ParameterSetName -eq 'ById') { $GroupId } else { Resolve-GroupId -Name $GroupName }

# Fetch members.
$members = Get-GroupMembers -Id $resolvedId

# Build CSV rows.
$rows = [System.Collections.Generic.List[object]]::new()
$missingEmail = [System.Collections.Generic.List[string]]::new()
$missingDept = [System.Collections.Generic.List[string]]::new()

$usingGitHubUsername = -not [string]::IsNullOrWhiteSpace($GitHubUsernameProperty)

foreach ($m in $members) {
    $name = "$($m.displayName)".Trim()

    if ($usingGitHubUsername) {
        # Write the GitHub username (from the configured Entra property) to the
        # email column so the CSV can be used with -SkipUserLookup downstream.
        $email = "$(Get-UserPropertyValue -User $m -Property $GitHubUsernameProperty)".Trim()
    }
    else {
        # Prefer mail, fall back to userPrincipalName.
        $email = "$($m.mail)".Trim()
        if ([string]::IsNullOrWhiteSpace($email)) { $email = "$($m.userPrincipalName)".Trim() }
    }

    $dept = "$(Get-UserPropertyValue -User $m -Property $CostCentreProperty)".Trim()
    if ([string]::IsNullOrWhiteSpace($dept) -and $DefaultCostCentre) { $dept = $DefaultCostCentre }

    if ([string]::IsNullOrWhiteSpace($email)) {
        $reason = $usingGitHubUsername ? "no GitHub username ('$GitHubUsernameProperty')" : 'no email/UPN'
        Write-Log "Member '$name' has $reason - skipping." 'WARN'
        $missingEmail.Add(($name ? $name : '(unknown)'))
        continue
    }
    if ([string]::IsNullOrWhiteSpace($dept)) {
        Write-Log "Member '$name' <$email> has no value for '$CostCentreProperty' - cost centre will be blank." 'WARN'
        $missingDept.Add($email)
    }

    $row = [ordered]@{
        name              = $name
        $EmailColumn      = $email
        $CostCentreColumn = $dept
    }
    $rows.Add([pscustomobject]$row)
}

# Write the CSV.
$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
$rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8

# Summary.
Write-Log "===== Summary ====="
Write-Log "Members written : $($rows.Count)" 'SUCCESS'
if ($missingEmail.Count -gt 0) {
    $skipLabel = $usingGitHubUsername ? 'no GitHub username' : 'no email'
    Write-Log "Skipped ($skipLabel): $($missingEmail.Count) - $(( $missingEmail | Select-Object -Unique) -join ', ')" 'WARN'
}
if ($missingDept.Count -gt 0) {
    Write-Log "Blank cost centre ($($missingDept.Count)): $(( $missingDept | Select-Object -Unique) -join ', ')" 'WARN'
    Write-Log "Tip: pass -DefaultCostCentre or edit the CSV before running Set-CostCentreMembership.ps1." 'INFO'
}
Write-Log "CSV written to: $OutputPath" 'SUCCESS'
$nextCmd = $usingGitHubUsername `
    ? ".\Set-CostCentreMembership.ps1 -CsvPath '$OutputPath' -Enterprise <slug> -SkipUserLookup -WhatIf" `
    : ".\Set-CostCentreMembership.ps1 -CsvPath '$OutputPath' -Enterprise <slug> -WhatIf"
Write-Log "Next: $nextCmd" 'INFO'

#endregion Main ----------------------------------------------------------------
