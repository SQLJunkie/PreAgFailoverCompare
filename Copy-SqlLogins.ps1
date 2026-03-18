# Copy-SqlLogins.ps1
# Copies logins missing on the Target that exist on the Source.
# SQL logins are created with the original password hash and SID.
# Windows logins are created as FROM WINDOWS.
param(
    [Parameter(Mandatory=$true)]
    [string]$SourceInstance,

    [Parameter(Mandatory=$true)]
    [string]$TargetInstance,

    # Optional: restrict to specific login names
    [string[]]$LoginName,

    [switch]$WhatIf
)

$requiredDbatoolsVersion = [version]'2.7.25'
$loadedDbatools = Get-Module -Name dbatools
if (-not $loadedDbatools -or $loadedDbatools.Version -ne $requiredDbatoolsVersion) {
    Import-Module dbatools -RequiredVersion $requiredDbatoolsVersion -Force -ErrorAction Stop
}

function ConvertTo-HexString {
    param([byte[]]$Bytes)
    '0x' + [BitConverter]::ToString($Bytes).Replace('-', '')
}

$commonQueryParams = @{
    As          = 'DataTable'
    ErrorAction = 'Stop'
}

# -- Connect ---------------------------------------------------------------

Write-Host "Connecting to source: $SourceInstance" -ForegroundColor Cyan
try {
    $src = Connect-DbaInstance -SqlInstance $SourceInstance -TrustServerCertificate -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to source '$SourceInstance': $_"
    exit 1
}

Write-Host "Connecting to target: $TargetInstance" -ForegroundColor Cyan
try {
    $tgt = Connect-DbaInstance -SqlInstance $TargetInstance -TrustServerCertificate -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to target '$TargetInstance': $_"
    exit 1
}

# -- Query logins ----------------------------------------------------------

$srcLoginQuery = @"
SELECT
    sp.name,
    sp.type,
    sp.type_desc,
    sp.default_database_name,
    sp.default_language_name,
    sl.password_hash,
    sp.sid,
    ISNULL(sl.is_policy_checked,    0) AS is_policy_checked,
    ISNULL(sl.is_expiration_checked, 0) AS is_expiration_checked
FROM sys.server_principals sp
LEFT JOIN sys.sql_logins sl ON sp.principal_id = sl.principal_id
WHERE sp.type IN ('S', 'U', 'G')
  AND sp.name NOT LIKE '##%'
  AND sp.name <> 'sa'
  AND sp.name NOT LIKE 'NT AUTHORITY\%'
  AND sp.name NOT LIKE 'NT SERVICE\%'
ORDER BY sp.name
"@

Write-Host "Querying logins on source..." -ForegroundColor Cyan
$srcLogins = Invoke-DbaQuery -SqlInstance $src -Query $srcLoginQuery @commonQueryParams

Write-Host "Querying logins on target..." -ForegroundColor Cyan
$tgtLoginRows = Invoke-DbaQuery -SqlInstance $tgt `
    -Query "SELECT name FROM sys.server_principals WHERE type IN ('S','U','G') AND name NOT LIKE '##%'" `
    @commonQueryParams
$tgtLoginSet = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@($tgtLoginRows.name),
    [System.StringComparer]::OrdinalIgnoreCase
)

# Used to validate default_database_name before creating each login
$tgtDbSet = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@((Invoke-DbaQuery -SqlInstance $tgt -Query "SELECT name FROM sys.databases" @commonQueryParams).name),
    [System.StringComparer]::OrdinalIgnoreCase
)

# -- Filter ----------------------------------------------------------------

$toProcess = @($srcLogins | Where-Object { -not $tgtLoginSet.Contains($_.name) })

if ($LoginName) {
    $filterSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]$LoginName,
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $toProcess = @($toProcess | Where-Object { $filterSet.Contains($_.name) })
}

if ($toProcess.Count -eq 0) {
    Write-Host "No missing logins found — nothing to do." -ForegroundColor Green
    exit 0
}

Write-Host "$($toProcess.Count) login(s) to create on target." -ForegroundColor Yellow
if ($WhatIf) { Write-Host "(WhatIf mode — no changes will be made)" -ForegroundColor DarkCyan }
Write-Host ""

# -- Create logins ---------------------------------------------------------

$created = 0
$skipped = 0
$failed  = 0

foreach ($login in $toProcess) {
    $name     = $login.name
    $safeName = $name.Replace(']', ']]')

    # Fall back to master if the login's default DB doesn't exist on target
    $defaultDb   = if ($tgtDbSet.Contains($login.default_database_name)) { $login.default_database_name } else { 'master' }
    $safeDb      = $defaultDb.Replace(']', ']]')
    $langName    = if ($login.default_language_name) { $login.default_language_name } else { 'us_english' }
    $safeLang    = $langName.Replace(']', ']]')

    if ($login.type -eq 'S') {
        # SQL login — requires password hash and SID
        if (-not $login.password_hash -or -not $login.sid) {
            Write-Warning "  SKIP    [$name] — missing password_hash or SID."
            $skipped++
            continue
        }

        $hashHex = ConvertTo-HexString $login.password_hash
        $sidHex  = ConvertTo-HexString $login.sid
        $policy  = if ($login.is_policy_checked)    { 'ON' } else { 'OFF' }
        $expiry  = if ($login.is_expiration_checked) { 'ON' } else { 'OFF' }

        $sql = "CREATE LOGIN [$safeName] WITH PASSWORD = $hashHex HASHED, SID = $sidHex, DEFAULT_DATABASE = [$safeDb], DEFAULT_LANGUAGE = [$safeLang], CHECK_POLICY = $policy, CHECK_EXPIRATION = $expiry;"
    }
    elseif ($login.type -in 'U', 'G') {
        # Windows user / group
        $sql = "CREATE LOGIN [$safeName] FROM WINDOWS WITH DEFAULT_DATABASE = [$safeDb], DEFAULT_LANGUAGE = [$safeLang];"
    }
    else {
        Write-Warning "  SKIP    [$name] — unsupported type '$($login.type)'."
        $skipped++
        continue
    }

    if ($WhatIf) {
        Write-Host "  WHATIF  [$name] ($($login.type_desc))" -ForegroundColor DarkCyan
        Write-Host "          $sql" -ForegroundColor DarkGray
    }
    else {
        try {
            Invoke-DbaQuery -SqlInstance $tgt -Query $sql -ErrorAction Stop -WarningAction Stop
            Write-Host "  CREATED [$name] ($($login.type_desc))" -ForegroundColor Green
            $created++
        }
        catch {
            Write-Warning "  FAILED  [$name]: $_"
            $failed++
        }
    }
}

# -- Summary ---------------------------------------------------------------

Write-Host ""
if ($WhatIf) {
    Write-Host "WhatIf mode — no logins were created." -ForegroundColor Yellow
}
else {
    Write-Host "Done.  Created: $created  |  Skipped: $skipped  |  Failed: $failed" -ForegroundColor Cyan
}
