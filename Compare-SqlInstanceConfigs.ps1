param(
    [Parameter(Mandatory=$true)]
    [string]$SourceInstance,

    [Parameter(Mandatory=$true)]
    [string]$TargetInstance
)

# Ensure dbatools is loaded at exactly the required version.
# -Force reloads if a different version is already in the session.
$requiredDbatoolsVersion = [version]'2.7.25'
$loadedDbatools = Get-Module -Name dbatools
if (-not $loadedDbatools -or $loadedDbatools.Version -ne $requiredDbatoolsVersion) {
    Import-Module dbatools -RequiredVersion $requiredDbatoolsVersion -Force -ErrorAction Stop
}

Write-Host "Loading comparison functions..." -ForegroundColor Cyan

# Clear old functions to force reload
Get-ChildItem Function:\Compare-* | Remove-Item -Force -ErrorAction SilentlyContinue
Remove-Item Function:\Add-Row -Force -ErrorAction SilentlyContinue
Remove-Item Function:\Write-HtmlReport -Force -ErrorAction SilentlyContinue
Remove-Item Function:\Get-AgDbNamesSafe -Force -ErrorAction SilentlyContinue


# Dot-source the functions file located in the same directory as this script
$functionsFile = Join-Path $PSScriptRoot "SqlCompareFunctions.ps1"

if (-not (Test-Path $functionsFile)) {
    Write-Error "Functions file not found: $functionsFile"
    exit 1
}

. $functionsFile

Write-Host "Connecting to SQL instances..." -ForegroundColor Cyan

try {
    $src = Connect-DbaInstance -SqlInstance $SourceInstance -TrustServerCertificate -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to source instance $SourceInstance"
    throw
}

try {
    $tgt = Connect-DbaInstance -SqlInstance $TargetInstance -TrustServerCertificate -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to target instance $TargetInstance"
    throw
}


# Clear previous results
$script:rows = @()

# Run comparisons
Compare-ServerLevel    -Source $src -Target $tgt
Compare-Logins         -Source $src -Target $tgt
Compare-NonAgDatabases -Source $src -Target $tgt

# Generate HTML report
$myReport = Write-HtmlReport -SourceInstance $SourceInstance -TargetInstance $TargetInstance

Write-Host "Comparison complete. HTML report generated." -ForegroundColor Green

# Automatically open the report
Start-Process $myReport
