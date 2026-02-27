<#
.SYNOPSIS
    Debug wrapper for Kopano IMAP to Graph Migration.
    Runs the migration script with all diagnostic options enabled.

.DESCRIPTION
    This script is a convenience wrapper that runs Kopano-IMAP-to-Graph-Migration.ps1
    with full diagnostic logging enabled. Use it to troubleshoot migration issues.

    It enables:
    - DiagnosticMode: Full IMAP/Graph API logging
    - SaveMimeToFile: Saves raw MIME content to disk
    - VerboseLogging: All debug messages shown
    - ContinueOnError: Don't stop on individual message failures

.PARAMETER ImapServer
    IMAP server hostname

.PARAMETER TestSource
    Source IMAP email/username

.PARAMETER TestTarget
    Target M365 mailbox

.PARAMETER TestPassword
    IMAP password

.PARAMETER TenantId
    Microsoft 365 tenant ID

.PARAMETER ClientId
    Azure AD application client ID

.PARAMETER ClientSecret
    Azure AD application client secret

.PARAMETER TestSingleMessage
    Fetch only one message and display details (default: true)

.PARAMETER SkipGraphImport
    Test IMAP fetch only, skip Graph API import

.EXAMPLE
    .\Kopano-IMAP-Migration-Debug.ps1 `
        -ImapServer "imap.elkw.de" `
        -TestSource "testuser@elkw.de" `
        -TestTarget "testuser@elkw.de" `
        -TestPassword "xxx" `
        -TenantId "xxx" `
        -ClientId "xxx" `
        -ClientSecret "xxx"

.EXAMPLE
    # Test IMAP connectivity only (no Graph API needed):
    .\Kopano-IMAP-Migration-Debug.ps1 `
        -ImapServer "imap.elkw.de" `
        -TestSource "testuser@elkw.de" `
        -TestTarget "testuser@elkw.de" `
        -TestPassword "xxx" `
        -TenantId "dummy" `
        -ClientId "dummy" `
        -ClientSecret "dummy" `
        -SkipGraphImport
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ImapServer,

    [Parameter(Mandatory)]
    [string]$TestSource,

    [Parameter(Mandatory)]
    [string]$TestTarget,

    [Parameter(Mandatory)]
    [string]$TestPassword,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$ClientSecret,

    [int]$ImapPort = 993,
    [switch]$ImapSkipCertValidation,
    [switch]$TestSingleMessage = $true,
    [switch]$SkipGraphImport,
    [int]$MaxMessages = 3,
    [string[]]$FoldersToMigrate = @("INBOX")
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Kopano IMAP Migration - DEBUG MODE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "IMAP Server: $ImapServer`:$ImapPort" -ForegroundColor Yellow
Write-Host "Source: $TestSource" -ForegroundColor Yellow
Write-Host "Target: $TestTarget" -ForegroundColor Yellow
Write-Host "Folders: $($FoldersToMigrate -join ', ')" -ForegroundColor Yellow
Write-Host "Max Messages: $MaxMessages" -ForegroundColor Yellow
Write-Host "Test Single Message: $TestSingleMessage" -ForegroundColor Yellow
Write-Host "Skip Graph Import: $SkipGraphImport" -ForegroundColor Yellow
Write-Host ""

# Build parameters
$params = @{
    TenantId             = $TenantId
    ClientId             = $ClientId
    ClientSecret         = $ClientSecret
    ImapServer           = $ImapServer
    ImapPort             = $ImapPort
    ImapUseSsl           = $true
    TestMode             = $true
    TestSource           = $TestSource
    TestTarget           = $TestTarget
    TestPassword         = $TestPassword
    MaxMessagesPerMailbox = $MaxMessages
    FoldersToMigrate     = $FoldersToMigrate
    VerboseLogging       = $true
    DiagnosticMode       = $true
    SaveMimeToFile       = $true
    ContinueOnError      = $true
}

if ($ImapSkipCertValidation) {
    $params.ImapSkipCertValidation = $true
}

if ($TestSingleMessage) {
    $params.TestSingleMessage = $true
}

if ($SkipGraphImport) {
    $params.SkipGraphImport = $true
}

# Run the migration script with debug options
$scriptPath = Join-Path $PSScriptRoot "Kopano-IMAP-to-Graph-Migration.ps1"

if (!(Test-Path $scriptPath)) {
    Write-Host "ERROR: Migration script not found at: $scriptPath" -ForegroundColor Red
    exit 1
}

Write-Host "Running: $scriptPath" -ForegroundColor Gray
Write-Host "With parameters:" -ForegroundColor Gray
$params.GetEnumerator() | Where-Object { $_.Key -notmatch 'Secret|Password' } | Sort-Object Key | ForEach-Object {
    Write-Host "  -$($_.Key) $($_.Value)" -ForegroundColor Gray
}
Write-Host ""

try {
    & $scriptPath @params
}
catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  DEBUG: Migration failed with error:" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Full error:" -ForegroundColor Yellow
    Write-Host $_ -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Debug session complete." -ForegroundColor Cyan

# Check if MIME files were saved
$mimeDumpPath = Join-Path $PSScriptRoot "mime_dump"
if (Test-Path $mimeDumpPath) {
    $emlFiles = Get-ChildItem "$mimeDumpPath\*.eml" -ErrorAction SilentlyContinue
    if ($emlFiles) {
        Write-Host "Saved MIME files ($($emlFiles.Count)):" -ForegroundColor Green
        foreach ($f in $emlFiles) {
            $sizeKb = "{0:N1} KB" -f ($f.Length / 1024)
            Write-Host "  $($f.Name) ($sizeKb)" -ForegroundColor Green
        }
    }
}

# Check for log files
$logPath = Join-Path $PSScriptRoot "migration_log"
if (Test-Path $logPath) {
    $latestLog = Get-ChildItem "$logPath\*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) {
        Write-Host "Latest log: $($latestLog.FullName)" -ForegroundColor Green
    }
}
