<#
.SYNOPSIS
    Minimal splatting example for quick setup
#>

# ============================================
# OPTION 1: Full CSV-based migration
# ============================================
$params = @{
    # Graph API Auth
    TenantId                = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    ClientId                = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    ClientSecret            = "your-secret-here"

    # IMAP Source
    ImapServer              = "mail.kopano.local"
    ImapPort                = 993
    ImapUseSsl              = $true
    ImapSkipCertValidation  = $false                # $true for self-signed certs

    # Users (CSV mode)
    UserCsvPath             = ".\users.csv"

    # Options
    PreserveFolderStructure = $true
    PreserveReceivedDate    = $true
    ContinueOnError         = $true

    # Logging
    LogPath                 = ".\migration_logs"
    StateFile               = ".\migration_state.json"

    # Uncomment for testing
    # WhatIf                = $true
    # MaxMessagesPerMailbox = 50
    # StartDate             = [datetime]"2024-01-01"
}

# ============================================
# OPTION 2: Single user test mode
# Uncomment this block and comment out OPTION 1
# ============================================
<#
$params = @{
    # Graph API Auth
    TenantId                = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    ClientId                = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    ClientSecret            = "your-secret-here"

    # IMAP Source
    ImapServer              = "mail.kopano.local"
    ImapPort                = 993
    ImapUseSsl              = $true
    ImapSkipCertValidation  = $false

    # Single User Test Mode
    TestMode                = $true
    TestSource              = "testuser@kopano.local"       # Source IMAP mailbox
    TestTarget              = "testuser@company.onmicrosoft.com"  # Target M365 mailbox
    TestUsername            = "testuser"                    # IMAP username (optional, defaults to TestSource)
    TestPassword            = "TestUserPassword123"         # IMAP password

    # Options
    PreserveFolderStructure = $true
    MaxMessagesPerMailbox   = 10                            # Limit for testing
    WhatIf                  = $true                         # Dry run first!

    # Logging
    LogPath                 = ".\migration_logs"
}
#>

# Run migration
& "$PSScriptRoot\Kopano-IMAP-to-Graph-Migration.ps1" @params
