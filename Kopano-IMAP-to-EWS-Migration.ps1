<#
.SYNOPSIS
    Kopano IMAP to Microsoft 365 Migration (EWS Version)
    
.DESCRIPTION
    Migrates emails from Kopano IMAP to M365 using:
    - MailKit for IMAP source (reliable IMAP client)
    - EWS for M365 target (true MIME import, no conversion)
    
    This is the cleanest approach - MIME bytes go directly into Exchange
    without any parsing or JSON conversion.
    
.NOTES
    Requires: 
    - lib/MailKit.dll, lib/MimeKit.dll, lib/BouncyCastle.Crypto.dll
    - lib/Microsoft.Exchange.WebServices.dll
    - Azure AD App with MailboxItem.ImportExport.All (Application permission)
#>

[CmdletBinding()]
param(
    # === Azure AD Authentication ===
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$ClientSecret,

    # === IMAP Source Configuration ===
    [Parameter(Mandatory)]
    [string]$ImapServer,

    [int]$ImapPort = 993,

    [switch]$ImapUseSsl = $true,

    [switch]$ImapSkipCertValidation,

    # === User List (CSV mode) ===
    [string]$UserCsvPath,

    # === Single User Test Mode ===
    [string]$TestSource,
    [string]$TestTarget,
    [string]$TestUsername,
    [string]$TestPassword,
    [switch]$TestMode,

    # === Migration Options ===
    [string[]]$FoldersToMigrate,

    [string[]]$ExcludeFolders = @(
        "Junk", "Spam", "Trash", "Deleted Items",
        "Drafts", "Entwürfe", "Papierkorb", "Junk-E-Mail"
    ),

    [datetime]$StartDate,
    [datetime]$EndDate,
    [int]$MaxMessagesPerMailbox,
    [switch]$PreserveFolderStructure = $true,

    # === Processing Options ===
    [int]$ThrottleMs = 100,
    [int]$MaxRetries = 3,
    [switch]$WhatIf,
    [switch]$ContinueOnError,

    # === Logging ===
    [string]$LogPath = ".\migration_log",
    [switch]$VerboseLogging,

    # === Resume Support ===
    [string]$StateFile,
    [switch]$Resume
)

# ================================
# Initialize
# ================================
$ErrorActionPreference = 'Stop'
$script:accessToken = $null
$script:tokenExpiry = [datetime]::MinValue

$script:stats = @{
    TotalUsers = 0
    ProcessedUsers = 0
    TotalMessages = 0
    MigratedMessages = 0
    SkippedMessages = 0
    FailedMessages = 0
    StartTime = Get-Date
}

# ================================
# Logging
# ================================
$script:logFile = $null

function Initialize-Logging {
    if (!(Test-Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:logFile = Join-Path $LogPath "migration_ews_$timestamp.log"
    
    $header = @"
========================================
Kopano IMAP to M365 Migration (EWS)
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
IMAP Server: ${ImapServer}:${ImapPort}
========================================

"@
    Set-Content -Path $script:logFile -Value $header
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error", "Debug")]
        [string]$Level = "Info",
        [string]$User = ""
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $userPrefix = if ($User) { "[$User] " } else { "" }
    $logMessage = "[$timestamp] [$Level] $userPrefix$Message"

    switch ($Level) {
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error"   { Write-Host $logMessage -ForegroundColor Red }
        "Debug"   { if ($VerboseLogging) { Write-Host $logMessage -ForegroundColor Gray } }
        default   { Write-Host $logMessage -ForegroundColor Cyan }
    }

    if ($script:logFile -and ($Level -ne "Debug" -or $VerboseLogging)) {
        Add-Content -Path $script:logFile -Value $logMessage -ErrorAction SilentlyContinue
    }
}

# ================================
# Load Assemblies
# ================================

function Initialize-Libraries {
    $libPath = Join-Path $PSScriptRoot "lib"
    
    # Required DLLs in load order
    $dlls = @(
        "BouncyCastle.Crypto.dll",
        "MimeKit.dll",
        "MailKit.dll",
        "Microsoft.Exchange.WebServices.dll"
    )
    
    foreach ($dll in $dlls) {
        $dllPath = Join-Path $libPath $dll
        
        if (!(Test-Path $dllPath)) {
            Write-Log "Missing DLL: $dllPath" -Level Error
            return $false
        }
        
        $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($dll)
        $loaded = [System.AppDomain]::CurrentDomain.GetAssemblies() | 
            Where-Object { $_.GetName().Name -eq $assemblyName }
        
        if (!$loaded) {
            try {
                Add-Type -Path $dllPath -ErrorAction Stop
                Write-Log "Loaded: $dll" -Level Debug
            }
            catch [System.Reflection.ReflectionTypeLoadException] {
                Write-Log "Already loaded: $dll" -Level Debug
            }
            catch {
                Write-Log "Failed to load $dll : $_" -Level Error
                return $false
            }
        }
    }
    
    Write-Log "All libraries loaded successfully" -Level Success
    return $true
}

# ================================
# OAuth Token
# ================================

function Get-OAuthToken {
    if ($script:accessToken -and $script:tokenExpiry -gt (Get-Date).AddMinutes(5)) {
        return $script:accessToken
    }

    Write-Log "Acquiring OAuth token..." -Level Info

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://outlook.office365.com/.default"
        grant_type    = "client_credentials"
    }

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body
        $script:accessToken = $response.access_token
        $script:tokenExpiry = (Get-Date).AddSeconds($response.expires_in - 300)
        Write-Log "Token acquired" -Level Success
        return $script:accessToken
    }
    catch {
        Write-Log "Failed to acquire token: $_" -Level Error
        throw
    }
}

# ================================
# EWS Service
# ================================

function Get-EwsService {
    param([string]$TargetMailbox)
    
    $token = Get-OAuthToken
    
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
    
    # OAuth credentials
    $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($token)
    
    # EWS URL
    $service.Url = New-Object Uri("https://outlook.office365.com/EWS/Exchange.asmx")
    
    # Impersonate target mailbox
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
        $TargetMailbox
    )
    
    # Set headers for better compatibility
    $service.HttpHeaders.Add("X-AnchorMailbox", $TargetMailbox)
    
    return $service
}

# ================================
# EWS Folder Management
# ================================

function Get-OrCreateEwsFolder {
    param(
        $Service,
        [string]$FolderPath,
        [hashtable]$FolderCache = @{}
    )
    
    $cacheKey = $FolderPath.ToLower()
    if ($FolderCache.ContainsKey($cacheKey)) {
        return $FolderCache[$cacheKey]
    }
    
    $normalizedPath = $FolderPath -replace '/', '\' -replace '\\+', '\'
    $parts = $normalizedPath.Split('\') | Where-Object { $_ -ne '' }
    
    # Well-known folder mapping
    $wellKnownFolders = @{
        'INBOX'              = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
        'Sent'               = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        'Sent Items'         = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        'Gesendete Objekte'  = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        'Gesendete Elemente' = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        'Drafts'             = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Drafts
        'Entwürfe'           = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Drafts
        'Trash'              = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        'Deleted Items'      = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        'Gelöschte Objekte'  = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        'Gelöschte Elemente' = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        'Junk'               = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail
        'Junk E-Mail'        = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail
        'Junk-E-Mail'        = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail
        'Archive'            = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox  # Fallback
        'Archiv'             = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
        'Postausgang'        = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Outbox
    }
    
    $currentFolderId = $null
    
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $folderName = $parts[$i]
        
        # Check well-known folder at root
        if ($i -eq 0 -and $wellKnownFolders.ContainsKey($folderName)) {
            $wellKnown = $wellKnownFolders[$folderName]
            $currentFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($wellKnown)
            continue
        }
        
        # Search for existing folder
        $parentId = if ($currentFolderId) { $currentFolderId } else {
            New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
        }
        
        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
            $folderName
        )
        
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        
        try {
            $searchResults = $Service.FindFolders($parentId, $searchFilter, $folderView)
            
            if ($searchResults.TotalCount -gt 0) {
                $currentFolderId = $searchResults.Folders[0].Id
                continue
            }
        }
        catch {
            Write-Log "Error searching for folder '$folderName': $_" -Level Debug
        }
        
        # Create folder
        try {
            Write-Log "Creating folder: $folderName" -Level Debug
            $newFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
            $newFolder.DisplayName = $folderName
            $newFolder.Save($parentId)
            $currentFolderId = $newFolder.Id
        }
        catch {
            Write-Log "Failed to create folder '$folderName': $_" -Level Warning
            # Fallback to Inbox
            $currentFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
        }
    }
    
    $FolderCache[$cacheKey] = $currentFolderId
    return $currentFolderId
}

# ================================
# EWS MIME Import
# ================================

function Import-MimeToEws {
    param(
        $Service,
        $TargetFolderId,
        [byte[]]$MimeContent,
        [datetime]$ReceivedDate,
        [bool]$IsRead = $true
    )
    
    # Create email message from MIME - this is the magic!
    # EWS imports the MIME bytes directly, no conversion needed
    $email = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($Service)
    
    # Set MIME content directly
    $email.MimeContent = New-Object Microsoft.Exchange.WebServices.Data.MimeContent("UTF-8", $MimeContent)
    
    # Set as read if needed
    $email.IsRead = $IsRead
    
    # Save to target folder (the MIME already contains all headers, body, attachments)
    $email.Save($TargetFolderId)
    
    return $email.Id.UniqueId
}

# ================================
# CSV Processing
# ================================

function Import-UserCsv {
    param([string]$CsvPath)

    if (!(Test-Path $CsvPath)) {
        throw "User CSV not found: $CsvPath"
    }

    Write-Log "Loading users from: $CsvPath" -Level Info

    $firstLine = Get-Content $CsvPath -First 1
    $delimiter = if ($firstLine -match ';') { ';' } else { ',' }

    $users = Import-Csv -Path $CsvPath -Delimiter $delimiter
    Write-Log "Loaded $($users.Count) users" -Level Success

    return $users
}

# ================================
# IMAP Migration with MailKit + EWS
# ================================

function Test-FolderExcluded {
    param([string]$FolderName)

    foreach ($exclude in $ExcludeFolders) {
        if ($FolderName -ieq $exclude -or $FolderName -ilike "*$exclude*") {
            return $true
        }
    }
    return $false
}

function Migrate-UserMailbox {
    param(
        [hashtable]$User,
        [hashtable]$FolderCache = @{}
    )

    $sourceEmail = $User.Email
    $imapUsername = $User.Username
    $imapPassword = $User.Password
    $targetEmail = if ($User.TargetEmail) { $User.TargetEmail } else { $sourceEmail }

    Write-Log "Starting migration: $sourceEmail -> $targetEmail" -Level Info -User $sourceEmail

    $userStats = @{ TotalMessages = 0; Migrated = 0; Skipped = 0; Failed = 0; Folders = 0 }
    $imapClient = $null
    $ewsService = $null

    try {
        # === Connect to source IMAP (MailKit) ===
        $imapClient = New-Object MailKit.Net.Imap.ImapClient
        $imapClient.Timeout = 30000

        if ($ImapSkipCertValidation) {
            $imapClient.ServerCertificateValidationCallback = {
                param($sender, $certificate, $chain, $sslPolicyErrors)
                return $true
            }
            Write-Log "SSL certificate validation disabled" -Level Warning -User $sourceEmail
        }

        $secureSocket = if ($ImapUseSsl) {
            [MailKit.Security.SecureSocketOptions]::SslOnConnect
        } else {
            [MailKit.Security.SecureSocketOptions]::StartTlsWhenAvailable
        }

        Write-Log "Connecting to IMAP $ImapServer`:$ImapPort..." -Level Info -User $sourceEmail
        $imapClient.Connect($ImapServer, $ImapPort, $secureSocket)
        $imapClient.Authenticate($imapUsername, $imapPassword)
        Write-Log "IMAP connected" -Level Success -User $sourceEmail

        # === Connect to target EWS ===
        Write-Log "Connecting to EWS for $targetEmail..." -Level Info -User $sourceEmail
        $ewsService = Get-EwsService -TargetMailbox $targetEmail
        Write-Log "EWS connected" -Level Success -User $sourceEmail

        # === Get IMAP folders ===
        $personalNamespace = $imapClient.PersonalNamespaces[0]
        $folders = $imapClient.GetFolders($personalNamespace)

        $folderNames = @($folders | ForEach-Object { $_.FullName })
        Write-Log "Found $($folders.Count) folders: $($folderNames -join ', ')" -Level Info -User $sourceEmail

        foreach ($folder in $folders) {
            $folderName = $folder.FullName

            # Skip GUID folders
            if ($folderName -match '^\{[0-9A-Fa-f-]{36}\}$') {
                Write-Log "Skipping GUID folder: $folderName" -Level Debug -User $sourceEmail
                continue
            }

            # Skip system folders
            if ($folderName -in @('Conversation History', 'Sync Issues', 'Conflicts', 'Local Failures', 'Server Failures')) {
                Write-Log "Skipping system folder: $folderName" -Level Debug -User $sourceEmail
                continue
            }

            # Check exclusions
            if (Test-FolderExcluded -FolderName $folderName) {
                Write-Log "Skipping excluded: $folderName" -Level Debug -User $sourceEmail
                continue
            }

            # Filter folders if specified
            if ($FoldersToMigrate -and $FoldersToMigrate.Count -gt 0) {
                $match = $FoldersToMigrate | Where-Object { $folderName -ilike $_ }
                if (!$match) {
                    Write-Log "Skipping (not in filter): $folderName" -Level Debug -User $sourceEmail
                    continue
                }
            }

            # Reconnect IMAP if needed
            if (!$imapClient.IsConnected) {
                Write-Log "Reconnecting IMAP..." -Level Warning -User $sourceEmail
                $imapClient.Connect($ImapServer, $ImapPort, $secureSocket)
                $imapClient.Authenticate($imapUsername, $imapPassword)
                $personalNamespace = $imapClient.PersonalNamespaces[0]
                $folder = $imapClient.GetFolder($folderName)
            }

            # Open IMAP folder
            try {
                $folder.Open([MailKit.FolderAccess]::ReadOnly)
            }
            catch {
                Write-Log "Cannot open folder: $folderName - $_" -Level Warning -User $sourceEmail
                continue
            }

            if ($folder.Count -eq 0) {
                Write-Log "Empty folder: $folderName" -Level Debug -User $sourceEmail
                $folder.Close()
                continue
            }

            Write-Log "Processing: $folderName ($($folder.Count) messages)" -Level Info -User $sourceEmail
            $userStats.Folders++

            # Build search query
            $query = [MailKit.Search.SearchQuery]::All
            if ($StartDate) {
                $query = $query.And([MailKit.Search.SearchQuery]::DeliveredAfter($StartDate))
            }
            if ($EndDate) {
                $query = $query.And([MailKit.Search.SearchQuery]::DeliveredBefore($EndDate))
            }

            # Search messages
            $uids = $folder.Search($query)
            $userStats.TotalMessages += $uids.Count

            Write-Log "Found $($uids.Count) messages matching criteria" -Level Info -User $sourceEmail

            if ($uids.Count -eq 0) {
                $folder.Close()
                continue
            }

            # Limit if specified
            $uidsToProcess = $uids
            if ($MaxMessagesPerMailbox -and $uids.Count -gt $MaxMessagesPerMailbox) {
                $remaining = $MaxMessagesPerMailbox - $userStats.Migrated
                if ($remaining -le 0) {
                    Write-Log "Reached message limit" -Level Warning -User $sourceEmail
                    $folder.Close()
                    break
                }
                $uidsToProcess = $uids | Select-Object -First $remaining
            }

            # Get/create EWS target folder
            $ewsFolderId = $null
            if ($PreserveFolderStructure) {
                $ewsFolderId = Get-OrCreateEwsFolder -Service $ewsService -FolderPath $folderName -FolderCache $FolderCache
            }
            else {
                $ewsFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
            }

            # Process messages
            $msgIndex = 0
            foreach ($uid in $uidsToProcess) {
                $msgIndex++

                try {
                    # Fetch full message with MailKit
                    $message = $folder.GetMessage($uid)

                    # Get MIME bytes directly
                    $memStream = New-Object System.IO.MemoryStream
                    $message.WriteTo($memStream)
                    $mimeBytes = $memStream.ToArray()
                    $memStream.Dispose()

                    $subject = if ($message.Subject) { 
                        if ($message.Subject.Length -gt 50) { $message.Subject.Substring(0,47) + "..." } 
                        else { $message.Subject }
                    } else { "(No Subject)" }

                    Write-Log "Fetched UID $uid : $($mimeBytes.Length) bytes - $subject" -Level Debug -User $sourceEmail

                    if ($mimeBytes.Length -eq 0) {
                        Write-Log "Empty message UID $uid, skipping" -Level Warning -User $sourceEmail
                        $userStats.Skipped++
                        continue
                    }

                    if ($WhatIf) {
                        Write-Log "[WHATIF] Would migrate: $subject" -Level Info -User $sourceEmail
                        $userStats.Migrated++
                        continue
                    }

                    # Get read status
                    $isRead = $false
                    try {
                        $summary = $folder.Fetch(@($uid), [MailKit.MessageSummaryItems]::Flags)
                        if ($summary -and $summary.Count -gt 0) {
                            $isRead = ($summary[0].Flags -band [MailKit.MessageFlags]::Seen) -eq [MailKit.MessageFlags]::Seen
                        }
                    }
                    catch { }

                    # Get date
                    $receivedDate = if ($message.Date.DateTime -ne [DateTime]::MinValue) {
                        $message.Date.DateTime
                    } else {
                        Get-Date
                    }

                    # === THE MAGIC: Direct MIME import via EWS ===
                    $importedId = Import-MimeToEws `
                        -Service $ewsService `
                        -TargetFolderId $ewsFolderId `
                        -MimeContent $mimeBytes `
                        -ReceivedDate $receivedDate `
                        -IsRead $isRead

                    $userStats.Migrated++
                    $script:stats.MigratedMessages++

                    Write-Log "Migrated [$msgIndex/$($uidsToProcess.Count)]: $subject" -Level Success -User $sourceEmail

                    # Throttle
                    if ($ThrottleMs -gt 0) {
                        Start-Sleep -Milliseconds $ThrottleMs
                    }
                }
                catch {
                    $userStats.Failed++
                    $script:stats.FailedMessages++
                    Write-Log "Failed UID $uid : $_" -Level Error -User $sourceEmail

                    if (!$ContinueOnError) { throw }
                }
            }

            $folder.Close()
        }

        Write-Log "Complete. Migrated: $($userStats.Migrated), Failed: $($userStats.Failed), Skipped: $($userStats.Skipped)" -Level Success -User $sourceEmail
    }
    catch {
        Write-Log "Migration failed: $_" -Level Error -User $sourceEmail
        throw
    }
    finally {
        if ($imapClient -and $imapClient.IsConnected) {
            try { $imapClient.Disconnect($true) } catch { }
            $imapClient.Dispose()
        }
    }

    return $userStats
}

# ================================
# State Management
# ================================

function Save-MigrationState {
    param([string]$StatePath, [hashtable]$State)
    $State | ConvertTo-Json -Depth 10 | Set-Content -Path $StatePath
}

function Load-MigrationState {
    param([string]$StatePath)
    if (!(Test-Path $StatePath)) { return $null }
    return Get-Content $StatePath -Raw | ConvertFrom-Json
}

# ================================
# Main
# ================================

try {
    Initialize-Logging

    Write-Log "=== Kopano IMAP to M365 Migration (EWS) ===" -Level Info
    Write-Log "IMAP Server: ${ImapServer}:${ImapPort}" -Level Info
    Write-Log "Tenant: $TenantId" -Level Info
    Write-Log "Method: EWS MIME Import (cleanest approach)" -Level Info

    # Load libraries
    if (!(Initialize-Libraries)) {
        throw "Failed to load libraries"
    }

    # Validate parameters
    if ($TestMode -or $TestSource -or $TestTarget -or $TestPassword) {
        if (!$TestSource) { throw "TestSource required" }
        if (!$TestTarget) { throw "TestTarget required" }
        if (!$TestPassword) { throw "TestPassword required" }
        $TestMode = $true
        Write-Log "*** TEST MODE ***" -Level Warning
    }
    elseif (!$UserCsvPath) {
        throw "Either -UserCsvPath or test mode parameters required"
    }

    if ($WhatIf) {
        Write-Log "*** WHATIF MODE ***" -Level Warning
    }

    # Test OAuth
    Write-Log "Testing OAuth..." -Level Info
    $null = Get-OAuthToken

    # Load users
    $users = @()
    if ($TestMode) {
        $users = @([PSCustomObject]@{
            Email       = $TestSource
            Username    = if ($TestUsername) { $TestUsername } else { $TestSource }
            Password    = $TestPassword
            TargetEmail = $TestTarget
        })
    }
    else {
        $users = Import-UserCsv -CsvPath $UserCsvPath
    }
    $script:stats.TotalUsers = $users.Count

    # Resume support
    $processedUsers = @{}
    if ($Resume -and $StateFile -and (Test-Path $StateFile)) {
        $previousState = Load-MigrationState -StatePath $StateFile
        if ($previousState) {
            foreach ($u in $previousState.ProcessedUsers) {
                $processedUsers[$u] = $true
            }
            Write-Log "Resuming. Already processed: $($processedUsers.Count) users" -Level Info
        }
    }

    # Folder cache
    $folderCache = @{}

    # Process users
    $userIndex = 0
    foreach ($user in $users) {
        $userIndex++

        $userHash = @{
            Email       = $user.Email
            Username    = if ($user.Username) { $user.Username } else { $user.Email }
            Password    = $user.Password
            TargetEmail = if ($user.TargetEmail) { $user.TargetEmail } else { $user.Email }
        }

        if ($processedUsers.ContainsKey($userHash.Email)) {
            Write-Log "Skipping already processed: $($userHash.Email)" -Level Info
            continue
        }

        Write-Log "`n========================================" -Level Info
        Write-Log "User $userIndex of $($users.Count): $($userHash.Email)" -Level Info
        Write-Log "========================================" -Level Info

        try {
            $userStats = Migrate-UserMailbox -User $userHash -FolderCache $folderCache
            $script:stats.ProcessedUsers++
            $processedUsers[$userHash.Email] = $true

            if ($StateFile) {
                Save-MigrationState -StatePath $StateFile -State @{
                    ProcessedUsers = $processedUsers.Keys
                    LastProcessed  = $userHash.Email
                    Timestamp      = Get-Date -Format 'o'
                    Stats          = $script:stats
                }
            }
        }
        catch {
            Write-Log "Failed: $($userHash.Email) - $_" -Level Error
            if (!$ContinueOnError) { throw }
        }
    }

    # Summary
    $duration = (Get-Date) - $script:stats.StartTime

    Write-Log "`n========================================" -Level Info
    Write-Log "Migration Complete" -Level Success
    Write-Log "========================================" -Level Info
    Write-Log "Duration: $($duration.ToString('hh\:mm\:ss'))" -Level Info
    Write-Log "Users: $($script:stats.ProcessedUsers) / $($script:stats.TotalUsers)" -Level Info
    Write-Log "Messages migrated: $($script:stats.MigratedMessages)" -Level Info
    Write-Log "Messages failed: $($script:stats.FailedMessages)" -Level Info
    Write-Log "Log: $script:logFile" -Level Info
}
catch {
    Write-Log "FATAL: $_" -Level Error
    Write-Log "Stack: $($_.ScriptStackTrace)" -Level Error
    throw
}
