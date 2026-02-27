<#
.SYNOPSIS
    Kopano IMAP to Exchange Online Migration (EWS Version)
    
.DESCRIPTION
    Migrates emails from Kopano IMAP server to Microsoft 365 via EWS.
    Uses direct MIME import for 1:1 email preservation.
    
.NOTES
    Requires: 
    - lib/MailKit.dll, lib/MimeKit.dll, lib/BouncyCastle.Crypto.dll (for IMAP)
    - Azure AD App with MailboxItem.ImportExport.All permission
#>

[CmdletBinding()]
param(
    # === Microsoft Graph/EWS Authentication ===
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
# MailKit Loading (for IMAP source)
# ================================

function Initialize-MailKit {
    $libPath = Join-Path $PSScriptRoot "lib"
    
    $dlls = @(
        "BouncyCastle.Crypto.dll",
        "MimeKit.dll",
        "MailKit.dll"
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
    
    Write-Log "MailKit libraries loaded" -Level Success
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
# EWS Functions
# ================================

function Get-EwsFolderId {
    param(
        [string]$TargetMailbox,
        [string]$FolderPath,
        [string]$Token,
        [hashtable]$FolderCache = @{}
    )

    $cacheKey = "$TargetMailbox|$FolderPath"
    if ($FolderCache.ContainsKey($cacheKey)) {
        return $FolderCache[$cacheKey]
    }

    # Well-known folder mapping
    $wellKnownFolders = @{
        'INBOX'              = 'inbox'
        'Sent'               = 'sentitems'
        'Sent Items'         = 'sentitems'
        'Gesendete Objekte'  = 'sentitems'
        'Drafts'             = 'drafts'
        'Entwürfe'           = 'drafts'
        'Trash'              = 'deleteditems'
        'Deleted Items'      = 'deleteditems'
        'Gelöschte Objekte'  = 'deleteditems'
        'Junk'               = 'junkemail'
        'Junk E-Mail'        = 'junkemail'
        'Archive'            = 'archive'
        'Archiv'             = 'archive'
        'Outbox'             = 'outbox'
        'Postausgang'        = 'outbox'
    }

    $normalizedPath = $FolderPath -replace '/', '\' -replace '\\+', '\'
    $parts = $normalizedPath.Split('\') | Where-Object { $_ -ne '' }
    
    if ($parts.Count -eq 0) {
        return $null
    }

    $rootFolder = $parts[0]
    $parentFolderId = $null

    # Check if root is a well-known folder
    if ($wellKnownFolders.ContainsKey($rootFolder)) {
        $wellKnownName = $wellKnownFolders[$rootFolder]
        
        # Get well-known folder via EWS
        $ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
        
        $getFolderXml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>$TargetMailbox</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
  </soap:Header>
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="$wellKnownName">
          <t:Mailbox>
            <t:EmailAddress>$TargetMailbox</t:EmailAddress>
          </t:Mailbox>
        </t:DistinguishedFolderId>
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>
"@

        try {
            $headers = @{
                "Authorization" = "Bearer $Token"
                "Content-Type"  = "text/xml; charset=utf-8"
            }
            
            $response = Invoke-WebRequest -Uri $ewsUrl -Method POST -Headers $headers -Body $getFolderXml -UseBasicParsing
            $xml = [xml]$response.Content
            
            $ns = @{
                soap = "http://schemas.xmlsoap.org/soap/envelope/"
                m = "http://schemas.microsoft.com/exchange/services/2006/messages"
                t = "http://schemas.microsoft.com/exchange/services/2006/types"
            }
            
            $folderId = $xml.SelectSingleNode("//t:FolderId", (New-XmlNamespaceManager $xml $ns))
            if ($folderId) {
                $parentFolderId = $folderId.GetAttribute("Id")
            }
        }
        catch {
            Write-Log "Failed to get well-known folder $wellKnownName : $_" -Level Warning
        }
    }

    # If we need to create subfolders or the root wasn't well-known
    if ($parts.Count -gt 1 -or !$parentFolderId) {
        # For now, just use inbox as fallback for complex paths
        if (!$parentFolderId) {
            $parentFolderId = Get-EwsFolderId -TargetMailbox $TargetMailbox -FolderPath "INBOX" -Token $Token -FolderCache $FolderCache
        }
        
        # TODO: Create subfolder structure if needed
    }

    $FolderCache[$cacheKey] = $parentFolderId
    return $parentFolderId
}

function New-XmlNamespaceManager {
    param($xml, $namespaces)
    $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    foreach ($key in $namespaces.Keys) {
        $nsmgr.AddNamespace($key, $namespaces[$key])
    }
    return $nsmgr
}

function Import-MessageViaEws {
    param(
        [string]$TargetMailbox,
        [string]$FolderId,
        [byte[]]$MimeContent,
        [datetime]$ReceivedDate,
        [bool]$IsRead = $true
    )

    $token = Get-OAuthToken
    $ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
    
    # Base64 encode the MIME content
    $mimeBase64 = [Convert]::ToBase64String($MimeContent)
    
    # Use msgfolderroot if no specific folder
    $folderIdElement = if ($FolderId) {
        "<t:FolderId Id=`"$FolderId`"/>"
    } else {
        @"
<t:DistinguishedFolderId Id="inbox">
  <t:Mailbox>
    <t:EmailAddress>$TargetMailbox</t:EmailAddress>
  </t:Mailbox>
</t:DistinguishedFolderId>
"@
    }

    # CreateItem with MIME content
    $createItemXml = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:ExchangeImpersonation>
      <t:ConnectingSID>
        <t:SmtpAddress>$TargetMailbox</t:SmtpAddress>
      </t:ConnectingSID>
    </t:ExchangeImpersonation>
  </soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SaveOnly">
      <m:SavedItemFolderId>
        $folderIdElement
      </m:SavedItemFolderId>
      <m:Items>
        <t:Message>
          <t:MimeContent CharacterSet="UTF-8">$mimeBase64</t:MimeContent>
          <t:IsRead>$($IsRead.ToString().ToLower())</t:IsRead>
        </t:Message>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>
"@

    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "text/xml; charset=utf-8"
    }

    try {
        $response = Invoke-WebRequest -Uri $ewsUrl -Method POST -Headers $headers -Body $createItemXml -UseBasicParsing
        
        if ($response.StatusCode -eq 200) {
            $xml = [xml]$response.Content
            
            # Check for success
            if ($response.Content -match 'ResponseClass="Success"') {
                # Extract ItemId
                if ($response.Content -match 'ItemId Id="([^"]+)"') {
                    return $matches[1]
                }
                return "success"
            }
            else {
                # Extract error
                if ($response.Content -match '<m:MessageText>([^<]+)</m:MessageText>') {
                    throw "EWS Error: $($matches[1])"
                }
                throw "EWS Error: Unknown error in response"
            }
        }
    }
    catch {
        Write-Log "EWS Import failed: $_" -Level Error
        throw
    }
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
    $client = $null
    $secureSocket = $null

    # Get OAuth token for EWS
    $ewsToken = Get-OAuthToken

    # Helper function to connect/reconnect IMAP
    function Connect-ImapClient {
        param($client)
        
        if ($client.IsConnected) { return }
        
        Write-Log "Connecting to IMAP..." -Level Debug -User $sourceEmail
        $client.Connect($ImapServer, $ImapPort, $secureSocket)
        $client.Authenticate($imapUsername, $imapPassword)
    }

    try {
        # Create MailKit IMAP client for source
        $client = New-Object MailKit.Net.Imap.ImapClient
        $client.Timeout = 30000

        if ($ImapSkipCertValidation) {
            $client.ServerCertificateValidationCallback = {
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

        Connect-ImapClient $client

        Write-Log "Connected to IMAP source" -Level Success -User $sourceEmail

        # Get all folders
        $personalNamespace = $client.PersonalNamespaces[0]
        $folders = $client.GetFolders($personalNamespace)

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

            # Reconnect if needed
            if (!$client.IsConnected) {
                Write-Log "Reconnecting IMAP..." -Level Warning -User $sourceEmail
                Connect-ImapClient $client
                $personalNamespace = $client.PersonalNamespaces[0]
                $folder = $client.GetFolder($folderName)
            }

            # Open folder
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

            # Search messages
            $query = [MailKit.Search.SearchQuery]::All
            if ($StartDate) {
                $query = $query.And([MailKit.Search.SearchQuery]::DeliveredAfter($StartDate))
            }
            if ($EndDate) {
                $query = $query.And([MailKit.Search.SearchQuery]::DeliveredBefore($EndDate))
            }

            $uids = $folder.Search($query)
            $userStats.TotalMessages += $uids.Count

            Write-Log "Found $($uids.Count) messages" -Level Info -User $sourceEmail

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

            # Get target folder ID via EWS
            $targetFolderId = $null
            if ($PreserveFolderStructure) {
                $targetFolderId = Get-EwsFolderId -TargetMailbox $targetEmail -FolderPath $folderName -Token $ewsToken -FolderCache $FolderCache
            }

            # Process messages
            $msgIndex = 0
            foreach ($uid in $uidsToProcess) {
                $msgIndex++

                try {
                    # Fetch full message
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
                        Write-Log "Empty message, skipping" -Level Warning -User $sourceEmail
                        $userStats.Skipped++
                        continue
                    }

                    if ($WhatIf) {
                        Write-Log "[WHATIF] Would migrate: $subject" -Level Info -User $sourceEmail
                        $userStats.Migrated++
                        continue
                    }

                    # Get date and read status
                    $receivedDate = if ($message.Date.DateTime -ne [DateTime]::MinValue) {
                        $message.Date.DateTime
                    } else {
                        Get-Date
                    }

                    $isRead = $false
                    try {
                        $summary = $folder.Fetch(@($uid), [MailKit.MessageSummaryItems]::Flags)
                        if ($summary -and $summary.Count -gt 0) {
                            $isRead = ($summary[0].Flags -band [MailKit.MessageFlags]::Seen) -eq [MailKit.MessageFlags]::Seen
                        }
                    }
                    catch { }

                    # Import via EWS - MIME stays MIME!
                    $importedId = Import-MessageViaEws `
                        -TargetMailbox $targetEmail `
                        -FolderId $targetFolderId `
                        -MimeContent $mimeBytes `
                        -ReceivedDate $receivedDate `
                        -IsRead $isRead

                    $userStats.Migrated++
                    $script:stats.MigratedMessages++

                    Write-Log "Migrated [$msgIndex/$($uidsToProcess.Count)]: $subject" -Level Success -User $sourceEmail

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

        Write-Log "Complete. Migrated: $($userStats.Migrated), Failed: $($userStats.Failed)" -Level Success -User $sourceEmail
    }
    catch {
        Write-Log "Migration failed: $_" -Level Error -User $sourceEmail
        throw
    }
    finally {
        if ($client -and $client.IsConnected) {
            try { $client.Disconnect($true) } catch { }
            $client.Dispose()
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
    Write-Log "Method: EWS MIME Import (1:1 preservation)" -Level Info

    # Load MailKit for IMAP source
    if (!(Initialize-MailKit)) {
        throw "Failed to load MailKit"
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

    # Test EWS connectivity
    Write-Log "Testing EWS connectivity..." -Level Info
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
