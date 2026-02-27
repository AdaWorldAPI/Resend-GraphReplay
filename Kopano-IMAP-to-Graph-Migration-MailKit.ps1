<#
.SYNOPSIS
    Kopano IMAP to Microsoft Graph Migration (MailKit Version)
    
.DESCRIPTION
    Migrates emails from Kopano IMAP server to Microsoft 365 via Graph API.
    Uses MailKit/MimeKit for reliable IMAP operations.
    
.NOTES
    Requires: lib/MailKit.dll, lib/MimeKit.dll, lib/BouncyCastle.Crypto.dll
    Run Setup-MailKit.ps1 first to download dependencies.
#>

[CmdletBinding()]
param(
    # === Microsoft Graph Authentication ===
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
    [int]$ThrottleMs = 200,
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
    $script:logFile = Join-Path $LogPath "migration_$timestamp.log"
    
    $header = @"
========================================
Kopano IMAP to Graph Migration (MailKit)
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
# MailKit Loading
# ================================

function Initialize-MailKit {
    $libPath = Join-Path $PSScriptRoot "lib"
    
    # Required DLLs in load order
    $dlls = @(
        "BouncyCastle.Crypto.dll",
        "MimeKit.dll",
        "MailKit.dll"
    )
    
    foreach ($dll in $dlls) {
        $dllPath = Join-Path $libPath $dll
        
        if (!(Test-Path $dllPath)) {
            Write-Log "Missing DLL: $dllPath" -Level Error
            Write-Log "Run Setup-MailKit.ps1 first to download dependencies." -Level Error
            return $false
        }
        
        # Check if already loaded
        $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($dll)
        $loaded = [System.AppDomain]::CurrentDomain.GetAssemblies() | 
            Where-Object { $_.GetName().Name -eq $assemblyName }
        
        if (!$loaded) {
            try {
                Add-Type -Path $dllPath -ErrorAction Stop
                Write-Log "Loaded: $dll" -Level Debug
            }
            catch [System.Reflection.ReflectionTypeLoadException] {
                # Already loaded via dependency, ignore
                Write-Log "Already loaded: $dll" -Level Debug
            }
            catch {
                Write-Log "Failed to load $dll : $_" -Level Error
                return $false
            }
        }
    }
    
    Write-Log "MailKit libraries loaded successfully" -Level Success
    return $true
}

# ================================
# Graph API Functions
# ================================

function Get-GraphToken {
    if ($script:accessToken -and $script:tokenExpiry -gt (Get-Date).AddMinutes(5)) {
        return $script:accessToken
    }

    Write-Log "Acquiring Graph API token..." -Level Info

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
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

function Invoke-GraphRequest {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [object]$Body,
        [hashtable]$Headers = @{},
        [string]$ContentType = "application/json",
        [int]$RetryCount = 0
    )

    $token = Get-GraphToken
    $Headers["Authorization"] = "Bearer $token"

    $params = @{
        Method  = $Method
        Uri     = $Uri
        Headers = $Headers
    }

    if ($Body) {
        if ($ContentType -eq "application/json" -and $Body -isnot [string]) {
            $jsonString = $Body | ConvertTo-Json -Depth 20 -Compress
            $params.Body = [System.Text.Encoding]::UTF8.GetBytes($jsonString)
            $params.ContentType = "application/json; charset=utf-8"
        }
        else {
            $params.Body = $Body
            $params.ContentType = $ContentType
        }
    }

    try {
        return Invoke-RestMethod @params
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__

        if ($statusCode -eq 429) {
            $retryAfter = $_.Exception.Response.Headers["Retry-After"]
            $waitTime = if ($retryAfter) { [int]$retryAfter } else { 60 }
            Write-Log "Throttled. Waiting $waitTime seconds..." -Level Warning
            Start-Sleep -Seconds $waitTime
            return Invoke-GraphRequest @PSBoundParameters -RetryCount ($RetryCount + 1)
        }

        if ($RetryCount -lt $MaxRetries -and $statusCode -in @(500, 502, 503, 504)) {
            $waitTime = [math]::Pow(2, $RetryCount) * 2
            Write-Log "Transient error ($statusCode). Retrying in $waitTime seconds..." -Level Warning
            Start-Sleep -Seconds $waitTime
            return Invoke-GraphRequest @PSBoundParameters -RetryCount ($RetryCount + 1)
        }

        throw
    }
}

function Get-OrCreateMailFolder {
    param(
        [string]$TargetMailbox,
        [string]$FolderPath,
        [hashtable]$FolderCache = @{}
    )

    $cacheKey = "$TargetMailbox|$FolderPath"
    if ($FolderCache.ContainsKey($cacheKey)) {
        return $FolderCache[$cacheKey]
    }

    $normalizedPath = $FolderPath -replace '/', '\' -replace '\\+', '\'
    $parts = $normalizedPath.Split('\') | Where-Object { $_ -ne '' }

    # Well-known folder mapping
    $folderMapping = @{
        'INBOX'              = 'Inbox'
        'Sent'               = 'SentItems'
        'Sent Items'         = 'SentItems'
        'Gesendete Elemente' = 'SentItems'
        'Drafts'             = 'Drafts'
        'Entwürfe'           = 'Drafts'
        'Trash'              = 'DeletedItems'
        'Deleted Items'      = 'DeletedItems'
        'Gelöschte Elemente' = 'DeletedItems'
        'Junk'               = 'JunkEmail'
        'Spam'               = 'JunkEmail'
        'Junk-E-Mail'        = 'JunkEmail'
        'Archive'            = 'Archive'
        'Archiv'             = 'Archive'
    }

    $currentParentId = $null
    $currentFolderId = $null

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $folderName = $parts[$i]

        # Check well-known folder at root
        if ($i -eq 0 -and $folderMapping.ContainsKey($folderName)) {
            $wellKnownName = $folderMapping[$folderName]
            $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$wellKnownName"
            try {
                $folder = Invoke-GraphRequest -Uri $uri
                $currentFolderId = $folder.id
                $currentParentId = $folder.id
                continue
            }
            catch { }
        }

        # Search for existing folder
        $searchUri = if ($currentParentId) {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$currentParentId/childFolders?`$filter=displayName eq '$folderName'"
        }
        else {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders?`$filter=displayName eq '$folderName'"
        }

        try {
            $existingFolders = Invoke-GraphRequest -Uri $searchUri
            if ($existingFolders.value -and $existingFolders.value.Count -gt 0) {
                $currentFolderId = $existingFolders.value[0].id
                $currentParentId = $currentFolderId
                continue
            }
        }
        catch { }

        # Create folder
        $createUri = if ($currentParentId) {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$currentParentId/childFolders"
        }
        else {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders"
        }

        try {
            Write-Log "Creating folder: $folderName" -Level Debug
            $created = Invoke-GraphRequest -Uri $createUri -Method POST -Body @{ displayName = $folderName }
            $currentFolderId = $created.id
            $currentParentId = $currentFolderId
        }
        catch {
            Write-Log "Failed to create folder '$folderName': $_" -Level Warning
            $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/Inbox"
            $inbox = Invoke-GraphRequest -Uri $uri
            $currentFolderId = $inbox.id
        }
    }

    $FolderCache[$cacheKey] = $currentFolderId
    return $currentFolderId
}

function Import-MessageToGraph {
    param(
        [string]$TargetMailbox,
        [string]$FolderId,
        [byte[]]$MimeContent,
        [datetime]$ReceivedDate,
        [bool]$IsRead = $true
    )

    $token = Get-GraphToken
    $createUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages"

    # Method 1: Direct MIME import
    try {
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type"  = "text/plain"
        }

        $mimeString = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetString($MimeContent)
        $response = Invoke-WebRequest -Method POST -Uri $createUri -Headers $headers -Body $mimeString -UseBasicParsing

        if ($response.StatusCode -in @(200, 201)) {
            $createdMessage = $response.Content | ConvertFrom-Json
            $messageId = $createdMessage.id

            # Move to target folder
            if ($FolderId) {
                try {
                    $moveUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId/move"
                    $movedMessage = Invoke-GraphRequest -Uri $moveUri -Method POST -Body @{ destinationId = $FolderId }
                    $messageId = $movedMessage.id
                }
                catch {
                    Write-Log "Failed to move message: $_" -Level Warning
                }
            }

            # Set date and flags
            Set-MessageDateAndFlags -TargetMailbox $TargetMailbox -MessageId $messageId -ReceivedDate $ReceivedDate -IsRead $IsRead
            return $messageId
        }
    }
    catch {
        Write-Log "MIME import failed: $($_.Exception.Message)" -Level Debug
    }

    # Method 2: Fallback with HttpClient
    try {
        Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue

        $httpClient = New-Object System.Net.Http.HttpClient
        $httpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $token)
        $httpClient.Timeout = [TimeSpan]::FromMinutes(5)

        $content = New-Object System.Net.Http.ByteArrayContent(,$MimeContent)
        $content.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("text/plain")

        $response = $httpClient.PostAsync($createUri, $content).Result

        if ($response.IsSuccessStatusCode) {
            $responseContent = $response.Content.ReadAsStringAsync().Result
            $createdMessage = $responseContent | ConvertFrom-Json
            $messageId = $createdMessage.id

            if ($FolderId) {
                try {
                    $moveUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId/move"
                    $movedMessage = Invoke-GraphRequest -Uri $moveUri -Method POST -Body @{ destinationId = $FolderId }
                    $messageId = $movedMessage.id
                }
                catch { }
            }

            Set-MessageDateAndFlags -TargetMailbox $TargetMailbox -MessageId $messageId -ReceivedDate $ReceivedDate -IsRead $IsRead
            $httpClient.Dispose()
            return $messageId
        }

        $httpClient.Dispose()
    }
    catch {
        Write-Log "HttpClient import failed: $_" -Level Debug
    }

    throw "All import methods failed for message"
}

function Set-MessageDateAndFlags {
    param(
        [string]$TargetMailbox,
        [string]$MessageId,
        [datetime]$ReceivedDate,
        [bool]$IsRead = $true
    )

    $updateUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$MessageId"
    $dateString = $ReceivedDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $updateBody = @{
        isRead = $IsRead
        singleValueExtendedProperties = @(
            @{ id = "Integer 0x0E07"; value = "1" }
            @{ id = "SystemTime 0x0E06"; value = $dateString }
            @{ id = "SystemTime 0x0039"; value = $dateString }
        )
    }

    try {
        Invoke-GraphRequest -Uri $updateUri -Method PATCH -Body $updateBody | Out-Null
    }
    catch {
        Write-Log "Failed to set message flags: $_" -Level Warning
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
# IMAP Migration with MailKit
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

    try {
        # Create MailKit IMAP client
        $client = New-Object MailKit.Net.Imap.ImapClient

        # Certificate validation callback for self-signed certs
        if ($ImapSkipCertValidation) {
            $client.ServerCertificateValidationCallback = {
                param($sender, $certificate, $chain, $sslPolicyErrors)
                return $true
            }
            Write-Log "SSL certificate validation disabled" -Level Warning -User $sourceEmail
        }

        # Connect
        $secureSocket = if ($ImapUseSsl) {
            [MailKit.Security.SecureSocketOptions]::SslOnConnect
        } else {
            [MailKit.Security.SecureSocketOptions]::StartTlsWhenAvailable
        }

        Write-Log "Connecting to $ImapServer`:$ImapPort..." -Level Info -User $sourceEmail
        $client.Connect($ImapServer, $ImapPort, $secureSocket)

        # Authenticate
        Write-Log "Authenticating..." -Level Debug -User $sourceEmail
        $client.Authenticate($imapUsername, $imapPassword)

        Write-Log "Connected and authenticated via MailKit" -Level Success -User $sourceEmail

        # Get all folders
        $personalNamespace = $client.PersonalNamespaces[0]
        $folders = $client.GetFolders($personalNamespace)

        Write-Log "Found $($folders.Count) folders" -Level Info -User $sourceEmail

        foreach ($folder in $folders) {
            $folderName = $folder.FullName

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

            # Open folder read-only
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

            # Get/create target folder
            $targetFolderId = $null
            if ($PreserveFolderStructure) {
                $targetFolderId = Get-OrCreateMailFolder -TargetMailbox $targetEmail -FolderPath $folderName -FolderCache $FolderCache
            }
            else {
                $targetFolderId = Get-OrCreateMailFolder -TargetMailbox $targetEmail -FolderPath "Inbox" -FolderCache $FolderCache
            }

            # Process messages
            $msgIndex = 0
            foreach ($uid in $uidsToProcess) {
                $msgIndex++

                try {
                    # Fetch full message with MailKit
                    $message = $folder.GetMessage($uid)

                    # Get MIME bytes
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

                    # Get received date
                    $receivedDate = if ($message.Date.DateTime -ne [DateTime]::MinValue) {
                        $message.Date.DateTime
                    } else {
                        Get-Date
                    }

                    # Get read status via summary
                    $isRead = $false
                    try {
                        $summary = $folder.Fetch(@($uid), [MailKit.MessageSummaryItems]::Flags)
                        if ($summary -and $summary.Count -gt 0) {
                            $isRead = ($summary[0].Flags -band [MailKit.MessageFlags]::Seen) -eq [MailKit.MessageFlags]::Seen
                        }
                    }
                    catch { }

                    # Import to Graph
                    $importedId = Import-MessageToGraph `
                        -TargetMailbox $targetEmail `
                        -FolderId $targetFolderId `
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
        if ($client -and $client.IsConnected) {
            try {
                $client.Disconnect($true)
            }
            catch { }
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

    Write-Log "=== Kopano IMAP to Graph Migration (MailKit) ===" -Level Info
    Write-Log "IMAP Server: ${ImapServer}:${ImapPort} (SSL: $ImapUseSsl)" -Level Info
    Write-Log "Tenant: $TenantId" -Level Info

    # Load MailKit
    if (!(Initialize-MailKit)) {
        throw "Failed to load MailKit. Run Setup-MailKit.ps1 first."
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
        Write-Log "*** WHATIF MODE - No actual migration ***" -Level Warning
    }

    # Test Graph API
    Write-Log "Testing Graph API..." -Level Info
    $null = Get-GraphToken

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

        # Normalize
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
