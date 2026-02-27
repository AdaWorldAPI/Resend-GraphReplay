<#
Kopano-IMAP-to-Graph-Migration.ps1 — IMAP to Microsoft Graph Email Migration
Migrates emails from Kopano IMAP server to Microsoft 365 via Graph API

Uses MailKit/MimeKit for reliable IMAP operations. Run Setup-MailKit.ps1 first.

Features:
  - Bulk migration from CSV user list
  - Preserves original sent/received dates
  - Maintains folder structure
  - Supports SSL/TLS IMAP connections
  - Progress tracking and detailed logging
  - Diagnostic mode for troubleshooting
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
    [string]$ImapServer,                    # IMAP server hostname (e.g., mail.kopano.local)

    [int]$ImapPort = 993,                   # IMAP port (993 for SSL, 143 for STARTTLS)

    [switch]$ImapUseSsl = $true,            # Use SSL/TLS connection

    [switch]$ImapSkipCertValidation,        # Skip SSL certificate validation (for self-signed certs)

    # === User List (CSV mode) ===
    [string]$UserCsvPath,                   # CSV file with: Email,Username,Password,TargetEmail (optional)

    # === Single User Test Mode ===
    [string]$TestSource,                    # Test: Source IMAP email/username
    [string]$TestTarget,                    # Test: Target M365 mailbox
    [string]$TestUsername,                  # Test: IMAP username (if different from TestSource)
    [string]$TestPassword,                  # Test: IMAP password
    [switch]$TestMode,                      # Enable single-user test mode

    # === Migration Options ===
    [string[]]$FoldersToMigrate,            # Specific folders to migrate (empty = all folders)

    [string[]]$ExcludeFolders = @(          # Folders to exclude
        "Junk", "Spam", "Trash", "Deleted Items",
        "Drafts", "Entwürfe", "Papierkorb"
    ),

    [switch]$IncludeSubfolders = $true,     # Include subfolders

    [datetime]$StartDate,                   # Only migrate emails after this date

    [datetime]$EndDate,                     # Only migrate emails before this date

    [int]$MaxMessagesPerMailbox,            # Limit messages per mailbox (for testing)

    [switch]$PreserveFolderStructure = $true, # Create matching folder structure in target

    # === Date Handling ===
    [switch]$PreserveReceivedDate = $true,  # Preserve original received date (default: true)

    # === Processing Options ===
    [int]$BatchSize = 25,                   # Messages to process in batch

    [int]$ThrottleMs = 200,                 # Delay between API calls (ms)

    [int]$MaxRetries = 3,                   # Max retries on failure

    [switch]$WhatIf,                        # Dry run - no actual migration

    [switch]$ContinueOnError,               # Continue processing on errors

    # === Logging ===
    [string]$LogPath = ".\migration_log",   # Log directory

    [switch]$VerboseLogging,                # Enable verbose logging

    # === Diagnostic Options ===
    [switch]$DiagnosticMode,                # Enable full diagnostic logging (IMAP commands, byte counts, API details)

    [switch]$TestSingleMessage,             # Fetch and display ONE message for debugging, then stop

    [switch]$SaveMimeToFile,                # Save fetched MIME content to disk before importing

    [switch]$SkipGraphImport,               # Test IMAP fetch only, skip Graph API import

    [string]$MimeSavePath = ".\mime_dump",  # Directory to save MIME files when -SaveMimeToFile is used

    # === Resume Support ===
    [string]$StateFile,                     # State file for resume capability

    [switch]$Resume                         # Resume from previous state
)

# ================================
# Initialize
# ================================
$ErrorActionPreference = 'Stop'
$script:accessToken = $null
$script:tokenExpiry = [datetime]::MinValue

# Enable verbose logging in diagnostic mode
if ($DiagnosticMode) {
    $VerboseLogging = $true
}

# Statistics
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
# MailKit Assembly Loading
# ================================

function Initialize-MailKit {
    <#
    .SYNOPSIS
    Loads MailKit/MimeKit libraries for IMAP operations.
    Run Setup-MailKit.ps1 first to download the required DLLs.
    #>

    $libPath = "$PSScriptRoot\lib"

    # Load in dependency order
    $dlls = @(
        "BouncyCastle.Crypto.dll",
        "MimeKit.dll",
        "MailKit.dll"
    )

    foreach ($dll in $dlls) {
        $dllPath = Join-Path $libPath $dll

        if (Test-Path $dllPath) {
            try {
                # Check if already loaded
                $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($dll)
                $alreadyLoaded = [System.AppDomain]::CurrentDomain.GetAssemblies() |
                    Where-Object { $_.GetName().Name -eq $assemblyName }

                if ($alreadyLoaded) {
                    Write-Log "Already loaded: $dll" -Level Debug
                    continue
                }

                Add-Type -Path $dllPath -ErrorAction Stop
                Write-Log "Loaded: $dll" -Level Debug
            }
            catch [System.Reflection.ReflectionTypeLoadException] {
                # Already loaded, ignore
                Write-Log "Assembly $dll already loaded (ReflectionTypeLoadException)" -Level Debug
            }
            catch {
                Write-Log "Warning loading $dll : $_ (may already be loaded)" -Level Debug
            }
        }
        else {
            throw "Required DLL not found: $dllPath. Run Setup-MailKit.ps1 first to download dependencies."
        }
    }

    Write-Log "MailKit/MimeKit loaded successfully" -Level Info
    return $true
}

function Get-MailKitClient {
    <#
    .SYNOPSIS
    Creates and connects a MailKit IMAP client.
    #>
    param(
        [string]$Server,
        [int]$Port,
        [bool]$UseSsl,
        [bool]$SkipCertValidation,
        [string]$Username,
        [string]$Password
    )

    $client = New-Object MailKit.Net.Imap.ImapClient

    # Set timeout for reliability
    $client.Timeout = 120000  # 2 minutes

    # Skip certificate validation if requested (for self-signed certs)
    if ($SkipCertValidation) {
        $client.ServerCertificateValidationCallback = {
            param($sender, $certificate, $chain, $sslPolicyErrors)
            return $true
        }
    }

    # Determine connection security
    $secureSocket = if ($UseSsl) {
        [MailKit.Security.SecureSocketOptions]::SslOnConnect
    } else {
        [MailKit.Security.SecureSocketOptions]::StartTlsWhenAvailable
    }

    Write-Log "Connecting to IMAP: $Server`:$Port (SSL: $UseSsl, SkipCert: $SkipCertValidation)" -Level Debug

    $client.Connect($Server, $Port, $secureSocket)

    Write-Log "Connected. Authenticating as: $Username" -Level Debug

    $client.Authenticate($Username, $Password)

    Write-Log "IMAP authentication successful" -Level Debug

    return $client
}

# ================================
# Logging Functions
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
Kopano IMAP to Graph Migration Log
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
IMAP Server: $ImapServer`:$ImapPort
Engine: MailKit/MimeKit
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

    # Console output with color
    switch ($Level) {
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error"   { Write-Host $logMessage -ForegroundColor Red }
        "Debug"   { if ($VerboseLogging) { Write-Host $logMessage -ForegroundColor Gray } }
        default   { Write-Host $logMessage -ForegroundColor Cyan }
    }

    # File logging
    if ($script:logFile -and ($Level -ne "Debug" -or $VerboseLogging)) {
        Add-Content -Path $script:logFile -Value $logMessage -ErrorAction SilentlyContinue
    }
}

# ================================
# Graph API Functions
# ================================

function Get-GraphToken {
    if ($script:accessToken -and $script:tokenExpiry -gt (Get-Date).AddMinutes(5)) {
        return $script:accessToken
    }

    Write-Log "Acquiring new Graph API token..." -Level Info

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
        Write-Log "Token acquired successfully" -Level Success
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
            # Convert to JSON and encode as UTF-8 bytes
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

        # Handle throttling (429)
        if ($statusCode -eq 429) {
            $retryAfter = $_.Exception.Response.Headers["Retry-After"]
            $waitTime = if ($retryAfter) { [int]$retryAfter } else { 60 }
            Write-Log "Throttled. Waiting $waitTime seconds..." -Level Warning
            Start-Sleep -Seconds $waitTime
            return Invoke-GraphRequest @PSBoundParameters -RetryCount ($RetryCount + 1)
        }

        # Retry on transient errors
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

    # Check cache first
    $cacheKey = "$TargetMailbox|$FolderPath"
    if ($FolderCache.ContainsKey($cacheKey)) {
        return $FolderCache[$cacheKey]
    }

    # Normalize folder path — handle both / and . delimiters
    $normalizedPath = $FolderPath -replace '/', '\' -replace '\.', '\' -replace '\\+', '\'
    $parts = $normalizedPath.Split('\') | Where-Object { $_ -ne '' }

    # Map common folder names to well-known Graph folder IDs
    $folderMapping = @{
        'INBOX'          = 'Inbox'
        'Sent'           = 'SentItems'
        'Sent Items'     = 'SentItems'
        'Gesendete Elemente' = 'SentItems'
        'Drafts'         = 'Drafts'
        'Entwürfe'       = 'Drafts'
        'Trash'          = 'DeletedItems'
        'Deleted Items'  = 'DeletedItems'
        'Gelöschte Elemente' = 'DeletedItems'
        'Junk'           = 'JunkEmail'
        'Spam'           = 'JunkEmail'
        'Junk-E-Mail'    = 'JunkEmail'
        'Archive'        = 'Archive'
        'Archiv'         = 'Archive'
    }

    $currentParentId = $null
    $currentFolderId = $null

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $folderName = $parts[$i]

        # Check for well-known folder at root level
        if ($i -eq 0 -and $folderMapping.ContainsKey($folderName)) {
            $wellKnownName = $folderMapping[$folderName]
            $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$wellKnownName"

            try {
                $folder = Invoke-GraphRequest -Uri $uri
                $currentFolderId = $folder.id
                $currentParentId = $folder.id
                continue
            }
            catch {
                # Well-known folder not found, will create it
            }
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
        catch {
            Write-Log "Error searching for folder '$folderName': $_" -Level Debug
        }

        # Create folder if it doesn't exist
        $createUri = if ($currentParentId) {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$currentParentId/childFolders"
        }
        else {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders"
        }

        $newFolder = @{
            displayName = $folderName
        }

        try {
            Write-Log "Creating folder: $folderName" -Level Debug
            $created = Invoke-GraphRequest -Uri $createUri -Method POST -Body $newFolder
            $currentFolderId = $created.id
            $currentParentId = $currentFolderId
        }
        catch {
            Write-Log "Failed to create folder '$folderName': $_" -Level Warning
            # Try to get Inbox as fallback
            $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/Inbox"
            $inbox = Invoke-GraphRequest -Uri $uri
            $currentFolderId = $inbox.id
        }
    }

    # Cache result
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
    $messageId = $null
    $lastError = $null

    if ($DiagnosticMode) {
        Write-Log "DIAG: Import-MessageToGraph called. MimeContent size: $($MimeContent.Length) bytes, FolderId: $FolderId" -Level Debug
    }

    # Method 1: Direct MIME import via Invoke-WebRequest
    try {
        $headers = @{
            "Authorization" = "Bearer $token"
            "Content-Type"  = "text/plain"
        }

        # Convert bytes to string preserving all byte values (ISO-8859-1 is 1:1 mapping for 0-255)
        $mimeString = [System.Text.Encoding]::GetEncoding("ISO-8859-1").GetString($MimeContent)

        if ($DiagnosticMode) {
            Write-Log "DIAG: Method 1 - Sending MIME string ($($mimeString.Length) chars) to $createUri" -Level Debug
        }

        $response = Invoke-WebRequest -Method POST -Uri $createUri -Headers $headers -Body $mimeString -UseBasicParsing

        if ($response.StatusCode -in @(200, 201)) {
            $createdMessage = $response.Content | ConvertFrom-Json
            $messageId = $createdMessage.id

            if ($DiagnosticMode) {
                Write-Log "DIAG: Method 1 SUCCESS - Message ID: $messageId" -Level Debug
            }

            # Move to target folder if not default
            if ($FolderId) {
                try {
                    $moveUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId/move"
                    $moveBody = @{ destinationId = $FolderId }
                    $movedMessage = Invoke-GraphRequest -Uri $moveUri -Method POST -Body $moveBody
                    $messageId = $movedMessage.id
                }
                catch {
                    Write-Log "Failed to move message to target folder: $_" -Level Warning
                }
            }

            # Update message with original date and mark as NOT draft
            Set-MessageDateAndFlags -TargetMailbox $TargetMailbox -MessageId $messageId -ReceivedDate $ReceivedDate -IsRead $IsRead

            return $messageId
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        $errorDetails = ""
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorDetails = $reader.ReadToEnd()
                $reader.Close()
            } catch {}
        }
        $lastError = $_
        Write-Log "MIME import method 1 failed: $errorMsg $errorDetails" -Level Warning
    }

    # Method 2: Try with HttpClient and raw bytes
    try {
        Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue

        $httpClient = New-Object System.Net.Http.HttpClient
        $httpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $token)
        $httpClient.Timeout = [TimeSpan]::FromMinutes(5)

        $content = New-Object System.Net.Http.ByteArrayContent(,$MimeContent)
        $content.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("text/plain")

        if ($DiagnosticMode) {
            Write-Log "DIAG: Method 2 - Sending raw bytes ($($MimeContent.Length) bytes) via HttpClient" -Level Debug
        }

        $response = $httpClient.PostAsync($createUri, $content).Result

        if ($response.IsSuccessStatusCode) {
            $responseContent = $response.Content.ReadAsStringAsync().Result
            $createdMessage = $responseContent | ConvertFrom-Json
            $messageId = $createdMessage.id

            if ($DiagnosticMode) {
                Write-Log "DIAG: Method 2 SUCCESS - Message ID: $messageId" -Level Debug
            }

            # Move to target folder if specified
            if ($FolderId) {
                try {
                    $moveUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId/move"
                    $moveBody = @{ destinationId = $FolderId }
                    $movedMessage = Invoke-GraphRequest -Uri $moveUri -Method POST -Body $moveBody
                    $messageId = $movedMessage.id
                }
                catch {
                    Write-Log "Failed to move message to target folder: $_" -Level Warning
                }
            }

            # Update message with original date and mark as NOT draft
            Set-MessageDateAndFlags -TargetMailbox $TargetMailbox -MessageId $messageId -ReceivedDate $ReceivedDate -IsRead $IsRead

            $httpClient.Dispose()
            return $messageId
        }
        else {
            $errorContent = $response.Content.ReadAsStringAsync().Result
            $lastError = "HTTP $($response.StatusCode): $errorContent"
            Write-Log "MIME import method 2 failed: $($response.StatusCode) - $errorContent" -Level Warning
        }

        $httpClient.Dispose()
    }
    catch {
        $lastError = $_
        Write-Log "MIME import method 2 exception: $_" -Level Warning
    }

    # Method 3: Fallback to wrapper with .eml attachment
    Write-Log "Using fallback method 3 (wrapper with .eml attachment)" -Level Warning
    try {
        $messageId = Import-MessageToGraphBase64 -TargetMailbox $TargetMailbox -FolderId $FolderId -MimeContent $MimeContent -ReceivedDate $ReceivedDate -IsRead $IsRead

        if ($DiagnosticMode) {
            Write-Log "DIAG: Method 3 SUCCESS - Message ID: $messageId" -Level Debug
        }

        return $messageId
    }
    catch {
        $lastError = $_
        Write-Log "MIME import method 3 (Base64 fallback) also failed: $_" -Level Error
    }

    # All methods failed
    throw "All 3 import methods failed for message. Last error: $lastError"
}

function Set-MessageDateAndFlags {
    <#
    .SYNOPSIS
    Set the original date and mark message as received (not draft) using MAPI extended properties
    #>
    param(
        [string]$TargetMailbox,
        [string]$MessageId,
        [datetime]$ReceivedDate,
        [bool]$IsRead = $true
    )

    $updateUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$MessageId"

    # Format date for MAPI property (ISO 8601)
    $dateString = $ReceivedDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $updateBody = @{
        isRead = $IsRead
        # Use singleValueExtendedProperties to set MAPI properties
        singleValueExtendedProperties = @(
            @{
                # PR_MESSAGE_FLAGS - Mark as received (not draft)
                # MSGFLAG_READ (0x01) + MSGFLAG_UNMODIFIED (0x02) = 3
                # Without MSGFLAG_UNSENT (0x08) = not a draft
                id = "Integer 0x0E07"
                value = "1"
            },
            @{
                # PR_MESSAGE_DELIVERY_TIME - When message was delivered/received
                id = "SystemTime 0x0E06"
                value = $dateString
            },
            @{
                # PR_CLIENT_SUBMIT_TIME - When message was sent
                id = "SystemTime 0x0039"
                value = $dateString
            }
        )
    }

    try {
        Invoke-GraphRequest -Uri $updateUri -Method PATCH -Body $updateBody | Out-Null
    }
    catch {
        Write-Log "Failed to set message date/flags: $_" -Level Warning
    }
}

function Import-MessageToGraphBase64 {
    <#
    .SYNOPSIS
    Fallback method using Base64 encoded MIME content with .eml attachment
    #>
    param(
        [string]$TargetMailbox,
        [string]$FolderId,
        [byte[]]$MimeContent,
        [datetime]$ReceivedDate,
        [bool]$IsRead = $true
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$FolderId/messages"

    # Parse MIME headers to extract basic info for fallback message creation
    $mimeString = [System.Text.Encoding]::UTF8.GetString($MimeContent)

    # Extract headers for wrapper message
    $subject = "Imported Message"
    $from = ""
    $to = @()

    if ($mimeString -match '(?m)^Subject:\s*(.+?)(?:\r?\n(?!\s)|$)') {
        $subject = $matches[1].Trim()
        # Decode MIME encoded words if present
        if ($subject -match '=\?') {
            try {
                # Basic MIME word decoding
                $subject = $subject -replace '=\?([^?]+)\?([BQ])\?([^?]+)\?=', {
                    param($m)
                    $charset = $m.Groups[1].Value
                    $encoding = $m.Groups[2].Value
                    $data = $m.Groups[3].Value
                    if ($encoding -eq 'B') {
                        [System.Text.Encoding]::GetEncoding($charset).GetString([Convert]::FromBase64String($data))
                    }
                    else {
                        # Q encoding
                        $decoded = $data -replace '_', ' ' -replace '=([0-9A-F]{2})', { [char][convert]::ToInt32($_.Groups[1].Value, 16) }
                        $decoded
                    }
                }
            }
            catch { }
        }
    }

    if ($mimeString -match '(?m)^From:\s*(.+?)(?:\r?\n(?!\s)|$)') {
        $fromHeader = $matches[1].Trim()
        if ($fromHeader -match '<([^>]+)>') {
            $from = $matches[1]
        }
        elseif ($fromHeader -match '[\w\.-]+@[\w\.-]+') {
            $from = $matches[0]
        }
    }

    if ($mimeString -match '(?m)^To:\s*(.+?)(?:\r?\n(?!\s)|$)') {
        $toHeader = $matches[1].Trim()
        $toAddresses = [regex]::Matches($toHeader, '[\w\.-]+@[\w\.-]+')
        $to = $toAddresses | ForEach-Object { $_.Value }
    }

    # Use the ReceivedDate parameter (from IMAP INTERNALDATE) - more reliable than parsing MIME
    $date = if ($ReceivedDate -and $ReceivedDate -ne [datetime]::MinValue) { $ReceivedDate } else { Get-Date }

    # Create message with .eml attachment (preserves original exactly)
    $emlBase64 = [Convert]::ToBase64String($MimeContent)

    # Build HTML with proper encoding (use HTML entities for safety)
    $htmlBody = "<html><head><meta charset='UTF-8'></head><body>" +
        "<div style='padding:10px;background:#fff3cd;border-left:4px solid #ffc107;margin-bottom:15px;'>" +
        "<strong>Migrierte E-Mail</strong><br/>" +
        "<small>Original-Datum: $($date.ToString('dd.MM.yyyy HH:mm'))</small><br/>" +
        "<small>Die Original-E-Mail ist als .eml-Datei angeh&#228;ngt.</small>" +
        "</div></body></html>"

    $message = @{
        subject = "[Migrated] $subject"
        body    = @{
            contentType = "HTML"
            content     = $htmlBody
        }
        toRecipients = @()
        isRead       = $IsRead
        internetMessageHeaders = @(
            @{ name = "X-Migrated-From"; value = "Kopano-IMAP" }
            @{ name = "X-Original-Date"; value = $date.ToString("r") }
        )
        attachments  = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name          = "original_email.eml"
                contentType   = "message/rfc822"
                contentBytes  = $emlBase64
            }
        )
    }

    if ($from) {
        $message.from = @{
            emailAddress = @{ address = $from }
        }
    }

    if ($to.Count -gt 0) {
        $message.toRecipients = @($to | ForEach-Object {
            @{ emailAddress = @{ address = $_ } }
        })
    }

    try {
        $created = Invoke-GraphRequest -Uri $uri -Method POST -Body $message

        # Set original date and mark as NOT draft using MAPI extended properties
        Set-MessageDateAndFlags -TargetMailbox $TargetMailbox -MessageId $created.id -ReceivedDate $ReceivedDate -IsRead $IsRead

        return $created.id
    }
    catch {
        $errorDetails = ""
        if ($_.Exception.Response) {
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorDetails = $reader.ReadToEnd()
                $reader.Close()
            }
            catch { }
        }

        if ($errorDetails) {
            throw "Failed to import message (Base64 fallback): $_ - Details: $errorDetails"
        }
        throw "Failed to import message (Base64 fallback): $_"
    }
}

# ================================
# CSV Processing
# ================================

function Import-UserCsv {
    param(
        [string]$CsvPath
    )

    if (!(Test-Path $CsvPath)) {
        throw "User CSV file not found: $CsvPath"
    }

    Write-Log "Loading user list from: $CsvPath" -Level Info

    # Try to detect delimiter
    $firstLine = Get-Content $CsvPath -First 1
    $delimiter = if ($firstLine -match ';') { ';' } else { ',' }

    $users = Import-Csv -Path $CsvPath -Delimiter $delimiter

    # Validate required columns
    $requiredColumns = @('Email', 'Username', 'Password')
    $actualColumns = $users[0].PSObject.Properties.Name

    foreach ($col in $requiredColumns) {
        # Case-insensitive check
        $found = $actualColumns | Where-Object { $_ -ieq $col }
        if (!$found) {
            throw "CSV is missing required column: $col. Found columns: $($actualColumns -join ', ')"
        }
    }

    Write-Log "Loaded $($users.Count) users from CSV" -Level Success

    return $users
}

# ================================
# Folder Exclusion
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

# ================================
# IMAP Migration (MailKit-based)
# ================================

function Migrate-UserMailbox {
    param(
        [hashtable]$User,
        [hashtable]$FolderCache = @{}
    )

    $sourceEmail = $User.Email
    $imapUsername = $User.Username
    $imapPassword = $User.Password
    $targetEmail = if ($User.TargetEmail) { $User.TargetEmail } else { $sourceEmail }

    Write-Log "Starting migration for user: $sourceEmail -> $targetEmail" -Level Info -User $sourceEmail

    $userStats = @{
        TotalMessages = 0
        Migrated = 0
        Skipped = 0
        Failed = 0
        Folders = 0
    }

    $client = $null

    try {
        # Connect to IMAP using MailKit
        Write-Log "Connecting to IMAP server via MailKit..." -Level Info -User $sourceEmail
        $client = Get-MailKitClient `
            -Server $ImapServer `
            -Port $ImapPort `
            -UseSsl $ImapUseSsl `
            -SkipCertValidation $ImapSkipCertValidation `
            -Username $imapUsername `
            -Password $imapPassword

        Write-Log "Successfully connected to IMAP via MailKit" -Level Success -User $sourceEmail

        if ($DiagnosticMode) {
            Write-Log "DIAG: IMAP capabilities: $($client.Capabilities)" -Level Debug -User $sourceEmail
        }

        # Get all folders using MailKit's namespace support
        $personalNamespace = $client.PersonalNamespaces[0]
        $folders = $client.GetFolders($personalNamespace)

        Write-Log "Found $($folders.Count) folders" -Level Info -User $sourceEmail

        if ($DiagnosticMode) {
            foreach ($f in $folders) {
                Write-Log "DIAG: Folder: '$($f.FullName)' (Attributes: $($f.Attributes))" -Level Debug -User $sourceEmail
            }
        }

        # Filter folders if specified
        if ($FoldersToMigrate -and $FoldersToMigrate.Count -gt 0) {
            $folders = $folders | Where-Object {
                $folderName = $_.FullName
                $FoldersToMigrate | Where-Object { $folderName -ilike $_ -or $folderName -ieq $_ }
            }
            Write-Log "Filtered to $($folders.Count) folders matching criteria" -Level Info -User $sourceEmail
        }

        # Process each folder
        foreach ($folder in $folders) {
            $folderName = $folder.FullName

            # Check exclusions
            if (Test-FolderExcluded -FolderName $folderName) {
                Write-Log "Skipping excluded folder: $folderName" -Level Debug -User $sourceEmail
                continue
            }

            # Skip non-selectable folders
            if ($folder.Attributes -band [MailKit.FolderAttributes]::NonExistent) {
                Write-Log "Skipping non-existent folder: $folderName" -Level Debug -User $sourceEmail
                continue
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
                Write-Log "Folder is empty: $folderName" -Level Debug -User $sourceEmail
                $folder.Close()
                continue
            }

            Write-Log "Processing folder: $folderName ($($folder.Count) messages)" -Level Info -User $sourceEmail
            $userStats.Folders++

            try {
                # Build search query using MailKit's search API
                $query = [MailKit.Search.SearchQuery]::All

                if ($StartDate) {
                    $query = [MailKit.Search.SearchQuery]::And($query, [MailKit.Search.SearchQuery]::DeliveredAfter($StartDate))
                }
                if ($EndDate) {
                    $query = [MailKit.Search.SearchQuery]::And($query, [MailKit.Search.SearchQuery]::DeliveredBefore($EndDate))
                }

                # Search messages
                $uids = $folder.Search($query)

                if ($uids.Count -eq 0) {
                    Write-Log "No messages match criteria in folder: $folderName" -Level Debug -User $sourceEmail
                    $folder.Close()
                    continue
                }

                Write-Log "Found $($uids.Count) messages matching criteria" -Level Info -User $sourceEmail
                $userStats.TotalMessages += $uids.Count

                # Limit messages if specified
                $uidsToProcess = $uids
                if ($MaxMessagesPerMailbox -and ($userStats.Migrated + $uids.Count) -gt $MaxMessagesPerMailbox) {
                    $remaining = $MaxMessagesPerMailbox - $userStats.Migrated
                    if ($remaining -le 0) {
                        Write-Log "Reached message limit, stopping folder processing" -Level Warning -User $sourceEmail
                        $folder.Close()
                        break
                    }
                    $uidsToProcess = @($uids | Select-Object -First $remaining)
                }

                # Get or create target folder in Graph
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
                        Write-Log "Fetching message $msgIndex/$($uidsToProcess.Count) (UID: $uid)..." -Level Debug -User $sourceEmail

                        # Fetch full MimeMessage using MailKit
                        $message = $folder.GetMessage($uid)

                        # Get subject for logging
                        $subject = if ($message.Subject) { $message.Subject } else { "(no subject)" }
                        if ($subject.Length -gt 60) {
                            $subject = $subject.Substring(0, 57) + "..."
                        }

                        # Serialize to raw MIME bytes
                        $memStream = New-Object System.IO.MemoryStream
                        $message.WriteTo($memStream)
                        $mimeBytes = $memStream.ToArray()
                        $memStream.Dispose()

                        Write-Log "Fetched UID $uid: $($mimeBytes.Length) bytes - $subject" -Level Debug -User $sourceEmail

                        if ($DiagnosticMode) {
                            Write-Log "DIAG: Message Date: $($message.Date), From: $($message.From), MessageId: $($message.MessageId)" -Level Debug -User $sourceEmail
                        }

                        if ($mimeBytes.Length -eq 0) {
                            Write-Log "WARNING: Empty message content for UID $uid, skipping" -Level Warning -User $sourceEmail
                            $userStats.Skipped++
                            continue
                        }

                        # Save MIME to file if requested
                        if ($SaveMimeToFile) {
                            if (!(Test-Path $MimeSavePath)) {
                                New-Item -ItemType Directory -Path $MimeSavePath -Force | Out-Null
                            }
                            $safeSubject = ($subject -replace '[^\w\-\.]', '_')
                            if ($safeSubject.Length -gt 40) { $safeSubject = $safeSubject.Substring(0, 40) }
                            $mimeFilePath = Join-Path $MimeSavePath "uid_${uid}_${safeSubject}.eml"
                            [System.IO.File]::WriteAllBytes($mimeFilePath, $mimeBytes)
                            Write-Log "Saved MIME to: $mimeFilePath ($($mimeBytes.Length) bytes)" -Level Info -User $sourceEmail
                        }

                        # TestSingleMessage mode: display details and stop
                        if ($TestSingleMessage) {
                            Write-Log "=== TEST SINGLE MESSAGE ===" -Level Info -User $sourceEmail
                            Write-Log "  UID: $uid" -Level Info -User $sourceEmail
                            Write-Log "  Subject: $($message.Subject)" -Level Info -User $sourceEmail
                            Write-Log "  From: $($message.From)" -Level Info -User $sourceEmail
                            Write-Log "  To: $($message.To)" -Level Info -User $sourceEmail
                            Write-Log "  Date: $($message.Date)" -Level Info -User $sourceEmail
                            Write-Log "  MessageId: $($message.MessageId)" -Level Info -User $sourceEmail
                            Write-Log "  MIME size: $($mimeBytes.Length) bytes" -Level Info -User $sourceEmail
                            Write-Log "  Content-Type: $($message.Body.ContentType)" -Level Info -User $sourceEmail

                            # Show first 500 chars of MIME
                            $mimePreview = [System.Text.Encoding]::UTF8.GetString($mimeBytes, 0, [Math]::Min(500, $mimeBytes.Length))
                            Write-Log "  MIME preview:`n$mimePreview" -Level Info -User $sourceEmail
                            Write-Log "=== END TEST ===" -Level Info -User $sourceEmail

                            $folder.Close()
                            return $userStats
                        }

                        if ($WhatIf) {
                            Write-Log "[WHATIF] Would migrate: $subject (Date: $($message.Date.ToString('yyyy-MM-dd')))" -Level Info -User $sourceEmail
                            $userStats.Migrated++
                            continue
                        }

                        if ($SkipGraphImport) {
                            Write-Log "[SKIP-IMPORT] IMAP fetch OK: $subject ($($mimeBytes.Length) bytes)" -Level Success -User $sourceEmail
                            $userStats.Migrated++
                            continue
                        }

                        # Determine received date
                        $receivedDate = if ($message.Date.DateTime -ne [DateTime]::MinValue) {
                            $message.Date.DateTime
                        } else {
                            Get-Date
                        }

                        # Check if message was read
                        $msgSummary = $folder.Fetch(@($uid), [MailKit.MessageSummaryItems]::Flags)
                        $isRead = $false
                        if ($msgSummary -and $msgSummary.Count -gt 0) {
                            $isRead = [bool]($msgSummary[0].Flags -band [MailKit.MessageFlags]::Seen)
                        }

                        # Import to Graph
                        Write-Log "Importing: $subject" -Level Debug -User $sourceEmail

                        $importedId = Import-MessageToGraph `
                            -TargetMailbox $targetEmail `
                            -FolderId $targetFolderId `
                            -MimeContent $mimeBytes `
                            -ReceivedDate $receivedDate `
                            -IsRead $isRead

                        if ($importedId) {
                            $userStats.Migrated++
                            $script:stats.MigratedMessages++
                            Write-Log "Migrated [$msgIndex/$($uidsToProcess.Count)]: $subject" -Level Success -User $sourceEmail
                        }
                        else {
                            throw "Import returned no message ID"
                        }

                        # Throttle
                        if ($ThrottleMs -gt 0) {
                            Start-Sleep -Milliseconds $ThrottleMs
                        }
                    }
                    catch {
                        $userStats.Failed++
                        $script:stats.FailedMessages++
                        Write-Log "Failed to migrate message UID $uid : $_" -Level Error -User $sourceEmail

                        if ($DiagnosticMode) {
                            Write-Log "DIAG: Stack trace: $($_.ScriptStackTrace)" -Level Debug -User $sourceEmail
                        }

                        if (!$ContinueOnError) {
                            throw
                        }
                    }
                }
            }
            catch {
                Write-Log "Error processing folder $folderName : $_" -Level Error -User $sourceEmail

                if (!$ContinueOnError) {
                    $folder.Close()
                    throw
                }
            }

            # Close folder after processing
            try { $folder.Close() } catch { }
        }

        Write-Log "User migration complete. Migrated: $($userStats.Migrated), Failed: $($userStats.Failed), Skipped: $($userStats.Skipped)" -Level Success -User $sourceEmail
    }
    catch {
        Write-Log "Migration failed for user $sourceEmail : $_" -Level Error -User $sourceEmail
        throw
    }
    finally {
        # Cleanup IMAP connection
        if ($client) {
            try {
                if ($client.IsConnected) {
                    $client.Disconnect($true)
                }
                $client.Dispose()
            }
            catch {
                Write-Log "Error during IMAP cleanup: $_" -Level Debug -User $sourceEmail
            }
        }
    }

    return $userStats
}

# ================================
# State Management (Resume Support)
# ================================

function Save-MigrationState {
    param(
        [string]$StatePath,
        [hashtable]$State
    )

    $State | ConvertTo-Json -Depth 10 | Set-Content -Path $StatePath
    Write-Log "Migration state saved to: $StatePath" -Level Debug
}

function Load-MigrationState {
    param(
        [string]$StatePath
    )

    if (!(Test-Path $StatePath)) {
        return $null
    }

    $state = Get-Content $StatePath -Raw | ConvertFrom-Json
    return $state
}

# ================================
# Main Execution
# ================================

try {
    # Initialize logging
    Initialize-Logging

    Write-Log "=== Kopano IMAP to Graph Migration ===" -Level Info
    Write-Log "IMAP Server: $ImapServer`:$ImapPort (SSL: $ImapUseSsl)" -Level Info
    Write-Log "Tenant ID: $TenantId" -Level Info
    Write-Log "Client ID: $ClientId" -Level Info
    Write-Log "Engine: MailKit/MimeKit" -Level Info

    if ($DiagnosticMode) {
        Write-Log "*** DIAGNOSTIC MODE ENABLED ***" -Level Warning
    }

    # Load MailKit/MimeKit assemblies
    $mailkitReady = Initialize-MailKit
    if (!$mailkitReady) {
        throw "Failed to load MailKit/MimeKit. Run Setup-MailKit.ps1 first."
    }

    # === Validate Parameters ===
    if ($TestMode -or $TestSource -or $TestTarget -or $TestPassword) {
        # Test mode - validate test parameters
        if (!$TestSource) { throw "TestSource is required in test mode" }
        if (!$TestTarget) { throw "TestTarget is required in test mode" }
        if (!$TestPassword) { throw "TestPassword is required in test mode" }

        $TestMode = $true  # Ensure flag is set
        Write-Log "*** TEST MODE - Single user migration ***" -Level Warning
        Write-Log "Source: $TestSource" -Level Info
        Write-Log "Target: $TestTarget" -Level Info
    }
    elseif (!$UserCsvPath) {
        throw "Either -UserCsvPath or test mode parameters (-TestSource, -TestTarget, -TestPassword) are required"
    }

    if ($WhatIf) {
        Write-Log "*** WHATIF MODE - No actual migration will occur ***" -Level Warning
    }

    if ($SkipGraphImport) {
        Write-Log "*** SKIP GRAPH IMPORT - IMAP fetch testing only ***" -Level Warning
    }

    if ($TestSingleMessage) {
        Write-Log "*** TEST SINGLE MESSAGE MODE - Will fetch one message and stop ***" -Level Warning
    }

    if ($SaveMimeToFile) {
        Write-Log "*** SAVING MIME FILES TO: $MimeSavePath ***" -Level Warning
    }

    if ($StartDate -or $EndDate) {
        Write-Log "Date filter: $StartDate to $EndDate" -Level Info
    }

    if ($MaxMessagesPerMailbox) {
        Write-Log "Max messages per mailbox: $MaxMessagesPerMailbox" -Level Info
    }

    # Test Graph API connectivity (skip if not needed)
    if (!$SkipGraphImport) {
        Write-Log "Testing Graph API connectivity..." -Level Info
        $null = Get-GraphToken
    }
    else {
        Write-Log "Skipping Graph API token (SkipGraphImport mode)" -Level Info
    }

    # Load user list (CSV or Test mode)
    $users = @()
    if ($TestMode) {
        # Create single test user object
        $testUser = [PSCustomObject]@{
            Email       = $TestSource
            Username    = if ($TestUsername) { $TestUsername } else { $TestSource }
            Password    = $TestPassword
            TargetEmail = $TestTarget
        }
        $users = @($testUser)
        Write-Log "Test mode: Single user configured" -Level Info
    }
    else {
        $users = Import-UserCsv -CsvPath $UserCsvPath
    }
    $script:stats.TotalUsers = $users.Count

    # Load previous state if resuming
    $processedUsers = @{}
    if ($Resume -and $StateFile -and (Test-Path $StateFile)) {
        $previousState = Load-MigrationState -StatePath $StateFile
        if ($previousState) {
            Write-Log "Resuming from previous state..." -Level Info
            $processedUsers = @{}
            foreach ($u in $previousState.ProcessedUsers) {
                $processedUsers[$u] = $true
            }
            Write-Log "Already processed: $($processedUsers.Count) users" -Level Info
        }
    }

    # Folder cache for efficiency
    $folderCache = @{}

    # Process each user
    $userIndex = 0
    foreach ($user in $users) {
        $userIndex++

        # Normalize user data (handle case-insensitive column names)
        $normalizedUser = @{}
        foreach ($prop in $user.PSObject.Properties) {
            $normalizedUser[$prop.Name] = $prop.Value
        }

        # Handle case-insensitive lookups
        $email = $normalizedUser.Email
        if (!$email) { $email = $normalizedUser.email }

        $username = $normalizedUser.Username
        if (!$username) { $username = $normalizedUser.username }

        $password = $normalizedUser.Password
        if (!$password) { $password = $normalizedUser.password }

        $targetEmail = $normalizedUser.TargetEmail
        if (!$targetEmail) { $targetEmail = $normalizedUser.targetemail }

        $userHash = @{
            Email = $email
            Username = $username
            Password = $password
            TargetEmail = $targetEmail
        }

        # Check if already processed (resume support)
        if ($processedUsers.ContainsKey($email)) {
            Write-Log "Skipping already processed user: $email" -Level Info
            continue
        }

        Write-Log "`n========================================" -Level Info
        Write-Log "Processing user $userIndex of $($users.Count): $email" -Level Info
        Write-Log "========================================" -Level Info

        try {
            $userStats = Migrate-UserMailbox -User $userHash -FolderCache $folderCache

            $script:stats.ProcessedUsers++
            $processedUsers[$email] = $true

            # Save state after each user
            if ($StateFile) {
                $state = @{
                    ProcessedUsers = $processedUsers.Keys
                    LastProcessed = $email
                    Timestamp = Get-Date -Format 'o'
                    Stats = $script:stats
                }
                Save-MigrationState -StatePath $StateFile -State $state
            }
        }
        catch {
            Write-Log "Failed to migrate user $email : $_" -Level Error

            if (!$ContinueOnError) {
                throw
            }
        }
    }

    # Final summary
    $duration = (Get-Date) - $script:stats.StartTime

    Write-Log "`n========================================" -Level Info
    Write-Log "Migration Complete" -Level Success
    Write-Log "========================================" -Level Info
    Write-Log "Duration: $($duration.ToString('hh\:mm\:ss'))" -Level Info
    Write-Log "Users processed: $($script:stats.ProcessedUsers) of $($script:stats.TotalUsers)" -Level Info
    Write-Log "Messages migrated: $($script:stats.MigratedMessages)" -Level Info
    Write-Log "Messages failed: $($script:stats.FailedMessages)" -Level Info
    Write-Log "Messages skipped: $($script:stats.SkippedMessages)" -Level Info
    Write-Log "Log file: $script:logFile" -Level Info

    if ($StateFile) {
        Write-Log "State file: $StateFile" -Level Info
    }
}
catch {
    Write-Log "Fatal error: $_" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error

    # Save error state
    if ($StateFile) {
        $state = @{
            ProcessedUsers = $processedUsers.Keys
            LastError = $_.ToString()
            Timestamp = Get-Date -Format 'o'
            Stats = $script:stats
        }
        Save-MigrationState -StatePath $StateFile -State $state
    }

    throw
}

<#
.SYNOPSIS
    Migrates emails from Kopano IMAP server to Microsoft 365 via Graph API

.DESCRIPTION
    This script connects to an IMAP server (designed for Kopano) and migrates
    emails to Microsoft 365 mailboxes using the Microsoft Graph API.

    Uses MailKit/MimeKit for reliable IMAP operations. Run Setup-MailKit.ps1 first.

    Features:
    - Bulk migration from CSV user list
    - Preserves original email dates via MIME import
    - Maintains folder structure
    - Resume capability for interrupted migrations
    - Detailed logging and error handling
    - Diagnostic mode for troubleshooting

.PARAMETER TenantId
    Microsoft 365 tenant ID

.PARAMETER ClientId
    Azure AD application client ID (requires Mail.ReadWrite application permission)

.PARAMETER ClientSecret
    Azure AD application client secret

.PARAMETER ImapServer
    Kopano IMAP server hostname

.PARAMETER ImapPort
    IMAP port (default: 993 for SSL)

.PARAMETER ImapUseSsl
    Use SSL/TLS for IMAP connection (default: true)

.PARAMETER ImapSkipCertValidation
    Skip SSL certificate validation (for self-signed certificates)

.PARAMETER UserCsvPath
    Path to CSV file with columns: Email, Username, Password, TargetEmail (optional)

.PARAMETER FoldersToMigrate
    Specific folders to migrate (empty = all folders)

.PARAMETER ExcludeFolders
    Folders to exclude from migration

.PARAMETER StartDate
    Only migrate emails after this date

.PARAMETER EndDate
    Only migrate emails before this date

.PARAMETER MaxMessagesPerMailbox
    Limit number of messages per mailbox (for testing)

.PARAMETER PreserveFolderStructure
    Create matching folder structure in target (default: true)

.PARAMETER PreserveReceivedDate
    Preserve original received date (default: true)

.PARAMETER DiagnosticMode
    Enable full diagnostic logging (IMAP commands, byte counts, Graph API details)

.PARAMETER TestSingleMessage
    Fetch and display ONE message for debugging, then stop

.PARAMETER SaveMimeToFile
    Save fetched MIME content to disk before importing to Graph

.PARAMETER SkipGraphImport
    Test IMAP fetch only, skip Graph API import

.PARAMETER MimeSavePath
    Directory to save MIME files when -SaveMimeToFile is used (default: .\mime_dump)

.PARAMETER WhatIf
    Dry run - show what would be migrated without actually migrating

.PARAMETER ContinueOnError
    Continue processing other messages/users on errors

.PARAMETER StateFile
    Path to state file for resume capability

.PARAMETER Resume
    Resume from previous state file

.EXAMPLE
    # First run setup to download MailKit:
    .\Setup-MailKit.ps1

    # Then run migration:
    .\Kopano-IMAP-to-Graph-Migration.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -ClientSecret "your-secret" `
        -ImapServer "mail.kopano.local" `
        -UserCsvPath ".\users.csv" `
        -WhatIf

    Dry run to test configuration

.EXAMPLE
    .\Kopano-IMAP-to-Graph-Migration.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -ClientSecret "your-secret" `
        -ImapServer "mail.kopano.local" `
        -TestMode `
        -TestSource "user@company.com" `
        -TestTarget "user@company.com" `
        -TestPassword "password" `
        -TestSingleMessage `
        -DiagnosticMode

    Debug mode: fetch one message with full diagnostics

.EXAMPLE
    .\Kopano-IMAP-to-Graph-Migration.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -ClientSecret "your-secret" `
        -ImapServer "mail.kopano.local" `
        -ImapPort 993 `
        -ImapUseSsl `
        -UserCsvPath ".\users.csv" `
        -StartDate "2023-01-01" `
        -StateFile ".\migration_state.json" `
        -ContinueOnError

    Full migration with date filter and resume support

.EXAMPLE
    .\Kopano-IMAP-to-Graph-Migration.ps1 `
        -TenantId "your-tenant-id" `
        -ClientId "your-client-id" `
        -ClientSecret "your-secret" `
        -ImapServer "mail.kopano.local" `
        -UserCsvPath ".\users.csv" `
        -FoldersToMigrate @("INBOX", "Sent") `
        -MaxMessagesPerMailbox 100

    Migrate only specific folders with message limit

.NOTES
    Prerequisites:
    - Run Setup-MailKit.ps1 to download MailKit/MimeKit DLLs
    - PowerShell 5.1 or later (or PowerShell 7+)

    CSV Format:
    Email,Username,Password,TargetEmail
    user@company.com,user,password123,user@company.onmicrosoft.com

    If TargetEmail is omitted, the Email value is used as target.

    Required Azure AD App Permissions:
    - Mail.ReadWrite (Application)
    - User.Read.All (Application) - optional, for user validation

.LINK
    https://docs.microsoft.com/en-us/graph/api/user-post-messages
#>
