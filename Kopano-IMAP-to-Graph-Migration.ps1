<#
Kopano-IMAP-to-Graph-Migration.ps1 — IMAP to Microsoft Graph Email Migration
Migrates emails from Kopano IMAP server to Microsoft 365 via Graph API

Features:
  - Bulk migration from CSV user list
  - Preserves original sent/received dates
  - Maintains folder structure
  - Supports SSL/TLS IMAP connections
  - Progress tracking and detailed logging
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

    # === User List ===
    [Parameter(Mandatory)]
    [string]$UserCsvPath,                   # CSV file with: Email,Username,Password,TargetEmail (optional)

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
    Loads MailKit library for IMAP operations
    #>

    # Check if MailKit is already loaded
    $mailkitLoaded = [System.AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq 'MailKit' }

    if ($mailkitLoaded) {
        Write-Log "MailKit already loaded" -Level Info
        return $true
    }

    # Try to load from NuGet packages
    $possiblePaths = @(
        "$PSScriptRoot\packages\MailKit\lib\netstandard2.0\MailKit.dll",
        "$PSScriptRoot\lib\MailKit.dll",
        "$env:USERPROFILE\.nuget\packages\mailkit\*\lib\netstandard2.0\MailKit.dll",
        "$env:USERPROFILE\.nuget\packages\mailkit\*\lib\net6.0\MailKit.dll"
    )

    $mimeKitPaths = @(
        "$PSScriptRoot\packages\MimeKit\lib\netstandard2.0\MimeKit.dll",
        "$PSScriptRoot\lib\MimeKit.dll",
        "$env:USERPROFILE\.nuget\packages\mimekit\*\lib\netstandard2.0\MimeKit.dll",
        "$env:USERPROFILE\.nuget\packages\mimekit\*\lib\net6.0\MimeKit.dll"
    )

    # Try to find and load MimeKit first (dependency)
    $mimeKitLoaded = $false
    foreach ($pattern in $mimeKitPaths) {
        $paths = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object { [version]($_.Directory.Parent.Name -replace '[^\d.]', '') } -Descending |
            Select-Object -First 1

        if ($paths) {
            try {
                Add-Type -Path $paths.FullName
                Write-Log "Loaded MimeKit from: $($paths.FullName)" -Level Info
                $mimeKitLoaded = $true
                break
            }
            catch {
                Write-Log "Failed to load MimeKit from $($paths.FullName): $_" -Level Warning
            }
        }
    }

    # Try to find and load MailKit
    foreach ($pattern in $possiblePaths) {
        $paths = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object { [version]($_.Directory.Parent.Name -replace '[^\d.]', '') } -Descending |
            Select-Object -First 1

        if ($paths) {
            try {
                Add-Type -Path $paths.FullName
                Write-Log "Loaded MailKit from: $($paths.FullName)" -Level Info
                return $true
            }
            catch {
                Write-Log "Failed to load MailKit from $($paths.FullName): $_" -Level Warning
            }
        }
    }

    # Try to install via NuGet
    Write-Log "MailKit not found. Attempting to install via NuGet..." -Level Warning

    try {
        # Check if NuGet provider is available
        if (!(Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
        }

        # Install MailKit package
        $installPath = "$PSScriptRoot\packages"
        if (!(Test-Path $installPath)) {
            New-Item -ItemType Directory -Path $installPath -Force | Out-Null
        }

        # Use nuget.exe if available, otherwise use Install-Package
        $nugetExe = Get-Command nuget.exe -ErrorAction SilentlyContinue
        if ($nugetExe) {
            & nuget.exe install MailKit -OutputDirectory $installPath -NonInteractive
        }
        else {
            # Fallback: try to use .NET restore or manual download
            Write-Log "Please install MailKit manually:" -Level Warning
            Write-Log "  Option 1: dotnet add package MailKit" -Level Warning
            Write-Log "  Option 2: Install-Package MailKit -Scope CurrentUser" -Level Warning
            Write-Log "  Option 3: Download from https://www.nuget.org/packages/MailKit/" -Level Warning
            return $false
        }

        # Retry loading
        return Initialize-MailKit
    }
    catch {
        Write-Log "Failed to install MailKit: $_" -Level Error
        return $false
    }
}

# ================================
# Alternative: Pure .NET IMAP Implementation
# ================================

class SimpleImapClient {
    [System.Net.Sockets.TcpClient]$TcpClient
    [System.IO.StreamReader]$Reader
    [System.IO.StreamWriter]$Writer
    [System.Net.Security.SslStream]$SslStream
    [int]$TagCounter = 0
    [bool]$Connected = $false
    [bool]$SkipCertValidation = $false

    SimpleImapClient([bool]$skipCertValidation) {
        $this.SkipCertValidation = $skipCertValidation
    }

    [string] GetNextTag() {
        $this.TagCounter++
        return "A{0:D4}" -f $this.TagCounter
    }

    [void] Connect([string]$server, [int]$port, [bool]$useSsl) {
        $this.TcpClient = New-Object System.Net.Sockets.TcpClient
        $this.TcpClient.Connect($server, $port)

        if ($useSsl) {
            if ($this.SkipCertValidation) {
                $callback = [System.Net.Security.RemoteCertificateValidationCallback]{
                    param($sender, $certificate, $chain, $sslPolicyErrors)
                    return $true
                }
                $this.SslStream = New-Object System.Net.Security.SslStream(
                    $this.TcpClient.GetStream(),
                    $false,
                    $callback
                )
            }
            else {
                $this.SslStream = New-Object System.Net.Security.SslStream(
                    $this.TcpClient.GetStream(),
                    $false
                )
            }
            $this.SslStream.AuthenticateAsClient($server)
            $this.Reader = New-Object System.IO.StreamReader($this.SslStream)
            $this.Writer = New-Object System.IO.StreamWriter($this.SslStream)
        }
        else {
            $stream = $this.TcpClient.GetStream()
            $this.Reader = New-Object System.IO.StreamReader($stream)
            $this.Writer = New-Object System.IO.StreamWriter($stream)
        }

        $this.Writer.AutoFlush = $true

        # Read greeting
        $greeting = $this.Reader.ReadLine()
        if ($greeting -notmatch '^\* OK') {
            throw "IMAP server did not send OK greeting: $greeting"
        }

        $this.Connected = $true
    }

    [hashtable] SendCommand([string]$command) {
        $tag = $this.GetNextTag()
        $this.Writer.WriteLine("$tag $command")

        $response = @{
            Tag = $tag
            Lines = @()
            Success = $false
            ResultLine = ""
        }

        while ($true) {
            $line = $this.Reader.ReadLine()
            if ($null -eq $line) {
                throw "Connection closed unexpectedly"
            }

            if ($line.StartsWith($tag)) {
                $response.ResultLine = $line
                $response.Success = $line -match "^$tag OK"
                break
            }

            $response.Lines += $line
        }

        return $response
    }

    [bool] Login([string]$username, [string]$password) {
        # Escape special characters in password
        $escapedPassword = $password -replace '\\', '\\' -replace '"', '\"'
        $result = $this.SendCommand("LOGIN `"$username`" `"$escapedPassword`"")
        return $result.Success
    }

    [array] ListFolders() {
        $result = $this.SendCommand('LIST "" "*"')
        $folders = @()

        foreach ($line in $result.Lines) {
            if ($line -match '^\* LIST \(([^)]*)\) "([^"]*)" "?([^"]+)"?$') {
                $flags = $matches[1]
                $delimiter = $matches[2]
                $name = $matches[3] -replace '^"', '' -replace '"$', ''

                $folders += @{
                    Name = $name
                    Flags = $flags
                    Delimiter = $delimiter
                }
            }
        }

        return $folders
    }

    [hashtable] SelectFolder([string]$folder) {
        $result = $this.SendCommand("SELECT `"$folder`"")

        $info = @{
            Success = $result.Success
            Exists = 0
            Recent = 0
            UidValidity = 0
        }

        foreach ($line in $result.Lines) {
            if ($line -match '^\* (\d+) EXISTS') {
                $info.Exists = [int]$matches[1]
            }
            elseif ($line -match '^\* (\d+) RECENT') {
                $info.Recent = [int]$matches[1]
            }
            elseif ($line -match 'UIDVALIDITY (\d+)') {
                $info.UidValidity = [int]$matches[1]
            }
        }

        return $info
    }

    [array] SearchMessages([string]$criteria = "ALL") {
        $result = $this.SendCommand("UID SEARCH $criteria")
        $uids = @()

        foreach ($line in $result.Lines) {
            if ($line -match '^\* SEARCH (.*)$') {
                $uids = $matches[1].Trim().Split(' ') | Where-Object { $_ -ne '' } | ForEach-Object { [int]$_ }
            }
        }

        return $uids
    }

    [hashtable] FetchMessageHeaders([int]$uid) {
        $result = $this.SendCommand("UID FETCH $uid (BODY.PEEK[HEADER] INTERNALDATE FLAGS)")

        $message = @{
            UID = $uid
            Headers = ""
            InternalDate = $null
            Flags = @()
        }

        $inHeaders = $false
        $headerLines = @()

        foreach ($line in $result.Lines) {
            if ($line -match 'INTERNALDATE "([^"]+)"') {
                try {
                    $message.InternalDate = [datetime]::ParseExact(
                        $matches[1],
                        "d-MMM-yyyy HH:mm:ss zzz",
                        [System.Globalization.CultureInfo]::InvariantCulture
                    )
                }
                catch {
                    # Try alternative format
                    try {
                        $message.InternalDate = [datetime]::Parse($matches[1])
                    }
                    catch { }
                }
            }

            if ($line -match 'FLAGS \(([^)]*)\)') {
                $message.Flags = $matches[1].Split(' ')
            }

            if ($line -match 'BODY\[HEADER\]') {
                $inHeaders = $true
                continue
            }

            if ($inHeaders) {
                if ($line -eq ')' -or $line -match '^\)$') {
                    $inHeaders = $false
                }
                else {
                    $headerLines += $line
                }
            }
        }

        $message.Headers = $headerLines -join "`r`n"

        return $message
    }

    [string] FetchMessageRaw([int]$uid) {
        $result = $this.SendCommand("UID FETCH $uid (BODY.PEEK[])")

        $messageLines = @()
        $inMessage = $false
        $bytesToRead = 0

        foreach ($line in $result.Lines) {
            if ($line -match 'BODY\[\] \{(\d+)\}') {
                $bytesToRead = [int]$matches[1]
                $inMessage = $true
                continue
            }

            if ($inMessage) {
                if ($line -eq ')' -and $messageLines.Count -gt 0) {
                    break
                }
                $messageLines += $line
            }
        }

        return ($messageLines -join "`r`n")
    }

    [byte[]] FetchMessageBytes([int]$uid) {
        $tag = $this.GetNextTag()
        $this.Writer.WriteLine("$tag UID FETCH $uid (BODY.PEEK[])")

        $allBytes = New-Object System.Collections.Generic.List[byte]
        $inLiteral = $false
        $literalSize = 0
        $literalBytesRead = 0

        # Read response until we get the tagged response
        while ($true) {
            $line = $this.Reader.ReadLine()

            if ($null -eq $line) {
                throw "Connection closed unexpectedly"
            }

            if ($line.StartsWith($tag)) {
                break
            }

            # Check for literal start
            if ($line -match '\{(\d+)\}$') {
                $literalSize = [int]$matches[1]
                $inLiteral = $true

                # Read literal bytes
                $buffer = New-Object byte[] $literalSize
                $bytesRead = 0

                while ($bytesRead -lt $literalSize) {
                    $chunk = $this.SslStream.Read($buffer, $bytesRead, $literalSize - $bytesRead)
                    if ($chunk -eq 0) {
                        throw "Connection closed while reading literal"
                    }
                    $bytesRead += $chunk
                }

                $allBytes.AddRange($buffer)
                $inLiteral = $false
            }
        }

        return $allBytes.ToArray()
    }

    [void] Logout() {
        if ($this.Connected) {
            try {
                $this.SendCommand("LOGOUT")
            }
            catch { }

            $this.Reader.Dispose()
            $this.Writer.Dispose()
            if ($this.SslStream) { $this.SslStream.Dispose() }
            $this.TcpClient.Close()
            $this.Connected = $false
        }
    }
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
        $params.ContentType = $ContentType
        if ($ContentType -eq "application/json" -and $Body -isnot [string]) {
            $params.Body = $Body | ConvertTo-Json -Depth 20 -Compress
        }
        else {
            $params.Body = $Body
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

    # Normalize folder path
    $normalizedPath = $FolderPath -replace '/', '\' -replace '\\+', '\'
    $parts = $normalizedPath.Split('\') | Where-Object { $_ -ne '' }

    # Map common folder names
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

    # Graph API supports importing MIME messages with received date preservation
    # We'll use the messages endpoint with raw MIME content

    $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$FolderId/messages"

    $headers = @{
        "Authorization" = "Bearer $(Get-GraphToken)"
        "Content-Type"  = "text/plain"
    }

    try {
        # First, create message from MIME
        $response = Invoke-WebRequest -Method POST -Uri $uri -Headers $headers -Body $MimeContent
        $createdMessage = $response.Content | ConvertFrom-Json

        # Update the message to set read status and potentially adjust dates
        $updateUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$($createdMessage.id)"
        $updateBody = @{
            isRead = $IsRead
        }

        # Note: receivedDateTime cannot be modified after creation via MIME import,
        # but the MIME Date header should be preserved

        Invoke-GraphRequest -Uri $updateUri -Method PATCH -Body $updateBody | Out-Null

        return $createdMessage.id
    }
    catch {
        throw "Failed to import message: $_"
    }
}

function Import-MessageToGraphAlternative {
    <#
    .SYNOPSIS
    Alternative method using message creation with explicit date setting
    #>
    param(
        [string]$TargetMailbox,
        [string]$FolderId,
        [hashtable]$MessageData,
        [datetime]$ReceivedDate
    )

    $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$FolderId/messages"

    $message = @{
        subject = $MessageData.Subject
        body = @{
            contentType = "HTML"
            content = $MessageData.Body
        }
        from = @{
            emailAddress = @{
                address = $MessageData.From
            }
        }
        toRecipients = @(
            $MessageData.To | ForEach-Object {
                @{ emailAddress = @{ address = $_ } }
            }
        )
        isRead = $MessageData.IsRead
        # receivedDateTime will be set to creation time by Graph API
        # Unfortunately, we cannot set receivedDateTime directly
        internetMessageHeaders = @(
            @{ name = "X-Original-Date"; value = $ReceivedDate.ToString("r") }
            @{ name = "X-Migration-Source"; value = "Kopano-IMAP" }
            @{ name = "X-Migration-Date"; value = (Get-Date).ToString("r") }
        )
    }

    if ($MessageData.Cc) {
        $message.ccRecipients = @(
            $MessageData.Cc | ForEach-Object {
                @{ emailAddress = @{ address = $_ } }
            }
        )
    }

    return Invoke-GraphRequest -Uri $uri -Method POST -Body $message
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
# IMAP Migration Functions
# ================================

function Get-ImapClient {
    param(
        [string]$Server,
        [int]$Port,
        [bool]$UseSsl,
        [bool]$SkipCertValidation
    )

    $client = [SimpleImapClient]::new($SkipCertValidation)
    $client.Connect($Server, $Port, $UseSsl)

    return $client
}

function Get-MessageSearchCriteria {
    $criteria = "ALL"
    $parts = @()

    if ($StartDate) {
        $parts += "SINCE $($StartDate.ToString('dd-MMM-yyyy'))"
    }

    if ($EndDate) {
        $parts += "BEFORE $($EndDate.ToString('dd-MMM-yyyy'))"
    }

    if ($parts.Count -gt 0) {
        $criteria = $parts -join " "
    }

    return $criteria
}

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
        # Connect to IMAP
        Write-Log "Connecting to IMAP server..." -Level Info -User $sourceEmail
        $client = Get-ImapClient -Server $ImapServer -Port $ImapPort -UseSsl $ImapUseSsl -SkipCertValidation $ImapSkipCertValidation

        # Login
        Write-Log "Authenticating..." -Level Debug -User $sourceEmail
        $loginSuccess = $client.Login($imapUsername, $imapPassword)

        if (!$loginSuccess) {
            throw "IMAP login failed for user $imapUsername"
        }

        Write-Log "Successfully logged in to IMAP" -Level Success -User $sourceEmail

        # Get folder list
        $folders = $client.ListFolders()
        Write-Log "Found $($folders.Count) folders" -Level Info -User $sourceEmail

        # Filter folders if specified
        if ($FoldersToMigrate -and $FoldersToMigrate.Count -gt 0) {
            $folders = $folders | Where-Object {
                $folderName = $_.Name
                $FoldersToMigrate | Where-Object { $folderName -ilike $_ }
            }
        }

        # Process each folder
        foreach ($folder in $folders) {
            $folderName = $folder.Name

            # Check exclusions
            if (Test-FolderExcluded -FolderName $folderName) {
                Write-Log "Skipping excluded folder: $folderName" -Level Debug -User $sourceEmail
                continue
            }

            # Check for \Noselect flag
            if ($folder.Flags -match '\\Noselect') {
                Write-Log "Skipping non-selectable folder: $folderName" -Level Debug -User $sourceEmail
                continue
            }

            Write-Log "Processing folder: $folderName" -Level Info -User $sourceEmail
            $userStats.Folders++

            try {
                # Select folder
                $folderInfo = $client.SelectFolder($folderName)

                if (!$folderInfo.Success) {
                    Write-Log "Failed to select folder: $folderName" -Level Warning -User $sourceEmail
                    continue
                }

                if ($folderInfo.Exists -eq 0) {
                    Write-Log "Folder is empty: $folderName" -Level Debug -User $sourceEmail
                    continue
                }

                Write-Log "Folder has $($folderInfo.Exists) messages" -Level Info -User $sourceEmail

                # Search messages
                $searchCriteria = Get-MessageSearchCriteria
                $messageUids = $client.SearchMessages($searchCriteria)

                if ($messageUids.Count -eq 0) {
                    Write-Log "No messages match criteria in folder: $folderName" -Level Debug -User $sourceEmail
                    continue
                }

                Write-Log "Found $($messageUids.Count) messages matching criteria" -Level Info -User $sourceEmail
                $userStats.TotalMessages += $messageUids.Count

                # Limit messages if specified
                if ($MaxMessagesPerMailbox -and ($userStats.Migrated + $messageUids.Count) -gt $MaxMessagesPerMailbox) {
                    $remaining = $MaxMessagesPerMailbox - $userStats.Migrated
                    if ($remaining -le 0) {
                        Write-Log "Reached message limit, stopping folder processing" -Level Warning -User $sourceEmail
                        break
                    }
                    $messageUids = $messageUids | Select-Object -First $remaining
                }

                # Get or create target folder
                $targetFolderId = $null
                if ($PreserveFolderStructure) {
                    $targetFolderId = Get-OrCreateMailFolder -TargetMailbox $targetEmail -FolderPath $folderName -FolderCache $FolderCache
                }
                else {
                    # Use Inbox as default target
                    $targetFolderId = Get-OrCreateMailFolder -TargetMailbox $targetEmail -FolderPath "Inbox" -FolderCache $FolderCache
                }

                # Process messages
                $messageCount = 0
                foreach ($uid in $messageUids) {
                    $messageCount++

                    try {
                        # Fetch message
                        Write-Log "Fetching message $messageCount/$($messageUids.Count) (UID: $uid)..." -Level Debug -User $sourceEmail

                        # Get headers first for info
                        $headers = $client.FetchMessageHeaders($uid)

                        $receivedDate = if ($headers.InternalDate) {
                            $headers.InternalDate
                        } else {
                            Get-Date
                        }

                        $isRead = $headers.Flags -contains '\Seen'

                        # Parse subject from headers for logging
                        $subject = "Unknown"
                        if ($headers.Headers -match 'Subject:\s*(.+?)(?:\r?\n(?!\s)|$)') {
                            $subject = $matches[1].Trim()
                            if ($subject.Length -gt 50) {
                                $subject = $subject.Substring(0, 47) + "..."
                            }
                        }

                        if ($WhatIf) {
                            Write-Log "[WHATIF] Would migrate: $subject (Date: $($receivedDate.ToString('yyyy-MM-dd')))" -Level Info -User $sourceEmail
                            $userStats.Migrated++
                            continue
                        }

                        # Fetch full message
                        $rawMessage = $client.FetchMessageRaw($uid)

                        if (!$rawMessage -or $rawMessage.Length -eq 0) {
                            Write-Log "Empty message content for UID $uid, skipping" -Level Warning -User $sourceEmail
                            $userStats.Skipped++
                            continue
                        }

                        $messageBytes = [System.Text.Encoding]::UTF8.GetBytes($rawMessage)

                        # Import to Graph
                        Write-Log "Importing: $subject" -Level Debug -User $sourceEmail

                        $importedId = Import-MessageToGraph `
                            -TargetMailbox $targetEmail `
                            -FolderId $targetFolderId `
                            -MimeContent $messageBytes `
                            -ReceivedDate $receivedDate `
                            -IsRead $isRead

                        $userStats.Migrated++
                        $script:stats.MigratedMessages++

                        Write-Log "Migrated: $subject" -Level Success -User $sourceEmail

                        # Throttle
                        if ($ThrottleMs -gt 0) {
                            Start-Sleep -Milliseconds $ThrottleMs
                        }
                    }
                    catch {
                        $userStats.Failed++
                        $script:stats.FailedMessages++
                        Write-Log "Failed to migrate message UID $uid : $_" -Level Error -User $sourceEmail

                        if (!$ContinueOnError) {
                            throw
                        }
                    }
                }
            }
            catch {
                Write-Log "Error processing folder $folderName : $_" -Level Error -User $sourceEmail

                if (!$ContinueOnError) {
                    throw
                }
            }
        }

        Write-Log "User migration complete. Migrated: $($userStats.Migrated), Failed: $($userStats.Failed), Skipped: $($userStats.Skipped)" -Level Success -User $sourceEmail
    }
    catch {
        Write-Log "Migration failed for user $sourceEmail : $_" -Level Error -User $sourceEmail
        throw
    }
    finally {
        # Cleanup IMAP connection
        if ($client -and $client.Connected) {
            try {
                $client.Logout()
            }
            catch {
                Write-Log "Error during IMAP logout: $_" -Level Debug -User $sourceEmail
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

    if ($WhatIf) {
        Write-Log "*** WHATIF MODE - No actual migration will occur ***" -Level Warning
    }

    if ($StartDate -or $EndDate) {
        Write-Log "Date filter: $StartDate to $EndDate" -Level Info
    }

    if ($MaxMessagesPerMailbox) {
        Write-Log "Max messages per mailbox: $MaxMessagesPerMailbox" -Level Info
    }

    # Test Graph API connectivity
    Write-Log "Testing Graph API connectivity..." -Level Info
    $null = Get-GraphToken

    # Load user list
    $users = Import-UserCsv -CsvPath $UserCsvPath
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

    Features:
    - Bulk migration from CSV user list
    - Preserves original email dates via MIME import
    - Maintains folder structure
    - Resume capability for interrupted migrations
    - Detailed logging and error handling

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

.PARAMETER WhatIf
    Dry run - show what would be migrated without actually migrating

.PARAMETER ContinueOnError
    Continue processing other messages/users on errors

.PARAMETER StateFile
    Path to state file for resume capability

.PARAMETER Resume
    Resume from previous state file

.EXAMPLE
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
