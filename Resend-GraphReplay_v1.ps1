<# 
Resend-GraphReplay.ps1 — Microsoft Graph API Email Replay
Uses App Registration with Mail.Read and Mail.Send permissions
Two modes: 
  1) Transparent replay (preserves MIME)
  2) Wrapper with banner + .eml attachment
#>

[CmdletBinding()]
param(
    # Configuration file (overrides individual parameters if provided)
    [Parameter(ParameterSetName='ConfigFile')]
    [string]$Config,                     # Path to JSON config file

    [switch]$DebugOutput,
    
    [string]$ReplayHeader = "X-Graph-Replay:2025-Migration",  # "HeaderName:Value" or $null for none

    [switch]$TrueTransparentReplay,   # Use modern transparent replay via /messages + text/plain

    # Authentication (required if no config file)
    [Parameter(ParameterSetName='Direct', Mandatory)]
    [string]$TenantId,
    
    [Parameter(ParameterSetName='Direct', Mandatory)]
    [string]$ClientId,
    
    [Parameter(ParameterSetName='Direct', Mandatory)]
    [string]$ClientSecret,
    
    # Source configuration
    [Parameter(ParameterSetName='Direct', Mandatory)]
    [Parameter(ParameterSetName='ConfigFile')]
    [string[]]$SourceMailboxes,          # Array of mailboxes to process
    
    [Parameter(ParameterSetName='Direct')]
    [Parameter(ParameterSetName='ConfigFile')]
    [string]$FolderName = "Inbox",       # Folder to process (Inbox/Posteingang)
    
    # Target configuration
    [Parameter(ParameterSetName='Direct', Mandatory)]
    [Parameter(ParameterSetName='ConfigFile')]
    [string]$TargetMailbox,              # Where to send replayed emails
    
    # Processing options
    [switch]$AttachmentsOnly,            # Only process emails with attachments
    $StartDate,                          # Optional date range start (accepts string or datetime)
    $EndDate,                            # Optional date range end (accepts string or datetime)
    [string]$SubjectFilter,              # Optional subject filter
    [int]$MaxMessages,                   # Limit number of messages
    
    # Replay mode
    [ValidateSet("Transparent", "Wrapper")]
    [string]$ReplayMode = "Transparent", # Transparent = as-is, Wrapper = banner+attachment
    
    # Logging
    [string]$LogPath,                    # Path for success logging
    [switch]$LogSuccessful = $true,      # Log successful sends
    
    # Testing
    [switch]$TestMode,                   # Send single test email only
    [string]$TestMailbox,                # Specific test mailbox
    [switch]$WhatIf,                     # Dry run mode
    
    # Advanced
    [string[]]$BccAlways,                # Always BCC these addresses
    [switch]$SkipAlreadyProcessed,       # Skip if custom header exists
    [string]$ProcessedHeader = "X-GraphReplay-Processed",
    [int]$BatchSize = 50,                # Graph API batch size
    [int]$ThrottleMs = 100               # Throttle between sends
)

# ================================
# Configuration File Management
# ================================

$LogRoot = "c:\data\logs"   # Or wherever you want
if (!(Test-Path $LogRoot)) { New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null }

function Get-LogFilePath($mailbox) {
    $safe = $mailbox -replace '[^\w\-\.@]', '_'
    return Join-Path $LogRoot "$safe.csv"
}

$DedupFolder = Join-Path $LogRoot "dedup"
if (!(Test-Path $DedupFolder)) { New-Item -ItemType Directory -Path $DedupFolder -Force | Out-Null }

function Already-Sent($messageId, $recipient) {
    $marker = Join-Path $DedupFolder ("$($messageId)_$($recipient -replace '[^\w\-\.@]', '_').txt")
    return Test-Path $marker
}

function Mark-Sent($messageId, $recipient) {
    $marker = Join-Path $DedupFolder ("$($messageId)_$($recipient -replace '[^\w\-\.@]', '_').txt")
    Set-Content -Path $marker -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -NoNewline
}

function Save-ReplayConfig {
    param(
        [string]$Path,
        [hashtable]$Configuration
    )
    
    $safeConfig = $Configuration.Clone()
    
    if ($Configuration.ClientSecret) {
        try {
            $secureString = ConvertTo-SecureString $Configuration.ClientSecret -AsPlainText -Force
            $encryptedSecret = ConvertFrom-SecureString $secureString
            $safeConfig.ClientSecretEncrypted = $encryptedSecret
            $safeConfig.Remove('ClientSecret')
        }
        catch {
            Write-Warning "Could not encrypt ClientSecret - storing in plain text"
        }
    }
    
    $safeConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $Path
    Write-Host "Configuration saved to: $Path" -ForegroundColor Green
}

function Load-ReplayConfig {
    param(
        [string]$Path
    )
    
    if (!(Test-Path $Path)) {
        throw "Configuration file not found: $Path"
    }
    
    $configData = Get-Content $Path -Raw | ConvertFrom-Json
    
    $config = @{}
    $configData.PSObject.Properties | ForEach-Object {
        $config[$_.Name] = $_.Value
    }
    
    if ($config.ClientSecretEncrypted -and !$config.ClientSecret) {
        try {
            $secureString = ConvertTo-SecureString $config.ClientSecretEncrypted
            $config.ClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
            )
            $config.Remove('ClientSecretEncrypted')
        }
        catch {
            throw "Could not decrypt ClientSecret. If the config was created on a different machine or user account, you'll need to update the ClientSecret."
        }
    }
    
    return $config
}

# ================================
# Date Parameter Hardening
# ================================
function Convert-ToSafeDate {
    param(
        [Parameter(Mandatory)]
        $InputValue,
        [string]$ParamName
    )
    
    # Already a datetime? Return untouched.
    if ($InputValue -is [datetime]) {
        return $InputValue
    }
    
    # Null or empty → return null (for optional params)
    if ($null -eq $InputValue -or [string]::IsNullOrWhiteSpace($InputValue)) {
        return $null
    }
    
    # Normalize input
    $raw = $InputValue.ToString().Trim()
    
    # Supported formats (ISO + German + variations)
    $formats = @(
        'yyyy-MM-dd',
        'yyyy-MM-ddTHH:mm:ss',
        'yyyy-MM-ddTHH:mm:ssZ',
        'yyyy-MM-ddTHH:mm:ss.fffZ',
        'dd.MM.yyyy',
        'dd.MM.yyyy HH:mm',
        'dd.MM.yyyy HH:mm:ss',
        'dd.MM.yyyyTHH:mm:ss',
        'dd.MM.yyyyTHH:mm:ssZ'
    )
    
    # Try explicit formats first (avoids ambiguity)
    foreach ($fmt in $formats) {
        try {
            $parsed = [datetime]::ParseExact($raw, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            Write-Verbose "[$ParamName] Parsed '$raw' using format '$fmt' → $($parsed.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
            return $parsed
        }
        catch {
            # Continue to next format
        }
    }
    
    # Fallback: Try universal parser (handles many ISO variants)
    try {
        $parsed = [datetime]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Verbose "[$ParamName] Parsed '$raw' using universal parser → $($parsed.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
        return $parsed
    }
    catch {
        $errorMsg = @"
Invalid date format for parameter '$ParamName': '$InputValue'

Accepted formats:
  ISO formats:
    • 2025-11-01
    • 2025-11-01T13:05:00
    • 2025-11-01T13:05:00Z
  
  German formats:
    • 01.11.2025
    • 01.11.2025 13:05
    • 01.11.2025 13:05:00

Examples:
  -StartDate '2025-11-01'
  -StartDate '01.11.2025'
  -StartDate (Get-Date '2025-11-01')

"@
        throw $errorMsg
    }
}

# Load configuration if provided
if ($Config) {
    Write-Host "Loading configuration from: $Config" -ForegroundColor Cyan
    try {
        $loadedConfig = Load-ReplayConfig -Path $Config
        
        # Apply loaded configuration (command-line parameters override config file)
        if (!$TenantId -and $loadedConfig.TenantId) { $TenantId = $loadedConfig.TenantId }
        if (!$ClientId -and $loadedConfig.ClientId) { $ClientId = $loadedConfig.ClientId }
        if (!$ClientSecret -and $loadedConfig.ClientSecret) { $ClientSecret = $loadedConfig.ClientSecret }
        if (!$SourceMailboxes -and $loadedConfig.SourceMailboxes) { $SourceMailboxes = $loadedConfig.SourceMailboxes }
        if (!$TargetMailbox -and $loadedConfig.TargetMailbox) { $TargetMailbox = $loadedConfig.TargetMailbox }
        if ($FolderName -eq "Inbox" -and $loadedConfig.FolderName) { $FolderName = $loadedConfig.FolderName }
        if (!$ReplayMode -and $loadedConfig.ReplayMode) { $ReplayMode = $loadedConfig.ReplayMode }
        if (!$BccAlways -and $loadedConfig.BccAlways) { $BccAlways = $loadedConfig.BccAlways }
        if (!$LogPath -and $loadedConfig.LogPath) { $LogPath = $loadedConfig.LogPath }
        if (!$ProcessedHeader -and $loadedConfig.ProcessedHeader) { $ProcessedHeader = $loadedConfig.ProcessedHeader }
        if ($loadedConfig.AttachmentsOnly) { $AttachmentsOnly = $loadedConfig.AttachmentsOnly }
        if ($loadedConfig.SkipAlreadyProcessed) { $SkipAlreadyProcessed = $loadedConfig.SkipAlreadyProcessed }
        if ($loadedConfig.MaxMessages) { $MaxMessages = $loadedConfig.MaxMessages }
        if ($loadedConfig.BatchSize) { $BatchSize = $loadedConfig.BatchSize }
        if ($loadedConfig.ThrottleMs) { $ThrottleMs = $loadedConfig.ThrottleMs }
        
        # Load dates from config if not provided via command line
        if (!$StartDate -and $loadedConfig.StartDate) { 
            $StartDate = Convert-ToSafeDate -InputValue $loadedConfig.StartDate -ParamName 'StartDate (from config)'
        }
        if (!$EndDate -and $loadedConfig.EndDate) { 
            $EndDate = Convert-ToSafeDate -InputValue $loadedConfig.EndDate -ParamName 'EndDate (from config)'
        }
        
        Write-Host "Configuration loaded successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to load configuration: $_"
        throw
    }
    
    # Validate required parameters after loading config
    if (!$TenantId) { throw "TenantId is required (not found in config or parameters)" }
    if (!$ClientId) { throw "ClientId is required (not found in config or parameters)" }
    if (!$ClientSecret) { throw "ClientSecret is required (not found in config or parameters)" }
    if (!$SourceMailboxes) { throw "SourceMailboxes is required (not found in config or parameters)" }
    if (!$TargetMailbox) { throw "TargetMailbox is required (not found in config or parameters)" }
}

# ================================
# Apply Date Validation (for command-line params)
# ================================
if ($PSBoundParameters.ContainsKey('StartDate') -and $StartDate) {
    try {
        $StartDate = Convert-ToSafeDate -InputValue $StartDate -ParamName 'StartDate'
        Write-Host "StartDate parsed: $($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" -ForegroundColor Yellow
    }
    catch {
        Write-Error $_.Exception.Message
        throw
    }
}

if ($PSBoundParameters.ContainsKey('EndDate') -and $EndDate) {
    try {
        $EndDate = Convert-ToSafeDate -InputValue $EndDate -ParamName 'EndDate'
        Write-Host "EndDate parsed:   $($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" -ForegroundColor Yellow
    }
    catch {
        Write-Error $_.Exception.Message
        throw
    }
}

# Cross-validation: StartDate must be before EndDate
if ($StartDate -and $EndDate -and $StartDate -gt $EndDate) {
    $errorMsg = "StartDate ($($StartDate.ToString('yyyy-MM-dd'))) cannot be after EndDate ($($EndDate.ToString('yyyy-MM-dd')))"
    Write-Error $errorMsg
    throw $errorMsg
}

# ================================
# Initialize
# ================================
$ErrorActionPreference = 'Stop'
$global:accessToken = $null
$global:tokenExpiry = [datetime]::MinValue
$processedCount = 0
$errorCount = 0
$skippedCount = 0

# Setup logging
if ($LogPath -and $LogSuccessful) {
    $logFile = if ([System.IO.Path]::IsPathRooted($LogPath)) {
        $LogPath
    } else {
        Join-Path (Get-Location).Path $LogPath
    }
    
    $logDir = [System.IO.Path]::GetDirectoryName($logFile)
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    $logHeader = @"
========================================
Graph Email Replay Log
Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
========================================
"@
    Add-Content -Path $logFile -Value $logHeader
}

# ================================
# Helper Functions
# ================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error"   { Write-Host $logMessage -ForegroundColor Red }
        default   { Write-Host $logMessage -ForegroundColor Cyan }
    }
    
    if ($LogPath -and $LogSuccessful) {
        Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
    }
}

function Get-GraphToken {
    if ($global:accessToken -and $global:tokenExpiry -gt (Get-Date).AddMinutes(5)) {
        return $global:accessToken
    }
    
    Write-Log "Acquiring new Graph API token..." -Level Info
    
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope        = "https://graph.microsoft.com/.default"
        grant_type   = "client_credentials"
    }
    
    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    
    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body
        $global:accessToken = $response.access_token
        $global:tokenExpiry = (Get-Date).AddSeconds($response.expires_in - 300)
        Write-Log "Token acquired successfully (expires: $($global:tokenExpiry))" -Level Success
        return $global:accessToken
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
        [switch]$ReturnHeaders
    )
    
    $token = Get-GraphToken
    $Headers["Authorization"] = "Bearer $token"
    
    if ($Body -and $Method -in @("POST", "PATCH", "PUT")) {
        $Headers["Content-Type"] = "application/json"
        $Body = $Body | ConvertTo-Json -Depth 10 -Compress
    }
    
    $params = @{
        Method  = $Method
        Uri     = $Uri
        Headers = $Headers
    }
    
    if ($Body) {
        $params.Body = $Body
    }
    
    try {
        if ($ReturnHeaders) {
            $response = Invoke-WebRequest @params
            return @{
                Body = $response.Content | ConvertFrom-Json
                Headers = $response.Headers
            }
        }
        else {
            return Invoke-RestMethod @params
        }
    }
    catch {
        $err = $_
        $resp = $null
        if ($err.Exception -and $err.Exception.Response) {
            $resp = $err.Exception.Response
        }

        $status = $null
        if ($resp) {
            try { $status = $resp.StatusCode } catch {}
        }

        if ($status -eq 429) {
            $retryAfter = $null
            try { $retryAfter = $resp.Headers["Retry-After"] } catch {}
            $waitTime = if ($retryAfter) { [int]$retryAfter } else { 60 }
            Write-Log "Throttled. Waiting $waitTime seconds..." -Level Warning
            Start-Sleep -Seconds $waitTime
            return Invoke-GraphRequest @PSBoundParameters
        }

        throw $err
    }
}

function Get-MailboxMessages {
    param(
        [string]$Mailbox,
        [string]$Folder = "Inbox",
        [datetime]$StartDate,
        [datetime]$EndDate,
        [string]$SubjectFilter,
        [switch]$HasAttachments,
        [int]$Top = 100
    )
    
    Write-Log "Fetching messages from $Mailbox/$Folder" -Level Info
    
    $filters = @()
    
    if ($StartDate) {
        $filters += "receivedDateTime ge $($StartDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))"
    }
    
    if ($EndDate) {
        $filters += "receivedDateTime le $($EndDate.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ'))"
    }
    
    if ($SubjectFilter) {
        $filters += "contains(subject, '$SubjectFilter')"
    }
    
    if ($HasAttachments) {
        $filters += "hasAttachments eq true"
    }
    
    $filter = if ($filters) { "&`$filter=" + ($filters -join " and ") } else { "" }
    
    $folderName = switch -Regex ($Folder) {
        "^(Inbox|Posteingang)$" { "inbox" }
        "^(Sent|Gesendete)" { "sentitems" }
        "^(Draft|Entwurf)" { "drafts" }
        "^(Deleted|Gelöscht)" { "deleteditems" }
        default { $Folder }
    }
    
    $uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/mailFolders/$folderName/messages?`$top=$Top&`$select=id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,hasAttachments,internetMessageHeaders$filter"
    
    $allMessages = @()
    
    do {
        $response = Invoke-GraphRequest -Uri $uri
        $allMessages += $response.value
        $uri = $response.'@odata.nextLink'
        
        if ($MaxMessages -and $allMessages.Count -ge $MaxMessages) {
            $allMessages = $allMessages[0..($MaxMessages - 1)]
            break
        }
    } while ($uri)
    
    Write-Log "Found $($allMessages.Count) messages in $Mailbox/$Folder" -Level Info
    return $allMessages
}

function Test-AlreadyProcessed {
    param(
        [object]$Message
    )
    
    if (!$SkipAlreadyProcessed) {
        return $false
    }
    
    foreach ($header in $Message.internetMessageHeaders) {
        if ($header.name -eq $ProcessedHeader) {
            return $true
        }
    }
    
    return $false
}

function Get-MessageMimeContent {
    param(
        [string]$Mailbox,
        [string]$MessageId
    )
    
    $uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/messages/$MessageId/`$value"
    
    $token = Get-GraphToken
    $headers = @{
        "Authorization" = "Bearer $token"
    }
    
    try {
        $response = Invoke-WebRequest -Method GET -Uri $uri -Headers $headers
        return $response.Content
    }
    catch {
        Write-Log "Failed to get MIME content for message $MessageId" -Level Warning
        return $null
    }
}

function Send-TrueTransparentReplay {
    param(
        [string]$SourceMailbox,
        [string]$MessageId,
        [string]$TargetMailbox,
        [string]$ReplayHeader
    )
    
    $mimeContent = Get-MessageMimeContent -Mailbox $SourceMailbox -MessageId $MessageId
    if (-not $mimeContent) { throw "Could not retrieve MIME content for message $MessageId" }

    $mimeBytes  = [System.Text.Encoding]::UTF8.GetBytes($mimeContent)
    $mimeBase64 = [Convert]::ToBase64String($mimeBytes)
    $uriBase    = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages"
    $token      = Get-GraphToken
    $headers    = @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "text/plain"
    }

    $draftResp = Invoke-WebRequest -Method POST -Uri $uriBase -Headers $headers -Body $mimeBase64 -ErrorAction Stop
    $draftId   = ($draftResp.Content | ConvertFrom-Json).id

    if ($ReplayHeader) {
        if ($ReplayHeader -match "^([^:]+):(.*)$") {
            $headerName  = $matches[1].Trim()
            $headerValue = $matches[2].Trim()
            $patchBody = @{
                internetMessageHeaders = @(
                    @{
                        name  = $headerName
                        value = $headerValue
                    }
                )
            } | ConvertTo-Json -Depth 10
            Invoke-GraphRequest -Uri "$uriBase/$draftId" -Method PATCH -Body $patchBody
        }
        else {
            Write-Log "ReplayHeader is not in 'HeaderName:Value' format. Skipping header stamp." -Level Warning
        }
    }

    Invoke-GraphRequest -Uri "$uriBase/$draftId/send" -Method POST
    Write-Log "TRUE TRANSPARENT replay sent (Target: $TargetMailbox, DraftId: $draftId, Source: $SourceMailbox/$MessageId)" -Level Success

    return $draftId
}

function Send-TransparentReplay {
    param(
        [string]$SourceMailbox,
        [string]$MessageId,
        [string]$TargetMailbox,
        [string[]]$BccAddresses
    )
    
    $mimeContent = Get-MessageMimeContent -Mailbox $SourceMailbox -MessageId $MessageId
    if (-not $mimeContent) {
        throw "Could not retrieve MIME content for message $MessageId"
    }

    $additionalHeaders = @"
Resent-Date: $(Get-Date -Format 'r')
Resent-From: $SourceMailbox
Resent-To: $TargetMailbox
Auto-Submitted: auto-generated
X-Resent-Via: GraphAPI/TransparentReplay
${ProcessedHeader}: true

"@

    $finalMime = $additionalHeaders + $mimeContent
    $mimeBytes  = [System.Text.Encoding]::UTF8.GetBytes($finalMime)
    $mimeBase64 = [Convert]::ToBase64String($mimeBytes)

    $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/sendMail"

    $token = Get-GraphToken
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "text/plain"
    }

    $null = Invoke-WebRequest -Method POST -Uri $uri -Headers $headers -Body $mimeBase64

    return "mime-" + ([guid]::NewGuid().ToString())
}

function As-Array {
    param($maybeArray)
    if ($null -eq $maybeArray) { return @() }
    if ($maybeArray -is [System.Collections.IEnumerable] -and $maybeArray -isnot [string]) { return $maybeArray }
    return @($maybeArray)
}

function Fix-Encoding {
    param($str)
    if ($null -eq $str) { return "" }
    $bytes = [System.Text.Encoding]::Default.GetBytes($str)
    try {
        $utf8Str = [System.Text.Encoding]::UTF8.GetString($bytes)
        if ($utf8Str -match '[\u00C0-\u017F]') { return $utf8Str }
    } catch { }
    return $str
}

function Safe-EmailName {
    param(
        [object]$emailObj,
        [string]$fallback = ""
    )
    if ($null -eq $emailObj) { return $fallback }
    $name = ""
    $addr = ""
    try { $name = Fix-Encoding $emailObj.name } catch {}
    try { $addr = $emailObj.address } catch {}
    if ([string]::IsNullOrWhiteSpace($name) -and [string]::IsNullOrWhiteSpace($addr)) {
        return $fallback
    }
    if (-not [string]::IsNullOrWhiteSpace($addr)) {
        if (-not [string]::IsNullOrWhiteSpace($name) -and $name -ne $addr) {
            return "$name &lt;$addr&gt;"
        } else {
            return $addr
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        return $name
    }
    return $fallback
}

function Send-WrapperReplay {
    param(
        [string]$SourceMailbox,
        [object]$Message,
        [string]$TargetMailbox,
        [string[]]$BccAddresses
    )

    $mimeContent = Get-MessageMimeContent -Mailbox $SourceMailbox -MessageId $Message.id
    if (-not $mimeContent) { throw "Could not retrieve MIME content" }
    $mimeBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($mimeContent))

    $uri = "https://graph.microsoft.com/v1.0/users/$SourceMailbox/messages/$($Message.id)?`$expand=attachments"
    $fullMessage = Invoke-GraphRequest -Uri $uri
    $attachments = $fullMessage.attachments
    $originalBody = $fullMessage.body.content
    $bodyType = $fullMessage.body.contentType

    $originalFrom = Safe-EmailName $Message.from.emailAddress
    $originalTo = (As-Array $Message.toRecipients | ForEach-Object {
        Safe-EmailName $_.emailAddress
    }) -join ", "

    # FIX: Use locale-safe datetime parsing
  $formats = @("yyyy-MM-ddTHH:mm:ssZ", "MM/dd/yyyy HH:mm:ss", "dd.MM.yyyy HH:mm:ss")
$culture = [System.Globalization.CultureInfo]::InvariantCulture
try {
    $receivedTime = [datetime]::ParseExact($Message.receivedDateTime, $formats, $culture, [System.Globalization.DateTimeStyles]::AssumeUniversal)
    $receivedStr = $receivedTime.ToString("dd.MM.yyyy HH:mm")
} catch {
    $receivedStr = $Message.receivedDateTime # fallback
}

    $renderedOriginalBody =
        if ($bodyType -eq "HTML") { $originalBody }
        else { [System.Web.HttpUtility]::HtmlEncode($originalBody) }

        $htmlBody = @"
<table border='0' cellpadding='8' bgcolor='#fef3e2' style='border-left:5px solid #f39c12;'>
  <tr>
    <td>
      <b style='color:#d68910;'>&#9888; Diese E-Mail wurde erneut zugestellt</b><br/>
      <div style='margin-top:8px;background:#fff;'>
        <b>Ursprünglicher Absender:</b> $originalFrom<br/>
        <b>Ursprüngliche Empfänger:</b> $originalTo<br/>
        <b>Empfangen am:</b> $receivedTime<br/>
        <b>Betreff:</b> $([System.Web.HttpUtility]::HtmlEncode($Message.subject))
      </div>
      <div style='margin-top:8px;font-size:12px;'>
        <i>Die ursprüngliche E-Mail ist als <b>.eml-Datei</b> im Anhang enthalten.</i><br/>
        <span>Bitte antworten Sie bei Bedarf dem ursprünglichen Absender.</span>
      </div>
    </td>
  </tr>
</table>

<hr>
<b>Ursprüngliche Nachricht:</b><br/>
$renderedOriginalBody
<hr>
<i>Alle ursprünglichen Anhänge sind zusammen mit der .eml-Datei unten angefügt.</i>
"@

<#
    $htmlBody = @"
<table border='0' cellpadding='8' bgcolor='#fef3e2' style='border-left:5px solid #f39c12;'>
  <tr>
    <td>
      <b style='color:#d68910;'>&#9888; This email was replayed</b><br/>
      <div style='margin-top:8px;background:#fff;'>
        <b>Original Sender:</b> $originalFrom<br/>
        <b>Original Recipients:</b> $originalTo<br/>
        <b>Received:</b> $receivedTime<br/>
        <b>Subject:</b> $([System.Web.HttpUtility]::HtmlEncode($Message.subject))
      </div>
      <div style='margin-top:8px;font-size:12px;'>
        <i>The original email is attached as <b>.eml</b> file.</i><br/>
        <span>Please reply to the original sender if needed.</span>
      </div>
    </td>
  </tr>
</table>

<hr>
<b>Original message:</b><br/>
$renderedOriginalBody
<hr>
<i>All original attachments are attached below along with the .eml file.</i>
"@

#>

    $attachmentsList = @()

    foreach ($att in $attachments) {
        if ($att.'@odata.type' -eq "#microsoft.graph.fileAttachment") {
            $attachmentsList += @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name = $att.name
                contentType = $att.contentType
                contentBytes = $att.contentBytes
            }
        }
    }

    $attachmentsList += @{
        "@odata.type" = "#microsoft.graph.fileAttachment"
        name = "Original_Message.eml"
        contentType = "message/rfc822"
        contentBytes = $mimeBase64
    }

    $newMessage = @{
        subject = "[Weiterleitung] $($Message.subject)"
        body = @{
            contentType = "HTML"
            content = $htmlBody
        }
        toRecipients = @(@{ emailAddress = @{ address = $TargetMailbox } })
        attachments = $attachmentsList
        importance = "normal"
        internetMessageHeaders = @(
            @{ name = $ProcessedHeader; value = "true" }
            @{ name = "X-Original-MessageId"; value = $Message.id }
        )
    }

    if ($BccAddresses) {
        $newMessage.bccRecipients =
            $BccAddresses | ForEach-Object {
                @{ emailAddress = @{ address = $_ } }
            }
    }

    $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/sendMail"
    $body = @{
        message = $newMessage
        saveToSentItems = $false
    }

    Invoke-GraphRequest -Uri $uri -Method POST -Body $body

    return "wrapper-$(New-Guid)"
}

function Send-TestEmail {
    param(
        [string]$TestMailbox,
        [string]$TargetMailbox
    )
    
    $testMessage = @{
        subject = "[TEST] Graph Email Replay - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        body = @{
            contentType = "HTML"
            content = @"
<div style='font-family:Segoe UI,Arial'>
    <h2>Test Email - Graph Replay Script</h2>
    <p>This is a test message from the Graph Email Replay script.</p>
    <ul>
        <li><strong>Timestamp:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</li>
        <li><strong>Source Mailbox:</strong> $TestMailbox</li>
        <li><strong>Target Mailbox:</strong> $TargetMailbox</li>
        <li><strong>Replay Mode:</strong> $ReplayMode</li>
    </ul>
    <p style='color:#28a745;'>If you received this message, the configuration is working correctly.</p>
</div>
"@
        }
        toRecipients = @(
            @{
                emailAddress = @{
                    address = $TargetMailbox
                }
            }
        )
        importance = "normal"
    }
    
    $uri = "https://graph.microsoft.com/v1.0/users/$TestMailbox/sendMail"
    $body = @{
        message = $testMessage
        saveToSentItems = $true
    }
    
    try {
        Invoke-GraphRequest -Uri $uri -Method POST -Body $body
        Write-Log "Test email sent successfully from $TestMailbox to $TargetMailbox" -Level Success
        return $true
    }
    catch {
        Write-Log "Test email failed: $_" -Level Error
        return $false
    }
}

# ================================
# Main Processing
# ================================

try {
    Write-Log "=== Graph Email Replay Configuration ===" -Level Info
    Write-Log "Tenant ID: $TenantId" -Level Info
    Write-Log "Client ID: $ClientId" -Level Info
    Write-Log "Target Mailbox: $TargetMailbox" -Level Info
    Write-Log "Replay Mode: $ReplayMode" -Level Info
    Write-Log "Source Mailboxes: $($SourceMailboxes -join ', ')" -Level Info
    
    if ($AttachmentsOnly) {
        Write-Log "Filter: Attachments Only" -Level Info
    }
    
    if ($StartDate -or $EndDate) {
        Write-Log "Date Range: $StartDate to $EndDate" -Level Info
    }
    
    if ($SubjectFilter) {
        Write-Log "Subject Filter: $SubjectFilter" -Level Info
    }
    
    if ($BccAlways) {
        Write-Log "BCC Always: $($BccAlways -join ', ')" -Level Info
    }
    
    if ($WhatIf) {
        Write-Log "*** WHATIF MODE - No emails will be sent ***" -Level Warning
    }
    
    if ($TestMode) {
        $testSource = if ($TestMailbox) { $TestMailbox } else { $SourceMailboxes[0] }
        Write-Log "Running in TEST MODE - Sending test email only" -Level Warning
        
        if (Send-TestEmail -TestMailbox $testSource -TargetMailbox $TargetMailbox) {
            Write-Log "Test completed successfully" -Level Success
        }
        else {
            Write-Log "Test failed" -Level Error
        }
        
        return
    }
    
    foreach ($sourceMailbox in $SourceMailboxes) {
        Write-Log "`nProcessing mailbox: $sourceMailbox" -Level Info
        
        try {
            $messages = Get-MailboxMessages `
                -Mailbox $sourceMailbox `
                -Folder $FolderName `
                -StartDate $StartDate `
                -EndDate $EndDate `
                -SubjectFilter $SubjectFilter `
                -HasAttachments:$AttachmentsOnly `
                -Top $BatchSize
            
            if ($messages.Count -eq 0) {
                Write-Log "No messages found in $sourceMailbox/$FolderName" -Level Warning
                continue
            }
            
            Write-Log "Processing $($messages.Count) messages from $sourceMailbox" -Level Info
            
            foreach ($message in $messages) {

    # 1) Skip messages tagged with ProcessedHeader (Graph header in original message)
    if (Test-AlreadyProcessed -Message $message) {
        Write-Log "Skipping (already processed): $($message.subject)" -Level Info
        $skippedCount++
        continue
    }

    # 2) DEDUP: skip if MessageId + TargetMailbox already sent in previous runs
    if (Already-Sent $message.id $TargetMailbox) {
        Write-Log "Skipping (already resent): $($message.subject) to $TargetMailbox [MessageId: $($message.id)]" -Level Warning
        $skippedCount++
        continue
    }

    # 3) Max message limit
    if ($MaxMessages -and $processedCount -ge $MaxMessages) {
        Write-Log "Reached maximum message limit ($MaxMessages)" -Level Warning
        break
    }

    $action = if ($WhatIf) { "[WHATIF]" } else { "[SENDING]" }
    Write-Log "$action $($message.subject) (from: $($message.from.emailAddress.address))" -Level Info

    if (-not $WhatIf) {
        try {

            # Fetch attachments if ReplayMode = Wrapper
            $attachments = @()
            if ($ReplayMode -eq "Wrapper") {
                $uri = "https://graph.microsoft.com/v1.0/users/$sourceMailbox/messages/$($message.id)?`$expand=attachments"
                $fullMessage = Invoke-GraphRequest -Uri $uri
                $attachments = $fullMessage.attachments
            }

            # SEND based on mode
            $sentId = if ($ReplayMode -eq "Transparent") {
                if ($TrueTransparentReplay) {
                    Send-TrueTransparentReplay `
                        -SourceMailbox $sourceMailbox `
                        -MessageId $message.id `
                        -TargetMailbox $TargetMailbox `
                        -ReplayHeader $ReplayHeader
                }
                else {
                    Send-TransparentReplay `
                        -SourceMailbox $sourceMailbox `
                        -MessageId $message.id `
                        -TargetMailbox $TargetMailbox `
                        -BccAddresses $BccAlways
                }
            }
            else {
                Send-WrapperReplay `
                    -SourceMailbox $sourceMailbox `
                    -Message $message `
                    -TargetMailbox $TargetMailbox `
                    -BccAddresses $BccAlways
            }

            $processedCount++

            # 4) Per-mailbox CSV logging
            $csvPath = Get-LogFilePath $sourceMailbox
            $csvLine = [PSCustomObject]@{
                Timestamp     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                MessageId     = $message.id
                Subject       = $message.subject
                From          = $message.from.emailAddress.address
                To            = $TargetMailbox
                ReplayMode    = $ReplayMode
                Attachments   = ($attachments | Where-Object { $_.name } | ForEach-Object { $_.name }) -join ';'
                Status        = "SENT"
            }
            $writeHeader = -not (Test-Path $csvPath)
            $csvLine | Export-Csv -Path $csvPath -Append -NoTypeInformation -Encoding UTF8
            if ($writeHeader) {
                (Get-Content $csvPath) | Set-Content $csvPath
            }

            # 5) Dedup marker creation
            Mark-Sent $message.id $TargetMailbox

            Write-Log "Successfully sent: $($message.subject)" -Level Success

            if ($ThrottleMs -gt 0) {
                Start-Sleep -Milliseconds $ThrottleMs
            }

        }
        catch {
            $errorCount++
            Write-Log "Failed to send: $($message.subject) - Error: $_" -Level Error

            # 6) Mailbox-specific error CSV
            $errorCsv = Join-Path $LogRoot "$($sourceMailbox -replace '[^\w\-\.@]', '_')_errors.csv"
            $attNames = ($attachments | Where-Object { $_.name } | ForEach-Object { $_.name }) -join ';'

            $errRecord = [PSCustomObject]@{
                Timestamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                MessageId   = $message.id
                Subject     = $message.subject
                From        = $message.from.emailAddress.address
                To          = $TargetMailbox
                Attachments = $attNames
                Error       = $_.ToString()
            }

            $writeHeader = -not (Test-Path $errorCsv)
            $errRecord | Export-Csv -Path $errorCsv -Append -NoTypeInformation -Encoding UTF8
            if ($writeHeader) {
                (Get-Content $errorCsv) | Set-Content $errorCsv
            }
        }
    }
    else {
        $processedCount++
    }
}
        }
        catch {
            Write-Log "Error processing mailbox $sourceMailbox : $_" -Level Error
        }
    }
    
    Write-Log "`n=== Summary ===" -Level Info
    Write-Log "Total Processed: $processedCount" -Level Info
    Write-Log "Total Skipped:   $skippedCount" -Level Info
    Write-Log "Total Errors:    $errorCount" -Level Info
    Write-Log "Script completed." -Level Info
}
catch {
    Write-Log "Fatal error: $_" -Level Error
    throw
}   