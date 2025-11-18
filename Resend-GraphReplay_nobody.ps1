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
    [datetime]$StartDate,                # Optional date range start
    [datetime]$EndDate,                  # Optional date range end
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
    $BccAlways,                          # Always BCC these addresses (string or array)
    [switch]$SkipAlreadyProcessed,       # Skip if custom header exists
    [string]$ProcessedHeader = "X-GraphReplay-Processed",
    [int]$BatchSize = 50,                # Graph API batch size
    [int]$ThrottleMs = 100               # Throttle between sends
)

# ================================
# Configuration File Management
# ================================

function Save-ReplayConfig {
    param(
        [string]$Path,
        [hashtable]$Configuration
    )
    
    # Remove sensitive data if saving template
    $safeConfig = $Configuration.Clone()
    
    # Encrypt sensitive values if possible
    if ($Configuration.ClientSecret) {
        try {
            $secureString = ConvertTo-SecureString $Configuration.ClientSecret -AsPlainText -Force
            $encryptedSecret = ConvertFrom-SecureString $secureString
            $safeConfig.ClientSecretEncrypted = $encryptedSecret
            $safeConfig.Remove('ClientSecret')
        }
        catch {
            # If encryption fails, warn user
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
    
    # Convert PSCustomObject to Hashtable
    $config = @{}
    $configData.PSObject.Properties | ForEach-Object {
        $config[$_.Name] = $_.Value
    }
    
    # Decrypt sensitive values if encrypted
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
        if (!$BccAlways -and $loadedConfig.BccAlways) { 
            # Handle BccAlways from config - could be string or array
            if ($loadedConfig.BccAlways -is [string]) {
                $BccAlways = @($loadedConfig.BccAlways)
            } else {
                $BccAlways = $loadedConfig.BccAlways
            }
        }
        if (!$LogPath -and $loadedConfig.LogPath) { $LogPath = $loadedConfig.LogPath }
        if (!$ProcessedHeader -and $loadedConfig.ProcessedHeader) { $ProcessedHeader = $loadedConfig.ProcessedHeader }
        if ($loadedConfig.AttachmentsOnly) { $AttachmentsOnly = $loadedConfig.AttachmentsOnly }
        if ($loadedConfig.SkipAlreadyProcessed) { $SkipAlreadyProcessed = $loadedConfig.SkipAlreadyProcessed }
        if ($loadedConfig.MaxMessages) { $MaxMessages = $loadedConfig.MaxMessages }
        if ($loadedConfig.BatchSize) { $BatchSize = $loadedConfig.BatchSize }
        if ($loadedConfig.ThrottleMs) { $ThrottleMs = $loadedConfig.ThrottleMs }
        
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
    
    # Ensure directory exists
    $logDir = [System.IO.Path]::GetDirectoryName($logFile)
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Initialize log with timestamp
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
    
    # Console output with color
    switch ($Level) {
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error"   { Write-Host $logMessage -ForegroundColor Red }
        default   { Write-Host $logMessage -ForegroundColor Cyan }
    }
    
    # File logging if enabled
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
        if ($_.Exception.Response.StatusCode -eq 429) {
            # Handle throttling
            $retryAfter = $_.Exception.Response.Headers["Retry-After"]
            $waitTime = if ($retryAfter) { [int]$retryAfter } else { 60 }
            Write-Log "Throttled. Waiting $waitTime seconds..." -Level Warning
            Start-Sleep -Seconds $waitTime
            return Invoke-GraphRequest @PSBoundParameters
        }
        throw
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
    
    # Build filter
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
    
    # Folder name normalization
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

function Send-TransparentReplay {
    param(
        [string]$SourceMailbox,
        [string]$MessageId,
        [string]$TargetMailbox,
        [string[]]$BccAddresses
    )
    
    # Get MIME content
    $mimeContent = Get-MessageMimeContent -Mailbox $SourceMailbox -MessageId $MessageId
    
    if (!$mimeContent) {
        throw "Could not retrieve MIME content"
    }
    
    # Convert to base64
    $mimeBytes = [System.Text.Encoding]::UTF8.GetBytes($mimeContent)
    
    # Add resent headers
    $headers = @"
Resent-Date: $(Get-Date -Format 'r')
Resent-From: $SourceMailbox
Resent-To: $TargetMailbox
Auto-Submitted: auto-generated
X-Resent-Via: GraphAPI/TransparentReplay
$ProcessedHeader: true

"@
    
    $headerBytes = [System.Text.Encoding]::UTF8.GetBytes($headers)
    $combinedBytes = $headerBytes + $mimeBytes
    $mimeBase64 = [Convert]::ToBase64String($combinedBytes)
    
    # Send using Graph API raw send
    $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/sendMail"
    
    $body = @{
        message = @{
            internetMessageId = "<replay-$(New-Guid)@graphreplay>"
        }
    }
    
    # Use MIME format
    $mimeBody = @"
{
    "message": {
        "@odata.type": "#microsoft.graph.message"
    },
    "saveToSentItems": false
}
"@
    
    # Actually, for raw MIME we need different approach
    # Using the /messages endpoint with MIME
    $createUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages"
    
    $token = Get-GraphToken
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "text/plain"
        "Prefer" = "IdType='ImmutableId'"
    }
    
    # Create message from MIME
    $response = Invoke-WebRequest -Method POST -Uri $createUri -Headers $headers -Body $mimeContent
    $createdMessage = $response.Content | ConvertFrom-Json
    
    # Send the created message
    $sendUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$($createdMessage.id)/send"
    Invoke-GraphRequest -Uri $sendUri -Method POST
    
    return $createdMessage.id
}

function Send-WrapperReplay {
    param(
        [string]$SourceMailbox,
        [object]$Message,
        [string]$TargetMailbox,
        $BccAddresses  # Can be string, array, or null
    )
    
    # Get original MIME as attachment
    $mimeContent = Get-MessageMimeContent -Mailbox $SourceMailbox -MessageId $Message.id
    
    if (!$mimeContent) {
        throw "Could not retrieve MIME content"
    }
    
    $mimeBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($mimeContent))
    
    # Build wrapper email
    $originalFrom = if ($Message.from.emailAddress.address) { 
        "$($Message.from.emailAddress.name) <$($Message.from.emailAddress.address)>" 
    } else { 
        "Unknown Sender" 
    }
    
    $originalTo = ($Message.toRecipients | ForEach-Object { $_.emailAddress.address }) -join ", "
    
    # Simple date parsing
    $receivedStr = "Unknown"
    try {
        $receivedTime = [datetime]::Parse($Message.receivedDateTime)
        $receivedStr = $receivedTime.ToString("dd.MM.yyyy HH:mm")
    } catch {
        $receivedStr = $Message.receivedDateTime
    }
    
    # German template with simpler HTML
    $htmlBody = @"
<table border='0' cellpadding='8' bgcolor='#fef3e2' style='border-left:5px solid #f39c12;'>
  <tr>
    <td>
      <b style='color:#d68910;'>&#9888; Diese E-Mail wurde erneut zugestellt</b><br/>
      <div style='margin-top:8px;background:#fff;padding:8px;'>
        <b>Ursprünglicher Absender:</b> $originalFrom<br/>
        <b>Ursprüngliche Empfänger:</b> $originalTo<br/>
        <b>Empfangen am:</b> $receivedStr<br/>
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
<p><i>Ursprüngliche Nachricht siehe Anhang.</i></p>
"@
    
    # Build recipients
    $toRecipients = @(@{
        emailAddress = @{
            address = $TargetMailbox
        }
    })
    
    # SIMPLIFIED BCC LOGIC - exactly like working version
    $bccRecipients = @()
    if ($BccAddresses) {
        # Handle both string and array
        $bccList = @()
        if ($BccAddresses -is [string]) {
            $bccList = @($BccAddresses)
        } elseif ($BccAddresses -is [array]) {
            $bccList = $BccAddresses
        }
        
        foreach ($bcc in $bccList) {
            if ($bcc) {  # Just check not null/empty
                $bccRecipients += @{
                    emailAddress = @{
                        address = $bcc
                    }
                }
            }
        }
    }
    
    # Create message
    $newMessage = @{
        subject = "[Weiterleitung] $($Message.subject)"
        body = @{
            contentType = "HTML"
            content = $htmlBody
        }
        toRecipients = $toRecipients
        importance = "normal"
        attachments = @(
            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                name = "Original_Message.eml"
                contentType = "message/rfc822"
                contentBytes = $mimeBase64
            }
        )
        internetMessageHeaders = @(
            @{
                name = $ProcessedHeader
                value = "true"
            }
            @{
                name = "X-Original-MessageId"
                value = $Message.id
            }
        )
    }
    
    # Only add bccRecipients if we have some
    if ($bccRecipients.Count -gt 0) {
        $newMessage.bccRecipients = $bccRecipients
    }
    
    # Send
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
        Write-Log "Failed to send test email: $_" -Level Error
        return $false
    }
}

# ================================
# Main Processing
# ================================

try {
    # Display configuration
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
    
    if ($BccAlways -and $BccAlways.Count -gt 0) {
        Write-Log "BCC Always: $($BccAlways -join ', ')" -Level Info
    }
    
    if ($WhatIf) {
        Write-Log "*** WHATIF MODE - No emails will be sent ***" -Level Warning
    }
    
    # Test mode
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
    
    # Process each source mailbox
    foreach ($sourceMailbox in $SourceMailboxes) {
        Write-Log "`nProcessing mailbox: $sourceMailbox" -Level Info
        
        try {
            # Get messages
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
                # Check if already processed
                if (Test-AlreadyProcessed -Message $message) {
                    Write-Log "Skipping (already processed): $($message.subject)" -Level Info
                    $skippedCount++
                    continue
                }
                
                # Check max messages limit
                if ($MaxMessages -and $processedCount -ge $MaxMessages) {
                    Write-Log "Reached maximum message limit ($MaxMessages)" -Level Warning
                    break
                }
                
                # Display what we're doing
                $action = if ($WhatIf) { "[WHATIF]" } else { "[SENDING]" }
                Write-Log "$action $($message.subject) (from: $($message.from.emailAddress.address))" -Level Info
                
                if (!$WhatIf) {
                    try {
                        # Send based on mode
                        $sentId = if ($ReplayMode -eq "Transparent") {
                            Send-TransparentReplay `
                                -SourceMailbox $sourceMailbox `
                                -MessageId $message.id `
                                -TargetMailbox $TargetMailbox `
                                -BccAddresses $BccAlways
                        }
                        else {
                            Send-WrapperReplay `
                                -SourceMailbox $sourceMailbox `
                                -Message $message `
                                -TargetMailbox $TargetMailbox `
                                -BccAddresses $BccAlways
                        }
                        
                        $processedCount++
                        
                        # Log success
                        if ($LogSuccessful) {
                            $logEntry = @{
                                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                                SourceMailbox = $sourceMailbox
                                MessageId = $message.id
                                Subject = $message.subject
                                From = $message.from.emailAddress.address
                                ReceivedDate = $message.receivedDateTime
                                TargetMailbox = $TargetMailbox
                                ReplayMode = $ReplayMode
                                SentId = $sentId
                            }
                            
                            $logEntry | ConvertTo-Json -Compress | Add-Content -Path $logFile
                        }
                        
                        Write-Log "Successfully sent: $($message.subject)" -Level Success
                        
                        # Throttle
                        if ($ThrottleMs -gt 0) {
                            Start-Sleep -Milliseconds $ThrottleMs
                        }
                    }
                    catch {
                        $errorCount++
                        Write-Log "Failed to send: $($message.subject) - Error: $_" -Level Error
                        
                        # Log error
                        if ($LogPath) {
                            $errorEntry = @{
                                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                                SourceMailbox = $sourceMailbox
                                MessageId = $message.id
                                Subject = $message.subject
                                Error = $_.ToString()
                            }
                            
                            $errorEntry | ConvertTo-Json -Compress | 
                                Add-Content -Path ($logFile -replace '\.log$', '_errors.log')
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
    
    # Final summary
    Write-Log "`n========================================" -Level Info
    Write-Log "Processing Complete" -Level Success
    Write-Log "========================================" -Level Info
    Write-Log "Processed: $processedCount messages" -Level Info
    Write-Log "Skipped: $skippedCount messages" -Level Info
    Write-Log "Errors: $errorCount messages" -Level Info
    
    if ($LogPath -and $LogSuccessful) {
        Write-Log "Log file: $logFile" -Level Info
    }
    
    # Final summary
    Write-Log "`n========================================" -Level Info
    Write-Log "Processing Complete" -Level Success
    Write-Log "========================================" -Level Info
    Write-Log "Processed: $processedCount messages" -Level Info
    Write-Log "Skipped: $skippedCount messages" -Level Info
    Write-Log "Errors: $errorCount messages" -Level Info
    
    if ($LogPath -and $LogSuccessful) {
        Write-Log "Log file: $logFile" -Level Info
        
        # Write footer to log
        $footer = @"

========================================
Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Total Processed: $processedCount
Total Errors: $errorCount
Total Skipped: $skippedCount
========================================
"@
        Add-Content -Path $logFile -Value $footer -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Log "Fatal error: $_" -Level Error
    
    # Log fatal error if possible
    if ($LogPath) {
        $errorEntry = @{
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            FatalError = $_.ToString()
            ScriptStack = $_.ScriptStackTrace
        }
        $errorEntry | ConvertTo-Json -Compress | 
            Add-Content -Path ($logFile -replace '\.log$', '_fatal.log') -ErrorAction SilentlyContinue
    }
    
    throw
}