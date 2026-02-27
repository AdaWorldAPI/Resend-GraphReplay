<#
.SYNOPSIS
    Kopano IMAP to Microsoft 365 Migration (Graph MIME Import)
    
.DESCRIPTION
    Migrates emails using Graph API with true MIME import.
    Tests both methods to find which one creates proper emails (not drafts).
    
    Method A: Direct MIME POST with Content-Type: text/plain
    Method B: Import Session API (backup/restore style)
    
.NOTES
    Requires: MailboxItem.ImportExport.All (Application permission)
    Or: Mail.ReadWrite (Application permission)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter(Mandatory)]
    [string]$ClientSecret,

    [Parameter(Mandatory)]
    [string]$ImapServer,

    [int]$ImapPort = 993,
    [switch]$ImapUseSsl = $true,
    [switch]$ImapSkipCertValidation,

    [string]$UserCsvPath,

    [string]$TestSource,
    [string]$TestTarget,
    [string]$TestUsername,
    [string]$TestPassword,
    [switch]$TestMode,

    [string[]]$FoldersToMigrate,
    [string[]]$ExcludeFolders = @("Junk", "Spam", "Trash", "Drafts", "Entwürfe", "Papierkorb", "Junk-E-Mail"),

    [datetime]$StartDate,
    [datetime]$EndDate,
    [int]$MaxMessagesPerMailbox,

    # === Import Method Selection ===
    [ValidateSet("MethodA", "MethodB", "Both")]
    [string]$ImportMethod = "Both",

    [int]$ThrottleMs = 200,
    [switch]$WhatIf,
    [switch]$ContinueOnError,

    [string]$LogPath = ".\migration_log",
    [switch]$VerboseLogging
)

$ErrorActionPreference = 'Stop'
$script:accessToken = $null
$script:tokenExpiry = [datetime]::MinValue

# ================================
# Logging
# ================================
$script:logFile = $null

function Initialize-Logging {
    if (!(Test-Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:logFile = Join-Path $LogPath "migration_graph_mime_$timestamp.log"
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

    if ($script:logFile) {
        Add-Content -Path $script:logFile -Value $logMessage -ErrorAction SilentlyContinue
    }
}

# ================================
# MailKit Loading
# ================================

function Initialize-MailKit {
    $libPath = Join-Path $PSScriptRoot "lib"
    $dlls = @("BouncyCastle.Crypto.dll", "MimeKit.dll", "MailKit.dll")
    
    foreach ($dll in $dlls) {
        $dllPath = Join-Path $libPath $dll
        if (!(Test-Path $dllPath)) {
            Write-Log "Missing: $dllPath" -Level Error
            return $false
        }
        
        $assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($dll)
        $loaded = [System.AppDomain]::CurrentDomain.GetAssemblies() | 
            Where-Object { $_.GetName().Name -eq $assemblyName }
        
        if (!$loaded) {
            try { Add-Type -Path $dllPath -ErrorAction Stop }
            catch [System.Reflection.ReflectionTypeLoadException] { }
            catch { Write-Log "Failed to load $dll : $_" -Level Error; return $false }
        }
    }
    
    Write-Log "MailKit loaded" -Level Success
    return $true
}

# ================================
# OAuth Token
# ================================

function Get-GraphToken {
    if ($script:accessToken -and $script:tokenExpiry -gt (Get-Date).AddMinutes(5)) {
        return $script:accessToken
    }

    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body
    $script:accessToken = $response.access_token
    $script:tokenExpiry = (Get-Date).AddSeconds($response.expires_in - 300)
    
    return $script:accessToken
}

# ================================
# Graph Folder Management
# ================================

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

    $token = Get-GraphToken
    $headers = @{ "Authorization" = "Bearer $token" }

    $normalizedPath = $FolderPath -replace '/', '\' -replace '\\+', '\'
    $parts = $normalizedPath.Split('\') | Where-Object { $_ -ne '' }

    $folderMapping = @{
        'INBOX'              = 'Inbox'
        'Sent'               = 'SentItems'
        'Sent Items'         = 'SentItems'
        'Gesendete Objekte'  = 'SentItems'
        'Gesendete Elemente' = 'SentItems'
        'Drafts'             = 'Drafts'
        'Entwürfe'           = 'Drafts'
        'Trash'              = 'DeletedItems'
        'Deleted Items'      = 'DeletedItems'
        'Gelöschte Objekte'  = 'DeletedItems'
        'Gelöschte Elemente' = 'DeletedItems'
        'Archive'            = 'Archive'
        'Archiv'             = 'Archive'
    }

    $currentParentId = $null
    $currentFolderId = $null

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $folderName = $parts[$i]

        if ($i -eq 0 -and $folderMapping.ContainsKey($folderName)) {
            $wellKnownName = $folderMapping[$folderName]
            try {
                $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$wellKnownName"
                $folder = Invoke-RestMethod -Uri $uri -Headers $headers
                $currentFolderId = $folder.id
                $currentParentId = $folder.id
                continue
            }
            catch { }
        }

        $searchUri = if ($currentParentId) {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$currentParentId/childFolders?`$filter=displayName eq '$folderName'"
        } else {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders?`$filter=displayName eq '$folderName'"
        }

        try {
            $existingFolders = Invoke-RestMethod -Uri $searchUri -Headers $headers
            if ($existingFolders.value -and $existingFolders.value.Count -gt 0) {
                $currentFolderId = $existingFolders.value[0].id
                $currentParentId = $currentFolderId
                continue
            }
        }
        catch { }

        $createUri = if ($currentParentId) {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$currentParentId/childFolders"
        } else {
            "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders"
        }

        try {
            $body = @{ displayName = $folderName } | ConvertTo-Json
            $created = Invoke-RestMethod -Uri $createUri -Method POST -Headers $headers -Body $body -ContentType "application/json"
            $currentFolderId = $created.id
            $currentParentId = $currentFolderId
        }
        catch {
            $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/Inbox"
            $inbox = Invoke-RestMethod -Uri $uri -Headers $headers
            $currentFolderId = $inbox.id
        }
    }

    $FolderCache[$cacheKey] = $currentFolderId
    return $currentFolderId
}

# ================================
# METHOD A: Direct MIME POST to /messages with Base64
# Graph API expects MIME content as Base64-encoded string
# ================================

function Import-MimeMethodA {
    param(
        [string]$TargetMailbox,
        [string]$FolderId,
        [byte[]]$MimeContent,
        [bool]$IsRead = $true
    )

    $token = Get-GraphToken
    
    # MIME import goes to /messages directly
    $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages"
    
    # BASE64 ENCODE the MIME content - this is what Graph expects!
    $mimeBase64 = [Convert]::ToBase64String($MimeContent)
    
    Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
    
    $httpClient = New-Object System.Net.Http.HttpClient
    $httpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $token)
    $httpClient.Timeout = [TimeSpan]::FromMinutes(5)
    
    # Send Base64-encoded MIME with text/plain content type
    $content = New-Object System.Net.Http.StringContent($mimeBase64, [System.Text.Encoding]::UTF8, "text/plain")
    
    $response = $httpClient.PostAsync($uri, $content).Result
    $responseBody = $response.Content.ReadAsStringAsync().Result
    
    $httpClient.Dispose()
    
    if ($response.IsSuccessStatusCode) {
        $createdMessage = $responseBody | ConvertFrom-Json
        $messageId = $createdMessage.id
        
        # Move to target folder if specified
        if ($FolderId) {
            try {
                $moveUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId/move"
                $headers = @{ "Authorization" = "Bearer $token" }
                $moveBody = @{ destinationId = $FolderId } | ConvertTo-Json
                $moved = Invoke-RestMethod -Uri $moveUri -Method POST -Headers $headers -Body $moveBody -ContentType "application/json"
                $messageId = $moved.id
            }
            catch {
                Write-Log "    Move failed: $_" -Level Warning
            }
        }
        
        # Mark as read if needed
        if ($IsRead) {
            try {
                $updateUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId"
                $headers = @{ "Authorization" = "Bearer $token" }
                Invoke-RestMethod -Uri $updateUri -Method PATCH -Headers $headers -Body '{"isRead":true}' -ContentType "application/json" | Out-Null
            }
            catch { }
        }
        
        return @{
            Success = $true
            MessageId = $messageId
            IsDraft = $createdMessage.isDraft
            Subject = $createdMessage.subject
            From = $createdMessage.from.emailAddress.address
        }
    }
    else {
        return @{
            Success = $false
            Error = "$($response.StatusCode): $responseBody"
        }
    }
}

# ================================
# METHOD B: Try different approaches - JSON body with mimeContent property
# This is how Graph API documentation suggests for programmatic creation
# ================================

function Import-MimeMethodB {
    param(
        [string]$TargetMailbox,
        [string]$FolderId,
        [byte[]]$MimeContent,
        [bool]$IsRead = $true
    )

    $token = Get-GraphToken
    $headers = @{ 
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
    }
    
    # Base64 encode the MIME
    $mimeBase64 = [Convert]::ToBase64String($MimeContent)
    
    # Method B: Use PATCH with mimeContent property on a draft
    # First create an empty draft, then update with MIME content
    
    try {
        # Step 1: Create minimal draft
        $draftBody = @{
            subject = "Temp"
        } | ConvertTo-Json
        
        $uri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages"
        $draft = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $draftBody -ContentType "application/json"
        $draftId = $draft.id
        
        # Step 2: Update with MIME content using $value endpoint
        # The $value endpoint accepts raw content
        $valueUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$draftId/`$value"
        
        Add-Type -AssemblyName System.Net.Http -ErrorAction SilentlyContinue
        $httpClient = New-Object System.Net.Http.HttpClient
        $httpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $token)
        
        # PUT raw MIME bytes to $value endpoint
        $byteContent = New-Object System.Net.Http.ByteArrayContent(,$MimeContent)
        $byteContent.Headers.ContentType = New-Object System.Net.Http.Headers.MediaTypeHeaderValue("text/plain")
        
        $response = $httpClient.PutAsync($valueUri, $byteContent).Result
        $responseBody = $response.Content.ReadAsStringAsync().Result
        
        $httpClient.Dispose()
        
        if ($response.IsSuccessStatusCode) {
            # Get updated message
            $getUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$draftId"
            $updatedMessage = Invoke-RestMethod -Uri $getUri -Headers $headers
            
            $messageId = $updatedMessage.id
            
            # Move to target folder
            if ($FolderId) {
                try {
                    $moveUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId/move"
                    $moveBody = @{ destinationId = $FolderId } | ConvertTo-Json
                    $moved = Invoke-RestMethod -Uri $moveUri -Method POST -Headers $headers -Body $moveBody -ContentType "application/json"
                    $messageId = $moved.id
                }
                catch { }
            }
            
            # Mark as read
            if ($IsRead) {
                try {
                    $updateUri = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$messageId"
                    Invoke-RestMethod -Uri $updateUri -Method PATCH -Headers $headers -Body '{"isRead":true}' -ContentType "application/json" | Out-Null
                }
                catch { }
            }
            
            return @{
                Success = $true
                MessageId = $messageId
                IsDraft = $updatedMessage.isDraft
                Subject = $updatedMessage.subject
                From = $updatedMessage.from.emailAddress.address
            }
        }
        else {
            # Cleanup draft
            try { Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$TargetMailbox/messages/$draftId" -Method DELETE -Headers $headers | Out-Null } catch { }
            
            return @{
                Success = $false
                Error = "$($response.StatusCode): $responseBody"
            }
        }
    }
    catch {
        return @{
            Success = $false
            Error = $_.Exception.Message
        }
    }
}

# ================================
# Combined Import Test
# ================================

function Import-MessageToGraph {
    param(
        [string]$TargetMailbox,
        [string]$FolderId,
        [byte[]]$MimeContent,
        [datetime]$ReceivedDate,
        [bool]$IsRead = $true,
        [string]$Method = "Both"
    )

    $results = @{}

    # Try Method A
    if ($Method -in @("MethodA", "Both")) {
        Write-Log "  Trying Method A (text/plain MIME string)..." -Level Debug
        $resultA = Import-MimeMethodA -TargetMailbox $TargetMailbox -FolderId $FolderId -MimeContent $MimeContent -IsRead $IsRead
        $results["MethodA"] = $resultA
        
        if ($resultA.Success) {
            Write-Log "  Method A: SUCCESS - isDraft=$($resultA.IsDraft), From=$($resultA.From)" -Level Info
        }
        else {
            Write-Log "  Method A: FAILED - $($resultA.Error)" -Level Warning
        }
    }

    # Try Method B
    if ($Method -in @("MethodB", "Both")) {
        Write-Log "  Trying Method B (octet-stream / rfc822)..." -Level Debug
        $resultB = Import-MimeMethodB -TargetMailbox $TargetMailbox -FolderId $FolderId -MimeContent $MimeContent -IsRead $IsRead
        $results["MethodB"] = $resultB
        
        if ($resultB.Success) {
            Write-Log "  Method B: SUCCESS - isDraft=$($resultB.IsDraft), From=$($resultB.From)" -Level Info
        }
        else {
            Write-Log "  Method B: FAILED - $($resultB.Error)" -Level Warning
        }
    }

    # Return first successful result
    if ($results["MethodA"]?.Success) {
        return @{ Success = $true; Method = "A"; MessageId = $results["MethodA"].MessageId; IsDraft = $results["MethodA"].IsDraft }
    }
    if ($results["MethodB"]?.Success) {
        return @{ Success = $true; Method = "B"; MessageId = $results["MethodB"].MessageId; IsDraft = $results["MethodB"].IsDraft }
    }

    return @{ Success = $false; Results = $results }
}

# ================================
# IMAP Functions
# ================================

function Test-FolderExcluded {
    param([string]$FolderName)
    foreach ($exclude in $ExcludeFolders) {
        if ($FolderName -ieq $exclude -or $FolderName -ilike "*$exclude*") { return $true }
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

    Write-Log "Starting: $sourceEmail -> $targetEmail" -Level Info -User $sourceEmail

    $stats = @{ Migrated = 0; Failed = 0; MethodA_Success = 0; MethodB_Success = 0; MethodA_Draft = 0; MethodB_Draft = 0 }
    $client = $null

    try {
        # Connect IMAP
        $client = New-Object MailKit.Net.Imap.ImapClient
        $client.Timeout = 30000

        if ($ImapSkipCertValidation) {
            $client.ServerCertificateValidationCallback = { param($s,$c,$ch,$e) return $true }
        }

        $secureSocket = if ($ImapUseSsl) { [MailKit.Security.SecureSocketOptions]::SslOnConnect } 
                        else { [MailKit.Security.SecureSocketOptions]::StartTlsWhenAvailable }

        $client.Connect($ImapServer, $ImapPort, $secureSocket)
        $client.Authenticate($imapUsername, $imapPassword)
        Write-Log "IMAP connected" -Level Success -User $sourceEmail

        # Get folders
        $personalNamespace = $client.PersonalNamespaces[0]
        $folders = $client.GetFolders($personalNamespace)

        foreach ($folder in $folders) {
            $folderName = $folder.FullName

            # Skip problematic folders
            if ($folderName -match '^\{[0-9A-Fa-f-]{36}\}$') { continue }
            if ($folderName -in @('Conversation History', 'Sync Issues')) { continue }
            if (Test-FolderExcluded -FolderName $folderName) { continue }

            # Reconnect if needed
            if (!$client.IsConnected) {
                $client.Connect($ImapServer, $ImapPort, $secureSocket)
                $client.Authenticate($imapUsername, $imapPassword)
                $folder = $client.GetFolder($folderName)
            }

            try { $folder.Open([MailKit.FolderAccess]::ReadOnly) }
            catch { continue }

            if ($folder.Count -eq 0) { $folder.Close(); continue }

            Write-Log "Processing: $folderName ($($folder.Count) messages)" -Level Info -User $sourceEmail

            $uids = $folder.Search([MailKit.Search.SearchQuery]::All)
            
            if ($MaxMessagesPerMailbox -and $uids.Count -gt $MaxMessagesPerMailbox) {
                $uids = $uids | Select-Object -First $MaxMessagesPerMailbox
            }

            $targetFolderId = Get-OrCreateMailFolder -TargetMailbox $targetEmail -FolderPath $folderName -FolderCache $FolderCache

            $msgIndex = 0
            foreach ($uid in $uids) {
                $msgIndex++

                try {
                    $message = $folder.GetMessage($uid)
                    
                    $memStream = New-Object System.IO.MemoryStream
                    $message.WriteTo($memStream)
                    $mimeBytes = $memStream.ToArray()
                    $memStream.Dispose()

                    $subject = if ($message.Subject) { 
                        if ($message.Subject.Length -gt 40) { $message.Subject.Substring(0,37) + "..." } 
                        else { $message.Subject }
                    } else { "(No Subject)" }

                    if ($WhatIf) {
                        Write-Log "[WHATIF] $subject" -Level Info -User $sourceEmail
                        continue
                    }

                    # Get flags
                    $isRead = $false
                    try {
                        $summary = $folder.Fetch(@($uid), [MailKit.MessageSummaryItems]::Flags)
                        if ($summary -and $summary.Count -gt 0) {
                            $isRead = ($summary[0].Flags -band [MailKit.MessageFlags]::Seen) -eq [MailKit.MessageFlags]::Seen
                        }
                    }
                    catch { }

                    $receivedDate = if ($message.Date.DateTime -ne [DateTime]::MinValue) { $message.Date.DateTime } else { Get-Date }

                    Write-Log "[$msgIndex/$($uids.Count)] $subject" -Level Info -User $sourceEmail

                    # Import with selected method(s)
                    $result = Import-MessageToGraph `
                        -TargetMailbox $targetEmail `
                        -FolderId $targetFolderId `
                        -MimeContent $mimeBytes `
                        -ReceivedDate $receivedDate `
                        -IsRead $isRead `
                        -Method $ImportMethod

                    if ($result.Success) {
                        $stats.Migrated++
                        
                        if ($result.Method -eq "A") {
                            $stats.MethodA_Success++
                            if ($result.IsDraft) { $stats.MethodA_Draft++ }
                        }
                        elseif ($result.Method -eq "B") {
                            $stats.MethodB_Success++
                            if ($result.IsDraft) { $stats.MethodB_Draft++ }
                        }
                        
                        $draftStatus = if ($result.IsDraft) { " [DRAFT!]" } else { " [OK]" }
                        Write-Log "  -> Method $($result.Method)$draftStatus" -Level Success -User $sourceEmail
                    }
                    else {
                        $stats.Failed++
                        Write-Log "  -> FAILED both methods" -Level Error -User $sourceEmail
                    }

                    if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
                }
                catch {
                    $stats.Failed++
                    Write-Log "Failed UID $uid : $_" -Level Error -User $sourceEmail
                    if (!$ContinueOnError) { throw }
                }
            }

            $folder.Close()
        }

        Write-Log "Complete. Migrated: $($stats.Migrated), Failed: $($stats.Failed)" -Level Success -User $sourceEmail
        Write-Log "Method A: $($stats.MethodA_Success) success, $($stats.MethodA_Draft) drafts" -Level Info -User $sourceEmail
        Write-Log "Method B: $($stats.MethodB_Success) success, $($stats.MethodB_Draft) drafts" -Level Info -User $sourceEmail
    }
    finally {
        if ($client -and $client.IsConnected) {
            try { $client.Disconnect($true) } catch { }
            $client.Dispose()
        }
    }

    return $stats
}

# ================================
# Main
# ================================

try {
    Initialize-Logging

    Write-Log "=== Graph MIME Import Test ===" -Level Info
    Write-Log "Import Method: $ImportMethod" -Level Info
    Write-Log "IMAP: ${ImapServer}:${ImapPort}" -Level Info

    if (!(Initialize-MailKit)) { throw "Failed to load MailKit" }

    # Validate
    if ($TestMode -or $TestSource) {
        if (!$TestSource -or !$TestTarget -or !$TestPassword) { throw "Test mode requires TestSource, TestTarget, TestPassword" }
        $TestMode = $true
        Write-Log "*** TEST MODE ***" -Level Warning
    }
    elseif (!$UserCsvPath) {
        throw "Provide -UserCsvPath or test parameters"
    }

    # Test token
    Write-Log "Testing OAuth..." -Level Info
    $null = Get-GraphToken
    Write-Log "Token OK" -Level Success

    # Users
    $users = if ($TestMode) {
        @([PSCustomObject]@{
            Email = $TestSource
            Username = if ($TestUsername) { $TestUsername } else { $TestSource }
            Password = $TestPassword
            TargetEmail = $TestTarget
        })
    }
    else {
        $firstLine = Get-Content $UserCsvPath -First 1
        $delimiter = if ($firstLine -match ';') { ';' } else { ',' }
        Import-Csv -Path $UserCsvPath -Delimiter $delimiter
    }

    $folderCache = @{}

    foreach ($user in $users) {
        $userHash = @{
            Email = $user.Email
            Username = if ($user.Username) { $user.Username } else { $user.Email }
            Password = $user.Password
            TargetEmail = if ($user.TargetEmail) { $user.TargetEmail } else { $user.Email }
        }

        Write-Log "`n========================================" -Level Info
        Write-Log "User: $($userHash.Email)" -Level Info
        Write-Log "========================================" -Level Info

        try {
            $stats = Migrate-UserMailbox -User $userHash -FolderCache $folderCache
        }
        catch {
            Write-Log "Failed: $_" -Level Error
            if (!$ContinueOnError) { throw }
        }
    }

    Write-Log "`n=== COMPARISON SUMMARY ===" -Level Info
    Write-Log "Check your mailbox to see which method created proper emails vs drafts" -Level Info
    Write-Log "Method A = text/plain with MIME string" -Level Info
    Write-Log "Method B = octet-stream / message/rfc822" -Level Info
}
catch {
    Write-Log "FATAL: $_" -Level Error
    throw
}
