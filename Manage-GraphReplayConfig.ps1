<#
Manage-GraphReplayConfig.ps1 - Configuration Management for Graph Replay
Creates, updates, and tests configuration files for the Graph Email Replay script
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet("Create", "Update", "Test", "List", "Show", "Encrypt")]
    [string]$Action,
    
    [string]$ConfigPath,
    [string]$ConfigName,
    [string]$ConfigDirectory = "C:\GraphReplay\Configs"
)

# Ensure config directory exists
if (!(Test-Path $ConfigDirectory)) {
    New-Item -ItemType Directory -Path $ConfigDirectory -Force | Out-Null
}

function New-ConfigFile {
    param([string]$Path)
    
    Write-Host "Creating new configuration file..." -ForegroundColor Cyan
    Write-Host "Please provide the following information:" -ForegroundColor Yellow
    
    $config = @{}
    
    # Required fields
    $config.TenantId = Read-Host "Tenant ID (required)"
    $config.ClientId = Read-Host "Client ID (required)"
    
    # Handle secret securely
    $secretResponse = Read-Host "Client Secret (required)" -AsSecureString
    $config.ClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secretResponse)
    )
    
    # Source mailboxes
    $mailboxInput = Read-Host "Source Mailboxes (comma-separated)"
    $config.SourceMailboxes = $mailboxInput -split ',' | ForEach-Object { $_.Trim() }
    
    $config.TargetMailbox = Read-Host "Target Mailbox (required)"
    
    # Optional fields
    Write-Host "`nOptional settings (press Enter to skip):" -ForegroundColor Yellow
    
    $folder = Read-Host "Folder Name [Inbox]"
    if ($folder) { $config.FolderName = $folder }
    
    $mode = Read-Host "Replay Mode (Transparent/Wrapper) [Transparent]"
    if ($mode) { $config.ReplayMode = $mode }
    
    $attachOnly = Read-Host "Attachments Only (true/false) [false]"
    if ($attachOnly -eq 'true') { $config.AttachmentsOnly = $true }
    
    $skipProcessed = Read-Host "Skip Already Processed (true/false) [false]"
    if ($skipProcessed -eq 'true') { $config.SkipAlreadyProcessed = $true }
    
    $bcc = Read-Host "BCC Always (comma-separated)"
    if ($bcc) { 
        $config.BccAlways = $bcc -split ',' | ForEach-Object { $_.Trim() }
    }
    
    $logPath = Read-Host "Log Path"
    if ($logPath) { $config.LogPath = $logPath }
    
    $maxMsg = Read-Host "Max Messages"
    if ($maxMsg) { $config.MaxMessages = [int]$maxMsg }
    
    $batchSize = Read-Host "Batch Size [50]"
    if ($batchSize) { $config.BatchSize = [int]$batchSize }
    
    $throttle = Read-Host "Throttle MS [100]"
    if ($throttle) { $config.ThrottleMs = [int]$throttle }
    
    $header = Read-Host "Processed Header [X-GraphReplay-Processed]"
    if ($header) { $config.ProcessedHeader = $header }
    
    # Add metadata
    $config.CreatedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $config.CreatedBy = $env:USERNAME
    $config.Description = Read-Host "Description/Notes"
    
    # Encrypt the secret
    try {
        $secureString = ConvertTo-SecureString $config.ClientSecret -AsPlainText -Force
        $config.ClientSecretEncrypted = ConvertFrom-SecureString $secureString
        $config.Remove('ClientSecret')
        Write-Host "Client Secret encrypted successfully" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not encrypt Client Secret - will be stored in plain text"
    }
    
    # Save configuration
    $config | ConvertTo-Json -Depth 10 | Set-Content -Path $Path
    Write-Host "`nConfiguration saved to: $Path" -ForegroundColor Green
    
    # Offer to test
    $testNow = Read-Host "`nWould you like to test this configuration now? (y/n)"
    if ($testNow -eq 'y') {
        Test-ConfigFile -Path $Path
    }
}

function Update-ConfigFile {
    param([string]$Path)
    
    if (!(Test-Path $Path)) {
        Write-Error "Configuration file not found: $Path"
        return
    }
    
    $config = Get-Content $Path -Raw | ConvertFrom-Json | ConvertTo-HashTable
    
    Write-Host "Current configuration:" -ForegroundColor Cyan
    Show-ConfigFile -Path $Path
    
    Write-Host "`nEnter new values (press Enter to keep current):" -ForegroundColor Yellow
    
    # Update each field
    $newTenant = Read-Host "Tenant ID [$($config.TenantId)]"
    if ($newTenant) { $config.TenantId = $newTenant }
    
    $newClient = Read-Host "Client ID [$($config.ClientId)]"
    if ($newClient) { $config.ClientId = $newClient }
    
    $updateSecret = Read-Host "Update Client Secret? (y/n)"
    if ($updateSecret -eq 'y') {
        $secretResponse = Read-Host "Client Secret" -AsSecureString
        $newSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secretResponse)
        )
        
        # Encrypt the new secret
        $secureString = ConvertTo-SecureString $newSecret -AsPlainText -Force
        $config.ClientSecretEncrypted = ConvertFrom-SecureString $secureString
        $config.Remove('ClientSecret')
    }
    
    $newTarget = Read-Host "Target Mailbox [$($config.TargetMailbox)]"
    if ($newTarget) { $config.TargetMailbox = $newTarget }
    
    # Update metadata
    $config.ModifiedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $config.ModifiedBy = $env:USERNAME
    
    # Save updated configuration
    $config | ConvertTo-Json -Depth 10 | Set-Content -Path $Path
    Write-Host "`nConfiguration updated successfully" -ForegroundColor Green
}

function Test-ConfigFile {
    param([string]$Path)
    
    if (!(Test-Path $Path)) {
        Write-Error "Configuration file not found: $Path"
        return
    }
    
    Write-Host "Testing configuration: $Path" -ForegroundColor Cyan
    
    try {
        # Load the config
        $configData = Get-Content $Path -Raw | ConvertFrom-Json
        
        # Decrypt secret if needed
        $secret = $null
        if ($configData.ClientSecretEncrypted) {
            try {
                $secureString = ConvertTo-SecureString $configData.ClientSecretEncrypted
                $secret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
                )
            }
            catch {
                Write-Error "Could not decrypt Client Secret"
                return
            }
        }
        else {
            $secret = $configData.ClientSecret
        }
        
        # Test authentication
        Write-Host "Testing authentication..." -ForegroundColor Yellow
        $body = @{
            client_id     = $configData.ClientId
            client_secret = $secret
            scope        = "https://graph.microsoft.com/.default"
            grant_type   = "client_credentials"
        }
        
        $tokenUrl = "https://login.microsoftonline.com/$($configData.TenantId)/oauth2/v2.0/token"
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body
        
        if ($response.access_token) {
            Write-Host "✓ Authentication successful" -ForegroundColor Green
            
            # Test mailbox access
            Write-Host "Testing mailbox access..." -ForegroundColor Yellow
            $headers = @{
                "Authorization" = "Bearer $($response.access_token)"
            }
            
            foreach ($mailbox in $configData.SourceMailboxes) {
                try {
                    $uri = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$top=1"
                    $testResponse = Invoke-RestMethod -Uri $uri -Headers $headers
                    Write-Host "  ✓ $mailbox - Accessible" -ForegroundColor Green
                }
                catch {
                    Write-Host "  ✗ $mailbox - Not accessible: $_" -ForegroundColor Red
                }
            }
            
            # Test target mailbox
            try {
                $uri = "https://graph.microsoft.com/v1.0/users/$($configData.TargetMailbox)"
                $testResponse = Invoke-RestMethod -Uri $uri -Headers $headers
                Write-Host "✓ Target mailbox accessible: $($configData.TargetMailbox)" -ForegroundColor Green
            }
            catch {
                Write-Host "✗ Target mailbox not accessible: $_" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Error "Configuration test failed: $_"
    }
}

function Show-ConfigFile {
    param([string]$Path)
    
    if (!(Test-Path $Path)) {
        Write-Error "Configuration file not found: $Path"
        return
    }
    
    $config = Get-Content $Path -Raw | ConvertFrom-Json
    
    Write-Host "`n=== Configuration Details ===" -ForegroundColor Cyan
    Write-Host "File: $Path" -ForegroundColor Gray
    
    $config.PSObject.Properties | ForEach-Object {
        if ($_.Name -eq 'ClientSecretEncrypted') {
            Write-Host "$($_.Name): [ENCRYPTED]" -ForegroundColor Yellow
        }
        elseif ($_.Name -eq 'ClientSecret') {
            Write-Host "$($_.Name): [HIDDEN]" -ForegroundColor Yellow
        }
        else {
            $value = if ($_.Value -is [array]) { $_.Value -join ', ' } else { $_.Value }
            Write-Host "$($_.Name): $value"
        }
    }
}

function ConvertTo-HashTable {
    param(
        [Parameter(ValueFromPipeline)]
        $InputObject
    )
    
    process {
        if ($null -eq $InputObject) { return $null }
        
        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $collection = @(
                foreach ($object in $InputObject) {
                    ConvertTo-HashTable $object
                }
            )
            return ,$collection
        }
        elseif ($InputObject -is [psobject]) {
            $hash = @{}
            foreach ($property in $InputObject.PSObject.Properties) {
                $hash[$property.Name] = ConvertTo-HashTable $property.Value
            }
            return $hash
        }
        else {
            return $InputObject
        }
    }
}

# Main execution
switch ($Action) {
    "Create" {
        if (!$ConfigPath) {
            if (!$ConfigName) {
                $ConfigName = Read-Host "Enter config name (e.g., Company1)"
            }
            $ConfigPath = Join-Path $ConfigDirectory "$ConfigName.json"
        }
        New-ConfigFile -Path $ConfigPath
    }
    
    "Update" {
        if (!$ConfigPath) {
            if (!$ConfigName) {
                $ConfigName = Read-Host "Enter config name"
            }
            $ConfigPath = Join-Path $ConfigDirectory "$ConfigName.json"
        }
        Update-ConfigFile -Path $ConfigPath
    }
    
    "Test" {
        if (!$ConfigPath) {
            if (!$ConfigName) {
                $ConfigName = Read-Host "Enter config name"
            }
            $ConfigPath = Join-Path $ConfigDirectory "$ConfigName.json"
        }
        Test-ConfigFile -Path $ConfigPath
    }
    
    "List" {
        Write-Host "`nAvailable configurations:" -ForegroundColor Cyan
        Get-ChildItem -Path $ConfigDirectory -Filter "*.json" | ForEach-Object {
            $config = Get-Content $_.FullName -Raw | ConvertFrom-Json
            Write-Host "`n$($_.BaseName)" -ForegroundColor Green
            Write-Host "  File: $($_.FullName)"
            Write-Host "  Tenant: $($config.TenantId)"
            Write-Host "  Target: $($config.TargetMailbox)"
            Write-Host "  Sources: $($config.SourceMailboxes -join ', ')"
            if ($config.Description) {
                Write-Host "  Description: $($config.Description)"
            }
            if ($config.CreatedDate) {
                Write-Host "  Created: $($config.CreatedDate)"
            }
        }
    }
    
    "Show" {
        if (!$ConfigPath) {
            if (!$ConfigName) {
                $ConfigName = Read-Host "Enter config name"
            }
            $ConfigPath = Join-Path $ConfigDirectory "$ConfigName.json"
        }
        Show-ConfigFile -Path $ConfigPath
    }
    
    "Encrypt" {
        # Re-encrypt all configs (useful after moving to new machine)
        Write-Host "Re-encrypting all configuration files..." -ForegroundColor Cyan
        Get-ChildItem -Path $ConfigDirectory -Filter "*.json" | ForEach-Object {
            try {
                $config = Get-Content $_.FullName -Raw | ConvertFrom-Json | ConvertTo-HashTable
                
                if ($config.ClientSecret -and !$config.ClientSecretEncrypted) {
                    $secureString = ConvertTo-SecureString $config.ClientSecret -AsPlainText -Force
                    $config.ClientSecretEncrypted = ConvertFrom-SecureString $secureString
                    $config.Remove('ClientSecret')
                    $config | ConvertTo-Json -Depth 10 | Set-Content -Path $_.FullName
                    Write-Host "✓ Encrypted: $($_.Name)" -ForegroundColor Green
                }
                else {
                    Write-Host "○ Already encrypted: $($_.Name)" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "✗ Failed: $($_.Name) - $_" -ForegroundColor Red
            }
        }
    }
}