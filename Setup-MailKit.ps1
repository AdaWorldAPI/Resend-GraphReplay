<#
.SYNOPSIS
    Downloads MailKit/MimeKit dependencies for the migration script
    
.DESCRIPTION
    Downloads required DLLs from NuGet:
    - BouncyCastle.Crypto.dll (cryptography)
    - MimeKit.dll (MIME parsing)
    - MailKit.dll (IMAP client)
    
.EXAMPLE
    .\Setup-MailKit.ps1
#>

$libPath = Join-Path $PSScriptRoot "lib"

Write-Host "`n=== MailKit Setup ===" -ForegroundColor Cyan
Write-Host "Downloading dependencies to: $libPath`n" -ForegroundColor Gray

# Create lib directory
if (!(Test-Path $libPath)) {
    New-Item -ItemType Directory -Path $libPath -Force | Out-Null
}

# Packages to download
$packages = @(
    @{ Name = "Portable.BouncyCastle"; Version = "1.9.0"; Dll = "BouncyCastle.Crypto.dll" },
    @{ Name = "MimeKit"; Version = "4.3.0"; Dll = "MimeKit.dll" },
    @{ Name = "MailKit"; Version = "4.3.0"; Dll = "MailKit.dll" }
)

$success = $true

foreach ($pkg in $packages) {
    Write-Host "  $($pkg.Name) v$($pkg.Version)... " -NoNewline
    
    $url = "https://www.nuget.org/api/v2/package/$($pkg.Name)/$($pkg.Version)"
    $tempZip = Join-Path $libPath "$($pkg.Name).nupkg.zip"
    $extractPath = Join-Path $libPath "$($pkg.Name)_temp"
    
    try {
        # Download
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $tempZip -UseBasicParsing -ErrorAction Stop
        
        # Extract
        Expand-Archive -Path $tempZip -DestinationPath $extractPath -Force
        
        # Find DLL in netstandard2.0
        $dllSource = Get-ChildItem -Path "$extractPath\lib\netstandard2.0\$($pkg.Dll)" -ErrorAction SilentlyContinue
        
        if (!$dllSource) {
            # Try alternative name
            $dllSource = Get-ChildItem -Path "$extractPath\lib\netstandard2.0\*.dll" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        
        if ($dllSource) {
            Copy-Item $dllSource.FullName (Join-Path $libPath $pkg.Dll) -Force
            Write-Host "OK" -ForegroundColor Green
        }
        else {
            Write-Host "DLL not found!" -ForegroundColor Red
            $success = $false
        }
    }
    catch {
        Write-Host "FAILED: $_" -ForegroundColor Red
        $success = $false
    }
    finally {
        # Cleanup
        Remove-Item $tempZip -Force -ErrorAction SilentlyContinue
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`n--- Installed DLLs ---" -ForegroundColor Cyan
Get-ChildItem "$libPath\*.dll" -ErrorAction SilentlyContinue | ForEach-Object {
    $size = [math]::Round($_.Length / 1KB, 0)
    Write-Host "  $($_.Name) ($size KB)" -ForegroundColor Green
}

if ($success) {
    Write-Host "`nSetup complete! Run the migration script now." -ForegroundColor Green
    Write-Host "  .\Kopano-IMAP-to-Graph-Migration-MailKit.ps1 -TestMode ..." -ForegroundColor Gray
}
else {
    Write-Host "`nSetup had errors. Check above." -ForegroundColor Red
}
