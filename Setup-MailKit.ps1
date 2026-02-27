<#
.SYNOPSIS
    Downloads MailKit/MimeKit DLLs required for the Kopano IMAP to Graph migration script.

.DESCRIPTION
    This script downloads the MailKit, MimeKit, and BouncyCastle .NET libraries from NuGet
    and places them in the lib/ folder alongside the migration script.

    Run this script once before using Kopano-IMAP-to-Graph-Migration.ps1.

.EXAMPLE
    .\Setup-MailKit.ps1
#>

$libPath = "$PSScriptRoot\lib"
New-Item -ItemType Directory -Path $libPath -Force | Out-Null

Write-Host "=== MailKit/MimeKit Setup ===" -ForegroundColor Cyan
Write-Host "Downloading dependencies to: $libPath" -ForegroundColor Cyan
Write-Host ""

# NuGet packages to download (order matters for dependencies)
$packages = @(
    @{ Name = "Portable.BouncyCastle"; Version = "1.9.0"; Framework = "netstandard2.0" },
    @{ Name = "MimeKit"; Version = "4.3.0"; Framework = "netstandard2.0" },
    @{ Name = "MailKit"; Version = "4.3.0"; Framework = "netstandard2.0" }
)

$allSuccess = $true

foreach ($pkg in $packages) {
    Write-Host "  Downloading $($pkg.Name) v$($pkg.Version)..." -NoNewline

    $url = "https://www.nuget.org/api/v2/package/$($pkg.Name)/$($pkg.Version)"
    $tempPath = "$libPath\$($pkg.Name).nupkg.zip"
    $extractPath = "$libPath\$($pkg.Name)_temp"

    try {
        # Download NuGet package (nupkg is a ZIP)
        Invoke-WebRequest -Uri $url -OutFile $tempPath -UseBasicParsing

        # Extract
        Expand-Archive -Path $tempPath -DestinationPath $extractPath -Force

        # Find the DLL for the target framework
        $dllSource = Get-ChildItem -Path "$extractPath\lib\$($pkg.Framework)\*.dll" -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if ($dllSource) {
            Copy-Item $dllSource.FullName $libPath -Force
            Write-Host " OK" -ForegroundColor Green
        }
        else {
            # Try net6.0 as fallback
            $dllSource = Get-ChildItem -Path "$extractPath\lib\net6.0\*.dll" -ErrorAction SilentlyContinue |
                Select-Object -First 1

            if ($dllSource) {
                Copy-Item $dllSource.FullName $libPath -Force
                Write-Host " OK (net6.0)" -ForegroundColor Green
            }
            else {
                Write-Host " DLL not found in package!" -ForegroundColor Red
                $allSuccess = $false
            }
        }

        # Cleanup temp files
        Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host " FAILED: $_" -ForegroundColor Red
        $allSuccess = $false

        # Cleanup on failure
        Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
        Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "Installed DLLs:" -ForegroundColor Cyan
$dlls = Get-ChildItem "$libPath\*.dll" -ErrorAction SilentlyContinue
if ($dlls) {
    foreach ($dll in $dlls) {
        $size = "{0:N0} KB" -f ($dll.Length / 1024)
        Write-Host "  $($dll.Name) ($size)" -ForegroundColor Green
    }
}
else {
    Write-Host "  No DLLs found!" -ForegroundColor Red
    $allSuccess = $false
}

Write-Host ""
if ($allSuccess) {
    Write-Host "Setup complete. You can now run the migration script." -ForegroundColor Green
}
else {
    Write-Host "Setup completed with errors. Check output above." -ForegroundColor Yellow
    Write-Host "You may need to download DLLs manually from https://www.nuget.org/" -ForegroundColor Yellow
}
