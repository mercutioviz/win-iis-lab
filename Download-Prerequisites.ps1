#Requires -RunAsAdministrator
#Requires -Version 7.0

<#
.SYNOPSIS
    IIS Lab Environment - Download & Extract Prerequisites
.DESCRIPTION
    Automates the manual prep steps:
    - Installs 7-Zip for fast extraction
    - Downloads all prerequisite files
    - Installs MSI/EXE packages (Node.js, URL Rewrite, ARR, VC++ Redist)
    - Extracts ZIP archives to their target locations using 7-Zip
    - Downloads the correct Juice Shop build (auto-detects node24 asset)
.NOTES
    Run AFTER installing .NET 4.8 and PowerShell 7 (requires reboot).
	Download .NET 4.8 installer at https://go.microsoft.com/fwlink/?linkid=2088631
	Download Powershell 7 installer at https://github.com/PowerShell/PowerShell/releases/latest
    Run this script in elevated PowerShell 7 (pwsh).
    After this script completes, run Install-Prerequisites.ps1.
#>

# ============================================================================
# CONFIGURATION — Adjust versions/URLs as needed
# ============================================================================

$Config = @{
    # Working directory
    LabSetupPath        = "C:\LabSetup"

    # 7-Zip
    SevenZipURL         = "https://www.7-zip.org/a/7z2600-x64.msi"
    SevenZipExe         = "C:\Program Files\7-Zip\7z.exe"

    # Node.js 24 LTS
    NodeURL             = "https://nodejs.org/dist/v24.14.0/node-v24.14.0-x64.msi"
    NodeMSI             = "C:\LabSetup\nodejs.msi"
    ExpectedNodeMajor   = 24
    ExpectedModuleVer   = 137

    # URL Rewrite Module 2.1
    RewriteURL          = "https://download.microsoft.com/download/1/2/8/128E2E22-C1B9-44A4-BE2A-5859ED1D4592/rewrite_amd64_en-US.msi"
    RewriteMSI          = "C:\LabSetup\rewrite_amd64.msi"

    # Application Request Routing 3.0
    ARRURL              = "https://download.microsoft.com/download/E/9/8/E9849D6A-020E-47E4-9FD0-A023E99B54EB/requestRouter_amd64.msi"
    ARRMSI              = "C:\LabSetup\requestRouter_amd64.msi"

    # Visual C++ Redistributable 2015-2022
    VCRedistURL         = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
    VCRedistEXE         = "C:\LabSetup\vc_redist_x64.exe"

    # PHP 8.5 NTS x64 (vs17)
    PHPURL              = "https://windows.php.net/downloads/releases/php-8.5.1-nts-Win32-vs17-x64.zip"
    PHPZip              = "C:\LabSetup\php.zip"
    PHPTarget           = "C:\PHP"

    # MariaDB 11.4 LTS
    MariaDBURL          = "https://archive.mariadb.org/mariadb-11.4.10/winx64-packages/mariadb-11.4.10-winx64.zip"
    MariaDBZip          = "C:\LabSetup\mariadb.zip"
    MariaDBTarget       = "C:\MariaDB"
    MariaDBInnerFolder  = "mariadb-11.4.10-winx64"   # Folder name inside the zip

    # WordPress
    WordPressURL        = "https://wordpress.org/latest.zip"
    WordPressZip        = "C:\LabSetup\wordpress.zip"
    WordPressTarget     = "C:\inetpub\wordpress"
    WordPressInnerFolder = "wordpress"   # Folder name inside the zip

    # OWASP Juice Shop (auto-detected from GitHub API)
    JuiceShopZip        = "C:\LabSetup\juice-shop.zip"
    JuiceShopTarget     = "C:\inetpub\juiceshop"

    # WinSW
    WinSWURL            = "https://github.com/winsw/winsw/releases/download/v2.12.0/WinSW-x64.exe"
    WinSWEXE            = "C:\LabSetup\WinSW-x64.exe"
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-Step {
    param([string]$Message)
    Write-Host "`n$("=" * 60)" -ForegroundColor Cyan
    Write-Host "  $Message" -ForegroundColor Cyan
    Write-Host "$("=" * 60)" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "  ✅ $Message" -ForegroundColor Green
}

function Write-Failure {
    param([string]$Message)
    Write-Host "  ❌ $Message" -ForegroundColor Red
}

function Write-Info {
    param([string]$Message)
    Write-Host "  ℹ️  $Message" -ForegroundColor Yellow
}

function Download-File {
    param(
        [string]$URL,
        [string]$Destination,
        [string]$Description
    )

    if (Test-Path $Destination) {
        Write-Info "$Description already downloaded — skipping"
        return $true
    }

    Write-Info "Downloading $Description..."
    Write-Host "    URL: $URL" -ForegroundColor Gray
    try {
        Invoke-WebRequest -Uri $URL -OutFile $Destination -TimeoutSec 300
        if (Test-Path $Destination) {
            $size = [math]::Round((Get-Item $Destination).Length / 1MB, 1)
            Write-Success "$Description downloaded (${size} MB)"
            return $true
        } else {
            Write-Failure "$Description download failed — file not created"
            return $false
        }
    } catch {
        Write-Failure "$Description download failed: $($_.Exception.Message)"
        return $false
    }
}

function Extract-WithSevenZip {
    param(
        [string]$Archive,
        [string]$Destination,
        [string]$Description,
        [string]$InnerFolder = "",      # If the zip has a nested folder to flatten
        [string]$VerifyFile = ""        # A file that should exist after extraction
    )

    $sz = $Config.SevenZipExe

    # Check if already extracted
    if ($VerifyFile -and (Test-Path $VerifyFile)) {
        Write-Info "$Description already extracted — skipping"
        return $true
    }

    if (-not (Test-Path $Archive)) {
        Write-Failure "$Description archive not found: $Archive"
        return $false
    }

    Write-Info "Extracting $Description..."

    if ($InnerFolder) {
        # Extract to a temp location, then move contents from the inner folder
        $tempDir = "$($Config.LabSetupPath)\temp-extract-$([guid]::NewGuid().ToString().Substring(0,8))"
        New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

        $proc = Start-Process $sz -ArgumentList "x `"$Archive`" -o`"$tempDir`" -y" -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -ne 0) {
            Write-Failure "$Description extraction failed (exit code: $($proc.ExitCode))"
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            return $false
        }

        # Create destination if needed
        New-Item -Path $Destination -ItemType Directory -Force | Out-Null

        # Find the inner folder (may have slight name variations)
        $innerPath = Get-ChildItem $tempDir -Directory | Where-Object { $_.Name -like "$InnerFolder*" } | Select-Object -First 1

        if ($innerPath) {
            # Move contents from inner folder to destination
            Get-ChildItem $innerPath.FullName | Move-Item -Destination $Destination -Force
            Write-Success "$Description extracted and flattened to $Destination"
        } else {
            # No inner folder found — move everything
            Get-ChildItem $tempDir | Move-Item -Destination $Destination -Force
            Write-Success "$Description extracted to $Destination"
        }

        # Clean up temp
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        # Extract directly to destination
        New-Item -Path $Destination -ItemType Directory -Force | Out-Null

        $proc = Start-Process $sz -ArgumentList "x `"$Archive`" -o`"$Destination`" -y" -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -ne 0) {
            Write-Failure "$Description extraction failed (exit code: $($proc.ExitCode))"
            return $false
        }

        Write-Success "$Description extracted to $Destination"
    }

    # Verify
    if ($VerifyFile) {
        if (Test-Path $VerifyFile) {
            Write-Success "Verified: $VerifyFile exists"
            return $true
        } else {
            Write-Failure "Verification failed: $VerifyFile not found"
            return $false
        }
    }

    return $true
}

# ============================================================================
# PRE-FLIGHT CHECKS
# ============================================================================

Write-Step "PRE-FLIGHT CHECKS"

$preflight = @(
    @{ Name = "PowerShell 7+";    Test = { $PSVersionTable.PSVersion.Major -ge 7 } },
    @{ Name = ".NET 4.8";         Test = { (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 528040 } },
    @{ Name = "Internet access";  Test = { (Test-NetConnection -ComputerName "www.7-zip.org" -Port 443 -WarningAction SilentlyContinue).TcpTestSucceeded } }
)

$allPassed = $true
foreach ($check in $preflight) {
    try {
        if (& $check.Test) {
            Write-Success $check.Name
        } else {
            Write-Failure $check.Name
            $allPassed = $false
        }
    } catch {
        Write-Failure "$($check.Name) — $($_.Exception.Message)"
        $allPassed = $false
    }
}

if (-not $allPassed) {
    Write-Failure "PRE-FLIGHT CHECKS FAILED."
    exit 1
}

Write-Success "All pre-flight checks passed!`n"

# ============================================================================
# STEP 1: CREATE DIRECTORY STRUCTURE
# ============================================================================

Write-Step "STEP 1: Creating Directory Structure"

$dirs = @(
    $Config.LabSetupPath,
    $Config.PHPTarget,
    $Config.MariaDBTarget,
    "C:\inetpub\juiceshop",
    "C:\inetpub\wordpress",
    "C:\inetpub\echospa",
    "C:\Tools"
)

foreach ($dir in $dirs) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Success "Created: $dir"
    } else {
        Write-Info "Already exists: $dir"
    }
}

# ============================================================================
# STEP 2: INSTALL 7-ZIP
# ============================================================================

Write-Step "STEP 2: Installing 7-Zip"

if (Test-Path $Config.SevenZipExe) {
    Write-Info "7-Zip already installed — skipping"
} else {
    $szMSI = "$($Config.LabSetupPath)\7zip.msi"

    if (-not (Download-File -URL $Config.SevenZipURL -Destination $szMSI -Description "7-Zip")) {
        Write-Failure "Cannot continue without 7-Zip"
        exit 1
    }

    Write-Info "Installing 7-Zip..."
    $proc = Start-Process msiexec.exe -ArgumentList "/i `"$szMSI`" /qn /norestart" -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -eq 0 -and (Test-Path $Config.SevenZipExe)) {
        Write-Success "7-Zip installed to C:\Program Files\7-Zip\"
    } else {
        Write-Failure "7-Zip installation failed (exit code: $($proc.ExitCode))"
        exit 1
    }
}

# ============================================================================
# STEP 3: DOWNLOAD ALL FILES
# ============================================================================

Write-Step "STEP 3: Downloading All Prerequisites"

$downloads = @(
    @{ URL = $Config.NodeURL;      Dest = $Config.NodeMSI;     Desc = "Node.js 24 LTS" },
    @{ URL = $Config.RewriteURL;   Dest = $Config.RewriteMSI;  Desc = "URL Rewrite Module 2.1" },
    @{ URL = $Config.ARRURL;       Dest = $Config.ARRMSI;      Desc = "Application Request Routing 3.0" },
    @{ URL = $Config.VCRedistURL;  Dest = $Config.VCRedistEXE; Desc = "Visual C++ Redistributable" },
    @{ URL = $Config.PHPURL;       Dest = $Config.PHPZip;      Desc = "PHP 8.5 NTS x64" },
    @{ URL = $Config.MariaDBURL;   Dest = $Config.MariaDBZip;  Desc = "MariaDB 11.4 LTS" },
    @{ URL = $Config.WordPressURL; Dest = $Config.WordPressZip; Desc = "WordPress" },
    @{ URL = $Config.WinSWURL;     Dest = $Config.WinSWEXE;   Desc = "WinSW" }
)

$downloadFailed = $false
foreach ($dl in $downloads) {
    if (-not (Download-File -URL $dl.URL -Destination $dl.Dest -Description $dl.Desc)) {
        $downloadFailed = $true
    }
}

# Juice Shop — dynamically find the correct asset from GitHub
Write-Info "Querying GitHub for latest Juice Shop release..."
try {
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/juice-shop/juice-shop/releases/latest" -TimeoutSec 30
    $jsVersion = $release.tag_name
    Write-Info "Latest Juice Shop version: $jsVersion"

    $jsAsset = $release.assets | Where-Object { $_.name -like "*node24_win32_x64.zip" } | Select-Object -First 1

    if (-not $jsAsset) {
        # Fallback: try node22 if node24 not available
        $jsAsset = $release.assets | Where-Object { $_.name -like "*node22_win32_x64.zip" } | Select-Object -First 1
        if ($jsAsset) {
            Write-Info "node24 build not available — using node22 build"
            Write-Info "⚠️  You may need to install Node.js 22 instead of 24!"
        }
    }

    if ($jsAsset) {
        if (-not (Download-File -URL $jsAsset.browser_download_url -Destination $Config.JuiceShopZip -Description "Juice Shop ($($jsAsset.name))")) {
            $downloadFailed = $true
        }
    } else {
        Write-Failure "No compatible Juice Shop Windows build found!"
        Write-Info "Available assets:"
        $release.assets | Where-Object { $_.name -like "*win32*" } | ForEach-Object { Write-Host "    $($_.name)" -ForegroundColor Gray }
        $downloadFailed = $true
    }
} catch {
    Write-Failure "Failed to query GitHub API: $($_.Exception.Message)"
    $downloadFailed = $true
}

if ($downloadFailed) {
    Write-Failure "Some downloads failed — check errors above"
    Write-Info "You can fix the failing URLs in the `$Config block and re-run"
    Write-Info "The script is idempotent — successful downloads will be skipped"
    exit 1
}

Write-Success "All downloads complete!"

# ============================================================================
# STEP 4: INSTALL MSI/EXE PACKAGES
# ============================================================================

Write-Step "STEP 4: Installing MSI/EXE Packages"

# --- Node.js ---
$existingNode = Get-Command node -ErrorAction SilentlyContinue
if ($existingNode) {
    $nodeVer = node --version 2>$null
    if ($nodeVer -match "^v$($Config.ExpectedNodeMajor)\.") {
        Write-Info "Node.js $nodeVer already installed — skipping"
    } else {
        Write-Info "Node.js $nodeVer found but expected v$($Config.ExpectedNodeMajor).x — upgrading..."
        $proc = Start-Process msiexec.exe -ArgumentList "/i `"$($Config.NodeMSI)`" /qn /norestart" -Wait -PassThru -NoNewWindow
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        Write-Success "Node.js upgraded (exit code: $($proc.ExitCode))"
    }
} else {
    Write-Info "Installing Node.js..."
    $proc = Start-Process msiexec.exe -ArgumentList "/i `"$($Config.NodeMSI)`" /qn /norestart" -Wait -PassThru -NoNewWindow

    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

    # Ensure Node.js is in Machine PATH (MSI sometimes adds to User PATH only)
    $nodePath = "C:\Program Files\nodejs"
    if (Test-Path "$nodePath\node.exe") {
        $machinePath = [System.Environment]::GetEnvironmentVariable("Path","Machine")
        if ($machinePath -notlike "*$nodePath*") {
            [System.Environment]::SetEnvironmentVariable("Path", "$machinePath;$nodePath", "Machine")
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
            Write-Info "Node.js added to Machine PATH"
        }
    }

    if ($proc.ExitCode -eq 0) {
        Write-Success "Node.js installed"
    } else {
        Write-Failure "Node.js install failed (exit code: $($proc.ExitCode))"
        exit 1
    }
}

# Verify Node
$nodeVer = node --version 2>$null
$modulesVer = (node -e "console.log(process.versions.modules)" 2>$null).Trim()
Write-Success "Node.js $nodeVer (MODULE_VERSION: $modulesVer)"

if ($modulesVer -ne "$($Config.ExpectedModuleVer)") {
    Write-Failure "Expected MODULE_VERSION $($Config.ExpectedModuleVer) but got $modulesVer"
    Write-Info "Juice Shop may not work — check version compatibility"
}

# --- VC++ Redistributable ---
# Check if already installed via vcruntime140.dll
if (Test-Path "$env:SystemRoot\System32\vcruntime140.dll") {
    Write-Info "VC++ Redistributable already present — installing latest anyway..."
}

$proc = Start-Process $Config.VCRedistEXE -ArgumentList "/install /quiet /norestart" -Wait -PassThru -NoNewWindow
if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 1638) {
    Write-Success "VC++ Redistributable installed (exit code: $($proc.ExitCode))"
} else {
    Write-Failure "VC++ Redistributable failed (exit code: $($proc.ExitCode))"
}

# Note: URL Rewrite and ARR are NOT installed here — they require IIS to be installed first.
# They will be installed by Install-Prerequisites.ps1 after IIS features are enabled.
Write-Info "URL Rewrite and ARR MSIs downloaded — will be installed by Install-Prerequisites.ps1 (after IIS)"

# ============================================================================
# STEP 5: EXTRACT ZIP ARCHIVES
# ============================================================================

Write-Step "STEP 5: Extracting ZIP Archives with 7-Zip"

# --- PHP ---
Extract-WithSevenZip `
    -Archive $Config.PHPZip `
    -Destination $Config.PHPTarget `
    -Description "PHP 8.5 NTS x64" `
    -InnerFolder "" `
    -VerifyFile "$($Config.PHPTarget)\php-cgi.exe"

# --- MariaDB ---
Extract-WithSevenZip `
    -Archive $Config.MariaDBZip `
    -Destination $Config.MariaDBTarget `
    -Description "MariaDB 11.4 LTS" `
    -InnerFolder $Config.MariaDBInnerFolder `
    -VerifyFile "$($Config.MariaDBTarget)\bin\mysqld.exe"

# --- Juice Shop ---
# Juice Shop zips contain a nested folder like "juice-shop_19.1.1"
Extract-WithSevenZip `
    -Archive $Config.JuiceShopZip `
    -Destination $Config.JuiceShopTarget `
    -Description "OWASP Juice Shop" `
    -InnerFolder "juice-shop" `
    -VerifyFile "$($Config.JuiceShopTarget)\package.json"

# --- WordPress ---
Extract-WithSevenZip `
    -Archive $Config.WordPressZip `
    -Destination $Config.WordPressTarget `
    -Description "WordPress" `
    -InnerFolder $Config.WordPressInnerFolder `
    -VerifyFile "$($Config.WordPressTarget)\wp-login.php"

# ============================================================================
# FINAL VERIFICATION
# ============================================================================

Write-Step "FINAL VERIFICATION"

$checks = @(
    @{ Name = "7-Zip installed";                 Test = { Test-Path $Config.SevenZipExe } },
    @{ Name = "Node.js $($Config.ExpectedNodeMajor)"; Test = { (node --version 2>$null) -match "^v$($Config.ExpectedNodeMajor)\." } },
    @{ Name = "Node MODULE_VERSION $($Config.ExpectedModuleVer)"; Test = { (node -e "console.log(process.versions.modules)" 2>$null).Trim() -eq "$($Config.ExpectedModuleVer)" } },
    @{ Name = "VC++ Redistributable";            Test = { Test-Path "$env:SystemRoot\System32\vcruntime140.dll" } },

    @{ Name = "--- MSIs Ready for Install-Prerequisites.ps1 ---"; Test = { $true } },
    @{ Name = "URL Rewrite MSI";                 Test = { Test-Path $Config.RewriteMSI } },
    @{ Name = "ARR MSI";                         Test = { Test-Path $Config.ARRMSI } },
    @{ Name = "WinSW executable";                Test = { Test-Path $Config.WinSWEXE } },

    @{ Name = "--- Extracted Archives ---";      Test = { $true } },
    @{ Name = "PHP (php-cgi.exe)";               Test = { Test-Path "$($Config.PHPTarget)\php-cgi.exe" } },
    @{ Name = "MariaDB (mysqld.exe)";            Test = { Test-Path "$($Config.MariaDBTarget)\bin\mysqld.exe" } },
    @{ Name = "Juice Shop (package.json)";       Test = { Test-Path "$($Config.JuiceShopTarget)\package.json" } },
    @{ Name = "WordPress (wp-login.php)";        Test = { Test-Path "$($Config.WordPressTarget)\wp-login.php" } }
)

$passCount = 0
$failCount = 0

foreach ($check in $checks) {
    if ($check.Name -match "^---") {
        Write-Host "`n  $($check.Name)" -ForegroundColor Yellow
        continue
    }
    try {
        if (& $check.Test) {
            Write-Success $check.Name
            $passCount++
        } else {
            Write-Failure $check.Name
            $failCount++
        }
    } catch {
        Write-Failure "$($check.Name) — $($_.Exception.Message)"
        $failCount++
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n$("=" * 60)" -ForegroundColor Cyan
Write-Host "  DOWNLOAD & EXTRACT COMPLETE" -ForegroundColor Cyan
Write-Host "$("=" * 60)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Results: $passCount passed, $failCount failed" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Red" })
Write-Host ""
Write-Host "  Files downloaded to:   $($Config.LabSetupPath)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Installed:" -ForegroundColor White
Write-Host "    • 7-Zip → C:\Program Files\7-Zip\" -ForegroundColor Gray
Write-Host "    • Node.js $(node --version 2>$null) → C:\Program Files\nodejs\" -ForegroundColor Gray
Write-Host "    • VC++ Redistributable → System" -ForegroundColor Gray
Write-Host ""
Write-Host "  Extracted:" -ForegroundColor White
Write-Host "    • PHP → $($Config.PHPTarget)\" -ForegroundColor Gray
Write-Host "    • MariaDB → $($Config.MariaDBTarget)\" -ForegroundColor Gray
Write-Host "    • Juice Shop → $($Config.JuiceShopTarget)\" -ForegroundColor Gray
Write-Host "    • WordPress → $($Config.WordPressTarget)\" -ForegroundColor Gray
Write-Host ""
Write-Host "  Ready for Install-Prerequisites.ps1:" -ForegroundColor White
Write-Host "    • URL Rewrite MSI (will install after IIS is enabled)" -ForegroundColor Gray
Write-Host "    • ARR MSI (will install after IIS is enabled)" -ForegroundColor Gray
Write-Host "    • WinSW (will be used by Install-JuiceShop.ps1)" -ForegroundColor Gray
Write-Host ""

if ($failCount -eq 0) {
    Write-Host "  ✅ Ready to run Install-Prerequisites.ps1" -ForegroundColor Green
} else {
    Write-Host "  ⚠️  Fix failed items and re-run (script is idempotent)." -ForegroundColor Red
}

Write-Host ""
Write-Host "$("=" * 60)" -ForegroundColor Cyan
