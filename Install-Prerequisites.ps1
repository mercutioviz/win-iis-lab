#Requires -RunAsAdministrator
#Requires -Version 7.0

<#
.SYNOPSIS
    IIS Lab Environment - Prerequisites Installation Script
.DESCRIPTION
    Installs and configures all prerequisites for the IIS lab environment:
    - IIS Role Features (including logging, tracing, FastCGI, ASP.NET, WebSockets)
    - URL Rewrite Module 2.1
    - Application Request Routing 3.0
    - PHP (configuration + IIS registration)
    - MariaDB (configuration, initialization, service)
    - Self-signed wildcard certificate
    - Hosts file entries
    - IIS logging with custom WAF/proxy header fields
.NOTES
    Run Download-Prerequisites.ps1 first to download/extract all required files.
    Run this script in elevated PowerShell 7 (pwsh).
#>

# ============================================================================
# CONFIGURATION — Adjust these values as needed
# ============================================================================

$Config = @{
    # Domain and hostnames
    LabDomain           = "lab.local"
    JuiceShopHost       = "juiceshop.lab.local"
    WordPressHost       = "wordpress.lab.local"
    EchoSPAHost         = "echo.lab.local"

    # Certificate
    CertValidityYears   = 5
    CertFriendlyName    = "Lab Local Wildcard"

    # MariaDB
    MariaDBPath         = "C:\MariaDB"
    MariaDBDataDir      = "C:\MariaDB\data"
    MariaDBPort         = 3306
    MariaDBServiceName  = "MariaDB"

    # PHP
    PHPPath             = "C:\PHP"

    # Paths
    LabSetupPath        = "C:\LabSetup"
    ToolsPath           = "C:\Tools"
    ThumbprintFile      = "C:\LabSetup\cert-thumbprint.txt"

    # Files expected from Download-Prerequisites.ps1
    RewriteMSI          = "C:\LabSetup\rewrite_amd64.msi"
    ARRMSI              = "C:\LabSetup\requestRouter_amd64.msi"
    WinSW               = "C:\LabSetup\WinSW-x64.exe"

    # Node.js
    ExpectedNodeMajor   = 24
    ExpectedModuleVer   = 137
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

$appcmd = "$env:SystemRoot\System32\inetsrv\appcmd.exe"

# ============================================================================
# ENVIRONMENT SETUP — Refresh PATH for elevated sessions
# ============================================================================

Write-Step "ENVIRONMENT SETUP"

# Refresh PATH to pick up any recent installations (Node.js, etc.)
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
Write-Success "PATH refreshed from system environment"

# Ensure Node.js is in Machine PATH (MSI sometimes adds to User PATH only)
$nodePath = "C:\Program Files\nodejs"
if (Test-Path "$nodePath\node.exe") {
    $machinePath = [System.Environment]::GetEnvironmentVariable("Path","Machine")
    if ($machinePath -notlike "*$nodePath*") {
        [System.Environment]::SetEnvironmentVariable("Path", "$machinePath;$nodePath", "Machine")
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        Write-Success "Node.js added to Machine PATH (was in User PATH only)"
    } else {
        Write-Info "Node.js already in Machine PATH"
    }
} else {
    Write-Failure "Node.js not found at $nodePath\node.exe"
    Write-Failure "Ensure Download-Prerequisites.ps1 was run first"
    exit 1
}

# ============================================================================
# PRE-FLIGHT CHECKS
# ============================================================================

Write-Step "PRE-FLIGHT CHECKS"

$preflight = @(
    @{ Name = "PowerShell 7+";               Test = { $PSVersionTable.PSVersion.Major -ge 7 } },
    @{ Name = ".NET 4.8";                     Test = { (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 528040 } },
    @{ Name = "Node.js $($Config.ExpectedNodeMajor)"; Test = { (node --version 2>$null) -match "^v$($Config.ExpectedNodeMajor)\." } },
    @{ Name = "Node MODULE_VERSION $($Config.ExpectedModuleVer)"; Test = { (node -e "console.log(process.versions.modules)" 2>$null).Trim() -eq "$($Config.ExpectedModuleVer)" } },
    @{ Name = "URL Rewrite MSI";              Test = { Test-Path $Config.RewriteMSI } },
    @{ Name = "ARR MSI";                      Test = { Test-Path $Config.ARRMSI } },
    @{ Name = "WinSW executable";             Test = { Test-Path $Config.WinSW } },
    @{ Name = "PHP extracted";                Test = { Test-Path "$($Config.PHPPath)\php-cgi.exe" } },
    @{ Name = "MariaDB extracted";            Test = { Test-Path "$($Config.MariaDBPath)\bin\mysqld.exe" } },
    @{ Name = "Juice Shop extracted";         Test = { Test-Path "C:\inetpub\juiceshop\package.json" } },
    @{ Name = "WordPress extracted";          Test = { Test-Path "C:\inetpub\wordpress\wp-login.php" } }
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
    Write-Host "`n  ❌ PRE-FLIGHT CHECKS FAILED. Fix issues above and re-run." -ForegroundColor Red
    exit 1
}

Write-Success "All pre-flight checks passed!`n"

# ============================================================================
# STEP 1: INSTALL IIS ROLE FEATURES
# ============================================================================

Write-Step "STEP 1: Installing IIS Role Features"

$iisFeatures = @(
    "Web-Server",
    "Web-CGI",
    "Web-Asp-Net45",
    "Web-Net-Ext45",
    "Web-ISAPI-Ext",
    "Web-ISAPI-Filter",
    "Web-WebSockets",
    "Web-Mgmt-Console",
    "Web-Http-Tracing",
    "Web-Custom-Logging",
    "Web-Log-Libraries",
    "NET-Framework-45-ASPNET"
)

$result = Install-WindowsFeature -Name $iisFeatures -IncludeManagementTools

if ($result.Success) {
    Write-Success "IIS features installed successfully"
    if ($result.RestartNeeded -eq "Yes") {
        Write-Failure "REBOOT REQUIRED. Reboot and re-run this script."
        exit 1
    }
} else {
    Write-Failure "IIS feature installation failed"
    exit 1
}

# Verify key features
$installedFeatures = Get-WindowsFeature $iisFeatures | Where-Object { $_.InstallState -eq "Installed" }
Write-Info "$($installedFeatures.Count) of $($iisFeatures.Count) features confirmed installed"

# ============================================================================
# STEP 2: INSTALL URL REWRITE MODULE 2.1
# ============================================================================

Write-Step "STEP 2: Installing URL Rewrite Module 2.1"

if (Test-Path "$env:SystemRoot\System32\inetsrv\rewrite.dll") {
    Write-Info "URL Rewrite Module already installed — skipping"
} else {
    $proc = Start-Process msiexec.exe -ArgumentList "/i `"$($Config.RewriteMSI)`" /qn /norestart" -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -eq 0) {
        Write-Success "URL Rewrite Module installed"
    } else {
        Write-Failure "URL Rewrite install failed with exit code $($proc.ExitCode)"
        exit 1
    }

    if (Test-Path "$env:SystemRoot\System32\inetsrv\rewrite.dll") {
        Write-Success "rewrite.dll verified"
    } else {
        Write-Failure "rewrite.dll not found"
        exit 1
    }
}

# ============================================================================
# STEP 3: INSTALL APPLICATION REQUEST ROUTING 3.0
# ============================================================================

Write-Step "STEP 3: Installing Application Request Routing 3.0"

$arrExists = (Test-Path "C:\Program Files\IIS\Application Request Routing\requestRouter.dll") -or `
             (Test-Path "C:\Program Files (x86)\IIS\Application Request Routing\requestRouter.dll")

if ($arrExists) {
    Write-Info "ARR 3.0 already installed — skipping"
} else {
    $proc = Start-Process msiexec.exe -ArgumentList "/i `"$($Config.ARRMSI)`" /qn /norestart" -Wait -PassThru -NoNewWindow
    if ($proc.ExitCode -eq 0) {
        Write-Success "ARR 3.0 installed"
    } else {
        Write-Failure "ARR install failed with exit code $($proc.ExitCode)"
        exit 1
    }

    $arrFound = @(
        "C:\Program Files\IIS\Application Request Routing\requestRouter.dll",
        "C:\Program Files (x86)\IIS\Application Request Routing\requestRouter.dll"
    ) | Where-Object { Test-Path $_ }

    if ($arrFound) {
        Write-Success "ARR DLL verified"
    } else {
        Write-Failure "ARR DLL not found"
        exit 1
    }
}

# ============================================================================
# STEP 4: ENABLE ARR PROXY & UNLOCK WEBSOCKET SECTION
# ============================================================================

Write-Step "STEP 4: Configuring ARR Proxy & WebSocket"

# Enable ARR proxy
& $appcmd set config -section:system.webServer/proxy /enabled:"True" /commit:apphost 2>$null
& $appcmd set config -section:system.webServer/proxy /reverseRewriteHostInResponseHeaders:"False" /commit:apphost 2>$null

Write-Success "ARR proxy enabled"
Write-Success "reverseRewriteHostInResponseHeaders disabled"

# Unlock WebSocket section (prevents 500.19 errors in site-level web.config)
& $appcmd unlock config -section:system.webServer/webSocket 2>$null

Write-Success "WebSocket config section unlocked"

# ============================================================================
# STEP 5: CONFIGURE PHP FOR IIS
# ============================================================================

Write-Step "STEP 5: Configuring PHP for IIS"

# Create php.ini from production template (only if not already created)
$phpIniSource = "$($Config.PHPPath)\php.ini-production"
$phpIniDest = "$($Config.PHPPath)\php.ini"

if (Test-Path $phpIniDest) {
    Write-Info "php.ini already exists — skipping configuration"
    Write-Info "To force reconfiguration, delete $phpIniDest and re-run"
} else {
    if (-not (Test-Path $phpIniSource)) {
        Write-Failure "php.ini-production not found at $phpIniSource"
        exit 1
    }

    Copy-Item $phpIniSource $phpIniDest -Force
    $phpIni = Get-Content $phpIniDest

    # Enable extensions required by WordPress
    $extensionReplacements = @(
        @{ Find = ';extension_dir = "ext"';   Replace = "extension_dir = `"$($Config.PHPPath)\ext`"" },
        @{ Find = ';extension=curl';          Replace = 'extension=curl' },
        @{ Find = ';extension=fileinfo';      Replace = 'extension=fileinfo' },
        @{ Find = ';extension=gd';            Replace = 'extension=gd' },
        @{ Find = ';extension=mbstring';      Replace = 'extension=mbstring' },
        @{ Find = ';extension=mysqli';        Replace = 'extension=mysqli' },
        @{ Find = ';extension=openssl';       Replace = 'extension=openssl' },
        @{ Find = ';extension=pdo_mysql';     Replace = 'extension=pdo_mysql' },
        @{ Find = ';extension=soap';          Replace = 'extension=soap' },
        @{ Find = ';extension=exif';          Replace = 'extension=exif' }
    )

    foreach ($r in $extensionReplacements) {
        $phpIni = $phpIni -replace [regex]::Escape($r.Find), $r.Replace
    }

    # Set WordPress-friendly PHP limits
    $limitReplacements = @(
        @{ Find = 'upload_max_filesize = 2M';   Replace = 'upload_max_filesize = 64M' },
        @{ Find = 'post_max_size = 8M';         Replace = 'post_max_size = 64M' },
        @{ Find = 'max_execution_time = 30';    Replace = 'max_execution_time = 300' },
        @{ Find = 'max_input_time = 60';        Replace = 'max_input_time = 300' },
        @{ Find = 'memory_limit = 128M';        Replace = 'memory_limit = 256M' }
    )

    foreach ($r in $limitReplacements) {
        $phpIni = $phpIni -replace [regex]::Escape($r.Find), $r.Replace
    }

    Set-Content -Path $phpIniDest -Value $phpIni
    Write-Success "php.ini configured with extensions and WordPress limits"
}

# Add PHP to system PATH
$currentPath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
if ($currentPath -notlike "*$($Config.PHPPath)*") {
    [System.Environment]::SetEnvironmentVariable("Path", "$currentPath;$($Config.PHPPath)", "Machine")
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Write-Success "PHP added to system PATH"
} else {
    Write-Info "PHP already in system PATH"
}

# Register PHP with IIS FastCGI (only if not already registered)
$existingFastCGI = & $appcmd list config /section:system.webServer/fastCGI 2>$null
if ($existingFastCGI -match "php-cgi.exe") {
    Write-Info "PHP FastCGI application already registered — skipping"
} else {
    & $appcmd set config /section:system.webServer/fastCGI `
        /+"[fullPath='$($Config.PHPPath)\php-cgi.exe',maxInstances='4',activityTimeout='600',requestTimeout='600']" 2>$null
    Write-Success "PHP FastCGI application registered"
}

$existingHandler = & $appcmd list config /section:system.webServer/handlers 2>$null
if ($existingHandler -match "PHP_via_FastCGI") {
    Write-Info "PHP handler mapping already registered — skipping"
} else {
    & $appcmd set config /section:system.webServer/handlers `
        /+"[name='PHP_via_FastCGI',path='*.php',verb='*',modules='FastCgiModule',scriptProcessor='$($Config.PHPPath)\php-cgi.exe',resourceType='Either']" 2>$null
    Write-Success "PHP handler mapping registered"
}

# Verify PHP
$phpVersion = & "$($Config.PHPPath)\php.exe" -v 2>&1 | Select-Object -First 1
Write-Success "PHP version: $phpVersion"

$phpModules = & "$($Config.PHPPath)\php.exe" -m 2>&1
$requiredModules = @("curl", "gd", "mbstring", "mysqli", "openssl")
$missingModules = $requiredModules | Where-Object { $phpModules -notcontains $_ }
if ($missingModules.Count -eq 0) {
    Write-Success "All required PHP modules enabled"
} else {
    Write-Failure "Missing PHP modules: $($missingModules -join ', ')"
    exit 1
}

# ============================================================================
# STEP 6: CONFIGURE AND INITIALIZE MARIADB
# ============================================================================

Write-Step "STEP 6: Configuring and Initializing MariaDB"

# Create my.ini
$myCnf = @"
[mysqld]
basedir=$($Config.MariaDBPath -replace '\\','/')
datadir=$($Config.MariaDBDataDir -replace '\\','/')
port=$($Config.MariaDBPort)
max_allowed_packet=64M

[client]
port=$($Config.MariaDBPort)
"@

Set-Content -Path "$($Config.MariaDBPath)\my.ini" -Value $myCnf -Encoding ASCII
Write-Success "my.ini created"

# Initialize data directory (only if not already initialized)
if (Test-Path "$($Config.MariaDBDataDir)\mysql") {
    Write-Info "Data directory already initialized — skipping"
} else {
    Write-Info "Initializing MariaDB data directory..."

    # Determine which install tool is available
    $installDb = $null
    $candidates = @(
        "$($Config.MariaDBPath)\bin\mysql_install_db.exe",
        "$($Config.MariaDBPath)\bin\mariadb-install-db.exe"
    )
    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            $installDb = $candidate
            break
        }
    }

    if (-not $installDb) {
        Write-Failure "Cannot find mysql_install_db.exe or mariadb-install-db.exe in $($Config.MariaDBPath)\bin\"
        exit 1
    }

    Write-Info "Using: $installDb"

    $initOutput = & $installDb --datadir="$($Config.MariaDBDataDir)" --password="" 2>&1

    # Check if data directory was actually created
    if (Test-Path "$($Config.MariaDBDataDir)\mysql") {
        Write-Success "MariaDB data directory initialized"
    } else {
        Write-Failure "MariaDB initialization failed. Output:"
        $initOutput | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
        exit 1
    }
}

# Determine which daemon executable to use
$mariaDBDaemon = $null
$daemonCandidates = @(
    "$($Config.MariaDBPath)\bin\mysqld.exe",
    "$($Config.MariaDBPath)\bin\mariadbd.exe"
)
foreach ($candidate in $daemonCandidates) {
    if (Test-Path $candidate) {
        $mariaDBDaemon = $candidate
        break
    }
}

if (-not $mariaDBDaemon) {
    Write-Failure "Cannot find mysqld.exe or mariadbd.exe in $($Config.MariaDBPath)\bin\"
    exit 1
}

Write-Info "Using daemon: $mariaDBDaemon"

# Install as Windows service (if not already installed)
$existingService = Get-Service -Name $Config.MariaDBServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Info "MariaDB service already exists — skipping install"
} else {
    & $mariaDBDaemon --install $Config.MariaDBServiceName 2>$null
    Write-Success "MariaDB installed as Windows service"
}

# Start the service
$svc = Get-Service -Name $Config.MariaDBServiceName -ErrorAction SilentlyContinue
if (-not $svc) {
    Write-Failure "MariaDB service not found after install"
    exit 1
}

if ($svc.Status -ne "Running") {
    Start-Service $Config.MariaDBServiceName
    Start-Sleep -Seconds 3
    $svc = Get-Service -Name $Config.MariaDBServiceName
}

if ($svc.Status -eq "Running") {
    Write-Success "MariaDB service is running"
} else {
    Write-Failure "MariaDB service failed to start (Status: $($svc.Status))"
    exit 1
}

# Verify MariaDB
$mariaVersion = & "$($Config.MariaDBPath)\bin\mysql.exe" -u root -e "SELECT VERSION();" 2>$null
if ($mariaVersion) {
    Write-Success "MariaDB responding: $($mariaVersion | Select-Object -Last 1)"
} else {
    Write-Failure "Cannot connect to MariaDB"
    exit 1
}

# ============================================================================
# STEP 7: GENERATE SELF-SIGNED WILDCARD CERTIFICATE
# ============================================================================

Write-Step "STEP 7: Generating Self-Signed Wildcard Certificate"

# Check if cert already exists
$existingCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.FriendlyName -eq $Config.CertFriendlyName }

if ($existingCert) {
    Write-Info "Certificate already exists — using existing (Thumbprint: $($existingCert.Thumbprint))"
    $cert = $existingCert
} else {
    $cert = New-SelfSignedCertificate `
        -DnsName "*.$($Config.LabDomain)", $Config.LabDomain `
        -CertStoreLocation "Cert:\LocalMachine\My" `
        -FriendlyName $Config.CertFriendlyName `
        -NotAfter (Get-Date).AddYears($Config.CertValidityYears) `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -HashAlgorithm SHA256 `
        -KeyExportPolicy Exportable

    Write-Success "Certificate created"
}

# Save thumbprint
$cert.Thumbprint | Out-File $Config.ThumbprintFile -Encoding ASCII -NoNewline
Write-Success "Thumbprint saved to $($Config.ThumbprintFile)"
Write-Info "Thumbprint: $($cert.Thumbprint)"
Write-Info "Subject: $($cert.Subject)"
Write-Info "Expires: $($cert.NotAfter)"

# ============================================================================
# STEP 8: CONFIGURE HOSTS FILE
# ============================================================================

Write-Step "STEP 8: Configuring Hosts File"

$hostsFile = "$env:SystemRoot\System32\drivers\etc\hosts"
$hostsContent = Get-Content $hostsFile

$entries = @(
    "127.0.0.1    $($Config.JuiceShopHost)",
    "127.0.0.1    $($Config.WordPressHost)",
    "127.0.0.1    $($Config.EchoSPAHost)"
)

foreach ($entry in $entries) {
    if ($hostsContent -notcontains $entry) {
        Add-Content -Path $hostsFile -Value $entry
        Write-Success "Added: $entry"
    } else {
        Write-Info "Already exists: $entry"
    }
}

# ============================================================================
# STEP 9: REMOVE DEFAULT WEB SITE & CONFIGURE IIS LOGGING
# ============================================================================

Write-Step "STEP 9: Configuring IIS (Default Site, Logging)"

# Remove Default Web Site
$defaultSiteExists = & $appcmd list site "Default Web Site" 2>$null
if ($defaultSiteExists) {
    & $appcmd stop site "Default Web Site" 2>$null
    & $appcmd delete site "Default Web Site" 2>$null
    Write-Success "Default Web Site removed"
} else {
    Write-Info "Default Web Site already removed"
}

# Configure enhanced W3C logging at server level
& $appcmd set config -section:system.applicationHost/sites `
    /siteDefaults.logFile.logFormat:"W3C" /commit:apphost 2>$null

& $appcmd set config -section:system.applicationHost/sites `
    /siteDefaults.logFile.logExtFileFlags:"Date,Time,ClientIP,UserName,SiteName,ServerIP,Method,UriStem,UriQuery,HttpStatus,HttpSubStatus,Win32Status,TimeTaken,ServerPort,UserAgent,Referer,Host,BytesSent,BytesRecv,ProtocolVersion" `
    /commit:apphost 2>$null

Write-Success "W3C logging standard fields configured"

# Add custom log fields for WAF/proxy headers (only if not already present)
$existingLogConfig = & $appcmd list config -section:system.applicationHost/sites /text:* 2>$null
$existingLogConfigText = $existingLogConfig -join "`n"

$customFields = @(
    @{ Name = "X-Forwarded-For";    Source = "X-Forwarded-For";    SourceType = "RequestHeader" },
    @{ Name = "X-Forwarded-Proto";  Source = "X-Forwarded-Proto";  SourceType = "RequestHeader" },
    @{ Name = "X-Forwarded-Host";   Source = "X-Forwarded-Host";   SourceType = "RequestHeader" },
    @{ Name = "True-Client-IP";     Source = "True-Client-IP";     SourceType = "RequestHeader" }
)

foreach ($field in $customFields) {
    if ($existingLogConfigText -match $field.Name) {
        Write-Info "Custom log field already exists: $($field.Name) — skipping"
    } else {
        & $appcmd set config -section:system.applicationHost/sites `
            /+"siteDefaults.logFile.customFields.[logFieldName='$($field.Name)',sourceName='$($field.Source)',sourceType='$($field.SourceType)']" `
            /commit:apphost 2>$null
        Write-Success "Custom log field added: $($field.Name)"
    }
}

# ============================================================================
# STEP 10: CREATE SITE DIRECTORIES
# ============================================================================

Write-Step "STEP 10: Creating Site Directories"

$siteDirs = @(
    "C:\inetpub\juiceshop",
    "C:\inetpub\wordpress",
    "C:\inetpub\echospa"
)

foreach ($dir in $siteDirs) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
        Write-Success "Created: $dir"
    } else {
        Write-Info "Already exists: $dir"
    }
}

# ============================================================================
# STEP 11: IIS RESET
# ============================================================================

Write-Step "STEP 11: Restarting IIS"

iisreset /restart 2>$null | Out-Null
Start-Sleep -Seconds 3

$w3svc = Get-Service W3SVC
if ($w3svc.Status -eq "Running") {
    Write-Success "IIS restarted and running"
} else {
    Write-Failure "IIS may not have restarted properly (Status: $($w3svc.Status))"
}

# ============================================================================
# FINAL VERIFICATION
# ============================================================================

Write-Step "FINAL VERIFICATION"

$verifications = @(
    @{ Name = "IIS Feature: Web-CGI";           Test = { (Get-WindowsFeature Web-CGI).InstallState -eq "Installed" } },
    @{ Name = "IIS Feature: Web-Asp-Net45";     Test = { (Get-WindowsFeature Web-Asp-Net45).InstallState -eq "Installed" } },
    @{ Name = "IIS Feature: Web-WebSockets";    Test = { (Get-WindowsFeature Web-WebSockets).InstallState -eq "Installed" } },
    @{ Name = "IIS Feature: Web-Http-Tracing";  Test = { (Get-WindowsFeature Web-Http-Tracing).InstallState -eq "Installed" } },
    @{ Name = "IIS Feature: Web-Custom-Logging"; Test = { (Get-WindowsFeature Web-Custom-Logging).InstallState -eq "Installed" } },
    @{ Name = "URL Rewrite DLL";                Test = { Test-Path "$env:SystemRoot\System32\inetsrv\rewrite.dll" } },
    @{ Name = "ARR DLL";                        Test = { (Test-Path "C:\Program Files\IIS\Application Request Routing\requestRouter.dll") -or (Test-Path "C:\Program Files (x86)\IIS\Application Request Routing\requestRouter.dll") } },
    @{ Name = "VC++ Redistributable";           Test = { Test-Path "$env:SystemRoot\System32\vcruntime140.dll" } },
    @{ Name = "PHP accessible";                 Test = { (& "$($Config.PHPPath)\php.exe" -v 2>$null) -match "PHP" } },
    @{ Name = "PHP in IIS handlers";            Test = { (& $appcmd list config /section:handlers 2>$null) -match "PHP_via_FastCGI" } },
    @{ Name = "MariaDB service running";        Test = { (Get-Service $Config.MariaDBServiceName).Status -eq "Running" } },
    @{ Name = "MariaDB responds";               Test = { (& "$($Config.MariaDBPath)\bin\mysql.exe" -u root -e "SELECT 1;" 2>$null) -match "1" } },
    @{ Name = "SSL Certificate";                Test = { Test-Path $Config.ThumbprintFile } },
    @{ Name = "Hosts file entries";             Test = { (Get-Content "$env:SystemRoot\System32\drivers\etc\hosts") -match $Config.LabDomain } },
    @{ Name = "IIS running";                    Test = { (Get-Service W3SVC).Status -eq "Running" } },
    @{ Name = "Default Web Site removed";       Test = { -not (& $appcmd list site "Default Web Site" 2>$null) } }
)

$passCount = 0
$failCount = 0

foreach ($v in $verifications) {
    try {
        if (& $v.Test) {
            Write-Success $v.Name
            $passCount++
        } else {
            Write-Failure $v.Name
            $failCount++
        }
    } catch {
        Write-Failure "$($v.Name) — $($_.Exception.Message)"
        $failCount++
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n$("=" * 60)" -ForegroundColor Cyan
Write-Host "  PREREQUISITES INSTALLATION COMPLETE" -ForegroundColor Cyan
Write-Host "$("=" * 60)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Results: $passCount passed, $failCount failed" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Red" })
Write-Host ""
Write-Host "  Installed Components:" -ForegroundColor White
Write-Host "    • IIS with CGI, ASP.NET 4.5, WebSockets, Tracing" -ForegroundColor Gray
Write-Host "    • URL Rewrite Module 2.1" -ForegroundColor Gray
Write-Host "    • Application Request Routing 3.0 (proxy enabled)" -ForegroundColor Gray
Write-Host "    • PHP $((& "$($Config.PHPPath)\php.exe" -v 2>$null | Select-Object -First 1) -replace 'PHP ([0-9.]+).*','$1') (FastCGI)" -ForegroundColor Gray
Write-Host "    • MariaDB $((& "$($Config.MariaDBPath)\bin\mysql.exe" -u root -e "SELECT VERSION();" 2>$null | Select-Object -Last 1))" -ForegroundColor Gray
Write-Host "    • SSL Certificate: *.$($Config.LabDomain)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Certificate Thumbprint:" -ForegroundColor White
Write-Host "    $(Get-Content $Config.ThumbprintFile)" -ForegroundColor Gray
Write-Host ""

if ($failCount -eq 0) {
    Write-Host "  ✅ Ready to run Install-JuiceShop.ps1" -ForegroundColor Green
    Write-Host "  ✅ Ready to run Install-WordPress.ps1" -ForegroundColor Green
    Write-Host "  ✅ Ready to run Install-EchoSPA.ps1" -ForegroundColor Green
} else {
    Write-Host "  ⚠️  Fix failed items before running app install scripts." -ForegroundColor Red
}

Write-Host ""
Write-Host "$("=" * 60)" -ForegroundColor Cyan
