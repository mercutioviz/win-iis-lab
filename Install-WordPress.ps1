#Requires -RunAsAdministrator
#Requires -Version 7.0

<#
.SYNOPSIS
    IIS Lab Environment - WordPress Installation Script
.DESCRIPTION
    Deploys WordPress on IIS with PHP FastCGI and MariaDB:
    - Creates MariaDB database and user
    - Generates wp-config.php with salt keys and WAF-aware HTTPS detection
    - Sets IIS file permissions
    - Creates IIS site with HTTPS/SNI binding
    - Configures URL Rewrite rules for WordPress permalinks
    - Enables Failed Request Tracing (FREB)
    - Optionally adds a real domain binding for WAF testing
.NOTES
    Run Install-Prerequisites.ps1 first.
    WordPress must already be extracted to C:\inetpub\wordpress\
#>

# ============================================================================
# CONFIGURATION — Adjust these values as needed
# ============================================================================

$Config = @{
    # Site name and paths
    SiteName            = "WordPress"
    SiteRoot            = "C:\inetpub\wordpress"

    # Hostnames
    LocalHostname       = "wordpress.lab.local"
    RealDomain          = ""   # Set to e.g. "wordpress.waaslab.com" to add a WAF-facing binding, or leave empty

    # MariaDB
    MariaDBBin          = "C:\MariaDB\bin"
    DBName              = "wordpress"
    DBUser              = "wpuser"
    DBPass              = "WPlab2024!"
    DBHost              = "localhost"

    # Certificate
    ThumbprintFile      = "C:\LabSetup\cert-thumbprint.txt"

    # PHP
    PHPPath             = "C:\PHP"

    # FREB
    FrebEnabled         = $true
    FrebStatusCodes     = "400-599"
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
$mysql = "$($Config.MariaDBBin)\mysql.exe"

# ============================================================================
# ENVIRONMENT SETUP
# ============================================================================

Write-Step "ENVIRONMENT SETUP"

$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
Write-Success "PATH refreshed"

# ============================================================================
# PRE-FLIGHT CHECKS
# ============================================================================

Write-Step "PRE-FLIGHT CHECKS"

$preflight = @(
    @{ Name = "WordPress files extracted";       Test = { Test-Path "$($Config.SiteRoot)\wp-login.php" } },
    @{ Name = "PHP accessible";                  Test = { (& "$($Config.PHPPath)\php.exe" -v 2>$null) -match "PHP" } },
    @{ Name = "PHP registered in IIS";           Test = { (& $appcmd list config /section:handlers 2>$null) -match "PHP_via_FastCGI" } },
    @{ Name = "MariaDB service running";         Test = { (Get-Service MariaDB -ErrorAction SilentlyContinue).Status -eq "Running" } },
    @{ Name = "MariaDB mysql.exe exists";        Test = { Test-Path $mysql } },
    @{ Name = "Certificate thumbprint file";     Test = { Test-Path $Config.ThumbprintFile } },
    @{ Name = "URL Rewrite installed";           Test = { Test-Path "$env:SystemRoot\System32\inetsrv\rewrite.dll" } },
    @{ Name = "IIS running";                     Test = { (Get-Service W3SVC).Status -eq "Running" } }
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
    Write-Failure "PRE-FLIGHT CHECKS FAILED. Fix issues above and re-run."
    exit 1
}

Write-Success "All pre-flight checks passed!`n"

# ============================================================================
# STEP 1: CREATE MARIADB DATABASE AND USER
# ============================================================================

Write-Step "STEP 1: Creating MariaDB Database and User"

# Check if database already exists
$existingDB = & $mysql -u root -e "SHOW DATABASES LIKE '$($Config.DBName)';" 2>$null
if ($existingDB -match $Config.DBName) {
    Write-Info "Database '$($Config.DBName)' already exists — skipping creation"
} else {
    $createDB = "CREATE DATABASE $($Config.DBName) DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;"
    $createDB | & $mysql -u root 2>$null

    # Verify
    $verifyDB = & $mysql -u root -e "SHOW DATABASES LIKE '$($Config.DBName)';" 2>$null
    if ($verifyDB -match $Config.DBName) {
        Write-Success "Database '$($Config.DBName)' created"
    } else {
        Write-Failure "Failed to create database '$($Config.DBName)'"
        exit 1
    }
}

# Check if user already exists
$existingUser = & $mysql -u root -e "SELECT User FROM mysql.user WHERE User='$($Config.DBUser)' AND Host='localhost';" 2>$null
if ($existingUser -match $Config.DBUser) {
    Write-Info "User '$($Config.DBUser)' already exists — skipping creation"
} else {
    $createUser = @"
CREATE USER '$($Config.DBUser)'@'localhost' IDENTIFIED BY '$($Config.DBPass)';
GRANT ALL PRIVILEGES ON $($Config.DBName).* TO '$($Config.DBUser)'@'localhost';
FLUSH PRIVILEGES;
"@
    $createUser | & $mysql -u root 2>$null

    # Verify
    $verifyUser = & $mysql -u root -e "SELECT User FROM mysql.user WHERE User='$($Config.DBUser)' AND Host='localhost';" 2>$null
    if ($verifyUser -match $Config.DBUser) {
        Write-Success "User '$($Config.DBUser)' created with grants on '$($Config.DBName)'"
    } else {
        Write-Failure "Failed to create user '$($Config.DBUser)'"
        exit 1
    }
}

# Test the user can connect
$testConn = & $mysql -u $Config.DBUser -p"$($Config.DBPass)" -e "SELECT 1;" 2>$null
if ($testConn -match "1") {
    Write-Success "User '$($Config.DBUser)' can connect to MariaDB"
} else {
    Write-Failure "User '$($Config.DBUser)' cannot connect — check credentials"
    exit 1
}

# ============================================================================
# STEP 2: GENERATE WP-CONFIG.PHP
# ============================================================================

Write-Step "STEP 2: Generating wp-config.php"

if (Test-Path "$($Config.SiteRoot)\wp-config.php") {
    Write-Info "wp-config.php already exists — skipping"
    Write-Info "To force regeneration, delete $($Config.SiteRoot)\wp-config.php and re-run"
} else {
    # Fetch salt keys from WordPress API
    Write-Info "Fetching salt keys from api.wordpress.org..."
    try {
        $saltKeys = (Invoke-WebRequest -Uri "https://api.wordpress.org/secret-key/1.1/salt/" -UseBasicParsing -TimeoutSec 15).Content
        Write-Success "Salt keys retrieved"
    } catch {
        Write-Info "Could not fetch salt keys from API — generating locally"
        # Generate fallback salt keys locally
        $chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+[]{}|;:,.<>?'
        function Get-RandomSalt {
            -join (1..64 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
        }
        $saltNames = @("AUTH_KEY","SECURE_AUTH_KEY","LOGGED_IN_KEY","NONCE_KEY","AUTH_SALT","SECURE_AUTH_SALT","LOGGED_IN_SALT","NONCE_SALT")
        $saltKeys = ($saltNames | ForEach-Object { "define( '$_', '$(Get-RandomSalt)' );" }) -join "`n"
        Write-Success "Local salt keys generated"
    }

    $wpConfig = @"
<?php
/**
 * WordPress Configuration - Generated by Install-WordPress.ps1
 */

// Database settings
define( 'DB_NAME', '$($Config.DBName)' );
define( 'DB_USER', '$($Config.DBUser)' );
define( 'DB_PASSWORD', '$($Config.DBPass)' );
define( 'DB_HOST', '$($Config.DBHost)' );
define( 'DB_CHARSET', 'utf8mb4' );
define( 'DB_COLLATE', '' );

// Authentication unique keys and salts
$saltKeys

// Table prefix
`$table_prefix = 'wp_';

// Debug mode
define( 'WP_DEBUG', false );

// Handle HTTPS behind reverse proxy / WAF
// Barracuda WaaS and other proxies send X-Forwarded-Proto
if ( isset( `$_SERVER['HTTP_X_FORWARDED_PROTO'] ) && `$_SERVER['HTTP_X_FORWARDED_PROTO'] === 'https' ) {
    `$_SERVER['HTTPS'] = 'on';
}

// Absolute path to the WordPress directory
if ( ! defined( 'ABSPATH' ) ) {
    define( 'ABSPATH', __DIR__ . '/' );
}

// Load WordPress
require_once ABSPATH . 'wp-settings.php';
"@

    Set-Content -Path "$($Config.SiteRoot)\wp-config.php" -Value $wpConfig -Encoding UTF8
    Write-Success "wp-config.php created"
}

# ============================================================================
# STEP 3: SET FILE PERMISSIONS
# ============================================================================

Write-Step "STEP 3: Setting File Permissions"

$acl = Get-Acl $Config.SiteRoot

# Grant IIS_IUSRS
$iisUsersRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "IIS_IUSRS", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
)

# Grant IUSR
$iusrRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "IUSR", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
)

# Check if permissions already set
$existingRules = $acl.Access | Where-Object { $_.IdentityReference -match "IIS_IUSRS|IUSR" -and $_.FileSystemRights -match "FullControl" }

if ($existingRules.Count -ge 2) {
    Write-Info "IIS permissions already set — skipping"
} else {
    $acl.SetAccessRule($iisUsersRule)
    $acl.SetAccessRule($iusrRule)
    Set-Acl $Config.SiteRoot $acl
    Write-Success "IIS_IUSRS granted FullControl"
    Write-Success "IUSR granted FullControl"
}

# ============================================================================
# STEP 4: CREATE IIS SITE WITH HTTPS/SNI BINDING
# ============================================================================

Write-Step "STEP 4: Creating IIS Site"

$thumbprint = (Get-Content $Config.ThumbprintFile).Trim()

# Check if site already exists
$existingSite = & $appcmd list site $Config.SiteName 2>$null

if ($existingSite) {
    Write-Info "IIS site '$($Config.SiteName)' already exists — skipping creation"
} else {
    # Create the site with correct binding format: https/*:443:hostname
    & $appcmd add site /name:"$($Config.SiteName)" `
        /physicalPath:"$($Config.SiteRoot)" `
        /bindings:"https/*:443:$($Config.LocalHostname)" 2>$null | Out-Null

    # Enable SNI on the binding
    & $appcmd set site "$($Config.SiteName)" `
        /"bindings.[protocol='https',bindingInformation='*:443:$($Config.LocalHostname)'].sslFlags:1" 2>$null | Out-Null

    Write-Success "IIS site '$($Config.SiteName)' created with SNI"
}

# Bind SSL certificate via netsh (delete first for idempotency)
netsh http delete sslcert hostnameport="$($Config.LocalHostname):443" 2>$null | Out-Null

$appId = [guid]::NewGuid().ToString()
$netshResult = netsh http add sslcert hostnameport="$($Config.LocalHostname):443" `
    certhash=$thumbprint `
    certstorename=MY `
    appid="{$appId}" 2>&1

if ($netshResult -match "successfully") {
    Write-Success "SSL certificate bound to $($Config.LocalHostname):443"
} else {
    Write-Failure "SSL cert binding failed: $netshResult"
    exit 1
}

# Start the site
& $appcmd start site "$($Config.SiteName)" 2>$null | Out-Null
Write-Success "IIS site started"

# ============================================================================
# STEP 5: CONFIGURE URL REWRITE FOR WORDPRESS PERMALINKS
# ============================================================================

Write-Step "STEP 5: Configuring URL Rewrite Rules"

$wpWebConfig = @'
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <rewrite>
            <rules>
                <rule name="WordPress" stopProcessing="true">
                    <match url=".*" />
                    <conditions logicalGrouping="MatchAll">
                        <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
                        <add input="{REQUEST_FILENAME}" matchType="IsDirectory" negate="true" />
                    </conditions>
                    <action type="Rewrite" url="index.php" />
                </rule>
            </rules>
        </rewrite>
        <defaultDocument>
            <files>
                <clear />
                <add value="index.php" />
                <add value="index.html" />
            </files>
        </defaultDocument>
    </system.webServer>
</configuration>
'@

Set-Content -Path "$($Config.SiteRoot)\web.config" -Value $wpWebConfig -Encoding UTF8
Write-Success "WordPress web.config with permalink rewrite rules created"

# ============================================================================
# STEP 6: ADD REAL DOMAIN BINDING (OPTIONAL)
# ============================================================================

if ($Config.RealDomain) {
    Write-Step "STEP 6: Adding Real Domain Binding ($($Config.RealDomain))"

    # Check if binding already exists
    $existingBindings = & $appcmd list site "$($Config.SiteName)" /text:bindings 2>$null
    if ($existingBindings -match [regex]::Escape($Config.RealDomain)) {
        Write-Info "Binding for $($Config.RealDomain) already exists — skipping"
    } else {
        & $appcmd set site "$($Config.SiteName)" `
            /+"bindings.[protocol='https',bindingInformation='*:443:$($Config.RealDomain)',sslFlags='1']" 2>$null | Out-Null

        Write-Success "HTTPS binding added for $($Config.RealDomain)"
    }

    # Bind SSL cert via netsh
    netsh http delete sslcert hostnameport="$($Config.RealDomain):443" 2>$null | Out-Null

    $appId2 = [guid]::NewGuid().ToString()
    $netshResult2 = netsh http add sslcert hostnameport="$($Config.RealDomain):443" `
        certhash=$thumbprint `
        certstorename=MY `
        appid="{$appId2}" 2>&1

    if ($netshResult2 -match "successfully") {
        Write-Success "SSL certificate bound to $($Config.RealDomain):443"
    } else {
        Write-Info "SSL cert binding note: $netshResult2"
    }
} else {
    Write-Step "STEP 6: Real Domain Binding — SKIPPED (not configured)"
    Write-Info "Set `$Config.RealDomain to add a WAF-facing binding"
}

# ============================================================================
# STEP 7: ENABLE FAILED REQUEST TRACING (FREB)
# ============================================================================

if ($Config.FrebEnabled) {
    Write-Step "STEP 7: Enabling Failed Request Tracing"

    & $appcmd configure trace "$($Config.SiteName)" /enablesite 2>$null | Out-Null
    Write-Success "FREB enabled for $($Config.SiteName)"

    $existingRules = & $appcmd list config "$($Config.SiteName)" /section:system.webServer/tracing/traceFailedRequests 2>$null

    if ($existingRules -match "path=") {
        Write-Info "FREB rules already configured — skipping"
    } else {
        & $appcmd set config "$($Config.SiteName)" /section:system.webServer/tracing/traceFailedRequests `
            /+"[path='*']" 2>$null | Out-Null

        & $appcmd set config "$($Config.SiteName)" /section:system.webServer/tracing/traceFailedRequests `
            /"[path='*'].traceAreas.[provider='WWW Server',areas='Authentication,Security,Filter,StaticFile,CGI,Compression,Cache,RequestNotifications,Module,Rewrite,iisGeneral',verbosity='Verbose']" 2>$null | Out-Null

        & $appcmd set config "$($Config.SiteName)" /section:system.webServer/tracing/traceFailedRequests `
            /"[path='*'].failureDefinitions.statusCodes:$($Config.FrebStatusCodes)" 2>$null | Out-Null

        Write-Success "FREB rule created: status codes $($Config.FrebStatusCodes)"
    }
} else {
    Write-Step "STEP 7: Failed Request Tracing — SKIPPED (disabled in config)"
}

# ============================================================================
# STEP 8: RESTART SITE & VALIDATE
# ============================================================================

Write-Step "STEP 8: Validation"

& $appcmd stop site "$($Config.SiteName)" 2>$null | Out-Null
Start-Sleep -Seconds 1
& $appcmd start site "$($Config.SiteName)" 2>$null | Out-Null
Start-Sleep -Seconds 2

# Test HTTPS through IIS
Write-Info "Testing https://$($Config.LocalHostname)..."

try {
    $response = Invoke-WebRequest -Uri "https://$($Config.LocalHostname)" `
        -SkipCertificateCheck -TimeoutSec 15 -MaximumRedirection 5

    Write-Success "HTTP Status: $($response.StatusCode)"
    Write-Success "Content Length: $($response.Content.Length) bytes"

    if ($response.Content -match "WordPress" -or $response.Content -match "wp-") {
        Write-Success "WordPress content confirmed!"
    } elseif ($response.Content -match "install.php" -or $response.Content -match "setup-config") {
        Write-Success "WordPress installation wizard detected — ready for setup!"
    } else {
        Write-Info "Got a response but WordPress content not detected — check manually"
    }
} catch {
    Write-Failure "HTTPS test failed: $($_.Exception.Message)"
    Write-Info "Check that the site is running: & $appcmd list site $($Config.SiteName)"
}

# Quick PHP test
Write-Info "Testing PHP execution..."
$phpTestFile = "$($Config.SiteRoot)\iis-php-test.php"
Set-Content -Path $phpTestFile -Value "<?php echo 'PHP_OK_' . phpversion(); ?>"

try {
    $phpResponse = Invoke-WebRequest -Uri "https://$($Config.LocalHostname)/iis-php-test.php" `
        -SkipCertificateCheck -TimeoutSec 10

    if ($phpResponse.Content -match "PHP_OK_") {
        $detectedVersion = ($phpResponse.Content -replace "PHP_OK_", "").Trim()
        Write-Success "PHP executing correctly (version $detectedVersion)"
    } else {
        Write-Failure "PHP not executing — response was: $($phpResponse.Content)"
    }
} catch {
    Write-Failure "PHP test failed: $($_.Exception.Message)"
}

# Clean up test file
Remove-Item $phpTestFile -Force -ErrorAction SilentlyContinue

# Test real domain if configured
if ($Config.RealDomain) {
    Write-Info "Testing https://$($Config.RealDomain)..."
    try {
        $response2 = Invoke-WebRequest -Uri "https://$($Config.RealDomain)" -SkipCertificateCheck -TimeoutSec 10
        Write-Success "Real domain test: HTTP $($response2.StatusCode)"
    } catch {
        Write-Info "Real domain not reachable yet (expected if DNS not configured)"
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n$("=" * 60)" -ForegroundColor Cyan
Write-Host "  WORDPRESS INSTALLATION COMPLETE" -ForegroundColor Cyan
Write-Host "$("=" * 60)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  IIS HTTPS URL:        https://$($Config.LocalHostname)" -ForegroundColor Gray
if ($Config.RealDomain) {
    Write-Host "  WAF Domain URL:       https://$($Config.RealDomain)" -ForegroundColor Gray
}
Write-Host "  Database:             $($Config.DBName) @ $($Config.DBHost)" -ForegroundColor Gray
Write-Host "  Database User:        $($Config.DBUser)" -ForegroundColor Gray
Write-Host "  Site Root:            $($Config.SiteRoot)" -ForegroundColor Gray
Write-Host "  FREB:                 $(if ($Config.FrebEnabled) {'Enabled'} else {'Disabled'})" -ForegroundColor Gray
Write-Host ""
Write-Host "  ⚠️  ACTION REQUIRED:" -ForegroundColor Yellow
Write-Host "     Open https://$($Config.LocalHostname) in a browser" -ForegroundColor Yellow
Write-Host "     to complete the WordPress installation wizard." -ForegroundColor Yellow
Write-Host ""
Write-Host "     Suggested settings:" -ForegroundColor Gray
Write-Host "       Site Title:  Lab WordPress" -ForegroundColor Gray
Write-Host "       Username:    admin" -ForegroundColor Gray
Write-Host "       Email:       admin@lab.local" -ForegroundColor Gray
Write-Host "       Search:      ☑ Discourage search engines" -ForegroundColor Gray
Write-Host ""
Write-Host "$("=" * 60)" -ForegroundColor Cyan
