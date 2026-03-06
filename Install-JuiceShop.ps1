#Requires -RunAsAdministrator
#Requires -Version 7.0

<#
.SYNOPSIS
    IIS Lab Environment - OWASP Juice Shop Installation Script
.DESCRIPTION
    Deploys OWASP Juice Shop behind IIS as a reverse proxy with HTTPS:
    - Registers Juice Shop as a persistent Windows service (WinSW)
    - Creates IIS site with HTTPS/SNI binding
    - Configures reverse proxy rules (URL Rewrite + ARR)
    - Enables Failed Request Tracing (FREB)
    - Optionally adds a real domain binding for WAF testing
.NOTES
    Run Install-Prerequisites.ps1 first.
    Juice Shop must already be extracted to C:\inetpub\juiceshop\
#>

# ============================================================================
# CONFIGURATION — Adjust these values as needed
# ============================================================================

$Config = @{
    # Site name and paths
    SiteName            = "JuiceShop"
    SiteRoot            = "C:\inetpub\juiceshop"
    LogsPath            = "C:\inetpub\juiceshop\logs"

    # Hostnames
    LocalHostname       = "juiceshop.lab.local"
    RealDomain          = ""   # Set to e.g. "juiceshop.waaslab.com" to add a WAF-facing binding, or leave empty

    # Node.js backend
    NodePort            = 3000
    StartupWaitSeconds  = 60

    # Service
    ServiceName         = "JuiceShop"
    ServiceDisplayName  = "OWASP Juice Shop"
    ServiceDescription  = "OWASP Juice Shop - Intentionally Vulnerable Web Application"

    # Files from prerequisites
    WinSW               = "C:\LabSetup\WinSW-x64.exe"
    ThumbprintFile      = "C:\LabSetup\cert-thumbprint.txt"

    # Node.js
    ExpectedNodeMajor   = 24
    ExpectedModuleVer   = 137

    # FREB
    FrebEnabled         = $true
    FrebMaxFiles        = 50
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
    @{ Name = "Node.js $($Config.ExpectedNodeMajor)";   Test = { (node --version 2>$null) -match "^v$($Config.ExpectedNodeMajor)\." } },
    @{ Name = "Node MODULE_VERSION $($Config.ExpectedModuleVer)"; Test = { (node -e "console.log(process.versions.modules)" 2>$null).Trim() -eq "$($Config.ExpectedModuleVer)" } },
    @{ Name = "Juice Shop extracted";                    Test = { Test-Path "$($Config.SiteRoot)\package.json" } },
    @{ Name = "Certificate thumbprint file";             Test = { Test-Path $Config.ThumbprintFile } },
    @{ Name = "WinSW executable";                        Test = { Test-Path $Config.WinSW } },
    @{ Name = "ARR proxy enabled";                       Test = { (& $appcmd list config /section:system.webServer/proxy 2>$null) -match 'enabled="true"' } },
    @{ Name = "URL Rewrite installed";                   Test = { Test-Path "$env:SystemRoot\System32\inetsrv\rewrite.dll" } },
    @{ Name = "IIS running";                             Test = { (Get-Service W3SVC).Status -eq "Running" } }
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
# STEP 1: IDENTIFY JUICE SHOP ENTRY POINT
# ============================================================================

Write-Step "STEP 1: Identifying Juice Shop Entry Point"

$packageJson = Get-Content "$($Config.SiteRoot)\package.json" | ConvertFrom-Json
$jsVersion = $packageJson.version
Write-Info "Juice Shop version: $jsVersion"

$startScript = $packageJson.scripts.start
Write-Info "npm start command: $startScript"

# Parse the entry point from "node build/app" or similar
if ($startScript -match "node\s+(.+)$") {
    $entryPoint = $Matches[1].Trim()
} else {
    Write-Failure "Could not parse entry point from start script: $startScript"
    exit 1
}

Write-Success "Entry point: $entryPoint"

# Convert forward slashes to backslashes for Windows
$entryPointWindows = $entryPoint -replace "/", "\"

# Verify the entry point file exists (with or without .js extension)
$entryPointPath = "$($Config.SiteRoot)\$entryPointWindows"
if (-not (Test-Path $entryPointPath) -and -not (Test-Path "$entryPointPath.js")) {
    Write-Failure "Entry point file not found: $entryPointPath (or $entryPointPath.js)"
    exit 1
}

Write-Success "Entry point file verified"

# ============================================================================
# STEP 2: INSTALL WINSW SERVICE
# ============================================================================

Write-Step "STEP 2: Configuring Juice Shop as a Windows Service"

$serviceExe = "$($Config.SiteRoot)\$($Config.ServiceName)Service.exe"
$serviceXml = "$($Config.SiteRoot)\$($Config.ServiceName)Service.xml"

# Check if service already exists and is running
$existingService = Get-Service -Name $Config.ServiceName -ErrorAction SilentlyContinue
if ($existingService -and $existingService.Status -eq "Running") {
    Write-Info "Juice Shop service already running — skipping service setup"
} else {
    # Copy WinSW executable (if not already there)
    if (-not (Test-Path $serviceExe)) {
        Copy-Item $Config.WinSW -Destination $serviceExe -Force
        Write-Success "WinSW executable copied to $serviceExe"
    } else {
        Write-Info "WinSW executable already in place"
    }

    # Create WinSW XML config
    $nodePath = (Get-Command node -ErrorAction Stop).Source

    $winswConfig = @"
<service>
    <id>$($Config.ServiceName)</id>
    <name>$($Config.ServiceDisplayName)</name>
    <description>$($Config.ServiceDescription)</description>
    <executable>$nodePath</executable>
    <arguments>$entryPointWindows</arguments>
    <workingdirectory>$($Config.SiteRoot)</workingdirectory>
    <logpath>$($Config.LogsPath)</logpath>
    <log mode="roll-by-size">
        <sizeThreshold>10240</sizeThreshold>
        <keepFiles>3</keepFiles>
    </log>
    <startmode>Automatic</startmode>
    <onfailure action="restart" delay="10 sec"/>
    <onfailure action="restart" delay="20 sec"/>
    <onfailure action="none"/>
</service>
"@

    Set-Content -Path $serviceXml -Value $winswConfig -Encoding UTF8
    Write-Success "WinSW config created"

    # Create logs directory
    New-Item -Path $Config.LogsPath -ItemType Directory -Force | Out-Null

    # Install the service (if not already installed)
    if (-not $existingService) {
        Set-Location $Config.SiteRoot
        $installOutput = & $serviceExe install 2>&1
        Write-Success "Service installed: $installOutput"
    } else {
        Write-Info "Service already registered — restarting"
    }

    # Start the service
    Start-Service $Config.ServiceName -ErrorAction Stop
    Write-Success "Service started"
}

# ============================================================================
# STEP 3: WAIT FOR JUICE SHOP TO BE READY
# ============================================================================

Write-Step "STEP 3: Waiting for Juice Shop to Start"

$ready = $false
$elapsed = 0
$interval = 5

while ($elapsed -lt $Config.StartupWaitSeconds) {
    try {
        $response = Invoke-WebRequest -Uri "http://localhost:$($Config.NodePort)" -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        if ($response.StatusCode -eq 200) {
            $ready = $true
            break
        }
    } catch {
        # Not ready yet
    }
    Write-Host "    Waiting... ($elapsed/$($Config.StartupWaitSeconds) seconds)" -ForegroundColor Gray
    Start-Sleep -Seconds $interval
    $elapsed += $interval
}

if ($ready) {
    Write-Success "Juice Shop is responding on http://localhost:$($Config.NodePort)"
} else {
    Write-Failure "Juice Shop did not start within $($Config.StartupWaitSeconds) seconds"
    Write-Info "Check logs at: $($Config.LogsPath)"
    Write-Info "Try manually: Set-Location '$($Config.SiteRoot)'; node $entryPoint"
    exit 1
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
# STEP 5: CONFIGURE REVERSE PROXY RULES
# ============================================================================

Write-Step "STEP 5: Configuring Reverse Proxy Rules"

$webConfig = @"
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <rewrite>
            <rules>
                <rule name="ReverseProxyToJuiceShop" stopProcessing="true">
                    <match url="(.*)" />
                    <action type="Rewrite" url="http://localhost:$($Config.NodePort)/{R:1}" />
                </rule>
            </rules>
            <outboundRules>
                <rule name="RewriteLocationHeader" preCondition="IsRedirection">
                    <match serverVariable="RESPONSE_Location" pattern="http://localhost:$($Config.NodePort)/(.*)" />
                    <action type="Rewrite" value="https://{HTTP_HOST}/{R:1}" />
                </rule>
                <preConditions>
                    <preCondition name="IsRedirection">
                        <add input="{RESPONSE_STATUS}" pattern="3\d\d" />
                    </preCondition>
                </preConditions>
            </outboundRules>
        </rewrite>
        <proxy preserveHostHeader="true" />
        <webSocket enabled="true" />
    </system.webServer>
</configuration>
"@

Set-Content -Path "$($Config.SiteRoot)\web.config" -Value $webConfig -Encoding UTF8
Write-Success "Reverse proxy web.config written"

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
        # Add the HTTPS binding
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

    # Enable FREB for the site
    & $appcmd configure trace "$($Config.SiteName)" /enablesite 2>$null | Out-Null
    Write-Success "FREB enabled for $($Config.SiteName)"

    # Check if a tracing rule already exists
    $existingRules = & $appcmd list config "$($Config.SiteName)" /section:system.webServer/tracing/traceFailedRequests 2>$null

    if ($existingRules -match "path=") {
        Write-Info "FREB rules already configured — skipping"
    } else {
        # Add a tracing rule for error status codes
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
    $response = Invoke-WebRequest -Uri "https://$($Config.LocalHostname)" -SkipCertificateCheck -TimeoutSec 15
    Write-Success "HTTP Status: $($response.StatusCode)"
    Write-Success "Content Length: $($response.Content.Length) bytes"

    if ($response.Content -match "OWASP Juice Shop") {
        Write-Success "Juice Shop content confirmed!"
    } else {
        Write-Info "Got a response but 'OWASP Juice Shop' not found in content — check manually"
    }
} catch {
    Write-Failure "HTTPS test failed: $($_.Exception.Message)"
    Write-Info "Check that the site is running: & $appcmd list site $($Config.SiteName)"
    Write-Info "Check Juice Shop logs: $($Config.LogsPath)"
}

# Test real domain if configured
if ($Config.RealDomain) {
    Write-Info "Testing https://$($Config.RealDomain) (will only work if DNS resolves)..."
    try {
        $response2 = Invoke-WebRequest -Uri "https://$($Config.RealDomain)" -SkipCertificateCheck -TimeoutSec 10
        Write-Success "Real domain test: HTTP $($response2.StatusCode)"
    } catch {
        Write-Info "Real domain not reachable yet (expected if DNS not configured): $($_.Exception.Message)"
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n$("=" * 60)" -ForegroundColor Cyan
Write-Host "  JUICE SHOP INSTALLATION COMPLETE" -ForegroundColor Cyan
Write-Host "$("=" * 60)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Juice Shop Version:   $jsVersion" -ForegroundColor Gray
Write-Host "  Node.js Backend:      http://localhost:$($Config.NodePort)" -ForegroundColor Gray
Write-Host "  IIS HTTPS URL:        https://$($Config.LocalHostname)" -ForegroundColor Gray
if ($Config.RealDomain) {
    Write-Host "  WAF Domain URL:       https://$($Config.RealDomain)" -ForegroundColor Gray
}
Write-Host "  Windows Service:      $($Config.ServiceName) (Auto Start)" -ForegroundColor Gray
Write-Host "  Service Logs:         $($Config.LogsPath)" -ForegroundColor Gray
Write-Host "  FREB:                 $(if ($Config.FrebEnabled) {'Enabled'} else {'Disabled'})" -ForegroundColor Gray
Write-Host ""
Write-Host "  ⚠️  Juice Shop is INTENTIONALLY VULNERABLE." -ForegroundColor Red
Write-Host "     Restrict access via Azure NSG to Barracuda WaaS IPs only." -ForegroundColor Red
Write-Host ""
Write-Host "$("=" * 60)" -ForegroundColor Cyan
