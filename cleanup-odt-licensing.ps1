# Cleanup script for Office licensing (run elevated)
# - Prefers ospp.vbs in the same folder as this script (OfficeDeploymentTool)
# - Uninstalls installed product keys (ospp.vbs /unpkey:xxxxx)
# - Removes KMS host override registry values
# - Clears OfficeSoftwareProtectionPlatform ProgramData token cache
# - Shows final ospp.vbs /dstatusall

if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Script must be run as Administrator."
    exit 1
}

# Prefer the copy placed next to this script (OfficeDeploymentTool), then fall back to common locations
$scriptDir = $PSScriptRoot
if (-not $scriptDir) {
    # fallback when run in contexts where $PSScriptRoot is not set
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

$possible = @(
    Join-Path $scriptDir 'ospp.vbs',
    Join-Path $env:ProgramFiles 'Common Files\Microsoft Shared\OfficeSoftwareProtectionPlatform\ospp.vbs',
    Join-Path $env:ProgramFiles 'Microsoft Office\Office16\ospp.vbs',
    Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\ospp.vbs'
)

if ($env:'ProgramFiles(x86)') {
    $possible += @(
        Join-Path $env:'ProgramFiles(x86)' 'Microsoft Office\Office16\ospp.vbs',
        Join-Path $env:'ProgramFiles(x86)' 'Microsoft Office\root\Office16\ospp.vbs'
    )
}

$ospp = $possible | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $ospp) {
    Write-Error "ospp.vbs not found. Place your copy in the same folder as this script (`$scriptDir`) or update the script."
    exit 1
}
Write-Output "Using ospp.vbs at: $ospp"

function Run-OSPP([string]$args) {
    & cscript.exe //Nologo $ospp $args 2>&1
}

# 1) Query current status and collect last-5 product key suffixes
$status = Run-OSPP "/dstatusall" | Out-String
Write-Output "Current activation status collected."

# Regex: Last 5 characters of installed product key: XXXXX
$keys = @()
foreach ($m in [regex]::Matches($status, 'Last 5 characters of installed product key:\s*([A-Za-z0-9]{5})')) {
    $keys += $m.Groups[1].Value
}
$keys = $keys | Select-Object -Unique

if ($keys.Count -gt 0) {
    Write-Output "Found installed key suffixes: $($keys -join ', ')"
    foreach ($k in $keys) {
        Write-Output "Uninstalling key suffix $k ..."
        $out = Run-OSPP "/unpkey:$k" | Out-String
        Write-Output $out
    }
} else {
    Write-Output "No installed product-key suffixes detected to uninstall."
}

# 2) Remove KMS host override values from registry (HKLM)
$regPaths = @(
    'HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeSoftwareProtectionPlatform'
)

foreach ($rk in $regPaths) {
    if (Test-Path $rk) {
        Write-Output "Cleaning KMS override values under $rk"
        foreach ($val in @('KeyManagementServiceName','KeyManagementServicePort','KMSHost')) {
            try {
                Remove-ItemProperty -Path $rk -Name $val -ErrorAction SilentlyContinue
            } catch {
                # ignore
            }
        }
    } else {
        Write-Output "Registry path not present: $rk"
    }
}

# 3) Clear ProgramData token/cache (stop service if necessary)
$ospPath = 'C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform'
$svc = Get-Service -Name osppsvc -ErrorAction SilentlyContinue
$svcWasRunning = $false
if ($svc) {
    $svcWasRunning = $svc.Status -eq 'Running'
    if ($svcWasRunning) {
        Write-Output "Stopping osppsvc ..."
        Stop-Service -Name osppsvc -Force -ErrorAction SilentlyContinue
    }
}

if (Test-Path $ospPath) {
    Write-Output "Backing up ProgramData to $env:TEMP\OSPPL_Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
    $backup = Join-Path $env:TEMP ("OSPPL_Backup_{0}.zip" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    try {
        Compress-Archive -Path (Join-Path $ospPath '*') -DestinationPath $backup -Force -ErrorAction Stop
        Write-Output "Backup created: $backup"
    } catch {
        Write-Warning "Backup failed: $_"
    }

    Write-Output "Removing $ospPath ..."
    try {
        Remove-Item -Path $ospPath -Recurse -Force -ErrorAction Stop
        Write-Output "Removed $ospPath"
    } catch {
        Write-Warning "Failed to remove $ospPath: $_"
    }
} else {
    Write-Output "ProgramData path not present: $ospPath"
}

if ($svc -and $svcWasRunning) {
    Write-Output "Starting osppsvc ..."
    Start-Service -Name osppsvc -ErrorAction SilentlyContinue
}

# 4) Final verification
Write-Output "`n=== Final ospp.vbs /dstatusall ==="
Run-OSPP "/dstatusall"
Write-Output "=== End ==="