#Requires -RunAsAdministrator
<#
.SYNOPSIS
    WimWizard-DefenderExclusions.ps1 - Add or remove Windows Defender exclusions for WimWizard

.DESCRIPTION
    Adds path and process exclusions to Windows Defender so that real-time protection
    does not interfere with WimWizard.ps1 during WIM servicing.

    Must be run BEFORE starting WimWizard.ps1 (without -Remove), and AFTER it finishes
    (with -Remove) to clean up the exclusions.

    Path exclusions are recursive - subfolders are automatically covered.

    Exclusions added:
      Paths:
        <OutputPath>\WIMServicing_Work\   (covers Mount, WinREMount, Scratch, LCU_temp*, winre*.wim)
        <OutputPath>\                      (covers the final .wim output file)
        <ScriptRoot>\Updates\              (covers downloaded MSU/CAB files)

      Processes:
        dism.exe
        dismhost.exe
        wusa.exe
        wimserv.exe

.PARAMETER OutputPath
    The output folder used by WimWizard.ps1.
    Default: the folder this script lives in + \Output

.PARAMETER Remove
    Remove the exclusions instead of adding them.

.EXAMPLE
    # Before running WimWizard:
    .\WimWizard-DefenderExclusions.ps1 -OutputPath "C:\wimwizard\Output"

    # After WimWizard finishes:
    .\WimWizard-DefenderExclusions.ps1 -OutputPath "C:\wimwizard\Output" -Remove

.NOTES
    Author  : Mathias Haas, Fidelity Consulting AB
    Contact : bWF0aGlhcy5oYWFzQGZpZGVsaXR5Y29uc3VsdGluZy5zZQ== (base64)
    License : GNU General Public License v3.0 (GPL-3.0)
              https://www.gnu.org/licenses/gpl-3.0.html
    Version : 1.0
    Product : WIM Wizard (tribute to WIM Witch by Donna Ryan)
    Requires: Windows PowerShell 5.1+, Administrator privileges

    CHANGELOG
    1.0     Initial release. Adds/removes Defender path and process exclusions
            for WimWizard.ps1 to prevent 0x80070241 / exit 577 during LCU application.
#>

param(
    [string]$OutputPath = "",
    [switch]$Remove
)

# ── Resolve paths ─────────────────────────────────────────────────────────────
$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) {
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

if (-not $OutputPath) {
    $OutputPath = Join-Path $ScriptRoot "Output"
}

# Normalize - remove trailing backslash if present
$OutputPath = $OutputPath.TrimEnd('\')

$workRoot    = Join-Path $OutputPath "WIMServicing_Work"
$updatesPath = Join-Path $ScriptRoot "Updates"

$pathExclusions = @(
    $workRoot,              # Recursive: covers Mount, WinREMount, Scratch\<GUID>, LCU_temp, LCU_temp2, LCU_winre_temp, winre*.wim
    $OutputPath,            # Covers final .wim output file
    $updatesPath,           # Covers downloaded MSU/CAB files read during patching
    $env:TEMP,              # wusa.exe extracts MSU contents here independently of DISM /ScratchDir
    "$env:SystemRoot\Temp"  # Same - wusa.exe may use either depending on process context
)

$processExclusions = @(
    "dism.exe",
    "dismhost.exe",
    "wusa.exe",
    "wimserv.exe"
)

# ── Output helpers ─────────────────────────────────────────────────────────────
function Write-OK   { param([string]$msg) Write-Host "  [OK]  $msg" -ForegroundColor Green }
function Write-Warn { param([string]$msg) Write-Host "  [!]   $msg" -ForegroundColor Yellow }
function Write-Info { param([string]$msg) Write-Host "  [i]   $msg" -ForegroundColor Cyan }
function Write-Fail { param([string]$msg) Write-Host "  [ERR] $msg" -ForegroundColor Red }

# ── Header ─────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "+==================================================================+" -ForegroundColor Cyan
Write-Host "|  WIM WIZARD - Defender Exclusions v1.0                           |" -ForegroundColor Cyan
Write-Host "+==================================================================+" -ForegroundColor Cyan
Write-Host ""

$action = if ($Remove) { "REMOVING" } else { "ADDING" }
Write-Host "  $action Windows Defender exclusions for WimWizard" -ForegroundColor White
Write-Host ""

# ── Check for Tamper Protection ────────────────────────────────────────────────
# Tamper Protection blocks Add-MpPreference even as admin. Warn if it's on.
# Note: Get-MpComputerStatus may not be available if Defender is managed by MDE/Sense.
try {
    $mpStatus = Get-MpComputerStatus -ErrorAction Stop
    if ($mpStatus.IsTamperProtected) {
        Write-Warn "Tamper Protection is ENABLED."
        Write-Warn "Exclusion changes may be blocked by Tamper Protection or MDE/Sense."
        Write-Warn "If the script fails, disable Tamper Protection first in Windows Security settings,"
        Write-Warn "or via Intune/MDE policy if managed."
        Write-Host ""
    }
} catch {
    Write-Warn "Could not read Defender status (possibly managed by MDE/Sense): $($_.Exception.Message)"
    Write-Host ""
}

# ── Apply exclusions ───────────────────────────────────────────────────────────
$anyFailed = $false

Write-Host "  Paths (recursive):" -ForegroundColor White
foreach ($p in $pathExclusions) {
    try {
        if ($Remove) {
            Remove-MpPreference -ExclusionPath $p -ErrorAction Stop
        } else {
            Add-MpPreference -ExclusionPath $p -ErrorAction Stop
        }
        Write-OK "$p"
    } catch {
        Write-Fail "$p"
        Write-Fail "  $($_.Exception.Message)"
        $anyFailed = $true
    }
}

Write-Host ""
Write-Host "  Processes:" -ForegroundColor White
foreach ($proc in $processExclusions) {
    try {
        if ($Remove) {
            Remove-MpPreference -ExclusionProcess $proc -ErrorAction Stop
        } else {
            Add-MpPreference -ExclusionProcess $proc -ErrorAction Stop
        }
        Write-OK "$proc"
    } catch {
        Write-Fail "$proc"
        Write-Fail "  $($_.Exception.Message)"
        $anyFailed = $true
    }
}

Write-Host ""

# ── Summary ────────────────────────────────────────────────────────────────────
if ($anyFailed) {
    Write-Warn "One or more exclusions could not be $($action.ToLower()). See errors above."
    Write-Warn "WimWizard may still encounter Defender interference during LCU application."
} else {
    if ($Remove) {
        Write-OK "All exclusions removed successfully."
    } else {
        Write-OK "All exclusions added successfully."
        Write-Host ""
        Write-Info "You can now run WimWizard.ps1."
        Write-Info "When WimWizard finishes, run this script again with -Remove to clean up:"
        Write-Host ""
        Write-Host "    .\WimWizard-DefenderExclusions.ps1 -OutputPath `"$OutputPath`" -Remove" -ForegroundColor Yellow
    }
}

Write-Host ""