#Requires -RunAsAdministrator
<#
.SYNOPSIS
    WIM Wizard - Windows 11 Image Servicing Tool for SCCM/MECM

.DESCRIPTION
    Services a Windows 11 25H2/24H2 WIM file for distribution via SCCM/MECM.

    The user only needs to point to ONE folder containing the downloaded ISOs.
    Supports both x64 and ARM64 builds. The folder can contain all four ISOs
    at once; the correct ones are selected by architecture automatically.
    The script automatically:
      - Finds and mounts the correct Windows ISO (x64 or ARM64)
      - Finds and mounts the matching Language Pack ISO (also contains FOD packages)
        and searches the LanguagesAndOptionalFeatures subfolder for all cab files
      - Auto-selects the Enterprise edition (or asks if not found)
      - Downloads the latest Patch Tuesday updates from Microsoft Update Catalog
      - Injects language packs and FOD for chosen languages (default: se,no,dk,fi)
      - Removes unnecessary provisioned Appx packages
      - Exports and compresses the finished WIM

.PARAMETER SourceFolder
    Folder containing the ISOs downloaded from Microsoft.
    Can contain up to four ISOs (x64 Windows, ARM64 Windows, x64 LP, ARM64 LP).
    The correct ISO pair is selected automatically based on -X64 / -ARM64.
    Default: <ScriptFolder>\ISO-Source

.PARAMETER UpdatePath
    Optional manual override for update files (.msu/.cab).
    Leave empty to download updates automatically from Microsoft Update Catalog.
    Downloaded files are saved persistently to <ScriptFolder>\Updates\ and reused
    by KB number on subsequent runs - no re-downloading if the KB is unchanged.

.PARAMETER OutputPath
    Full path for the finished WIM file.
    Default: auto-generated from WIM version/build/languages/date

.PARAMETER WimIndex
    Force a specific WIM index. Default: 0 (auto-detect Enterprise).

.PARAMETER SkipUpdates
    Skip downloading and applying Patch Tuesday updates.

.PARAMETER SkipLanguagePacks
    Skip language pack injection.

.PARAMETER AppxListPath
    Path to an XML file containing a custom list of Appx package IDs to remove.
    Generate with WimWizard-GUI.ps1. If omitted, the built-in default list is used.

.PARAMETER SkipAppxRemoval
    Skip Appx package removal.

.PARAMETER PatchExistingWim
    Path to an existing serviced WIM file to patch with new updates only.
    Skips ISO discovery, language pack injection and Appx removal.
    Languages and OS version are read directly from the WIM.
    Use this for monthly patch cycles after the initial image build.

.PARAMETER Languages
    Comma-separated 2-letter country codes for language packs to inject.
    Example: "se,no,dk,fi"
    When omitted in interactive mode: prompts with default se,no,dk,fi.
    When omitted with -Unattended: skips language pack injection (English only).

.PARAMETER X64
    Build an x64 image. This is the default when neither -X64 nor -ARM64 is specified.
    Cannot be combined with -ARM64.

.PARAMETER ARM64
    Build an ARM64 image. Requires ARM64 Windows ISO and ARM64 LP ISO in the source folder.
    Cannot be combined with -X64.
    Note: Only NetFx3 is supported as a Feature on Demand for ARM64 builds.
    RSAT tools (RsatAD, RsatGPO, RsatSrvMgr) are not available offline for ARM64.

.PARAMETER FoDList
    Comma-separated Feature on Demand keys to enable in the image.
    Supported keys: NetFx3, RsatAD, RsatGPO, RsatSrvMgr
    Example: -FoDList "NetFx3,RsatAD"
    For ARM64 builds only NetFx3 is supported; other keys are silently skipped.
    FoD packages are sourced from the Language Pack ISO - no separate download needed.

.PARAMETER DebugBuild
    Enables extra diagnostics during the build:
      - Full DISM output printed to screen and log for both LCU passes
      - Full DISM output for component store cleanup
      - Pending package dump before /ResetBase (when no FoDs injected)
    Use when troubleshooting DISM failures or unexpected package states.

.PARAMETER Unattended
    Answers yes to all prompts and uses defaults for all path inputs.
    Suitable for scheduled tasks and automation pipelines.
    All paths (SourceFolder, OutputPath) must either be provided as parameters
    or resolve correctly from defaults based on $PSScriptRoot.
    If Enterprise edition cannot be auto-detected, the script will fail with
    an error rather than hang - use -WimIndex to specify the index explicitly.

    Version     : 5.1.8
    Date        : 2026-05-08
    Requires    : Windows PowerShell 5.1+, Administrator rights, DISM
    Tested on   : Windows 11 25H2 (OS build 26200.x), Windows Server 2022

    CHANGELOG
    5.1.8  Fix: Invoke-UpdateFolderCleanup no longer deletes old-named .msu/.cab
           files unless a canonical equivalent for the same KB already exists.
           Previously users upgrading from pre-5.1.6 with only an old-named cached
           LCU would have it deleted on the first run, forcing a re-download.
    5.1.7  New: Invoke-UpdateFolderCleanup runs automatically before each download
           pass. Removes: (1) old-named hash-less LCU/checkpoint .msu files left by
           older MSCatalogLTS code paths; (2) old-named SafeOS .cab files that were
           not cleaned up after canonical renaming; (3) superseded canonical files
           from previous months (0_Checkpoint / 1_LCU / 2_DotNet / 3_SafeOS) —
           only the newest KB per type/arch is kept. Existing users with stale files
           will have them cleaned automatically on the next run.
    5.1.6  Fix: $lcuFile selector now requires .msu extension. Previously any file
           matching "^windows11.0-kb<n>-" passed the filter, including old-named
           SafeOS .cab files (e.g. windows11.0-kb5083482-x64.cab). On runs where
           the SafeOS had been renamed to 3_SafeOS_... but the original remained,
           the .cab sorted before the real LCU .msu alphabetically and was selected
           as the LCU for both install.wim pass 1 and pass 2, leaving the WIM
           unpatched with no error.
           Fix: SafeOS rename now deletes the old-named original when the canonical
           3_SafeOS_... file already exists (from a prior run), preventing stale
           duplicates from accumulating in the Updates folder.
           Fix: Checkpoint download now removes any hash-less copy of kb5043080
           after downloading the hash-named version, matching the LCU stale-file
           cleanup behaviour.
           Fix: RunOnce language fix injection now skipped in patch mode and when
           -SkipAppxRemoval is set ($SkipAppxRemoval is forced true in patch mode).
           The block was unconditional, so patch runs always overwrote the RunOnce
           key and InstallSystemApps.ps1 baked in at full-build time.
    5.1.5  Fix: LP injection (Step 10) now handles 0x800f0952 gracefully. LTSC
           images (e.g. English International) ship with en-GB pre-installed;
           Add-WindowsPackage previously had no try/catch, so any error was
           fatal. Added try/catch with 0x800f0952 treated as ignorable (language
           already present in base image) and 0x800f081e also ignorable (not
           applicable). All other errors still rethrow and abort the pipeline.
    5.1.4  Fix: Windows ISO detection pattern extended to match LTSC 2024 ISOs.
           Previous patterns (win_pro_11, win.*11.*ent, win.*11.*business) did not
           match SW_DVD9_WIN_ENT_LTSC_2024_* because the filename contains neither
           "11" nor "ent" adjacent to "11". Added win.*ltsc to the pattern.
    5.1.3  Fix: Architecture filter regex corrected for update file selection.
           Patterns "_x64\." / "_arm64\." never matched real Microsoft LCU
           filenames (format: windows11.0-kb...-arm64_<hash>.msu). On an x64
           build with both arch files present the ARM64 LCU passed the filter and
           was selected first (alphabetical sort). Fixed to "-x64[_.]" and
           "-arm64[_.]" which correctly match Microsoft's dash-before/underscore-
           after separator convention.
           Fix: LCU failures now fatal (exit 1) instead of silently continuing.
           Previously, Add-WindowsPackage/dism.exe errors on WinRE SSU, LCU pass 1,
           and LCU pass 2 were caught and logged as warnings but did not abort the
           pipeline. The script would produce a WIM with an unpatched image and
           report success. Affected error codes include 0x80070241 and any other
           non-ignorable DISM failure. The two known-ignorable codes (0x8007007e
           and 0x800f081e) remain non-fatal as before.
    5.1.2  Fix: Appx removal count in pre-run summary now reflects the custom XML list
           count instead of always showing the default list count (27).
    5.1.1  Code cleanup pass - no functional pipeline changes.
           - Removed [Fk4] fork/branch labels from comment and Write-Section output
           - LCU DownloadDialog API fallback changed from silent Save-MSCatalogUpdate
             (which strips the hash from the filename, causing 0x80070241) to a hard
             fail with clear instructions to download the LCU manually
           - LCU cache hash-detection regex tightened: _[0-9a-f]{8,} instead of
             _[0-9a-f]{36,} (safer match against real Microsoft filenames)
           - Download success counter no longer increments on fallback path
           - Manual update instructions updated: LCU must be saved with its original
             Microsoft filename (windows11.0-kbXXXXXXX-x64_<hash>.msu) - the old
             1_LCU_ prefix naming is not recognised by the detection logic
           - Added comment explaining $matches2 naming (avoids $Matches collision)
           - Fixed stray 4-space indent on $CatalogSearchTerms assignment (line ~1223)
           - Section separator aligned to 66 chars to match banner width

    5.1.0  Complete pipeline rewrite per Microsoft documented offline servicing
           sequence (learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update).
           Resolves offline LCU failure (0x800401e3/0x80070241) on fully patched
           Windows 11 25H2 hosts.

           Key changes:
           - WinRE patched BEFORE install.wim (Microsofts dokumenterade ordning)
           - LCU applied twice: pass 1 (SSU) before LP/FoD, pass 2 (full) after
           - dism.exe /Add-Package replaces Add-WindowsPackage for LCU on main OS
             (Add-WindowsPackage fails with 0x800401e3 for KB5083769 Apr 2026)
           - LCU downloaded via DownloadDialog API preserving original filename
             with SHA1 hash (e.g. windows11.0-kb5083769-x64_57f4bd47...msu).
             DISM validates filename against package signature - MSCatalogLTS
             strips the hash causing 0x80070241/0x80070570.
           - Start-BitsTransfer replaces Invoke-WebRequest for LCU/checkpoint
             downloads (progress display, faster, no memory buffering)
           - Unblock-File called after download to remove Zone.Identifier MoTW
           - Language FoDs injected via Add-WindowsCapability with capability
             names (Language.Basic~~~sv-SE~0.0.1.0) instead of raw cab files
           - NetFx3 installed via Add-WindowsCapability after cleanup (step 15)
           - No /ResetBase on main OS cleanup (Microsoft spec: WinRE/WinPE only)
           - AppxListPath XML parsing fixed: block moved after Write-Log definition
             to avoid Set-StrictMode crash when script starts
           - BuildFilter added to CatalogSearchTerms to prevent 24H2/25H2 mix-up
             when they share same KB number
           - Removed -Architecture param from Get-MSCatalogUpdate (filtered wrong)

    5.0 beta3 -DebugBuild now writes full DISM output and pending-package dumps
             to the log file. -X64 and -ARM64 made mutually exclusive.
    4.9.x  Multiple fixes: ARM64 support, FoD tab in GUI, NuGet auto-install,
           dotNetFiles StrictMode fix, log file location, component cleanup order.
    1.0.0-4.8.x  See git history for full changelog.

.LINK
    https://admin.microsoft.com/adminportal/home#/subscriptions/vlnew/downloadsandkeys (Enterprise ISO + LP + FOD)
    https://catalog.update.microsoft.com (monthly updates)
#>

[CmdletBinding()]
param(
    [switch]$Help,
    [string]$SourceFolder = "",             # Defaults to <ScriptRoot>\ISO-Source (set after param)
    [string]$UpdatePath,                    # Manual override; empty = auto-download
    [string]$OutputPath,
    [int]   $WimIndex            = 0,       # 0 = auto-detect Enterprise
    [string]$Languages          = "",       # Comma-separated 2-letter country codes, e.g. "se,no,dk,fi"
                                            # Defaults to se,no,dk,fi when not -Unattended.
                                            # When -Unattended with no -Languages: skip LPs (English only).
    [switch]$SkipUpdates,
    [switch]$SkipLanguagePacks,
    [string]$AppxListPath = "",             # Path to XML file with custom Appx removal list
                                            # Generated by WimWizard-GUI.ps1. If not specified,
                                            # uses built-in default list.
    [switch]$SkipAppxRemoval,
    [string]$PatchExistingWim = "",         # Path to existing serviced WIM to patch (updates only).
                                            # When set: skips ISO discovery, LP injection, Appx removal.
                                            # Languages and build info are read from the WIM itself.
    [switch]$X64,                         # Build x64 image (default when neither switch set)
    [switch]$ARM64,                        # Build ARM64 image
    [string]$FoDList = "",                 # Comma-separated FoD keys to enable (e.g. "NetFx3,RsatAD")
    [switch]$Unattended,
    [switch]$DebugBuild                    # Extra diagnostics: show full DISM output, dump pending packages before cleanup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ScriptVersion = "5.1.8"

# Validate architecture selection
if ($X64 -and $ARM64) {
    Write-Host ""
    Write-Host "  [ERR] -X64 and -ARM64 cannot be used together. Please specify only one architecture." -ForegroundColor Red
    Write-Host ""
    exit 1
}
$Architecture = if ($ARM64) { "arm64" } else { "x64" }

# ARM64 only supports NetFx3 - strip unsupported FoD keys and warn
if ($ARM64 -and $FoDList -ne "") {
    $allowedArm64FoDs = @("NetFx3")
    $requestedKeys = $FoDList -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    $unsupported = $requestedKeys | Where-Object { $_ -notin $allowedArm64FoDs }
    if ($unsupported) {
        Write-Host ""
        Write-Host "  [!]   ARM64 only supports NetFx3 as a Feature on Demand." -ForegroundColor Yellow
        Write-Host "  [!]   The following FoD keys are not available for ARM64 and will be skipped: $($unsupported -join ', ')" -ForegroundColor Yellow
        Write-Host ""
    }
    $FoDList = ($requestedKeys | Where-Object { $_ -in $allowedArm64FoDs }) -join ","
}

if ($Help) {
    Write-Host ""
    Write-Host "  WIM Wizard v$ScriptVersion" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  USAGE" -ForegroundColor White
    Write-Host "  -----"
    Write-Host "  Full build (interactive):"
    Write-Host "    .\WimWizard.ps1"
    Write-Host ""
    Write-Host "  Full build (unattended, 4 Nordic languages):"
    Write-Host "    .\WimWizard.ps1 -Languages `"da,fi,no,se`" -Unattended"
    Write-Host ""
    Write-Host "  Full build (ARM64, unattended):"
    Write-Host "    .\WimWizard.ps1 -ARM64 -Languages `"da,fi,no,se`" -Unattended"
    Write-Host ""
    Write-Host "  Full build with Features on Demand:"
    Write-Host "    .\WimWizard.ps1 -Languages `"da,fi,no,se`" -FoDList `"NetFx3,RsatAD`" -Unattended"
    Write-Host ""
    Write-Host "  Patch existing WIM (updates only):"
    Write-Host "    .\WimWizard.ps1 -PatchExistingWim `"Output\Win11_25H2_..._20260409.wim`""
    Write-Host ""
    Write-Host "  Build using manually downloaded updates:"
    Write-Host "    .\WimWizard.ps1 -Languages `"da,fi,no,se`" -UpdatePath `"C:\Updates\`" -Unattended"
    Write-Host ""
    Write-Host "  PARAMETERS" -ForegroundColor White
    Write-Host "  ----------"
    Write-Host "  -SourceFolder     <path>   Folder with Windows ISO + LP ISO (default: .\ISO-Source\)"
    Write-Host "  -Languages        <codes>  Comma-separated language codes: da,fi,no,se,de,fr ..."
    Write-Host "                             Omit for English-only build. Interactive mode prompts if not set."
    Write-Host "  -OutputPath       <path>   Output WIM path (auto-generated from version/langs/date if omitted)"
    Write-Host "  -UpdatePath       <path>   Folder with pre-downloaded .msu/.cab files (skips auto-download)"
    Write-Host "  -PatchExistingWim <path>   Patch this WIM with latest updates only (skips LP/Appx steps)"
    Write-Host "  -AppxListPath     <path>   XML app removal list (generated by GUI; uses built-in list if omitted)"
    Write-Host "  -WimIndex         <int>    WIM index to service (default: auto-detect Enterprise)"
    Write-Host "  -X64                       Build x64 image (default when neither -X64 nor -ARM64 is set)"
    Write-Host "  -ARM64                     Build ARM64 image (mutually exclusive with -X64)"
    Write-Host "  -FoDList          <keys>   Features on Demand: NetFx3,RsatAD,RsatGPO,RsatSrvMgr"
    Write-Host "                             ARM64 note: only NetFx3 is supported; RSAT keys are silently skipped"
    Write-Host "  -SkipUpdates               Do not download or apply updates"
    Write-Host "  -SkipLanguagePacks         Skip language pack and FOD injection"
    Write-Host "  -SkipAppxRemoval           Skip removal of provisioned Appx packages"
    Write-Host "  -DebugBuild                Extra diagnostics: full DISM output for all steps + pending package dump"
    Write-Host "  -Unattended                No interactive prompts (for GUI/automation use)"
    Write-Host "  -Help                      Show this help"
    Write-Host ""
    exit 0
}

# Derive base folder from wherever this script lives.
# All default paths (ISO-Source, Updates, Output) are relative to this.
$ScriptRoot = $PSScriptRoot
if (-not $ScriptRoot) {
    # Fallback if run dot-sourced or from ISE without a saved path
    $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
Write-Verbose "Script root: $ScriptRoot"

# Apply defaults that depend on $ScriptRoot (cannot be set in param() block)
if (-not $SourceFolder) { $SourceFolder = "$ScriptRoot\ISO-Source" }
if (-not $OutputPath)   { $OutputPath   = "$ScriptRoot\Output\install.wim" }

# Full locale tag table - maps 2-letter country code to all matching locale tags in the LP ISO.
# Multiple matches (e.g. "in" -> bn-IN, en-IN, hi-IN, ta-IN) are all installed.
# Norwegian: user passes "no" which maps to nb-NO (Bokmål) as that is what the LP ISO contains.
$LocaleMap = @{
    # Full language packs (Microsoft-Windows-Client-Language-Pack_x64_xx-xx.cab)
    'ar' = @('ar-SA')
    'bg' = @('bg-BG')
    'cs' = @('cs-CZ')
    'da' = @('da-DK')
    'de' = @('de-DE')
    'dk' = @('da-DK')   # alias: country code for Denmark
    'el' = @('el-GR')
    'en' = @('en-GB','en-US')
    'es' = @('es-ES','es-MX')
    'et' = @('et-EE')
    'fi' = @('fi-FI')
    'fr' = @('fr-CA','fr-FR')
    'gb' = @('en-GB')
    'gr' = @('el-GR')
    'he' = @('he-IL')
    'hr' = @('hr-HR')
    'hu' = @('hu-HU')
    'il' = @('he-IL')
    'it' = @('it-IT')
    'ja' = @('ja-JP')   # Japanese (ISO 639-1)
    'jp' = @('ja-JP')   # Japanese (country code alias)
    'ko' = @('ko-KR')
    'kr' = @('ko-KR')
    'lt' = @('lt-LT')
    'lv' = @('lv-LV')
    'mx' = @('es-MX')
    'nl' = @('nl-NL')
    'no' = @('nb-NO')   # nb-NO is the LP ISO locale for Norwegian
    'pl' = @('pl-PL')
    'pt' = @('pt-BR','pt-PT')
    'ro' = @('ro-RO')
    'ru' = @('ru-RU')
    'se' = @('sv-SE')
    'sk' = @('sk-SK')
    'sl' = @('sl-SI')
    'sr' = @('sr-Latn-RS')
    'sv' = @('sv-SE')   # alias: language code for Swedish
    'th' = @('th-TH')
    'tr' = @('tr-TR')
    'tw' = @('zh-TW')
    'hk' = @('zh-TW')   # zh-HK: uses zh-TW LP as display language base
    'ua' = @('uk-UA')
    'uk' = @('uk-UA')
    'us' = @('en-US')
    'vn' = @('vi-VN')
    'zh' = @('zh-CN','zh-TW')
    'cn' = @('zh-CN')
    # Language Interface Packs (Microsoft-Windows-Lip-Language-Pack_x64_xx-xx.cab)
    'ca' = @('ca-ES')   # Catalan
    'eu' = @('eu-ES')   # Basque
    'gl' = @('gl-ES')   # Galician
    'id' = @('id-ID')   # Indonesian
    'vi' = @('vi-VN')   # Vietnamese
}

# Build the one-line supported list for display
$SupportedCodesLine = ($LocaleMap.Keys | Sort-Object) -join ", "

# -- Output helpers (needed before any validation runs) -----------------------
function Write-OK   { param([string]$msg) Write-Host "  [OK]  $msg" -ForegroundColor Green }
function Write-Warn { param([string]$msg) Write-Host "  [!]   $msg" -ForegroundColor Yellow }
function Write-Info { param([string]$msg) Write-Host "  [i]   $msg" -ForegroundColor Cyan }
function Write-Fail { param([string]$msg) Write-Host "  [ERR] $msg" -ForegroundColor Red }

function Resolve-LanguageCodes {
    param([string]$CodeString)
    $codes  = $CodeString.ToLower().Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    $result = [System.Collections.Generic.List[string]]::new()
    $errors = @()
    foreach ($code in $codes) {
        if ($LocaleMap.ContainsKey($code)) {
            foreach ($tag in $LocaleMap[$code]) { if (-not $result.Contains($tag)) { $result.Add($tag) } }
        } else {
            $errors += $code
        }
    }
    if ($errors.Count -gt 0) {
        Write-Fail "Unknown language code(s): $($errors -join ', ')"
        Write-Host ""
        Write-Host "  Supported codes: $SupportedCodesLine" -ForegroundColor Yellow
        Write-Host ""
        Dismount-AllISOs; exit 1
    }
    return @($result)
}

# Resolve $ResolvedLocales - the list of full locale tags to install (e.g. @("sv-SE","nb-NO","da-DK","fi-FI"))
$ResolvedLocales = @()

# ── Patch mode: read existing WIM metadata ────────────────────────────────────
$PatchMode = $PatchExistingWim -ne "" -and (Test-Path $PatchExistingWim)

if ($PatchExistingWim -ne "" -and -not (Test-Path $PatchExistingWim)) {
    Write-Host ""
    Write-Fail "WIM file not found: $PatchExistingWim"
    Write-Host ""
    Write-Host "  Usage: .\WimWizard.ps1 -PatchExistingWim `"Output\Win11_25H2_..._20260409.wim`"" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

if ($PatchMode) {
    Write-Info "Patch mode: reading metadata from existing WIM..."
    try {
        $patchWimInfo = Get-WindowsImage -ImagePath $PatchExistingWim -Index 1 -ErrorAction Stop
        # Use index 1 since serviced WIMs exported by WimWizard have a single index
        Write-OK "WIM: $($patchWimInfo.ImageName) ($($patchWimInfo.Version))"
        Write-OK "Languages in WIM: $($patchWimInfo.Languages -join ', ')"
    } catch {
        Write-Fail "Cannot read WIM metadata from: $PatchExistingWim"
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    # Force-skip LP injection and Appx removal in patch mode
    $SkipLanguagePacks = $true
    $SkipAppxRemoval   = $true
}

# Trap ensures ISOs are always dismounted even if the script exits early (Ctrl+C, error, user cancel)
trap {
    Write-Host ""
    Write-Fail "Unhandled error: $($_.Exception.Message)"
    Dismount-AllISOs
    break
}

function Exit-Script {
    param([int]$Code = 0)
    Dismount-AllISOs
    exit $Code
}

$CatalogSearchTerms = $null   # populated after WIM index selection (see Get-CatalogSearchTerms)

function Get-CatalogSearchTerms {
    param([string]$WimVersion)   # e.g. "10.0.26200.6584"

    # Extract the 5-digit build number (third segment)
    $build = 0
    if ($WimVersion -match '^\d+\.\d+\.(\d+)\.') {
        $build = [int]$Matches[1]
    }

    # Map build number to Windows version string as used in the Update Catalog
    # Microsoft uses "Version XX" (capital V) in LCU/SafeOS titles and
    # "version XX" (lowercase v) in .NET titles - match exactly.
    $versionString = switch ($true) {
        ($build -ge 27000)              { "26H2" }   # placeholder for future
        ($build -ge 26200 -and $build -lt 26300) { "25H2" }
        ($build -ge 26100 -and $build -lt 26200) { "24H2" }
        ($build -ge 22631 -and $build -lt 26100) { "23H2" }
        default {
            Write-Warn "Unknown build number $build - defaulting to 25H2 search terms"
            "25H2"
        }
    }

    Write-Info "Detected Windows 11 version: $versionString (build $build)"

    # IMPORTANT: SafeOS titles in the Catalog combine 24H2 and 25H2 in one entry
    # e.g. "Safe OS Dynamic Update for Windows 11, versions 24H2 and 25H2"
    # LCU and .NET have separate entries per version.
    $safeOSSearch = if ($versionString -in @("24H2","25H2")) {
        "Safe OS Dynamic Update for Windows 11, versions 24H2 and 25H2"
    } else {
        "Safe OS Dynamic Update for Windows 11 Version $versionString"
    }

    # ARM64 uses different catalog titles than x64
    $archLabel = if ($Architecture -eq "arm64") { "arm64" } else { "x64" }

    # Build number filter prevents 24H2/25H2 mix-up when they share the same KB number.
    $buildFilter = $build.ToString()   # e.g. "26200" or "26100"

    return @{
        LCU           = "Cumulative Update for Windows 11 Version $versionString for $archLabel"
        DotNet        = "Cumulative Update for .NET Framework 3.5 and 4.8.1 for Windows 11, version $versionString for $archLabel"
        SafeOS        = $safeOSSearch
        SafeOSVersion = $versionString
        BuildFilter   = $buildFilter
    }
}

#region ── UI helper functions ────────────────────────────────────────────────

function Write-Banner {
    Clear-Host
    $w = "=" * 66
    Write-Host "+$w+" -ForegroundColor Cyan
    Write-Host ("|  WIM WIZARD  v{0,-55}|" -f $ScriptVersion) -ForegroundColor Cyan
    Write-Host "|  A tribute to WIM Witch by Donna Ryan                             |" -ForegroundColor Cyan
    Write-Host "+$w+" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("-" * 66) -ForegroundColor DarkCyan
    Write-Host "  >> $Title" -ForegroundColor White
    Write-Host ("-" * 66) -ForegroundColor DarkCyan
}

function Confirm-Continue {
    param([string]$Message = "Continue?")
    if ($Unattended) { return $true }
    Write-Host ""
    $answer = Read-Host "  $Message [Y/n]"
    return ($answer -eq "" -or $answer -match "^[yY]")
}

function Read-PathWithDefault {
    param([string]$Prompt, [string]$Default, [switch]$MustExist)
    if ($Unattended) {
        Write-Info "$Prompt"
        Write-Info "  [Unattended] Using default: $Default"
        $path = $Default
    } else {
        Write-Host "  $Prompt" -ForegroundColor White
        Write-Host "  [Default: $Default]" -ForegroundColor DarkGray
        $inp  = Read-Host "  Path (press Enter for default)"
        $path = if ($inp.Trim() -eq "") { $Default } else { $inp.Trim() }
    }
    if ($MustExist -and -not (Test-Path $path)) {
        Write-Fail "Path not found: $path"
        return $null
    }
    return $path
}

#endregion

#region ── Appx packages to remove ───────────────────────────────────────────

# Default list used when no -AppxListPath is provided
$AppxToRemoveDefault = @(
    # Verified against Microsoft 25H2 provisioned app policy. Matches GUI defaults.
    "Microsoft.BingSearch"                          # Bing Search
    "Clipchamp.Clipchamp"                           # Clipchamp
    "Microsoft.Copilot"                             # Copilot (Consumer)
    "DevHome_8wekyb3d8bbwe"                         # Dev Home
    "MicrosoftCorporationII.MicrosoftFamily"        # Family Safety
    "Microsoft.WindowsFeedbackHub"                  # Feedback Hub
    "Microsoft.GamingApp"                           # Gaming App (Xbox)
    "Microsoft.GetHelp"                             # Get Help
    "Microsoft.MicrosoftJournal"                    # Journal
    "Microsoft.Messaging"                           # Messaging
    "Microsoft.BingWeather"                         # MSN Weather
    "Microsoft.BingNews"                            # Microsoft News
    "Microsoft.MicrosoftOfficeHub"                  # Microsoft Office Hub
    "Microsoft.MicrosoftPCManager"                  # Microsoft PC Manager
    "Microsoft.MicrosoftSolitaireCollection"        # Microsoft Solitaire
    "Microsoft.ZuneMusic"                           # Microsoft Store - Music
    "Microsoft.ZuneVideo"                           # Microsoft Store - Video
    "MicrosoftTeams"                                # Microsoft Teams (Personal)
    "Microsoft.OutlookForWindows"                   # Outlook for Windows
    "MicrosoftCorporationII.QuickAssist"            # Quick Assist
    "Microsoft.Whiteboard"                          # Whiteboard
    "Microsoft.XboxGamingOverlay"                   # Xbox Game Bar
    "Microsoft.XboxIdentityProvider"                # Xbox Identity Provider
    "Microsoft.XboxSpeechToTextOverlay"             # Xbox Speech to Text
    "Microsoft.Xbox.TCUI"                           # Xbox TCUI
    "Microsoft.YourPhone"                           # Your Phone (Phone Link)
    "Microsoft.QuickAssist"                      # Quick Assist (legacy package name)
)
# $AppxToRemove is populated after Write-Log is defined (see below)
# Initialize to default here so Summary block can reference it safely
$AppxToRemove = $AppxToRemoveDefault


#endregion

#region ── ISO auto-discovery and mounting ────────────────────────────────────

# Tracks all ISOs we have mounted so we can dismount them all on exit/error
$script:MountedISOs = @()

function Mount-ISOSafe {
    <#
    .SYNOPSIS
        Mounts an ISO and returns the drive letter. Tracks it for cleanup.
    #>
    param([string]$IsoPath)
    $result = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $letter = ($result | Get-Volume).DriveLetter + ":"
    $script:MountedISOs += $IsoPath
    return $letter
}

function Dismount-AllISOs {
    foreach ($iso in $script:MountedISOs) {
        try { Dismount-DiskImage -ImagePath $iso -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    $script:MountedISOs = @()
}

function Resolve-SourceFolder {
    param([string]$Folder, [switch]$SkipLP)

    $isoFiles = @(Get-ChildItem $Folder -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match "^\.iso$" })

    if (-not $isoFiles -or $isoFiles.Count -eq 0) {
        return $null
    }

    Write-Info "Found $($isoFiles.Count) ISO file(s) in $($Folder):"
    $isoFiles | ForEach-Object { Write-Host "    $($_.Name)" -ForegroundColor DarkGray }

    $result = @{ WimFile = $null; WimISO = $null; LPSearchRoot = $null; FODSearchRoot = $null }

    foreach ($iso in $isoFiles) {
        $name = $iso.Name.ToLower()

        # LP checked first - its filename pattern overlaps with the OS ISO pattern below.
        # LP ISO:  SW_DVD9_Win_11_*_LangPack_*.ISO  (also contains all FOD packages)
        # OS ISO:  SW_DVD9_Win_Pro_11_25H2_*.ISO / en-us_windows_11_business_*.iso
        $isLPISO = (
            $name -match "langpack"       -or  # SW_DVD9_Win_11_*_LangPack_*
            $name -match "language.?pack" -or  # mu_windows_11_language_pack_*
            $name -match "lang_pack"      -or
            $name -match "_lp_"
        )

        # Architecture filter for LP ISO: prefer the one matching requested arch.
        # ARM64 LP ISO has "arm64" in filename; x64 LP ISO does not.
        if ($isLPISO) {
            $lpIsArm64 = $name -match "arm64"
            $wantArm64 = $Architecture -eq "arm64"
            if ($lpIsArm64 -ne $wantArm64) { $isLPISO = $false }
        }

        # ── Windows OS ISO ────────────────────────────────────────────────────
        # Explicitly excluded: anything already matched as LP above
        $isWindowsISO = (
            -not $isLPISO
        ) -and (
            $name -match "win_pro_11"        -or  # SW_DVD9_Win_Pro_11_25H2_* / _26H2_*
            $name -match "win.*11.*ent"      -or  # en-us_windows_11_enterprise_*
            $name -match "win.*11.*business" -or  # en-us_windows_11_business_*
            $name -match "win.*ltsc"              # SW_DVD9_WIN_ENT_LTSC_2024_*
        )

        # Architecture filter: if -Architecture is set, only pick the matching ISO
        if ($isWindowsISO) {
            $isArm64 = $name -match "arm64"
            $wantArm64 = $Architecture -eq "arm64"
            if ($isArm64 -ne $wantArm64) { $isWindowsISO = $false }
        }

        if ($isWindowsISO -and -not $result.WimFile) {
            Write-Info "Mounting Windows ISO: $($iso.Name)"
            try {
                $drive   = Mount-ISOSafe -IsoPath $iso.FullName
                $wimPath = "$drive\sources\install.wim"
                if (-not (Test-Path $wimPath)) { $wimPath = "$drive\sources\install.esd" }
                if (Test-Path $wimPath) {
                    $result.WimFile = $wimPath
                    $result.WimISO  = $iso.FullName
                    Write-OK "Windows ISO mounted as ${drive} - $wimPath"
                } else {
                    Write-Warn "Mounted $($iso.Name) but no install.wim/esd found - skipping"
                }
            } catch {
                Write-Warn "Could not mount $($iso.Name): $_"
            }

        } elseif ($isLPISO -and -not $result.LPSearchRoot) {
            if ($SkipLP) {
                Write-Info "Skipping Language Pack ISO (language packs disabled)"
            } else {
                Write-Info "Mounting Language Pack ISO: $($iso.Name)"
                try {
                    $drive = Mount-ISOSafe -IsoPath $iso.FullName
                    # LP and FOD cabs are both under LanguagesAndOptionalFeatures\
                    $lpFodSubfolder = "$drive\LanguagesAndOptionalFeatures"
                    if (Test-Path $lpFodSubfolder) {
                        $result.LPSearchRoot  = $lpFodSubfolder
                        $result.FODSearchRoot = $lpFodSubfolder
                        Write-OK "LP ISO mounted as ${drive}, using subfolder: LanguagesAndOptionalFeatures"
                    } else {
                        # Subfolder not found - search whole ISO root
                        $result.LPSearchRoot  = $drive
                        $result.FODSearchRoot = $drive
                        Write-OK "LP ISO mounted as $($drive) (no LanguagesAndOptionalFeatures subfolder, searching root)"
                    }
                } catch {
                    Write-Warn "Could not mount $($iso.Name): $_"
                }
            }
        }
    }

    return $result
}

#endregion

#region ── Automatic download via MSCatalogLTS module ────────────────────────

function Install-MSCatalogLTS {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if (Get-Module -Name MSCatalogLTS -ListAvailable) {
        Import-Module MSCatalogLTS -Force
        Write-OK "MSCatalogLTS module loaded"
        return $true
    }

    Write-Info "MSCatalogLTS module not found - installing from PSGallery..."

    # Ensure NuGet provider DLL is on disk before any PackageManagement cmdlet runs.
    # Get-PackageProvider, Install-PackageProvider, and Install-Module all trigger
    # the same interactive prompt on a clean system. Avoid all of them by checking
    # for the DLL directly with Test-Path, and downloading it with Invoke-WebRequest
    # if missing. Only after the file exists do we touch PackageManagement at all.
    $nugetDir  = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget\2.8.5.208"
    $nugetDest = "$nugetDir\Microsoft.PackageManagement.NuGetProvider.dll"
    if (-not (Test-Path $nugetDest)) {
        Write-Info "Installing NuGet provider (required by PowerShellGet)..."
        New-Item $nugetDir -ItemType Directory -Force | Out-Null
        Invoke-WebRequest -Uri 'https://cdn.oneget.org/providers/Microsoft.PackageManagement.NuGetProvider-2.8.5.208.dll' `
                          -OutFile $nugetDest -UseBasicParsing
    }
    Import-PackageProvider -Name NuGet -RequiredVersion 2.8.5.208 -Force | Out-Null

    try {
        Install-Module -Name MSCatalogLTS -Force -Scope AllUsers -AllowClobber
        Import-Module MSCatalogLTS -Force
        Write-OK "MSCatalogLTS installed and loaded"
        return $true
    } catch {
        Write-Warn "Could not install MSCatalogLTS: $($_.Exception.Message)"
        return $false
    }
}

function Invoke-UpdateFolderCleanup {
    # Removes legacy / stale files from the Updates folder that accumulate over time.
    # Called automatically at the start of every download pass.
    #
    # Removes:
    #   1. Old-named LCU / checkpoint .msu files — original Microsoft filenames without
    #      the SHA1 hash segment (e.g. windows11.0-kb5083769-x64.msu). These were
    #      produced by older MSCatalogLTS code paths and cause DISM signature failures.
    #   2. Old-named SafeOS .cab files — original Microsoft filenames (e.g.
    #      windows11.0-kb5083482-x64.cab) that remained after canonical renaming to
    #      3_SafeOS_KB... and could be mistakenly selected as the LCU.
    #   3. Superseded canonical files — all 0_Checkpoint / 1_LCU / 2_DotNet / 3_SafeOS
    #      files from previous months that are no longer the current KB. Only the
    #      newest KB for each prefix+arch combination is kept.
    param([string]$DownloadDir)

    if (-not (Test-Path $DownloadDir)) { return }

    $cleaned = 0

    # ── 1. Old-named LCU/checkpoint .msu (windows11.0-kb<n>-<arch>.msu, no hash) ──
    # Only remove if a canonical file for the same KB already exists in the folder.
    # This protects users upgrading from pre-5.1.6 whose only cached copy is the
    # old-named file — we leave it in place; Invoke-AutoUpdateDownload will rename it.
    Get-ChildItem $DownloadDir -Filter "*.msu" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "^windows11\.0-kb\d+-" -and $_.Name -notmatch '_[0-9a-f]{8,}' } |
        ForEach-Object {
            $oldFile = $_
            $kb = ([regex]::Match($oldFile.Name, 'kb(\d+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Groups[1].Value
            # Check whether a canonical file for this KB already exists
            $hasCanonical = [bool](Get-ChildItem $DownloadDir -Filter "*_KB${kb}*" -ErrorAction SilentlyContinue |
                                   Where-Object { $_.Name -match '^[0-9]_' })
            if ($hasCanonical) {
                Write-Info "Cleanup: removing old-named MSU (canonical exists): $($oldFile.Name)"
                Remove-Item $oldFile.FullName -Force -ErrorAction SilentlyContinue
                $cleaned++
            }
        }

    # ── 2. Old-named SafeOS .cab (windows11.0-kb<n>-<arch>.cab) ──────────────────
    # Same safety check: only remove if the canonical 3_SafeOS_KB... already exists.
    Get-ChildItem $DownloadDir -Filter "*.cab" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match "^windows11\.0-kb\d+-" } |
        ForEach-Object {
            $oldFile = $_
            $kb = ([regex]::Match($oldFile.Name, 'kb(\d+)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Groups[1].Value
            $hasCanonical = [bool](Get-ChildItem $DownloadDir -Filter "3_SafeOS_KB${kb}*" -ErrorAction SilentlyContinue)
            if ($hasCanonical) {
                Write-Info "Cleanup: removing old-named CAB (canonical exists): $($oldFile.Name)"
                Remove-Item $oldFile.FullName -Force -ErrorAction SilentlyContinue
                $cleaned++
            }
        }

    # ── 3. Superseded canonical files ────────────────────────────────────────────
    # Group by prefix (e.g. "1_LCU") and architecture suffix, keep only the newest
    # KB in each group. This handles both .msu and .cab canonical files.
    # Canonical naming: <n>_<Label>_KB<number>_<arch>.<ext>
    $canonicalPattern = '^(\d+_[^_]+)_KB(\d+)_(\w+)\.(msu|cab)$'

    $allCanonical = Get-ChildItem $DownloadDir -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match $canonicalPattern }

    # Group by prefix + arch + ext (e.g. "1_LCU_x64_msu")
    $groups = $allCanonical | Group-Object {
        $m = [regex]::Match($_.Name, $canonicalPattern)
        "$($m.Groups[1].Value)_$($m.Groups[3].Value)_$($m.Groups[4].Value)"
    }

    foreach ($group in $groups) {
        if ($group.Count -le 1) { continue }

        # Sort by KB number descending — keep the highest
        $sorted = $group.Group | Sort-Object {
            $m = [regex]::Match($_.Name, $canonicalPattern)
            [int]$m.Groups[2].Value
        } -Descending

        # Remove all but the first (newest)
        $sorted | Select-Object -Skip 1 | ForEach-Object {
            Write-Info "Cleanup: removing superseded file: $($_.Name)"
            Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            $cleaned++
        }
    }

    if ($cleaned -gt 0) {
        Write-OK "Update folder cleanup: removed $cleaned stale/superseded file(s)"
        Write-Log "Update folder cleanup removed $cleaned file(s)"
    }
}

function Invoke-AutoUpdateDownload {
    param([string]$DownloadDir)

    if (-not (Test-Path $DownloadDir)) {
        New-Item $DownloadDir -ItemType Directory -Force | Out-Null
    }

    # Ensure MSCatalogLTS is available
    if (-not (Install-MSCatalogLTS)) {
        Write-Warn "Cannot download updates automatically without MSCatalogLTS."
        return $null
    }

    Write-Info "Searching Microsoft Update Catalog via MSCatalogLTS..."
    $downloaded = 0

    # MSCatalogLTS can find LCU and .NET CU reliably.
    # SafeOS Dynamic Updates are classified differently in the Catalog
    # (product: "Windows Safe OS Dynamic Update") and MSCatalogLTS cannot
    # find them regardless of search terms. Handle SafeOS separately below.
    foreach ($type in @("LCU", "DotNet")) {
        $label  = switch ($type) {
            "LCU"    { "Cumulative Update (LCU + SSU)" }
            "DotNet" { ".NET Framework CU" }
        }
        $prefix = switch ($type) { "LCU" { "1" } "DotNet" { "2" } }

        Write-Host ""
        Write-Host "  >> $label" -ForegroundColor White
        Write-Info "     Query: $($CatalogSearchTerms[$type])"

        try {
            # Do NOT pass -Architecture to Get-MSCatalogUpdate - it filters incorrectly.
            # Architecture and build filtering is handled by Where-Object on $_.Title below.
            $results = Get-MSCatalogUpdate -Search $CatalogSearchTerms[$type] -ErrorAction Stop
        } catch {
            Write-Warn "Search failed for $label`: $($_.Exception.Message)"
            continue
        }

        if (-not $results -or $results.Count -eq 0) {
            Write-Warn "No results found for $label."
            continue
        }

        # Filter by architecture, build number, and exclude Preview/x86.
        # BuildFilter prevents 24H2/25H2 mix-up when they share the same KB number.
        $buildFilter = $CatalogSearchTerms.BuildFilter
        $filtered = $results | Where-Object {
            $_.Title -notmatch "\bPreview\b" -and
            $_.Title -notmatch "\bx86\b"     -and
            ($Architecture -eq "arm64" -or $_.Title -notmatch "\barm64\b") -and
            ($Architecture -eq "x64"   -or $_.Title -notmatch "\bx64\b")   -and
            ($buildFilter -eq "" -or $_.Title -match [regex]::Escape("($buildFilter") -or $_.Title -match [regex]::Escape("($buildFilter."))
        }
        if (-not $filtered) {
            Write-Warn "Build filter ($buildFilter) matched nothing - falling back to unfiltered results"
            $filtered = $results | Where-Object {
                $_.Title -notmatch "\bPreview\b" -and
                $_.Title -notmatch "\bx86\b"     -and
                ($Architecture -eq "arm64" -or $_.Title -notmatch "\barm64\b") -and
                ($Architecture -eq "x64"   -or $_.Title -notmatch "\bx64\b")
            }
        }
        if (-not $filtered) { $filtered = $results }

        # Latest = first result
        $best = $filtered | Select-Object -First 1
        Write-Host "  Found   : $($best.Title)" -ForegroundColor Green
        Write-Host "  Date    : $($best.LastUpdated)" -ForegroundColor DarkGray

        # Build canonical filename for non-LCU types
        $kbMatch  = [regex]::Match($best.Title, "KB\d+")
        $kbNum    = if ($kbMatch.Success) { $kbMatch.Value } else { "KB_unknown" }
        $kbLower  = $kbNum.ToLower()
        $archMsu  = if ($Architecture -eq "arm64") { "arm64" } else { "x64" }
        $fileName = if ($type -eq "LCU") {
            "windows11.0-$kbLower-$archMsu.msu"   # placeholder, real name determined from URL
        } else {
            "${prefix}_${type}_${kbNum}_${Architecture}.msu"
        }
        $destPath = Join-Path $DownloadDir $fileName

        # ── LCU: download via direct API to preserve original filename with hash ──
        # MSCatalogLTS strips the hash from the filename when saving, causing DISM
        # signature validation to fail (0x80070241/0x80070570). We bypass Save-MSCatalogUpdate
        # for LCU and instead call the DownloadDialog API directly to get the real URL,
        # then download with Invoke-WebRequest preserving the full original filename.
        if ($type -eq "LCU") {
            # Check cache: look for any file matching KB number + arch with hash in name
            $alreadyHave = Get-ChildItem $DownloadDir -Filter "*.msu" -ErrorAction SilentlyContinue |
                           Where-Object { $_.Name -match $kbNum -and $_.Name -match $archMsu -and $_.Length -gt 1MB } |
                           Select-Object -First 1

            if ($alreadyHave) {
                # Real Microsoft LCU filenames contain a SHA1 hash segment after the KB number,
                # e.g. windows11.0-kb5083769-x64_57f4bd47...msu. Files without a hash were
                # downloaded by an older code path (MSCatalogLTS strips the hash) and must be
                # re-downloaded - DISM validates the filename against the package signature.
                if ($alreadyHave.Name -match '_[0-9a-f]{8,}') {
                    Write-OK "Already downloaded: $($alreadyHave.Name) ($([math]::Round($alreadyHave.Length/1MB,0)) MB)"
                    $destPath = $alreadyHave.FullName
                    $downloaded++
                    continue
                } else {
                    Write-Info "Cached LCU has no hash in filename - re-downloading with correct name..."
                    Remove-Item $alreadyHave.FullName -Force -ErrorAction SilentlyContinue
                }
            }

            # Get download URLs via DownloadDialog API (same as MSCatalogLTS Get-UpdateLinks)
            Write-Info "     Fetching download URLs for $kbNum via Update Catalog API..."
            try {
                $guid = $best.Guid
                $post = @{size = 0; UpdateID = $guid; UpdateIDInfo = $guid} | ConvertTo-Json -Compress
                $body = @{UpdateIDs = "[$post]"}
                $response = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/DownloadDialog.aspx" `
                    -Body $body -ContentType "application/x-www-form-urlencoded" `
                    -UseBasicParsing -ErrorAction Stop
                $content = $response.Content -replace "www.download.windowsupdate", "download.windowsupdate"

                # Extract all download URLs from response
                $regex = "downloadInformation\[(\d+)\]\.files\[(\d+)\]\.url\s*=\s*'([^']*)'"
                # Named $matches2 (not $matches) to avoid collision with PowerShell's
                # automatic $Matches variable populated by the -match operator above.
                $matches2 = [regex]::Matches($content, $regex)

                # Find the URL matching our architecture
                $lcuUrl = $null
                foreach ($m in $matches2) {
                    $url = $m.Groups[3].Value
                    if ($url -match $archMsu -and $url -match $kbLower -and $url -match "\.msu") {
                        $lcuUrl = $url
                        break
                    }
                }

                # If no arch-specific match, try first .msu URL
                if (-not $lcuUrl) {
                    foreach ($m in $matches2) {
                        $url = $m.Groups[3].Value
                        if ($url -match "\.msu") {
                            $lcuUrl = $url
                            break
                        }
                    }
                }

                if (-not $lcuUrl) {
                    Write-Warn "Could not find download URL for LCU from API - falling back to Save-MSCatalogUpdate"
                    throw "No URL found"
                }

                # Extract original filename with hash from URL
                $originalName = $lcuUrl.Split('/')[-1]
                $destPath = Join-Path $DownloadDir $originalName
                Write-Info "     Downloading $originalName..."
                Start-BitsTransfer -Source $lcuUrl -Destination $destPath -ErrorAction Stop
                # Unblock file to remove Zone.Identifier MoTW
                Unblock-File -Path $destPath -ErrorAction SilentlyContinue
                $sizeMB = [math]::Round((Get-Item $destPath).Length / 1MB, 0)
                Write-OK "Downloaded: $originalName ($sizeMB MB)"

                # Also download checkpoint via same API
                foreach ($m in $matches2) {
                    $url = $m.Groups[3].Value
                    if ($url -match "kb5043080" -and $url -match $archMsu -and $url -match "\.msu") {
                        $cpName = $url.Split('/')[-1]
                        $cpPath = Join-Path $DownloadDir $cpName
                        if (-not (Test-Path $cpPath)) {
                            Write-Info "     Downloading checkpoint: $cpName..."
                            Start-BitsTransfer -Source $url -Destination $cpPath -ErrorAction Stop
                            Unblock-File -Path $cpPath -ErrorAction SilentlyContinue
                            Write-OK "Checkpoint prerequisite: $cpName"
                        } else {
                            Write-OK "Checkpoint already cached: $cpName"
                        }
                        # Remove any old hash-less copy of the checkpoint left from a previous run
                        $cpKB = [regex]::Match($cpName, 'kb\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Value
                        Get-ChildItem $DownloadDir -Filter "*.msu" -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -match $cpKB -and $_.Name -notmatch '_[0-9a-f]{8,}' } |
                            ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue; Write-Info "Removed old-named checkpoint: $($_.Name)" }
                        break
                    }
                }

                $downloaded++
                continue
            } catch {
                # Hard fail: do NOT fall through to Save-MSCatalogUpdate.
                # That cmdlet strips the SHA1 hash from the filename, which causes DISM to
                # reject the package with 0x80070241/0x80070570 (signature validation failure).
                Write-Fail "LCU download via DownloadDialog API failed: $($_.Exception.Message)"
                Write-Host "" 
                Write-Host "  The LCU must be downloaded manually with its original Microsoft filename." -ForegroundColor Yellow
                Write-Host "  The filename must include the SHA1 hash, e.g.:" -ForegroundColor Yellow
                Write-Host "    windows11.0-$kbLower-$archMsu`_<hash>.msu" -ForegroundColor White
                Write-Host "" 
                Write-Host "  Steps:" -ForegroundColor White
                Write-Host "    1. Go to https://www.catalog.update.microsoft.com" -ForegroundColor Cyan
                Write-Host "    2. Search for: $kbNum" -ForegroundColor Cyan
                Write-Host "    3. Click Download -> right-click the link -> Save link as" -ForegroundColor Cyan
                Write-Host "       (do NOT use the Download button - it strips the hash)" -ForegroundColor Yellow
                Write-Host "    4. Save the file to: $DownloadDir" -ForegroundColor Cyan
                Write-Host "    5. Re-run the script - it will detect the cached file automatically" -ForegroundColor Cyan
                Write-Host ""
                return $null
            }
        }

        # ── Non-LCU types (DotNet): use Save-MSCatalogUpdate with canonical rename ──
        # Check cache
        $alreadyHave = Get-ChildItem $DownloadDir -Filter "*.msu" -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -match $kbNum -and $_.Name -match $archMsu -and $_.Length -gt 1MB } |
                       Select-Object -First 1

        if ($alreadyHave) {
            if ($alreadyHave.FullName -ne $destPath) {
                Rename-Item $alreadyHave.FullName $destPath -ErrorAction SilentlyContinue
                Write-OK "Found in cache (renamed to canonical): $fileName"
            } else {
                Write-OK "Already downloaded: $fileName ($([math]::Round($alreadyHave.Length/1MB,0)) MB)"
            }
            $downloaded++
            continue
        }

        Write-Info "     Downloading $fileName (this may take a while)..."
        try {
            Save-MSCatalogUpdate -Update $best -Destination $DownloadDir -ErrorAction Stop

            $allNewMSUs = Get-ChildItem $DownloadDir -Filter "*.msu" |
                          Where-Object { $_.Name -notmatch "^[0-9]_" }

            foreach ($newFile in $allNewMSUs) {
                $fileKB = ([regex]::Match($newFile.Name, 'KB\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Value.ToUpper()
                if ($fileKB -eq $kbNum -and $newFile.FullName -ne $destPath) {
                    Rename-Item $newFile.FullName $destPath -Force -ErrorAction SilentlyContinue
                }
            }

            if (Test-Path $destPath) {
                $sizeMB = [math]::Round((Get-Item $destPath).Length / 1MB, 0)
                Write-OK "Downloaded: $fileName ($sizeMB MB)"
                $downloaded++
            }
        } catch {
            Write-Warn "Download failed for $label`: $($_.Exception.Message)"
            if (Test-Path $destPath) { Remove-Item $destPath -Force -ErrorAction SilentlyContinue }
        }
    }

    # ── SafeOS Dynamic Update ──────────────────────────────────────────────────
    # MSCatalogLTS wraps all search terms in quotes. Long phrases fail because the
    # Catalog's API doesn't surface Dynamic Update packages by title phrase.
    # Short search terms like "Safe OS" work, and searching by exact KB number works.
    # Strategy:
    #   1. Check cache for existing SafeOS cab
    #   2. Search by short term "Safe OS" with -IncludeDynamic, filter x64 + version
    #   3. Fall back to manual instructions
    Write-Host ""
    Write-Host "  >> Safe OS Dynamic Update (WinRE)" -ForegroundColor White

    $safeOSExisting = Get-ChildItem $DownloadDir -Filter "3_SafeOS_*_${Architecture}.cab" -ErrorAction SilentlyContinue |
                      Select-Object -First 1
    if (-not $safeOSExisting) {
        # Also catch original MSCatalog download names (e.g. windows11.0-kb5083482-arm64.cab)
        # Both patterns must match the correct architecture.
        $safeOSExisting = Get-ChildItem $DownloadDir -Filter "*.cab" -ErrorAction SilentlyContinue |
                          Where-Object { ($_.Name -match "SafeOS.*_${Architecture}" -or $_.Name -match "windows11\.0-kb\d+-$Architecture") -and
                                         $_.Length -gt 1MB } |
                          Select-Object -First 1
    }

    if ($safeOSExisting) {
        Write-OK "Found SafeOS in cache: $($safeOSExisting.Name)"
        $canonical = Join-Path $DownloadDir "3_SafeOS_$(([regex]::Match($safeOSExisting.Name,'KB\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Value.ToUpper())_${Architecture}.cab"
        if ($safeOSExisting.FullName -ne $canonical -and [regex]::IsMatch($safeOSExisting.Name, 'KB\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            if (Test-Path $canonical) {
                # Canonical already exists (from a previous run) - just delete the old-named duplicate
                Remove-Item $safeOSExisting.FullName -Force -ErrorAction SilentlyContinue
                Write-Info "Removed old-named SafeOS duplicate: $($safeOSExisting.Name)"
            } else {
                Rename-Item $safeOSExisting.FullName $canonical -ErrorAction SilentlyContinue
            }
        }
        $downloaded++
    } else {
        # Search using short term - MSCatalogLTS can find SafeOS when the query is
        # short enough to match Catalog entries. Filter to x64 + our version string.
        Write-Info "     Searching for SafeOS (short-term search with -IncludeDynamic)..."
        $safeOSResult = $null
        $versionStr   = $CatalogSearchTerms.SafeOSVersion   # e.g. "25H2"

        $shortTerms = @("Safe OS $versionStr", "Safe OS Dynamic Update $versionStr", "Safe OS")
        foreach ($term in $shortTerms) {
            try {
                $candidates = Get-MSCatalogUpdate -Search $term -IncludeDynamic -ErrorAction Stop |
                              Where-Object {
                                  $_.Title -match "Safe OS" -and
                                  $_.Title -match $Architecture -and
                                  $_.Title -match $versionStr -and
                                  $_.Title -notmatch "\bPreview\b" -and
                                  ($Architecture -eq "arm64" -or $_.Title -notmatch "\barm64\b") -and
                                  ($Architecture -eq "x64"   -or $_.Title -notmatch "\bx64\b")
                              } |
                              Sort-Object LastUpdated -Descending |
                              Select-Object -First 1
                if ($candidates) {
                    $safeOSResult = $candidates
                    Write-Info "     Found using term: '$term'"
                    break
                }
            } catch {
                # Try next term
            }
        }

        if ($safeOSResult) {
            Write-OK "Found: $($safeOSResult.Title)"
            $safeKB   = [regex]::Match($safeOSResult.Title, 'KB\d+').Value
            $safeDest = Join-Path $DownloadDir "3_SafeOS_${safeKB}_${Architecture}.cab"
            try {
                Save-MSCatalogUpdate -Update $safeOSResult -Destination $DownloadDir -ErrorAction Stop
                # Rename to canonical
                $savedSafe = Get-ChildItem $DownloadDir -Filter "*.cab" |
                             Where-Object { $_.Name -match $safeKB -and $_.Name -ne "3_SafeOS_${safeKB}_${Architecture}.cab" } |
                             Select-Object -First 1
                if ($savedSafe) { Rename-Item $savedSafe.FullName $safeDest -Force }
                if (Test-Path $safeDest) {
                    $sizeMB = [math]::Round((Get-Item $safeDest).Length / 1MB, 0)
                    Write-OK "Downloaded: 3_SafeOS_${safeKB}_${Architecture}.cab ($sizeMB MB)"
                    $downloaded++
                }
            } catch {
                Write-Warn "SafeOS download failed: $($_.Exception.Message)"
            }
        } else {
            Write-Warn "SafeOS not found automatically. Manual download required."
            Write-Host ""
            Write-Host "  To download SafeOS manually:" -ForegroundColor Yellow
            Write-Host "  1. Go to https://www.catalog.update.microsoft.com" -ForegroundColor White
            Write-Host "  2. Search for: Safe OS Dynamic Update Windows 11 $versionStr" -ForegroundColor White
            Write-Host "  3. Download the .cab file for $Architecture" -ForegroundColor White
            Write-Host "  4. Save it to: $DownloadDir" -ForegroundColor White
            Write-Host "  5. Name it:    3_SafeOS_KBxxxxxxx_${Architecture}.cab" -ForegroundColor White
            Write-Host "  The script will use it on the next run, or you can add it now" -ForegroundColor DarkGray
            Write-Host "  and re-run with -SkipLanguagePacks to save time." -ForegroundColor DarkGray
            Write-Host ""
            Write-Info "Continuing without SafeOS - WinRE will be patched with LCU only."
        }
    }

    if ($downloaded -eq 0) { return $null }

    Write-Host ""
    Write-OK "$downloaded of 3 update types downloaded to: $DownloadDir"
    return $DownloadDir
}

#endregion

#region ── Step 0: Banner and prerequisites ───────────────────────────────────

Write-Banner

Write-Section "Checking prerequisites"

# Check for stale WIM mount points from previous runs and clean them up
$staleMounts = Get-WindowsImage -Mounted -ErrorAction SilentlyContinue |
               Where-Object { $_.MountStatus -ne "Ok" }
if ($staleMounts) {
    Write-Warn "Stale WIM mount point(s) detected from a previous run:"
    $staleMounts | ForEach-Object { Write-Host "    $($_.Path)" -ForegroundColor Yellow }
    Write-Info "Cleaning up stale mounts..."
    & dism.exe /Cleanup-Mountpoints | Out-Null
    Write-OK "Stale mounts cleaned"
}

# DISM
$dismOut = & dism.exe /English /? 2>&1
if ($LASTEXITCODE -gt 1) {
    Write-Fail "DISM not found."
    exit 1
}
Write-OK "DISM available"

# Disk space
$drive  = Get-PSDrive ($env:SystemDrive.TrimEnd(':'))
$freeGB = [math]::Round($drive.Free / 1GB, 1)
if ($freeGB -lt 30) {
    Write-Warn "Only $freeGB GB free. At least 30 GB recommended."
} else {
    Write-OK "$freeGB GB free disk space"
}

# Internet access
$internetOK = $false
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $null = Invoke-WebRequest "https://www.catalog.update.microsoft.com" -UseBasicParsing -TimeoutSec 10 `
                -Headers @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" }
    $internetOK = $true
    Write-OK "Internet access available"
} catch {
    Write-Warn "No internet access detected ($($_.Exception.Message))"
    Write-Warn "Automatic update download not available - manual mode only"
}

#endregion

#region ── Step 1: Source folder ──────────────────────────────────────────────

if ($PatchMode) {
    Write-Section "Step 1/4 - Patch Mode (existing WIM)"
    Write-Host "  Patching existing WIM - skipping ISO discovery." -ForegroundColor Cyan
    Write-Host "  Source WIM : $PatchExistingWim" -ForegroundColor Cyan
    Write-Host "  Edition    : $($patchWimInfo.ImageName)" -ForegroundColor Cyan
    Write-Host "  Version    : $($patchWimInfo.Version)" -ForegroundColor Cyan
    Write-Host "  Languages  : $($patchWimInfo.Languages -join ', ')" -ForegroundColor Cyan
    Write-Host ""

    # In patch mode, use the existing WIM directly
    $wimFile      = $PatchExistingWim
    $WimIndex     = 1
    $chosenImage  = $patchWimInfo
    $chosenImageFull = $patchWimInfo
    $CatalogSearchTerms = Get-CatalogSearchTerms -WimVersion $patchWimInfo.Version
    $LPSearchRoot = $null
    $FODSearchRoot = $null

} else {

Write-Section "Step 1/4 - Source folder"

Write-Host ""
Write-Host "  Place the required ISO file(s) in a single folder, e.g. $($ScriptRoot)\ISO-Source\" -ForegroundColor White
Write-Host ""
Write-Host "  [1] Windows 11 Enterprise 25H2 ISO" -ForegroundColor Yellow
Write-Host "      x64:   SW_DVD9_Win_Pro_11_25H2_64BIT_English_Pro_Ent_EDU_N_MLF_*.ISO" -ForegroundColor DarkGray
Write-Host "      ARM64: SW_DVD9_Win_Pro_11_25H2*Arm64_English_Pro_Ent_EDU_N_MLF_*.ISO" -ForegroundColor DarkGray
Write-Host "      (Future versions will follow same pattern: Win_Pro_11_26H2_...)" -ForegroundColor DarkGray
Write-Host ""
$needLPISOHint = -not $SkipLanguagePacks -or $FoDList -ne ""
if ($needLPISOHint) {
    Write-Host "  [2] Language Pack ISO (also contains all FOD packages - no separate FOD ISO needed)" -ForegroundColor Yellow
    Write-Host "      x64:   SW_DVD9_Win_11_24H2_25H2_x64_MultiLang_LangPackAll_LIP_LoF_*.ISO" -ForegroundColor DarkGray
    Write-Host "      ARM64: SW_DVD9_Win_11_24H2_25H2_Arm64_MultiLang_LangPackAll_LIP_LoF_*.ISO" -ForegroundColor DarkGray
    Write-Host "      (Future versions will follow same pattern: Win_11_26H2_..._LangPack_...)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Download both ISOs from the Microsoft 365 Admin Center:" -ForegroundColor White
} else {
    Write-Host "  (Language Pack ISO not required - language pack injection is disabled)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Download the Windows ISO from the Microsoft 365 Admin Center:" -ForegroundColor White
}
Write-Host "  https://admin.microsoft.com/adminportal/home#/subscriptions/vlnew/downloadsandkeys" -ForegroundColor Cyan
Write-Host "  Sign in -> Downloads & Keys -> Windows 11 Enterprise -> 25H2" -ForegroundColor DarkGray
Write-Host ""

if (-not (Test-Path $SourceFolder)) {
    # Check if this looks like a mapped drive letter that's unavailable in the elevated session
    if ($SourceFolder -match '^[A-Za-z]:\\' -and $SourceFolder -notmatch '^[Cc]:\\') {
        $driveLetter = $SourceFolder.Substring(0, 2)
        if (-not (Test-Path $driveLetter)) {
            Write-Host ""
            Write-Warn "Drive $driveLetter is not available in this elevated session."
            Write-Host ""
            Write-Host "  This usually means $driveLetter is a mapped network drive that was" -ForegroundColor Yellow
            Write-Host "  connected as your regular user, but PowerShell is running as Administrator." -ForegroundColor Yellow
            Write-Host "  Elevated sessions do not inherit user-mapped drives." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Solutions:" -ForegroundColor White
            Write-Host "    1. Use a UNC path instead:  \\\\server\\share\\ISO-folder" -ForegroundColor Cyan
            Write-Host "    2. Remap the drive in an elevated prompt:" -ForegroundColor Cyan
            Write-Host "       net use $driveLetter \\\\server\\share /persistent:no" -ForegroundColor Cyan
            Write-Host "    3. Copy the ISOs to a local folder (e.g. C:\WimWizard\ISO-Source)" -ForegroundColor Cyan
            Write-Host ""
            Dismount-AllISOs; exit 1
        }
    }
    # Folder doesn't exist - prompt the user
    $SourceFolder = Read-PathWithDefault `
        -Prompt  "Path to folder containing the two ISO files:" `
        -Default "$ScriptRoot\ISO-Source" `
        -MustExist
    if (-not $SourceFolder) { exit 1 }
}

# ── Discover and mount ISOs ───────────────────────────────────────────────────

Write-Info "Scanning $SourceFolder for ISO files..."
# Mount LP ISO if language packs are needed OR if FoDs are requested
# (FoD cabs are sourced from the LP ISO LanguagesAndOptionalFeatures folder)
$needLPISO  = -not $SkipLanguagePacks -or $FoDList -ne ""
$discovered = Resolve-SourceFolder -Folder $SourceFolder -SkipLP:(-not $needLPISO)

# Resolved variables used throughout the rest of the script
$wimFile      = $null
$wimISOPath   = $null
$LPSearchRoot = $null
$FODSearchRoot= $null

if ($discovered) {
    $wimFile       = $discovered.WimFile
    $wimISOPath    = $discovered.WimISO
    $LPSearchRoot  = $discovered.LPSearchRoot
    $FODSearchRoot = $discovered.FODSearchRoot

    # If LP ISO was not found (and not skipped), note it - the Step 2 pre-flight will handle
    # the error with a proper message. Don't silently fall back to the whole source folder.
    if (-not $LPSearchRoot) {
        if (-not $SkipLanguagePacks) {
            Write-Warn "No Language Pack ISO found in $SourceFolder"
        }
        $LPSearchRoot  = $SourceFolder   # fallback so later code doesn't crash on null
        $FODSearchRoot = $SourceFolder
    }
} else {
    # No ISOs found - maybe the folder contains an already-extracted WIM or subfolders
    Write-Warn "No ISO files found in $SourceFolder - looking for install.wim/esd directly..."
    $found = Get-ChildItem $SourceFolder -Filter "install.wim" -Recurse | Select-Object -First 1
    if (-not $found) { $found = Get-ChildItem $SourceFolder -Filter "install.esd" -Recurse | Select-Object -First 1 }
    if ($found) {
        $wimFile       = $found.FullName
        $LPSearchRoot  = $SourceFolder
        $FODSearchRoot = $SourceFolder
        Write-OK "Found WIM directly: $wimFile"
    } else {
        Write-Fail "No ISO files and no install.wim/esd found in: $SourceFolder"
        Write-Host ""
        if ($SkipLanguagePacks -and $FoDList -eq "") {
            Write-Host "  Make sure you have placed the Windows 11 ISO in the folder." -ForegroundColor Yellow
            Write-Host "  (Language Pack ISO is not required when using -SkipLanguagePacks)" -ForegroundColor DarkGray
        } elseif ($SkipLanguagePacks -and $FoDList -ne "") {
            Write-Host "  Make sure you have placed both ISO files in the folder." -ForegroundColor Yellow
            Write-Host "  (Language Pack ISO is still required because -FoDList was specified)" -ForegroundColor DarkGray
        } else {
            Write-Host "  Make sure you have placed both ISO files in the folder." -ForegroundColor Yellow
            Write-Host "  (Windows 11 Enterprise ISO + Language Pack ISO)" -ForegroundColor DarkGray
        }
        Dismount-AllISOs
        exit 1
    }
}

if (-not $wimFile -or -not (Test-Path $wimFile)) {
    Write-Fail "Could not locate install.wim or install.esd."
    Write-Host ""
    Write-Host "  The Windows ISO was not recognised. Check the filename contains" -ForegroundColor Yellow
    Write-Host "  'windows', '11', 'enterprise' or 'business' and is the OS ISO," -ForegroundColor Yellow
    Write-Host "  not the Language Pack or FOD ISO." -ForegroundColor Yellow
    Dismount-AllISOs
    exit 1
}

# ── Select WIM index (auto-detect Enterprise) ─────────────────────────────────

$images = Get-WindowsImage -ImagePath $wimFile

# Auto-detect Enterprise index
$enterpriseImage = $images | Where-Object { $_.ImageName -match "Enterprise" -and $_.ImageName -notmatch "Evaluation" } |
                   Select-Object -First 1

if ($WimIndex -ne 0) {
    # Caller forced a specific index
    $chosenImage = $images | Where-Object { $_.ImageIndex -eq $WimIndex }
    if (-not $chosenImage) {
        Write-Fail "Index $WimIndex not found in WIM."
        Dismount-AllISOs; exit 1
    }
    Write-OK "Using forced index: [$WimIndex] $($chosenImage.ImageName)"

} elseif ($enterpriseImage) {
    # Found Enterprise automatically
    $WimIndex    = $enterpriseImage.ImageIndex
    $chosenImage = $enterpriseImage
    Write-OK "Auto-selected: [$WimIndex] $($chosenImage.ImageName)"

} else {
    # Could not auto-detect Enterprise edition
    Write-Warn "Could not auto-detect Enterprise edition. Available editions:"
    foreach ($img in $images) {
        Write-Host ("    [{0}] {1}" -f $img.ImageIndex, $img.ImageName) -ForegroundColor White
    }
    Write-Host ""
    if ($Unattended) {
        Write-Fail "Unattended mode: cannot prompt for WIM index. Use -WimIndex to specify one."
        Dismount-AllISOs; exit 1
    }
    $WimIndex    = [int](Read-Host "  Enter index number")
    $chosenImage = $images | Where-Object { $_.ImageIndex -eq $WimIndex }
    if (-not $chosenImage) {
        Write-Fail "Index $WimIndex not found."
        Dismount-AllISOs; exit 1
    }
    Write-OK "Selected: [$WimIndex] $($chosenImage.ImageName)"
}

# Now that we know the WIM version, build the correct Catalog search terms.
# Get-WindowsImage without -Index returns a summary list without the Version property.
# Fetch the full info for the selected index to get the actual build number.
$chosenImageFull  = Get-WindowsImage -ImagePath $wimFile -Index $WimIndex
$CatalogSearchTerms = Get-CatalogSearchTerms -WimVersion $chosenImageFull.Version

} # end if ($PatchMode)

#endregion

#region ── Step 2: Language packs ────────────────────────────────────────────

if (-not $SkipLanguagePacks) {

    # Resolve which locales to install
    if ($Languages -ne "") {
        $ResolvedLocales = @(Resolve-LanguageCodes -CodeString $Languages)
    } elseif ($Unattended) {
        Write-Info "Unattended with no -Languages - skipping language packs (English only)."
        $SkipLanguagePacks = $true
    } else {
        Write-Host ""
        Write-Host "  Supported language codes:" -ForegroundColor White
        Write-Host "  $SupportedCodesLine" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  Enter comma-separated 2-letter country codes, e.g: se,no,dk,fi" -ForegroundColor White
        Write-Host "  (Press Enter for default: se,no,dk,fi)" -ForegroundColor DarkGray
        $langInput = Read-Host "  Languages"
        if ($langInput.Trim() -eq "") { $langInput = "se,no,dk,fi" }
        $Languages = $langInput   # store for supplemental FOD resolution below
        $ResolvedLocales = @(Resolve-LanguageCodes -CodeString $langInput)
    }

    # Build supplemental FOD locale list - locales that need extra FODs beyond their base LP.
    # e.g. 'hk' resolves to zh-TW for the LP but also needs zh-HK specific FODs.
    $SupplementalFODMap = @{
        'hk' = 'zh-HK'   # zh-HK FODs on top of zh-TW LP
    }
    $SupplementalFODLocales = @()
    foreach ($code in ($Languages.ToLower().Split(',') | ForEach-Object { $_.Trim() })) {
        if ($SupplementalFODMap.ContainsKey($code)) {
            $SupplementalFODLocales += $SupplementalFODMap[$code]
        }
    }

    if (-not $SkipLanguagePacks) {
        Write-Section "Step 2/4 - Language Packs + FOD"
        Write-Host "  Languages to install: $($ResolvedLocales -join ', ')" -ForegroundColor Cyan
        Write-Host ""

        # Scan LP/FOD folder ONCE and cache results - avoids rescanning 7800+ files per language.
        Write-Info "Scanning language pack folder (this may take a moment)..."
        $allLPCabs = @(Get-ChildItem $LPSearchRoot -Filter "*.cab" -ErrorAction SilentlyContinue)
        Write-Info "Found $($allLPCabs.Count) cab files in LP folder."

        # LP cab filenames use "x64" for x64 ISO and "arm64" for ARM64 ISO.
        # FOD cabs use "amd64" for x64 and "arm64" for ARM64.
        $lpArchStr  = if ($Architecture -eq "arm64") { "arm64" } else { "x64" }
        $fodArchStr = if ($Architecture -eq "arm64") { "arm64" } else { "amd64" }

        # Validate each language before installing
        $anyMissing = $false
        foreach ($lang in $ResolvedLocales) {
            $lp  = $allLPCabs | Where-Object { $_.Name -match "Language-Pack_${lpArchStr}_$lang" } | Select-Object -First 1
            $fod = $allLPCabs | Where-Object { $_.Name -match "LanguageFeatures-Basic-$lang" -and
                                               $_.Name -match "~${fodArchStr}~" } | Select-Object -First 1
            if (-not $lp)  { $anyMissing = $true }
            if (-not $fod) { $anyMissing = $true }
            $lc  = if ($lp)  { "Green" } else { "Red" }
            $fc  = if ($fod) { "Green" } else { "Red" }
            Write-Host ("  [$($lang.ToUpper())]  LP: ") -NoNewline
            Write-Host ("{0,-8}" -f $(if ($lp)  { "OK" } else { "MISSING" })) -ForegroundColor $lc -NoNewline
            Write-Host "  FOD Basic: " -NoNewline
            Write-Host ("{0,-8}" -f $(if ($fod) { "OK" } else { "MISSING" })) -ForegroundColor $fc -NoNewline
            Write-Host "  [$(if ($lp) { $lp.Name } else { 'NOT FOUND' })]" -ForegroundColor DarkGray
        }

        if ($anyMissing) {
            Write-Host ""
            Write-Fail "Language Pack ISO not found or missing cab files in: $SourceFolder"
            Write-Host ""
            Write-Host "  The Language Pack ISO must be present in the source folder." -ForegroundColor Yellow
            Write-Host "  It also contains all FOD packages - no separate FOD ISO is needed." -ForegroundColor Yellow
            Write-Host "  Expected filename like: SW_DVD9_Win_11_24H2_25H2_x64_Multilang_LangPack_*.ISO" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  Download from the Microsoft 365 Admin Center:" -ForegroundColor White
            Write-Host "  https://admin.microsoft.com/adminportal/home#/subscriptions/vlnew/downloadsandkeys" -ForegroundColor Cyan
            Write-Host "  Sign in -> Downloads & Keys -> Windows 11 Enterprise -> Language Pack" -ForegroundColor DarkGray
            Write-Host ""
            if ($Unattended) {
                Dismount-AllISOs; exit 1
            }
            if (-not (Confirm-Continue "Continue anyway (missing languages will be skipped)?")) {
                Dismount-AllISOs; exit 0
            }
        } else {
            Write-OK "All language pack files found"
        }

    }
}

#endregion

#region ── Step 3: Updates ────────────────────────────────────────────────────

$resolvedUpdatePath = $null

if (-not $SkipUpdates) {
    Write-Section "Step 3/4 - Updates (Patch Tuesday)"

    if ($UpdatePath) {
        # Explicit manual path provided
        $resolvedUpdatePath = $UpdatePath
        $updCount = @(
            Get-ChildItem $resolvedUpdatePath -Filter "*.msu" -ErrorAction SilentlyContinue
            Get-ChildItem $resolvedUpdatePath -Filter "*.cab" -ErrorAction SilentlyContinue
        ).Count
        Write-OK "Using specified update folder: $UpdatePath ($updCount files)"

    } elseif ($internetOK) {
        # Automatic download
        Write-Host ""
        Write-Host "  Automatic download from Microsoft Update Catalog." -ForegroundColor White
        Write-Host "  The script will search and download the latest version of:" -ForegroundColor DarkGray
        Write-Host "    1. LCU    - Cumulative Update (includes SSU)" -ForegroundColor White
        Write-Host "    2. .NET   - .NET Framework 3.5 / 4.8.1 Cumulative Update" -ForegroundColor White
        Write-Host "    3. SafeOS - Safe OS Dynamic Update (for WinRE)" -ForegroundColor White
        Write-Host ""
        Write-Host "  Files are saved to $($ScriptRoot)\Updates\" -ForegroundColor DarkGray
        Write-Host "  Already-downloaded KBs are reused automatically (checked by KB number)." -ForegroundColor DarkGray
        Write-Warn "  First-time download may take 10-20 min (LCU alone is 500-800 MB)."

        if ($PatchMode -or (Confirm-Continue "Download updates now?")) {
            $cacheDir = "$ScriptRoot\Updates"
            Invoke-UpdateFolderCleanup -DownloadDir $cacheDir
            $resolvedUpdatePath = Invoke-AutoUpdateDownload -DownloadDir $cacheDir
            if (-not $resolvedUpdatePath) {
                Write-Warn "Automatic download failed. Falling back to manual mode..."
            }
        }
    }

    # Manual fallback
    if (-not $resolvedUpdatePath) {
        Write-Host ""
        Write-Host "  MANUAL DOWNLOAD - from Microsoft Update Catalog:" -ForegroundColor White
        Write-Host "  Link: https://www.catalog.update.microsoft.com" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Search for the following for Windows 11 Version 24H2/$($CatalogSearchTerms.SafeOSVersion) $Architecture):" -ForegroundColor White
        Write-Host "    LCU     - 'Cumulative Update for Windows 11 Version 24H2 for $Architecture'" -ForegroundColor Yellow
        Write-Host "    .NET CU - '.NET Framework 3.5 4.8.1 Windows 11 24H2 $Architecture'" -ForegroundColor Yellow
        Write-Host "    SafeOS  - 'Safe OS Dynamic Update Windows 11 24H2'" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  IMPORTANT - LCU filename must include the SHA1 hash from Microsoft:" -ForegroundColor White
        Write-Host "    Correct : windows11.0-kb5083769-x64_57f4bd47...msu" -ForegroundColor Green
        Write-Host "    Wrong   : windows11.0-kb5083769-x64.msu  (hash stripped - DISM will reject)" -ForegroundColor Red
        Write-Host "  To get the correct filename: click Download -> right-click the link -> Save link as." -ForegroundColor White
        Write-Host ""
        Write-Host "  For .NET and SafeOS, place the files in the folder as downloaded." -ForegroundColor DarkGray
        Write-Host "  NOTE: SSU is bundled inside LCU for 24H2/25H2 - no separate download needed." -ForegroundColor DarkGray
        Write-Host ""

        $manualPath = Read-PathWithDefault `
            -Prompt  "Path to folder with downloaded .msu/.cab files (press Enter to skip):" `
            -Default "$ScriptRoot\Updates"

        if ($manualPath -and (Test-Path $manualPath)) {
            $updFiles = @(
                Get-ChildItem $manualPath -Filter "*.msu" -ErrorAction SilentlyContinue
                Get-ChildItem $manualPath -Filter "*.cab" -ErrorAction SilentlyContinue
            )
            if ($updFiles -and $updFiles.Count -gt 0) {
                $resolvedUpdatePath = $manualPath
                Write-Info "Found $($updFiles.Count) update file(s):"
                $updFiles | ForEach-Object { Write-Host "    $($_.Name)" -ForegroundColor White }
            } else {
                Write-Warn "No .msu/.cab files found - skipping updates."
            }
        } else {
            Write-Warn "Skipping updates."
        }
    }
}

#endregion

#region ── Step 4: Output path ────────────────────────────────────────────────

Write-Section "Step 4/4 - Output"

# Build a descriptive default filename
$_buildStr   = $chosenImageFull.Version -replace '^\d+\.\d+\.', ''   # e.g. 26200.8117
$_versionStr = $CatalogSearchTerms.SafeOSVersion                     # e.g. 25H2
$_dateStr    = Get-Date -Format "yyyyMMdd"

if ($PatchMode) {
    # Keep the same base filename as the source WIM but bump the date to today
    $_srcName  = [System.IO.Path]::GetFileName($PatchExistingWim)
    $_autoName = $_srcName -replace '_\d{8}\.wim$', "_${_dateStr}.wim"
    # If no date pattern found, just use the source name as-is
    if ($_autoName -eq $_srcName) { $_autoName = $_srcName }
} else {
    $_langStr = if ($SkipLanguagePacks -or @($ResolvedLocales).Count -eq 0) {
        "en"
    } else {
        # Use the raw user-supplied codes for the filename (e.g. da,fi,no,se -> da_fi_no_se)
        # This matches the GUI filename preview exactly.
        $rawCodes = $Languages.ToLower().Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($rawCodes) { $rawCodes -join "_" } else { "en" }
    }
    $_archSuffix = if ($Architecture -eq "arm64") { "_arm64" } else { "" }
    $_autoName = "Win11_${_versionStr}_${_buildStr}_${_langStr}${_archSuffix}_${_dateStr}.wim"
}
# If OutputPath points to a directory (root drive or existing folder with no filename),
# treat it as the output directory and append the auto-generated filename.
if ($OutputPath -and (Test-Path $OutputPath -PathType Container)) {
    $OutputPath = Join-Path $OutputPath $_autoName
} elseif ($OutputPath -and -not [System.IO.Path]::GetExtension($OutputPath)) {
    # No extension - assume it's a directory, even if it doesn't exist yet
    $OutputPath = Join-Path $OutputPath $_autoName
}

$_autoPath = Join-Path (Split-Path $OutputPath -Parent) $_autoName

if (-not $OutputPath -or $OutputPath -eq "$ScriptRoot\Output\install.wim") {
    $OutputPath = $_autoPath
}

$OutputPath = Read-PathWithDefault `
    -Prompt  "Path for the finished WIM file (including filename):" `
    -Default $OutputPath

$outputDir = Split-Path $OutputPath -Parent
if (-not $outputDir) { $outputDir = Split-Path $OutputPath -Qualifier }  # fallback for root drive paths
if (-not (Test-Path $outputDir)) { New-Item $outputDir -ItemType Directory -Force | Out-Null; Write-OK "Created: $outputDir" }
if (Test-Path $OutputPath) { Write-Warn "File already exists and will be overwritten." }
Write-OK "Output: $OutputPath"

#endregion

#region ── Summary and confirmation ───────────────────────────────────────────

Write-Section "Summary"

Write-Host ""
Write-Host "  The following actions will be performed:" -ForegroundColor White
Write-Host ""
if ($PatchMode) {
    Write-Host "    Mode          : PATCH (existing WIM - no LP/Appx changes)" -ForegroundColor Yellow
    Write-Host "    Source WIM    : $PatchExistingWim" -ForegroundColor Cyan
} else {
    Write-Host "    Source folder : $SourceFolder" -ForegroundColor Cyan
    Write-Host "    WIM source    : $wimFile" -ForegroundColor Cyan
}
Write-Host "    Edition       : [$WimIndex] $($chosenImage.ImageName)" -ForegroundColor Cyan

if (-not $SkipLanguagePacks) {
    Write-Host "    Language packs: $($ResolvedLocales -join ', ')  (+ FOD)" -ForegroundColor Cyan
    Write-Host "    LP/FOD source : $LPSearchRoot" -ForegroundColor Cyan
}

if (-not $SkipUpdates -and $resolvedUpdatePath) {
    $n = @(
        Get-ChildItem $resolvedUpdatePath -Filter "*.msu" -ErrorAction SilentlyContinue
        Get-ChildItem $resolvedUpdatePath -Filter "*.cab" -ErrorAction SilentlyContinue
    ).Count
    Write-Host "    Updates       : $resolvedUpdatePath ($n files)" -ForegroundColor Cyan
} elseif (-not $SkipUpdates) {
    Write-Host "    Updates       : Skipped" -ForegroundColor Yellow
}

if ($FoDList -ne "") {
    Write-Host "    Features (FoD): $FoDList" -ForegroundColor Cyan
}

Write-Host "    Output        : $OutputPath" -ForegroundColor Cyan
Write-Host ""
if (-not $PatchMode) {
    Write-Warn "This process typically takes ~60 minutes (4 langs). Do NOT abort while the WIM is mounted."
    Write-Host ""
}

if (-not (Confirm-Continue "Start servicing now?")) {
    Write-Info "Cancelled by user."
    Dismount-AllISOs
    exit 0
}

#endregion

#region ── Working directories and log ────────────────────────────────────────

$workRoot   = Join-Path $outputDir "WIMServicing_Work"
$mountDir   = Join-Path $workRoot "Mount"
$workWim    = Join-Path $workRoot "install_work.wim"
$scratchDir = Join-Path $workRoot "Scratch"   # DISM scratch space - avoids system drive disk-full

# Log file goes next to the script, not in the output folder.
# This avoids UNC/network share write issues when -OutputPath is a remote path.
$logFile    = Join-Path $outputDir ("WIMServicing_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmm"))

@($workRoot, $mountDir, $scratchDir) | ForEach-Object {
    if (-not (Test-Path $_)) { New-Item $_ -ItemType Directory -Force | Out-Null }
}

function Write-Log {
    param([string]$msg, [string]$level = "INFO")
    # Use File::AppendAllText instead of Add-Content - handles UNC paths reliably
    [System.IO.File]::AppendAllText($logFile, ("[{0}] [{1}] {2}`r`n" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $level, $msg))
}

Write-Log "=== WIM Wizard v$ScriptVersion ==="

# Load custom Appx list from XML - must be here, after Write-Log is defined
if (($AppxListPath -is [string]) -and ($AppxListPath.Length -gt 0) -and (Test-Path -LiteralPath $AppxListPath)) {
    try {
        [xml]$AppxXml = Get-Content $AppxListPath -ErrorAction Stop
        # Under StrictMode, a single XML child element returns XmlElement (not array) - no .Count.
        # Cast to [array] explicitly to ensure .Count always works.
        [array]$pkgNodes = $AppxXml.AppxRemovalList.Package
        if ($null -eq $pkgNodes -or $pkgNodes.Count -eq 0 -or $null -eq $pkgNodes[0]) {
            $AppxToRemove = @()
        } else {
            $AppxToRemove = @($pkgNodes | ForEach-Object { $_.Id } | Where-Object { $_ })
        }
        Write-Info "Loaded custom Appx list from: $AppxListPath ($($AppxToRemove.Count) packages)"
        Write-Log "Loaded custom Appx list: $AppxListPath ($($AppxToRemove.Count) packages)"
        if ($AppxToRemove.Count -eq 0) {
            Write-Info "Appx list is empty - skipping Appx removal"
            $SkipAppxRemoval = $true
        }
    } catch {
        Write-Warn "Could not load AppxListPath: $($_.Exception.Message) - using default list"
        Write-Log "WARN: Could not load AppxListPath: $($_.Exception.Message)" "WARN"
        $AppxToRemove = $AppxToRemoveDefault
    }
} else {
    $AppxToRemove = $AppxToRemoveDefault
}

# Print Appx summary here so it reflects the custom XML count (if loaded)
if (-not $SkipAppxRemoval) {
    Write-Host "    Appx removal  : $($AppxToRemove.Count) packages removed if present" -ForegroundColor Cyan
}

# Build a cut-pasteable command line for the log (line 2) - only includes params that were set
$_cmdParts = @(".\WimWizard.ps1")
if ($SourceFolder)     { $_cmdParts += "-SourceFolder `"$SourceFolder`"" }
if ($OutputPath)       { $_cmdParts += "-OutputPath `"$OutputPath`"" }
if ($WimIndex -ne 0)   { $_cmdParts += "-WimIndex $WimIndex" }
if ($Languages)        { $_cmdParts += "-Languages `"$Languages`"" }
if ($UpdatePath)       { $_cmdParts += "-UpdatePath `"$UpdatePath`"" }
if ($PatchExistingWim) { $_cmdParts += "-PatchExistingWim `"$PatchExistingWim`"" }
if ($FoDList)          { $_cmdParts += "-FoDList `"$FoDList`"" }
if ($AppxListPath)     { $_cmdParts += "-AppxListPath `"$AppxListPath`"" }
if ($SkipUpdates)      { $_cmdParts += "-SkipUpdates" }
if ($SkipLanguagePacks){ $_cmdParts += "-SkipLanguagePacks" }
if ($SkipAppxRemoval)  { $_cmdParts += "-SkipAppxRemoval" }
if ($ARM64)            { $_cmdParts += "-ARM64" } elseif ($X64) { $_cmdParts += "-X64" }
if ($Unattended)       { $_cmdParts += "-Unattended" }
if ($DebugBuild)       { $_cmdParts += "-DebugBuild" }
Write-Log ($_cmdParts -join " ")
Write-Log "Source: $SourceFolder  WIM: $wimFile  Index: $WimIndex  Edition: $($chosenImage.ImageName)"

#endregion

#region ── Main try block ─────────────────────────────────────────────────────

try {

    # Copy / convert WIM
    Write-Section "Copying source WIM"

    if ($PatchMode) {
        Copy-Item $PatchExistingWim $workWim -Force
        Set-ItemProperty $workWim -Name IsReadOnly -Value $false
        $workWimIndex = 1
        Write-OK "Existing WIM copied for patching"
    } elseif ($wimFile -match "\.esd$") {
        Write-Info "ESD format - exporting to WIM first (this may take a while)..."
        Export-WindowsImage -SourceImagePath $wimFile -SourceIndex $WimIndex `
            -DestinationImagePath $workWim -CompressionType fast
        $workWimIndex = 1
        Write-OK "ESD exported to WIM"
    } else {
        Copy-Item $wimFile $workWim -Force
        Set-ItemProperty $workWim -Name IsReadOnly -Value $false
        $workWimIndex = $WimIndex
        Write-OK "WIM copied"
    }
    Write-Log "WIM copied: $workWim (index $workWimIndex)"

    # Mount
    Write-Section "Mounting WIM"

    # Check for stale mount from a previous interrupted build and clean up
    $existingFiles = @(Get-ChildItem $mountDir -ErrorAction SilentlyContinue)
    if ($existingFiles.Count -gt 0) {
        Write-Warn "Mount directory is not empty - attempting to clean up stale mount..."
        try {
            Dismount-WindowsImage -Path $mountDir -Discard -ErrorAction Stop
            Write-OK "Stale mount dismounted (discarded)"
        } catch {
            Write-Warn "Could not dismount: $($_.Exception.Message)"
        }
        # Force-clean the directory regardless
        Remove-Item "$mountDir\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-OK "Mount directory cleared"
    }

    Mount-WindowsImage -ImagePath $workWim -Index $workWimIndex -Path $mountDir
    Write-OK "Mounted at: $mountDir"
    Write-Log "WIM mounted"

    # ═══════════════════════════════════════════════════════════════════════════
    # PIPELINE - Microsoft documented order
    # Source: learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update
    #
    # ORDER:
    #   1. Patch WinRE (winre.wim): SSU → SafeOS → Cleanup /ResetBase /Defer → Export
    #   2. Patch install.wim:       SSU → LP → FoD → Full LCU → Cleanup → NetFx3 → .NET CU
    # ═══════════════════════════════════════════════════════════════════════════

    $dotNetFiles = @()
    if (-not $SkipUpdates -and $resolvedUpdatePath) {

        # Enumerate update files
        $allUpdateFiles = @(
            Get-ChildItem $resolvedUpdatePath -Filter "*.msu" -ErrorAction SilentlyContinue
            Get-ChildItem $resolvedUpdatePath -Filter "*.cab" -ErrorAction SilentlyContinue
        ) | Where-Object {
            # Microsoft LCU filenames use "-x64_" and "-arm64_" (dash before, underscore after).
            # The old patterns "_x64\." / "_arm64\." never matched real filenames, causing
            # ARM64 LCUs to pass the filter on x64 builds (alphabetically sorted first → selected).
            $_.Name -notmatch "-x64[_.]"   -or $Architecture -eq "x64"
        } | Where-Object {
            $_.Name -notmatch "-arm64[_.]" -or $Architecture -eq "arm64"
        } | Sort-Object Name

        $safeOSFiles = @($allUpdateFiles | Where-Object { $_.Name -match "SafeOS" })
        $dotNetFiles = @($allUpdateFiles | Where-Object { $_.Name -match "^2_DotNet" })

        # LCU file - original Microsoft filename (with hash in name), always .msu
        # Explicitly require .msu to prevent old-named SafeOS .cab files (e.g.
        # windows11.0-kb5083482-x64.cab) from matching and being selected as the LCU.
        $lcuFile = $allUpdateFiles |
                   Where-Object { $_.Extension -eq ".msu" -and $_.Name -match "^windows11\.0-kb\d+-" -and $_.Name -notmatch "kb5043080" } |
                   Select-Object -First 1

        # Checkpoint file - original Microsoft filename (with hash in name)
        $checkpointFile = $allUpdateFiles |
                          Where-Object { $_.Name -match "^windows11\.0-kb5043080" } |
                          Select-Object -First 1

        # All LCU-related MSUs (LCU + checkpoint) for passing to DISM temp folder
        $lcuAllFiles = @($lcuFile) + @($checkpointFile) | Where-Object { $_ }

        # ── PART 1: Patch WinRE ────────────────────────────────────────────────
        # Per Microsoft docs: WinRE is patched FIRST, before install.wim.
        # Sequence: SSU via LCU → SafeOS → Cleanup /ResetBase /Defer → Export

        if ($lcuFile -or $safeOSFiles) {
            Write-Section "Patching WinRE (winre.wim) [step 1 of pipeline]"

            $winreSource   = "$mountDir\Windows\System32\Recovery\winre.wim"
            $winreWork     = Join-Path $workRoot "winre_work.wim"
            $winreExport   = Join-Path $workRoot "winre_export.wim"
            $winreMountDir = Join-Path $workRoot "WinREMount"

            if (-not (Test-Path $winreSource)) {
                Write-Warn "winre.wim not found in mounted image - skipping WinRE patching"
                Write-Log "WARN: winre.wim not found" "WARN"
            } else {
                New-Item $winreMountDir -ItemType Directory -Force | Out-Null
                Copy-Item $winreSource $winreWork -Force
                Set-ItemProperty $winreWork -Name IsReadOnly -Value $false
                Write-OK "Copied winre.wim to working location"
                Write-Log "winre.wim copied: $winreWork"

                Mount-WindowsImage -ImagePath $winreWork -Index 1 -Path $winreMountDir
                Write-OK "winre.wim mounted at: $winreMountDir"
                Write-Log "winre.wim mounted"

                try {
                    # Step 1: SSU via LCU
                    # Per Microsoft: ignore 0x8007007e (known issue with combined CU)
                    if ($lcuFile) {
                        # Copy LCU + checkpoint to isolated temp folder - both needed, original filenames preserved
                        $winreLCUTemp = Join-Path $workRoot "LCU_winre_temp"
                        New-Item $winreLCUTemp -ItemType Directory -Force | Out-Null
                        foreach ($f in $lcuAllFiles) { Copy-Item $f.FullName $winreLCUTemp }
                        $winreLCUTarget = Join-Path $winreLCUTemp $lcuFile.Name

                        Write-Host "  >> SSU -> WinRE: $($lcuFile.Name)" -ForegroundColor White
                        Write-Info "     Applying SSU to WinRE - may take several minutes..."
                        try {
                            Add-WindowsPackage -PackagePath $winreLCUTarget -Path $winreMountDir -ScratchDirectory $scratchDir | Out-Null
                            Write-OK "SSU applied to WinRE"
                            Write-Log "WinRE SSU OK: $($lcuFile.Name)"
                        } catch {
                            $err = $_.Exception.Message
                            if ($err -match "0x8007007e") {
                                Write-Info "SSU: 0x8007007e - known issue with combined CU, ignoring per Microsoft docs"
                                Write-Log "WinRE SSU 0x8007007e - ignored (known issue)" "INFO"
                            } elseif ($err -match "0x800f081e") {
                                Write-Info "SSU already current in WinRE (0x800f081e) - skipping"
                                Write-Log "WinRE SSU not applicable (already current)" "INFO"
                            } else {
                                throw "SSU failed on WinRE: $err"
                            }
                        }
                        Remove-Item $winreLCUTemp -Recurse -Force -ErrorAction SilentlyContinue
                    }

                    # Step 6: SafeOS Dynamic Update
                    foreach ($safeOS in $safeOSFiles) {
                        Write-Host "  >> SafeOS -> WinRE: $($safeOS.Name)" -ForegroundColor White
                        try {
                            Add-WindowsPackage -PackagePath $safeOS.FullName -Path $winreMountDir -ScratchDirectory $scratchDir -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                            Write-OK "SafeOS Dynamic Update applied to WinRE"
                            Write-Log "WinRE SafeOS OK: $($safeOS.Name)"
                        } catch {
                            $err = $_.Exception.Message
                            if ($err -match "0x800f081e") {
                                Write-Info "SafeOS not applicable (already current) - skipping"
                                Write-Log "WinRE SafeOS 0x800f081e - skipped" "INFO"
                            } else {
                                Write-Warn "SafeOS failed: $err"
                                Write-Log "WARN WinRE SafeOS: $err" "WARN"
                            }
                        }
                    }

                    # Step 7: Cleanup /ResetBase /Defer
                    Write-Info "Cleaning up WinRE image (/ResetBase /Defer)..."
                    & dism /Image:$winreMountDir /Cleanup-Image /StartComponentCleanup /ResetBase /Defer | Out-Null
                    if ($LASTEXITCODE -eq 0) {
                        Write-OK "WinRE cleanup OK"
                        Write-Log "WinRE cleanup OK"
                    } else {
                        Write-Warn "WinRE cleanup exit $LASTEXITCODE (non-fatal, continuing)"
                        Write-Log "WARN: WinRE cleanup exit $LASTEXITCODE" "WARN"
                    }

                    # Dismount + save
                    Dismount-WindowsImage -Path $winreMountDir -Save
                    Write-OK "winre.wim saved"
                    Write-Log "winre.wim dismounted (Save)"

                } catch {
                    Write-Warn "Error patching WinRE: $($_.Exception.Message)"
                    Write-Warn "Discarding WinRE changes - main OS is not affected"
                    Write-Log "WARN: WinRE patch failed: $($_.Exception.Message)" "WARN"
                    try { Dismount-WindowsImage -Path $winreMountDir -Discard -ErrorAction SilentlyContinue } catch {}
                }

                # Step 8: Export winre.wim
                if (-not (Get-WindowsImage -Mounted | Where-Object { $_.Path -eq $winreMountDir })) {
                    Write-Info "Exporting winre.wim (reduces size)..."
                    Export-WindowsImage -SourceImagePath $winreWork -SourceIndex 1 `
                        -DestinationImagePath $winreExport -CompressionType maximum
                    Write-OK "winre.wim exported"
                    Write-Log "winre.wim exported: $winreExport"

                    # Copy patched winre.wim back into mounted install.wim
                    Set-ItemProperty $winreSource -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
                    Copy-Item $winreExport $winreSource -Force
                    Write-OK "Patched winre.wim written back into install.wim"
                    Write-Log "winre.wim written back to: $winreSource"

                    Remove-Item $winreWork   -Force -ErrorAction SilentlyContinue
                    Remove-Item $winreExport -Force -ErrorAction SilentlyContinue
                }

                Remove-Item $winreMountDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        } else {
            Write-Info "No LCU or SafeOS files found - skipping WinRE patching"
        }

        # ── PART 2: Patch install.wim ──────────────────────────────────────────
        # Per Microsoft docs sequence for install.wim:
        #   Step 9:  SSU via LCU
        #   Step 10: Language Pack
        #   Step 11: Features on Demand
        #   Step 13: Full LCU (second pass)
        #   Step 14: Cleanup /StartComponentCleanup (NO /ResetBase)
        #   Step 15: NetFx3 + .NET CU

        # Step 9: SSU via LCU (first pass)
        if ($lcuFile) {
            Write-Section "install.wim - Step 9: SSU via LCU (pass 1)"
            # Copy LCU + checkpoint to isolated temp folder with original filenames
            $lcuTempDir = Join-Path $workRoot "LCU_temp"
            New-Item $lcuTempDir -ItemType Directory -Force | Out-Null
            foreach ($f in $lcuAllFiles) { Copy-Item $f.FullName $lcuTempDir }
            $lcuTempTarget = Join-Path $lcuTempDir $lcuFile.Name

            Write-Host "  >> LCU: $($lcuFile.Name)" -ForegroundColor White
            Write-Info "     Applying LCU pass 1 - this takes ~25 minutes..."
            try {
                # Use dism.exe - Add-WindowsPackage fails with 0x800401e3 on KB5083769 Apr 2026
                $dismResult = & dism.exe /Image:$mountDir /Add-Package /PackagePath:$lcuTempTarget /ScratchDir:$scratchDir 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-OK "LCU pass 1 installed"
                    Write-Log "LCU pass 1 OK: $($lcuFile.Name)"
                } elseif ($LASTEXITCODE -eq -2146498530) {
                    Write-Warn "LCU pass 1 not applicable (0x800f081e) - skipping"
                    Write-Log "WARN 0x800f081e LCU pass 1" "WARN"
                } else {
                    $errMsg = ($dismResult | Where-Object { $_ -match "Error" }) -join " "
                    Remove-Item $lcuTempDir -Recurse -Force -ErrorAction SilentlyContinue
                    throw "LCU pass 1 failed (exit $LASTEXITCODE): $errMsg"
                }
            } catch {
                Remove-Item $lcuTempDir -Recurse -Force -ErrorAction SilentlyContinue
                throw
            }
            Remove-Item $lcuTempDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Step 10: Language Packs
        if (-not $SkipLanguagePacks) {
            Write-Section "install.wim - Step 10: Language Packs"
            foreach ($lang in $ResolvedLocales) {
                Write-Host ""
                Write-Host "  >> $($lang.ToUpper())" -ForegroundColor White

                # LP cab - Add-WindowsPackage with file path (same as Microsoft)
                $lp = $allLPCabs | Where-Object { $_.Name -match "Language-Pack_${lpArchStr}_$lang" } | Select-Object -First 1
                if ($lp) {
                    try {
                        Add-WindowsPackage -PackagePath $lp.FullName -Path $mountDir -ScratchDirectory $scratchDir | Out-Null
                        Write-OK "LP: $($lp.Name)"
                        Write-Log "LP: $($lp.FullName)"
                    } catch {
                        $err = $_.Exception.Message
                        if ($err -match "0x800f0952") {
                            # Package already present in base image (e.g. en-GB in LTSC 2024 English International)
                            Write-Info "LP $lang already present in base image (0x800f0952) - skipping"
                            Write-Log "LP $lang skipped - already in base image (0x800f0952)" "INFO"
                        } elseif ($err -match "0x800f081e") {
                            Write-Info "LP $lang not applicable (0x800f081e) - skipping"
                            Write-Log "LP $lang skipped - not applicable (0x800f081e)" "INFO"
                        } else {
                            throw
                        }
                    }
                } else {
                    Write-Warn "LP not found for $lang"
                    Write-Log "WARN: LP not found for $lang" "WARN"
                }

                # Font support
                $FontTagMap = @{
                    'ar-SA'='Arab'; 'he-IL'='Hebr'; 'ja-JP'='Jpan'
                    'ko-KR'='Kore'; 'th-TH'='Thai'; 'zh-CN'='Hans'; 'zh-TW'='Hant'
                }
                if ($FontTagMap.ContainsKey($lang)) {
                    $tag = $FontTagMap[$lang]
                    Add-WindowsCapability -Name "Language.Fonts.$tag~~~und-$tag~0.0.1.0" `
                        -Path $mountDir -Source $LPSearchRoot -ScratchDirectory $scratchDir -ErrorAction SilentlyContinue | Out-Null
                    Write-OK "  Font capability [$tag]"
                    Write-Log "Font capability [$tag]: OK"
                }
            }

            # Offline registry: preserve installed language FODs
            Write-Info "Configuring offline registry to preserve language FODs..."
            try {
                $offlineSoftware = "$mountDir\Windows\System32\config\SOFTWARE"
                reg load "HKLM\WW_SOFTWARE" $offlineSoftware | Out-Null
                reg add "HKLM\WW_SOFTWARE\Policies\Microsoft\Control Panel\International" `
                    /v "BlockCleanupOfUnusedPreinstalledLangPacks" /t REG_DWORD /d 1 /f | Out-Null
                foreach ($locale in $ResolvedLocales) {
                    reg add "HKLM\WW_SOFTWARE\Microsoft\Windows NT\CurrentVersion\MUI\UILanguages\$locale" /f | Out-Null
                    Write-Log "MUI UILanguages registered: $locale"
                }
                [GC]::Collect(); Start-Sleep -Seconds 1
                reg unload "HKLM\WW_SOFTWARE" | Out-Null
                Write-OK "Offline registry: language cleanup suppressed for $($ResolvedLocales -join ', ')"
                Write-Log "Registry: BlockCleanupOfUnusedPreinstalledLangPacks=1, locales registered: $($ResolvedLocales -join ', ')"
            } catch {
                Write-Warn "Could not configure offline registry: $($_.Exception.Message)"
                try { reg unload "HKLM\WW_SOFTWARE" 2>$null | Out-Null } catch {}
            }
        }

        # Step 11: Features on Demand
        # Per Microsoft: Add-WindowsCapability with capability names, Source = FOD ISO path
        # Language FODs use capability names (Language.Basic~~~sv-SE~0.0.1.0 etc)
        # NOTE: NetFx3 goes to step 15 (after cleanup), not here
        Write-Section "install.wim - Step 11: Features on Demand"

        if (-not $SkipLanguagePacks) {
            foreach ($lang in $ResolvedLocales) {
                Write-Host "  >> Language FoDs: $($lang.ToUpper())" -ForegroundColor White
                foreach ($fodType in @("Basic","OCR","Handwriting","TextToSpeech","Speech")) {
                    $capName = "Language.$fodType~~~$lang~0.0.1.0"
                    try {
                        Add-WindowsCapability -Name $capName -Path $mountDir -Source $LPSearchRoot -ScratchDirectory $scratchDir | Out-Null
                        Write-OK "  $fodType"
                        Write-Log "Language FoD OK: $capName"
                    } catch {
                        $err = $_.Exception.Message
                        if ($err -match "0x800f081e" -or $err -match "not present\|not applicable") {
                            Write-Info "  $fodType not available for $lang - skipping"
                        } else {
                            Write-Warn "  $fodType failed: $err"
                            Write-Log "WARN Language FoD $capName`: $err" "WARN"
                        }
                    }
                }
            }
        }

        # RSAT and other non-NetFx3 FoDs
        if ($FoDList -ne "") {
            $FoDCatalog = @{
                'RsatAD'     = @{ Label = "RSAT: Active Directory"; Name = "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" }
                'RsatGPO'    = @{ Label = "RSAT: Group Policy Management"; Name = "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0" }
                'RsatSrvMgr' = @{ Label = "RSAT: Server Manager"; Name = "Rsat.ServerManager.Tools~~~~0.0.1.0" }
            }
            $fodKeys = $FoDList.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" -and $_ -ne "NetFx3" }
            foreach ($key in $fodKeys) {
                if (-not $FoDCatalog.ContainsKey($key)) { Write-Warn "Unknown FoD key: $key - skipping"; continue }
                $fod = $FoDCatalog[$key]
                Write-Host "  >> $($fod.Label)" -ForegroundColor White
                try {
                    Add-WindowsCapability -Name $fod.Name -Path $mountDir -Source $LPSearchRoot -ScratchDirectory $scratchDir | Out-Null
                    Write-OK "$($fod.Label) enabled"
                    Write-Log "FoD OK: $($fod.Label)"
                } catch {
                    $err = $_.Exception.Message
                    if ($err -match "0x800f081e") {
                        Write-Info "$($fod.Label) already present - skipping"
                    } else {
                        Write-Warn "$($fod.Label) failed: $err"
                        Write-Log "WARN FoD $($fod.Label): $err" "WARN"
                    }
                }
            }
        }

        # Step 13: Full LCU (second pass, after LP/FoD)
        if ($lcuFile) {
            Write-Section "install.wim - Step 13: Full LCU (pass 2, after LP/FoD)"
            $lcuTempDir2 = Join-Path $workRoot "LCU_temp2"
            New-Item $lcuTempDir2 -ItemType Directory -Force | Out-Null
            foreach ($f in $lcuAllFiles) { Copy-Item $f.FullName $lcuTempDir2 }
            $lcuTempTarget2 = Join-Path $lcuTempDir2 $lcuFile.Name

            Write-Host "  >> LCU: $($lcuFile.Name)" -ForegroundColor White
            Write-Info "     Applying full LCU - this takes ~25 minutes..."
            $dismResult = & dism.exe /Image:$mountDir /Add-Package /PackagePath:$lcuTempTarget2 /ScratchDir:$scratchDir 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-OK "Full LCU installed (pass 2)"
                Write-Log "LCU pass 2 OK: $($lcuFile.Name)"
            } elseif ($LASTEXITCODE -eq -2146498530) {
                Write-Warn "LCU pass 2 not applicable (0x800f081e) - skipping"
                Write-Log "WARN 0x800f081e LCU pass 2" "WARN"
            } else {
                $errMsg = ($dismResult | Where-Object { $_ -match "Error" }) -join " "
                Remove-Item $lcuTempDir2 -Recurse -Force -ErrorAction SilentlyContinue
                throw "LCU pass 2 failed (exit $LASTEXITCODE): $errMsg"
            }
            Remove-Item $lcuTempDir2 -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Step 14: Cleanup - /StartComponentCleanup only, NO /ResetBase on main OS
        # Per Microsoft sample script: only WinRE/WinPE get /ResetBase
        Write-Section "install.wim - Step 14: Component store cleanup"
        Write-Info "Running /StartComponentCleanup (no /ResetBase per Microsoft docs)..."
        if ($DebugBuild) {
            $cleanupResult = & dism.exe /Image:$mountDir /Cleanup-Image /StartComponentCleanup /ScratchDir:$scratchDir 2>&1
            $cleanupResult | ForEach-Object { Write-Host "    $_" -ForegroundColor Magenta; Write-Log "[DEBUG] $_" }
        } else {
            $cleanupResult = & dism.exe /Image:$mountDir /Cleanup-Image /StartComponentCleanup /ScratchDir:$scratchDir 2>&1
        }
        if ($LASTEXITCODE -eq 0) {
            Write-OK "Component cleanup complete"
            Write-Log "Component cleanup OK"
        } elseif ($LASTEXITCODE -eq -2146498554) {
            # 0x800F0806 CBS_E_PENDING - optional components require online operation
            Write-Warn "Cleanup: pending operations (0x800F0806) - image must boot to complete. Non-fatal, continuing."
            Write-Log "WARN: Component cleanup 0x800F0806 (pending ops)" "WARN"
        } else {
            Write-Warn "Component cleanup returned exit code $LASTEXITCODE (non-fatal, continuing)"
            Write-Log "WARN: Component cleanup exit $LASTEXITCODE" "WARN"
        }

        # Step 15: NetFx3 + .NET CU
        # NetFx3 via Add-WindowsCapability per Microsoft docs, AFTER cleanup
        $wantNetFx3 = $FoDList -ne "" -and ($FoDList.Split(',') | ForEach-Object { $_.Trim() }) -contains 'NetFx3'
        if ($wantNetFx3 -and $LPSearchRoot) {
            Write-Section "install.wim - Step 15a: NetFx3 (.NET Framework 3.5)"
            Write-Info "Installing via Add-WindowsCapability per Microsoft docs..."
            try {
                Add-WindowsCapability -Name "NetFX3~~~~" -Path $mountDir -Source $LPSearchRoot -ScratchDirectory $scratchDir | Out-Null
                Write-OK "NetFx3 installed"
                Write-Log "FoD OK: NetFx3 (step 15)"
            } catch {
                $err = $_.Exception.Message
                if ($err -match "0x800f081e" -or $err -match "already installed") {
                    Write-Info "NetFx3 already present - skipping"
                } else {
                    Write-Warn "NetFx3 failed: $err"
                    Write-Log "WARN FoD NetFx3: $err" "WARN"
                }
            }
        }

        if ($dotNetFiles -and $dotNetFiles.Count -gt 0) {
            Write-Section "install.wim - Step 15b: .NET CU"
            foreach ($upd in $dotNetFiles) {
                Write-Host "  >> .NET: $($upd.Name)" -ForegroundColor White
                try {
                    Add-WindowsPackage -PackagePath $upd.FullName -Path $mountDir -IgnoreCheck -ScratchDirectory $scratchDir | Out-Null
                    Write-OK "Installed"
                    Write-Log ".NET OK: $($upd.Name)"
                } catch {
                    $err = $_.Exception.Message
                    if ($err -match "0x800f081e") {
                        Write-Warn "Not applicable (0x800f081e) - skipping"
                        Write-Log "WARN 0x800f081e .NET" "WARN"
                    } else {
                        Write-Warn "Failed: $err"
                        Write-Log "ERROR .NET: $err" "ERROR"
                    }
                }
            }
        }

    } # end if (-not $SkipUpdates)

    # Appx removal
    if (-not $SkipAppxRemoval) {
        Write-Section "Removing Appx packages"
        $installed = Get-AppxProvisionedPackage -Path $mountDir
        $removed   = 0

        foreach ($appx in $AppxToRemove) {
            $pkg = $installed | Where-Object { $_.DisplayName -like "*$appx*" }
            if ($pkg) {
                try {
                    Remove-AppxProvisionedPackage -Path $mountDir -PackageName $pkg.PackageName | Out-Null
                    Write-OK "Removed: $($pkg.DisplayName)"
                    Write-Log "Appx removed: $($pkg.DisplayName)"
                    $removed++
                } catch {
                    Write-Warn "Could not remove: $($pkg.DisplayName)"
                    Write-Log "WARN Appx: $($pkg.DisplayName)" "WARN"
                }
            }
        }
        Write-Info "Removed: $removed of $($AppxToRemove.Count) candidates"
    }

    # ── RunOnce injection for language localization ───────────────────────────
    # Store-managed inbox apps (Notepad, Calculator, Paint etc.) display in English
    # at first logon even after LP injection, because their UI language comes from
    # a satellite MSIX downloaded by the Store framework at install time, not from
    # OS FODs. The fix is to uninstall and reinstall them via winget at first user
    # logon - winget triggers the AppX framework to fetch the correct language
    # satellites for the user's configured language.
    #
    # This section:
    # 1. Builds a list of apps that were KEPT (not in AppxToRemove)
    # 2. Generates InstallSystemApps.ps1 with winget reinstall commands for kept apps
    # 3. Copies the script into the WIM at C:\ProgramData\WimWizard\
    # 4. Injects a RunOnce key into the Default User profile hive so the script
    #    runs at first logon for every new user
    #
    # Skipped in patch mode / -SkipAppxRemoval: the app list is tied to Appx
    # removal decisions made at full-build time and must not be regenerated when
    # patching an existing image.
    if (-not $SkipAppxRemoval) {

    Write-Section "Injecting RunOnce language fix"

    # Full catalog: package name -> Store ID (null = no public winget/Store ID)
    $AppxStoreIds = @{
        "Microsoft.BingSearch"                   = "9NZBF4GT040C"
        "Microsoft.WindowsCalculator"            = "9WZDNCRFHVN5"
        "Microsoft.WindowsCamera"                = "9WZDNCRFJBBG"
        "Clipchamp.Clipchamp"                    = "9P1J8S7CCWWT"
        "Microsoft.Copilot"                      = "XP9CXNGPPJ97XX"
        "DevHome_8wekyb3d8bbwe"                  = "9N8MHTPHNGVV"
        "MicrosoftCorporationII.MicrosoftFamily" = "9PDJDJS743XF"
        "Microsoft.WindowsFeedbackHub"           = "9NBLGGH4R32N"
        "Microsoft.GamingApp"                    = "9NZKPSTSNW4P"
        "Microsoft.GetHelp"                      = "9PKDZBMV1H3T"
        "Microsoft.MicrosoftJournal"             = "9N318R854RHH"
        "Microsoft.Messaging"                    = $null   # Discontinued
        "Microsoft.BingWeather"                  = "9WZDNCRFJ3Q2"
        "Microsoft.BingNews"                     = "9WZDNCRFHVFW"
        "Microsoft.MicrosoftOfficeHub"           = $null   # No public Store ID
        "Microsoft.MicrosoftPCManager"           = "Microsoft.PCManager"  # winget repo ID
        "Microsoft.Windows.Photos"               = "9WZDNCRFJBH4"
        "Microsoft.MicrosoftSolitaireCollection" = "9NBLGGH4S79B"
        "Microsoft.MicrosoftStickyNotes"         = "9NBLGGH4QGHW"
        "Microsoft.ZuneMusic"                    = "9WZDNCRFJ3PT"
        "Microsoft.ZuneVideo"                    = "9WZDNCRFJ3P2"
        "MicrosoftTeams"                         = "9WZDNCRFJBMP"
        "Microsoft.Todos"                        = "9NBLGGH5R558"
        "Microsoft.WindowsNotepad"               = "9MSMLRH6LZF3"
        "Microsoft.OutlookForWindows"            = "9NRX63209R7B"
        "Microsoft.Paint"                        = "9PCFS5B6T72H"
        "MicrosoftCorporationII.QuickAssist"     = "9P7BP5VNWKX5"
        "Microsoft.ScreenSketch"                 = "9MZ95KL8MR0L"
        "Microsoft.WindowsSoundRecorder"         = "9WZDNCRFHWKN"
        "Microsoft.Whiteboard"                   = "9MSPC6MP8FM4"
        "Microsoft.XboxGamingOverlay"            = "9NZKPSTSNW4P"
        "Microsoft.XboxIdentityProvider"         = $null   # System component
        "Microsoft.XboxSpeechToTextOverlay"      = $null   # System component
        "Microsoft.Xbox.TCUI"                    = $null   # System component
        "Microsoft.YourPhone"                    = "9NMPJ99VJBWV"
        "Microsoft.WindowsTerminal"              = "9N0DX20HK701"
    }

    # Determine kept apps: full catalog minus removal list, with a known Store ID
    $keptApps = $AppxStoreIds.GetEnumerator() | Where-Object {
        $entryKey   = $_.Key
        $entryValue = $_.Value
        $entryValue -ne $null -and
        -not ($AppxToRemove | Where-Object { $entryKey -like "*$_*" -or $_ -like "*$entryKey*" })
    } | Sort-Object Key

    Write-Info "Apps kept in image with Store IDs: $($keptApps.Count)"
    foreach ($a in $keptApps) { Write-Log "  Kept app: $($a.Key) -> $($a.Value)" }

    # Build InstallSystemApps.ps1 content
    $scriptLines = @()
    $scriptLines += '# InstallSystemApps.ps1 - Generated by WimWizard'
    $scriptLines += '# Runs at first user logon via RunOnce.'
    $scriptLines += '# Reinstalls kept inbox apps so the Store framework downloads'
    $scriptLines += '# correct language satellites for the user''s configured language.'
    $scriptLines += "# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm') by WimWizard v$ScriptVersion"
    $scriptLines += ''
    $scriptLines += '$log = "C:\ProgramData\WimWizard\InstallSystemApps.log"'
    $scriptLines += 'function Write-Log { param([string]$msg) $ts = Get-Date -Format "HH:mm:ss"; Add-Content $log "[$ts] $msg" }'
    $scriptLines += ''
    $scriptLines += 'Write-Log "=== InstallSystemApps starting ==="'
    $scriptLines += ''
    $scriptLines += '# Bootstrap winget by downloading latest DesktopAppInstaller from GitHub'
    $scriptLines += 'Write-Log "Bootstrapping winget..."'
    $scriptLines += 'try {'
    $scriptLines += '    $releases = Invoke-RestMethod "https://api.github.com/repos/microsoft/winget-cli/releases/latest" -ErrorAction Stop'
    $scriptLines += '    $msixUrl  = ($releases.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1).browser_download_url'
    $scriptLines += '    $msixPath = "$env:TEMP\DesktopAppInstaller.msixbundle"'
    $scriptLines += '    Invoke-WebRequest $msixUrl -OutFile $msixPath -UseBasicParsing -ErrorAction Stop'
    $scriptLines += '    Add-AppxPackage $msixPath -ErrorAction Stop'
    $scriptLines += '    Write-Log "winget bootstrapped OK"'
    $scriptLines += '} catch { Write-Log "winget bootstrap failed: $($_.Exception.Message)" }'
    $scriptLines += ''
    $scriptLines += '# Update winget sources'
    $scriptLines += 'Write-Log "Updating winget sources..."'
    $scriptLines += 'try { winget source update --disable-interactivity 2>&1 | Out-Null; Write-Log "Sources updated" } catch { Write-Log "Source update failed" }'
    $scriptLines += ''

    foreach ($app in $keptApps) {
        $id     = $app.Value
        $pkg    = $app.Key
        # PC Manager uses winget repo, others use msstore
        $source = if ($id -notmatch '^9[A-Z0-9]+$') { 'winget' } else { 'msstore' }
        $scriptLines += "# $pkg"
        $scriptLines += "Write-Log `"Installing $pkg ($id)...`""
        $scriptLines += 'try {'
        $scriptLines += "    winget uninstall `"$id`" --silent --accept-source-agreements 2>&1 | Out-Null"
        $scriptLines += "    winget install `"$id`" --source $source --accept-package-agreements --accept-source-agreements --force --silent 2>&1 | Out-Null"
        $scriptLines += "    Write-Log `"  OK: $pkg`""
        $scriptLines += '} catch { Write-Log "  FAIL: ' + $pkg + ': $($_.Exception.Message)" }'
        $scriptLines += ''
    }

    $scriptLines += 'Write-Log "=== InstallSystemApps complete ==="'

    $scriptContent = $scriptLines -join "`r`n"

    # Create target directory in WIM and write script
    $wimScriptDir = Join-Path $mountDir "ProgramData\WimWizard"
    if (-not (Test-Path $wimScriptDir)) {
        New-Item $wimScriptDir -ItemType Directory -Force | Out-Null
    }
    $wimScriptPath = Join-Path $wimScriptDir "InstallSystemApps.ps1"
    [System.IO.File]::WriteAllText($wimScriptPath, $scriptContent, [System.Text.Encoding]::UTF8)
    Write-OK "Script written: ProgramData\WimWizard\InstallSystemApps.ps1 ($($keptApps.Count) apps)"
    Write-Log "InstallSystemApps.ps1 written with $($keptApps.Count) apps"

    # Inject RunOnce key into Default User hive
    $defaultHive    = Join-Path $mountDir "Users\Default\NTUSER.DAT"
    $hiveMountPoint = "HKLM\WimWizardDefaultUser_$(Get-Random)"
    $runOnceKey     = "$hiveMountPoint\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    $runOnceCmd     = 'powershell.exe -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "C:\ProgramData\WimWizard\InstallSystemApps.ps1"'

    try {
        & reg load $hiveMountPoint $defaultHive 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "reg load failed with exit code $LASTEXITCODE" }
        Write-Log "Default user hive loaded at $hiveMountPoint"

        & reg add $runOnceKey /v "WimWizardInstallApps" /t REG_SZ /d $runOnceCmd /f 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "reg add failed with exit code $LASTEXITCODE" }
        Write-OK "RunOnce key injected into Default User profile"
        Write-Log "RunOnce key set: $runOnceCmd"
    } catch {
        Write-Warn "RunOnce injection failed: $($_.Exception.Message)"
        Write-Log "WARN: RunOnce injection failed: $($_.Exception.Message)" "WARN"
    } finally {
        # Always unload the hive - leaving it mounted blocks future access
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 1
        & reg unload $hiveMountPoint 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Default user hive unloaded"
        } else {
            Write-Warn "reg unload failed (exit $LASTEXITCODE) - hive may remain locked"
            Write-Log "WARN: reg unload failed exit $LASTEXITCODE" "WARN"
        }
    }

    } # end if (-not $SkipAppxRemoval) - RunOnce injection

    # ── Final component store cleanup with /ResetBase ────────────────────────
    # /ResetBase permanently marks superseded component baselines as reclaimable,
    # reducing WIM size. Skipped when FoDs are injected: FoDs (especially NetFx3)
    # leave "Install Pending" state in CBS that causes DISM to reject BOTH
    # /ResetBase and plain /StartComponentCleanup with 0x800f0806. Step B already
    # ran /StartComponentCleanup after the LCU, which is the meaningful cleanup pass.
    # Running it again after FoDs would always fail, so we skip the section entirely.
    if ($FoDList -ne "") {
        Write-Section "Final component store cleanup"
        Write-Info "Skipped - FoDs leave CBS pending operations that block offline cleanup."
        Write-Info "Component store was already cleaned at Step B (after LCU, before FoDs)."
        Write-Log "Final cleanup skipped (FoDs injected - CBS pending ops would block 0x800f0806)"
    } else {
        Write-Section "Final component store cleanup (/ResetBase)"
        Write-Info "Running DISM /Cleanup-Image /StartComponentCleanup /ResetBase..."
        Write-Info "This permanently reduces WIM size by clearing superseded component baselines."
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 3
        if ($DebugBuild) {
            Write-Host ""
            Write-Host "  [DEBUG] Pending packages before /ResetBase:" -ForegroundColor Magenta
            Write-Log "[DEBUG] Pending packages before /ResetBase:"
            try {
                $pendingPkgs = Get-WindowsPackage -Path $mountDir |
                               Where-Object { $_.PackageState -eq 'InstallPending' -or $_.PackageState -eq 'UninstallPending' }
                if ($pendingPkgs) {
                    $pendingPkgs | ForEach-Object {
                        $line = "    $($_.PackageName)  [$($_.PackageState)]"
                        Write-Host $line -ForegroundColor Magenta
                        Write-Log "[DEBUG] $line"
                    }
                } else {
                    Write-Host "    (none)" -ForegroundColor Magenta
                    Write-Log "[DEBUG]   (no pending packages)"
                }
            } catch {
                Write-Host "    (failed to enumerate: $($_.Exception.Message))" -ForegroundColor Magenta
                Write-Log "[DEBUG]   pending pkg enum failed: $($_.Exception.Message)"
            }

            Write-Host ""
            Write-Host "  [DEBUG] Running /StartComponentCleanup /ResetBase with full output:" -ForegroundColor Magenta
            Write-Log "[DEBUG] Running /StartComponentCleanup /ResetBase:"
            $resetBaseResult = & dism.exe /Image:$mountDir /Cleanup-Image /StartComponentCleanup /ResetBase `
                                          /ScratchDir:$scratchDir 2>&1
            $resetBaseResult | ForEach-Object {
                Write-Host "    $_" -ForegroundColor Magenta
                Write-Log "[DEBUG] $_"
            }
        } else {
            $resetBaseResult = & dism.exe /Image:$mountDir /Cleanup-Image /StartComponentCleanup /ResetBase `
                                          /ScratchDir:$scratchDir 2>&1
        }
        if ($LASTEXITCODE -eq 0) {
            Write-OK "ResetBase complete"
            Write-Log "ResetBase OK"
        } else {
            Write-Warn "ResetBase returned exit code $LASTEXITCODE (non-fatal, continuing)"
            Write-Log "WARN: ResetBase exit $LASTEXITCODE" "WARN"
        }
    }

    # Dismount and save
    Write-Section "Saving and dismounting WIM"
    Write-Info "Writing changes back to WIM file..."
    Dismount-WindowsImage -Path $mountDir -Save -ScratchDirectory $scratchDir
    Write-OK "WIM saved"
    Write-Log "WIM dismounted (Save)"

    # Export and compress
    Write-Section "Exporting and compressing WIM"
    Write-Info "Compressing with 'maximum' - reduces file size significantly..."

    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

    Export-WindowsImage `
        -SourceImagePath      $workWim `
        -SourceIndex          $workWimIndex `
        -DestinationImagePath $OutputPath `
        -CompressionType      maximum `
        -ScratchDirectory     $scratchDir

    $sizeMB = [math]::Round((Get-Item $OutputPath).Length / 1MB, 0)
    Write-OK "Export complete: $OutputPath ($sizeMB MB)"
    Write-Log "Export OK: $OutputPath ($sizeMB MB)"

    # Clean up
    Remove-Item $workRoot -Recurse -Force -ErrorAction SilentlyContinue
    Write-OK "Temporary files cleaned up"

    # Dismount all source ISOs
    Dismount-AllISOs
    Write-OK "Source ISOs dismounted"

    # Done!
    $w = "=" * 66
    Write-Host ""
    Write-Host "+$w+" -ForegroundColor Green
    Write-Host ("|  [OK] DONE! The WIM file is ready to import into SCCM/MECM{0}|" -f (" " * 8)) -ForegroundColor Green
    Write-Host "+$w+" -ForegroundColor Green
    Write-Host ""
    Write-Host "  WIM file : $OutputPath  ($sizeMB MB)" -ForegroundColor White
    Write-Host "  Log file : $logFile" -ForegroundColor White
    Write-Host ""
    Write-Host "  Next steps in SCCM/MECM:" -ForegroundColor Yellow
    Write-Host "    1. Copy install.wim to your OS Image source folder" -ForegroundColor White
    Write-Host "    2. Software Library -> OS Images -> right-click -> Update Distribution Points" -ForegroundColor White
    Write-Host ""
    Write-Log "=== DONE. $OutputPath ($sizeMB MB) ==="

} catch {
    $errMsg = $_.Exception.Message
    Write-Host ""
    Write-Fail "AN ERROR OCCURRED:"
    Write-Host "  $errMsg" -ForegroundColor Red
    Write-Log "FATAL: $errMsg" "ERROR"
    Write-Host ""
    Write-Host "  Attempting to dismount without saving..." -ForegroundColor Yellow

    try {
        $mounted = Get-WindowsImage -Mounted -ErrorAction SilentlyContinue |
                   Where-Object { $_.Path -eq $mountDir }
        if ($mounted) {
            Dismount-WindowsImage -Path $mountDir -Discard
            Write-OK "WIM dismounted (changes discarded)"
            Write-Log "WIM dismounted with Discard" "ERROR"
        }
    } catch {
        Write-Fail "Could not dismount automatically - run manually:"
        Write-Host "    dism.exe /Unmount-Image /MountDir:`"$mountDir`" /Discard" -ForegroundColor White
        Write-Host "    dism.exe /Cleanup-Mountpoints" -ForegroundColor White
    }

    Dismount-AllISOs

    Write-Host ""
    Write-Host "  Log file: $logFile" -ForegroundColor White
    exit 1
}

#endregion