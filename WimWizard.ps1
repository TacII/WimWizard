#Requires -RunAsAdministrator
<#
.SYNOPSIS
    WIM Wizard - Windows 11 Image Servicing Tool for SCCM/MECM  v4.2

.DESCRIPTION
    Services a Windows 11 25H2/24H2 WIM file for distribution via SCCM/MECM.

    The user only needs to point to ONE folder containing the downloaded ISOs.
    The script automatically:
      - Finds and mounts the Windows ISO, extracts install.wim
      - Finds and mounts the Language Pack ISO (which also contains all FOD packages)
        and searches the LanguagesAndOptionalFeatures subfolder for all cab files
      - Auto-selects the Enterprise edition (or asks if not found)
      - Downloads the latest Patch Tuesday updates from Microsoft Update Catalog
      - Injects language packs and FOD for chosen languages (default: se,no,dk,fi)
      - Removes unnecessary provisioned Appx packages
      - Exports and compresses the finished WIM

.PARAMETER SourceFolder
    Folder containing the two ISOs downloaded from Microsoft:
      - Windows 11 Enterprise 25H2 ISO
      - Language Pack ISO (also contains all FOD packages)
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

.PARAMETER Unattended
    Answers yes to all prompts and uses defaults for all path inputs.
    Suitable for scheduled tasks and automation pipelines.
    All paths (SourceFolder, OutputPath) must either be provided as parameters
    or resolve correctly from defaults based on $PSScriptRoot.
    If Enterprise edition cannot be auto-detected, the script will fail with
    an error rather than hang - use -WimIndex to specify the index explicitly.

.NOTES
    Author      : Mathias Haas, Nobina IT
    Contact     : bWF0aGlhcy5oYWFzQGZpZGVsaXR5Y29uc3VsdGluZy5zZQ== (base64 - for urgent matters)
    License  : GNU General Public License v3.0 (GPL-3.0)
               https://www.gnu.org/licenses/gpl-3.0.html
    Product     : WIM Wizard (tribute to WIM Witch by Donna Ryan)
    Version     : 4.6.0
    Date        : 2026-04-01
    Requires    : Windows PowerShell 5.1+, Administrator rights, DISM
    Tested on   : Windows 11 25H2 (OS build 26200.x)

    CHANGELOG
    1.0.0  Initial release.
    1.1.0  Automatic update download from Microsoft Update Catalog.
    1.2.0  Translated all user-facing text and comments to English.
    1.3.0  Added direct Microsoft 365 Admin Center download link for Enterprise ISO.
    2.0.0  Single source folder, auto ISO discovery/mounting. Auto-selects Enterprise
           edition. -SourceFolder replaces -ISOPath / -LPPath / -FODPath.
    2.2.0  ISO auto-detection rewritten for actual Microsoft naming conventions.
           LP matched by 'LangPack', OS ISO by 'Win_Pro_11'.
    2.3.0  LP ISO contains both LP and FOD cabs - removed separate FOD ISO assumption.
           System FODs installed as language-tagged variants per language.
    2.4.0  Proper WinRE patching per Microsoft documented sequence: winre.wim
           extracted, mounted, patched with LCU + SafeOS, exported, written back.
    2.5.0  Persistent update cache in Updates\ folder keyed by KB number.
    2.6.0  All paths now derived from $PSScriptRoot - works on any drive letter.
    2.7.0  Update Catalog search terms derived dynamically from WIM build number.
    2.7.1  Fixed PropertyNotFoundException - Get-WindowsImage needs -Index for Version.
    2.7.2  Added TLS 1.2 and User-Agent to all Catalog requests.
    2.8.0  Replaced custom Catalog HTML scraping with MSCatalogLTS module.
    2.9.0  Fixed SafeOS search term. Removed neutral FOD crash (0x800f081e).
           Fixed ISO dismount on all exit paths (trap + Exit-Script helper).
    3.0.0  Fixed G:: double colon. Added -SetDefaultLanguage switch (off by default).
           Fixed LCU disk-full (0x80070070) with -ScratchDirectory on all DISM calls.
    3.1.0  SafeOS auto-download removed - MSCatalogLTS cannot find Dynamic Update
           packages. Script checks cache and shows manual instructions if missing.
    3.2.0  Added DISM component cleanup before dismount to reduce save time.
           Added -ScratchDirectory to Dismount-WindowsImage and Export-WindowsImage.
    3.2.1  Fixed stale version (2.0.0) in .SYNOPSIS and wrong filename reference.
    3.3.0  -Unattended fully wired up - Read-PathWithDefault uses defaults silently.
    3.4.0  Checkpoint cumulative update support: LCU applied from isolated temp
           folder so DISM auto-discovers prerequisites without .NET interference.
    3.5.0  Checkpoint prerequisites downloaded automatically via -DownloadAll.
           SafeOS now tries -IncludeDynamic short-term search before manual fallback.
    3.6.0  Fixed SafeOS auto-download using short-term search with -IncludeDynamic.
    3.6.1  Fixed "No LCU files found" - Get-ChildItem -Include silently ignores
           filter without -Recurse. Replaced with two separate -Filter calls.
    3.7.0  Fixed .Count crash under StrictMode (single Where-Object result).
           Output filename auto-generated from WIM build, version, languages, date.
    3.7.1  Fixed LCU 0x80073712 on cache hit: checkpoint prerequisite now fetched
           via -DownloadAll when no 0_Checkpoint_*.msu exists in Updates folder.
    3.8.0  Fixed empty KB number in checkpoint filename - MSCatalogLTS downloads
           with lowercase 'kb'; regex now case-insensitive with .ToUpper().
           WinRE LCU now uses same isolated-folder approach as main OS.
    3.8.1  Added -Optimize to Mount-WindowsImage calls.
    3.8.2  Removed -Optimize from main OS mount (cross-drive file access errors
           when adding language pack cabs from mounted ISO on separate drive).
    3.8.3  Removed -Optimize from WinRE mount - same issue as main OS mount.
    4.7.0  RunOnce language localization fix. After Appx removal, generates
           InstallSystemApps.ps1 containing winget reinstall commands for all
           apps that were KEPT in the image (not in removal list). Copies
           script into WIM at C:\ProgramData\WimWizard\. Injects RunOnce key
           into Default User NTUSER.DAT so script runs at every new user's
           first logon. Winget reinstall triggers AppX framework to download
           correct language satellites for the user's configured locale.
           winget is bootstrapped online at logon by downloading the latest
           DesktopAppInstaller release from GitHub.
    4.6.0  Expanded Appx removal list to 26 apps matching GUI defaults.
           Synced script defaults with GUI. Inbox Apps ISO approach removed
           (version mismatch with patched WIM causes app deletion).
    4.5.0  After LP/FOD injection, mounts offline SOFTWARE hive and sets
           BlockCleanupOfUnusedPreinstalledLangPacks=1 + registers each
           installed locale in MUI\UILanguages. Prevents Windows from removing
           app language FODs (Notepad, Calculator, Snipping Tool etc.) at first
           user logon. Installed languages are preserved; uninstalled ones can
           still be cleaned up by Windows.
    4.4.0  Fixed WinRE patching sequence to match Microsoft spec. Previously
           applied full LCU to WinRE (incorrect - causes 0x800f081e). Now correctly
           applies SSU via LCU package, then SafeOS Dynamic Update, then runs
           Cleanup-Image /StartComponentCleanup /ResetBase before dismount.
           Removed -IgnoreCheck from WinRE package installs. Final export
           changed from maximum to fast compression for SCCM compatibility.
    4.3.0  Patch mode: -PatchExistingWim parameter accepts path to an existing
           serviced WIM. Reads build number and installed languages directly from
           the WIM. Skips ISO discovery, LP injection and Appx removal. Output
           filename includes "_patched_" and current date. Compatible with GUI.
    4.2.0  Rebuilt $LocaleMap from actual LP ISO contents (295 cabs, 42 locales).
           Removed 20+ locales that don't exist in the ISO (de-AT, de-CH, fr-BE,
           nl-BE, en-AU, en-CA, bn-IN, hi-IN etc). Added missing locales:
           bg-BG, et-EE, eu-ES, gl-ES, hr-HR, lt-LT, lv-LV, ru-RU, sl-SI,
           sr-Latn-RS, uk-UA, vi-VN. Added country-code aliases (dk->da-DK,
           sv->sv-SE, gb->en-GB etc). Fixed LP cab filename search pattern:
           actual filename is "Language-Pack_x64_sv-se.cab" not "Language-Pack.*".
    4.1.3  Fixed $ResolvedLocales always empty: Resolve-LanguageCodes used
           $Input as parameter name which is a reserved PowerShell automatic
           variable (current pipeline object). Renamed to $CodeString.
    4.1.2  Fixed output filename showing "en" instead of selected languages.
           $anyMissing = $true was embedded inside a colour ternary expression,
           causing $lc to receive an array value instead of a string, which
           corrupted the validation loop and left $ResolvedLocales unreadable
           at Step 4. Side effect now separated from colour assignment.
    4.1.1  Fixed .Count PropertyNotFoundException under StrictMode when
           Resolve-LanguageCodes returns a single locale. All assignments
           now wrapped in @(). ResolvedLocales.Count guarded with @().
    4.1.0  Performance: LP/FOD folder scanned once and cached - avoids rescanning
           7800+ files dozens of times per language. Typical LP install time reduced
           significantly. $systemFODNames moved outside language loop. Merged double
           if (-not $SkipLanguagePacks) blocks into one. Fixed remaining -Include
           bugs in manual update path and summary count. Trimmed verbose comments.
           Fixed .PARAMETER help referencing stale FOD ISO / three-ISO model.
    4.0.0  Major: -Languages parameter replaces hardcoded Nordic locale list.
           Pass comma-separated 2-letter country codes (e.g. "se,no,dk,fi").
           Full locale mapping table covers all 53 Windows LP ISO languages.
           Interactive mode shows supported code list and prompts if omitted.
           -Unattended without -Languages skips LPs entirely (English only).
           Unknown codes show supported list and exit cleanly.
           Output filename uses country codes: Win11_25H2_..._se_no_dk_fi_....wim
           LCU progress message added so users know the script has not hung.

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
    [switch]$Unattended
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ScriptVersion = "4.7.0"

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
    Write-Host "  Patch existing WIM (updates only):"
    Write-Host "    .\WimWizard.ps1 -PatchExistingWim `"Output\Win11_25H2_..._20260409.wim`""
    Write-Host ""
    Write-Host "  PARAMETERS" -ForegroundColor White
    Write-Host "  ----------"
    Write-Host "  -SourceFolder     <path>   Folder with Windows ISO + LP ISO (default: .\ISO-Source\)"
    Write-Host "  -Languages        <codes>  Comma-separated language codes: da,fi,no,se,de,fr ..."
    Write-Host "  -OutputPath       <path>   Output WIM path (auto-generated if omitted)"
    Write-Host "  -PatchExistingWim <path>   Patch this WIM with latest updates (skips LP/Appx)"
    Write-Host "  -AppxListPath     <path>   XML app removal list (generated by GUI)"
    Write-Host "  -WimIndex         <int>    WIM index to service (default: auto-detect Enterprise)"
    Write-Host "  -SkipUpdates                Do not download or apply updates"
    Write-Host "  -SkipLanguagePacks          Skip language pack and FOD injection"
    Write-Host "  -SkipAppxRemoval            Skip removal of provisioned Appx packages"
    Write-Host "  -Unattended                 No interactive prompts (for GUI/automation use)"
    Write-Host "  -Help                       Show this help"
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
    'jp' = @('ja-JP')
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

    return @{
        LCU           = "Cumulative Update for Windows 11 Version $versionString for x64"
        DotNet        = "Cumulative Update for .NET Framework 3.5 and 4.8.1 for Windows 11, version $versionString for x64"
        SafeOS        = $safeOSSearch
        SafeOSVersion = $versionString   # used in manual download instructions
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
    Write-Host ("-" * 67) -ForegroundColor DarkCyan
    Write-Host "  >> $Title" -ForegroundColor White
    Write-Host ("-" * 67) -ForegroundColor DarkCyan
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
# Load custom list from XML if provided, otherwise use default
if ($AppxListPath -and (Test-Path $AppxListPath)) {
    try {
        [xml]$AppxXml  = Get-Content $AppxListPath -ErrorAction Stop
        $AppxToRemove  = @($AppxXml.AppxRemovalList.Package | ForEach-Object { $_.Id })
        Write-Info "Loaded custom Appx list from: $AppxListPath ($($AppxToRemove.Count) packages)"
    } catch {
        Write-Warn "Could not load AppxListPath: $($_.Exception.Message) - using default list"
        $AppxToRemove = $AppxToRemoveDefault
    }
} else {
    $AppxToRemove = $AppxToRemoveDefault
}

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
        try { Dismount-DiskImage -ImagePath $iso -ErrorAction SilentlyContinue } catch {}
    }
    $script:MountedISOs = @()
}

function Resolve-SourceFolder {
    param([string]$Folder)

    $isoFiles = Get-ChildItem $Folder -Filter "*.iso" -ErrorAction SilentlyContinue

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

        # ── Windows OS ISO ────────────────────────────────────────────────────
        # Explicitly excluded: anything already matched as LP above
        $isWindowsISO = (
            -not $isLPISO
        ) -and (
            $name -match "win_pro_11"        -or  # SW_DVD9_Win_Pro_11_25H2_* / _26H2_*
            $name -match "win.*11.*ent"      -or  # en-us_windows_11_enterprise_*
            $name -match "win.*11.*business"      # en-us_windows_11_business_*
        )

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
            $results = Get-MSCatalogUpdate -Search $CatalogSearchTerms[$type] -ErrorAction Stop
        } catch {
            Write-Warn "Search failed for $label`: $($_.Exception.Message)"
            continue
        }

        if (-not $results -or $results.Count -eq 0) {
            Write-Warn "No results found for $label."
            continue
        }

        # Filter out Preview builds and non-x64
        $filtered = $results | Where-Object {
            $_.Title -notmatch "\bPreview\b" -and
            $_.Title -notmatch "\bx86\b"     -and
            $_.Title -notmatch "\barm64\b"
        }
        if (-not $filtered) { $filtered = $results }

        # Latest = first result
        $best = $filtered | Select-Object -First 1
        Write-Host "  Found   : $($best.Title)" -ForegroundColor Green
        Write-Host "  Date    : $($best.LastUpdated)" -ForegroundColor DarkGray

        # Build canonical filename
        $kbMatch  = [regex]::Match($best.Title, "KB\d+")
        $kbNum    = if ($kbMatch.Success) { $kbMatch.Value } else { "KB_unknown" }
        $fileName = "${prefix}_${type}_${kbNum}.msu"
        $destPath = Join-Path $DownloadDir $fileName

        # Check if this KB is already in the cache folder
        $alreadyHave = Get-ChildItem $DownloadDir -Filter "*.msu" -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -match $kbNum -and $_.Length -gt 1MB } |
                       Select-Object -First 1

        if ($alreadyHave) {
            if ($alreadyHave.FullName -ne $destPath) {
                Rename-Item $alreadyHave.FullName $destPath -ErrorAction SilentlyContinue
                Write-OK "Found in cache (renamed to canonical): $fileName"
            } else {
                Write-OK "Already downloaded: $fileName ($([math]::Round($alreadyHave.Length/1MB,0)) MB)"
            }

            # For LCU: also check that checkpoint prerequisites are present.
            # If the LCU was cached before -DownloadAll was introduced, prerequisites
            # may be missing. Re-run Save-MSCatalogUpdate -DownloadAll to fetch them.
            if ($type -eq "LCU") {
                $hasCheckpoint = Get-ChildItem $DownloadDir -Filter "0_Checkpoint_*.msu" -ErrorAction SilentlyContinue |
                                 Select-Object -First 1
                if (-not $hasCheckpoint) {
                    Write-Info "     No checkpoint prerequisites found - checking if required..."
                    try {
                        Save-MSCatalogUpdate -Update $best -Destination $DownloadDir -DownloadAll -ErrorAction Stop
                        # Rename any newly downloaded checkpoint MSUs
                        $newFiles = Get-ChildItem $DownloadDir -Filter "*.msu" |
                                    Where-Object { $_.Name -notmatch "^[0-9]_" }
                        foreach ($newFile in $newFiles) {
                            $fileKB = ([regex]::Match($newFile.Name, 'KB\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Value.ToUpper()
                            if ($fileKB -ne $kbNum) {
                                $cpDest = Join-Path $DownloadDir "0_Checkpoint_${fileKB}.msu"
                                if (-not (Test-Path $cpDest)) {
                                    Rename-Item $newFile.FullName $cpDest -Force -ErrorAction SilentlyContinue
                                    Write-OK "Checkpoint prerequisite downloaded: 0_Checkpoint_${fileKB}.msu"
                                }
                            } else {
                                # Remove duplicate of already-cached LCU
                                if ($newFile.FullName -ne $destPath) {
                                    Remove-Item $newFile.FullName -Force -ErrorAction SilentlyContinue
                                }
                            }
                        }
                    } catch {
                        Write-Info "     No checkpoint prerequisites needed or download skipped."
                    }
                }
            }

            $downloaded++
            continue
        }

        Write-Info "     Downloading $fileName (this may take a while)..."
        try {
            # -DownloadAll ensures checkpoint prerequisite MSUs are downloaded
            # alongside the target LCU when the update is a post-checkpoint release.
            $dlArgs = @{ Update = $best; Destination = $DownloadDir; ErrorAction = 'Stop' }
            if ($type -eq "LCU") { $dlArgs['DownloadAll'] = $true }
            Save-MSCatalogUpdate @dlArgs

            # MSCatalogLTS saves with original Catalog filenames. Rename:
            # - Target LCU       → 1_LCU_KBxxxxxxx.msu
            # - Checkpoint MSUs  → 0_Checkpoint_KBxxxxxxx.msu (sort before target)
            $allNewMSUs = Get-ChildItem $DownloadDir -Filter "*.msu" |
                          Where-Object { $_.Name -notmatch "^[0-9]_" }

            foreach ($newFile in $allNewMSUs) {
                $fileKB = ([regex]::Match($newFile.Name, 'KB\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Value.ToUpper()
                if ($fileKB -eq $kbNum) {
                    if ($newFile.FullName -ne $destPath) {
                        Rename-Item $newFile.FullName $destPath -Force -ErrorAction SilentlyContinue
                    }
                } elseif ($type -eq "LCU") {
                    $cpDest = Join-Path $DownloadDir "0_Checkpoint_${fileKB}.msu"
                    if (-not (Test-Path $cpDest)) {
                        Rename-Item $newFile.FullName $cpDest -Force -ErrorAction SilentlyContinue
                        Write-OK "Checkpoint prerequisite: 0_Checkpoint_${fileKB}.msu"
                    }
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

    $safeOSExisting = Get-ChildItem $DownloadDir -Filter "3_SafeOS_*.cab" -ErrorAction SilentlyContinue |
                      Select-Object -First 1
    if (-not $safeOSExisting) {
        $safeOSExisting = Get-ChildItem $DownloadDir -Filter "*.cab" -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -match "SafeOS" -and $_.Length -gt 1MB } |
                          Select-Object -First 1
    }

    if ($safeOSExisting) {
        Write-OK "Found SafeOS in cache: $($safeOSExisting.Name)"
        $canonical = Join-Path $DownloadDir "3_SafeOS_$(([regex]::Match($safeOSExisting.Name,'KB\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Value.ToUpper()).cab"
        if ($safeOSExisting.FullName -ne $canonical -and [regex]::IsMatch($safeOSExisting.Name, 'KB\d+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            Rename-Item $safeOSExisting.FullName $canonical -ErrorAction SilentlyContinue
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
                                  $_.Title -match "x64"     -and
                                  $_.Title -match $versionStr -and
                                  $_.Title -notmatch "\bPreview\b" -and
                                  $_.Title -notmatch "\barm64\b"
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
            $safeDest = Join-Path $DownloadDir "3_SafeOS_${safeKB}.cab"
            try {
                Save-MSCatalogUpdate -Update $safeOSResult -Destination $DownloadDir -ErrorAction Stop
                # Rename to canonical
                $savedSafe = Get-ChildItem $DownloadDir -Filter "*.cab" |
                             Where-Object { $_.Name -match $safeKB -and $_.Name -ne "3_SafeOS_${safeKB}.cab" } |
                             Select-Object -First 1
                if ($savedSafe) { Rename-Item $savedSafe.FullName $safeDest -Force }
                if (Test-Path $safeDest) {
                    $sizeMB = [math]::Round((Get-Item $safeDest).Length / 1MB, 0)
                    Write-OK "Downloaded: 3_SafeOS_${safeKB}.cab ($sizeMB MB)"
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
            Write-Host "  3. Download the .cab file for x64" -ForegroundColor White
            Write-Host "  4. Save it to: $DownloadDir" -ForegroundColor White
            Write-Host "  5. Name it:    3_SafeOS_KBxxxxxxx.cab" -ForegroundColor White
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
Write-Host "  Place these TWO ISO files in a single folder, e.g. $($ScriptRoot)\ISO-Source\" -ForegroundColor White
Write-Host ""
Write-Host "  [1] Windows 11 Enterprise 25H2 ISO" -ForegroundColor Yellow
Write-Host "      Filename like: SW_DVD9_Win_Pro_11_25H2_64BIT_English_Pro_Ent_EDU_N_MLF_*.ISO" -ForegroundColor DarkGray
Write-Host "      (Future versions will follow same pattern: Win_Pro_11_26H2_...)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  [2] Language Pack ISO (also contains all FOD packages - no separate FOD ISO needed)" -ForegroundColor Yellow
Write-Host "      Filename like: SW_DVD9_Win_11_24H2_25H2_x64_Multilang_LangPack_All_LIP_LoF_*.ISO" -ForegroundColor DarkGray
Write-Host "      (Future versions will follow same pattern: Win_11_26H2_..._LangPack_...)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Download both ISOs from the Microsoft 365 Admin Center:" -ForegroundColor White
Write-Host "  https://admin.microsoft.com/adminportal/home#/subscriptions/vlnew/downloadsandkeys" -ForegroundColor Cyan
Write-Host "  Sign in -> Downloads & Keys -> Windows 11 Enterprise -> 25H2" -ForegroundColor DarkGray
Write-Host ""

if (-not (Test-Path $SourceFolder)) {
    # Folder doesn't exist - prompt the user
    $SourceFolder = Read-PathWithDefault `
        -Prompt  "Path to folder containing the two ISO files:" `
        -Default "$ScriptRoot\ISO-Source" `
        -MustExist
    if (-not $SourceFolder) { exit 1 }
}

# ── Discover and mount ISOs ───────────────────────────────────────────────────

Write-Info "Scanning $SourceFolder for ISO files..."
$discovered = Resolve-SourceFolder -Folder $SourceFolder

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

    # If LP ISO was not found, note it - the Step 2 pre-flight will handle the error
    # with a proper message. Don't silently fall back to searching the whole source folder
    # as that would give confusing errors instead of a clear "LP ISO missing" message.
    if (-not $LPSearchRoot) {
        Write-Warn "No Language Pack ISO found in $SourceFolder"
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
        Write-Host "  Make sure you have placed both ISO files in the folder." -ForegroundColor Yellow
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

} # end if -not PatchMode

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
        $ResolvedLocales = @(Resolve-LanguageCodes -CodeString $langInput)
    }

    if (-not $SkipLanguagePacks) {
        Write-Section "Step 2/4 - Language Packs + FOD"
        Write-Host "  Languages to install: $($ResolvedLocales -join ', ')" -ForegroundColor Cyan
        Write-Host ""

        # Scan LP/FOD folder ONCE and cache results - avoids rescanning 7800+ files per language.
        Write-Info "Scanning language pack folder (this may take a moment)..."
        $allLPCabs = @(Get-ChildItem $LPSearchRoot -Filter "*.cab" -ErrorAction SilentlyContinue)
        Write-Info "Found $($allLPCabs.Count) cab files in LP folder."

        # Validate each language before installing
        $anyMissing = $false
        foreach ($lang in $ResolvedLocales) {
            $lp  = $allLPCabs | Where-Object { $_.Name -match "Language-Pack_x64_$lang" } | Select-Object -First 1
            $fod = $allLPCabs | Where-Object { $_.Name -match "LanguageFeatures-Basic-$lang" } | Select-Object -First 1
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
        Write-Host "  Search for the following for Windows 11 Version 24H2 x64 (also applies to 25H2):" -ForegroundColor White
        Write-Host "    1_LCU_kbXXXXXXX.msu  - 'Cumulative Update for Windows 11 Version 24H2 for x64'" -ForegroundColor Yellow
        Write-Host "    2_DotNet_kbXXXX.msu  - '.NET Framework 3.5 4.8.1 Windows 11 24H2 x64'" -ForegroundColor Yellow
        Write-Host "    3_SafeOS_kbXXXX.cab  - 'Safe OS Dynamic Update Windows 11 24H2'" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Name files with prefix 1_, 2_, 3_ to control installation order." -ForegroundColor DarkGray
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
        ($ResolvedLocales | ForEach-Object { $_.Split('-')[1].ToLower() }) -join "_"
    }
    $_autoName = "Win11_${_versionStr}_${_buildStr}_${_langStr}_${_dateStr}.wim"
}
$_autoPath   = Join-Path (Split-Path $OutputPath -Parent) $_autoName

if (-not $OutputPath -or $OutputPath -eq "$ScriptRoot\Output\install.wim") {
    $OutputPath = $_autoPath
}

$OutputPath = Read-PathWithDefault `
    -Prompt  "Path for the finished WIM file (including filename):" `
    -Default $OutputPath

$outputDir = Split-Path $OutputPath -Parent
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
    Write-Host "    LP source     : $LPSearchRoot" -ForegroundColor Cyan
    Write-Host "    FOD source    : $FODSearchRoot" -ForegroundColor Cyan
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

if (-not $SkipAppxRemoval) {
    Write-Host "    Appx removal  : $($AppxToRemove.Count) packages removed if present" -ForegroundColor Cyan
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
$logFile    = Join-Path $outputDir ("WIMServicing_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmm"))

@($workRoot, $mountDir, $scratchDir) | ForEach-Object {
    if (-not (Test-Path $_)) { New-Item $_ -ItemType Directory -Force | Out-Null }
}

function Write-Log {
    param([string]$msg, [string]$level = "INFO")
    Add-Content -Path $logFile -Value ("[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $level, $msg)
}

Write-Log "=== WIM Wizard v$ScriptVersion ==="
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
    Mount-WindowsImage -ImagePath $workWim -Index $workWimIndex -Path $mountDir
    Write-OK "Mounted at: $mountDir"
    Write-Log "WIM mounted"

    # Language packs
    if (-not $SkipLanguagePacks) {
        Write-Section "Installing language packs"

        # $allLPCabs was built during the Step 2 validation scan above.
        # Reuse it here - no need to rescan 7800+ files again.
        # System FOD names defined once outside the language loop.
        $systemFODNames = @(
            "Microsoft-Windows-MSPaint-FoD-Package"
            "Microsoft-Windows-MediaPlayer-Package"
            "Microsoft-Windows-Notepad-FoD-Package"
            "Microsoft-Windows-Notepad-System-FoD-Package"
            "Microsoft-Windows-PowerShell-ISE-FOD-Package"
            "Microsoft-Windows-SnippingTool-FoD-Package"
            "Microsoft-Windows-StepsRecorder-Package"
            "Microsoft-Windows-WinOcr-FOD-Package"
            "Microsoft-Windows-WirelessDisplay-FOD-Package"
        )

        foreach ($lang in $ResolvedLocales) {
            Write-Host ""
            Write-Host "  >> $($lang.ToUpper())" -ForegroundColor White

            # LP - use cached scan
            $lp = $allLPCabs | Where-Object { $_.Name -match "Language-Pack_x64_$lang" } | Select-Object -First 1
            if ($lp) {
                Add-WindowsPackage -PackagePath $lp.FullName -Path $mountDir -ScratchDirectory $scratchDir | Out-Null
                Write-OK "LP: $($lp.Name)"
                Write-Log "LP: $($lp.FullName)"
            } else {
                Write-Warn "LP not found for $lang"
                Write-Log "WARN: LP not found for $lang" "WARN"
            }

            # Language feature FODs - use cached scan
            foreach ($fodType in @("Basic","OCR","Handwriting","Speech","TextToSpeech")) {
                $fod = $allLPCabs | Where-Object { $_.Name -match "LanguageFeatures-$fodType-$lang" } | Select-Object -First 1
                if ($fod) {
                    Add-WindowsPackage -PackagePath $fod.FullName -Path $mountDir -ScratchDirectory $scratchDir | Out-Null
                    Write-OK "${fodType}: $($fod.Name)"
                    Write-Log "${fodType} FOD: $($fod.FullName)"
                }
            }

            # System FODs (language-tagged variants e.g. ~amd64~sv-SE~.cab) - use cached scan
            foreach ($sysFod in $systemFODNames) {
                $tagged = $allLPCabs |
                          Where-Object { $_.Name -match [regex]::Escape($sysFod) -and
                                         $_.Name -match "~amd64~$([regex]::Escape($lang))~" } |
                          Select-Object -First 1
                if ($tagged) {
                    Add-WindowsPackage -PackagePath $tagged.FullName -Path $mountDir -ScratchDirectory $scratchDir | Out-Null
                    Write-OK "  $sysFod [$lang]"
                    Write-Log "System FOD [$lang]: $($tagged.FullName)"
                }
            }
        } # end foreach lang

        # ── Block language pack cleanup for installed languages ────────────────
        # Without this, Windows scheduled tasks remove "unused" language FODs at
        # first user logon - causing Notepad, Calculator, Snipping Tool etc. to
        # revert to English even though we injected Swedish/Nordic FOD packages.
        #
        # Strategy:
        #   1. Set BlockCleanupOfUnusedPreinstalledLangPacks = 1 in offline registry
        #      so the LPRemove and Pre-staged app cleanup tasks are suppressed.
        #   2. Register each installed locale in the MUI UILanguages key so Windows
        #      considers them intentionally installed, not candidates for removal.
        #   3. We do NOT block cleanup of languages we didn't install - Windows can
        #      still clean up any other pre-staged languages.
        Write-Info "Configuring offline registry to preserve installed language FODs..."
        try {
            # Load offline SOFTWARE hive
            $offlineSoftware = "$mountDir\Windows\System32\config\SOFTWARE"
            reg load "HKLM\WW_SOFTWARE" $offlineSoftware | Out-Null

            # Block cleanup of pre-installed language packs
            reg add "HKLM\WW_SOFTWARE\Policies\Microsoft\Control Panel\International" `
                /v "BlockCleanupOfUnusedPreinstalledLangPacks" /t REG_DWORD /d 1 /f | Out-Null

            # Register each installed locale in the MUI language list
            foreach ($locale in $ResolvedLocales) {
                reg add "HKLM\WW_SOFTWARE\Microsoft\Windows NT\CurrentVersion\MUI\UILanguages\$locale" `
                    /f | Out-Null
                Write-Log "MUI UILanguages registered: $locale"
            }

            # Unload hive (must be done cleanly)
            [GC]::Collect()
            Start-Sleep -Seconds 1
            reg unload "HKLM\WW_SOFTWARE" | Out-Null
            Write-OK "Offline registry: language cleanup suppressed for $($ResolvedLocales -join ', ')"
            Write-Log "Registry: BlockCleanupOfUnusedPreinstalledLangPacks=1, locales registered: $($ResolvedLocales -join ', ')"
        } catch {
            Write-Warn "Could not configure offline registry for language cleanup: $($_.Exception.Message)"
            Write-Warn "Language FODs may be removed at first user logon - consider re-running the build."
            Write-Log "WARN: Offline registry language config failed: $($_.Exception.Message)" "WARN"
            # Attempt to unload hive if it was loaded
            try { reg unload "HKLM\WW_SOFTWARE" 2>$null | Out-Null } catch {}
        }


    }

    # Updates
    if (-not $SkipUpdates -and $resolvedUpdatePath) {

        # Separate update files by type.
        # SafeOS goes to WinRE only. LCU and .NET go to the main OS (LCU also to WinRE).
        # Note: Get-ChildItem -Include requires -Recurse to work. Use two separate calls instead.
        $allUpdateFiles = @(
            Get-ChildItem $resolvedUpdatePath -Filter "*.msu" -ErrorAction SilentlyContinue
            Get-ChildItem $resolvedUpdatePath -Filter "*.cab" -ErrorAction SilentlyContinue
        ) | Sort-Object Name
        $safeOSFiles = @($allUpdateFiles | Where-Object { $_.Name -match "SafeOS" })
        $lcuFiles    = @($allUpdateFiles | Where-Object { $_.Name -match "^1_LCU" -or $_.Name -match "^0_" })
        $dotNetFiles = @($allUpdateFiles | Where-Object { $_.Name -match "^2_DotNet" })

        # ── Step A: Apply LCU to main OS ─────────────────────────────────────
        # IMPORTANT: Checkpoint cumulative updates (24H2+) require ALL checkpoint
        # prerequisite MSUs to be in the same folder as the target MSU, with NO
        # other MSU files present. DISM discovers prerequisites automatically.
        # We copy LCU files to an isolated temp folder to satisfy this requirement.
        Write-Section "Installing updates (main OS)"

        if ($lcuFiles -and $lcuFiles.Count -gt 0) {
            $lcuTempDir = Join-Path $scratchDir "LCU_temp"
            New-Item $lcuTempDir -ItemType Directory -Force | Out-Null

            foreach ($f in $lcuFiles) {
                Copy-Item $f.FullName $lcuTempDir
            }

            # Target = the highest-numbered (latest) LCU MSU. Checkpoint prerequisites
            # (prefixed 0_) are in the same folder and discovered automatically by DISM.
            $lcuTarget = Get-ChildItem $lcuTempDir -Filter "1_LCU_*.msu" | Sort-Object Name | Select-Object -Last 1

            # Always show what's in the temp folder so failures are diagnosable
            $lcuTempContents = @(Get-ChildItem $lcuTempDir -Filter "*.msu")
            Write-Info "     LCU temp folder contents ($($lcuTempContents.Count) file(s)):"
            foreach ($f in $lcuTempContents) {
                Write-Host "       $($f.Name) ($([math]::Round($f.Length/1MB,0)) MB)" -ForegroundColor DarkGray
            }

            if ($lcuTarget) {
                Write-Host "  >> LCU: $($lcuTarget.Name)" -ForegroundColor White
                Write-Info "     Applying LCU - this takes ~25 minutes. The script has not hung, please wait..."
                try {
                    Add-WindowsPackage -PackagePath $lcuTarget.FullName -Path $mountDir -IgnoreCheck -ScratchDirectory $scratchDir | Out-Null
                    Write-OK "LCU installed"
                    Write-Log "LCU OK: $($lcuTarget.Name)"
                } catch {
                    $err = $_.Exception.Message
                    if ($err -match "0x800f081e") {
                        Write-Warn "LCU not applicable to this image (0x800f081e) - skipping"
                        Write-Log "WARN 0x800f081e LCU" "WARN"
                    } else {
                        Write-Warn "LCU failed: $err"
                        Write-Log "ERROR LCU: $err" "ERROR"
                    }
                }
            } else {
                Write-Warn "No 1_LCU_*.msu found in LCU temp folder."
            }

            Remove-Item $lcuTempDir -Recurse -Force -ErrorAction SilentlyContinue

        } else {
            Write-Info "No LCU files found - skipping."
        }

        # ── Step B: Apply .NET CU to main OS ─────────────────────────────────
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

        # ── Step C: Patch WinRE (winre.wim) ──────────────────────────────────
        # Microsoft documented sequence (learn.microsoft.com/windows/deployment/update/media-dynamic-update):
        #   1. Apply SSU (servicing stack, via LCU package) to winre.wim
        #      NOTE: Apply the LCU package to winre.wim - DISM extracts and applies only
        #      the servicing stack component. The full LCU does NOT apply to WinRE (0x800f081e).
        #   2. Apply SafeOS Dynamic Update to winre.wim
        #   3. Cleanup image (StartComponentCleanup /ResetBase)
        #   4. Dismount + save winre.wim
        #   5. Export winre.wim (reduces size)
        #   6. Copy patched winre.wim back into mounted install.wim

        if ($lcuFiles -or $safeOSFiles) {
            Write-Section "Patching WinRE (winre.wim)"

            $winreSource  = "$mountDir\Windows\System32\Recovery\winre.wim"
            $winreWork    = Join-Path $workRoot "winre_work.wim"
            $winreExport  = Join-Path $workRoot "winre_export.wim"
            $winreMountDir= Join-Path $workRoot "WinREMount"

            if (-not (Test-Path $winreSource)) {
                Write-Warn "winre.wim not found in mounted image - skipping WinRE patching"
                Write-Log "WARN: winre.wim not found at $winreSource" "WARN"
            } else {
                New-Item $winreMountDir -ItemType Directory -Force | Out-Null

                # 1. Copy winre.wim to working location (it is read-only inside the WIM)
                Copy-Item $winreSource $winreWork -Force
                Set-ItemProperty $winreWork -Name IsReadOnly -Value $false
                Write-OK "Copied winre.wim to working location"
                Write-Log "winre.wim copied: $winreWork"

                # 2. Mount winre.wim (always index 1)
                Mount-WindowsImage -ImagePath $winreWork -Index 1 -Path $winreMountDir
                Write-OK "winre.wim mounted at: $winreMountDir"
                Write-Log "winre.wim mounted"

                try {
                    # 3. Apply SSU to WinRE by pointing at the LCU folder.
                    # Per Microsoft spec step 1: use the combined cumulative update package.
                    # DISM applies only the servicing stack component to WinRE - the full
                    # LCU is not applicable (0x800f081e) and will be silently skipped by DISM.
                    # Use isolated temp folder so checkpoint MSUs are available for dependency resolution.
                    if ($lcuFiles -and $lcuFiles.Count -gt 0) {
                        $winreLCUTemp = Join-Path $scratchDir "LCU_winre_temp"
                        New-Item $winreLCUTemp -ItemType Directory -Force | Out-Null
                        foreach ($f in $lcuFiles) { Copy-Item $f.FullName $winreLCUTemp }

                        $winreLCUTarget = Get-ChildItem $winreLCUTemp -Filter "1_LCU_*.msu" |
                                          Sort-Object Name | Select-Object -Last 1

                        if ($winreLCUTarget) {
                            Write-Host "  >> SSU -> WinRE (via LCU package): $($winreLCUTarget.Name)" -ForegroundColor White
                            Write-Info "     Applying SSU to WinRE - may take several minutes..."
                            try {
                                Add-WindowsPackage -PackagePath $winreLCUTarget.FullName -Path $winreMountDir -ScratchDirectory $scratchDir | Out-Null
                                Write-OK "SSU applied to WinRE"
                                Write-Log "WinRE SSU OK: $($winreLCUTarget.Name)"
                            } catch {
                                $err = $_.Exception.Message
                                if ($err -match "0x800f081e") {
                                    Write-Info "SSU already current in WinRE (0x800f081e) - skipping"
                                    Write-Log "WinRE SSU not applicable (already current)" "INFO"
                                } else {
                                    Write-Warn "SSU failed on WinRE: $err"
                                    Write-Log "WARN WinRE SSU: $err" "WARN"
                                }
                            }
                        }
                        Remove-Item $winreLCUTemp -Recurse -Force -ErrorAction SilentlyContinue
                    }

                    # 4. Apply SafeOS Dynamic Update to WinRE (Microsoft spec step 2)
                    foreach ($safeOS in $safeOSFiles) {
                        Write-Host "  >> SafeOS -> WinRE: $($safeOS.Name)" -ForegroundColor White
                        try {
                            Add-WindowsPackage -PackagePath $safeOS.FullName -Path $winreMountDir -ScratchDirectory $scratchDir | Out-Null
                            Write-OK "SafeOS Dynamic Update applied to WinRE"
                            Write-Log "WinRE SafeOS OK: $($safeOS.Name)"
                        } catch {
                            $err = $_.Exception.Message
                            if ($err -match "0x800f081e") {
                                Write-Info "SafeOS already current in WinRE - skipping"
                                Write-Log "WinRE SafeOS not applicable (already current)" "INFO"
                            } else {
                                Write-Warn "SafeOS failed: $err"
                                Write-Log "WARN WinRE SafeOS: $err" "WARN"
                            }
                        }
                    }

                    # 5. Cleanup WinRE image (Microsoft spec step 3)
                    Write-Info "Cleaning up WinRE image..."
                    & dism /Image:$winreMountDir /Cleanup-Image /StartComponentCleanup /ResetBase | Out-Null
                    Write-OK "WinRE image cleaned up"
                    Write-Log "WinRE cleanup OK"

                    # 6. Dismount winre.wim and save
                    Dismount-WindowsImage -Path $winreMountDir -Save
                    Write-OK "winre.wim saved"
                    Write-Log "winre.wim dismounted (Save)"

                } catch {
                    # On any error, discard WinRE changes and continue - main OS is unaffected
                    Write-Warn "Error patching WinRE: $($_.Exception.Message)"
                    Write-Warn "Discarding WinRE changes - main OS image is not affected"
                    Write-Log "WARN: WinRE patch failed, discarding: $($_.Exception.Message)" "WARN"
                    try { Dismount-WindowsImage -Path $winreMountDir -Discard -ErrorAction SilentlyContinue } catch {}
                }

                # 7. Export winre.wim to reduce size (skip if dismount failed)
                if (-not (Get-WindowsImage -Mounted | Where-Object { $_.Path -eq $winreMountDir })) {
                    Write-Info "Exporting winre.wim (reduces size)..."
                    Export-WindowsImage -SourceImagePath $winreWork -SourceIndex 1 `
                        -DestinationImagePath $winreExport -CompressionType maximum
                    Write-OK "winre.wim exported"
                    Write-Log "winre.wim exported: $winreExport"

                    # 8. Copy patched winre.wim back into the mounted install.wim
                    Set-ItemProperty $winreSource -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
                    Copy-Item $winreExport $winreSource -Force
                    Write-OK "Patched winre.wim written back into install.wim"
                    Write-Log "winre.wim written back to: $winreSource"

                    # Clean up working files
                    Remove-Item $winreWork   -Force -ErrorAction SilentlyContinue
                    Remove-Item $winreExport -Force -ErrorAction SilentlyContinue
                }

                Remove-Item $winreMountDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        } else {
            Write-Info "No LCU or SafeOS files found - skipping WinRE patching"
        }
    }

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

    # Component cleanup before dismount.
    # Removes superseded components from the store - dramatically reduces the
    # amount of data DISM has to write back during dismount (can cut 50-70%).
    # /ResetBase removes all superseded versions; this is a one-way operation
    # but fine for a deployment image.
    Write-Section "Component store cleanup"
    Write-Info "Running DISM /Cleanup-Image /StartComponentCleanup /ResetBase..."
    Write-Info "This takes ~5 min but significantly reduces dismount time and final WIM size."
    $cleanupResult = & dism.exe /Image:$mountDir /Cleanup-Image /StartComponentCleanup /ResetBase `
                                /ScratchDir:$scratchDir 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-OK "Component cleanup complete"
        Write-Log "Component cleanup OK"
    } else {
        # Non-fatal - log and continue, dismount will still work
        Write-Warn "Component cleanup returned exit code $LASTEXITCODE (non-fatal, continuing)"
        Write-Log "WARN: Component cleanup exit $LASTEXITCODE" "WARN"
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
        -CompressionType      fast `
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
