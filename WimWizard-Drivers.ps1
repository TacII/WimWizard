<#
    Shared driver discovery, metadata, filtering and offline servicing helpers.
    This file intentionally contains no parameter block or Windows-only startup
    code so that the functions can be loaded by Pester without running a build.
#>

$script:WimWizardDriverClassCatalog = [ordered]@{
    Network = [ordered]@{
        Classes = @('Net')
        ClassGuids = @('{4D36E972-E325-11CE-BFC1-08002BE10318}')
    }
    Storage = [ordered]@{
        Classes = @('SCSIAdapter', 'HDC', 'NvmeDisk')
        ClassGuids = @(
            '{4D36E97B-E325-11CE-BFC1-08002BE10318}'
            '{4D36E96A-E325-11CE-BFC1-08002BE10318}'
            '{75416E63-5912-4DFA-AE8F-3EFACCAFFB14}'
        )
    }
}

function Assert-DriverSourceParameters {
    param(
        [string]$DriverPackageID = '',
        [string]$DriverPath = '',
        [switch]$AddStorageDriversToWinRE,
        [switch]$AddNetworkDriversToWinRE
    )
    $hasPackage = -not [string]::IsNullOrWhiteSpace($DriverPackageID)
    $hasPath = -not [string]::IsNullOrWhiteSpace($DriverPath)
    if ($hasPackage -and $hasPath) { throw '-DriverPackageID and -DriverPath cannot be used together.' }
    if (($AddStorageDriversToWinRE -or $AddNetworkDriversToWinRE) -and -not ($hasPackage -or $hasPath)) {
        throw 'WinRE driver integration requires a configured driver source.'
    }
    return [pscustomobject]@{ HasSource = ($hasPackage -or $hasPath); SourceType = if ($hasPackage) { 'MECM' } elseif ($hasPath) { 'Path' } else { 'None' } }
}

function Test-WinREServicingRequired {
    param(
        [object]$LcuFile,
        [object[]]$SafeOSFiles = @(),
        [object[]]$DriverInfFiles = @()
    )
    return [bool]($null -ne $LcuFile -or @($SafeOSFiles).Count -gt 0 -or @($DriverInfFiles).Count -gt 0)
}

function Invoke-DriverLog {
    param(
        [scriptblock]$LogAction,
        [string]$Message,
        [string]$Level = 'INFO'
    )
    if ($LogAction) { & $LogAction $Message $Level }
}

function Get-DriverFileText {
    param([Parameter(Mandatory)][string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        return [System.Text.Encoding]::BigEndianUnicode.GetString($bytes, 2, $bytes.Length - 2)
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }

    # UTF-8 without BOM is common for hand-authored/test INFs. If strict UTF-8
    # decoding fails, use the Windows ANSI code page used by PowerShell 5.1.
    try {
        $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
        return $strictUtf8.GetString($bytes)
    } catch {
        return [System.Text.Encoding]::Default.GetString($bytes)
    }
}

function Remove-InfInlineComment {
    param([string]$Value)
    $inQuotes = $false
    for ($i = 0; $i -lt $Value.Length; $i++) {
        if ($Value[$i] -eq '"') { $inQuotes = -not $inQuotes; continue }
        if ($Value[$i] -eq ';' -and -not $inQuotes) {
            return $Value.Substring(0, $i).Trim()
        }
    }
    return $Value.Trim()
}

function Convert-InfValue {
    param([string]$Value)
    $clean = (Remove-InfInlineComment -Value $Value).Trim()
    if ($clean.Length -ge 2 -and $clean[0] -eq '"' -and $clean[$clean.Length - 1] -eq '"') {
        $clean = $clean.Substring(1, $clean.Length - 2)
    }
    return $clean.Trim()
}

function Get-DriverCategory {
    param(
        [string]$Class,
        [string]$ClassGuid
    )

    $classValue = if ($Class) { $Class.Trim() } else { '' }
    $guidValue  = if ($ClassGuid) { $ClassGuid.Trim().ToUpperInvariant() } else { '' }

    foreach ($category in @('Network', 'Storage')) {
        $entry = $script:WimWizardDriverClassCatalog[$category]
        if ($entry.Classes -contains $classValue) { return $category }
    }

    # DCH dependency/component INF classes are intentionally not promoted to
    # Storage/Network solely because a package happens to carry a matching GUID.
    if ($classValue -match '^(Extension|SoftwareComponent|System)$') { return 'Other' }

    # ClassGuid is the fallback if Class is missing or not a recognized class.
    foreach ($category in @('Network', 'Storage')) {
        $entry = $script:WimWizardDriverClassCatalog[$category]
        if (($entry.ClassGuids | ForEach-Object { $_.ToUpperInvariant() }) -contains $guidValue) {
            return $category
        }
    }

    if (-not $classValue -and -not $guidValue) { return 'Unknown' }
    return 'Unknown'
}

function Get-DriverInfMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [scriptblock]$LogAction
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Driver INF file not found: $Path"
    }

    $class = ''
    $classGuid = ''
    $provider = ''
    $driverVer = ''
    $inVersion = $false
    $versionFound = $false

    try {
        $text = Get-DriverFileText -Path $Path
        foreach ($line in ($text -split "`r?`n")) {
            $trimmed = $line.Trim()
            if (-not $trimmed -or $trimmed.StartsWith(';')) { continue }
            if ($trimmed -match '^\s*\[([^]]+)\]') {
                $inVersion = $Matches[1].Trim() -ieq 'Version'
                if ($inVersion) { $versionFound = $true }
                continue
            }
            if (-not $inVersion -or $trimmed -notmatch '^\s*(ClassGuid|Class|Provider|DriverVer)\s*=\s*(.*?)\s*$') {
                continue
            }
            $key = $Matches[1]
            $value = Convert-InfValue -Value $Matches[2]
            switch -Regex ($key) {
                '^Class$'      { $class = $value; break }
                '^ClassGuid$'  { $classGuid = $value; break }
                '^Provider$'   { $provider = $value; break }
                '^DriverVer$'  { $driverVer = $value; break }
            }
        }
    } catch {
        Invoke-DriverLog $LogAction "WARN: Could not parse INF '$Path': $($_.Exception.Message)" 'WARN'
        return [pscustomobject]@{
            FullPath = [System.IO.Path]::GetFullPath($Path); FileName = [System.IO.Path]::GetFileName($Path)
            SourceDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($Path))
            Class = ''; ClassGuid = ''; Provider = ''; DriverVer = ''; Category = 'Unknown'
            VersionSectionFound = $false; IsValid = $false; ParseError = $_.Exception.Message
        }
    }

    if (-not $class -and -not $classGuid) {
        Invoke-DriverLog $LogAction "WARN: INF '$Path' contains no usable Class or ClassGuid in [Version]; classified as Unknown." 'WARN'
    }

    return [pscustomobject]@{
        FullPath = [System.IO.Path]::GetFullPath($Path)
        FileName = [System.IO.Path]::GetFileName($Path)
        SourceDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($Path))
        Class = $class
        ClassGuid = $classGuid
        Provider = $provider
        DriverVer = $driverVer
        Category = Get-DriverCategory -Class $class -ClassGuid $classGuid
        VersionSectionFound = $versionFound
        IsValid = $versionFound
        ParseError = $null
    }
}

function Find-DriverInfFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [scriptblock]$LogAction
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Path -match '^[A-Za-z]:\\' -and -not (Test-Path -LiteralPath $Path.Substring(0, 2))) {
            throw "Driver path '$Path' is not available. The mapped drive may not be present in this elevated session; use a UNC path or make the drive available to the administrator session."
        }
        throw "Driver path not found: $Path"
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Driver source must be a directory: $Path"
    }

    try {
        $files = @(Get-ChildItem -LiteralPath $Path -Filter '*.inf' -File -Recurse -ErrorAction Stop | Sort-Object FullName)
    } catch {
        throw "Could not enumerate driver source '$Path': $($_.Exception.Message)"
    }
    Invoke-DriverLog $LogAction "Found $($files.Count) INF file(s) under '$Path'." 'INFO'
    return $files
}

function Select-DriverInfFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Metadata,
        [ValidateSet('All', 'StorageAndNetwork')][string]$DriverFilter = 'All',
        [scriptblock]$LogAction
    )

    $valid = @($Metadata | Where-Object { $_.IsValid -ne $false })
    $selected = @()
    $skipped = @()
    if ($DriverFilter -eq 'All') {
        $selected = $valid
    } else {
        $selected = @($valid | Where-Object { $_.Category -in @('Storage', 'Network') })
        $skipped = @($valid | Where-Object { $_.Category -notin @('Storage', 'Network') })
    }

    $skipByCategory = @($skipped | Group-Object Category | ForEach-Object { "$($_.Name): $($_.Count)" })
    if ($skipped.Count -gt 0) {
        Invoke-DriverLog $LogAction "Skipped by DriverFilter '$DriverFilter': $($skipByCategory -join ', ')." 'INFO'
        $skipped | Where-Object { $_.Category -in @('Other', 'Unknown') } | ForEach-Object {
            if ($_.Class -match '^(Extension|SoftwareComponent|System)$' -or $_.Category -eq 'Other') {
                Invoke-DriverLog $LogAction "Skipped dependency/component INF: $($_.FullPath) [$($_.Class)]" 'INFO'
            }
        }
    }

    return [pscustomobject]@{
        Found = $Metadata.Count
        Valid = $valid.Count
        Selected = @($selected)
        Skipped = @($skipped)
        Storage = @($selected | Where-Object { $_.Category -eq 'Storage' })
        Network = @($selected | Where-Object { $_.Category -eq 'Network' })
        Other = @($selected | Where-Object { $_.Category -eq 'Other' })
        Unknown = @($selected | Where-Object { $_.Category -eq 'Unknown' })
        SkippedByCategory = $skipByCategory
        Filter = $DriverFilter
    }
}

function Resolve-CMDriverPackageSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DriverPackageID,
        [Parameter(Mandatory)][string]$SCCMServer,
        [Parameter(Mandatory)][string]$SCCMSiteCode,
        [string]$ModulePath,
        [scriptblock]$LogAction
    )

    if (-not $ModulePath) {
        if (-not $env:SMS_ADMIN_UI_PATH) { throw 'The Configuration Manager console is not installed (SMS_ADMIN_UI_PATH is missing).' }
        $ModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
    }
    if (-not (Test-Path -LiteralPath $ModulePath -PathType Leaf)) {
        throw "ConfigurationManager.psd1 not found: $ModulePath"
    }
    if ([string]::IsNullOrWhiteSpace($SCCMServer) -or [string]::IsNullOrWhiteSpace($SCCMSiteCode)) {
        throw 'Driver package resolution requires both -SCCMServer and -SCCMSiteCode.'
    }

    $previousLocation = Get-Location
    $siteDrive = $SCCMSiteCode.Trim().ToUpperInvariant()
    $locationPushed = $false
    try {
        Import-Module $ModulePath -ErrorAction Stop
        if (-not (Get-PSDrive -Name $siteDrive -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $siteDrive -PSProvider CMSite -Root $SCCMServer -ErrorAction Stop | Out-Null
        }
        Push-Location "$siteDrive`:" -ErrorAction Stop
        $locationPushed = $true
        # Do not use -Fast here: PkgSourcePath is a lazy property on some
        # Configuration Manager versions. The cmdlet warns when -Fast is not
        # used, but full properties are intentional for package resolution.
        $package = Get-CMDriverPackage -Id $DriverPackageID -WarningAction SilentlyContinue -ErrorAction Stop
        if (-not $package) { throw "MECM Driver Package '$DriverPackageID' was not found." }

        $typeProperty = $package.PSObject.Properties['PackageType']
        if ($typeProperty -and [string]$package.PackageType) {
            $packageTypeText = [string]$package.PackageType
            $packageTypeNumber = 0
            $isNumericPackageType = [int]::TryParse($packageTypeText, [ref]$packageTypeNumber)
            # SMS_PackageBaseClass uses PackageType 3 for Driver Packages.
            $isDriverPackage = ($packageTypeText -match 'driver') -or
                               ($isNumericPackageType -and $packageTypeNumber -eq 3)
            if (-not $isDriverPackage) {
                throw "MECM object '$DriverPackageID' is not a Driver Package (type: $($package.PackageType))."
            }
        }
        $sourcePath = [string]$package.PkgSourcePath
        if ([string]::IsNullOrWhiteSpace($sourcePath)) { throw "Driver Package '$DriverPackageID' has no PkgSourcePath." }

        # Validate the source through the normal FileSystem provider. The
        # package was retrieved from the CMSite provider above; validating an
        # UNC path while that provider is still current can hide the actual
        # SMB/access error on some Configuration Manager console versions.
        $packageName = [string]$package.Name
        if ($locationPushed) {
            Pop-Location -ErrorAction Stop
            $locationPushed = $false
        }
        try {
            $sourceItem = Get-Item -LiteralPath $sourcePath -ErrorAction Stop
        } catch {
            throw "PkgSourcePath is not reachable from this build host: $sourcePath ($($_.Exception.Message))"
        }
        if (-not $sourceItem.PSIsContainer) { throw "PkgSourcePath is not a directory: $sourcePath" }

        Invoke-DriverLog $LogAction "MECM Driver Package: $packageName [$DriverPackageID]" 'INFO'
        Invoke-DriverLog $LogAction "Resolved driver source path: $sourcePath" 'INFO'
        return [pscustomobject]@{
            SourceType = 'MECM Driver Package'; PackageName = $packageName
            PackageID = $DriverPackageID; SourcePath = $sourcePath; Package = $package
        }
    } finally {
        if ($locationPushed) { Pop-Location -ErrorAction SilentlyContinue }
        Set-Location -LiteralPath $previousLocation.Path -ErrorAction SilentlyContinue
    }
}

function Add-DriversToOfflineImage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MountPath,
        [Parameter(Mandatory)][object[]]$DriverInfFiles,
        [Parameter(Mandatory)][string]$TargetLabel,
        [int]$FoundCount = 0,
        [int]$SkippedCount = 0,
        [scriptblock]$LogAction,
        [switch]$ThrowOnFailure
    )

    $result = [pscustomobject]@{
        Target = $TargetLabel; Found = $FoundCount; Selected = $DriverInfFiles.Count
        Successful = @(); Failed = @(); Skipped = @()
        SuccessCount = 0; FailedCount = 0; SkippedCount = $SkippedCount
    }
    foreach ($inf in $DriverInfFiles) {
        $path = if ($inf.FullPath) { [string]$inf.FullPath } else { [string]$inf }
        $class = if ($inf.Category) { "$($inf.Category) ($($inf.Class))" } else { 'Unknown' }
        Invoke-DriverLog $LogAction "Driver integration started [$TargetLabel]: $path [$class]" 'INFO'
        try {
            Add-WindowsDriver -Path $MountPath -Driver $path -ErrorAction Stop | Out-Null
            $result.Successful += $inf
            $result.SuccessCount++
            Invoke-DriverLog $LogAction "Driver integration succeeded [$TargetLabel]: $path" 'INFO'
        } catch {
            $result.Failed += [pscustomobject]@{ Inf = $inf; Error = $_.Exception.Message }
            $result.FailedCount++
            Invoke-DriverLog $LogAction "Driver integration failed [$TargetLabel]: $path - $($_.Exception.Message)" 'ERROR'
        }
    }
    Invoke-DriverLog $LogAction "Driver summary [$TargetLabel]: found=$($result.Found), selected=$($result.Selected), successful=$($result.SuccessCount), skipped=$($result.SkippedCount), failed=$($result.FailedCount)" 'INFO'
    if ($ThrowOnFailure -and $result.FailedCount -gt 0) {
        throw "Driver integration into $TargetLabel failed for $($result.FailedCount) INF file(s). The image will not be completed."
    }
    return $result
}

function Invoke-WinREDriverAndUpdateServicing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$InstallMountPath,
        [Parameter(Mandatory)][string]$WorkRoot,
        [Parameter(Mandatory)][string]$ScratchDirectory,
        [object]$LcuFile,
        [object[]]$LcuAllFiles = @(),
        [object[]]$SafeOSFiles = @(),
        [object[]]$DriverInfFiles = @(),
        [int]$DriverFoundCount = 0,
        [int]$DriverSkippedCount = 0,
        [scriptblock]$LogAction
    )

    $winreSource = Join-Path $InstallMountPath 'Windows\System32\Recovery\winre.wim'
    $driverRequested = @($DriverInfFiles).Count -gt 0
    if (-not (Test-Path -LiteralPath $winreSource)) {
        if ($driverRequested) { throw "winre.wim was not found in the mounted install.wim; requested WinRE drivers cannot be integrated." }
        Invoke-DriverLog $LogAction 'WARN: winre.wim not found; skipping WinRE servicing.' 'WARN'
        return $null
    }

    $winreWork = Join-Path $WorkRoot 'winre_work.wim'
    $winreExport = Join-Path $WorkRoot 'winre_export.wim'
    $winreMount = Join-Path $WorkRoot 'WinREMount'
    $mounted = $false
    $saved = $false
    New-Item -Path $winreMount -ItemType Directory -Force | Out-Null
    Copy-Item -LiteralPath $winreSource -Destination $winreWork -Force
    Set-ItemProperty -LiteralPath $winreWork -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
    try {
        Mount-WindowsImage -ImagePath $winreWork -Index 1 -Path $winreMount -ErrorAction Stop
        $mounted = $true
        Invoke-DriverLog $LogAction "winre.wim mounted: $winreMount" 'INFO'

        if ($LcuFile) {
            $temp = Join-Path $WorkRoot 'LCU_winre_temp'
            New-Item -Path $temp -ItemType Directory -Force | Out-Null
            foreach ($file in @($LcuAllFiles)) { Copy-Item -LiteralPath $file.FullName -Destination $temp -Force }
            $target = Join-Path $temp $LcuFile.Name
            try {
                Add-WindowsPackage -PackagePath $target -Path $winreMount -ScratchDirectory $ScratchDirectory -ErrorAction Stop | Out-Null
                Invoke-DriverLog $LogAction "WinRE SSU applied: $($LcuFile.Name)" 'INFO'
            } catch {
                $err = $_.Exception.Message
                if ($err -match '0x8007007e') {
                    Invoke-DriverLog $LogAction 'WinRE SSU 0x8007007e ignored per Microsoft guidance.' 'INFO'
                } elseif ($err -match '0x800f081e') {
                    Invoke-DriverLog $LogAction 'WinRE SSU not applicable (0x800f081e).' 'INFO'
                } else { throw "SSU failed on WinRE: $err" }
            } finally { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
        }
        foreach ($safeOS in @($SafeOSFiles)) {
            try {
                Add-WindowsPackage -PackagePath $safeOS.FullName -Path $winreMount -ScratchDirectory $ScratchDirectory -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
                Invoke-DriverLog $LogAction "WinRE SafeOS applied: $($safeOS.Name)" 'INFO'
            } catch {
                if ($_.Exception.Message -match '0x800f081e') {
                    Invoke-DriverLog $LogAction 'WinRE SafeOS not applicable (0x800f081e).' 'INFO'
                } else { Invoke-DriverLog $LogAction "WARN: WinRE SafeOS failed: $($_.Exception.Message)" 'WARN' }
            }
        }

        if ($driverRequested) {
            $null = Add-DriversToOfflineImage -MountPath $winreMount -DriverInfFiles $DriverInfFiles -TargetLabel 'winre.wim' -FoundCount $DriverFoundCount -SkippedCount $DriverSkippedCount -LogAction $LogAction -ThrowOnFailure
        }
        if ($null -ne $LcuFile -or @($SafeOSFiles).Count -gt 0) {
            & dism.exe "/Image:$winreMount" '/Cleanup-Image' '/StartComponentCleanup' '/ResetBase' '/Defer' | Out-Null
            if ($LASTEXITCODE -ne 0) { Invoke-DriverLog $LogAction "WARN: WinRE cleanup exit code $LASTEXITCODE" 'WARN' }
        }
        Dismount-WindowsImage -Path $winreMount -Save -ErrorAction Stop
        $mounted = $false; $saved = $true
        Export-WindowsImage -SourceImagePath $winreWork -SourceIndex 1 -DestinationImagePath $winreExport -CompressionType maximum -ErrorAction Stop
        Set-ItemProperty -LiteralPath $winreSource -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
        Copy-Item -LiteralPath $winreExport -Destination $winreSource -Force
        Invoke-DriverLog $LogAction 'winre.wim exported and written back to install.wim.' 'INFO'
    } catch {
        if ($mounted) { try { Dismount-WindowsImage -Path $winreMount -Discard -ErrorAction SilentlyContinue } catch {} ; $mounted = $false }
        if (-not $driverRequested) {
            # Preserve the existing update-only behavior: a WinRE update error
            # is logged and the original WinRE image is carried forward.
            Invoke-DriverLog $LogAction "WARN: WinRE servicing failed; retaining the original WinRE image: $($_.Exception.Message)" 'WARN'
            Export-WindowsImage -SourceImagePath $winreWork -SourceIndex 1 -DestinationImagePath $winreExport -CompressionType maximum -ErrorAction Stop
            Set-ItemProperty -LiteralPath $winreSource -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue
            Copy-Item -LiteralPath $winreExport -Destination $winreSource -Force
            return
        }
        throw
    } finally {
        Remove-Item -LiteralPath $winreMount -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $winreWork -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $winreExport -Force -ErrorAction SilentlyContinue
    }
}
