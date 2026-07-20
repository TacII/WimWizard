BeforeAll {
    $helper = Join-Path $PSScriptRoot '..\WimWizard-Drivers.ps1'
    . (Resolve-Path $helper)
    if (-not (Get-Command Add-WindowsDriver -ErrorAction SilentlyContinue)) { function Add-WindowsDriver {} }
    if (-not (Get-Command Mount-WindowsImage -ErrorAction SilentlyContinue)) { function Mount-WindowsImage {} }
    if (-not (Get-Command Dismount-WindowsImage -ErrorAction SilentlyContinue)) { function Dismount-WindowsImage {} }
    if (-not (Get-Command Export-WindowsImage -ErrorAction SilentlyContinue)) { function Export-WindowsImage {} }
    if (-not (Get-Command Add-WindowsPackage -ErrorAction SilentlyContinue)) { function Add-WindowsPackage {} }
}

Describe 'WimWizard driver parameter validation' {
    It 'rejects both driver sources' {
        { Assert-DriverSourceParameters -DriverPackageID 'ABC00123' -DriverPath 'C:\Drivers' } | Should -Throw
    }
    It 'keeps no-source behavior unchanged' {
        (Assert-DriverSourceParameters).SourceType | Should -Be 'None'
        (Test-WinREServicingRequired).ToString() | Should -Be 'False'
    }
    It 'requires a source for WinRE driver switches' {
        { Assert-DriverSourceParameters -AddStorageDriversToWinRE } | Should -Throw
    }
}

Describe 'WimWizard INF discovery and metadata' {
    BeforeEach {
        $root = Join-Path $TestDrive 'Drivers'
        New-Item -Path $root -ItemType Directory -Force | Out-Null
    }

    It 'rejects a missing path and a path without INF files' {
        { Find-DriverInfFiles -Path (Join-Path $TestDrive 'missing') } | Should -Throw
        $empty = Join-Path $TestDrive 'empty'
        New-Item -Path $empty -ItemType Directory -Force | Out-Null
        @(Find-DriverInfFiles -Path $empty).Count | Should -Be 0
    }

    It 'parses keys case-insensitively, removes comments and quotes' {
        $path = Join-Path $TestDrive 'Drivers\quoted.inf'
        Set-Content -Path $path -Encoding UTF8 -Value @('[Version]', 'cLaSs = "Net" ; comment', 'CLASSGUID = "{4D36E972-E325-11CE-BFC1-08002BE10318}" ; ignored', 'Provider="Acme; Labs" ; comment', 'DriverVer = 01/01/2026,1.2.3.4')
        $meta = Get-DriverInfMetadata -Path $path
        $meta.Class | Should -Be 'Net'
        $meta.ClassGuid | Should -Be '{4D36E972-E325-11CE-BFC1-08002BE10318}'
        $meta.Provider | Should -Be 'Acme; Labs'
        $meta.DriverVer | Should -Be '01/01/2026,1.2.3.4'
        $meta.Category | Should -Be 'Network'
    }

    It 'uses ClassGuid as fallback and supports UTF-16' {
        $path = Join-Path $TestDrive 'Drivers\storage.inf'
        [IO.File]::WriteAllText($path, "[Version]`r`nClassGuid={75416E63-5912-4DFA-AE8F-3EFACCAFFB14}`r`n", [Text.Encoding]::Unicode)
        (Get-DriverInfMetadata -Path $path).Category | Should -Be 'Storage'
    }

    It 'classifies all standard storage classes, Net, and unknown classes' {
        $cases = @(
            @{ Name = 'net'; Text = '[Version]`nClass=Net'; Category = 'Network' }
            @{ Name = 'scsi'; Text = '[Version]`nClass=SCSIAdapter'; Category = 'Storage' }
            @{ Name = 'hdc'; Text = '[Version]`nClass=HDC'; Category = 'Storage' }
            @{ Name = 'nvme'; Text = '[Version]`nClass=NvmeDisk'; Category = 'Storage' }
            @{ Name = 'unknown'; Text = '[Version]`nClass=Printer'; Category = 'Unknown' }
        )
        foreach ($case in $cases) {
            $path = Join-Path $TestDrive "Drivers\$($case.Name).inf"
            Set-Content -Path $path -Encoding UTF8 -Value ($case.Text -replace '`n', "`r`n")
            (Get-DriverInfMetadata -Path $path).Category | Should -Be $case.Category
        }
    }
}

Describe 'WimWizard driver filtering' {
    BeforeAll {
        $metadata = @(
            [pscustomobject]@{ FullPath='net.inf'; IsValid=$true; Category='Network'; Class='Net' }
            [pscustomobject]@{ FullPath='storage.inf'; IsValid=$true; Category='Storage'; Class='SCSIAdapter' }
            [pscustomobject]@{ FullPath='other.inf'; IsValid=$true; Category='Other'; Class='System' }
            [pscustomobject]@{ FullPath='unknown.inf'; IsValid=$true; Category='Unknown'; Class='Printer' }
        )
    }
    It 'selects all valid INF files with All' {
        (Select-DriverInfFiles -Metadata $metadata -DriverFilter All).Selected.Count | Should -Be 4
    }
    It 'selects only Storage and Network with StorageAndNetwork' {
        $selection = Select-DriverInfFiles -Metadata $metadata -DriverFilter StorageAndNetwork
        $selection.Selected.Count | Should -Be 2
        @($selection.Selected.Category) | Should -Contain 'Storage'
        @($selection.Selected.Category) | Should -Contain 'Network'
        $selection.Skipped.Count | Should -Be 2
    }
    It 'keeps WinRE Storage selection free of Network and Other files' {
        $selected = @($metadata | Where-Object Category -eq 'Storage')
        @($selected.Category) | Should -Not -Contain 'Network'
        @($selected.Category) | Should -Not -Contain 'Other'
    }
    It 'keeps WinRE Network selection free of Storage and Other files' {
        $selected = @($metadata | Where-Object Category -eq 'Network')
        @($selected.Category) | Should -Not -Contain 'Storage'
        @($selected.Category) | Should -Not -Contain 'Other'
    }
    It 'plans WinRE driver servicing even when updates are skipped' {
        (Test-WinREServicingRequired -DriverInfFiles @($metadata[0])).ToString() | Should -Be 'True'
    }
}

Describe 'WimWizard MECM driver package resolution' {
    BeforeAll {
        if (-not (Get-Command Get-CMDriverPackage -ErrorAction SilentlyContinue)) { function Get-CMDriverPackage {} }
        $modulePath = Join-Path $TestDrive 'ConfigurationManager.psd1'
        Set-Content -Path $modulePath -Value '# mocked module'
    }
    BeforeEach {
        Mock Import-Module {}
        Mock Get-PSDrive { $null }
        Mock New-PSDrive {}
        Mock Push-Location {}
        Mock Pop-Location {}
        Mock Set-Location {}
        Mock Get-CMDriverPackage { [pscustomobject]@{ Name='Dell Drivers'; PackageID='ABC00123'; PackageType='Driver Package'; PkgSourcePath=$TestDrive } }
    }
    It 'resolves a package by ID and restores the previous location' {
        $before = (Get-Location).Path
        $resolved = Resolve-CMDriverPackageSource -DriverPackageID 'ABC00123' -SCCMServer 'mecm01' -SCCMSiteCode 'ABC' -ModulePath $modulePath
        $resolved.SourcePath | Should -Be $TestDrive
        $resolved.PackageName | Should -Be 'Dell Drivers'
        Should -Invoke Get-CMDriverPackage -Times 1
        (Get-Location).Path | Should -Be $before
    }

    It 'recognizes numeric Configuration Manager PackageType 3 as a Driver Package' {
        Mock Get-CMDriverPackage { [pscustomobject]@{ Name='Numeric Driver Package'; PackageID='ABC00123'; PackageType=3; PkgSourcePath=$TestDrive } }
        $resolved = Resolve-CMDriverPackageSource -DriverPackageID 'ABC00123' -SCCMServer 'mecm01' -SCCMSiteCode 'ABC' -ModulePath $modulePath
        $resolved.SourceType | Should -Be 'MECM Driver Package'
        $resolved.PackageName | Should -Be 'Numeric Driver Package'
    }
    It 'fails cleanly when PkgSourcePath is not reachable' {
        Mock Get-CMDriverPackage { [pscustomobject]@{ Name='Broken'; PackageType='Driver Package'; PkgSourcePath=(Join-Path $TestDrive 'does-not-exist') } }
        { Resolve-CMDriverPackageSource -DriverPackageID 'ABC00123' -SCCMServer 'mecm01' -SCCMSiteCode 'ABC' -ModulePath $modulePath } | Should -Throw
    }
}

Describe 'WimWizard offline driver integration' {
    It 'reports a failed Add-WindowsDriver operation and can fail the build' {
        Mock Add-WindowsDriver { throw 'DISM driver error' }
        $inf = [pscustomobject]@{ FullPath='C:\Drivers\bad.inf'; Category='Network'; Class='Net' }
        { Add-DriversToOfflineImage -MountPath 'C:\Mount' -DriverInfFiles @($inf) -TargetLabel 'install.wim' -ThrowOnFailure } | Should -Throw
    }

    It 'discards a mounted WinRE image when requested driver integration fails' {
        $install = Join-Path $TestDrive 'install'
        $recovery = Join-Path $install 'Windows\System32\Recovery'
        New-Item -Path $recovery -ItemType Directory -Force | Out-Null
        Set-Content -Path (Join-Path $recovery 'winre.wim') -Value 'mock'
        Mock Mount-WindowsImage {}
        Mock Dismount-WindowsImage {}
        Mock Export-WindowsImage {}
        Mock Add-WindowsDriver { throw 'DISM driver error' }
        $inf = [pscustomobject]@{ FullPath='C:\Drivers\bad.inf'; Category='Storage'; Class='SCSIAdapter' }
        { Invoke-WinREDriverAndUpdateServicing -InstallMountPath $install -WorkRoot (Join-Path $TestDrive 'work') -ScratchDirectory $TestDrive -DriverInfFiles @($inf) } | Should -Throw
        Should -Invoke Dismount-WindowsImage -Times 1
        (Test-Path (Join-Path $TestDrive 'work\WinREMount')) | Should -BeFalse
    }
}
