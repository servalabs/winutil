param(
    [ValidateSet(
        'PostInstall',
        'Base',
        'MidUser',
        'PowerUser',
        'DevTools',
        'PrivacySuite',
        'MonitoringBenchmark',
        'Dependencies',
        'MultimediaStreaming',
        'FileTransferImaging',
        'NotesWriting',
        'UtilitiesOther',
        'OnHold'
    )]
    [string]
    $Category,

    [ValidateSet(
        'PostInstall',
        'Base',
        'MidUser',
        'PowerUser',
        'DevTools',
        'PrivacySuite',
        'MonitoringBenchmark',
        'Dependencies',
        'MultimediaStreaming',
        'FileTransferImaging',
        'NotesWriting',
        'UtilitiesOther',
        'OnHold'
    )]
    [string[]]
    $Categories,

    [switch]
    $All
)

function Install-WinGetPackages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]
        $PackageIds
    )

    $uniqueIds = $PackageIds | Where-Object { $_ -and $_.Trim() -ne '' } | Select-Object -Unique

    foreach ($id in $uniqueIds) {
        Write-Host "Installing $id ..." -ForegroundColor Cyan
        winget install --id $id --exact --accept-source-agreements --accept-package-agreements --silent --disable-interactivity
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Installation reported a non-zero exit code for $id ($LASTEXITCODE). Continuing..."
        }
    }
}

function Install-PostInstall {
    Install-WinGetPackages @(
        'IObit.DriverBooster'
    )
}

function Install-Base {
    Install-WinGetPackages @(
        'Nilesoft.Shell',
        'Giorgiotani.Peazip',
        'DuongDieuPhap.ImageGlass',
        'Starpine.Screenbox'
    )
}

function Install-MidUser {
    Install-WinGetPackages @(
        'Bopsoft.Listary',
        'PDFgear.PDFgear',
        'AntibodySoftware.WizTree',
        'CodeSector.TeraCopy',
        'Vivaldi.Vivaldi',
        'SoftDeluxe.FreeDownloadManager',
        'Klocman.BulkCrapUninstaller',
        'flux.flux'
    )
}

function Install-PowerUser {
    Install-WinGetPackages @(
        'Microsoft.PowerToys',
        'ShareX.ShareX',
        'Espanso.Espanso',
        'AutoHotkey.AutoHotkey',
        'QL-Win.QuickLook',
        'hluk.CopyQ'
        # Macrorit Partition Master free is provided as a direct download in the markdown list.
        # It is not installed here since it's not a Winget ID.
    )

    try {
        Write-Host "Fetching Macrorit Partition Master Free zip..." -ForegroundColor Cyan
        $tempRoot = Join-Path $env:TEMP ("mde-" + [System.Guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        $zipPath = Join-Path $tempRoot 'mde-free-setup.zip'
        $extractPath = Join-Path $tempRoot 'extracted'

        $uri = 'https://disk-tool.com/download/mde/mde-free-setup.zip'
        Invoke-WebRequest -Uri $uri -OutFile $zipPath -UseBasicParsing

        Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

        $desktop = [Environment]::GetFolderPath('Desktop')
        if (-not (Test-Path -Path $desktop)) {
            $desktop = Join-Path $env:USERPROFILE 'Desktop'
        }

        Write-Host "Moving extracted files to Desktop: $desktop" -ForegroundColor Cyan
        Get-ChildItem -Path $extractPath -Force | ForEach-Object {
            Move-Item -Path $_.FullName -Destination $desktop -Force -ErrorAction Continue
        }
    } catch {
        Write-Warning "Failed to fetch or extract Macrorit Partition Master zip: $($_.Exception.Message)"
    }
}

function Install-DevTools {
    Install-WinGetPackages @(
        'GitHub.GitHubDesktop',
        'PostgreSQL.pgAdmin',
        'Anysphere.Cursor',
        'Microsoft.WindowsTerminal',
        'OpenJS.NodeJS',
        'Microsoft.PowerShell',
        'Git.Git',
        'Oracle.JDK.25',
        'Python.Python.3.10',
        'Python.Launcher',
        'Mobatek.MobaXterm',
        'junegunn.fzf',
        'BurntSushi.ripgrep.MSVC'
    )
}

function Install-PrivacySuite {
    Install-WinGetPackages @(
        'Proton.ProtonDrive',
        'IDRIX.VeraCrypt',
        'Proton.ProtonPass',
        'Proton.ProtonMail',
        'Proton.ProtonVPN',
        'Tailscale.Tailscale',
        'OpenWhisperSystems.Signal',
        'Cryptomator.Cryptomator'
    )
}

function Install-MonitoringBenchmark {
    Install-WinGetPackages @(
        'REALiX.HWiNFO',
        'CrystalDewWorld.CrystalDiskInfo',
        'CrystalDewWorld.CrystalDiskMark',
        'WinsiderSS.SystemInformer',
        'Resplendence.WhoCrashed',
        'Famatech.AdvancedIPScanner'
    )
}

function Install-Dependencies {
    Install-WinGetPackages @(
        'Microsoft.VCRedist.All',
        'Microsoft.DotNet.DesktopRuntime.6',
        'Microsoft.DotNet.DesktopRuntime.8',
        'Microsoft.DotNet.Runtime.7',
        'Microsoft.DotNet.Runtime.8',
        'Microsoft.DotNet.Runtime.9'
    )
}

function Install-MultimediaStreaming {
    Install-WinGetPackages @(
        'Stremio.Stremio.Beta',
        'flux.flux'
    )
}

function Install-FileTransferImaging {
    Install-WinGetPackages @(
        'Google.QuickShare',
        'LocalSend.LocalSend'
    )
}

function Install-NotesWriting {
    Install-WinGetPackages @(
        'Obsidian.Obsidian'
    )
}

function Install-UtilitiesOther {
    Install-WinGetPackages @(
        'Nlitesoft.NTLite',
        'BillStewart.SyncthingWindowsSetup',
        'Balena.Etcher'
    )
}

function Install-OnHold {
    Install-WinGetPackages @(
        'BlastApps.FluentSearch',
        'Ablaze.Floorp'
    )
}

function Install-WinGetCategory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet(
            'PostInstall',
            'Base',
            'MidUser',
            'PowerUser',
            'DevTools',
            'PrivacySuite',
            'MonitoringBenchmark',
            'Dependencies',
            'MultimediaStreaming',
            'FileTransferImaging',
            'NotesWriting',
            'UtilitiesOther',
            'OnHold'
        )]
        [string[]]
        $Name
    )

    foreach ($n in $Name) {
        & "Install-$n"
    }
}

# Back-compat alias keeping the old function name available to callers
Set-Alias -Name Install-Category -Value Install-WinGetCategory -Scope Local -Force

if ($PSBoundParameters.ContainsKey('All') -and $All.IsPresent) {
    Install-Base
    Install-Dependencies
    Install-MidUser
    Install-PowerUser
    Install-DevTools
    Install-PrivacySuite
    Install-MonitoringBenchmark
    Install-MultimediaStreaming
    Install-FileTransferImaging
    Install-NotesWriting
    Install-UtilitiesOther
    return
}

if ($PSBoundParameters.ContainsKey('Categories') -and $Categories) {
    Install-WinGetCategory -Name $Categories
} elseif ($PSBoundParameters.ContainsKey('Category') -and $Category) {
    Install-WinGetCategory -Name $Category
}


