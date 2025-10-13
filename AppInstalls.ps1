param(
    [string[]]
    $Categories,
    [switch]
    $All,
    [int]
    $MaxParallel = 4,
    [switch]
    $Update
)

$CategoryMap = @{
    'PostInstall' = @(
        'IObit.DriverBooster'
    )
    'Base' = @(
        'Nilesoft.Shell',
        'Giorgiotani.Peazip',
        'DuongDieuPhap.ImageGlass',
        'Starpine.Screenbox'
    )
    'MidUser' = @(
        'Bopsoft.Listary',
        'PDFgear.PDFgear',
        'AntibodySoftware.WizTree',
        'CodeSector.TeraCopy',
        'Vivaldi.Vivaldi',
        'SoftDeluxe.FreeDownloadManager',
        'Klocman.BulkCrapUninstaller',
        'flux.flux'
    )
    'PowerUser' = @(
        'Microsoft.PowerToys',
        'ShareX.ShareX',
        'Espanso.Espanso',
        'AutoHotkey.AutoHotkey',
        'QL-Win.QuickLook',
        'hluk.CopyQ'
    )
    'DevTools' = @(
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
    'PrivacySuite' = @(
        'Proton.ProtonDrive',
        'IDRIX.VeraCrypt',
        'Proton.ProtonPass',
        'Proton.ProtonMail',
        'Proton.ProtonVPN',
        'Tailscale.Tailscale',
        'OpenWhisperSystems.Signal',
        'Cryptomator.Cryptomator'
    )
    'MonitoringBenchmark' = @(
        'REALiX.HWiNFO',
        'CrystalDewWorld.CrystalDiskInfo',
        'CrystalDewWorld.CrystalDiskMark',
        'WinsiderSS.SystemInformer',
        'Resplendence.WhoCrashed',
        'Famatech.AdvancedIPScanner'
    )
    'Dependencies' = @(
        'Microsoft.VCRedist.All',
        'Microsoft.DotNet.DesktopRuntime.6',
        'Microsoft.DotNet.DesktopRuntime.8',
        'Microsoft.DotNet.Runtime.7',
        'Microsoft.DotNet.Runtime.8',
        'Microsoft.DotNet.Runtime.9'
    )
    'MultimediaStreaming' = @(
        'Stremio.Stremio.Beta',
        'flux.flux'
    )
    'FileTransferImaging' = @(
        'Google.QuickShare',
        'LocalSend.LocalSend'
    )
    'NotesWriting' = @(
        'Obsidian.Obsidian'
    )
    'UtilitiesOther' = @(
        'Nlitesoft.NTLite',
        'BillStewart.SyncthingWindowsSetup',
        'Balena.Etcher'
    )
    'OnHold' = @(
        'BlastApps.FluentSearch',
        'Ablaze.Floorp'
    )
}

function Install-WinGetPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Id,
        [switch]
        $Update
    )
    $commonArgs = @('--id', $Id, '--exact', '--source', 'winget', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity', '--silent')
    if ($Update) {
        winget upgrade @commonArgs
    } else {
        winget install @commonArgs
    }
}

function Invoke-Category {
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Name,
        [switch]
        $Update,
        [int]
        $MaxParallel = 4
    )
    if (-not $CategoryMap.ContainsKey($Name)) { throw "Unknown category '$Name'" }
    $ids = $CategoryMap[$Name]
    $ids | ForEach-Object -Parallel {
        param($using:Update)
        Install-WinGetPackage -Id $_ -Update:$using:Update
    } -ThrottleLimit $MaxParallel
}

# Per-category convenience functions so you can call PostInstall, Base, etc. directly
function PostInstall { Invoke-Category -Name 'PostInstall' -MaxParallel 4 }
function Base { Invoke-Category -Name 'Base' -MaxParallel 4 }
function MidUser { Invoke-Category -Name 'MidUser' -MaxParallel 4 }
function PowerUser { Invoke-Category -Name 'PowerUser' -MaxParallel 4 }
function DevTools { Invoke-Category -Name 'DevTools' -MaxParallel 4 }
function PrivacySuite { Invoke-Category -Name 'PrivacySuite' -MaxParallel 4 }
function MonitoringBenchmark { Invoke-Category -Name 'MonitoringBenchmark' -MaxParallel 4 }
function Dependencies { Invoke-Category -Name 'Dependencies' -MaxParallel 4 }
function MultimediaStreaming { Invoke-Category -Name 'MultimediaStreaming' -MaxParallel 4 }
function FileTransferImaging { Invoke-Category -Name 'FileTransferImaging' -MaxParallel 4 }
function NotesWriting { Invoke-Category -Name 'NotesWriting' -MaxParallel 4 }
function UtilitiesOther { Invoke-Category -Name 'UtilitiesOther' -MaxParallel 4 }
function OnHold { Invoke-Category -Name 'OnHold' -MaxParallel 4 }

# If script is invoked directly with parameters (from menu), dispatch accordingly
if ($PSBoundParameters.Count -gt 0) {
    $selected = @()
    if ($All) {
        $selected = @($CategoryMap.Keys | Sort-Object)
    } elseif ($Categories -and $Categories.Count -gt 0) {
        $selected = $Categories
    }
    foreach ($c in $selected) {
        Invoke-Category -Name $c -Update:$Update -MaxParallel $MaxParallel
    }
}