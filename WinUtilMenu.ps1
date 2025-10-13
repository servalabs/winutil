$ErrorActionPreference = 'Stop'

${script:RemoteHostsBaseUrl} = 'https://raw.githubusercontent.com/servalabs/winutil/refs/heads/main/hosts/'
${script:RemoteHostsNames} = @(
    'adobe',
    'autodesk',
    'corel',
    'glasswire',
    'lightburn'
)

function Assert-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        throw 'This action requires Administrator privileges. Please run PowerShell as Administrator.'
    }
}

function Read-MenuSelection {
    param(
        [Parameter(Mandatory = $true)]
        [int]
        $Min,
        [Parameter(Mandatory = $true)]
        [int]
        $Max
    )
    while ($true) {
        Write-Host -NoNewline 'Select an option: '
        $inputValue = Read-Host
        if ([int]::TryParse($inputValue, [ref]$null)) {
            $number = [int]$inputValue
            if ($number -ge $Min -and $number -le $Max) { return $number }
        }
        Write-Host "Please enter a number between $Min and $Max." -ForegroundColor Yellow
    }
}

function Parse-MultiSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $Input,
        [Parameter(Mandatory = $true)]
        [int]
        $Max
    )
    $indices = New-Object System.Collections.Generic.List[int]
    $parts = $Input -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    foreach ($p in $parts) {
        if ([int]::TryParse($p, [ref]$null)) {
            $n = [int]$p
            if ($n -ge 1 -and $n -le $Max -and -not $indices.Contains($n)) {
                $indices.Add($n)
            }
        }
    }
    return ,($indices.ToArray())
}

function Show-MainMenu {
    Clear-Host
    Write-Host '=== Windows Utility Menu ===' -ForegroundColor Cyan
    Write-Host '1) Activate Windows / Microsoft Office'
    Write-Host '2) Install Apps'
    Write-Host '3) Network Blocklists (hosts)'
    Write-Host '4) Clean Windows (Disk Cleanup)'
    Write-Host '5) Improve Privacy (O&O ShutUp10)'
    Write-Host '0) Exit'
}

function Invoke-AppInstallsSubmenu {
    # Categories and their full winget install commands
    $CategoryMap = @{
        'PostInstall' = @(
            'winget install --id IObit.DriverBooster --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'Base' = @(
            'winget install --id Nilesoft.Shell --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Giorgiotani.Peazip --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id DuongDieuPhap.ImageGlass --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Starpine.Screenbox --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'MidUser' = @(
            'winget install --id Bopsoft.Listary --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id PDFgear.PDFgear --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id AntibodySoftware.WizTree --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id CodeSector.TeraCopy --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Vivaldi.Vivaldi --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id SoftDeluxe.FreeDownloadManager --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Klocman.BulkCrapUninstaller --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id flux.flux --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'PowerUser' = @(
            'winget install --id Microsoft.PowerToys --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id ShareX.ShareX --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Espanso.Espanso --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id AutoHotkey.AutoHotkey --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id QL-Win.QuickLook --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id hluk.CopyQ --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'DevTools' = @(
            'winget install --id GitHub.GitHubDesktop --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id PostgreSQL.pgAdmin --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Anysphere.Cursor --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Microsoft.WindowsTerminal --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id OpenJS.NodeJS --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Microsoft.PowerShell --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Git.Git --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Oracle.JDK.25 --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Python.Python.3.10 --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Python.Launcher --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Mobatek.MobaXterm --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id junegunn.fzf --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id BurntSushi.ripgrep.MSVC --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'PrivacySuite' = @(
            'winget install --id Proton.ProtonDrive --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id IDRIX.VeraCrypt --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Proton.ProtonPass --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Proton.ProtonMail --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Proton.ProtonVPN --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Tailscale.Tailscale --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id OpenWhisperSystems.Signal --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Cryptomator.Cryptomator --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'MonitoringBenchmark' = @(
            'winget install --id REALiX.HWiNFO --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id CrystalDewWorld.CrystalDiskInfo --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id CrystalDewWorld.CrystalDiskMark --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id WinsiderSS.SystemInformer --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Resplendence.WhoCrashed --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Famatech.AdvancedIPScanner --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'Dependencies' = @(
            'winget install --id Microsoft.VCRedist.All --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Microsoft.DotNet.DesktopRuntime.6 --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Microsoft.DotNet.DesktopRuntime.8 --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Microsoft.DotNet.Runtime.7 --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Microsoft.DotNet.Runtime.8 --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Microsoft.DotNet.Runtime.9 --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'MultimediaStreaming' = @(
            'winget install --id Stremio.Stremio.Beta --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id flux.flux --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'FileTransferImaging' = @(
            'winget install --id Google.QuickShare --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id LocalSend.LocalSend --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'NotesWriting' = @(
            'winget install --id Obsidian.Obsidian --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'UtilitiesOther' = @(
            'winget install --id Nlitesoft.NTLite --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id BillStewart.SyncthingWindowsSetup --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Balena.Etcher --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
        'OnHold' = @(
            'winget install --id BlastApps.FluentSearch --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent',
            'winget install --id Ablaze.Floorp --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
        )
    }
    $categories = @(
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
    )

    while ($true) {
        Clear-Host
        Write-Host '=== Install Apps ===' -ForegroundColor Cyan
        Write-Host 'U) Upgrade ALL installed apps'
        $i = 1
        foreach ($c in $categories) {
            Write-Host ("{0}) {1}" -f $i, $c)
            $i++
        }
        Write-Host ("{0}) Install ALL categories" -f $i)
        Write-Host '0) Back'
        Write-Host -NoNewline 'Enter number(s) comma-separated, A for all, U to upgrade all, or 0 to back: '
        $raw = Read-Host
        if ($raw -eq '0') { return }
        if ($raw -match '^[Uu]$') {
            $upgradeAllCmd = 'winget upgrade --all --accept-source-agreements --accept-package-agreements --disable-interactivity --silent'
            cmd.exe /c $upgradeAllCmd
            Write-Host 'Done. Press Enter to continue...'
            [void][System.Console]::ReadLine()
            continue
        }
        if ($raw -match '^[Aa]$') {
            # Ask whether to Install or Update
            $mode = $null
            while ($null -eq $mode) {
                Write-Host 'Choose mode: 1) Install  2) Update existing' -ForegroundColor Cyan
                $m = Read-Host
                if ($m -eq '1') { $mode = 'Install' }
                elseif ($m -eq '2') { $mode = 'Update' }
            }
            $selected = $categories
            foreach ($cat in $selected) {
                $commands = $CategoryMap[$cat]
                foreach ($cmd in $commands) {
                    $toRun = if ($mode -eq 'Update') { ($cmd -replace '\binstall\b','upgrade') } else { $cmd }
                    cmd.exe /c $toRun
                }
            }
        } else {
            $sel = Parse-MultiSelection -Input $raw -Max ($i - 1)
            if (-not $sel -or $sel.Count -eq 0) {
                Write-Host 'Invalid selection.' -ForegroundColor Yellow
                continue
            }
            $chosen = @()
            foreach ($idx in $sel) { $chosen += $categories[$idx - 1] }
            # Ask whether to Install or Update
            $mode = $null
            while ($null -eq $mode) {
                Write-Host 'Choose mode: 1) Install  2) Update existing' -ForegroundColor Cyan
                $m = Read-Host
                if ($m -eq '1') { $mode = 'Install' }
                elseif ($m -eq '2') { $mode = 'Update' }
            }
            foreach ($cat in $chosen) {
                if (-not $CategoryMap.ContainsKey($cat)) { continue }
                $commands = $CategoryMap[$cat]
                foreach ($cmd in $commands) {
                    $toRun = if ($mode -eq 'Update') { ($cmd -replace '\binstall\b','upgrade') } else { $cmd }
                    cmd.exe /c $toRun
                }
            }
        }
        Write-Host 'Done. Press Enter to continue...'
        [void][System.Console]::ReadLine()
    }
}

function Get-HostsBlocklistFiles {
    $dir = Join-Path $PSScriptRoot 'hosts'
    $local = @()
    if (Test-Path $dir) {
        $local = Get-ChildItem -Path $dir -File | Sort-Object Name
    }

    # Compose combined set: prefer local files; include remote names not present locally
    $localNames = @{}
    foreach ($f in $local) { $localNames[$f.BaseName] = $true }
    $combined = New-Object System.Collections.Generic.List[object]
    foreach ($f in $local) { [void]$combined.Add($f) }
    foreach ($name in ${script:RemoteHostsNames}) {
        if (-not $localNames.ContainsKey($name)) {
            # represent remote-only item with a PSCustomObject carrying Name/BaseName and a marker
            $obj = [pscustomobject]@{
                Name     = $name
                BaseName = $name
                FullName = $null
                IsRemote = $true
            }
            [void]$combined.Add($obj)
        }
    }
    return ,($combined | Sort-Object Name)
}

function Apply-HostsBlocklists {
    Assert-IsAdmin
    $files = Get-HostsBlocklistFiles
    if (-not $files) {
        Write-Host 'No blocklist files found in hosts/.' -ForegroundColor Yellow
        return
    }

    Clear-Host
    Write-Host '=== Apply Network Blocklists ===' -ForegroundColor Cyan
    for ($i = 0; $i -lt $files.Count; $i++) {
        $label = $files[$i].Name
        if ($files[$i].PSObject.Properties.Match('IsRemote').Count -gt 0 -and $files[$i].IsRemote) {
            $label = "$label (remote)"
        }
        Write-Host ("{0}) {1}" -f ($i+1), $label)
    }
    Write-Host 'A) Apply ALL'
    Write-Host '0) Back'
    Write-Host -NoNewline 'Enter number(s) comma-separated, A for all, or 0 to back: '
    $choice = Read-Host
    if ($choice -eq '0') { return }

    $selectedFiles = @()
    if ($choice -match '^[Aa]$') {
        $selectedFiles = $files
    } else {
        $sel = Parse-MultiSelection -Input $choice -Max $files.Count
        if ($sel -and $sel.Count -gt 0) {
            foreach ($idx in $sel) { $selectedFiles += $files[$idx - 1] }
        }
    }

    if (-not $selectedFiles -or $selectedFiles.Count -eq 0) {
        Write-Host 'Invalid selection.' -ForegroundColor Yellow
        return
    }

    $systemHostsPath = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
    $backupPath = "$systemHostsPath.bak"

    if (-not (Test-Path $backupPath)) {
        Copy-Item -Path $systemHostsPath -Destination $backupPath -Force
        Write-Host "Backup created: $backupPath" -ForegroundColor DarkGray
    }

    $original = Get-Content -Path $systemHostsPath -Raw -ErrorAction SilentlyContinue
    $merged = [System.Text.StringBuilder]::new()
    if ($original) { [void]$merged.AppendLine($original.TrimEnd()) }
    [void]$merged.AppendLine()
    [void]$merged.AppendLine('# ----- BEGIN winutil blocklists -----')
    foreach ($f in $selectedFiles) {
        $isRemote = $false
        if ($f.PSObject.Properties.Match('IsRemote').Count -gt 0) { $isRemote = [bool]$f.IsRemote }
        $sourceTag = if ($isRemote) { "# from ${script:RemoteHostsBaseUrl}$($f.Name)" } else { "# from hosts/" + $f.Name }
        [void]$merged.AppendLine($sourceTag)

        $content = $null
        if ($isRemote) {
            try {
                $url = "${script:RemoteHostsBaseUrl}$($f.Name)"
                $resp = Invoke-WebRequest -UseBasicParsing -Uri $url -Method GET -TimeoutSec 30
                $content = $resp.Content
            } catch {
                $content = "# Failed to fetch remote blocklist: $($f.Name)"
            }
        } else {
            $content = Get-Content -Path $f.FullName -Raw
        }

        if ($content) { [void]$merged.AppendLine(($content.TrimEnd())) }
        [void]$merged.AppendLine()
    }
    [void]$merged.AppendLine('# ----- END winutil blocklists -----')

    Set-Content -Path $systemHostsPath -Value $merged.ToString() -Encoding ASCII -Force
    Write-Host 'Hosts file updated.' -ForegroundColor Green
}

function Remove-HostsBlocklists {
    Assert-IsAdmin
    $systemHostsPath = Join-Path $env:SystemRoot 'System32\drivers\etc\hosts'
    $backupPath = "$systemHostsPath.bak"
    if (Test-Path $backupPath) {
        Copy-Item -Path $backupPath -Destination $systemHostsPath -Force
        Write-Host 'Hosts file restored from backup.' -ForegroundColor Green
    } else {
        # Fallback: strip the marked section
        $raw = Get-Content -Path $systemHostsPath -Raw -ErrorAction SilentlyContinue
        if (-not $raw) { Write-Host 'No hosts content found.' -ForegroundColor Yellow; return }
        $pattern = '(?s)# ----- BEGIN winutil blocklists -----.*?# ----- END winutil blocklists -----\s*'
        $cleaned = [System.Text.RegularExpressions.Regex]::Replace($raw, $pattern, '')
        Set-Content -Path $systemHostsPath -Value $cleaned -Encoding ASCII -Force
        Write-Host 'Removed winutil blocklists section from hosts.' -ForegroundColor Green
    }
}

function Invoke-HostsMenu {
    while ($true) {
        Clear-Host
        Write-Host '=== Network Blocklists (hosts) ===' -ForegroundColor Cyan
        Write-Host '1) Apply blocklist(s)'
        Write-Host '2) Remove applied blocklists (restore backup or strip section)'
        Write-Host '0) Back'
        $selection = Read-MenuSelection -Min 0 -Max 2
        switch ($selection) {
            1 { Apply-HostsBlocklists }
            2 { Remove-HostsBlocklists }
            0 { return }
        }
        Write-Host 'Press Enter to continue...'
        [void][System.Console]::ReadLine()
    }
}

function Invoke-CleanWindows {
    try {
        Start-Process -FilePath 'cleanmgr.exe' -Verb RunAs
    } catch {
        Write-Host "Failed to launch Disk Cleanup: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-PrivacyImprovements {
    # Try to launch ShutUp10 if installed; otherwise offer to install via winget
    $candidatePaths = @(
        "$Env:ProgramFiles\\O&O Software\\ShutUp10\\OOSU10.exe",
        "$Env:ProgramFiles(x86)\\O&O Software\\ShutUp10\\OOSU10.exe"
    )
    $exe = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($exe) {
        try { Start-Process -FilePath $exe -Verb RunAs } catch {}
        return
    }

    Write-Host 'O&O ShutUp10 not found. Downloading and launching (no install)...' -ForegroundColor Cyan
    try {
        $dest = Join-Path $env:TEMP ('OOSU10-' + [System.Guid]::NewGuid().ToString('N') + '.exe')
        Invoke-WebRequest -Uri 'https://dl5.oo-software.com/files/ooshutup10/OOSU10.exe' -OutFile $dest -UseBasicParsing
        if (Test-Path $dest) {
            try { Start-Process -FilePath $dest -Verb RunAs } catch {}
            return
        }
    } catch {
        Write-Host "Failed to download ShutUp10: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Invoke-ActivationMenu {
    while ($true) {
        Clear-Host
        Write-Host '=== Activation ===' -ForegroundColor Cyan
        Write-Host '1) Run activation script (Windows / Office)'
        Write-Host '2) Open Windows Activation settings'
        Write-Host '0) Back'
        $sel = Read-MenuSelection -Min 0 -Max 2
        switch ($sel) {
            1 {
                try {
                    Write-Host 'Running activation script...' -ForegroundColor Cyan
                    Invoke-Expression (curl.exe -s --doh-url https://1.1.1.1/dns-query https://get.activated.win | Out-String)
                } catch {
                    Write-Host "Activation script failed: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
            2 {
                try { Start-Process 'ms-settings:activation' } catch { Write-Host 'Failed to open Settings.' -ForegroundColor Yellow }
            }
            0 { return }
        }
        Write-Host 'Press Enter to continue...'
        [void][System.Console]::ReadLine()
    }
}

while ($true) {
    Show-MainMenu
    $choice = Read-MenuSelection -Min 0 -Max 5
    $shouldPause = $true
    switch ($choice) {
        1 { Invoke-ActivationMenu; $shouldPause = $false }
        2 { Invoke-AppInstallsSubmenu; $shouldPause = $false }
        3 { Invoke-HostsMenu; $shouldPause = $false }
        4 { Invoke-CleanWindows }
        5 { Invoke-PrivacyImprovements }
        0 { return }
    }
    if ($choice -ne 0 -and $shouldPause) {
        Write-Host 'Press Enter to return to main menu...'
        [void][System.Console]::ReadLine()
    }
}


