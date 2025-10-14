$ErrorActionPreference = 'Stop'

${script:RemoteHostsBaseUrl} = 'https://raw.githubusercontent.com/servalabs/winutil/refs/heads/main/hosts/'
${script:RemoteHostsNames} = @(
    'adobe',
    'autodesk',
    'corel',
    'glasswire',
    'lightburn'
)

# Remote source for unified app manifest (apps.json at repo root)
${script:RemoteAppsBaseUrl} = 'https://raw.githubusercontent.com/servalabs/winutil/refs/heads/main/'

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
    # Accept comma, semicolon, or whitespace separated values
    $parts = $Input -split '[,;\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
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

function Get-PreferredPowerShellExe {
    # Prefer pwsh if available, otherwise fall back to Windows PowerShell
    $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwsh -and $pwsh.CommandType -in @('Application','ExternalScript')) {
        return $pwsh.Source
    }
    $ps = Get-Command powershell -ErrorAction SilentlyContinue
    if ($ps) { return $ps.Source }
    return 'powershell'
}

function Invoke-AppInstallsSubmenu {
    Assert-IsAdmin
    $categories = @('base', 'bench-mon', 'depend', 'dev', 'mid', 'misc', 'power', 'privacy')
    while ($true) {
        Clear-Host
        Write-Host '=== Install Apps ===' -ForegroundColor Cyan
        Write-Host 'U) Upgrade ALL installed apps'
        for ($i = 0; $i -lt $categories.Count; $i++) { Write-Host ("{0}) {1}" -f ($i + 1), $categories[$i]) }
        Write-Host 'A) Run ALL categories'
        Write-Host '0) Back'
        Write-Host -NoNewline 'Enter a number, A for all, U to upgrade all, or 0 to back: '
        $raw = Read-Host
        if ($raw -eq '0') { return }
        if ($raw -match '^[Uu]$') {
            winget upgrade --all --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent
            Write-Host 'Done. Press Enter to continue...'
            [void][System.Console]::ReadLine()
            continue
        }

        $toRun = @()
        if ($raw -match '^[Aa]$') {
            $toRun = $categories
        } else {
            $rawTrim = if ($raw) { $raw.Trim() } else { '' }
            $indices = Parse-MultiSelection -Input $rawTrim -Max $categories.Count
            if ($indices -and $indices.Count -gt 0) {
                foreach ($n in $indices) { $toRun += $categories[$n - 1] }
            }
        }

        if (-not $toRun -or $toRun.Count -eq 0) { Write-Host 'Invalid selection.' -ForegroundColor Yellow; Start-Sleep -Milliseconds 800; continue }

        foreach ($cat in $toRun) {
            $manifest = Get-AppManifest -Category $cat
            if (-not $manifest -or $manifest.Count -eq 0) {
                Write-Host ("No JSON manifest found for '{0}'. Skipping." -f $cat) -ForegroundColor Yellow
                continue
            }
            Install-AppsFromManifest -Category $cat -Apps $manifest
        }

        Write-Host 'Press Enter to continue...'
        [void][System.Console]::ReadLine()
    }
}

# Load unified JSON app manifest (apps.json at repo root), return entire set
function Get-AllAppsManifest {
    $localPath = Join-Path $PSScriptRoot 'apps.json'
    if (Test-Path $localPath) {
        try {
            $content = Get-Content -Path $localPath -Raw -ErrorAction Stop
            $data = $null
            try { $data = $content | ConvertFrom-Json -ErrorAction Stop } catch { $data = $null }
            if ($data) { return ,$data }
        } catch {}
    }
    try {
        $url = "${script:RemoteAppsBaseUrl}apps.json"
        $resp = Invoke-WebRequest -UseBasicParsing -Uri $url -Method GET -TimeoutSec 30
        if ($resp -and $resp.Content) {
            $data = $resp.Content | ConvertFrom-Json -ErrorAction Stop
            return ,$data
        }
    } catch {}
    return @()
}

# Get apps for a specific category from the unified manifest
function Get-AppManifest {
    param(
        [Parameter(Mandatory = $true)] [string] $Category
    )
    $all = Get-AllAppsManifest
    if (-not $all -or $all.Count -eq 0) {
        Write-Host "No apps.json manifest loaded (local or remote)." -ForegroundColor Yellow
        return @()
    }
    $filtered = @($all | Where-Object { $_.category -eq $Category })
    if (-not $filtered -or $filtered.Count -eq 0) {
        Write-Host ("No apps found under category '{0}' in apps.json." -f $Category) -ForegroundColor Yellow
        return @()
    }
    return ,$filtered
}

# Present selection UI and install selected apps via winget
function Install-AppsFromManifest {
    param(
        [Parameter(Mandatory = $true)] [string] $Category,
        [Parameter(Mandatory = $true)] [object[]] $Apps
    )
    $records = $Apps | ForEach-Object {
        [pscustomobject]@{
            Name        = $_.name
            Id          = $_.id
            Description = $_.description
        }
    }

    $selected = $null
    $ogv = Get-Command Out-GridView -ErrorAction SilentlyContinue
    if ($ogv) {
        try {
            $selected = $records | Out-GridView -Title ("Select apps to install from '{0}' (Ctrl/Shift to multi-select), then OK" -f $Category) -PassThru
        } catch {}
    }
    if (-not $selected) {
        Write-Host ("Out-GridView not available or canceled. Falling back to console selection for '{0}'." -f $Category) -ForegroundColor Yellow
        for ($i = 0; $i -lt $records.Count; $i++) {
            Write-Host ("{0,2}) {1}  -  {2}" -f ($i+1), $records[$i].Name, $records[$i].Description)
        }
        Write-Host -NoNewline 'Enter numbers to install (comma/space/semicolon-separated), or press Enter for ALL: '
        $choice = Read-Host
        if ([string]::IsNullOrWhiteSpace($choice)) {
            $selected = $records
        } else {
            $indices = @()
            foreach ($p in ($choice -split '[,;\s]+')) {
                $pt = $p.Trim()
                if ([int]::TryParse($pt, [ref]$null)) {
                    $n = [int]$pt
                    if ($n -ge 1 -and $n -le $records.Count -and -not $indices.Contains($n)) { $indices += $n }
                }
            }
            $selected = @()
            foreach ($n in $indices) { $selected += $records[$n - 1] }
        }
    }

    if (-not $selected -or $selected.Count -eq 0) {
        Write-Host 'No apps selected. Skipping.' -ForegroundColor Yellow
        return
    }

    foreach ($app in $selected) {
        try {
            Write-Host ("Installing: {0} ({1})" -f $app.Name, $app.Id) -ForegroundColor Cyan
            winget install --id $app.Id --exact --source winget --accept-source-agreements --accept-package-agreements --disable-interactivity --silent
        } catch {
            Write-Host ("Failed to install {0}: {1}" -f $app.Name, $_.Exception.Message) -ForegroundColor Red
        }
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
    $choiceTrim = if ($choice) { $choice.Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($choiceTrim)) {
        Write-Host 'Invalid selection.' -ForegroundColor Yellow
        return
    }
    if ($choiceTrim -match '^[Aa]$') {
        $selectedFiles = $files
    } else {
        $indices = @()
        foreach ($p in ($choiceTrim -split ',')) {
            $pt = $p.Trim()
            if ([int]::TryParse($pt, [ref]$null)) {
                $n = [int]$pt
                if ($n -ge 1 -and $n -le $files.Count -and -not $indices.Contains($n)) {
                    $indices += $n
                }
            }
        }
        if ($indices.Count -gt 0) {
            foreach ($n in $indices) { $selectedFiles += $files[$n - 1] }
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
    if ($original) {
        $pattern = '(?s)# ----- BEGIN winutil blocklists -----.*?# ----- END winutil blocklists -----\s*'
        $original = [System.Text.RegularExpressions.Regex]::Replace($original, $pattern, '')
    }
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


