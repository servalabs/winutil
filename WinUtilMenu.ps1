$ErrorActionPreference = 'Stop'

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
    Write-Host '1) Install Apps'
    Write-Host '2) Network Blocklists (hosts)'
    Write-Host '3) Clean Windows (Disk Cleanup)'
    Write-Host '4) Improve Privacy (O&O ShutUp10)'
    Write-Host '0) Exit'
}

function Invoke-AppInstallsSubmenu {
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
        $i = 1
        foreach ($c in $categories) {
            Write-Host ("{0}) {1}" -f $i, $c)
            $i++
        }
        Write-Host ("{0}) Install ALL categories" -f $i)
        Write-Host '0) Back'
        Write-Host -NoNewline 'Enter number(s) comma-separated, A for all, or 0 to back: '
        $raw = Read-Host
        if ($raw -eq '0') { return }
        if ($raw -match '^[Aa]$') {
            & (Join-Path $PSScriptRoot 'AppInstalls.ps1') -All
        } else {
            $sel = Parse-MultiSelection -Input $raw -Max ($i - 1)
            if (-not $sel -or $sel.Count -eq 0) {
                Write-Host 'Invalid selection.' -ForegroundColor Yellow
                continue
            }
            $chosen = @()
            foreach ($idx in $sel) { $chosen += $categories[$idx - 1] }
            & (Join-Path $PSScriptRoot 'AppInstalls.ps1') -Categories $chosen
        }
        Write-Host 'Done. Press Enter to continue...'
        [void][System.Console]::ReadLine()
    }
}

function Get-HostsBlocklistFiles {
    $dir = Join-Path $PSScriptRoot 'hosts'
    if (-not (Test-Path $dir)) { return @() }
    Get-ChildItem -Path $dir -File | Sort-Object Name
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
        Write-Host ("{0}) {1}" -f ($i+1), $files[$i].Name)
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
        [void]$merged.AppendLine("# from hosts/" + $f.Name)
        $content = Get-Content -Path $f.FullName -Raw
        [void]$merged.AppendLine($content.TrimEnd())
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

    Write-Host 'O&O ShutUp10 not found. Attempting installation via winget...' -ForegroundColor Cyan
    try {
        winget install --id OandO.ShutUp10 --exact --accept-source-agreements --accept-package-agreements --silent --disable-interactivity
        if ($LASTEXITCODE -ne 0) {
            Write-Host "winget exit code: $LASTEXITCODE" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "winget installation failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    $exe = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($exe) {
        try { Start-Process -FilePath $exe -Verb RunAs } catch {}
    } else {
        Write-Host 'Please install O&O ShutUp10 manually from https://www.oo-software.com/en/shutup10' -ForegroundColor Yellow
    }
}

while ($true) {
    Show-MainMenu
    $choice = Read-MenuSelection -Min 0 -Max 4
    switch ($choice) {
        1 { Invoke-AppInstallsSubmenu }
        2 { Invoke-HostsMenu }
        3 { Invoke-CleanWindows }
        4 { Invoke-PrivacyImprovements }
        0 { break }
    }
    if ($choice -ne 0) {
        Write-Host 'Press Enter to return to main menu...'
        [void][System.Console]::ReadLine()
    }
}


