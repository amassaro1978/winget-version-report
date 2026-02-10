#Requires -Version 5.1
<#
.SYNOPSIS
    Winget Version Report — checks latest available versions for a list of app IDs.
.DESCRIPTION
    Edit the $AppIDs array below to add/remove applications.
    Run: .\winget-version-report.ps1
    Opens a styled HTML report in your default browser automatically.
    Optional: .\winget-version-report.ps1 -ExportCsv "report.csv"
#>

param(
    [string]$ExportCsv,
    [string]$ReportPath = "$env:TEMP\WingetReport.html"
)

# ===== EDIT THIS LIST =====
$AppIDs = @(
    "Google.Chrome"
    "Mozilla.Firefox.ESR"
    "7zip.7zip"
    "Zoom.Zoom"
    "Cisco.Webex"
    "Notepad++.Notepad++"
    "WinSCP.WinSCP"
    "GIMP.GIMP.3"
    "Microsoft.DotNet.SDK.8"
    "GNU.Emacs"
    "Adobe.Acrobat.Reader.32-bit"
    "Oracle.VirtualBox"
    "WiresharkFoundation.Wireshark"
    "VideoLAN.VLC"
    "VMware.WorkstationPro"
    "Anaconda.Anaconda3"
    "SlackTechnologies.Slack"
    "Omnissa.HorizonClient"
    "NVAccess.NVDA"
    "Git.Git"
    "Microsoft.VisualStudioCode"
    "Docker.DockerDesktop"
    "TechSmith.Snagit.2023"
    "JetBrains.PyCharm.Community"
    "GNU.Octave"
    "RProject.R"
)
# ===========================

function Get-AdobeReaderMspUrl {
    param([string]$Version)
    # Adobe MSP URL pattern: version like 25.001.20997 becomes 2500120997
    # Full installer: AcroRdrDC{versionFlat}_MUI.exe
    # MSP patch:      AcroRdrDCUpd{versionFlat}_MUI.msp
    # URL: https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/{versionFlat}/AcroRdrDCUpd{versionFlat}_MUI.msp
    
    if (-not $Version -or $Version -eq 'N/A') { return $null }
    
    # Convert version to flat format (remove dots)
    $flat = $Version -replace '\.', ''
    
    $mspUrl = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/$flat/AcroRdrDCUpd${flat}_MUI.msp"
    
    # Verify URL is reachable
    try {
        $response = Invoke-WebRequest -Uri $mspUrl -Method Head -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        if ($response.StatusCode -eq 200) {
            return @{ Url = $mspUrl; Type = "MSP" }
        }
    } catch {}
    
    # Fallback: try the full MUI EXE installer
    $exeUrl = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/$flat/AcroRdrDC${flat}_MUI.exe"
    try {
        $response = Invoke-WebRequest -Uri $exeUrl -Method Head -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        if ($response.StatusCode -eq 200) {
            return @{ Url = $exeUrl; Type = "EXE" }
        }
    } catch {}
    
    return $null
}

Write-Host ""
Write-Host "  Winget Version Report" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "  Scanning $($AppIDs.Count) applications..." -ForegroundColor Gray
Write-Host ""

$results = @()

foreach ($id in $AppIDs) {
    Write-Host "  Checking $id ..." -NoNewline -ForegroundColor Yellow

    try {
        # Get latest version info
        $lines = & winget show --id $id --accept-source-agreements 2>$null

        $version = "N/A"
        $publisher = "N/A"
        $name = $id
        $releaseDate = "N/A"
        $releaseUrl = ""
        $description = ""
        $homepage = ""
        $downloadUrl = ""
        $installerType = ""
        $sha256 = ""
        $architecture = ""

        foreach ($line in $lines) {
            if ($null -eq $line) { continue }
            $s = $line.ToString().Trim()

            if ($s -like "Version:*") {
                $version = ($s -replace '^Version:\s*','').Trim()
            }
            elseif ($s -like "Publisher:*") {
                $publisher = ($s -replace '^Publisher:\s*','').Trim()
            }
            elseif ($s -like "Release Date:*") {
                $releaseDate = ($s -replace '^Release Date:\s*','').Trim()
            }
            elseif ($s -like "Release Notes Url:*") {
                $releaseUrl = ($s -replace '^Release Notes Url:\s*','').Trim()
            }
            elseif ($s -like "Description:*") {
                $description = ($s -replace '^Description:\s*','').Trim()
            }
            elseif ($s -like "Homepage:*") {
                $homepage = ($s -replace '^Homepage:\s*','').Trim()
            }
            elseif ($s -like "Installer Url:*") {
                $url = ($s -replace '^Installer Url:\s*','').Trim()
                # Detect installer type from URL (check MSIX before MSI to avoid false match)
                if ($url -match '\.(msix|msixbundle|appx|appxbundle)(\?|$)') {
                    $downloadUrl = $url
                    $installerType = "MSIX"
                }
                elseif ($url -match '\.msi(\?|$)' -and $installerType -ne 'MSIX') {
                    $downloadUrl = $url
                    $installerType = "MSI"
                }
                elseif (-not $downloadUrl -or $downloadUrl -eq '') {
                    $downloadUrl = $url
                    if ($url -match '\.exe') { $installerType = "EXE" }
                    elseif ($url -match '\.zip') { $installerType = "ZIP" }
                    else { $installerType = "Other" }
                }
            }
            elseif ($s -like "Installer Type:*") {
                $iType = ($s -replace '^Installer Type:\s*','').Trim().ToUpper()
                # MSIX type from winget always wins (URL might show .msi for msix packages)
                if ($iType -eq 'MSIX' -or $iType -eq 'APPX') {
                    $installerType = "MSIX"
                }
                elseif (-not $installerType -or $installerType -eq 'Other') { $installerType = $iType }
            }
            elseif ($s -like "Installer SHA256:*" -or $s -like "SHA256:*") {
                if (-not $sha256) {
                    $sha256 = ($s -replace '^(Installer )?SHA256:\s*','').Trim()
                }
            }
            elseif ($s -like "Architecture:*" -or $s -like "Installer Architecture:*") {
                $arch = ($s -replace '^(Installer )?Architecture:\s*','').Trim()
                if ($arch -and $architecture -notmatch [regex]::Escape($arch)) {
                    if ($architecture) { $architecture = "$architecture, $arch" }
                    else { $architecture = $arch }
                }
            }
            elseif ($s -match '^Found\s+(.+)\s+\[') {
                $name = $Matches[1].Trim()
            }
        }

        # Special case: Adobe Reader — get MSP patch URL
        if ($id -like "Adobe.Acrobat.Reader*" -and $version -ne "N/A") {
            Write-Host "" 
            Write-Host "    Checking Adobe MSP patch..." -NoNewline -ForegroundColor Gray
            $adobeResult = Get-AdobeReaderMspUrl -Version $version
            if ($adobeResult) {
                $downloadUrl = $adobeResult.Url
                $installerType = $adobeResult.Type
                Write-Host " Found $($adobeResult.Type)" -ForegroundColor Green
            } else {
                Write-Host " Not found (using winget URL)" -ForegroundColor Yellow
            }
        }

        # If no architecture found, try to infer from installer URL
        if (-not $architecture -and $downloadUrl) {
            if ($downloadUrl -match 'x64|amd64|win64') { $architecture = "x64" }
            elseif ($downloadUrl -match 'x86|win32|i386') { $architecture = "x86" }
            elseif ($downloadUrl -match 'arm64|aarch64') { $architecture = "arm64" }
        }

        # Get previous version from version list
        $previousVersion = "N/A"
        try {
            $verOutput = & winget show --id $id --versions --accept-source-agreements 2>$null
            $versionList = @()
            $headerFound = $false
            foreach ($vl in $verOutput) {
                if ($null -eq $vl) { continue }
                $vs = $vl.ToString().Trim()
                # Skip empty lines
                if ($vs -eq '') { continue }
                # Look for the dashes separator line
                if ($vs -match '^-') { $headerFound = $true; continue }
                # Skip the "Version" header text itself
                if ($vs -eq 'Version') { continue }
                if ($vs -match '^Found' -or $vs -match '^Name' -or $vs -match '^Available') { continue }
                if ($headerFound) {
                    $versionList += $vs
                }
            }
            # First entry = latest, second = previous
            if ($versionList.Count -ge 2) {
                $previousVersion = $versionList[1]
            }
        } catch {}

        $results += [PSCustomObject]@{
            AppID           = $id
            Name            = $name
            Version         = $version
            PreviousVersion = $previousVersion
            Publisher       = $publisher
            ReleaseDate     = $releaseDate
            ReleaseNotes    = $releaseUrl
            Description     = $description
            Homepage        = $homepage
            DownloadUrl     = $downloadUrl
            InstallerType   = $installerType
            SHA256          = $sha256
            Architecture    = $architecture
        }

        Write-Host " $version" -ForegroundColor Green
    }
    catch {
        $results += [PSCustomObject]@{
            AppID           = $id
            Name            = "ERROR"
            Version         = "ERROR"
            PreviousVersion = "N/A"
            Publisher       = "N/A"
            ReleaseDate     = "N/A"
            ReleaseNotes    = ""
            Description     = ""
            Homepage        = ""
            DownloadUrl     = ""
            InstallerType   = ""
            SHA256          = ""
            Architecture    = ""
        }
        Write-Host " Error: $_" -ForegroundColor Red
    }
}

# Console output
Write-Host ""
$results | Format-Table -AutoSize -Property `
    @{Label="App ID"; Expression={$_.AppID}},
    @{Label="Name"; Expression={$_.Name}},
    @{Label="Version"; Expression={$_.Version}},
    @{Label="Previous"; Expression={$_.PreviousVersion}},
    @{Label="Release Date"; Expression={$_.ReleaseDate}},
    @{Label="Publisher"; Expression={$_.Publisher}}

# Build HTML report
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
$errorCount = ($results | Where-Object {$_.Version -eq 'ERROR'}).Count

$tableRows = ""
$i = 0
foreach ($r in $results) {
    $i++
    $rowClass = if ($i % 2 -eq 0) { 'even' } else { 'odd' }
    $statusIcon = if ($r.Version -eq 'ERROR') { '&#10060;' } else { '&#9989;' }
    $dateDisplay = if ($r.ReleaseDate -eq 'N/A') { '<span class="na">N/A</span>' } else { $r.ReleaseDate }
    $prevDisplay = if ($r.PreviousVersion -eq 'N/A') { '<span class="na">N/A</span>' } else { $r.PreviousVersion }

    $nameDisplay = $r.Name
    if ($r.Homepage) {
        $nameDisplay = "<a href='$($r.Homepage)' target='_blank'>$($r.Name)</a>"
    }

    $notesLink = ""
    if ($r.ReleaseNotes) {
        $notesLink = "<a href='$($r.ReleaseNotes)' target='_blank' title='Release Notes'>&#128196;</a>"
    }

    $downloadLink = ""
    if ($r.DownloadUrl) {
        $badge = $r.InstallerType
        if (-not $badge) { $badge = "DL" }
        $badgeClass = switch ($r.InstallerType) { 'MSI' { 'badge-msi' } 'MSP' { 'badge-msp' } 'ZIP' { 'badge-zip' } 'MSIX' { 'badge-msix' } default { 'badge-exe' } }
        $downloadLink = "<a href='$($r.DownloadUrl)' target='_blank' class='dl-badge $badgeClass'>&#11015; $badge</a>"
    } else {
        $downloadLink = '<span class="na">N/A</span>'
    }

    $shaDisplay = if ($r.SHA256) { "<span class='sha' title='$($r.SHA256)'>$($r.SHA256.Substring(0, [Math]::Min(16, $r.SHA256.Length)))&hellip;</span>" } else { '<span class="na">N/A</span>' }

    $tableRows += @"
    <tr class="$rowClass">
        <td>$statusIcon</td>
        <td><strong>$nameDisplay</strong><br><span class="app-id">$($r.AppID)</span></td>
        <td class="version">$($r.Version)</td>
        <td class="prev-version">$prevDisplay</td>
        <td>$dateDisplay</td>
        <td>$($r.Publisher)</td>
        <td class="center arch">$($r.Architecture)</td>
        <td class="center">$downloadLink</td>
        <td>$shaDisplay</td>
        <td class="center">$notesLink</td>
    </tr>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Winget Version Report</title>
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: #0f172a;
        color: #e2e8f0;
        padding: 40px;
    }
    .container { max-width: 1400px; margin: 0 auto; }
    .header {
        background: linear-gradient(135deg, #1e3a5f, #2563eb);
        border-radius: 16px 16px 0 0;
        padding: 30px 40px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .header h1 { font-size: 24px; font-weight: 600; }
    .header .meta { text-align: right; font-size: 14px; opacity: 0.8; }
    .stats {
        display: flex;
        gap: 20px;
        padding: 20px 40px;
        background: #1e293b;
        border-bottom: 1px solid #334155;
    }
    .stat-card {
        background: #0f172a;
        border-radius: 10px;
        padding: 15px 25px;
        flex: 1;
        text-align: center;
    }
    .stat-card .num { font-size: 28px; font-weight: 700; color: #60a5fa; }
    .stat-card .label { font-size: 12px; text-transform: uppercase; color: #94a3b8; margin-top: 4px; }
    table {
        width: 100%;
        border-collapse: collapse;
        background: #1e293b;
    }
    th {
        background: #334155;
        padding: 14px 16px;
        text-align: left;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        color: #94a3b8;
    }
    td {
        padding: 14px 16px;
        border-bottom: 1px solid #293548;
        font-size: 14px;
    }
    tr.even { background: #1a2536; }
    tr:hover { background: #253349; }
    .version {
        font-family: 'Cascadia Code', 'Consolas', monospace;
        font-weight: 600;
        color: #34d399;
        font-size: 15px;
    }
    .prev-version {
        font-family: 'Cascadia Code', 'Consolas', monospace;
        color: #94a3b8;
        font-size: 13px;
    }
    .app-id {
        font-size: 11px;
        color: #64748b;
        font-family: monospace;
    }
    .na { color: #475569; font-style: italic; }
    .sha {
        font-family: 'Cascadia Code', 'Consolas', monospace;
        font-size: 11px;
        color: #818cf8;
        cursor: pointer;
    }
    .sha:hover { color: #a5b4fc; }
    .arch { font-size: 12px; color: #93c5fd; font-weight: 600; }
    .center { text-align: center; }
    a { color: #60a5fa; text-decoration: none; }
    a:hover { text-decoration: underline; }
    .dl-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 600;
        text-decoration: none !important;
    }
    .badge-msi {
        background: #065f46;
        color: #6ee7b7;
    }
    .badge-msi:hover { background: #047857; }
    .badge-exe {
        background: #713f12;
        color: #fcd34d;
    }
    .badge-exe:hover { background: #854d0e; }
    .badge-msp {
        background: #1e3a5f;
        color: #7dd3fc;
    }
    .badge-msp:hover { background: #1e4a7f; }
    .badge-zip {
        background: #5b21b6;
        color: #c4b5fd;
    }
    .badge-zip:hover { background: #6d28d9; }
    .badge-msix {
        background: #831843;
        color: #f9a8d4;
    }
    .badge-msix:hover { background: #9d174d; }
    .footer {
        background: #1e293b;
        border-radius: 0 0 16px 16px;
        padding: 20px 40px;
        text-align: center;
        font-size: 12px;
        color: #64748b;
    }
</style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>&#128230; Winget Version Report</h1>
        <div class="meta">
            Generated: $timestamp<br>
            Machine: $env:COMPUTERNAME
        </div>
    </div>
    <div class="stats">
        <div class="stat-card">
            <div class="num">$($results.Count)</div>
            <div class="label">Apps Scanned</div>
        </div>
        <div class="stat-card">
            <div class="num">$($results.Count - $errorCount)</div>
            <div class="label">Successful</div>
        </div>
        <div class="stat-card">
            <div class="num">$errorCount</div>
            <div class="label">Errors</div>
        </div>
    </div>
    <table>
        <thead>
            <tr>
                <th width="40"></th>
                <th>Application</th>
                <th>Latest Version</th>
                <th>Previous Version</th>
                <th>Release Date</th>
                <th>Publisher</th>
                <th>Arch</th>
                <th width="80">Download</th>
                <th>SHA256</th>
                <th width="50">Notes</th>
            </tr>
        </thead>
        <tbody>
            $tableRows
        </tbody>
    </table>
    <div class="footer">
        Winget Version Report &mdash; Last checked: $timestamp &mdash; Powered by PowerShell + Winget
    </div>
</div>
</body>
</html>
"@

$html | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Host ""
Write-Host "  Report saved to: $ReportPath" -ForegroundColor Green
Write-Host "  Opening in browser..." -ForegroundColor Gray
Start-Process $ReportPath

if ($ExportCsv) {
    $results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "  CSV exported to: $ExportCsv" -ForegroundColor Green
}
