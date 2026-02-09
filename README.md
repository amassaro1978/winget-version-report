# Winget Version Report

A PowerShell script that queries the Windows Package Manager (winget) for the latest available versions of a curated list of enterprise applications and generates a styled HTML report.

## Features

- 📊 **Dark-themed HTML report** — opens automatically in your default browser
- 📦 **Latest & previous versions** from winget
- 📅 **Release dates** when available
- 🔗 **Download links** — MSI preferred, EXE as fallback
- 🔐 **SHA256 hashes** for installer verification
- 🏗️ **Architecture detection** (x64, x86, arm64)
- 🔵 **Adobe Reader MSP patch** — special handling to pull the MUI MSP patch directly from Adobe
- 📄 **CSV export** option
- ✅ Easy to customize — just edit the app ID list at the top

## Usage

```powershell
# Basic — generates HTML report and opens it
.\winget-version-report.ps1

# With CSV export
.\winget-version-report.ps1 -ExportCsv "report.csv"

# Custom report path
.\winget-version-report.ps1 -ReportPath "C:\Reports\winget.html"
```

## Adding/Removing Apps

Edit the `$AppIDs` array at the top of the script:

```powershell
$AppIDs = @(
    "Google.Chrome"
    "Mozilla.Firefox.ESR"
    "YourApp.Id.Here"
)
```

Find app IDs with: `winget search <app name>`

## Requirements

- Windows 10/11 with [winget](https://aka.ms/getwinget) installed
- PowerShell 5.1+

## Screenshot

The report includes:
- Summary stats (apps scanned, successful, errors)
- Sortable table with version info, release dates, publishers
- Color-coded download badges (green=MSI, gold=EXE, blue=MSP)
- Truncated SHA256 with full hash on hover

## License

Open
