# PowerShell script to install dependencies on Windows
#
# Usage: .\install_deps.ps1
#
# This script:
#   1. Installs Python dependencies from requirements.txt
#   2. Checks for readpst availability (WSL2, Cygwin, native)
#   3. Offers to download native Windows binary if needed

param(
    [switch]$DownloadReadpst,
    [string]$InstallPath = "$env:USERPROFILE\Tools\libpst"
)

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "PST to Gmail Migration - Windows Setup" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# Check Python
# =============================================================================
Write-Host "Checking Python installation..." -ForegroundColor White
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    $python = Get-Command python3 -ErrorAction SilentlyContinue
}

if (-not $python) {
    Write-Host "  ERROR: Python not found!" -ForegroundColor Red
    Write-Host "  Please install Python from https://python.org"
    exit 1
}

Write-Host "  Python found: $($python.Source)" -ForegroundColor Green
Write-Host ""

# =============================================================================
# Install Python dependencies
# =============================================================================
Write-Host "Installing Python dependencies..." -ForegroundColor White
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RequirementsFile = Join-Path (Split-Path -Parent $ScriptDir) "requirements.txt"

if (Test-Path $RequirementsFile) {
    try {
        & pip install -r $RequirementsFile 2>&1 | Out-Null
        Write-Host "  Python dependencies installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "  Warning: Error installing some dependencies: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "  Warning: requirements.txt not found, installing GYB directly..." -ForegroundColor Yellow
    & pip install got-your-back 2>&1 | Out-Null
}

# Verify GYB
$gyb = Get-Command gyb -ErrorAction SilentlyContinue
if ($gyb) {
    Write-Host "  GYB installed: $($gyb.Source)" -ForegroundColor Green
} else {
    # Try as Python module
    $gybCheck = & python -m gyb --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  GYB installed as Python module" -ForegroundColor Green
    } else {
        Write-Host "  Warning: GYB may not be installed correctly" -ForegroundColor Yellow
    }
}

Write-Host ""

# =============================================================================
# Check for readpst
# =============================================================================
Write-Host "Checking for readpst..." -ForegroundColor White

$readpstFound = $false
$readpstPath = $null

# Check 1: Native Windows (in PATH or common locations)
$readpst = Get-Command readpst -ErrorAction SilentlyContinue
if ($readpst) {
    Write-Host "  readpst found (native): $($readpst.Source)" -ForegroundColor Green
    $readpstFound = $true
    $readpstPath = $readpst.Source
}

# Check 2: ezwinports location
if (-not $readpstFound) {
    $ezwinportsPath = "$InstallPath\bin\readpst.exe"
    if (Test-Path $ezwinportsPath) {
        Write-Host "  readpst found (ezwinports): $ezwinportsPath" -ForegroundColor Green
        $readpstFound = $true
        $readpstPath = $ezwinportsPath
    }
}

# Check 3: Cygwin
if (-not $readpstFound) {
    $cygwinPaths = @(
        "C:\cygwin64\bin\readpst.exe",
        "C:\cygwin\bin\readpst.exe"
    )
    foreach ($path in $cygwinPaths) {
        if (Test-Path $path) {
            Write-Host "  readpst found (Cygwin): $path" -ForegroundColor Green
            $readpstFound = $true
            $readpstPath = $path
            break
        }
    }
}

# Check 4: WSL2
if (-not $readpstFound) {
    try {
        $wslCheck = wsl which readpst 2>&1
        if ($LASTEXITCODE -eq 0 -and $wslCheck -match "/usr") {
            Write-Host "  readpst found (WSL2): $wslCheck" -ForegroundColor Green
            $readpstFound = $true
            $readpstPath = "wsl readpst"
        }
    } catch {
        # WSL not available
    }
}

Write-Host ""

# =============================================================================
# If readpst not found, offer options
# =============================================================================
if (-not $readpstFound) {
    Write-Host "==========================================" -ForegroundColor Yellow
    Write-Host "readpst NOT FOUND" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "readpst is required to convert PST files. Options:" -ForegroundColor White
    Write-Host ""
    Write-Host "Option 1: Download Native Binary (Quick)" -ForegroundColor Cyan
    Write-Host "  Run: .\install_deps.ps1 -DownloadReadpst"
    Write-Host "  Downloads pre-built Windows binary from ezwinports"
    Write-Host ""
    Write-Host "Option 2: Use WSL2 (Recommended for best compatibility)" -ForegroundColor Cyan
    Write-Host "  1. Run: wsl --install"
    Write-Host "  2. Restart your computer"
    Write-Host "  3. In WSL terminal: sudo apt install pst-utils"
    Write-Host ""
    Write-Host "Option 3: Use Cygwin" -ForegroundColor Cyan
    Write-Host "  1. Install Cygwin from https://cygwin.com/install.html"
    Write-Host "  2. During setup, select the 'readpst' package"
    Write-Host ""
    Write-Host "Option 4: Use Thunderbird (Manual conversion)" -ForegroundColor Cyan
    Write-Host "  See SETUP_READPST.md for instructions"
    Write-Host ""

    # Auto-download if flag is set
    if ($DownloadReadpst) {
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host "Downloading readpst from ezwinports..." -ForegroundColor White
        Write-Host "==========================================" -ForegroundColor Cyan

        $downloadUrl = "https://sourceforge.net/projects/ezwinports/files/libpst-0.6.63-w32-bin.zip/download"
        $zipPath = "$env:TEMP\libpst-w32-bin.zip"

        try {
            # Download
            Write-Host "  Downloading..." -ForegroundColor White
            Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing

            # Extract
            Write-Host "  Extracting to $InstallPath..." -ForegroundColor White
            if (Test-Path $InstallPath) {
                Remove-Item -Recurse -Force $InstallPath
            }
            Expand-Archive -Path $zipPath -DestinationPath $InstallPath -Force

            # Clean up
            Remove-Item $zipPath -Force

            # Verify
            $readpstExe = "$InstallPath\bin\readpst.exe"
            if (Test-Path $readpstExe) {
                Write-Host "  readpst installed successfully!" -ForegroundColor Green
                Write-Host "  Location: $readpstExe" -ForegroundColor Green
                Write-Host ""
                Write-Host "  To use, either:" -ForegroundColor White
                Write-Host "    1. Add to PATH: $InstallPath\bin"
                Write-Host "    2. Use full path: $readpstExe"
                Write-Host ""
                Write-Host "  Example:" -ForegroundColor White
                Write-Host "    & '$readpstExe' -S -e -o .\output 'backup.pst'"
                $readpstFound = $true
                $readpstPath = $readpstExe
            } else {
                Write-Host "  ERROR: Installation failed" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ERROR: Download failed: $_" -ForegroundColor Red
            Write-Host "  Try downloading manually from:" -ForegroundColor Yellow
            Write-Host "  https://sourceforge.net/projects/ezwinports/files/libpst-0.6.63-w32-bin.zip/"
        }
    }
} else {
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host "SETUP COMPLETE" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Set up GYB authentication:" -ForegroundColor White
Write-Host "     gyb --email your.email@gmail.com --action check"
Write-Host ""
Write-Host "  2. Test with a dry run:" -ForegroundColor White
Write-Host "     python pst_to_gmail.py backup.pst --email EMAIL --dry-run"
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
