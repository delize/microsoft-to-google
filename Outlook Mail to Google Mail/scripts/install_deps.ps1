# PowerShell script to install dependencies on Windows
# Note: readpst is not available on Windows natively - use WSL2 instead
#
# Usage: .\install_deps.ps1

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "PST to Gmail Migration - Windows Setup" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
Write-Host "Checking Python installation..."
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    $python = Get-Command python3 -ErrorAction SilentlyContinue
}

if (-not $python) {
    Write-Host "Error: Python not found!" -ForegroundColor Red
    Write-Host "Please install Python from https://python.org"
    exit 1
}

Write-Host "Python found: $($python.Source)" -ForegroundColor Green
Write-Host ""

# Install Python dependencies from requirements.txt
Write-Host "Installing Python dependencies from requirements.txt..."
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RequirementsFile = Join-Path (Split-Path -Parent $ScriptDir) "requirements.txt"

if (Test-Path $RequirementsFile) {
    try {
        & pip install -r $RequirementsFile
        Write-Host "Python dependencies installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Error installing dependencies: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Warning: requirements.txt not found at $RequirementsFile" -ForegroundColor Yellow
    Write-Host "Installing GYB directly..."
    try {
        & pip install got-your-back
        Write-Host "GYB installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Error installing GYB: $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "IMPORTANT: readpst is not available on Windows" -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "To convert PST files, you have two options:" -ForegroundColor White
Write-Host ""
Write-Host "Option 1: Use WSL2 (Recommended)" -ForegroundColor Cyan
Write-Host "  1. Install WSL2: wsl --install"
Write-Host "  2. Restart your computer"
Write-Host "  3. Run: .\install_deps_wsl.sh from WSL"
Write-Host ""
Write-Host "Option 2: Use Thunderbird (Manual)" -ForegroundColor Cyan
Write-Host "  1. Install Mozilla Thunderbird"
Write-Host "  2. Import your PST file"
Write-Host "  3. Export as MBOX"
Write-Host "  4. Use MBOX files with this tool"
Write-Host ""
Write-Host "For pre-converted EML/MBOX files, you can use GYB directly:"
Write-Host "  python pst_to_gmail.py .\emails\ --email your@gmail.com"
Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
