#!/bin/bash
#
# Install dependencies for PST to Gmail migration tool
# Supports macOS and Linux (Ubuntu/Debian, Fedora/RHEL)
#

set -e

echo "=========================================="
echo "PST to Gmail Migration - Dependency Setup"
echo "=========================================="

# Detect OS
OS="unknown"
if [[ "$OSTYPE" == "darwin"* ]]; then
    OS="macos"
elif [[ -f /etc/debian_version ]]; then
    OS="debian"
elif [[ -f /etc/redhat-release ]]; then
    OS="redhat"
elif [[ -f /etc/arch-release ]]; then
    OS="arch"
fi

echo "Detected OS: $OS"
echo

# Install readpst (libpst)
echo "Installing readpst (libpst)..."
case $OS in
    macos)
        if ! command -v brew &> /dev/null; then
            echo "Error: Homebrew not found. Install it from https://brew.sh"
            exit 1
        fi
        brew install libpst
        ;;
    debian)
        sudo apt update
        sudo apt install -y pst-utils
        ;;
    redhat)
        sudo dnf install -y libpst || sudo yum install -y libpst
        ;;
    arch)
        sudo pacman -S --noconfirm libpst
        ;;
    *)
        echo "Warning: Unknown OS. Please install libpst manually."
        echo "  See: https://www.five-ten-sg.com/libpst/"
        ;;
esac

# Verify readpst installation
if command -v readpst &> /dev/null; then
    echo "✓ readpst installed: $(readpst --version 2>&1 | head -1)"
else
    echo "✗ readpst installation failed"
fi

echo

# Get script directory to find requirements.txt
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REQUIREMENTS_FILE="$SCRIPT_DIR/../requirements.txt"

# Install Python dependencies from requirements.txt
echo "Installing Python dependencies from requirements.txt..."
if [[ -f "$REQUIREMENTS_FILE" ]]; then
    if command -v pip3 &> /dev/null; then
        pip3 install --user -r "$REQUIREMENTS_FILE"
    elif command -v pip &> /dev/null; then
        pip install --user -r "$REQUIREMENTS_FILE"
    else
        echo "Error: pip not found. Install Python first."
        exit 1
    fi
else
    echo "Warning: requirements.txt not found at $REQUIREMENTS_FILE"
    echo "Installing GYB directly..."
    pip3 install --user got-your-back || pip install --user got-your-back
fi

# Verify GYB installation
if command -v gyb &> /dev/null; then
    echo "✓ GYB installed: $(gyb --version 2>&1 | head -1)"
elif python3 -m gyb --version &> /dev/null; then
    echo "✓ GYB installed as Python module"
else
    echo "✗ GYB installation may have failed"
    echo "  Try: pip install -r requirements.txt"
fi

echo
echo "=========================================="
echo "Installation complete!"
echo
echo "Next steps:"
echo "  1. Set up GYB authentication:"
echo "     gyb --email your.email@gmail.com --action check"
echo
echo "  2. Test with a dry run:"
echo "     python pst_to_gmail.py backup.pst --email EMAIL --dry-run"
echo "=========================================="
