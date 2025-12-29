#!/bin/bash
#
# Install dependencies for PST to Gmail migration in WSL2
# Run this script from within your WSL2 Ubuntu environment
#

set -e

echo "=========================================="
echo "PST to Gmail Migration - WSL2 Setup"
echo "=========================================="
echo

# Check if running in WSL
if ! grep -qi microsoft /proc/version 2>/dev/null; then
    echo "Warning: This script is designed for WSL2."
    echo "If you're on native Linux, use install_deps.sh instead."
    read -p "Continue anyway? (y/n) " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi

echo "Updating package lists..."
sudo apt update

echo
echo "Installing pst-utils (readpst)..."
sudo apt install -y pst-utils

echo
echo "Installing Python and pip..."
sudo apt install -y python3 python3-pip

echo
echo "Installing GYB (Got Your Back)..."
pip3 install --user got-your-back

# Add local bin to PATH if not already there
if [[ ":$PATH:" != *":$HOME/.local/bin:"* ]]; then
    echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc
    export PATH="$HOME/.local/bin:$PATH"
    echo "Added ~/.local/bin to PATH"
fi

echo
echo "=========================================="
echo "Verifying installations..."
echo "=========================================="

# Check readpst
if command -v readpst &> /dev/null; then
    echo "✓ readpst: $(readpst --version 2>&1 | head -1)"
else
    echo "✗ readpst not found"
fi

# Check GYB
if command -v gyb &> /dev/null; then
    echo "✓ GYB: $(gyb --version 2>&1 | head -1)"
elif python3 -m gyb --version &> /dev/null 2>&1; then
    echo "✓ GYB (as Python module)"
else
    echo "✗ GYB not found"
fi

echo
echo "=========================================="
echo "WSL2 Setup Complete!"
echo "=========================================="
echo
echo "Your Windows files are accessible at /mnt/c/Users/<username>/"
echo
echo "Example usage:"
echo "  cd /mnt/c/Users/YourName/Documents"
echo "  python3 pst_to_gmail.py 'Outlook Backup.pst' --email your@gmail.com"
echo
echo "Next steps:"
echo "  1. Set up GYB authentication:"
echo "     gyb --email your.email@gmail.com --action check"
echo
echo "  2. Navigate to your PST file location:"
echo "     cd /mnt/c/Users/YourName/Documents"
echo
echo "  3. Run the migration:"
echo "     python3 pst_to_gmail.py backup.pst --email your@gmail.com"
echo "=========================================="
