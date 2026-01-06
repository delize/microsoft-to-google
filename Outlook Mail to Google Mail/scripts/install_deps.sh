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

# Get script and project directories
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

# Install GYB (Got Your Back)
# GYB is distributed as standalone builds (not on PyPI)
echo "Installing GYB (Got Your Back)..."

GYB_INSTALLED=false
GYB_DIR="$PROJECT_DIR/gyb"

# Check if GYB is already installed globally
if command -v gyb &> /dev/null; then
    echo "✓ GYB already installed globally"
    GYB_INSTALLED=true
fi

# Check if GYB is already installed locally
if [[ "$GYB_INSTALLED" == "false" ]] && [[ -f "$GYB_DIR/gyb" ]]; then
    echo "✓ GYB already installed locally at $GYB_DIR"
    GYB_INSTALLED=true
fi

# Download standalone build
if [[ "$GYB_INSTALLED" == "false" ]]; then
    echo "Downloading GYB standalone build..."

    # Get latest version from GitHub API
    echo "Fetching latest version..."
    GYB_VERSION=$(curl -s "https://api.github.com/repos/GAM-team/got-your-back/releases/latest" | grep '"tag_name"' | sed -E 's/.*"v([^"]+)".*/\1/')

    if [[ -z "$GYB_VERSION" ]]; then
        echo "Warning: Could not determine latest version, using fallback"
        GYB_VERSION="1.95"
    fi

    echo "Latest version: $GYB_VERSION"

    # Determine platform and architecture
    case $OS in
        macos)
            ARCH=$(uname -m)
            if [[ "$ARCH" == "arm64" ]]; then
                GYB_FILE="gyb-${GYB_VERSION}-macos-aarch64.tar.xz"
            else
                GYB_FILE="gyb-${GYB_VERSION}-macos-x86_64.tar.xz"
            fi
            ;;
        debian|redhat|arch)
            ARCH=$(uname -m)
            if [[ "$ARCH" == "aarch64" ]]; then
                GYB_FILE="gyb-${GYB_VERSION}-linux-aarch64-glibc2.35.tar.xz"
            else
                GYB_FILE="gyb-${GYB_VERSION}-linux-x86_64-glibc2.35.tar.xz"
            fi
            ;;
        *)
            echo "Warning: Unknown OS for GYB download. Please install manually."
            echo "  See: https://github.com/GAM-team/got-your-back/releases"
            ;;
    esac

    if [[ -n "$GYB_FILE" ]]; then
        GYB_URL="https://github.com/GAM-team/got-your-back/releases/download/v${GYB_VERSION}/${GYB_FILE}"

        echo "Downloading from: $GYB_URL"

        # Create temp directory for download
        TMP_DIR=$(mktemp -d)
        cd "$TMP_DIR"

        if curl -L --fail -o "$GYB_FILE" "$GYB_URL"; then
            echo "Extracting..."
            tar -xf "$GYB_FILE"

            # Move to project directory
            /bin/rm -rf "$GYB_DIR"
            mv gyb "$GYB_DIR"

            # Make executable
            chmod +x "$GYB_DIR/gyb"

            if [[ -f "$GYB_DIR/gyb" ]]; then
                GYB_INSTALLED=true
                echo "✓ GYB installed to $GYB_DIR"
            fi
        else
            echo "Failed to download GYB from $GYB_URL"
        fi

        # Cleanup
        cd "$PROJECT_DIR"
        /bin/rm -rf "$TMP_DIR"
    fi
fi

# Verify GYB installation
echo ""
if command -v gyb &> /dev/null; then
    echo "✓ GYB installed: $(gyb --version 2>&1 | head -1)"
elif [[ -f "$GYB_DIR/gyb" ]]; then
    echo "✓ GYB installed locally: $($GYB_DIR/gyb --version 2>&1 | head -1)"
    echo ""
    echo "  Use with: --gyb-path $GYB_DIR/gyb"
    echo "  Or add to PATH: export PATH=\"$GYB_DIR:\$PATH\""
else
    echo "✗ GYB installation failed"
    echo "  Download manually from: https://github.com/GAM-team/got-your-back/releases"
fi

echo
echo "=========================================="
echo "Installation complete!"
echo
echo "Next steps:"
echo "  1. Create GYB project (if not done already):"
if [[ -f "$GYB_DIR/gyb" ]] && ! command -v gyb &> /dev/null; then
    echo "     $GYB_DIR/gyb --action create-project --email your.email@gmail.com"
    echo ""
    echo "  2. Authenticate with Gmail:"
    echo "     $GYB_DIR/gyb --email your.email@gmail.com --action count"
else
    echo "     gyb --action create-project --email your.email@gmail.com"
    echo ""
    echo "  2. Authenticate with Gmail:"
    echo "     gyb --email your.email@gmail.com --action count"
fi
echo
echo "  3. Test with a dry run:"
if [[ -f "$GYB_DIR/gyb" ]] && ! command -v gyb &> /dev/null; then
    echo "     python pst_to_gmail.py backup.pst --email EMAIL --gyb-path $GYB_DIR/gyb --dry-run"
else
    echo "     python pst_to_gmail.py backup.pst --email EMAIL --dry-run"
fi
echo
echo "  See SETUP_GYB.md for detailed instructions."
echo "=========================================="
