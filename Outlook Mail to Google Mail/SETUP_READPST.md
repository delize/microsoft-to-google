# Setting Up readpst (libpst)

`readpst` is part of the libpst library, which reads Microsoft Outlook PST files and converts them to standard formats like EML and MBOX.

## Installation

### macOS

Using Homebrew:

```bash
brew install libpst
```

Verify installation:

```bash
readpst --version
```

### Linux (Ubuntu/Debian)

```bash
sudo apt update
sudo apt install pst-utils
```

Verify installation:

```bash
readpst --version
```

### Linux (Fedora/RHEL/CentOS)

```bash
sudo dnf install libpst
```

Or on older systems:

```bash
sudo yum install libpst
```

### Linux (Arch)

```bash
sudo pacman -S libpst
```

### Windows

Windows has several options for running readpst, listed from recommended to fallback:

---

#### Option 1: WSL2 (Recommended)

WSL2 (Windows Subsystem for Linux) provides the best experience with full Linux compatibility.

**Step 1: Install WSL2**

1. Open PowerShell as Administrator
2. Run:
   ```powershell
   wsl --install
   ```
3. Restart your computer when prompted
4. After restart, Ubuntu will install automatically
5. Create a username and password when prompted

**Step 2: Install readpst in WSL2**

Open the Ubuntu terminal (search for "Ubuntu" in Start menu):

```bash
sudo apt update
sudo apt install pst-utils
```

Verify installation:

```bash
readpst --version
```

**Step 3: Access Windows Files from WSL2**

Your Windows drives are mounted under `/mnt/`:

```bash
# Access C: drive
cd /mnt/c/Users/YourUsername/Documents

# List PST files
ls *.pst
```

**Example: Convert a PST File in WSL2**

```bash
# Navigate to your PST file location
cd /mnt/c/Users/YourUsername/Documents

# Convert PST to EML
readpst -S -e -o ./converted_mail "Outlook Backup.pst"
```

---

#### Option 2: Cygwin

Cygwin provides a Unix-like environment on Windows with readpst available as a package.

**Step 1: Install Cygwin**

1. Download the Cygwin installer from [cygwin.com](https://www.cygwin.com/install.html)
2. Run `setup-x86_64.exe`
3. Choose "Install from Internet"
4. Select a mirror and proceed

**Step 2: Install readpst Package**

During Cygwin setup (or run setup again to add packages):

1. In the package selection screen, search for `readpst`
2. Click the "Skip" text next to `readpst` to select it for installation
3. Complete the installation

Alternatively, if you have Cygwin's `apt-cyg` package manager:
```bash
apt-cyg install readpst
```

**Step 3: Use readpst**

Open Cygwin Terminal and run:

```bash
readpst --version

# Convert PST (use /cygdrive/c for C: drive)
cd /cygdrive/c/Users/YourUsername/Documents
readpst -S -e -o ./converted_mail "Outlook Backup.pst"
```

---

#### Option 3: Native Windows Binary (ezwinports)

Pre-compiled Windows binaries are available from ezwinports (older version 0.6.63, but functional).

**Step 1: Download**

Download from: [libpst-0.6.63-w32-bin.zip](https://sourceforge.net/projects/ezwinports/files/libpst-0.6.63-w32-bin.zip/)

**Step 2: Extract**

1. Extract the ZIP file to a folder (e.g., `C:\Tools\libpst`)
2. The `bin` folder contains `readpst.exe`

**Step 3: Add to PATH (Optional)**

Add `C:\Tools\libpst\bin` to your system PATH, or use the full path:

```powershell
C:\Tools\libpst\bin\readpst.exe --version
```

**Step 4: Use readpst**

```powershell
cd C:\Users\YourUsername\Documents
C:\Tools\libpst\bin\readpst.exe -S -e -o .\converted_mail "Outlook Backup.pst"
```

**Note:** This is an older version (0.6.63 vs current 0.6.76) but handles most PST files correctly.

---

#### Option 4: Manual Conversion (Thunderbird)

If none of the above options work, you can manually convert PST files using Mozilla Thunderbird:

1. **Install Thunderbird** from [thunderbird.net](https://www.thunderbird.net/)
2. **Import PST:**
   - Go to `Tools` > `Import`
   - Select "Mail" and click Next
   - Select "Outlook" and click Next
   - Thunderbird will import the PST
3. **Install ImportExportTools NG:**
   - Go to `Tools` > `Add-ons and Themes`
   - Search for "ImportExportTools NG"
   - Install and restart Thunderbird
4. **Export as MBOX:**
   - Right-click the imported folder
   - Select `ImportExportTools NG` > `Export folder with subfolders`
   - Choose MBOX format
5. **Use the MBOX files** with this tool:
   ```bash
   python pst_to_gmail.py exported_mail.mbox --email user@gmail.com
   ```

## Usage

### Basic Conversion

Convert a PST file to EML format:

```bash
readpst -S -e -o ./output_folder "backup.pst"
```

**Options explained:**
- `-S` - Output each email as a separate file
- `-e` - Include attachments in the EML files
- `-o` - Specify output directory

### Preserve Folder Structure

```bash
readpst -S -e -r -o ./output_folder "backup.pst"
```

The `-r` flag creates subdirectories matching the PST folder structure.

### List PST Contents (No Conversion)

Just see what's in the PST without converting:

```bash
readpst -d ./debug.txt "backup.pst"
cat debug.txt
```

### Convert to MBOX Instead

If you prefer MBOX format:

```bash
readpst -o ./output_folder "backup.pst"
```

Without the `-S` flag, readpst creates MBOX files instead of individual EMLs.

## Understanding the Output

### EML Output (with -S flag)

```
output_folder/
├── Inbox/
│   ├── 00000001.eml
│   ├── 00000002.eml
│   └── ...
├── Sent Items/
│   ├── 00000001.eml
│   └── ...
└── Archive/
    └── ...
```

### MBOX Output (without -S flag)

```
output_folder/
├── Inbox.mbox
├── Sent Items.mbox
└── Archive.mbox
```

## Troubleshooting

### "readpst: command not found"

The package might have a different name:
- Try `pst-utils` (Ubuntu/Debian)
- Try `libpst` (Fedora/macOS)

### "Error reading PST file"

The PST file might be corrupted. Try:
1. Run `scanpst.exe` on Windows to repair
2. Try `pffexport` from libpff as an alternative

### "Permission denied"

Make sure you have read access to the PST file:
```bash
chmod +r "backup.pst"
```

### Large PST Files (50GB+)

For very large files:
1. Ensure adequate disk space (2-3x the PST size)
2. The conversion may take hours - let it run
3. Consider processing in a screen/tmux session:
   ```bash
   screen -S pst-convert
   readpst -S -e -o ./output "large_backup.pst"
   # Press Ctrl+A, then D to detach
   # Reconnect with: screen -r pst-convert
   ```

### Encoding Issues

If email subjects/bodies show garbled text:
```bash
readpst -S -e -8 -o ./output "backup.pst"
```

The `-8` flag forces UTF-8 encoding.

## Command Reference

| Option | Description |
|--------|-------------|
| `-S` | Separate each email into its own file (.eml) |
| `-e` | Include attachments in output |
| `-r` | Recursive - create subdirectories for folders |
| `-o DIR` | Output directory |
| `-8` | Force UTF-8 encoding |
| `-b` | Don't save attachments |
| `-c FMT` | Output contact format (vcard, list) |
| `-j N` | Number of parallel jobs |
| `-d FILE` | Debug output to file |
| `-V` | Show version |

## Performance Tips

### Use Multiple Cores

For faster conversion on multi-core systems:

```bash
readpst -S -e -j 4 -o ./output "backup.pst"
```

The `-j 4` uses 4 parallel jobs.

### SSD vs HDD

If possible, put both the PST file and output directory on an SSD. This dramatically speeds up conversion for large files.

### Memory Usage

readpst typically uses minimal memory. If you're running out of memory, try:
1. Close other applications
2. Process one PST file at a time
3. Split very large PST files using Outlook first

## Next Steps

Once you have converted EML or MBOX files:

1. **Validate the output:** `python eml_validator.py ./output_folder`
2. **Dry run upload:** `python pst_to_gmail.py ./output_folder --email EMAIL --dry-run`
3. **Upload to Gmail:** `python pst_to_gmail.py ./output_folder --email EMAIL`
