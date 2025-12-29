# Outlook Mail to Google Mail Migration

Migrate email from Outlook PST files, EML files, or MBOX archives to Gmail using an automated pipeline.

## The Problem

Moving email from Outlook to Gmail is surprisingly difficult:

- **GWMMO** (Google Workspace Migration for Microsoft Outlook) only works with paid Workspace accounts, not personal @gmail.com
- **Gmail's web import** only works for cloud-to-cloud (Yahoo, Outlook.com), not local PST files
- **Manual IMAP drag-and-drop** is slow, unreliable for large mailboxes, and has daily limits
- **PST files** are Microsoft's proprietary format that Gmail can't read directly

## The Solution

This tool automates the migration by combining two battle-tested tools:

1. **readpst** (libpst) - Converts PST files to standard EML format
2. **GYB** (Got Your Back) - Uploads EML/MBOX to Gmail via the Gmail API

The orchestrator handles format detection, conversion, progress tracking, and error recovery.

## Supported Input Formats

| Format | Extension | Conversion Required |
|--------|-----------|---------------------|
| Outlook PST | `.pst` | Yes (via readpst) |
| Email Message | `.eml` | No |
| Mailbox Archive | `.mbox` | No |

## Quick Start

### 1. Install Dependencies

**macOS:**
```bash
brew install libpst
pip install got-your-back
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt install pst-utils
pip install got-your-back
```

**Windows:** See [SETUP_READPST.md](SETUP_READPST.md) for WSL2 setup.

### 2. Set Up GYB Authentication

Follow the detailed guide: **[SETUP_GYB.md](SETUP_GYB.md)**

Quick version:
```bash
# Authenticate GYB with your Gmail account
gyb --email your.email@gmail.com --action check
```

### 3. Analyze Your PST (Optional)

```bash
python pst_analyzer.py "Outlook Backup.pst"
```

This shows message counts, date ranges, folder structure, and size estimates.

### 4. Test with Dry Run

```bash
python pst_to_gmail.py "Outlook Backup.pst" --email your.email@gmail.com --dry-run
```

### 5. Run the Migration

```bash
python pst_to_gmail.py "Outlook Backup.pst" --email your.email@gmail.com
```

## Usage Examples

### Basic PST Migration

```bash
python pst_to_gmail.py backup.pst --email user@gmail.com
```

### Import with a Gmail Label

Organize imported mail under a specific label:

```bash
python pst_to_gmail.py backup.pst --email user@gmail.com --label "Imported/Outlook"
```

### Pre-converted EML Files

If you've already converted your PST to EML:

```bash
python pst_to_gmail.py ./eml_folder/ --email user@gmail.com
```

### MBOX Archive

Import from Thunderbird or other mail clients:

```bash
python pst_to_gmail.py archive.mbox --email user@gmail.com
```

### Resume Interrupted Upload

If the upload was interrupted, resume where you left off:

```bash
python pst_to_gmail.py backup.pst --email user@gmail.com --resume
```

### Keep Converted Files

Don't delete the EML files after upload (useful for backup):

```bash
python pst_to_gmail.py backup.pst --email user@gmail.com --keep-converted
```

### Custom Output Directory

Store converted EML files in a specific location:

```bash
python pst_to_gmail.py backup.pst --email user@gmail.com --output-dir /path/to/output
```

## Command Line Options

### pst_to_gmail.py

| Option | Description |
|--------|-------------|
| `--email EMAIL` | Target Gmail address (required) |
| `--dry-run` | Preview without uploading |
| `--label LABEL` | Gmail label to apply to imported messages |
| `--output-dir DIR` | Where to store converted EML files (default: ./converted_mail) |
| `--keep-converted` | Don't delete converted EML files after upload |
| `--resume` | Resume interrupted upload |
| `--batch-size N` | Messages per GYB batch (default: 100) |
| `--gyb-path PATH` | Path to GYB executable (default: auto-detect) |
| `--readpst-path PATH` | Path to readpst executable (default: auto-detect) |

### pst_analyzer.py

| Option | Description |
|--------|-------------|
| `--json` | Output analysis as JSON |
| `--folders-only` | Only show folder structure |

### eml_validator.py

| Option | Description |
|--------|-------------|
| `--fix` | Attempt to fix common issues |
| `--sample N` | Number of random files to validate (default: all) |

## How It Works

```
┌─────────────────────────────────────────────────────────────┐
│                        PST File                              │
│                    (Outlook Backup.pst)                      │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                    readpst (libpst)                          │
│              Converts PST → EML files                        │
│         Preserves folder structure & attachments             │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                   EML Files Directory                        │
│                ./converted_mail/Inbox/                       │
│                ./converted_mail/Sent/                        │
│                ./converted_mail/Archive/                     │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                  GYB (Got Your Back)                         │
│              Uploads EML → Gmail via API                     │
│         Handles rate limits & resume capability              │
└─────────────────────┬───────────────────────────────────────┘
                      │
                      ▼
┌─────────────────────────────────────────────────────────────┐
│                      Gmail Inbox                             │
│              All messages imported with labels               │
└─────────────────────────────────────────────────────────────┘
```

## Features

### Robust Migration
- **Resume capability** - Continue where you left off if interrupted
- **Progress tracking** - Real-time status updates
- **Error handling** - Graceful handling of malformed messages
- **Duplicate detection** - Avoids re-importing existing messages

### Format Support
- **PST files** - Microsoft Outlook data files
- **EML files** - Standard email format (single or directory)
- **MBOX archives** - Unix mailbox format (Thunderbird, etc.)

### Cross-Platform
- **macOS** - Native support via Homebrew
- **Linux** - Native support via apt/yum
- **Windows** - WSL2 support with detailed setup guide

### Safety
- **Dry-run mode** - Preview what will be imported
- **Keep converted files** - Optional backup of EML files
- **Validation tools** - Check files before import

## File Structure

```
Outlook Mail to Google Mail/
├── README.md                    # This file
├── SETUP_GYB.md                 # GYB installation & OAuth setup
├── SETUP_READPST.md             # readpst installation guide
├── pst_to_gmail.py              # Main orchestrator tool
├── pst_analyzer.py              # Pre-migration analysis
├── eml_validator.py             # EML validation tool
├── requirements.txt             # Python dependencies
└── scripts/
    ├── install_deps.sh          # macOS/Linux installer
    ├── install_deps.ps1         # Windows PowerShell (limited)
    └── install_deps_wsl.sh      # WSL2 setup script
```

## Troubleshooting

### "readpst: command not found"

Install libpst:
- macOS: `brew install libpst`
- Ubuntu/Debian: `sudo apt install pst-utils`
- Windows: Use WSL2 (see [SETUP_READPST.md](SETUP_READPST.md))

### "GYB not authenticated"

Run GYB authentication:
```bash
gyb --email your.email@gmail.com --action check
```

### "Rate limit exceeded"

GYB handles rate limiting automatically. If you're hitting limits frequently:
- Reduce batch size: `--batch-size 50`
- Wait and resume: `--resume`

### "PST file is corrupted"

Try running `scanpst.exe` (Windows) to repair the PST first, or use a tool like `pffexport` as an alternative to readpst.

### Large PST files (50GB+)

For very large files:
1. Ensure adequate disk space (2x PST size for conversion)
2. Consider splitting by date range
3. Use `--keep-converted` to avoid re-conversion if upload fails

## Comparison with Other Methods

| Method | Personal Gmail | Large Files | Resume | Progress |
|--------|----------------|-------------|--------|----------|
| **This Tool** | Yes | Yes | Yes | Yes |
| GWMMO | No (Workspace only) | Yes | Yes | Yes |
| Outlook IMAP | Yes | Slow/Unreliable | No | No |
| Gmail Web Import | No (cloud only) | No | No | No |

## Security

- GYB credentials are stored locally in `~/.gyb/`
- PST files may contain sensitive data - handle securely
- Converted EML files are deleted by default after upload
- No data is sent to third parties - direct Gmail API only

## Contributing

Found an edge case that isn't handled? Have an improvement? Contributions are welcome!

1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - See [LICENSE](../LICENSE) for details.
