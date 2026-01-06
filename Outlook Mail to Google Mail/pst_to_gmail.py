#!/usr/bin/env python3
"""
PST to Gmail Migration Tool

This script orchestrates the migration of email from Outlook PST files to Gmail
using readpst for PST conversion and GYB (Got Your Back) for Gmail upload.

Supported input formats:
    - PST files (.pst) - Converted to EML via readpst
    - EML files (.eml) - Uploaded directly via GYB
    - MBOX files (.mbox) - Uploaded directly via GYB

Requirements:
    - readpst (libpst) - For PST conversion
    - GYB (Got Your Back) - For Gmail upload
    - Python 3.7+

Usage:
    python pst_to_gmail.py backup.pst --email user@gmail.com
    python pst_to_gmail.py ./eml_folder/ --email user@gmail.com
    python pst_to_gmail.py archive.mbox --email user@gmail.com
"""

import os
import sys
import argparse
import subprocess
import shutil
import time
import re
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple, List


class MigrationStats:
    """Track migration statistics"""
    def __init__(self):
        self.pst_size = 0
        self.messages_found = 0
        self.messages_uploaded = 0
        self.messages_skipped = 0
        self.messages_failed = 0
        self.start_time = None
        self.end_time = None

    def start(self):
        self.start_time = datetime.now()

    def finish(self):
        self.end_time = datetime.now()

    @property
    def duration(self) -> str:
        if not self.start_time or not self.end_time:
            return "N/A"
        delta = self.end_time - self.start_time
        hours, remainder = divmod(int(delta.total_seconds()), 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours > 0:
            return f"{hours}h {minutes}m {seconds}s"
        elif minutes > 0:
            return f"{minutes}m {seconds}s"
        else:
            return f"{seconds}s"


def find_executable(name: str, custom_path: Optional[str] = None) -> Optional[str]:
    """Find an executable in PATH or custom location"""
    if custom_path:
        if os.path.isfile(custom_path) and os.access(custom_path, os.X_OK):
            return custom_path
        else:
            return None

    # Try common locations
    result = shutil.which(name)
    if result:
        return result

    # Try Python module for GYB
    if name == 'gyb':
        # GYB might be installed as a Python module
        try:
            result = subprocess.run(
                [sys.executable, '-m', 'gyb', '--version'],
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                return f"{sys.executable} -m gyb"
        except Exception:
            pass

    return None


def detect_input_format(path: str) -> Tuple[str, str]:
    """
    Detect input format based on file extension or directory contents.

    Returns:
        Tuple of (format, description) where format is 'pst', 'eml', 'mbox', or 'eml_dir'
    """
    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"Input path does not exist: {path}")

    if path.is_file():
        ext = path.suffix.lower()
        if ext == '.pst':
            return 'pst', f"PST file ({path.stat().st_size / 1024 / 1024:.1f} MB)"
        elif ext == '.eml':
            return 'eml', "Single EML file"
        elif ext == '.mbox':
            return 'mbox', f"MBOX file ({path.stat().st_size / 1024 / 1024:.1f} MB)"
        else:
            # Try to detect by content
            with open(path, 'rb') as f:
                header = f.read(100)
                if b'!BDN' in header:  # PST magic bytes
                    return 'pst', f"PST file ({path.stat().st_size / 1024 / 1024:.1f} MB)"
                elif b'From ' in header[:5]:  # MBOX format
                    return 'mbox', f"MBOX file ({path.stat().st_size / 1024 / 1024:.1f} MB)"
            raise ValueError(f"Unknown file format: {path}")

    elif path.is_dir():
        # Check if directory contains EML files
        eml_files = list(path.rglob('*.eml'))
        if eml_files:
            return 'eml_dir', f"Directory with {len(eml_files)} EML files"

        mbox_files = list(path.rglob('*.mbox'))
        if mbox_files:
            return 'mbox_dir', f"Directory with {len(mbox_files)} MBOX files"

        raise ValueError(f"Directory contains no EML or MBOX files: {path}")

    raise ValueError(f"Invalid input path: {path}")


def count_eml_files(directory: Path) -> int:
    """Count EML files in a directory recursively"""
    return len(list(directory.rglob('*.eml')))


def count_mbox_files(directory: Path) -> int:
    """Count MBOX files in a directory recursively"""
    return len(list(directory.rglob('*.mbox')))


def count_messages_in_mbox(mbox_path: Path) -> int:
    """Count messages in an MBOX file by counting 'From ' lines"""
    count = 0
    try:
        with open(mbox_path, 'rb') as f:
            for line in f:
                if line.startswith(b'From '):
                    count += 1
    except Exception:
        pass
    return count


def count_all_mbox_messages(directory: Path) -> int:
    """Count total messages across all MBOX files in a directory"""
    total = 0
    for mbox_file in directory.rglob('*.mbox'):
        total += count_messages_in_mbox(mbox_file)
    return total


def convert_pst_to_eml(pst_path: str, output_dir: str, readpst_path: str,
                        dry_run: bool = False) -> Tuple[bool, int]:
    """
    Convert PST file to EML files using readpst.

    Returns:
        Tuple of (success, message_count)
    """
    pst_path = Path(pst_path)
    output_dir = Path(output_dir)

    print(f"\n{'=' * 60}")
    print("PST CONVERSION")
    print('=' * 60)
    print(f"Input:  {pst_path}")
    print(f"Output: {output_dir}")
    print(f"Size:   {pst_path.stat().st_size / 1024 / 1024:.1f} MB")

    if dry_run:
        print("\n[DRY RUN] Would convert PST to EML")
        # Estimate message count (rough heuristic: ~50KB per message average)
        estimated = int(pst_path.stat().st_size / 50000)
        print(f"[DRY RUN] Estimated messages: ~{estimated}")
        return True, estimated

    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)

    # Build readpst command
    # Default output is MBOX format, which GYB's restore-mbox can import
    # -o: Output directory
    cmd = [readpst_path, '-o', str(output_dir), str(pst_path)]

    print(f"\nRunning: {' '.join(cmd)}")
    print("-" * 60)

    start_time = time.time()

    try:
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1
        )

        # Stream output
        for line in process.stdout:
            line = line.strip()
            if line:
                print(f"  {line}")

        process.wait()

        if process.returncode != 0:
            print(f"\nError: readpst exited with code {process.returncode}")
            return False, 0

    except Exception as e:
        print(f"\nError running readpst: {e}")
        return False, 0

    elapsed = time.time() - start_time

    # Count converted files (MBOX format)
    message_count = count_mbox_files(output_dir)

    print("-" * 60)
    print(f"Conversion complete in {elapsed:.1f}s")
    print(f"MBOX files created: {message_count}")

    return True, message_count


def run_gyb_upload(email: str, local_folder: str, gyb_path: str,
                   action: str = 'restore', label: Optional[str] = None,
                   dry_run: bool = False) -> Tuple[bool, int, int]:
    """
    Run GYB to upload emails to Gmail.

    Args:
        email: Target Gmail address
        local_folder: Directory containing EML/MBOX files
        gyb_path: Path to GYB executable
        action: GYB action ('restore' for EML, 'restore-mbox' for MBOX)
        label: Optional Gmail label to apply
        dry_run: If True, just count files

    Returns:
        Tuple of (success, uploaded_count, failed_count)
    """
    print(f"\n{'=' * 60}")
    print("GMAIL UPLOAD")
    print('=' * 60)
    print(f"Email:  {email}")
    print(f"Source: {local_folder}")
    print(f"Action: {action}")
    if label:
        print(f"Label:  {label}")

    if dry_run:
        # Count what would be uploaded
        folder = Path(local_folder)
        if action == 'restore':
            count = count_eml_files(folder)
        else:
            count = len(list(folder.rglob('*.mbox')))
        print(f"\n[DRY RUN] Would upload {count} messages to {email}")
        return True, count, 0

    # Build GYB command
    if ' -m ' in gyb_path:
        # Running as Python module
        parts = gyb_path.split()
        cmd = parts + ['--email', email, '--action', action, '--local-folder', local_folder]
    else:
        cmd = [gyb_path, '--email', email, '--action', action, '--local-folder', local_folder]

    if label:
        cmd.extend(['--label-restored', label])

    print(f"\nRunning: {' '.join(cmd)}")
    print("-" * 60)

    start_time = time.time()
    uploaded = 0
    failed = 0

    try:
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1
        )

        # Stream and parse output
        for line in process.stdout:
            line = line.strip()
            if line:
                print(f"  {line}")

                # Parse GYB output for stats
                if 'restored' in line.lower() or 'uploaded' in line.lower():
                    # Try to extract numbers
                    match = re.search(r'(\d+)\s*(message|email|restored)', line.lower())
                    if match:
                        uploaded = int(match.group(1))
                elif 'error' in line.lower() or 'failed' in line.lower():
                    match = re.search(r'(\d+)\s*(error|failed)', line.lower())
                    if match:
                        failed = int(match.group(1))

        process.wait()

        if process.returncode != 0:
            print(f"\nWarning: GYB exited with code {process.returncode}")
            # Don't return False - some messages may have been uploaded

    except Exception as e:
        print(f"\nError running GYB: {e}")
        return False, uploaded, failed

    elapsed = time.time() - start_time

    print("-" * 60)
    print(f"Upload complete in {elapsed:.1f}s")

    return True, uploaded, failed


def check_gyb_auth(email: str, gyb_path: str) -> bool:
    """Check if GYB is authenticated for the given email"""
    print(f"\nChecking GYB authentication for {email}...")

    if ' -m ' in gyb_path:
        parts = gyb_path.split()
        cmd = parts + ['--email', email, '--action', 'count']
    else:
        cmd = [gyb_path, '--email', email, '--action', 'count']

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60
        )

        if result.returncode == 0:
            print("  GYB is authenticated")
            return True
        else:
            print(f"  GYB authentication check failed")
            print(f"  First create a project: gyb --action create-project --email {email}")
            print(f"  Then authenticate: gyb --email {email} --action count")
            return False

    except subprocess.TimeoutExpired:
        print("  GYB authentication check timed out")
        return False
    except Exception as e:
        print(f"  Error checking GYB auth: {e}")
        return False


def print_summary(stats: MigrationStats, dry_run: bool = False):
    """Print migration summary"""
    print(f"\n{'=' * 60}")
    if dry_run:
        print("DRY RUN SUMMARY")
    else:
        print("MIGRATION SUMMARY")
    print('=' * 60)

    if stats.pst_size > 0:
        print(f"PST file size:     {stats.pst_size / 1024 / 1024:.1f} MB")
    print(f"Messages found:    {stats.messages_found}")

    if dry_run:
        print(f"Would upload:      {stats.messages_found}")
    else:
        print(f"Successfully uploaded: {stats.messages_uploaded}")
        if stats.messages_skipped > 0:
            print(f"Skipped (duplicates):  {stats.messages_skipped}")
        if stats.messages_failed > 0:
            print(f"Failed:                {stats.messages_failed}")
        print(f"Duration:          {stats.duration}")

    print('=' * 60)


def main():
    parser = argparse.ArgumentParser(
        description='Migrate email from PST/EML/MBOX to Gmail',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s backup.pst --email user@gmail.com
  %(prog)s backup.pst --email user@gmail.com --dry-run
  %(prog)s ./eml_folder/ --email user@gmail.com
  %(prog)s archive.mbox --email user@gmail.com
  %(prog)s backup.pst --email user@gmail.com --label "Imported/Outlook"
        """
    )

    parser.add_argument('input_path',
                        help='PST file, EML file/directory, or MBOX file')
    parser.add_argument('--email', '-e', required=True,
                        help='Target Gmail address')
    parser.add_argument('--dry-run', action='store_true',
                        help='Preview without uploading')
    parser.add_argument('--output-dir', default='./converted_mail',
                        help='Directory for converted EML files (default: ./converted_mail)')
    parser.add_argument('--keep-converted', action='store_true',
                        help="Don't delete converted EML files after upload")
    parser.add_argument('--label',
                        help='Gmail label to apply to imported messages')
    parser.add_argument('--resume', action='store_true',
                        help='Resume interrupted upload (skip conversion if output exists)')
    parser.add_argument('--gyb-path',
                        help='Path to GYB executable (default: auto-detect)')
    parser.add_argument('--readpst-path',
                        help='Path to readpst executable (default: auto-detect)')

    args = parser.parse_args()

    stats = MigrationStats()
    stats.start()

    print("=" * 60)
    print("PST TO GMAIL MIGRATION TOOL")
    print("=" * 60)

    # Detect input format
    try:
        input_format, format_desc = detect_input_format(args.input_path)
        print(f"\nInput detected: {format_desc}")
    except (FileNotFoundError, ValueError) as e:
        print(f"\nError: {e}")
        sys.exit(1)

    # Find required tools
    readpst_path = None
    if input_format == 'pst':
        readpst_path = find_executable('readpst', args.readpst_path)
        if not readpst_path:
            print("\nError: readpst not found!")
            print("Install it with:")
            print("  macOS:  brew install libpst")
            print("  Ubuntu: sudo apt install pst-utils")
            print("  See SETUP_READPST.md for more options")
            sys.exit(1)
        print(f"readpst: {readpst_path}")

    gyb_path = find_executable('gyb', args.gyb_path)
    if not gyb_path:
        print("\nError: GYB (Got Your Back) not found!")
        print("Install it with: pip install got-your-back")
        print("See SETUP_GYB.md for setup instructions")
        sys.exit(1)
    print(f"GYB: {gyb_path}")

    # Check GYB authentication
    if not args.dry_run:
        if not check_gyb_auth(args.email, gyb_path):
            print("\nPlease authenticate GYB first:")
            print(f"  gyb --email {args.email} --action count")
            sys.exit(1)

    # Determine what to upload
    upload_path = args.input_path
    gyb_action = 'restore'
    needs_conversion = False

    if input_format == 'pst':
        needs_conversion = True
        upload_path = args.output_dir
        gyb_action = 'restore-mbox'  # PST converts to MBOX format
        stats.pst_size = Path(args.input_path).stat().st_size

        # Check if we should skip conversion (resume mode)
        if args.resume and Path(args.output_dir).exists():
            mbox_count = count_mbox_files(Path(args.output_dir))
            if mbox_count > 0:
                print(f"\nResume mode: Found {mbox_count} existing MBOX files, skipping conversion")
                needs_conversion = False
                stats.messages_found = mbox_count
        elif Path(args.output_dir).exists() and not args.resume:
            # Clean up old files when not resuming
            print(f"\nCleaning up previous conversion in {args.output_dir}...")
            shutil.rmtree(args.output_dir)

        if needs_conversion:
            success, count = convert_pst_to_eml(
                args.input_path,
                args.output_dir,
                readpst_path,
                dry_run=args.dry_run
            )
            if not success and not args.dry_run:
                print("\nPST conversion failed!")
                sys.exit(1)
            stats.messages_found = count

    elif input_format == 'mbox':
        gyb_action = 'restore-mbox'
        # Rough estimate for MBOX
        file_size = Path(args.input_path).stat().st_size
        stats.messages_found = int(file_size / 50000)  # Rough estimate

    elif input_format == 'eml':
        stats.messages_found = 1

    elif input_format == 'eml_dir':
        stats.messages_found = count_eml_files(Path(args.input_path))

    elif input_format == 'mbox_dir':
        gyb_action = 'restore-mbox'
        mbox_files = list(Path(args.input_path).rglob('*.mbox'))
        stats.messages_found = len(mbox_files)

    # Upload to Gmail
    success, uploaded, failed = run_gyb_upload(
        args.email,
        upload_path,
        gyb_path,
        action=gyb_action,
        label=args.label,
        dry_run=args.dry_run
    )

    stats.messages_uploaded = uploaded
    stats.messages_failed = failed

    # Cleanup converted files
    if input_format == 'pst' and not args.keep_converted and not args.dry_run:
        if Path(args.output_dir).exists():
            print(f"\nCleaning up converted files in {args.output_dir}")
            shutil.rmtree(args.output_dir)

    stats.finish()
    print_summary(stats, dry_run=args.dry_run)

    if args.dry_run:
        print("\n*** DRY RUN COMPLETE - No emails were uploaded ***")
        print(f"\nTo perform the actual migration, run without --dry-run:")
        print(f"  python {sys.argv[0]} {args.input_path} --email {args.email}")
    else:
        print("\nMigration complete!")

    # Exit with appropriate code
    if stats.messages_failed > 0 and stats.messages_uploaded == 0:
        sys.exit(1)
    sys.exit(0)


if __name__ == '__main__':
    main()
