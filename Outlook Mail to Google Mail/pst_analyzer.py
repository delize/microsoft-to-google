#!/usr/bin/env python3
"""
PST/EML/MBOX Analyzer

Analyze email archives before migration to understand:
- Total message count
- Date range of messages
- Folder structure
- Size breakdown
- Sender/recipient statistics

Usage:
    python pst_analyzer.py backup.pst
    python pst_analyzer.py ./eml_folder/
    python pst_analyzer.py archive.mbox
"""

import os
import sys
import argparse
import json
import subprocess
import shutil
import tempfile
import email
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Optional, Tuple, Any
from email.utils import parsedate_to_datetime


class MailAnalyzer:
    """Analyze email archives"""

    def __init__(self):
        self.stats = {
            'total_messages': 0,
            'total_size_bytes': 0,
            'date_range': {'earliest': None, 'latest': None},
            'folders': defaultdict(int),
            'senders': defaultdict(int),
            'recipients': defaultdict(int),
            'has_attachments': 0,
            'by_year': defaultdict(int),
            'by_month': defaultdict(int),
        }

    def analyze_pst(self, pst_path: str) -> Dict[str, Any]:
        """Analyze a PST file using readpst"""
        pst_path = Path(pst_path)

        if not pst_path.exists():
            raise FileNotFoundError(f"PST file not found: {pst_path}")

        self.stats['total_size_bytes'] = pst_path.stat().st_size

        # Check if readpst is available
        readpst = shutil.which('readpst')
        if not readpst:
            print("Warning: readpst not found. Using file size estimation only.")
            print("Install readpst for detailed analysis:")
            print("  macOS:  brew install libpst")
            print("  Ubuntu: sudo apt install pst-utils")

            # Estimate based on file size (~50KB per message average)
            self.stats['total_messages'] = int(self.stats['total_size_bytes'] / 50000)
            return self.stats

        # Create temp directory for conversion
        with tempfile.TemporaryDirectory() as tmpdir:
            print(f"Analyzing PST file: {pst_path}")
            print(f"Size: {self.stats['total_size_bytes'] / 1024 / 1024:.1f} MB")
            print("This may take a while for large files...")

            # Run readpst to convert to EML for analysis
            cmd = [readpst, '-S', '-e', '-r', '-o', tmpdir, str(pst_path)]

            try:
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=3600  # 1 hour timeout
                )

                if result.returncode != 0:
                    print(f"Warning: readpst returned code {result.returncode}")
                    if result.stderr:
                        print(f"Error: {result.stderr[:500]}")

            except subprocess.TimeoutExpired:
                print("Warning: PST analysis timed out after 1 hour")
                return self.stats
            except Exception as e:
                print(f"Error running readpst: {e}")
                return self.stats

            # Analyze the converted EML files
            self._analyze_eml_directory(Path(tmpdir))

        return self.stats

    def analyze_mbox(self, mbox_path: str) -> Dict[str, Any]:
        """Analyze an MBOX file"""
        import mailbox

        mbox_path = Path(mbox_path)

        if not mbox_path.exists():
            raise FileNotFoundError(f"MBOX file not found: {mbox_path}")

        self.stats['total_size_bytes'] = mbox_path.stat().st_size

        print(f"Analyzing MBOX file: {mbox_path}")
        print(f"Size: {self.stats['total_size_bytes'] / 1024 / 1024:.1f} MB")

        try:
            mbox = mailbox.mbox(str(mbox_path))
            for message in mbox:
                self._analyze_message(message)
                self.stats['total_messages'] += 1

                if self.stats['total_messages'] % 1000 == 0:
                    print(f"  Analyzed {self.stats['total_messages']} messages...")

        except Exception as e:
            print(f"Error analyzing MBOX: {e}")

        return self.stats

    def analyze_eml_directory(self, eml_dir: str) -> Dict[str, Any]:
        """Analyze a directory of EML files"""
        eml_dir = Path(eml_dir)

        if not eml_dir.exists():
            raise FileNotFoundError(f"Directory not found: {eml_dir}")

        print(f"Analyzing EML directory: {eml_dir}")

        # Calculate total size
        for eml_file in eml_dir.rglob('*.eml'):
            self.stats['total_size_bytes'] += eml_file.stat().st_size

        print(f"Total size: {self.stats['total_size_bytes'] / 1024 / 1024:.1f} MB")

        self._analyze_eml_directory(eml_dir)

        return self.stats

    def _analyze_eml_directory(self, directory: Path):
        """Analyze all EML files in a directory"""
        eml_files = list(directory.rglob('*.eml'))

        for i, eml_file in enumerate(eml_files):
            try:
                # Get folder from path
                rel_path = eml_file.relative_to(directory)
                if len(rel_path.parts) > 1:
                    folder = '/'.join(rel_path.parts[:-1])
                else:
                    folder = 'Inbox'
                self.stats['folders'][folder] += 1

                # Parse the EML file
                with open(eml_file, 'rb') as f:
                    msg = email.message_from_binary_file(f)
                    self._analyze_message(msg)

                self.stats['total_messages'] += 1

                if (i + 1) % 1000 == 0:
                    print(f"  Analyzed {i + 1}/{len(eml_files)} messages...")

            except Exception as e:
                # Skip malformed messages
                self.stats['total_messages'] += 1
                continue

    def _analyze_message(self, msg):
        """Extract stats from a single email message"""
        # Date
        date_str = msg.get('Date')
        if date_str:
            try:
                msg_date = parsedate_to_datetime(date_str)

                if self.stats['date_range']['earliest'] is None:
                    self.stats['date_range']['earliest'] = msg_date
                    self.stats['date_range']['latest'] = msg_date
                else:
                    if msg_date < self.stats['date_range']['earliest']:
                        self.stats['date_range']['earliest'] = msg_date
                    if msg_date > self.stats['date_range']['latest']:
                        self.stats['date_range']['latest'] = msg_date

                self.stats['by_year'][msg_date.year] += 1
                self.stats['by_month'][f"{msg_date.year}-{msg_date.month:02d}"] += 1

            except Exception:
                pass

        # Sender
        sender = msg.get('From', '')
        if sender:
            # Extract email address
            if '<' in sender:
                sender = sender.split('<')[1].split('>')[0]
            sender = sender.lower().strip()
            if sender:
                self.stats['senders'][sender] += 1

        # Recipients
        for header in ['To', 'Cc']:
            recipients = msg.get(header, '')
            if recipients:
                for recip in recipients.split(','):
                    recip = recip.strip()
                    if '<' in recip:
                        recip = recip.split('<')[1].split('>')[0]
                    recip = recip.lower().strip()
                    if recip:
                        self.stats['recipients'][recip] += 1

        # Attachments
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    self.stats['has_attachments'] += 1
                    break

    def print_report(self):
        """Print analysis report"""
        print("\n" + "=" * 60)
        print("EMAIL ARCHIVE ANALYSIS")
        print("=" * 60)

        print(f"\nTotal messages: {self.stats['total_messages']:,}")
        print(f"Total size: {self.stats['total_size_bytes'] / 1024 / 1024:.1f} MB")

        if self.stats['date_range']['earliest']:
            print(f"\nDate range:")
            print(f"  Earliest: {self.stats['date_range']['earliest'].strftime('%Y-%m-%d')}")
            print(f"  Latest:   {self.stats['date_range']['latest'].strftime('%Y-%m-%d')}")

        if self.stats['folders']:
            print(f"\nFolders ({len(self.stats['folders'])}):")
            sorted_folders = sorted(self.stats['folders'].items(), key=lambda x: -x[1])
            for folder, count in sorted_folders[:10]:
                print(f"  {folder}: {count:,}")
            if len(sorted_folders) > 10:
                print(f"  ... and {len(sorted_folders) - 10} more folders")

        if self.stats['by_year']:
            print(f"\nMessages by year:")
            for year in sorted(self.stats['by_year'].keys()):
                count = self.stats['by_year'][year]
                bar = '#' * min(50, int(count / max(self.stats['by_year'].values()) * 50))
                print(f"  {year}: {count:>6,} {bar}")

        if self.stats['senders']:
            print(f"\nTop senders:")
            sorted_senders = sorted(self.stats['senders'].items(), key=lambda x: -x[1])
            for sender, count in sorted_senders[:5]:
                print(f"  {sender}: {count:,}")

        if self.stats['has_attachments']:
            print(f"\nMessages with attachments: {self.stats['has_attachments']:,}")

        print("\n" + "=" * 60)

        # Estimate migration time (rough: ~5 messages/second for upload)
        est_seconds = self.stats['total_messages'] / 5
        if est_seconds < 60:
            est_time = f"{int(est_seconds)} seconds"
        elif est_seconds < 3600:
            est_time = f"{int(est_seconds / 60)} minutes"
        else:
            est_time = f"{est_seconds / 3600:.1f} hours"

        print(f"Estimated upload time: ~{est_time}")
        print("=" * 60)

    def to_json(self) -> str:
        """Export stats as JSON"""
        output = {
            'total_messages': self.stats['total_messages'],
            'total_size_bytes': self.stats['total_size_bytes'],
            'total_size_mb': round(self.stats['total_size_bytes'] / 1024 / 1024, 1),
            'date_range': {
                'earliest': self.stats['date_range']['earliest'].isoformat() if self.stats['date_range']['earliest'] else None,
                'latest': self.stats['date_range']['latest'].isoformat() if self.stats['date_range']['latest'] else None,
            },
            'folders': dict(self.stats['folders']),
            'messages_by_year': dict(self.stats['by_year']),
            'top_senders': dict(sorted(self.stats['senders'].items(), key=lambda x: -x[1])[:20]),
            'messages_with_attachments': self.stats['has_attachments'],
        }
        return json.dumps(output, indent=2)


def main():
    parser = argparse.ArgumentParser(
        description='Analyze PST/EML/MBOX files before migration',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s backup.pst
  %(prog)s ./eml_folder/
  %(prog)s archive.mbox
  %(prog)s backup.pst --json > analysis.json
        """
    )

    parser.add_argument('input_path',
                        help='PST file, EML directory, or MBOX file')
    parser.add_argument('--json', action='store_true',
                        help='Output analysis as JSON')
    parser.add_argument('--folders-only', action='store_true',
                        help='Only show folder structure')

    args = parser.parse_args()

    path = Path(args.input_path)

    if not path.exists():
        print(f"Error: Path not found: {path}")
        sys.exit(1)

    analyzer = MailAnalyzer()

    # Detect format and analyze
    if path.is_file():
        ext = path.suffix.lower()
        if ext == '.pst':
            analyzer.analyze_pst(str(path))
        elif ext == '.mbox':
            analyzer.analyze_mbox(str(path))
        elif ext == '.eml':
            # Single EML file
            analyzer.analyze_eml_directory(str(path.parent))
        else:
            # Try to detect by content
            with open(path, 'rb') as f:
                header = f.read(100)
                if b'!BDN' in header:
                    analyzer.analyze_pst(str(path))
                elif b'From ' in header[:5]:
                    analyzer.analyze_mbox(str(path))
                else:
                    print(f"Error: Unknown file format: {path}")
                    sys.exit(1)
    else:
        # Directory
        eml_files = list(path.rglob('*.eml'))
        if eml_files:
            analyzer.analyze_eml_directory(str(path))
        else:
            print(f"Error: No EML files found in {path}")
            sys.exit(1)

    # Output
    if args.json:
        print(analyzer.to_json())
    elif args.folders_only:
        print("\nFolder structure:")
        for folder, count in sorted(analyzer.stats['folders'].items()):
            print(f"  {folder}: {count:,}")
    else:
        analyzer.print_report()


if __name__ == '__main__':
    main()
