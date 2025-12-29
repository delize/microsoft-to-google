#!/usr/bin/env python3
"""
EML/MBOX Validator

Validate email files before uploading to Gmail to catch issues early.

Checks:
- Malformed headers
- Missing required fields (From, Date)
- Encoding issues
- Duplicate detection
- Oversized messages

Usage:
    python eml_validator.py ./eml_folder/
    python eml_validator.py archive.mbox
    python eml_validator.py ./eml_folder/ --sample 100
"""

import os
import sys
import argparse
import email
import hashlib
import random
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Set, Tuple, Any
from email.utils import parsedate_to_datetime


class ValidationResult:
    """Store validation results for a single message"""
    def __init__(self, path: str):
        self.path = path
        self.warnings: List[str] = []
        self.errors: List[str] = []

    def add_warning(self, msg: str):
        self.warnings.append(msg)

    def add_error(self, msg: str):
        self.errors.append(msg)

    @property
    def is_valid(self) -> bool:
        return len(self.errors) == 0

    @property
    def has_issues(self) -> bool:
        return len(self.errors) > 0 or len(self.warnings) > 0


class EMLValidator:
    """Validate EML and MBOX files"""

    # Gmail limits
    MAX_MESSAGE_SIZE = 25 * 1024 * 1024  # 25 MB
    MAX_RECIPIENTS = 500

    def __init__(self):
        self.results: List[ValidationResult] = []
        self.seen_ids: Set[str] = set()
        self.duplicates: List[Tuple[str, str]] = []  # (path, message_id)
        self.stats = {
            'total': 0,
            'valid': 0,
            'warnings': 0,
            'errors': 0,
            'duplicates': 0,
        }

    def validate_directory(self, directory: str, sample: int = 0) -> Dict[str, Any]:
        """Validate all EML files in a directory"""
        directory = Path(directory)

        if not directory.exists():
            raise FileNotFoundError(f"Directory not found: {directory}")

        eml_files = list(directory.rglob('*.eml'))

        if not eml_files:
            print(f"No EML files found in {directory}")
            return self.stats

        # Sample if requested
        if sample > 0 and len(eml_files) > sample:
            print(f"Sampling {sample} of {len(eml_files)} files...")
            eml_files = random.sample(eml_files, sample)
        else:
            print(f"Validating {len(eml_files)} EML files...")

        for i, eml_file in enumerate(eml_files):
            self._validate_eml_file(eml_file)

            if (i + 1) % 1000 == 0:
                print(f"  Validated {i + 1}/{len(eml_files)}...")

        return self.stats

    def validate_mbox(self, mbox_path: str, sample: int = 0) -> Dict[str, Any]:
        """Validate all messages in an MBOX file"""
        import mailbox

        mbox_path = Path(mbox_path)

        if not mbox_path.exists():
            raise FileNotFoundError(f"MBOX file not found: {mbox_path}")

        print(f"Validating MBOX file: {mbox_path}")

        try:
            mbox = mailbox.mbox(str(mbox_path))
            messages = list(mbox)

            # Sample if requested
            if sample > 0 and len(messages) > sample:
                print(f"Sampling {sample} of {len(messages)} messages...")
                messages = random.sample(messages, sample)
            else:
                print(f"Validating {len(messages)} messages...")

            for i, message in enumerate(messages):
                result = self._validate_message(message, f"mbox:{i}")
                self.results.append(result)
                self._update_stats(result)

                if (i + 1) % 1000 == 0:
                    print(f"  Validated {i + 1}/{len(messages)}...")

        except Exception as e:
            print(f"Error reading MBOX: {e}")

        return self.stats

    def _validate_eml_file(self, eml_path: Path):
        """Validate a single EML file"""
        result = ValidationResult(str(eml_path))

        try:
            # Check file size
            file_size = eml_path.stat().st_size
            if file_size > self.MAX_MESSAGE_SIZE:
                result.add_error(f"Message too large: {file_size / 1024 / 1024:.1f} MB (max 25 MB)")

            if file_size == 0:
                result.add_error("Empty file")
                self.results.append(result)
                self._update_stats(result)
                return

            # Parse the message
            with open(eml_path, 'rb') as f:
                msg = email.message_from_binary_file(f)
                self._validate_message_content(msg, result)

        except Exception as e:
            result.add_error(f"Parse error: {e}")

        self.results.append(result)
        self._update_stats(result)

    def _validate_message(self, msg, path: str) -> ValidationResult:
        """Validate a message object"""
        result = ValidationResult(path)
        self._validate_message_content(msg, result)
        return result

    def _validate_message_content(self, msg, result: ValidationResult):
        """Validate message content"""

        # Check for required headers
        if not msg.get('From'):
            result.add_error("Missing From header")
        else:
            from_addr = msg.get('From', '')
            if '@' not in from_addr and '<' not in from_addr:
                result.add_warning(f"Malformed From address: {from_addr[:50]}")

        if not msg.get('Date'):
            result.add_warning("Missing Date header")
        else:
            date_str = msg.get('Date')
            try:
                parsedate_to_datetime(date_str)
            except Exception:
                result.add_warning(f"Invalid Date format: {date_str[:50]}")

        # Check Message-ID for duplicates
        msg_id = msg.get('Message-ID', '')
        if msg_id:
            if msg_id in self.seen_ids:
                result.add_warning("Duplicate Message-ID detected")
                self.duplicates.append((result.path, msg_id))
                self.stats['duplicates'] += 1
            else:
                self.seen_ids.add(msg_id)
        else:
            # Generate a pseudo-ID from content for duplicate detection
            pseudo_id = self._generate_pseudo_id(msg)
            if pseudo_id in self.seen_ids:
                result.add_warning("Likely duplicate (same content hash)")
                self.duplicates.append((result.path, pseudo_id))
                self.stats['duplicates'] += 1
            else:
                self.seen_ids.add(pseudo_id)

        # Check recipient count
        all_recipients = []
        for header in ['To', 'Cc', 'Bcc']:
            recips = msg.get(header, '')
            if recips:
                all_recipients.extend(recips.split(','))

        if len(all_recipients) > self.MAX_RECIPIENTS:
            result.add_warning(f"Many recipients: {len(all_recipients)} (may be slow to process)")

        # Check for encoding issues
        subject = msg.get('Subject', '')
        if subject:
            try:
                # Try to decode subject
                if '=?' in subject:
                    from email.header import decode_header
                    decoded = decode_header(subject)
                    for part, charset in decoded:
                        if isinstance(part, bytes) and charset:
                            part.decode(charset)
            except Exception as e:
                result.add_warning(f"Subject encoding issue: {e}")

        # Check content type
        content_type = msg.get_content_type()
        if not content_type:
            result.add_warning("Missing Content-Type header")

        # Check for problematic attachments
        if msg.is_multipart():
            for part in msg.walk():
                filename = part.get_filename()
                if filename:
                    # Check for potentially problematic filenames
                    if len(filename) > 255:
                        result.add_warning(f"Long attachment filename: {len(filename)} chars")
                    if any(c in filename for c in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']):
                        result.add_warning(f"Special characters in attachment filename: {filename[:50]}")

    def _generate_pseudo_id(self, msg) -> str:
        """Generate a pseudo Message-ID from message content"""
        key_parts = [
            msg.get('From', ''),
            msg.get('Date', ''),
            msg.get('Subject', ''),
        ]
        content = '|'.join(key_parts).encode('utf-8', errors='ignore')
        return hashlib.md5(content).hexdigest()

    def _update_stats(self, result: ValidationResult):
        """Update statistics from a validation result"""
        self.stats['total'] += 1

        if result.is_valid:
            if result.warnings:
                self.stats['warnings'] += 1
            else:
                self.stats['valid'] += 1
        else:
            self.stats['errors'] += 1

    def print_report(self):
        """Print validation report"""
        print("\n" + "=" * 60)
        print("VALIDATION REPORT")
        print("=" * 60)

        print(f"\nTotal messages validated: {self.stats['total']}")
        print(f"Valid (no issues):        {self.stats['valid']}")
        print(f"Warnings only:            {self.stats['warnings']}")
        print(f"Errors:                   {self.stats['errors']}")
        print(f"Duplicates detected:      {self.stats['duplicates']}")

        # Show error details
        error_results = [r for r in self.results if r.errors]
        if error_results:
            print(f"\n{'-' * 60}")
            print("ERRORS (first 10):")
            print('-' * 60)
            for result in error_results[:10]:
                print(f"\n  File: {result.path}")
                for err in result.errors:
                    print(f"    ERROR: {err}")

            if len(error_results) > 10:
                print(f"\n  ... and {len(error_results) - 10} more files with errors")

        # Show warning details
        warning_results = [r for r in self.results if r.warnings and not r.errors]
        if warning_results:
            print(f"\n{'-' * 60}")
            print("WARNINGS (first 10):")
            print('-' * 60)
            for result in warning_results[:10]:
                print(f"\n  File: {result.path}")
                for warn in result.warnings:
                    print(f"    WARNING: {warn}")

            if len(warning_results) > 10:
                print(f"\n  ... and {len(warning_results) - 10} more files with warnings")

        # Summary
        print("\n" + "=" * 60)
        if self.stats['errors'] == 0:
            if self.stats['warnings'] == 0:
                print("All messages are valid!")
            else:
                print("All messages are valid but some have warnings.")
                print("These should still upload successfully.")
        else:
            print(f"Found {self.stats['errors']} messages with errors.")
            print("These may fail to upload. Consider reviewing them.")

        if self.stats['duplicates'] > 0:
            print(f"\nNote: {self.stats['duplicates']} potential duplicates detected.")
            print("Gmail may skip these automatically during import.")

        print("=" * 60)

    def get_problematic_files(self) -> List[str]:
        """Get list of files with errors"""
        return [r.path for r in self.results if r.errors]


def main():
    parser = argparse.ArgumentParser(
        description='Validate EML/MBOX files before Gmail migration',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s ./eml_folder/
  %(prog)s archive.mbox
  %(prog)s ./eml_folder/ --sample 100
        """
    )

    parser.add_argument('input_path',
                        help='EML directory or MBOX file')
    parser.add_argument('--sample', type=int, default=0,
                        help='Number of random files to validate (default: all)')
    parser.add_argument('--json', action='store_true',
                        help='Output results as JSON')
    parser.add_argument('--list-errors', action='store_true',
                        help='Only list files with errors')

    args = parser.parse_args()

    path = Path(args.input_path)

    if not path.exists():
        print(f"Error: Path not found: {path}")
        sys.exit(1)

    validator = EMLValidator()

    # Detect format and validate
    if path.is_file():
        ext = path.suffix.lower()
        if ext == '.mbox':
            validator.validate_mbox(str(path), sample=args.sample)
        elif ext == '.eml':
            validator.validate_directory(str(path.parent), sample=args.sample)
        else:
            # Try as MBOX
            with open(path, 'rb') as f:
                if f.read(5) == b'From ':
                    validator.validate_mbox(str(path), sample=args.sample)
                else:
                    print(f"Error: Unknown file format: {path}")
                    sys.exit(1)
    else:
        # Directory
        validator.validate_directory(str(path), sample=args.sample)

    # Output
    if args.json:
        import json
        output = {
            'stats': validator.stats,
            'errors': [r.path for r in validator.results if r.errors],
            'warnings': [r.path for r in validator.results if r.warnings and not r.errors],
        }
        print(json.dumps(output, indent=2))
    elif args.list_errors:
        for path in validator.get_problematic_files():
            print(path)
    else:
        validator.print_report()

    # Exit code based on errors
    if validator.stats['errors'] > 0:
        sys.exit(1)
    sys.exit(0)


if __name__ == '__main__':
    main()
