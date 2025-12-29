#!/usr/bin/env python3
"""
Unit tests for PST to Gmail migration tools.

Run with: pytest tests/ -v
"""

import os
import sys
import pytest
import tempfile
import shutil
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from pst_to_gmail import detect_input_format, find_executable, MigrationStats
from eml_validator import EMLValidator, ValidationResult
from pst_analyzer import MailAnalyzer


# Path to test data
TEST_DATA_DIR = Path(__file__).parent / "test_data"


class TestInputFormatDetection:
    """Tests for detect_input_format()"""

    def test_detect_pst_file(self):
        """Should detect PST files correctly"""
        pst_path = TEST_DATA_DIR / "sample.pst"
        if pst_path.exists():
            format_type, description = detect_input_format(str(pst_path))
            assert format_type == "pst"
            assert "PST" in description

    def test_detect_eml_file(self):
        """Should detect single EML files"""
        eml_path = TEST_DATA_DIR / "valid_simple.eml"
        format_type, description = detect_input_format(str(eml_path))
        assert format_type == "eml"
        assert "EML" in description

    def test_detect_mbox_file(self):
        """Should detect MBOX files"""
        mbox_path = TEST_DATA_DIR / "sample.mbox"
        format_type, description = detect_input_format(str(mbox_path))
        assert format_type == "mbox"
        assert "MBOX" in description

    def test_detect_eml_directory(self):
        """Should detect directory containing EML files"""
        format_type, description = detect_input_format(str(TEST_DATA_DIR))
        assert format_type == "eml_dir"
        assert "EML" in description

    def test_nonexistent_path_raises(self):
        """Should raise FileNotFoundError for nonexistent paths"""
        with pytest.raises(FileNotFoundError):
            detect_input_format("/nonexistent/path/file.pst")

    def test_unknown_format_raises(self):
        """Should raise ValueError for unknown formats"""
        # Create a temporary file with unknown content
        with tempfile.NamedTemporaryFile(suffix=".xyz", delete=False) as f:
            f.write(b"random unknown content that is not any mail format")
            temp_path = f.name

        try:
            with pytest.raises(ValueError):
                detect_input_format(temp_path)
        finally:
            os.unlink(temp_path)


class TestFindExecutable:
    """Tests for find_executable()"""

    def test_find_python(self):
        """Should find Python executable"""
        result = find_executable("python3") or find_executable("python")
        assert result is not None

    def test_custom_path_valid(self):
        """Should return custom path if it exists and is executable"""
        # Use Python as a test executable
        python_path = shutil.which("python3") or shutil.which("python")
        if python_path:
            result = find_executable("anything", custom_path=python_path)
            assert result == python_path

    def test_custom_path_invalid(self):
        """Should return None for invalid custom path"""
        result = find_executable("anything", custom_path="/nonexistent/path")
        assert result is None

    def test_nonexistent_executable(self):
        """Should return None for nonexistent executables"""
        result = find_executable("this_executable_definitely_does_not_exist_12345")
        assert result is None


class TestMigrationStats:
    """Tests for MigrationStats class"""

    def test_initial_state(self):
        """Stats should start at zero"""
        stats = MigrationStats()
        assert stats.messages_found == 0
        assert stats.messages_uploaded == 0
        assert stats.messages_failed == 0

    def test_duration_before_start(self):
        """Duration should be N/A before timing"""
        stats = MigrationStats()
        assert stats.duration == "N/A"

    def test_duration_calculation(self):
        """Duration should calculate correctly"""
        import time
        stats = MigrationStats()
        stats.start()
        time.sleep(0.1)  # Sleep 100ms
        stats.finish()
        # Duration should be non-N/A
        assert stats.duration != "N/A"
        assert "s" in stats.duration  # Should contain seconds


class TestEMLValidator:
    """Tests for EML validation"""

    def test_valid_simple_email(self):
        """Valid email should pass validation"""
        validator = EMLValidator()
        validator.validate_directory(str(TEST_DATA_DIR), sample=0)

        # Find result for valid_simple.eml
        valid_results = [r for r in validator.results
                        if "valid_simple.eml" in r.path]
        assert len(valid_results) == 1
        assert valid_results[0].is_valid

    def test_missing_from_header(self):
        """Email without From header should have error"""
        validator = EMLValidator()
        validator.validate_directory(str(TEST_DATA_DIR), sample=0)

        # Find result for malformed_no_from.eml
        results = [r for r in validator.results
                  if "malformed_no_from.eml" in r.path]
        assert len(results) == 1
        assert not results[0].is_valid
        assert any("From" in err for err in results[0].errors)

    def test_missing_date_header(self):
        """Email without Date header should have warning"""
        validator = EMLValidator()
        validator.validate_directory(str(TEST_DATA_DIR), sample=0)

        # Find result for malformed_no_date.eml
        results = [r for r in validator.results
                  if "malformed_no_date.eml" in r.path]
        assert len(results) == 1
        # Missing Date is a warning, not error
        assert any("Date" in warn for warn in results[0].warnings)

    def test_duplicate_detection(self):
        """Duplicate Message-IDs should be detected"""
        validator = EMLValidator()
        validator.validate_directory(str(TEST_DATA_DIR), sample=0)

        # Should detect duplicate
        assert validator.stats["duplicates"] >= 1

    def test_attachment_detection(self):
        """Email with attachment should be detected"""
        validator = EMLValidator()
        validator.validate_directory(str(TEST_DATA_DIR), sample=0)

        # Find result for valid_with_attachment.eml
        results = [r for r in validator.results
                  if "valid_with_attachment.eml" in r.path]
        assert len(results) == 1
        # Should be valid (attachments are fine)
        assert results[0].is_valid

    def test_validation_stats(self):
        """Validation should produce correct stats"""
        validator = EMLValidator()
        validator.validate_directory(str(TEST_DATA_DIR), sample=0)

        # Should have processed all EML files
        eml_count = len(list(TEST_DATA_DIR.glob("*.eml")))
        assert validator.stats["total"] == eml_count

    def test_mbox_validation(self):
        """Should validate MBOX files"""
        validator = EMLValidator()
        mbox_path = TEST_DATA_DIR / "sample.mbox"
        validator.validate_mbox(str(mbox_path))

        # Should have found messages in MBOX
        assert validator.stats["total"] >= 2


class TestMailAnalyzer:
    """Tests for mail analysis"""

    def test_analyze_eml_directory(self):
        """Should analyze EML directory"""
        analyzer = MailAnalyzer()
        stats = analyzer.analyze_eml_directory(str(TEST_DATA_DIR))

        # Should have found messages
        assert stats["total_messages"] > 0
        assert stats["total_size_bytes"] > 0

    def test_analyze_mbox(self):
        """Should analyze MBOX file"""
        analyzer = MailAnalyzer()
        mbox_path = TEST_DATA_DIR / "sample.mbox"
        stats = analyzer.analyze_mbox(str(mbox_path))

        # Should have found messages
        assert stats["total_messages"] >= 2

    def test_date_extraction(self):
        """Should extract date range from messages"""
        analyzer = MailAnalyzer()
        analyzer.analyze_eml_directory(str(TEST_DATA_DIR))

        # Should have date range
        assert analyzer.stats["date_range"]["earliest"] is not None
        assert analyzer.stats["date_range"]["latest"] is not None

    def test_sender_extraction(self):
        """Should extract sender statistics"""
        analyzer = MailAnalyzer()
        analyzer.analyze_eml_directory(str(TEST_DATA_DIR))

        # Should have found senders
        assert len(analyzer.stats["senders"]) > 0

    def test_json_export(self):
        """Should export to JSON format"""
        import json
        analyzer = MailAnalyzer()
        analyzer.analyze_eml_directory(str(TEST_DATA_DIR))

        json_output = analyzer.to_json()
        data = json.loads(json_output)

        assert "total_messages" in data
        assert "total_size_bytes" in data
        assert "date_range" in data

    @pytest.mark.skipif(
        shutil.which("readpst") is None,
        reason="readpst not installed"
    )
    def test_analyze_pst(self):
        """Should analyze PST file (requires readpst)"""
        analyzer = MailAnalyzer()
        pst_path = TEST_DATA_DIR / "sample.pst"

        if pst_path.exists():
            stats = analyzer.analyze_pst(str(pst_path))
            # Should have size at minimum
            assert stats["total_size_bytes"] > 0


class TestValidationResult:
    """Tests for ValidationResult class"""

    def test_initial_state(self):
        """New result should be valid with no issues"""
        result = ValidationResult("/test/path.eml")
        assert result.is_valid
        assert not result.has_issues

    def test_add_warning(self):
        """Adding warning should mark has_issues but still valid"""
        result = ValidationResult("/test/path.eml")
        result.add_warning("Test warning")
        assert result.is_valid
        assert result.has_issues
        assert len(result.warnings) == 1

    def test_add_error(self):
        """Adding error should mark as invalid"""
        result = ValidationResult("/test/path.eml")
        result.add_error("Test error")
        assert not result.is_valid
        assert result.has_issues
        assert len(result.errors) == 1


class TestIntegration:
    """Integration tests"""

    @pytest.mark.skipif(
        shutil.which("readpst") is None,
        reason="readpst not installed"
    )
    def test_pst_conversion(self):
        """Test PST to EML conversion (requires readpst)"""
        from pst_to_gmail import convert_pst_to_eml

        pst_path = TEST_DATA_DIR / "sample.pst"
        if not pst_path.exists():
            pytest.skip("Sample PST not available")

        readpst = shutil.which("readpst")

        with tempfile.TemporaryDirectory() as tmpdir:
            success, count = convert_pst_to_eml(
                str(pst_path),
                tmpdir,
                readpst,
                dry_run=False
            )

            assert success
            # readpst may create EML files or mbox files depending on PST structure
            # Check for any output files (eml or mbox)
            eml_files = list(Path(tmpdir).rglob("*.eml"))
            mbox_files = list(Path(tmpdir).rglob("mbox")) + list(Path(tmpdir).rglob("*.mbox"))
            total_files = len(eml_files) + len(mbox_files)
            assert total_files > 0, "Should have created EML or mbox files"

    def test_dry_run_mode(self):
        """Dry run should not modify anything"""
        from pst_to_gmail import convert_pst_to_eml

        pst_path = TEST_DATA_DIR / "sample.pst"
        if not pst_path.exists():
            pytest.skip("Sample PST not available")

        readpst = shutil.which("readpst")
        if not readpst:
            pytest.skip("readpst not installed")

        with tempfile.TemporaryDirectory() as tmpdir:
            success, count = convert_pst_to_eml(
                str(pst_path),
                tmpdir,
                readpst,
                dry_run=True
            )

            assert success
            # Dry run should not create files
            eml_files = list(Path(tmpdir).rglob("*.eml"))
            assert len(eml_files) == 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
