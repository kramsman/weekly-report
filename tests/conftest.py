"""
Shared pytest fixtures for weekly-report tests.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
HOW TO RUN TESTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Run all unit tests (no external services):
    pytest tests/ -s

Run integration tests only (hits real APIs — needs credentials):
    pytest tests/ -s -m integration

Run everything:
    pytest tests/ -s -m "integration or not integration"

Run a single test file:
    pytest tests/test_error_logging.py -s

Run a single test by name:
    pytest tests/test_error_logging.py::test_send_drive_notification_logs_on_failure -s

Verbose output (shows test names):
    pytest tests/ -s -v

MARKERS
    integration  Tests that call external services.
                 Skipped automatically if required env vars are not set.

CREDENTIALS NEEDED FOR INTEGRATION TESTS
    None currently — all tests use mocks.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import pytest


@pytest.fixture
def api_key_file(tmp_path):
    """A temp file containing a fake SendGrid API key."""
    key_file = tmp_path / "sendgrid_api_key.txt"
    key_file.write_text("fake-sendgrid-key")
    return key_file


@pytest.fixture
def error_log_file(tmp_path):
    """A temp path for the upload error log (does not exist until written)."""
    return tmp_path / "upload_errors.txt"
