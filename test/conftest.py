"""Setup app for testing."""

import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pytest #pylint: disable=wrong-import-position
from PyQt6.QtWidgets import QApplication #pylint: disable=wrong-import-position

@pytest.fixture(scope="session")
def qapp():
    """Test Qapp"""
    app = QApplication.instance() or QApplication([])
    yield app
