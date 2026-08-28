"""Pytest configuration: isolate FORGE_HOME from the real user directory."""

import os
import tempfile

_TEST_FORGE_HOME = tempfile.mkdtemp(prefix="forge_home_tests_")
os.environ["FORGE_HOME"] = _TEST_FORGE_HOME
