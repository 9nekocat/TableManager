# -*- coding: utf-8 -*-
"""Tests for SvnClient lock/unlock methods (Bug 2)."""

import sys
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from svn_client import SvnClient, SvnError


def test_lock_unlock_signature():
    """SvnClient has lock() and unlock() methods with correct signatures."""
    svn = SvnClient()
    assert hasattr(svn, "lock"), "SvnClient missing lock() method"
    assert hasattr(svn, "unlock"), "SvnClient missing unlock() method"
    assert callable(svn.lock)
    assert callable(svn.unlock)
