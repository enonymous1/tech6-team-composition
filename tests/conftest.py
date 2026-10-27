"""Shared fixtures."""

import pytest


@pytest.fixture()
def examples_dir():
    from pathlib import Path

    return Path(__file__).resolve().parent.parent / "examples"
