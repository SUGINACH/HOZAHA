#!/usr/bin/env python3
"""Shared parsing and validation for Word review-comment labels."""
from __future__ import annotations

import re
from collections.abc import Iterable

_COMMENT_LABEL = re.compile(r"^\s*((?:E-\d{4}|S-\d{8}-\d{2,}))\s*\|")
_SESSION_LABEL = re.compile(r"S-\d{8}-\d{2,}\Z")


def extract_comment_label(comment_text: str) -> str | None:
    """Return the exact leading label from a Word comment, or None."""
    match = _COMMENT_LABEL.match(comment_text)
    return match.group(1) if match else None


def validate_session_labels(labels: Iterable[str]) -> None:
    """Raise ValueError for malformed or duplicate session labels."""
    values = list(labels)
    if any(_SESSION_LABEL.fullmatch(label) is None for label in values):
        raise ValueError("Invalid session comment label format")
    if len(values) != len(set(values)):
        raise ValueError("Duplicate comment labels in session configuration")
