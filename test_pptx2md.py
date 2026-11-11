"""Tests for pptx2md (run with: pytest test_pptx2md.py -v)."""
from pathlib import Path

from pptx import Presentation

import pptx2md


def test_parse_empty_presentation(tmp_path):
    """A presentation with zero slides converts to header-only Markdown."""
    pptx_path = tmp_path / "empty.pptx"
    Presentation().save(pptx_path)

    md, format_string = pptx2md.parse_preso(Path(pptx_path))

    assert "title: empty" in md
    assert "<section" not in md
    assert format_string == "Slide%01d.svg"
