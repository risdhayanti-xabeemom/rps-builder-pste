"""Tema neo-brutalism bersama untuk app Streamlit.

File ini dan ``assets/neobrutal.css`` identik di repo RPS Builder dan Dashboard CPL.
Pakai: ``apply_theme()`` sekali di awal ``main()``, lalu ``page_header(...)`` untuk
header halaman. Tidak ada logika bisnis di sini.
"""

from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Iterable, Optional

import streamlit as st

_CSS_PATH = Path(__file__).parent / "assets" / "neobrutal.css"

BADGE_TONES = {"neutral", "yellow", "mint", "coral", "lilac", "indigo"}


def apply_theme() -> None:
    """Suntikkan CSS tema ke halaman. Jika file CSS hilang, app tetap jalan dengan tema bawaan."""
    try:
        css = _CSS_PATH.read_text(encoding="utf-8")
    except OSError:
        return
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def badge(text: str, tone: str = "neutral") -> str:
    """Kembalikan HTML badge kecil. ``tone``: neutral, yellow, mint, coral, lilac, indigo."""
    tone = tone if tone in BADGE_TONES else "neutral"
    cls = "nb-badge" if tone == "neutral" else f"nb-badge is-{tone}"
    return f'<span class="{cls}">{escape(text)}</span>'


def page_header(
    title: str,
    subtitle_lines: Iterable[str] = (),
    chips: Iterable[tuple[str, str]] = (),
    sticker: Optional[str] = None,
) -> None:
    """Header halaman: judul, subjudul (beberapa baris), badge, dan stiker miring opsional."""
    subtitle = "<br>".join(escape(line) for line in subtitle_lines)
    chip_html = "".join(badge(text, tone) for text, tone in chips)
    sticker_html = (
        f'<span class="nb-badge is-coral nb-sticker">{escape(sticker)}</span>' if sticker else ""
    )
    st.markdown(
        f"""
<div class="nb-hero">
{sticker_html}
<h1>{escape(title)}</h1>
<div class="subtitle">{subtitle}</div>
<div class="nb-chips">{chip_html}</div>
</div>
""",
        unsafe_allow_html=True,
    )
