"""Aesthetic "vibes" for VEPIP funder reports.

Each vibe is a complete visual system: palette + typography + motif + layout
biases that drive both the PPTX and HTML report builders. Picking a vibe is
the agent's main creative decision — same project data renders dramatically
differently across vibes.

Vibe names map to html-ppt-skill themes 1:1 so the HTML preview and the
editable PPTX share a consistent identity.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal


@dataclass(frozen=True)
class Palette:
    bg: str         # main slide background
    surface: str    # card / panel fill
    surface_alt: str  # alternate row / band
    border: str
    text_primary: str
    text_secondary: str
    text_muted: str
    accent: str      # primary accent (e.g. gold, saffron, navy)
    accent_soft: str  # de-saturated accent for big-display text on dark bg
    danger: str
    success: str

    def as_hex(self, key: str) -> str:
        v = getattr(self, key)
        return v.lstrip("#")


@dataclass(frozen=True)
class Typography:
    headline: str         # display / cover
    body: str             # paragraph
    eyebrow_spacing: float = 5.0  # char spacing for SECTION EYEBROWS in pt
    headline_italic: bool = False
    headline_weight_bold: bool = True


@dataclass(frozen=True)
class LayoutBias:
    """Per-vibe layout knobs. The renderer reads these to vary structure.

    Two flavours of knob live here:

    1. Visual knobs (cover_dark, motif, deliverable_style, ...) — how a slide
       looks once we've decided to render it.
    2. Structural knobs (slide_count_cap, prefers_prose, ...) — *whether* a
       slide is rendered at all. The manifest in build_pptx.main() reads
       these so different vibes produce different shape decks, not just
       reskins of the same 9 slides.
    """
    cover_dark: bool                     # dark cover vs light
    cover_motif: Literal["braille", "grid", "noise", "stripes", "blocks"]
    chrome_left_bar: bool                # left accent bar on content slides
    ghost_number: bool                   # giant ghost slide number on each slide
    overview_card_count: int             # 0 / 2 / 3 stat cards on slide 2
    metrics_orientation: Literal["2x2", "1x4", "asymmetric"]
    deliverable_style: Literal["table", "cards"]
    activities_style: Literal["timeline", "cards", "list"]
    geographic_hero: bool                # giant state count on slide 7
    financial_bars_color: Literal["accent", "monochrome", "gradient"]
    way_forward_layout: Literal["two-col-cream", "stacked-dark", "split-with-image"]
    closing_panel: Literal["thanks-dark", "minimal-light", "gradient", "stamp"]

    # ── Structural knobs (Phase 4) ────────────────────────────────────────
    # Hard cap on emitted slides (including cover). Manifest trims lowest
    # priority slides when over.
    slide_count_cap: int = 11
    # Prose-heavy vibes emit overview & way-forward when text is borderline;
    # number-led vibes skip them unless data is substantial.
    prefers_prose: bool = True
    # Quote-led vibes always show a Stories slide if any testimonial exists;
    # number-led vibes only show it when the testimonial is strong (≥ 100 chars).
    prefers_quotes: bool = True
    # Chart-led vibes never skip Financials if any budget data exists, even
    # when approved=0 and only spent is known.
    prefers_charts: bool = False
    # Photo-led vibes will allocate a second activities slide just for a
    # gallery grid when ≥ 4 photos exist.
    prefers_photos: bool = False


@dataclass(frozen=True)
class Vibe:
    key: str            # canonical id, must match an html-ppt theme name
    name: str           # human label
    palette: Palette
    typography: Typography
    layout: LayoutBias
    html_theme: str     # which html-ppt theme to use for the HTML preview
    motif_label: str    # one-line designer note shown to humans

    # Optional list of html-ppt animation classes to sprinkle on slides
    animations: tuple[str, ...] = field(default_factory=tuple)


# ── Vibe 1: editorial-serif (warm, magazine, Vision Empower's default brand) ─
EDITORIAL_SERIF = Vibe(
    key="editorial-serif",
    name="Editorial Serif",
    palette=Palette(
        bg="#FAF7F2",        surface="#FFFFFF",     surface_alt="#F5EDD0",
        border="#E8DDD0",    text_primary="#2A1508", text_secondary="#3D2010",
        text_muted="#9B7B5A", accent="#C49A32",      accent_soft="#EDD98A",
        danger="#C0392B",    success="#2D7D46",
    ),
    typography=Typography(headline="Fraunces", body="Aptos", headline_italic=True),
    layout=LayoutBias(
        cover_dark=True, cover_motif="braille", chrome_left_bar=True,
        ghost_number=True, overview_card_count=2,
        metrics_orientation="2x2", deliverable_style="table",
        activities_style="timeline", geographic_hero=True,
        financial_bars_color="accent",
        way_forward_layout="two-col-cream", closing_panel="thanks-dark",
        # Editorial / magazine-feel: prose welcome, quotes lead, full deck.
        slide_count_cap=11, prefers_prose=True, prefers_quotes=True,
        prefers_charts=False, prefers_photos=False,
    ),
    html_theme="editorial-serif",
    motif_label="Warm cream + gold serif, italic display headlines, braille texture",
    animations=("fade-up", "fade-in"),
)

# ── Vibe 2: dark-premium (tokyo-night feel, luxury, big numbers) ─────────────
DARK_PREMIUM = Vibe(
    key="dark-premium",
    name="Dark Premium",
    palette=Palette(
        bg="#0F1419",        surface="#1A2027",     surface_alt="#252C36",
        border="#2D3540",    text_primary="#F5F1E8", text_secondary="#C9C2B0",
        text_muted="#7E8694", accent="#D4A537",      accent_soft="#F5E6A8",
        danger="#E5484D",    success="#46B788",
    ),
    typography=Typography(headline="Aptos Display", body="Aptos"),
    layout=LayoutBias(
        cover_dark=True, cover_motif="grid", chrome_left_bar=False,
        ghost_number=True, overview_card_count=3,
        metrics_orientation="1x4", deliverable_style="cards",
        activities_style="cards", geographic_hero=True,
        financial_bars_color="gradient",
        way_forward_layout="stacked-dark", closing_panel="thanks-dark",
        # Numbers-forward, board-pitch: drop prose-only slides unless the
        # text is substantial; only show stories when the quote is real.
        slide_count_cap=8, prefers_prose=False, prefers_quotes=False,
        prefers_charts=True, prefers_photos=False,
    ),
    html_theme="tokyo-night",
    motif_label="Deep navy/charcoal background, gold accent, jumbo numerals",
    animations=("fade-up", "scale-in"),
)

# ── Vibe 3: magazine-bold (saffron + black, brutalist confidence) ────────────
MAGAZINE_BOLD = Vibe(
    key="magazine-bold",
    name="Magazine Bold",
    palette=Palette(
        bg="#FFFFFF",        surface="#FFF8E7",     surface_alt="#FFE8B0",
        border="#1A1A1A",    text_primary="#0A0A0A", text_secondary="#1A1A1A",
        text_muted="#5A5A5A", accent="#F59E0B",      accent_soft="#FBBF24",
        danger="#DC2626",    success="#059669",
    ),
    typography=Typography(headline="Arial Black", body="Calibri",
                          eyebrow_spacing=8.0, headline_italic=False),
    layout=LayoutBias(
        cover_dark=False, cover_motif="blocks", chrome_left_bar=False,
        ghost_number=False, overview_card_count=3,
        metrics_orientation="asymmetric", deliverable_style="cards",
        activities_style="list", geographic_hero=True,
        financial_bars_color="accent",
        way_forward_layout="split-with-image", closing_panel="stamp",
        # Tight, photo-led, energetic. Cap at 6 — no time for overview prose
        # or geographic recap; quote it, show it, ship it.
        slide_count_cap=6, prefers_prose=False, prefers_quotes=True,
        prefers_charts=False, prefers_photos=True,
    ),
    html_theme="magazine-bold",
    motif_label="Saffron + black, ultra-bold sans, asymmetric blocks",
    animations=("slide-right", "fade-up"),
)

# ── Vibe 4: ocean-corporate (deep blue, swiss grid, professional) ────────────
OCEAN_CORPORATE = Vibe(
    key="ocean-corporate",
    name="Ocean Corporate",
    palette=Palette(
        bg="#FFFFFF",        surface="#F1F5F9",     surface_alt="#E2E8F0",
        border="#CBD5E1",    text_primary="#0F172A", text_secondary="#1E293B",
        text_muted="#64748B", accent="#0369A1",      accent_soft="#7DD3FC",
        danger="#DC2626",    success="#059669",
    ),
    typography=Typography(headline="Calibri", body="Calibri",
                          eyebrow_spacing=4.0, headline_weight_bold=True),
    layout=LayoutBias(
        cover_dark=False, cover_motif="grid", chrome_left_bar=True,
        ghost_number=False, overview_card_count=2,
        metrics_orientation="2x2", deliverable_style="table",
        activities_style="timeline", geographic_hero=False,
        financial_bars_color="monochrome",
        way_forward_layout="two-col-cream", closing_panel="minimal-light",
        # Audit-friendly: data-led, full coverage, no quote-driven slides
        # unless quote is strong. Charts always shown when budget exists.
        slide_count_cap=11, prefers_prose=True, prefers_quotes=False,
        prefers_charts=True, prefers_photos=False,
    ),
    html_theme="corporate-clean",
    motif_label="Deep blue + slate, grid alignment, restrained corporate trust",
    animations=("fade-up",),
)


VIBES: dict[str, Vibe] = {
    v.key: v for v in (EDITORIAL_SERIF, DARK_PREMIUM, MAGAZINE_BOLD, OCEAN_CORPORATE)
}
DEFAULT_VIBE = "editorial-serif"


def get_vibe(key: str | None) -> Vibe:
    """Return the named vibe, falling back to the default. Unknown keys
    log a warning (via print) but still return the default rather than crash."""
    if not key:
        return VIBES[DEFAULT_VIBE]
    if key in VIBES:
        return VIBES[key]
    print(f"[vibes] unknown vibe '{key}', falling back to '{DEFAULT_VIBE}'", flush=True)
    return VIBES[DEFAULT_VIBE]


def list_vibes() -> list[dict]:
    """Cheap human-readable listing for skill docs / agent prompts."""
    return [
        {
            "key": v.key,
            "name": v.name,
            "motif": v.motif_label,
            "html_theme": v.html_theme,
            "best_for": _best_for(v.key),
        }
        for v in VIBES.values()
    ]


def _best_for(key: str) -> str:
    return {
        "editorial-serif":
            "Default — most VEPIP funder reports. Warm, brand-consistent. Grant from a foundation, narrative-heavy reports, year-end summaries.",
        "dark-premium":
            "High-stakes pitches, tech-forward projects (assistive devices, software platforms), board presentations where impact numbers should feel weighty.",
        "magazine-bold":
            "Activity-rich, photo-heavy projects, public-facing reports, advocacy decks. When the story should feel energetic.",
        "ocean-corporate":
            "Corporate CSR funders (banks, multinationals), audit-friendly reports, anything that needs to read as conservative and data-led.",
    }.get(key, "General-purpose.")
