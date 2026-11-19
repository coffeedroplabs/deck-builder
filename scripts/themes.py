"""Deck profiles, theme presets and the slide grid, shared by every script.

Every threshold a gate enforces lives here, so a profile can be tuned in one
place. Evidence strength for each number is recorded in
references/design-rules.md; the weakly sourced ones (read-ahead sizes, title
length, margin) are conventions, not research results.
"""

VERSION = "1.0.0"

EMU_PER_INCH = 914400
EMU_PER_PT = 12700

# Office's 16:9 size. Template decks may use another size, so geometry below is
# expressed as fractions of the slide and scaled at build time.
SLIDE_W = 12192000
SLIDE_H = 6858000

# Each size role is (base, smallest). The builder starts at base and steps down
# 2 pt at a time to the smallest before it refuses to build; it never goes
# lower, because shrinking text to make it fit is the failure the floors exist
# to stop.
PROFILES = {
    "live": {
        "description": "Presented in a room or on a call; read from a distance while someone talks.",
        "sizes": {
            "cover_title": (44, 40), "title": (36, 32), "statement": (40, 32),
            "body": (28, 24), "body2": (24, 20), "small": (20, 18),
            "chart": (18, 18), "table": (18, 18), "source": (12, 12),
            "big_number": (96, 72), "quote": (32, 28),
        },
        "floor": 18, "target": 24, "source_floor": 10,
        "max_bullets": 4, "max_levels": 2, "max_bullet_lines": 2,
        "max_title_words": 15, "max_title_lines": 2,
        "max_words": 50, "notes_required": True,
        "max_table_rows": 6, "max_table_cols": 5,
        "max_categories": 8, "max_series": 3, "max_pie_slices": 5,
    },
    "pitch": {
        "description": "Short persuasive deck, presented or sent; one idea per slide, large type.",
        "sizes": {
            "cover_title": (40, 36), "title": (32, 28), "statement": (36, 28),
            "body": (24, 20), "body2": (20, 18), "small": (18, 16),
            "chart": (16, 16), "table": (16, 16), "source": (11, 11),
            "big_number": (88, 64), "quote": (28, 24),
        },
        "floor": 16, "target": 20, "source_floor": 9,
        "max_bullets": 4, "max_levels": 2, "max_bullet_lines": 2,
        "max_title_words": 15, "max_title_lines": 2,
        "max_words": 60, "notes_required": True,
        "max_table_rows": 7, "max_table_cols": 5,
        "max_categories": 10, "max_series": 3, "max_pie_slices": 5,
    },
    "read-ahead": {
        "description": "Read on a screen without a presenter; denser, but still one message per slide.",
        "sizes": {
            "cover_title": (36, 32), "title": (24, 20), "statement": (32, 24),
            "body": (16, 14), "body2": (14, 12), "small": (12, 11),
            "chart": (12, 10), "table": (12, 10), "source": (9, 9),
            "big_number": (60, 48), "quote": (24, 20),
        },
        "floor": 10, "target": 14, "source_floor": 8,
        "max_bullets": 6, "max_levels": 2, "max_bullet_lines": 3,
        "max_title_words": 15, "max_title_lines": 2,
        "max_words": 120, "notes_required": False,
        "max_table_rows": 12, "max_table_cols": 7,
        "max_categories": 14, "max_series": 4, "max_pie_slices": 5,
    },
}

# WCAG 2.x 1.4.3 and 1.4.11. Large text is 18 pt, or 14 pt bold.
CONTRAST = {"text": 4.5, "large": 3.0, "nontext": 3.0, "strict": 7.0}

MAX_FONT_FAMILIES = 2
MAX_SAME_LAYOUT_RUN = 2
MAX_BIG_NUMBERS = 2
SAFE_MARGIN = 0.05  # fraction of slide width/height; a synthesis, so a warning only

# Fonts whose metrics match on Windows, Mac and the metric-compatible Linux
# substitutes (Liberation, Carlito, Caladea, Gelasio), so line breaks survive a
# trip to another machine.
METRIC_SAFE_FONTS = {"Arial", "Calibri", "Cambria", "Georgia", "Times New Roman", "Courier New"}

# Colours are 6-digit hex without '#', because a '#' or 8-digit value makes
# PowerPoint ask to repair the file.
#   bg      slide background        text    body and titles
#   muted   secondary text (source lines, labels); must still pass 4.5:1
#   rule    hairlines and table rules; decorative, no contrast requirement
#   accent  the one emphasis colour: highlighted series, big numbers
#   series  non-highlighted chart marks; must pass 3:1 against bg
#   third   a second data colour for the rare chart that needs two emphases
THEMES = {
    "paper": {
        "description": "White page, near-black text, one blue accent, serif titles. Safe default.",
        "dark": False,
        "fonts": {"major": "Georgia", "minor": "Arial"},
        "title_bold": False,
        "colors": {
            "bg": "FFFFFF", "text": "1F2328", "muted": "5B6370", "rule": "C9CDD3",
            "accent": "1F5FAD", "series": "868D97", "third": "B4441C",
        },
    },
    "consulting": {
        "description": "White page, navy accent, bold sans titles. Dense read-aheads and board updates.",
        "dark": False,
        "fonts": {"major": "Arial", "minor": "Arial"},
        "title_bold": True,
        "colors": {
            "bg": "FFFFFF", "text": "1A1A1A", "muted": "595959", "rule": "BFBFBF",
            "accent": "0B3D91", "series": "8A8A8A", "third": "2E7D4F",
        },
    },
    "ink": {
        "description": "Near-black page, light text, amber accent. Keynotes in dim rooms.",
        "dark": True,
        "fonts": {"major": "Calibri", "minor": "Calibri"},
        "title_bold": True,
        "colors": {
            "bg": "15181D", "text": "EDEEF0", "muted": "A3AAB5", "rule": "3A4048",
            "accent": "F0B429", "series": "7C8591", "third": "6CB4EE",
        },
    },
}

# Every colour is written into the deck as a theme slot, never as a fixed RGB,
# so recolouring the theme (or a client template) recolours the whole deck.
ROLE_SCHEME = {
    "text": "tx1", "bg": "bg1", "muted": "tx2", "rule": "bg2",
    "accent": "accent1", "series": "accent2", "third": "accent3",
}

# The slide grid as fractions of slide width (x, w) and height (y, h). Side
# margins follow Office's 16:9 master (0.917 in of 13.333 in). Office's own
# footer sits lower, but that leaves less than the 5% safe margin, so the
# footer band here ends 0.4 in above the slide edge.
GRID = {
    "margin_x": 0.06875,
    "title_y": 0.0533, "title_h": 0.1933,
    "body_y": 0.2667, "body_bottom": 0.8867,
    "footer_y": 0.9067, "footer_h": 0.04,
    "gutter": 0.03,
}


def profile(name):
    try:
        return PROFILES[name]
    except KeyError:
        raise ValueError(f"unknown profile {name!r}; choose one of {', '.join(PROFILES)}") from None


def resolve_theme(spec_theme):
    """Return a full theme dict from a preset name or a {"base", "colors", "fonts"} override."""
    if spec_theme is None:
        spec_theme = "paper"
    if isinstance(spec_theme, str):
        if spec_theme not in THEMES:
            raise ValueError(f"unknown theme {spec_theme!r}; choose one of {', '.join(THEMES)}")
        base = THEMES[spec_theme]
        return {**base, "name": spec_theme, "colors": dict(base["colors"]), "fonts": dict(base["fonts"])}
    if not isinstance(spec_theme, dict):
        raise ValueError("theme must be a preset name or an object with 'base', 'colors', 'fonts'")
    unknown = set(spec_theme) - {"base", "colors", "fonts", "title_bold"}
    if unknown:
        raise ValueError(f"unknown theme keys: {', '.join(sorted(unknown))}")
    theme = resolve_theme(spec_theme.get("base", "paper"))
    for key, value in (spec_theme.get("colors") or {}).items():
        if key not in theme["colors"]:
            raise ValueError(f"unknown theme colour {key!r}; known: {', '.join(theme['colors'])}")
        value = str(value).lstrip("#").upper()
        if len(value) != 6 or any(c not in "0123456789ABCDEF" for c in value):
            raise ValueError(f"theme colour {key} must be 6 hex digits, got {value!r}")
        theme["colors"][key] = value
    for key, value in (spec_theme.get("fonts") or {}).items():
        if key not in ("major", "minor"):
            raise ValueError(f"unknown theme font {key!r}; use 'major' (titles) and 'minor' (body)")
        theme["fonts"][key] = value
    if "title_bold" in spec_theme:
        theme["title_bold"] = bool(spec_theme["title_bold"])
    theme["name"] = "custom"
    return theme


def relative_luminance(hex6):
    def channel(v):
        v = v / 255
        return v / 12.92 if v <= 0.04045 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hex6[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def contrast_ratio(a, b):
    la, lb = sorted((relative_luminance(a), relative_luminance(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def theme_contrast_problems(theme):
    """Pairs in a theme that fail WCAG, so a brand override cannot ship unreadable colours."""
    c = theme["colors"]
    checks = [
        ("text", c["text"], CONTRAST["text"]),
        ("muted", c["muted"], CONTRAST["text"]),
        ("accent", c["accent"], CONTRAST["large"]),
        ("series", c["series"], CONTRAST["nontext"]),
        ("third", c["third"], CONTRAST["nontext"]),
    ]
    problems = []
    for name, colour, need in checks:
        ratio = contrast_ratio(colour, c["bg"])
        if ratio < need:
            problems.append(f"{name} {colour} on bg {c['bg']} is {ratio:.2f}:1, needs {need}:1")
    return problems
