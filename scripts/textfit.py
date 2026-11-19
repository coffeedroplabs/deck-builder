"""Predict how PowerPoint wraps and stacks text, from real font metrics via Pillow.

The builder uses this to refuse text that will not fit before PowerPoint ever
opens the file. render.py's measurement through actual PowerPoint is the
authority when it is available; this is the cheap offline estimate.
"""

from functools import lru_cache
from glob import glob
import os
import shutil
import subprocess

from PIL import ImageFont

# PowerPoint's single line spacing is 1.2 x the font size. Measured through
# PowerPoint for Arial, Georgia, Calibri, Times New Roman and Cambria at 12 to
# 40 pt, regular and bold: every combination came out at exactly 1.2.
LINE_HEIGHT = 1.2

# Style order is (regular, bold, italic, bold italic).
FONT_FILES = {
    "arial": ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "georgia": ("georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf"),
    "calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "cambria": ("cambria.ttc", "cambriab.ttf", "cambriai.ttf", "cambriaz.ttf"),
    "times new roman": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "courier new": ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"),
}

# Metric-compatible Linux families; substituting one keeps line breaks intact.
SUBSTITUTES = {
    "arial": "Liberation Sans",
    "times new roman": "Liberation Serif",
    "courier new": "Liberation Mono",
    "calibri": "Carlito",
    "cambria": "Caladea",
    "georgia": "Gelasio",
}

FONT_DIRS = [
    "/mnt/c/Windows/Fonts",
    "C:/Windows/Fonts",
    os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts"),
    "/Library/Fonts",
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    os.path.expanduser("~/Library/Fonts"),
    "/usr/share/fonts",
    "/usr/local/share/fonts",
    os.path.expanduser("~/.fonts"),
    os.path.expanduser("~/.local/share/fonts"),
] + sorted(glob("/mnt/c/Users/*/AppData/Local/Microsoft/Windows/Fonts"))

FONT_SUFFIXES = (".ttf", ".otf", ".ttc")


@lru_cache(maxsize=None)
def _font_index():
    """Lowercase file name -> path, first directory wins, so FONT_DIRS order is precedence."""
    index = {}
    for directory in FONT_DIRS:
        if not os.path.isdir(directory):
            continue
        for root, _dirs, files in os.walk(directory):
            for name in files:
                if name.lower().endswith(FONT_SUFFIXES):
                    index.setdefault(name.lower(), os.path.join(root, name))
    return index


# Names fonts give their upright and slanted faces, per style slot.
STYLE_NAMES = {
    "regular": ("regular", "normal", "roman", "book"),
    "bold": ("bold",),
    "italic": ("italic", "oblique"),
    "bold:italic": ("bold italic", "bold oblique"),
}


@lru_cache(maxsize=None)
def _family_index():
    """(lowercase family, lowercase style) -> path, read from each font's own
    name table. Template fonts outside FONT_FILES (Trebuchet MS, Segoe UI)
    live in the Windows fonts folder, which fontconfig cannot see from WSL;
    without this they were measured in DejaVu Sans, about 9% wider, so their
    text shrank a step it did not need. Slow (seconds over /mnt/c), so it is
    built only when a family is missing everywhere else."""
    index = {}
    for path in _font_index().values():
        try:
            # Not _load: its cache would keep every installed face open.
            family, style = ImageFont.truetype(path, 10).getname()
        except (OSError, ValueError):
            continue
        if family and style:
            index.setdefault((family.lower(), style.lower()), path)
    return index


def _from_family_index(name, style):
    index = _family_index()
    for wanted in (style, "regular"):
        for label in STYLE_NAMES[wanted]:
            path = index.get((name, label))
            if path:
                return path
    return None


def _fc_match(query):
    """Ask fontconfig for a <file, families> pair; None when fontconfig is absent or silent."""
    if not shutil.which("fc-match"):
        return None
    try:
        done = subprocess.run(
            ["fc-match", "-f", "%{file}|%{family}", query],
            capture_output=True, text=True, timeout=10,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    line = (done.stdout or "").strip().split("\n")[0]
    if "|" not in line:
        return None
    path, families = line.split("|", 1)
    return (path.strip() or None, [f.strip() for f in families.split(",") if f.strip()])


@lru_cache(maxsize=None)
def font_file(family, bold=False, italic=False):
    """Return (path, exact) for a family and style; exact is False for a substitute."""
    name = (family or "").strip().lower()
    if not name:
        return (None, False)
    style = "bold:italic" if (bold and italic) else "bold" if bold else "italic" if italic else "regular"

    if name in FONT_FILES:
        index = _font_index()
        files = FONT_FILES[name]
        slot = 3 if (bold and italic) else 1 if bold else 2 if italic else 0
        path = index.get(files[slot].lower())
        if path:
            return (path, True)
        path = index.get(files[0].lower())
        if path:
            return (path, True)

    match = _fc_match(f"{family}:{style}")
    path, families = match if match else (None, [])
    if path and os.path.exists(path) and any(f.lower() == name for f in families):
        return (path, True)
    found = _from_family_index(name, style)
    if found:
        return (found, True)
    if not path or not os.path.exists(path):
        return (None, False)
    substitute = SUBSTITUTES.get(name)
    if substitute:
        alt = _fc_match(f"{substitute}:{style}")
        if alt and alt[0] and os.path.exists(alt[0]):
            if any(f.lower() == substitute.lower() for f in alt[1]):
                return (alt[0], False)
    return (path, False)


@lru_cache(maxsize=None)
def _load(path):
    """A 1000 px face; widths scale linearly from there. Cached because loading is slow."""
    try:
        if path.lower().endswith(".ttc"):
            return ImageFont.truetype(path, 1000, index=0)
        return ImageFont.truetype(path, 1000)
    except (OSError, ValueError):
        return None


def text_width_pt(text, family, size_pt, bold=False, italic=False):
    if not text:
        return 0.0
    path, _exact = font_file(family, bold, italic)
    font = _load(path) if path else None
    if font is None:
        return len(text) * 0.5 * size_pt  # 0.5 em per character, a rough average
    return font.getlength(text) * size_pt / 1000.0


def _hyphen_parts(token):
    parts, start = [], 0
    for i, ch in enumerate(token):
        if ch == "-":
            parts.append(token[start:i + 1])
            start = i + 1
    if start < len(token):
        parts.append(token[start:])
    return parts or [token]


def _char_chunks(text, width_pt, measure):
    chunks, rest = [], text
    while rest:
        cut = 1
        for i in range(1, len(rest) + 1):
            if measure(rest[:i]) > width_pt + 0.01:
                break
            cut = i
        chunks.append(rest[:cut])
        rest = rest[cut:]
    return chunks or [""]


def _split_long(token, width_pt, measure):
    """Break an over-long token: after hyphens first, since PowerPoint prefers that, then per character."""
    chunks, run = [], ""
    for part in (_hyphen_parts(token) if "-" in token else [token]):
        if measure(run + part) <= width_pt + 0.01:
            run += part
            continue
        if run:
            chunks.append(run)
            run = ""
        if measure(part) <= width_pt + 0.01:
            run = part
        else:
            pieces = _char_chunks(part, width_pt, measure)
            chunks.extend(pieces[:-1])
            run = pieces[-1]
    if run:
        chunks.append(run)
    return chunks or [""]


def wrap_lines(text, family, size_pt, width_pt, bold=False, italic=False) -> list[str]:
    def measure(s):
        return text_width_pt(s, family, size_pt, bold, italic)

    out = []
    # \v is PowerPoint's soft line break; both it and \n end a line hard.
    for piece in text.replace("\v", "\n").split("\n"):
        lines, cur = [], ""
        for token in piece.split(" "):
            candidate = f"{cur} {token}" if cur else token
            if measure(candidate.rstrip(" ")) <= width_pt + 0.01:
                cur = candidate
                continue
            if cur.strip(" "):
                lines.append(cur.rstrip(" "))
                cur = ""
                if measure(token.rstrip(" ")) <= width_pt + 0.01:
                    cur = token
                    continue
            chunks = _split_long(token, width_pt, measure)
            lines.extend(chunks[:-1])
            cur = chunks[-1]
        if cur.strip(" ") or not lines:
            lines.append(cur.rstrip(" "))
        out.extend(lines)
    return out


def paragraph_lines(p, width_pt):
    return wrap_lines(
        p.get("text", ""), p.get("family", "Arial"), p.get("size", 18),
        width_pt - p.get("indent", 0.0), p.get("bold", False), p.get("italic", False),
    )


def block_height_pt(paragraphs, width_pt):
    height, counts = 0.0, []
    for i, p in enumerate(paragraphs):
        n = len(paragraph_lines(p, width_pt))
        counts.append(n)
        height += n * LINE_HEIGHT * p["size"] * p.get("line_spacing", 1.0)
        if i:  # the builder never puts space before the first paragraph
            height += p.get("space_before", 0.0)
    return height, counts


def fits(paragraphs, box_w_pt, box_h_pt, insets=(0, 0, 0, 0), max_lines=None):
    left, top, right, bottom = insets
    width = box_w_pt - left - right
    available = box_h_pt - top - bottom
    height, per_paragraph = block_height_pt(paragraphs, width)
    lines = sum(per_paragraph)
    ok = height <= available + 0.5  # half a point of rounding slack
    if max_lines is not None and lines > max_lines:
        ok = False
    return {
        "fits": ok,
        "needed": round(height, 1),
        "available": round(available, 1),
        "lines": lines,
        "per_paragraph": per_paragraph,
    }


def fit_size(make_paragraphs, sizes, box_w_pt, box_h_pt, insets=(0, 0, 0, 0), max_lines=None):
    result = None
    for size in sizes:
        result = fits(make_paragraphs(size), box_w_pt, box_h_pt, insets, max_lines)
        if result["fits"]:
            return size, result
    return None, result


def size_steps(base, smallest, step=2):
    steps, size = [], base
    while size > smallest:
        steps.append(size)
        size -= step
    if not steps or steps[-1] != smallest:
        steps.append(smallest)
    return steps
