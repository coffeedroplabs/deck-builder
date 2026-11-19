# /// script
# requires-python = ">=3.10"
# dependencies = ["pillow==13.0.0"]
# ///
"""Render a deck to PNGs and measure its text with the real layout engine.

    uv run render.py deck.pptx -o render/

Writes slide-NN.png, contact.png and measure.json. With PowerPoint (Windows,
or WSL calling Windows), measure.json holds every shape's position and its
text's measured bounds and line counts, which is what the lint uses to catch
overflow; a guess from font metrics is not enough to sign off a deck.
LibreOffice is the fallback: pictures only, no measurement.

PowerPoint is driven read-only and invisibly. Only the presentation this
script opened is closed, and PowerPoint is quit only if it was not already
running, so a user's open decks are never touched.
"""

import argparse
import fcntl
import glob
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from contextlib import contextmanager

from PIL import Image, ImageDraw, ImageFont

WIDTH = 1600
THUMB = (400, 225)
THUMB_COLS = 4
LOCK_PATH = os.path.expanduser("~/.cache/deck-builder/render.lock")

PS1 = r"""
param([string]$Deck, [string]$Out, [int]$Width)
$ErrorActionPreference = 'Stop'
$before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count
$app = New-Object -ComObject PowerPoint.Application
$pres = $null
try {
  # ReadOnly, not untitled, no window: nothing appears in the user's PowerPoint.
  $pres = $app.Presentations.Open($Deck, -1, 0, 0)
  $w = $pres.PageSetup.SlideWidth
  $h = $pres.PageSetup.SlideHeight
  $height = [int][math]::Round($Width * $h / $w)
  $slides = @()
  foreach ($s in $pres.Slides) {
    $s.Export((Join-Path $Out ('slide-{0:D2}.png' -f $s.SlideIndex)), 'PNG', $Width, $height)
    $shapes = @()
    foreach ($sh in $s.Shapes) {
      $item = [ordered]@{ name = $sh.Name; id = $sh.Id; left = $sh.Left; top = $sh.Top;
                          width = $sh.Width; height = $sh.Height; has_text = $false;
                          has_table = ($sh.HasTable -eq -1); has_chart = ($sh.HasChart -eq -1) }
      if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame2.HasText -eq -1) {
        $tf = $sh.TextFrame2
        $tr = $tf.TextRange
        $item.has_text = $true
        $item.bound_left = $tr.BoundLeft
        $item.bound_top = $tr.BoundTop
        $item.bound_width = $tr.BoundWidth
        $item.bound_height = $tr.BoundHeight
        $item.margin_left = $tf.MarginLeft
        $item.margin_right = $tf.MarginRight
        $item.margin_top = $tf.MarginTop
        $item.margin_bottom = $tf.MarginBottom
        $item.autosize = $tf.AutoSize
        $item.word_wrap = $tf.WordWrap
        $old = $sh.TextFrame.TextRange
        $item.lines = $old.Lines().Count
        $per = @()
        for ($i = 1; $i -le $old.Paragraphs().Count; $i++) { $per += $old.Paragraphs($i).Lines().Count }
        $item.paragraph_lines = $per
        $texts = @()
        for ($i = 1; $i -le $item.lines; $i++) { $texts += $old.Lines($i, 1).Text }
        $item.line_texts = $texts
        $sizes = @{}; $fonts = @{}
        foreach ($run in $tr.Runs()) { $sizes[[string]$run.Font.Size] = 1; $fonts[[string]$run.Font.Name] = 1 }
        $item.font_sizes = @($sizes.Keys)
        $item.font_names = @($fonts.Keys)
      }
      $shapes += [pscustomobject]$item
    }
    $slides += [pscustomobject]@{ index = $s.SlideIndex; shapes = $shapes }
  }
  $result = [ordered]@{ backend = 'powerpoint'; app_version = $app.Version; slide_width = $w;
                        slide_height = $h; slides = $slides }
  $result | ConvertTo-Json -Depth 8 | Out-File -FilePath (Join-Path $Out 'measure.json') -Encoding utf8
}
finally {
  if ($pres -ne $null) { $pres.Close() }
  if ($before -eq 0) { $app.Quit() }
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
"""


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


@contextmanager
def render_lock():
    # PowerPoint is one COM server per machine; two renders at once fight
    # over it, so parallel agents queue here instead.
    os.makedirs(os.path.dirname(LOCK_PATH), exist_ok=True)
    with open(LOCK_PATH, "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)


def _run(cmd, timeout):
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)


def powerpoint_available():
    exe = shutil.which("powershell.exe")
    if not exe:
        return False
    try:
        r = _run([exe, "-NoProfile", "-Command",
                  "if ([type]::GetTypeFromProgID('PowerPoint.Application')) { 'yes' } else { 'no' }"], 60)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return r.stdout.strip().endswith("yes")


def _windows_temp():
    r = _run(["powershell.exe", "-NoProfile", "-Command", "[System.IO.Path]::GetTempPath()"], 60)
    win = r.stdout.strip().replace("\r", "")
    return subprocess.run(["wslpath", "-u", win], capture_output=True, text=True, check=True).stdout.strip()


def _to_windows(path):
    return subprocess.run(["wslpath", "-w", path], capture_output=True, text=True, check=True).stdout.strip()


def render_powerpoint(deck, out_dir, width, timeout):
    on_windows = os.name == "nt"
    base = tempfile.gettempdir() if on_windows else _windows_temp()
    # PowerPoint cannot reliably open files on the WSL file system, so the
    # deck and the script run from a fresh folder in Windows' temp.
    work = os.path.join(base, f"deck-builder-{uuid.uuid4().hex}")
    os.makedirs(work)
    try:
        shutil.copy2(deck, os.path.join(work, "deck.pptx"))
        script = os.path.join(work, "render.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(PS1)
        conv = (lambda p: p) if on_windows else _to_windows
        exe = "powershell.exe" if not on_windows else "powershell"
        r = _run([exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", conv(script),
                  "-Deck", conv(os.path.join(work, "deck.pptx")), "-Out", conv(work), "-Width", str(width)],
                 timeout)
        if r.returncode != 0 or not os.path.exists(os.path.join(work, "measure.json")):
            raise RuntimeError(f"PowerPoint render failed:\n{r.stdout}\n{r.stderr}".replace("\r", ""))
        with open(os.path.join(work, "measure.json"), encoding="utf-8-sig") as f:
            measure = json.load(f)
        for png in glob.glob(os.path.join(work, "slide-*.png")):
            shutil.copy2(png, out_dir)
        return measure
    finally:
        shutil.rmtree(work, ignore_errors=True)


def render_libreoffice(deck, out_dir, width, timeout):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        raise RuntimeError("neither PowerPoint nor LibreOffice (soffice + pdftoppm) is available to render")
    with tempfile.TemporaryDirectory() as work:
        r = _run([soffice, "--headless", "--convert-to", "pdf", "--outdir", work, deck], timeout)
        pdf = os.path.join(work, os.path.splitext(os.path.basename(deck))[0] + ".pdf")
        if r.returncode != 0 or not os.path.exists(pdf):
            raise RuntimeError(f"LibreOffice conversion failed:\n{r.stdout}\n{r.stderr}")
        _run([pdftoppm, "-png", "-scale-to-x", str(width), "-scale-to-y", "-1", pdf,
              os.path.join(work, "p")], timeout)
        pages = sorted(glob.glob(os.path.join(work, "p-*.png")), key=lambda p: int(p.rsplit("-", 1)[1][:-4]))
        for i, page in enumerate(pages, 1):
            shutil.copy2(page, os.path.join(out_dir, f"slide-{i:02d}.png"))
    return {"backend": "libreoffice", "slides": []}


def contact_sheet(out_dir):
    pngs = sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))
    if not pngs:
        return None
    gap, label_h = 12, 22
    rows = (len(pngs) + THUMB_COLS - 1) // THUMB_COLS
    sheet = Image.new("RGB", (THUMB_COLS * (THUMB[0] + gap) + gap, rows * (THUMB[1] + label_h + gap) + gap),
                      (236, 238, 241))
    draw = ImageDraw.Draw(sheet)
    font = ImageFont.load_default()
    for i, path in enumerate(pngs):
        with Image.open(path) as im:
            thumb = im.convert("RGB").resize(THUMB, Image.LANCZOS)
        x = gap + (i % THUMB_COLS) * (THUMB[0] + gap)
        y = gap + (i // THUMB_COLS) * (THUMB[1] + label_h + gap)
        draw.text((x, y), str(i + 1), fill=(60, 64, 72), font=font)
        sheet.paste(thumb, (x, y + label_h))
        draw.rectangle([x - 1, y + label_h - 1, x + THUMB[0], y + label_h + THUMB[1]], outline=(180, 185, 192))
    path = os.path.join(out_dir, "contact.png")
    sheet.save(path)
    return path


def render(deck, out_dir, backend="auto", width=WIDTH, timeout=300):
    deck = os.path.abspath(deck)
    os.makedirs(out_dir, exist_ok=True)
    for old in glob.glob(os.path.join(out_dir, "slide-*.png")):
        os.remove(old)
    with render_lock():
        if backend == "powerpoint" or (backend == "auto" and powerpoint_available()):
            measure = render_powerpoint(deck, out_dir, width, timeout)
            measure["measured"] = True
        else:
            measure = render_libreoffice(deck, out_dir, width, timeout)
            measure["measured"] = False
    measure["deck"] = deck
    measure["deck_sha256"] = sha256_file(deck)
    with open(os.path.join(out_dir, "measure.json"), "w", encoding="utf-8") as f:
        json.dump(measure, f, indent=2, ensure_ascii=False)
    contact_sheet(out_dir)
    return measure


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("deck")
    ap.add_argument("-o", "--out", required=True, help="folder for slide PNGs, contact.png and measure.json")
    ap.add_argument("--backend", choices=("auto", "powerpoint", "libreoffice"), default="auto")
    ap.add_argument("--width", type=int, default=WIDTH)
    ap.add_argument("--timeout", type=int, default=300)
    args = ap.parse_args(argv)
    try:
        measure = render(args.deck, args.out, args.backend, args.width, args.timeout)
    except (RuntimeError, subprocess.TimeoutExpired) as e:
        print(f"RENDER FAILED: {e}", file=sys.stderr)
        return 2
    n = len(glob.glob(os.path.join(args.out, "slide-*.png")))
    kind = "measured" if measure["measured"] else "NOT measured (no PowerPoint); lint will warn"
    print(f"Rendered {n} slides with {measure['backend']}, text {kind}. "
          f"Look at {os.path.join(args.out, 'contact.png')} and each slide-NN.png.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
