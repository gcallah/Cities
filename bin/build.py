#!/usr/bin/env python3
"""Build the "Cities, Culture and Economy" static site from course material kept elsewhere.

Collects:
  * lecture slides saved in slides/<slug>.html -> lectures/<slug>.html
    (with --extract, first re-reads the decks: .pptx directly, .ppt/.key via
    Keynote, cached in .deck-cache/; images go to lectures/<slug>/)
  * handouts (.docx, .rtf, .txt)  -> handouts/<slug>.html (via pandoc)
  * spreadsheets (.xlsx)          -> spreadsheets/
  * quizzes                       -> quizzes/<slug>.html (via pandoc)
  * photographs                   -> gallery/ (downscaled)
and writes index.html, organised around the order of the course.

Usage: python3 bin/build.py [--src DIR] [--out DIR] [--extract]
Needs: python-pptx and Pillow (pip install python-pptx pillow), pandoc, and on macOS
       Keynote (to convert .ppt/.key decks) and sips.
"""

import argparse
import hashlib
import html
import re
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

DEFAULT_SRC = Path.home() / "Documents/CoursesTaught/CitiesCourse"

# ---------------------------------------------------------------- course data
#
# Everything about *what* goes on the site lives in these tables.
#
# Deliberately left out of the source folder:
#   *.key                                Keynote twins of the .ppt decks used below
#   EvolutionOfCities copy.ppt           duplicate of EvolutionOfCities.ppt
#   Syllabus-2013.docx                   admin details (phone, office, policies written by a
#                                        colleague); its course description is quoted on the index
#   GrandCentral.MOV                     94 MB video, too large for the repo
#   Graphics/, IllustratedJacobs.key/Data/   images already embedded in the decks

# Lecture decks: slug -> source path relative to SRC, in course order.
DECKS = {
    "evolution-of-cities": "EvolutionOfCities.ppt",
    "cities-first": "CitiesFirst.ppt",
    "urban-modernism": "UrbanModernism.ppt",
    "use-of-knowledge": "UseOfKnowledge.pptx",
    "jacobs-critique": "JacobsCritique.ppt",
    "illustrated-jacobs": "IllustratedJacobs.ppt",
}

# Handouts: slug -> (source path, title).
HANDOUTS = {
    "i-dreamed-a-dream": ("IDreamedADream.txt", "I Dreamed a Dream (a planner's lament)"),
}

# Spreadsheets: slug -> (source path, title, description). This course has none.
SPREADSHEETS = {}

# Quizzes, in order. Titles come from each document's first line.
QUIZZES = ["Quiz2_2013.docx", "Quiz3_2013.docx"]

# Photographs for the gallery on the index: (source path, caption).
GALLERY_HEADING = "Siena: a city made for walking"
GALLERY = [
    ("Siena/Roofs.jpg", "Roofs of Siena"),
    ("Siena/Tower.jpg", "The Torre del Mangia"),
    ("Siena/Alley.jpg", "An alley"),
    ("Siena/Arches.jpg", "Arches"),
    ("Siena/Cathedral.jpg", "The cathedral"),
    ("Siena/Arches2.jpg", "More arches"),
    ("Siena/Roofs2.jpg", "Roofs and towers"),
]

# The course, in order. The 2013 syllabus gives no weekly schedule, so these
# parts follow the decks. Readings are given only where a deck names its text. Each topic lists its decks and handouts; topics
# with none would still be shown.
TOPICS = [
    {"title": "How cities arose, vanished and returned: from the polis to Engels's Manchester",
     "decks": ["evolution-of-cities"]},
    {"title": "Cities first, rural development later",
     "reading": "Jacobs, The Economy of Cities, ch. 1, “Cities First—Rural Development Later”",
     "decks": ["cities-first"]},
    {"title": "Urban planning modernism: Howard, Le Corbusier and Wright",
     "decks": ["urban-modernism"], "handouts": ["i-dreamed-a-dream"]},
    {"title": "Hayek: knowledge, prices and spontaneous order",
     "reading": "Hayek, “The Use of Knowledge in Society”",
     "decks": ["use-of-knowledge"]},
    {"title": "Jane Jacobs against urban rationalism",
     "decks": ["jacobs-critique"]},
    {"title": "The Death and Life of Great American Cities, chapter by chapter",
     "reading": "Jacobs, The Death and Life of Great American Cities",
     "decks": ["illustrated-jacobs"]},
]

TEXTBOOK = ("Richard LeGates and Frederic Stout (eds.), <cite>The City Reader</cite>, and "
            "Jane Jacobs, <cite>The Death and Life of Great American Cities</cite>")
DESCRIPTION = ("Is there a common set of social institutions or environments that gives rise to "
               "both successful urban economies and flourishing arts and culture? Proposing that "
               "such a set exists, this course attempts to identify it; traces its implications "
               "for cultural and economic development; and explores the interrelations of "
               "capitalism, cities, and culture.")

# ---------------------------------------------------------------- site chrome

SITE_TITLE = "Cities, Culture and Economy"
TAGLINE = "What Is a City?"
SITE_URL = "https://gcallah.github.io/Cities/"  # absolute links for social-media previews
OG_IMAGE = "graphics/og-image.jpg"  # drawn by bin/og_image.py
FONTS_URL = ("https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,800"
             "&family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;1,8..60,400"
             "&family=Public+Sans:wght@500;600&display=swap")
BACKDROP_CREDIT = ("Background: Johannes Vermeer, <cite>View of Delft</cite> (c. 1660–61), "
                   "Mauritshuis, The Hague. Public domain, via "
                   '<a href="https://commons.wikimedia.org/wiki/File:Vermeer-view-of-delft.jpg">'
                   "Wikimedia Commons</a>. Photographs of Siena by Gene Callahan.")
FOOTER = (f"{SITE_TITLE}: lectures, handouts and quizzes from a course on cities, "
          "from the first towns to Jane Jacobs.")
N_CHART_SERIES = 4  # chart series colours are --series-0 .. --series-3 in style.css
MAX_IMAGE_PX = 1280  # longest side of images published on the site
PNG_TO_JPEG_BYTES = 300_000  # opaque PNGs bigger than this are photos: re-encode as JPEG
JPEG_QUALITY = 80
CACHE_DIR = ".deck-cache"  # under the output dir; gitignored
SLIDES_DIR = "slides"      # saved slide text, one file per deck; edit these to fix slides
SLIDES_HEADER = ("<!-- Saved from the lecture deck by `make extract`; `make html` builds the lecture\n"
                 "     page from this file. Edit it to fix slides (each slide's title is its <h2>).\n"
                 "     Running `make extract` again overwrites it. -->\n")

# Keynote opens .ppt and .key files and exports them as .pptx.
TO_PPTX_SCRIPT = """
on run argv
  with timeout of 900 seconds -- opening an old .ppt can take minutes
    tell application "Keynote"
      set d to open (POSIX file (item 1 of argv))
      export d to (POSIX file (item 2 of argv)) as Microsoft PowerPoint
      close d saving no
    end tell
  end timeout
end run
"""

WEB_IMAGE_EXTS = {"png", "jpg", "jpeg", "gif", "svg", "webp"}
DOWNSCALE_EXTS = {".png", ".jpg", ".jpeg"}
CONVERT_WITH_SIPS = {"tif", "tiff", "bmp"}

# Set in main() from style.css: its hash (so browsers refetch it after changes)
# and the slide card colour that transparent photos are flattened onto.
CSS_VERSION = ""
CARD_COLOUR = "#ffffff"

warnings = []


def warn(msg):
    warnings.append(msg)


def esc(text):
    return html.escape(text, quote=True)


def page(title, body, root="", body_class="", hero=None, path=""):
    """Wrap body in the site chrome.

    root:  relative path from this page to the site root ("" or "../").
    hero:  optional (kicker, heading, subtitle) for the page header.
    path:  this page's path from the site root, for its og:url.
    """
    hero_html = ""
    if hero:
        kicker, heading, sub = hero
        hero_html = f"""<header class="hero">
  {f'<p class="kicker">{kicker}</p>' if kicker else ''}
  <h1>{heading}</h1>
  {f'<p class="sub">{sub}</p>' if sub else ''}
</header>"""
    doc_title = SITE_TITLE if title == SITE_TITLE else f"{title} · {SITE_TITLE}"
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{esc(doc_title)}</title>
    <meta name="description" content="{esc(FOOTER)}">
    <meta property="og:type" content="website">
    <meta property="og:site_name" content="{esc(SITE_TITLE)}">
    <meta property="og:title" content="{esc(doc_title)}">
    <meta property="og:description" content="{esc(FOOTER)}">
    <meta property="og:url" content="{SITE_URL}{path}">
    <meta property="og:image" content="{SITE_URL}{OG_IMAGE}">
    <meta property="og:image:width" content="1200">
    <meta property="og:image:height" content="630">
    <meta property="og:image:alt" content="{esc(SITE_TITLE)}: Vermeer's View of Delft">
    <meta name="twitter:card" content="summary_large_image">
    <meta name="twitter:title" content="{esc(doc_title)}">
    <meta name="twitter:description" content="{esc(FOOTER)}">
    <meta name="twitter:image" content="{SITE_URL}{OG_IMAGE}">
    <meta name="twitter:image:alt" content="{esc(SITE_TITLE)}: Vermeer's View of Delft">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet" href="{FONTS_URL}">
    <link rel="stylesheet" href="{root}style.css?v={CSS_VERSION}">
</head>
<body class="{body_class}">
<nav class="site-nav">
  <a class="brand" href="{root}index.html">{SITE_TITLE}</a>
  <span class="links">
    <a href="{root}index.html#schedule">Schedule</a>
    <a href="{root}index.html#lectures">Lectures</a>
    <a href="{root}index.html#handouts">Handouts</a>
    <a href="{root}index.html#gallery">Siena</a>
    <a href="{root}index.html#quizzes">Quizzes</a>
  </span>
</nav>
<div class="sheet">
{hero_html}
<main>
{body}
</main>
</div>
<footer class="site-footer">
  <p>{FOOTER}</p>
  <p class="credit">{BACKDROP_CREDIT}</p>
</footer>
</body>
</html>
"""


def downscale(path):
    """Shrink an image so its longest side is at most MAX_IMAGE_PX.

    PNGs over PNG_TO_JPEG_BYTES are photos pasted into slides (Keynote often
    exports them with transparent margins): they are flattened onto the slide
    card colour from style.css and saved as JPEG. Returns the resulting path.
    """
    with Image.open(path) as im:
        im.load()
    im.thumbnail((MAX_IMAGE_PX, MAX_IMAGE_PX))
    if path.suffix.lower() != ".png":
        im.convert("RGB").save(path, quality=JPEG_QUALITY, optimize=True)
        return path
    im.save(path, optimize=True)
    if path.stat().st_size <= PNG_TO_JPEG_BYTES:
        return path
    flat = Image.new("RGB", im.size, CARD_COLOUR)
    rgba = im.convert("RGBA")
    flat.paste(rgba, mask=rgba.getchannel("A"))
    jpg = path.with_suffix(".jpg")
    flat.save(jpg, quality=JPEG_QUALITY, optimize=True)
    path.unlink()
    return jpg


def css_colour(css, name):
    """Read a colour such as --card from the :root block of style.css."""
    m = re.search(rf"--{name}:\s*(#[0-9a-fA-F]{{6}})", css)
    return m.group(1) if m else "#ffffff"


def is_decoration(text):
    """Text frames holding only punctuation (e.g. a lone "+" ornament)."""
    return not re.search(r"\w", text)


def slugify(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


# ---------------------------------------------------------------- slides

def run_html(run):
    text = esc(run.text).replace("\x0b", "<br>")
    if not text:
        return ""
    font = run.font
    if font.bold:
        text = f"<strong>{text}</strong>"
    if font.italic:
        text = f"<em>{text}</em>"
    addr = run.hyperlink.address if run.hyperlink else None
    if addr:
        text = f'<a href="{esc(addr)}">{text}</a>'
    return text


def text_frame_html(tf):
    """Render a text frame as nested lists keyed on paragraph level."""
    paras = [(p.level, "".join(run_html(r) for r in p.runs)) for p in tf.paragraphs]
    paras = [(lvl, t) for lvl, t in paras if t.strip()]
    if not paras:
        return ""
    out, depth = [], -1
    for lvl, text in paras:
        while depth < lvl:
            out.append("<ul>")
            depth += 1
        while depth > lvl:
            out.append("</li></ul>")
            depth -= 1
        if out[-1] != "<ul>":
            out.append("</li>")
        out.append(f"<li>{text}")
    out.append("</li>")
    while depth >= 0:
        out.append("</ul>")
        depth -= 1
    return "\n".join(out)


def table_html(tbl):
    rows = []
    for r_i, row in enumerate(tbl.rows):
        tag = "th" if r_i == 0 else "td"
        cells = "".join(
            f"<{tag}>{esc(cell.text).replace(chr(10), '<br>')}</{tag}>"
            for cell in row.cells)
        rows.append(f"<tr>{cells}</tr>")
    return "<table>\n" + "\n".join(rows) + "\n</table>"


C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def series_xy(series, index):
    """Return [(x, y)] for a chart series. Category charts use the point index as x."""
    ys = list(series.values)
    xval = series._element.find(f"{C_NS}xVal")  # XY scatter: explicit x values
    if xval is None:
        return [(i, y) for i, y in enumerate(ys)]
    xs = {int(pt.get("idx")): float(pt.find(f"{C_NS}v").text)
          for pt in xval.iter(f"{C_NS}pt")}
    return [(xs.get(i, i), y) for i, y in enumerate(ys)]


def chart_html(chart):
    """Draw a line or XY chart as inline SVG; colours come from style.css."""
    try:
        series = [(s.name, series_xy(s, k)) for k, s in enumerate(chart.plots[0].series)]
    except Exception as e:  # noqa: BLE001
        warn(f"chart could not be read: {e}")
        return ""
    pts = [(x, y) for _, ps in series for x, y in ps if y is not None]
    if not pts:
        return ""
    xlo, xhi = min(p[0] for p in pts), max(p[0] for p in pts)
    ylo, yhi = min([p[1] for p in pts] + [0]), max(p[1] for p in pts)
    xhi = xhi if xhi != xlo else xlo + 1
    yhi = yhi if yhi != ylo else ylo + 1
    W, H, L, R, T, B = 640, 320, 48, 16, 16, 32

    def sx(v):
        return L + (W - L - R) * (v - xlo) / (xhi - xlo)

    def sy(v):
        return T + (H - T - B) * (1 - (v - ylo) / (yhi - ylo))

    parts = [f'<svg class="chart" viewBox="0 0 {W} {H}" role="img" xmlns="http://www.w3.org/2000/svg">']
    for v in (ylo, (ylo + yhi) / 2, yhi):
        parts.append(f'<line x1="{L}" x2="{W - R}" y1="{sy(v):.1f}" y2="{sy(v):.1f}" class="grid"/>'
                     f'<text x="{L - 6}" y="{sy(v) + 4:.1f}" text-anchor="end">{v:,.4g}</text>')
    for v in (xlo, (xlo + xhi) / 2, xhi):
        parts.append(f'<text x="{sx(v):.1f}" y="{H - B + 18}" text-anchor="middle">{v:,.4g}</text>')
    for k, (_, ps) in enumerate(series):
        line = " ".join(f"{sx(x):.1f},{sy(y):.1f}" for x, y in ps if y is not None)
        parts.append(f'<polyline class="series-{k % N_CHART_SERIES}" points="{line}"/>')
    parts.append("</svg>")
    legend = '<p class="legend">' + " ".join(
        f'<span class="series-{k % N_CHART_SERIES}">&#9644; {esc(nm)}</span>'
        for k, (nm, _) in enumerate(series) if nm) + "</p>"
    return "\n".join(parts) + legend


def first_title_like(slide):
    for shape in slide.shapes:
        if shape.has_text_frame and not is_decoration(shape.text_frame.text):
            paras = [p for p in shape.text_frame.paragraphs if p.text.strip()]
            return shape if len(paras) == 1 and len(paras[0].text) <= 90 else None
    return None


def image_row(slide):
    """Return a slide's pictures if there are several and all sit side by side
    (their vertical extents overlap); otherwise []."""
    pics = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if len(pics) < 2:
        return []
    for a in pics:
        for b in pics:
            if a is not b and (a.top >= b.top + b.height or b.top >= a.top + a.height):
                return []
    return pics


def deck_pptx(path, cache):
    """Return a .pptx for a deck, converting .ppt/.key through Keynote (cached)."""
    if path.suffix.lower() == ".pptx":
        return path
    cache.mkdir(parents=True, exist_ok=True)
    out = cache / f"{path.stem}.pptx"
    if not out.exists() or out.stat().st_mtime < path.stat().st_mtime:
        print(f"  converting {path.name} with Keynote...")
        subprocess.run(["osascript", "-", str(path), str(out)], input=TO_PPTX_SCRIPT,
                       text=True, check=True, capture_output=True)
    return out


class Deck:
    def __init__(self, src, slug, outdir):
        self.src, self.slug = src, slug
        self.imgdir = outdir / slug
        self.n_images = 0

    def image_html(self, shape):
        img = shape.image
        ext = img.ext.lower()
        self.n_images += 1
        name = f"img{self.n_images:03d}"
        self.imgdir.mkdir(parents=True, exist_ok=True)
        if ext in CONVERT_WITH_SIPS:
            raw = self.imgdir / f"{name}.{ext}"
            raw.write_bytes(img.blob)
            dest = self.imgdir / f"{name}.png"
            subprocess.run(["sips", "-s", "format", "png", str(raw), "--out", str(dest)],
                           check=True, capture_output=True)
            raw.unlink()
        elif ext in WEB_IMAGE_EXTS:
            dest = self.imgdir / f"{name}.{ext}"
            dest.write_bytes(img.blob)
        else:
            warn(f"{self.src.name}: skipped .{ext} image (not viewable in browsers)")
            return ""
        if dest.suffix.lower() in DOWNSCALE_EXTS:
            dest = downscale(dest)
        alt = esc(getattr(shape, "name", "") or "")
        return f'<img src="{self.slug}/{dest.name}" alt="{alt}" loading="lazy">'

    def shape_html(self, shape, title_id):
        if shape.shape_id == title_id:
            return ""
        if shape.has_text_frame and is_decoration(shape.text_frame.text):
            return ""
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return "\n".join(filter(None, (self.shape_html(s, title_id) for s in shape.shapes)))
        parts = []
        if hasattr(shape, "image"):
            try:
                parts.append(self.image_html(shape))
            except ValueError:
                pass  # placeholder with no picture in it
        if getattr(shape, "has_table", False) and shape.has_table:
            parts.append(table_html(shape.table))
        if getattr(shape, "has_chart", False) and shape.has_chart:
            parts.append(chart_html(shape.chart))
        if shape.has_text_frame:
            parts.append(text_frame_html(shape.text_frame))
        return "\n".join(filter(None, parts))

    def row_html(self, row, slide_width):
        """Pictures placed side by side on the slide, kept in one row at their slide widths."""
        imgs = []
        for shape in sorted(row, key=lambda sh: sh.left):
            tag = self.image_html(shape)
            if tag:
                width = 100 * shape.width / slide_width
                imgs.append(tag.replace("<img ", f'<img style="--slide-width: {width:.0f}%" ', 1))
        return f'<div class="image-row">\n{chr(10).join(imgs)}\n</div>' if imgs else ""

    def build(self):
        """Return (deck title, [(slide num, slide title)], slides html)."""
        prs = Presentation(self.src)
        if self.imgdir.exists():
            shutil.rmtree(self.imgdir)
        deck_title, toc, sections = None, [], []
        for num, slide in enumerate(prs.slides, 1):
            title_shape = slide.shapes.title
            if title_shape is not None and not title_shape.has_text_frame:
                title_shape = None
            title = title_shape.text_frame.text.strip() if title_shape is not None else ""
            if not title:
                # Decks converted from Keynote have no title placeholders: the
                # first text box with a single paragraph serves as the title.
                title_shape = first_title_like(slide)
                title = title_shape.text_frame.text.strip() if title_shape is not None else ""
            title_id = title_shape.shape_id if title_shape is not None and title else None
            row = image_row(slide)
            row_ids = [sh.shape_id for sh in row]
            parts = []
            for shape in slide.shapes:
                if shape.shape_id in row_ids:
                    if shape.shape_id == row_ids[0]:
                        parts.append(self.row_html(row, prs.slide_width))
                    continue
                parts.append(self.shape_html(shape, title_id))
            body = "\n".join(filter(None, parts))
            title = re.sub(r"\s+", " ", title) or f"Slide {num}"
            if deck_title is None:
                deck_title = title
            toc.append((num, title))
            notes = ""
            if slide.has_notes_slide:
                n = slide.notes_slide.notes_text_frame
                if n is not None and n.text.strip():
                    notes = f'<aside class="notes">{text_frame_html(n)}</aside>'
            sections.append(
                f'<section class="slide" id="s{num}">\n'
                f'<h2><a class="num" href="#s{num}">{num}</a> {esc(title)}</h2>\n'
                f'{body}\n{notes}\n</section>')
        return deck_title, toc, "\n".join(sections)


# ---------------------------------------------------------------- documents

def pandoc_html(path, media_dir, link_prefix):
    """Convert a .docx/.rtf to an HTML fragment, extracting images to media_dir."""
    if path.suffix.lower() == ".txt":
        stanzas = re.split(r"\n\s*\n", path.read_text(encoding="utf-8").strip())
        return "\n".join(f'<p class="verse">{"<br>".join(esc(ln) for ln in st.splitlines())}</p>'
                         for st in stanzas)
    if path.suffix.lower() == ".rtf":  # pandoc 2.x cannot read RTF; macOS textutil can
        res = subprocess.run(["textutil", "-convert", "html", "-stdout", str(path)],
                             capture_output=True, text=True, check=True)
        res = subprocess.run(["pandoc", "-f", "html", "-t", "html5"], input=res.stdout,
                             capture_output=True, text=True, check=True)
        # Text pasted from a PDF keeps its line-end hyphens: "sat- isfaction".
        return re.sub(r"(?<=[a-z])- (?=[a-z])", "", res.stdout)
    res = subprocess.run(
        ["pandoc", str(path), "-t", "html5", f"--extract-media={media_dir}"],
        capture_output=True, text=True, check=True)
    return res.stdout.replace(f'src="{media_dir}/', f'src="{link_prefix}/')


def first_line(path):
    res = subprocess.run(["pandoc", str(path), "-t", "plain"],
                         capture_output=True, text=True, check=True)
    return next((ln.strip() for ln in res.stdout.splitlines() if ln.strip()), path.stem)


def fresh_dir(path):
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True)
    return path


# ---------------------------------------------------------------- topic lookups

def topic_of(kind, slug):
    """Return (index, topic) for the schedule topic that lists this item."""
    for i, t in enumerate(TOPICS):
        if slug in t.get(kind, []):
            return i, t
    return None, None


def topic_number(i):
    return f"Part {i + 1}"


def related_links(topic, root, built, exclude=None):
    """Chips linking to a topic's other decks, handouts and spreadsheets."""
    chips = []
    for kind, cls in (("decks", "deck"), ("handouts", "handout"), ("sheets", "spreadsheet")):
        for slug in topic.get(kind, []):
            if slug == exclude or slug not in built[kind]:
                continue
            href, title = built[kind][slug]
            dl = " download" if kind == "sheets" else ""
            chips.append(f'<a class="chip {cls}" href="{root}{href}"{dl}>{esc(title)}</a>')
    return chips


# ---------------------------------------------------------------- builders

def extract_lectures(src, out):
    """Read each deck once: slide text -> slides/<slug>.html, images -> lectures/<slug>/.

    Overwrites the saved slides, including any hand edits made to them.
    """
    slides_dir = out / SLIDES_DIR
    slides_dir.mkdir(parents=True, exist_ok=True)
    outdir = out / "lectures"
    outdir.mkdir(parents=True, exist_ok=True)
    for slug, rel in DECKS.items():
        path = src / rel
        if not path.exists():
            warn(f"missing deck: {path}")
            continue
        deck = Deck(deck_pptx(path, out / CACHE_DIR), slug, outdir)
        title, toc, body = deck.build()
        (slides_dir / f"{slug}.html").write_text(SLIDES_HEADER + body + "\n", encoding="utf-8")
        print(f"  {rel}: {len(toc)} slides, {deck.n_images} images")


SLIDE_HEADING = re.compile(r'<h2><a class="num" href="#s(\d+)">\d+</a> (.*?)</h2>')


def load_lectures(out):
    """Read the saved slides. Returns {slug: (deck title, [(num, title)], body html)};
    the deck title is the first slide's title."""
    built = {}
    for slug in DECKS:
        path = out / SLIDES_DIR / f"{slug}.html"
        if not path.exists():
            warn(f"no saved slides for {slug}: run `make extract`")
            continue
        body = path.read_text(encoding="utf-8")
        body = body[len(SLIDES_HEADER):] if body.startswith(SLIDES_HEADER) else body
        toc = [(int(n), html.unescape(t)) for n, t in SLIDE_HEADING.findall(body)]
        if not toc:
            warn(f"{path}: no slides found")
            continue
        built[slug] = (toc[0][1], toc, body)
        print(f"  {slug}: {len(toc)} slides")
    return built


def write_lectures(out, decks, built):
    outdir = out / "lectures"
    order = list(decks)
    for i, slug in enumerate(order):
        title, toc, body = decks[slug]
        prev = order[i - 1] if i > 0 else None
        nxt = order[i + 1] if i + 1 < len(order) else None
        ti, topic = topic_of("decks", slug)
        chips = related_links(topic, "../", built, exclude=slug) if topic else []
        related = (f'<p class="related"><span>With this lecture:</span> {" ".join(chips)}</p>'
                   if chips else "")
        contents = "\n".join(f'<li><a href="#s{n}">{esc(t)}</a></li>' for n, t in toc)
        content = f"""{pager(prev, nxt, decks)}
{related}
<details class="contents">
<summary>Contents &middot; {len(toc)} slides</summary>
<ol>
{contents}
</ol>
</details>
{body}
{pager(prev, nxt, decks)}"""
        kicker = esc(topic_number(ti)) if topic else "Lecture"
        sub = esc(topic["title"]) if topic else f"{len(toc)} slides"
        (outdir / f"{slug}.html").write_text(
            page(title, content, root="../", path=f"{outdir.name}/{slug}.html", body_class="lecture",
                 hero=(kicker, esc(title), sub)),
            encoding="utf-8")


def pager(prev, nxt, decks):
    def link(slug, cls, arrow):
        if not slug:
            return "<span></span>"
        return (f'<a class="{cls}" href="{slug}.html"><span class="dir">{arrow}</span>'
                f'<span class="t">{esc(decks[slug][0])}</span></a>')
    return (f'<nav class="pager">{link(prev, "prev", "&larr; Previous lecture")}'
            f'{link(nxt, "next", "Next lecture &rarr;")}</nav>')


def build_handouts(src, out, built):
    outdir = fresh_dir(out / "handouts")
    for slug, (rel, title) in HANDOUTS.items():
        path = src / rel
        if not path.exists():
            warn(f"missing handout: {path}")
            continue
        body = pandoc_html(path, outdir / slug, slug)
        ti, topic = topic_of("handouts", slug)
        chips = related_links(topic, "../", built, exclude=slug) if topic else []
        related = (f'<p class="related"><span>Goes with:</span> {" ".join(chips)}</p>'
                   if chips else "")
        (outdir / f"{slug}.html").write_text(
            page(title, f'{related}\n<article class="paper">\n{body}\n</article>', root="../", path=f"{outdir.name}/{slug}.html",
                 body_class="handout",
                 hero=(esc(topic_number(ti)) if topic else "Handout", esc(title),
                       esc(topic["title"]) if topic else "")),
            encoding="utf-8")
    print(f"  {len(list(outdir.glob('*.html')))} handouts")


def copy_spreadsheets(src, out):
    built = {}
    if not SPREADSHEETS:
        return built
    outdir = fresh_dir(out / "spreadsheets")
    for slug, (rel, title, _) in SPREADSHEETS.items():
        path = src / rel
        if not path.exists():
            warn(f"missing spreadsheet: {path}")
            continue
        shutil.copy2(path, outdir / path.name)
        built[slug] = (f"spreadsheets/{path.name}", title)
    print(f"  {len(built)} spreadsheets")
    return built


def build_quizzes(src, out):
    outdir = fresh_dir(out / "quizzes")
    entries = []
    for fname in QUIZZES:
        path = src / fname
        if not path.exists():
            warn(f"missing quiz: {path}")
            continue
        slug = slugify(path.stem)
        # Headings read "History of Economic Thought: Quiz 2"; keep the part after the colon.
        title = first_line(path).split(":", 1)[-1].strip()
        body = pandoc_html(path, outdir / slug, slug)
        (outdir / f"{slug}.html").write_text(
            page(title, f'<article class="paper">\n{body}\n</article>', root="../", path=f"{outdir.name}/{slug}.html",
                 body_class="quiz", hero=("Quizzes and exams", esc(title), "")),
            encoding="utf-8")
        entries.append((f"quizzes/{slug}.html", title))
    print(f"  {len(entries)} quizzes")
    return entries


# ---------------------------------------------------------------- index

def build_index(out, decks, built, quizzes, gallery):
    rows = []
    for i, t in enumerate(TOPICS):
        chips = related_links(t, "", built)
        reading = f'<span class="reading">Reading: {esc(t["reading"])}</span>' if t.get("reading") else ""
        rows.append(
            f'<li class="topic{"" if chips else " bare"}">'
            f'<span class="week">{i + 1}</span>'
            f'<div><span class="title">{esc(t["title"])}</span>{reading}'
            + (f'<span class="chips">{"".join(chips)}</span>' if chips else "")
            + "</div></li>")

    lec = "\n".join(
        f'<li><a class="card" href="lectures/{slug}.html">'
        f'<span class="label">{esc(topic_number(topic_of("decks", slug)[0]))}</span>'
        f'<span class="title">{esc(title)}</span>'
        f'<span class="meta">{len(toc)} slides</span></a></li>'
        for slug, (title, toc, _) in decks.items())

    hand = "\n".join(
        f'<li><a href="{built["handouts"][slug][0]}">{esc(built["handouts"][slug][1])}</a></li>'
        for slug in HANDOUTS if slug in built["handouts"])
    sheets = "\n".join(
        f'<li><a href="{built["sheets"][slug][0]}" download>{esc(built["sheets"][slug][1])}</a>'
        f' <span class="desc">{esc(SPREADSHEETS[slug][2])}</span></li>'
        for slug in SPREADSHEETS if slug in built["sheets"])
    qz = "".join(f'<a class="chip" href="{h}">{esc(t)}</a>' for h, t in quizzes)

    sheets_html = (f"""<h3>Spreadsheets</h3>
<ul class="plain-list">
{sheets}
</ul>""" if sheets else "")
    heading = "Handouts and spreadsheets" if sheets else "Handouts"
    photos = "\n".join(
        f'<figure><a href="{href}"><img src="{href}" alt="{esc(cap)}" loading="lazy"></a>'
        f"<figcaption>{esc(cap)}</figcaption></figure>"
        for href, cap in gallery)

    body = f"""<section id="schedule">
<h2>The course</h2>
<p class="lede">{esc(DESCRIPTION)} Books: {TEXTBOOK}.</p>
<ol class="schedule">
{chr(10).join(rows)}
</ol>
</section>
<section id="lectures">
<h2>Lectures</h2>
<ul class="cards">
{lec}
</ul>
</section>
<section id="handouts">
<h2>{heading}</h2>
<ul class="plain-list">
{hand}
</ul>
{sheets_html}
</section>
<section id="quizzes">
<h2>Quizzes</h2>
<p class="chips">{qz}</p>
</section>
<section id="gallery">
<h2>{esc(GALLERY_HEADING)}</h2>
<div class="gallery">
{photos}
</div>
</section>"""
    (out / "index.html").write_text(
        page(SITE_TITLE, body, body_class="home",
             hero=(esc(TAGLINE), SITE_TITLE,
                   "How cities arise, why they flourish, and what happens when planners try "
                   "to design them from above: from the polis to Le Corbusier, Hayek and "
                   "Jane Jacobs.")),
        encoding="utf-8")


def copy_gallery(src, out):
    """Copy the gallery photographs, downscaled for the web."""
    outdir = fresh_dir(out / "gallery")
    entries = []
    for rel, caption in GALLERY:
        path = src / rel
        if not path.exists():
            warn(f"missing photo: {path}")
            continue
        dest = outdir / f"{slugify(path.stem)}{path.suffix.lower()}"
        shutil.copy2(path, dest)
        dest = downscale(dest)
        entries.append((f"gallery/{dest.name}", caption))
    print(f"  {len(entries)} photos")
    return entries


def main():
    global CSS_VERSION, CARD_COLOUR
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--src", type=Path, default=DEFAULT_SRC)
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parent.parent)
    ap.add_argument("--extract", action="store_true",
                    help="re-read the lecture decks into slides/ and lectures/<slug>/ "
                         "(overwrites hand edits to the saved slides)")
    args = ap.parse_args()
    if not args.src.is_dir():
        sys.exit(f"source directory not found: {args.src}")
    css = (args.out / "style.css").read_text(encoding="utf-8")
    CSS_VERSION = hashlib.sha1(css.encode()).hexdigest()[:8]
    CARD_COLOUR = css_colour(css, "card")

    if args.extract:
        print("Extracting decks:")
        extract_lectures(args.src, args.out)
    print("Lectures:")
    decks = load_lectures(args.out)
    print("Spreadsheets:")
    built = {
        "decks": {s: (f"lectures/{s}.html", d[0]) for s, d in decks.items()},
        "handouts": {s: (f"handouts/{s}.html", t) for s, (r, t) in HANDOUTS.items()
                     if (args.src / r).exists()},
        "sheets": copy_spreadsheets(args.src, args.out),
    }
    write_lectures(args.out, decks, built)
    print("Handouts:")
    build_handouts(args.src, args.out, built)
    print("Quizzes:")
    quizzes = build_quizzes(args.src, args.out)
    print("Gallery:")
    gallery = copy_gallery(args.src, args.out)
    build_index(args.out, decks, built, quizzes, gallery)
    for w in warnings:
        print("warning:", w, file=sys.stderr)


if __name__ == "__main__":
    main()
