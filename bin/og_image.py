#!/usr/bin/env python3
"""Draw graphics/og-image.jpg, the 1200x630 picture shown when the site is shared.

Run by hand (`make og-image`) when the title or theme changes; the result is committed.
The theme's fonts (Fraunces, Public Sans) are fetched from Google Fonts into the
gitignored .deck-cache/fonts/ the first time.
"""
import re
import sys
import urllib.request
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).parent))
from build import SITE_TITLE, TAGLINE, SITE_URL, CACHE_DIR  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SIZE = (1200, 630)
FONTS = {
    "Fraunces.ttf": "https://github.com/google/fonts/raw/main/ofl/fraunces/Fraunces%5BSOFT,WONK,opsz,wght%5D.ttf",
    "PublicSans.ttf": "https://github.com/google/fonts/raw/main/ofl/publicsans/PublicSans%5Bwght%5D.ttf",
}


def css_var(css, name):
    return re.search(rf"--{name}:\s*([^;]+);", css).group(1).strip()


def hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def font(name, size, weight, opsz=None):
    path = ROOT / CACHE_DIR / "fonts" / name
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        urllib.request.urlretrieve(FONTS[name], path)
    f = ImageFont.truetype(str(path), size)
    axes = {a["name"]: a["default"] for a in f.get_variation_axes()}
    for key in axes:
        k = key.decode() if isinstance(key, bytes) else key
        if k.lower() == "weight":
            axes[key] = weight
        elif opsz and "optical" in k.lower():
            axes[key] = opsz
    f.set_variation_by_axes(list(axes.values()))
    return f


def main():
    css = (ROOT / "style.css").read_text()
    blue = hex_rgb(css_var(css, "binding"))
    brick = hex_rgb(css_var(css, "gilt"))
    light = hex_rgb(css_var(css, "binding-text"))

    # Fill the frame with the painting, keeping the skyline in view.
    art = Image.open(ROOT / "graphics" / "view-of-delft.jpg").convert("RGB")
    scale = SIZE[0] / art.width
    art = art.resize((SIZE[0], round(art.height * scale)), Image.LANCZOS)
    top = round(art.height * 0.30)
    img = art.crop((0, top, SIZE[0], top + SIZE[1]))

    # Delft-blue wash, dense at the bottom where the text sits.
    wash = Image.new("RGBA", SIZE, blue + (0,))
    d = ImageDraw.Draw(wash)
    for y in range(SIZE[1]):
        t = min(1.0, max(0.0, (y / SIZE[1] - 0.35) / 0.40))
        d.line([(0, y), (SIZE[0], y)], fill=blue + (round(240 * t ** 0.7),))
    img = Image.alpha_composite(img.convert("RGBA"), wash).convert("RGB")

    d = ImageDraw.Draw(img)
    x = 64
    d.rectangle([x, 410, x + 72, 416], fill=brick)
    d.text((x, 432), TAGLINE.upper(), font=font("PublicSans.ttf", 30, 600), fill=brick)
    d.text((x, 470), SITE_TITLE, font=font("Fraunces.ttf", 76, 800, opsz=144), fill=light)
    url = SITE_URL.removeprefix("https://").rstrip("/")
    d.text((x, 572), url, font=font("PublicSans.ttf", 26, 500), fill=light)

    out = ROOT / "graphics" / "og-image.jpg"
    img.save(out, quality=85, optimize=True)
    print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
