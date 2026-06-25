"""Generate the Jump List action icons for Paste Stuff.

These reproduce the thin, monochrome Windows 11 line style used by the shell's
own menu entries (Pin to taskbar, Close window). Each icon is a single white
glyph from the *Segoe Fluent Icons* font that ships with Windows 11:

    icon-edit.ico      pencil          -> "Edit config"
    icon-reload.ico    refresh arrow    -> "Reload config"
    icon-startup.ico   power button     -> "Enable / Disable run at startup"

Run this once (it needs Pillow, see requirements-dev.txt) to (re)create the
.ico files next to main.py:

    python generate_action_icons.py
"""

import os

from PIL import Image, ImageDraw, ImageFont

OUT_DIR = os.path.dirname(os.path.abspath(__file__))

# Segoe Fluent Icons (Windows 11); falls back to Segoe MDL2 Assets (Windows 10).
_FONTS = [
    r"C:\Windows\Fonts\SegoeIcons.ttf",
    r"C:\Windows\Fonts\segmdl2.ttf",
]
FONT_PATH = next((p for p in _FONTS if os.path.exists(p)), _FONTS[0])

SS = 512                                    # supersample working canvas.
SIZES = [16, 24, 32, 48, 64, 128, 256]      # sizes embedded in each .ico.
WHITE = (255, 255, 255, 255)

# Segoe Fluent Icons / MDL2 glyph code points.
GLYPHS = {
    "icon-edit.ico": "\uE70F",      # Edit (pencil).
    "icon-reload.ico": "\uE72C",    # Refresh (circular arrow).
    "icon-startup.ico": "\uE7E8",   # PowerButton.
}


def _render(char):
    font = ImageFont.truetype(FONT_PATH, int(SS * 0.74))
    img = Image.new("RGBA", (SS, SS), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    left, top, right, bottom = draw.textbbox((0, 0), char, font=font)
    x = (SS - (right - left)) // 2 - left
    y = (SS - (bottom - top)) // 2 - top
    draw.text((x, y), char, font=font, fill=WHITE)
    return img


def main():
    for name, char in GLYPHS.items():
        glyph = _render(char)
        icons = [glyph.resize((s, s), Image.LANCZOS) for s in SIZES]
        path = os.path.join(OUT_DIR, name)
        icons[0].save(path, format="ICO", sizes=[(s, s) for s in SIZES],
                      append_images=icons[1:])
        print(f"Wrote {path}")


if __name__ == "__main__":
    main()
