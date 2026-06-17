#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Engangsscript: genererer icon.ico - hvidt dokument-symbol med cyan flueben
på mørkeblå baggrund. Ikke Rambølls logo, bare et eget simpelt symbol."""

from PIL import Image, ImageDraw

OXFORD = (45, 55, 72)       # #2D3748
CYAN = (0, 157, 224)        # #009DE0
WHITE = (255, 255, 255)


def teg(size):
    img = Image.new("RGBA", (size, size), OXFORD + (255,))
    d = ImageDraw.Draw(img)

    # Rundede hjørner på baggrunden
    r = max(2, size // 8)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=255)
    bg = Image.new("RGBA", (size, size), OXFORD + (255,))
    img = Image.composite(bg, Image.new("RGBA", (size, size), (0, 0, 0, 0)), mask)
    d = ImageDraw.Draw(img)

    # Hvidt dokument (et omvendt "blad" med foldet hjørne)
    dx0, dy0 = size * 0.26, size * 0.16
    dx1, dy1 = size * 0.70, size * 0.86
    fold = size * 0.14
    d.polygon([
        (dx0, dy0), (dx1 - fold, dy0), (dx1, dy0 + fold),
        (dx1, dy1), (dx0, dy1),
    ], fill=WHITE + (255,))
    d.polygon([(dx1 - fold, dy0), (dx1, dy0 + fold), (dx1 - fold, dy0 + fold)],
              fill=(214, 224, 230, 255))

    # Et par "tekstlinjer" på dokumentet
    for i, frac in enumerate((0.36, 0.46, 0.56)):
        ly = dy0 + (dy1 - dy0) * frac
        d.line([(dx0 + size * 0.06, ly), (dx1 - size * 0.10, ly)],
               fill=(170, 185, 195, 255), width=max(1, size // 40))

    # Cyan flueben-cirkel i nederste højre hjørne, der overlapper dokumentet
    cx, cy, cr = size * 0.66, size * 0.70, size * 0.26
    d.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=CYAN + (255,))
    lw = max(2, size // 14)
    d.line([(cx - cr * 0.45, cy), (cx - cr * 0.08, cy + cr * 0.38)], fill=WHITE + (255,), width=lw)
    d.line([(cx - cr * 0.08, cy + cr * 0.38), (cx + cr * 0.5, cy - cr * 0.35)], fill=WHITE + (255,), width=lw)

    return img


storrelser = [16, 24, 32, 48, 64, 128, 256]
billeder = [teg(s) for s in storrelser]
billeder[-1].save("icon.ico", sizes=[(s, s) for s in storrelser])
print("icon.ico gemt.")
