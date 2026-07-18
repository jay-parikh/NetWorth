"""Rasterize page 1 of each sheet PDF to a cropped PNG.

Auto-crops trailing white space (rows/cols) with a small margin, so a
17-column sheet doesn't ship with a page of blank paper around it.
"""
import sys
from pathlib import Path

import fitz

SRC, DST, ZOOM = Path(sys.argv[1]), Path(sys.argv[2]), 2.2
PREFIX = sys.argv[3] if len(sys.argv) > 3 else ""
DST.mkdir(parents=True, exist_ok=True)


def content_bbox(pix):
    """Bounding box of non-white pixels (threshold near-white)."""
    w, h, n = pix.width, pix.height, pix.n
    s = pix.samples
    xmin, ymin, xmax, ymax = w, h, -1, -1
    step = 2                                     # sample every 2nd pixel
    for y in range(0, h, step):
        row = s[y * w * n:(y + 1) * w * n]
        for x in range(0, w, step):
            px = row[x * n:x * n + 3]
            if min(px) < 245:
                if x < xmin: xmin = x
                if x > xmax: xmax = x
                if y < ymin: ymin = y
                ymax = y
    if xmax < 0:
        return None
    m = 8                                        # margin px
    return (max(0, xmin - m), max(0, ymin - m),
            min(w, xmax + m), min(h, ymax + m))


for pdf in sorted(SRC.glob("*.pdf")):
    doc = fitz.open(pdf)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM))
    bbox = content_bbox(pix)
    if bbox:
        clip = fitz.Rect(bbox[0] / ZOOM, bbox[1] / ZOOM,
                         bbox[2] / ZOOM, bbox[3] / ZOOM)
        pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM), clip=clip)
    out = DST / f"{PREFIX}{pdf.stem}.png"
    pix.save(out)
    print(f"{out.name}: {pix.width}x{pix.height}, {doc.page_count} page(s)")
    doc.close()
