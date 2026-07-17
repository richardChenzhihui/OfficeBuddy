"""Page-level diffing between two renders, with change-region highlighting.

Used to send the verifier ONLY the pages that actually changed, with the
changed region outlined so it can focus.
"""
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

import fitz

from .pdf_to_images import PageImage


@dataclass
class PageDiff:
    page_index: int
    changed: bool
    bbox: Optional[Tuple[int, int, int, int]] = None  # x0, y0, x1, y1 in pixels
    reason: str = ""  # "content" | "added" | "removed"


def _pixels(img: PageImage):
    import numpy as np

    pix = fitz.Pixmap(str(img.path))
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    return arr[:, :, :3]  # drop alpha if present


def diff_pages(before: List[PageImage], after: List[PageImage]) -> List[PageDiff]:
    import numpy as np

    diffs: List[PageDiff] = []
    before_by_index = {img.index: img for img in before}
    after_by_index = {img.index: img for img in after}

    for index in sorted(set(before_by_index) | set(after_by_index)):
        if index not in before_by_index:
            diffs.append(PageDiff(index, True, None, "added"))
            continue
        if index not in after_by_index:
            diffs.append(PageDiff(index, True, None, "removed"))
            continue
        a_img, b_img = before_by_index[index], after_by_index[index]
        a, b = _pixels(a_img), _pixels(b_img)
        if a.shape != b.shape:
            diffs.append(PageDiff(index, True, None, "content"))
            continue
        delta = np.any(np.abs(a.astype(np.int16) - b.astype(np.int16)) > 8, axis=2)
        if not delta.any():
            diffs.append(PageDiff(index, False))
            continue
        rows = np.flatnonzero(delta.any(axis=1))
        cols = np.flatnonzero(delta.any(axis=0))
        pad = 12
        bbox = (
            max(0, int(cols[0]) - pad),
            max(0, int(rows[0]) - pad),
            min(b.shape[1], int(cols[-1]) + pad),
            min(b.shape[0], int(rows[-1]) + pad),
        )
        diffs.append(PageDiff(index, True, bbox, "content"))
    return diffs


def highlight_region(image: PageImage, bbox: Tuple[int, int, int, int], out_path: Path) -> Path:
    """Draw a red rectangle around the changed region; returns annotated PNG path."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return image.path  # Pillow unavailable: fall back to the raw image
    with Image.open(image.path) as img:
        img = img.convert("RGB")
        draw = ImageDraw.Draw(img)
        for offset in range(3):  # 3px border
            draw.rectangle(
                [bbox[0] - offset, bbox[1] - offset, bbox[2] + offset, bbox[3] + offset],
                outline=(255, 0, 0),
            )
        img.save(out_path)
    return out_path
