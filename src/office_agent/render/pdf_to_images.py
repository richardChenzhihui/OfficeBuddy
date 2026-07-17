"""PDF page rasterization via PyMuPDF."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

import fitz  # PyMuPDF

DEFAULT_ZOOM = 2.0  # 144 dpi ≈ 1200-1700 px long side: good for LLM vision


@dataclass
class PageImage:
    index: int  # 0-based page index
    path: Path
    width: int
    height: int
    _bytes: Optional[bytes] = field(default=None, repr=False)

    @property
    def png_bytes(self) -> bytes:
        if self._bytes is None:
            self._bytes = self.path.read_bytes()
        return self._bytes


def pdf_to_images(
    pdf_path: Path,
    out_dir: Path,
    zoom: float = DEFAULT_ZOOM,
    pages: Optional[List[int]] = None,
) -> List[PageImage]:
    out_dir.mkdir(parents=True, exist_ok=True)
    images: List[PageImage] = []
    with fitz.open(str(pdf_path)) as doc:
        indices = pages if pages is not None else range(len(doc))
        for i in indices:
            if not 0 <= i < len(doc):
                continue
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            path = out_dir / f"page_{i:03d}.png"
            pix.save(str(path))
            images.append(PageImage(index=i, path=path, width=pix.width, height=pix.height))
    return images


def locate_text_pages(pdf_path: Path, snippet: str, max_chars: int = 60) -> List[int]:
    """Find 0-based page indices containing the text snippet.

    Long snippets are trimmed: PDF extraction can differ in whitespace, so a
    shorter distinctive prefix matches more reliably.
    """
    needle = " ".join(snippet.split())[:max_chars]
    if not needle:
        return []
    hits: List[int] = []
    with fitz.open(str(pdf_path)) as doc:
        for i, page in enumerate(doc):
            if page.search_for(needle):
                hits.append(i)
    return hits
