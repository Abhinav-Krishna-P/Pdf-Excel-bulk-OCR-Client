"""Page rasterisation and text/graphic extraction.

Every page is reduced to the same primitives regardless of whether it has a
text layer (born-digital PDF) or is a scan (OCR via Tesseract):

* ``Word``    - a single token with a glyph-tight pixel box, an estimated em
                size (nominal points) and a baseline
* ``Rule``    - a horizontal or vertical line (table separators)
* ``Picture`` - a non-text raster region (logos), cropped as PNG bytes

Pages are *normalised*: each is rendered to a raster whose width corresponds
to an A4 sheet at ``NOMINAL_PPI`` (or the native scan resolution if higher),
independent of the PDF's own page box. Scans are often stored on tiny or
odd-sized pages, and this keeps OCR quality and the Excel geometry stable.

* ``PageData.scale``  pixels per PDF point  (px -> pt for the PDF overlay)
* ``PageData.ppi``    pixels per nominal inch (px -> Excel points / screen px)
"""
from __future__ import annotations

import io
import os
import re
import shutil
from dataclasses import dataclass, field

import cv2
import fitz  # PyMuPDF
import numpy as np
import pytesseract
from PIL import Image

NOMINAL_PPI = 200
A4_WIDTH_IN = 8.27
MIN_RENDER_W = int(A4_WIDTH_IN * NOMINAL_PPI)  # 1654 px
MAX_RENDER_W = 2600
MIN_OCR_CONF = 30
_AMOUNT_RE = re.compile(r"^\d{1,3}(,\d{2,3})*\.\d{2}(Cr|Dr|CR|DR)?$")

_TESS_CANDIDATES = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
]


def configure_tesseract() -> bool:
    """Point pytesseract at a Tesseract binary. Returns True if one was found."""
    if shutil.which("tesseract"):
        return True
    for cand in _TESS_CANDIDATES:
        if os.path.isfile(cand):
            pytesseract.pytesseract.tesseract_cmd = cand
            return True
    return False


@dataclass
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size_pt: float = 0.0  # em size in nominal points (Arial/Helvetica metrics)
    baseline: float = 0.0  # px
    bold: bool = False

    @property
    def xc(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def yc(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def h(self) -> float:
        return self.y1 - self.y0


@dataclass
class Rule:
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def horizontal(self) -> bool:
        return (self.x1 - self.x0) >= (self.y1 - self.y0)


@dataclass
class Picture:
    x0: float
    y0: float
    x1: float
    y1: float
    png: bytes


@dataclass
class PageData:
    index: int
    width: float  # px
    height: float
    scale: float  # px per PDF point
    ppi: float  # px per nominal inch
    words: list[Word] = field(default_factory=list)
    rules: list[Rule] = field(default_factory=list)
    pictures: list[Picture] = field(default_factory=list)
    source: str = "text"  # "text" | "ocr"


# --------------------------------------------------------------------------- #
# Rasterisation & masks
# --------------------------------------------------------------------------- #
def page_scale(page: fitz.Page) -> float:
    """Pixels per point so that the page renders at a useful width."""
    native = 0
    for info in page.get_image_info():
        bw = info["bbox"][2] - info["bbox"][0]
        if bw > 0.5 * page.rect.width:  # a full-page scan
            native = max(native, int(info["width"] * page.rect.width / bw))
    # 2x the native scan resolution gives Tesseract the best results
    target = min(MAX_RENDER_W, max(MIN_RENDER_W, 2 * native))
    return target / page.rect.width


def render_page(page: fitz.Page, scale: float) -> np.ndarray:
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = img[:, :, :3]
    return np.ascontiguousarray(img)  # RGB


def ink_mask(gray: np.ndarray) -> np.ndarray:
    _, mask = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
    return mask


def line_masks(mask: np.ndarray, ppi: float, min_len_in: float = 0.4) -> tuple[np.ndarray, np.ndarray]:
    """(horizontal, vertical) masks of long straight strokes."""
    L = max(8, int(min_len_in * ppi))
    h = cv2.morphologyEx(mask, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (L, 1)))
    v = cv2.morphologyEx(mask, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, L)))
    return h, v


# --------------------------------------------------------------------------- #
# Glyph metrics
# --------------------------------------------------------------------------- #
_ASCENDING = set("bdfhklt/()[]{}|!?\"'$%&@#0123456789")
_DESCENDING = set("gjpqy,;()[]{}@")
CAP = 0.716  # cap-height / em  (Arial & Helvetica)
XH = 0.519  # x-height / em
DESC = 0.21  # descender / em


def text_extents(text: str) -> tuple[float, float] | None:
    """(above_baseline, below_baseline) in em for a glyph-tight box, or None
    if the text carries no reliable height information (e.g. '-')."""
    up = None
    if any(c.isupper() for c in text) or _ASCENDING & set(text):
        up = CAP
    elif any(c.islower() for c in text):
        up = XH
    if up is None:
        return None
    down = DESC if _DESCENDING & set(text) else 0.0
    return up, down


def refine_and_measure(words: list[Word], mask: np.ndarray, ppi: float) -> list[Word]:
    """Tighten each box to its ink and derive em size + baseline from it."""
    H, W = mask.shape
    out: list[Word] = []
    for w in words:
        x0, y0 = max(0, int(w.x0) - 1), max(0, int(w.y0) - 1)
        x1, y1 = min(W, int(w.x1) + 2), min(H, int(w.y1) + 2)
        crop = mask[y0:y1, x0:x1]
        rows = np.where(crop.any(axis=1))[0]
        cols = np.where(crop.any(axis=0))[0]
        if len(rows) and len(cols):
            w.y0, w.y1 = y0 + rows[0], y0 + rows[-1] + 1
            w.x0, w.x1 = x0 + cols[0], x0 + cols[-1] + 1
        ext = text_extents(w.text)
        if ext is None:
            w.size_pt = 0.0  # filled from neighbours later
            w.baseline = w.y1
        else:
            up, down = ext
            em_px = w.h / (up + down)
            w.size_pt = em_px * 72.0 / ppi
            w.baseline = w.y1 - down * em_px
        out.append(w)
    # words without height info inherit the median size of the page
    known = [w.size_pt for w in out if w.size_pt > 0]
    fallback = float(np.median(known)) if known else 10.0
    for w in out:
        if w.size_pt <= 0:
            w.size_pt = fallback
    return out


# --------------------------------------------------------------------------- #
# Text layer
# --------------------------------------------------------------------------- #
def words_from_text_layer(page: fitz.Page, scale: float) -> list[Word]:
    """Words with positions from the PDF's own text layer (loose boxes)."""
    bold_rects: list[fitz.Rect] = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            for sp in line["spans"]:
                if bool(sp["flags"] & 16) or "bold" in sp["font"].lower():
                    bold_rects.append(fitz.Rect(sp["bbox"]))
    words: list[Word] = []
    for x0, y0, x1, y1, text, *_ in page.get_text("words"):
        text = text.strip()
        if not text:
            continue
        r = fitz.Rect(x0, y0, x1, y1)
        bold = any(br.intersects(r) and (br & r).get_area() > 0.5 * r.get_area() for br in bold_rects)
        words.append(Word(text, x0 * scale, y0 * scale, x1 * scale, y1 * scale, bold=bold))
    return words


# --------------------------------------------------------------------------- #
# OCR
# --------------------------------------------------------------------------- #
def words_from_ocr(gray: np.ndarray, lines: np.ndarray) -> list[Word]:
    clean = gray.copy()
    clean[cv2.dilate(lines, np.ones((3, 3), np.uint8)) > 0] = 255  # erase table lines
    data = pytesseract.image_to_data(
        clean, config="--psm 11 --oem 3", output_type=pytesseract.Output.DICT
    )
    words: list[Word] = []
    for i, text in enumerate(data["text"]):
        text = (text or "").strip()
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
        except (TypeError, ValueError):
            conf = -1
        if conf < MIN_OCR_CONF and not (conf >= 5 and _AMOUNT_RE.match(text)):
            continue  # low confidence, unless it is clearly a well-formed amount
        if not any(c.isalnum() for c in text):
            # bare punctuation is usually a line/noise artefact unless it is a
            # confident, legitimate symbol such as ':' or '-'
            if conf < 60 or any(c in text for c in "|_~\u2014\u2013[]{}"):
                continue
        x, y, w, h = (int(data[k][i]) for k in ("left", "top", "width", "height"))
        if w < 2 or h < 4:
            continue
        words.append(Word(text, x, y, x + w, y + h))
    return words


# --------------------------------------------------------------------------- #
# Graphics: rules and pictures (logos)
# --------------------------------------------------------------------------- #
def find_rules(hmask: np.ndarray, vmask: np.ndarray, ppi: float) -> list[Rule]:
    rules = []
    thick = 0.04 * ppi
    for m, horizontal in ((hmask, True), (vmask, False)):
        n, _, stats, _ = cv2.connectedComponentsWithStats(m, connectivity=8)
        for i in range(1, n):
            x, y, w, h, _ = stats[i]
            long_, short = (w, h) if horizontal else (h, w)
            if long_ >= 0.4 * ppi and short <= thick:
                rules.append(Rule(float(x), float(y), float(x + w), float(y + h)))
    return rules


def find_pictures(
    img_rgb: np.ndarray, mask: np.ndarray, lines: np.ndarray, words: list[Word], ppi: float
) -> list[Picture]:
    m = mask.copy()
    m[cv2.dilate(lines, np.ones((5, 5), np.uint8)) > 0] = 0
    pad = int(0.03 * ppi)
    for w in words:
        cv2.rectangle(m, (int(w.x0) - pad, int(w.y0) - pad), (int(w.x1) + pad, int(w.y1) + pad), 0, -1)
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (int(0.08 * ppi), int(0.08 * ppi)))
    closed = cv2.morphologyEx(m, cv2.MORPH_CLOSE, k)
    n, labels, stats, _ = cv2.connectedComponentsWithStats(closed, connectivity=8)
    pics = []
    min_dim = 0.12 * ppi  # fragments; the final size filter runs after merging
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if w < min_dim or h < min_dim or area < (0.15 * ppi) ** 2:
            continue
        density = (m[y : y + h, x : x + w] > 0).mean()
        if density < 0.06:  # stray strokes / noise, not an image
            continue
        mg = int(0.02 * ppi)
        x0, y0 = max(0, x - mg), max(0, y - mg)
        x1, y1 = min(img_rgb.shape[1], x + w + mg), min(img_rgb.shape[0], y + h + mg)
        crop = Image.fromarray(img_rgb[y0:y1, x0:x1])
        buf = io.BytesIO()
        crop.save(buf, format="PNG")
        pics.append(Picture(float(x0), float(y0), float(x1), float(y1), buf.getvalue()))
    return pics


def _near(a: tuple[float, float, float, float], b: tuple[float, float, float, float], gap: float) -> bool:
    return not (a[2] + gap < b[0] or b[2] + gap < a[0] or a[3] + gap < b[1] or b[3] + gap < a[1])


def absorb_logo_text(
    img_rgb: np.ndarray, pictures: list[Picture], words: list[Word], ppi: float
) -> tuple[list[Picture], list[Word]]:
    """Merge nearby picture fragments, and fold oversized words that sit next
    to a picture (stylised logo lettering) into it, so a logo is one image."""
    if not pictures:
        return pictures, words
    sizes = [w.size_pt for w in words]
    med = float(np.median(sizes)) if sizes else 10.0
    gap = 0.12 * ppi
    boxes = [[p.x0, p.y0, p.x1, p.y1] for p in pictures]
    absorbed: set[int] = set()
    changed = True
    while changed:
        changed = False
        # merge picture boxes that are close to each other
        i = 0
        while i < len(boxes):
            j = i + 1
            while j < len(boxes):
                if _near(boxes[i], boxes[j], gap):
                    a, b = boxes[i], boxes.pop(j)
                    boxes[i] = [min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])]
                    changed = True
                else:
                    j += 1
            i += 1
        # absorb big words adjacent to a picture
        for k, w in enumerate(words):
            if k in absorbed or w.size_pt < 1.8 * med:
                continue
            wb = (w.x0, w.y0, w.x1, w.y1)
            for box in boxes:
                if _near(box, wb, gap):
                    box[0], box[1] = min(box[0], w.x0), min(box[1], w.y0)
                    box[2], box[3] = max(box[2], w.x1), max(box[3], w.y1)
                    absorbed.add(k)
                    changed = True
                    break
    mg = int(0.02 * ppi)
    out: list[Picture] = []
    for x0, y0, x1, y1 in boxes:
        # anything still small after merging is unrecognised text, not artwork
        if min(x1 - x0, y1 - y0) < 0.25 * ppi or (x1 - x0) * (y1 - y0) < (0.35 * ppi) ** 2:
            continue
        x0, y0 = max(0, int(x0) - mg), max(0, int(y0) - mg)
        x1, y1 = min(img_rgb.shape[1], int(x1) + mg), min(img_rgb.shape[0], int(y1) + mg)
        buf = io.BytesIO()
        Image.fromarray(img_rgb[y0:y1, x0:x1]).save(buf, format="PNG")
        out.append(Picture(float(x0), float(y0), float(x1), float(y1), buf.getvalue()))
    kept = [w for k, w in enumerate(words) if k not in absorbed]
    return out, kept


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def extract_pages(pdf_bytes: bytes) -> list[PageData]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages: list[PageData] = []
    tess_ok = configure_tesseract()
    for idx, page in enumerate(doc):
        scale = page_scale(page)
        img = render_page(page, scale)
        ppi = img.shape[1] / A4_WIDTH_IN
        gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        mask = ink_mask(gray)
        hmask, vmask = line_masks(mask, ppi)
        lines = hmask | vmask

        words = words_from_text_layer(page, scale)
        source = "text"
        if len(words) < 3:
            if not tess_ok:
                raise RuntimeError(
                    "This PDF has no text layer and Tesseract OCR is not installed. "
                    "Install it from https://github.com/UB-Mannheim/tesseract/wiki"
                )
            words = words_from_ocr(gray, lines)
            source = "ocr"
        text_only = mask.copy()
        text_only[lines > 0] = 0
        words = refine_and_measure(words, text_only, ppi)
        rules = find_rules(hmask, vmask, ppi)
        pics = find_pictures(img, mask, lines, words, ppi)
        pics, words = absorb_logo_text(img, pics, words, ppi)
        pages.append(PageData(idx, img.shape[1], img.shape[0], scale, ppi, words, rules, pics, source))
    doc.close()
    return pages
