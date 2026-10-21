"""Turn positioned words into a spreadsheet grid that mirrors the page.

Pipeline:  words -> segments (one per visual text run) -> rows (y-bands)
           -> columns (x boundaries) -> cells (segment -> row/col span)

The grid keeps the *geometry* of the page: column widths and row heights are
proportional to the pixel distances on the page, so the sheet lines up with
the statement when viewed in Excel.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from statistics import median

from .extract import PageData, Word

_NUMERIC_RE = re.compile(r"^[\(\-₹$]?[\d,]*\.?\d+%?\)?(Cr|Dr|CR|DR)?$")


@dataclass
class Segment:
    id: int
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size_pt: float
    baseline: float
    bold: bool
    align: str = "left"  # "left" | "right"
    # free horizontal space the text may grow into without hitting a neighbour
    free_x0: float = 0.0
    free_x1: float = 0.0
    # grid placement (0-based)
    row: int = -1
    col0: int = -1
    col1: int = -1

    @property
    def yc(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def h(self) -> float:
        return self.y1 - self.y0


@dataclass
class Grid:
    page: PageData
    segments: list[Segment]
    row_bounds: list[float]  # len = nrows + 1, px
    col_bounds: list[float]  # len = ncols + 1, px
    text_rows: set[int] = field(default_factory=set)

    @property
    def nrows(self) -> int:
        return len(self.row_bounds) - 1

    @property
    def ncols(self) -> int:
        return len(self.col_bounds) - 1

    def row_at(self, y: float) -> int:
        return _index_of(self.row_bounds, y)

    def col_at(self, x: float) -> int:
        return _index_of(self.col_bounds, x)


def _index_of(bounds: list[float], v: float) -> int:
    lo, hi = 0, len(bounds) - 2
    while lo < hi:
        mid = (lo + hi + 1) // 2
        if bounds[mid] <= v:
            lo = mid
        else:
            hi = mid - 1
    return lo


# --------------------------------------------------------------------------- #
# 1. Words -> segments
# --------------------------------------------------------------------------- #
def _cluster_1d(values: list[float], tol: float) -> list[list[int]]:
    """Greedy clustering of sorted scalar values; returns index groups."""
    order = sorted(range(len(values)), key=lambda i: values[i])
    groups: list[list[int]] = []
    cur: list[int] = []
    cur_mean = 0.0
    for i in order:
        if cur and abs(values[i] - cur_mean) > tol:
            groups.append(cur)
            cur = []
        cur.append(i)
        cur_mean = sum(values[j] for j in cur) / len(cur)
    if cur:
        groups.append(cur)
    return groups


def build_segments(words: list[Word], ppi: float) -> list[Segment]:
    if not words:
        return []
    med_h = median(w.h for w in words)
    lines = _cluster_1d([w.yc for w in words], tol=max(0.55 * med_h, 3))
    segments: list[Segment] = []
    for line in lines:
        lw = sorted((words[i] for i in line), key=lambda w: w.x0)
        # typical character width on this line (robust to narrow single-letter words)
        cw = max(1.0, median((w.x1 - w.x0) / max(1, len(w.text)) for w in lw))
        run: list[Word] = [lw[0]]
        for w in lw[1:]:
            prev = run[-1]
            # split when the gap is wider than ~2 characters (a column gap),
            # or when the font size changes markedly (label vs value).
            size_jump = abs(w.size_pt - prev.size_pt) > 0.35 * max(w.size_pt, prev.size_pt)
            if w.x0 - prev.x1 > max(2.2 * cw, 0.04 * ppi) or size_jump:
                segments.append(_make_segment(run))
                run = []
            run.append(w)
        segments.append(_make_segment(run))
    segments.sort(key=lambda s: (s.y0, s.x0))
    for i, s in enumerate(segments):
        s.id = i
    return segments


def _make_segment(run: list[Word]) -> Segment:
    text = " ".join(w.text for w in run)
    return Segment(
        id=-1,
        text=text,
        x0=min(w.x0 for w in run),
        y0=min(w.y0 for w in run),
        x1=max(w.x1 for w in run),
        y1=max(w.y1 for w in run),
        size_pt=round(median(w.size_pt for w in run), 2),
        baseline=round(median(w.baseline for w in run), 2),
        bold=any(w.bold for w in run),
    )


# --------------------------------------------------------------------------- #
# 2. Rows
# --------------------------------------------------------------------------- #
def _row_bands(segments: list[Segment], page_h: float, ppi: float) -> tuple[list[float], list[list[Segment]], set[int]]:
    """Return row boundaries, the segments in each *text* row, and the set of
    row indices that carry text (the rest are spacer rows)."""
    med_h = median(s.h for s in segments)
    groups = _cluster_1d([s.yc for s in segments], tol=max(0.55 * med_h, 3))
    groups.sort(key=lambda g: sum(segments[i].yc for i in g) / len(g))
    bands: list[tuple[float, float, list[Segment]]] = []
    for g in groups:
        segs = [segments[i] for i in g]
        y0 = min(s.y0 for s in segs)
        y1 = max(s.y1 for s in segs)
        bands.append((y0, y1, segs))

    bounds: list[float] = [0.0]
    row_segs: list[list[Segment]] = []
    text_rows: set[int] = set()
    prev_bottom = 0.0
    for i, (y0, y1, segs) in enumerate(bands):
        h = y1 - y0
        pad = 0.45 * h
        next_top = bands[i + 1][0] if i + 1 < len(bands) else page_h
        top = max(prev_bottom, y0 - pad)
        bottom = min(y1 + pad, (y1 + next_top) / 2 if i + 1 < len(bands) else page_h)
        bottom = max(bottom, top + max(h, 1.0))  # bands overlapping in y must still advance
        if top - prev_bottom > 0.035 * ppi:  # spacer row for the gap above
            bounds.append(top)
            row_segs.append([])
        text_rows.add(len(bounds) - 1)
        bounds.append(bottom)
        row_segs.append(segs)
        prev_bottom = bottom
    if page_h - prev_bottom > 1:
        bounds.append(page_h)
        row_segs.append([])
    return bounds, row_segs, text_rows


# --------------------------------------------------------------------------- #
# 3. Columns
# --------------------------------------------------------------------------- #
def _choose_alignment(segments: list[Segment], tol: float) -> None:
    lefts = [s.x0 for s in segments]
    rights = [s.x1 for s in segments]

    def support(vals: list[float], v: float) -> int:
        return sum(1 for u in vals if abs(u - v) <= tol)

    for s in segments:
        ls, rs = support(lefts, s.x0), support(rights, s.x1)
        numeric = bool(_NUMERIC_RE.match(s.text.replace(" ", "")))
        if rs > ls or (numeric and rs >= 2):
            s.align = "right"
        else:
            s.align = "left"


def _col_bounds(segments: list[Segment], row_segs: list[list[Segment]], page_w: float, tol: float) -> list[float]:
    anchors = [s.x0 if s.align == "left" else s.x1 for s in segments]
    groups = _cluster_1d(anchors, tol)
    bounds = sorted({0.0, page_w} | {sum(anchors[i] for i in g) / len(g) for g in groups})

    # Resolve segments that would collide inside one column by adding bounds.
    for _ in range(10):
        extra: set[float] = set()
        for segs in row_segs:
            segs = sorted(segs, key=lambda s: s.x0)
            for a, b in zip(segs, segs[1:]):
                if _index_of(bounds, a.x1 - 0.01) >= _index_of(bounds, b.x0):
                    extra.add(round((a.x1 + b.x0) / 2, 1))
        if not extra:
            break
        bounds = sorted(set(bounds) | extra)
    return bounds


# --------------------------------------------------------------------------- #
# 4. Assemble
# --------------------------------------------------------------------------- #
def build_grid(page: PageData) -> Grid:
    segments = build_segments(page.words, page.ppi)
    if not segments:
        return Grid(page, [], [0.0, page.height], [0.0, page.width])

    tol = 0.008 * page.width
    row_bounds, row_segs, text_rows = _row_bands(segments, page.height, page.ppi)
    _choose_alignment(segments, tol)
    col_bounds = _col_bounds(segments, row_segs, page.width, tol)
    grid = Grid(page, segments, row_bounds, col_bounds, text_rows)

    for r, segs in enumerate(row_segs):
        segs = sorted(segs, key=lambda s: s.x0)
        for i, s in enumerate(segs):
            s.row = r
            s.col0 = grid.col_at(s.x0 + 0.01)
            s.col1 = max(s.col0, grid.col_at(s.x1 - 0.01))
            # a right-aligned segment ends exactly on its anchor boundary
            if s.align == "right":
                c = grid.col_at(s.x1 - 0.01)
                s.col1 = c
            s.free_x0 = segs[i - 1].x1 + 2 if i > 0 else 0.0
            s.free_x1 = segs[i + 1].x0 - 2 if i + 1 < len(segs) else page.width
        # prevent overlaps of merged spans between neighbours in a row
        for a, b in zip(segs, segs[1:]):
            if a.col1 >= b.col0:
                a.col1 = max(a.col0, b.col0 - 1)
    return grid
