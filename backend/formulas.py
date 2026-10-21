"""Turn the statement's amount columns into live numbers and formulas.

* amount cells under WITHDRAWALS / DEPOSITS / BALANCE become numeric cells
  whose number format reproduces the original text (grouping, decimals,
  Cr/Dr suffix) so that an unedited export is unchanged
* the "Page Total" row gets ``=SUM(...)`` over each amount column
* every balance after the first is ``=previous balance - withdrawal + deposit``

Also provides a tiny evaluator for exactly those formulas, used on export when
the workbook was never recalculated by Excel (no cached values).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string, get_column_letter

from .layout import Grid, Segment

AMOUNT_RE = re.compile(r"^(?P<num>\d{1,3}(?:,\d{2,3})*(?:\.\d+)?|\d+(?:\.\d+)?)\s*(?P<suffix>Cr|Dr|CR|DR)?$")
INDIAN_RE = re.compile(r"^\d{1,2}(,\d{2})+,\d{3}(\.\d+)?$")

_ROLES = {
    "withdrawals": {"withdrawals", "withdrawal", "debit", "debits", "dr", "debit amount", "withdrawal amount"},
    "deposits": {"deposits", "deposit", "credit", "credits", "cr", "credit amount", "deposit amount"},
    "balance": {"balance", "closing balance", "running balance", "bal"},
}
PAGE_TOTAL_RE = re.compile(r"page\s*total", re.I)
GRAND_TOTAL_RE = re.compile(r"grand\s*total|closing", re.I)


@dataclass
class Amount:
    value: float
    decimals: int
    grouped: bool
    indian: bool
    suffix: str  # "", "Cr", "Dr", ...


def parse_amount(text: str) -> Amount | None:
    m = AMOUNT_RE.match(text.strip())
    if not m:
        return None
    num = m.group("num")
    dec = len(num.split(".")[1]) if "." in num else 0
    return Amount(
        value=float(num.replace(",", "")),
        decimals=dec,
        grouped="," in num,
        indian=bool(INDIAN_RE.match(num)),
        suffix=m.group("suffix") or "",
    )


def format_amount(value: float, like: Amount) -> str:
    """Render ``value`` the way ``like`` was written on the statement."""
    neg = value < 0
    v = abs(value)
    body = f"{v:.{like.decimals}f}"
    if like.grouped:
        ip, _, fp = body.partition(".")
        if like.indian and len(ip) > 3:
            head, last3 = ip[:-3], ip[-3:]
            parts = []
            while len(head) > 2:
                parts.insert(0, head[-2:])
                head = head[:-2]
            if head:
                parts.insert(0, head)
            ip = ",".join(parts + [last3])
        else:
            ip = f"{int(ip):,}"
        body = ip + ("." + fp if fp else "")
    suffix = like.suffix
    if suffix and neg:  # a balance that flips sign flips Cr <-> Dr
        suffix = {"Cr": "Dr", "Dr": "Cr", "CR": "DR", "DR": "CR"}.get(suffix, suffix)
    elif neg:
        body = "-" + body
    return body + suffix


def number_format(like: Amount) -> str:
    dec = "." + "0" * like.decimals if like.decimals else ""
    if like.grouped and like.indian:
        pos = f"[>=10000000]##\\,##\\,##\\,##0{dec};[>=100000]##\\,##\\,##0{dec};#,##0{dec}"
        # conditional sections cannot carry a suffix cleanly; fall back to western grouping
        pos = f"#,##0{dec}"
    elif like.grouped:
        pos = f"#,##0{dec}"
    else:
        pos = f"0{dec}"
    if like.suffix:
        flip = {"Cr": "Dr", "Dr": "Cr", "CR": "DR", "DR": "CR"}.get(like.suffix, like.suffix)
        return f'{pos}"{like.suffix}";{pos}"{flip}"'
    return f"{pos};-{pos}"


# --------------------------------------------------------------------------- #
# Applying to a worksheet
# --------------------------------------------------------------------------- #
def _overlaps(a: Segment, b: Segment, tol: float) -> bool:
    return a.x1 > b.x0 - tol and a.x0 < b.x1 + tol


def apply_formulas(ws, grid: Grid, refs: dict[int, str]) -> None:
    """``refs`` maps segment id -> cell reference on ``ws``."""
    segs = [s for s in grid.segments if s.id in refs]
    headers: dict[str, Segment] = {}
    for s in segs:
        key = s.text.strip().lower().rstrip(":").strip()
        for role, names in _ROLES.items():
            if key in names and role not in headers:
                headers[role] = s
    if not ({"withdrawals", "deposits"} & headers.keys()):
        return
    header_row = max(h.row for h in headers.values())
    page_total = next((s for s in segs if PAGE_TOTAL_RE.search(s.text)), None)
    grand_total = next((s for s in segs if GRAND_TOTAL_RE.search(s.text)), None)
    last_row = min(r for r in (page_total.row if page_total else None,
                               grand_total.row if grand_total else None,
                               grid.nrows) if r is not None)
    tol = 0.01 * grid.page.width

    # transaction amount cells per role, in row order
    cells: dict[str, list[tuple[int, str, Amount]]] = {r: [] for r in headers}
    for s in sorted(segs, key=lambda s: s.row):
        if not (header_row < s.row < last_row):
            continue
        amt = parse_amount(s.text)
        if amt is None:
            continue
        for role, h in headers.items():
            if _overlaps(s, h, tol):
                cells[role].append((_xl_row(refs[s.id]), refs[s.id], amt))
                break

    def set_number(ref: str, amt: Amount) -> None:
        c = ws[ref]
        c.value = amt.value
        c.number_format = number_format(amt)

    for role, lst in cells.items():
        for _, ref, amt in lst:
            set_number(ref, amt)

    # running balance
    bal = cells.get("balance", [])
    if bal:
        # first row: opening balance (not printed) is derived from the statement
        # and written as a literal, so edits to the first amounts flow through too
        row, ref, amt = bal[0]
        w = _cell_in_row(cells.get("withdrawals", []), row)
        d = _cell_in_row(cells.get("deposits", []), row)
        if w or d:
            opening = _signed(amt) + (w[1].value if w else 0.0) - (d[1].value if d else 0.0)
            expr = f"{opening:.{amt.decimals}f}"
            if w:
                expr += f"-{w[0]}"
            if d:
                expr += f"+{d[0]}"
            ws[ref].value = f"={expr}"
            ws[ref].number_format = number_format(amt)
    for (prev_row, prev_ref, prev_amt), (row, ref, amt) in zip(bal, bal[1:]):
        w = _cell_in_row(cells.get("withdrawals", []), row)
        d = _cell_in_row(cells.get("deposits", []), row)
        if w is None and d is None:
            continue
        # only trust the formula where the statement's own arithmetic confirms it
        expected = _signed(prev_amt) - (w[1].value if w else 0.0) + (d[1].value if d else 0.0)
        if abs(expected - _signed(amt)) > 0.006:
            continue
        expr = prev_ref
        if w:
            expr += f"-{w[0]}"
        if d:
            expr += f"+{d[0]}"
        ws[ref].value = f"={expr}"
        ws[ref].number_format = number_format(amt)

    # page total row
    if page_total is None:
        return
    for role in ("withdrawals", "deposits"):
        lst = cells.get(role, [])
        if not lst:
            continue
        col = column_index_from_string(re.match(r"[A-Z]+", lst[0][1]).group())
        first, last = lst[0][0], lst[-1][0]
        tot = _segment_in_row(segs, headers[role], page_total.row, tol)
        if tot is None:
            continue
        amt = parse_amount(tot.text) or Amount(0.0, 2, True, False, "")
        rng = f"{get_column_letter(col)}{first}:{get_column_letter(col)}{last}"
        ws[refs[tot.id]].value = f"=SUM({rng})"
        ws[refs[tot.id]].number_format = number_format(amt)
    if bal and "balance" in headers:
        tot = _segment_in_row(segs, headers["balance"], page_total.row, tol)
        if tot is not None:
            amt = parse_amount(tot.text) or Amount(0.0, 2, True, False, "Cr")
            ws[refs[tot.id]].value = f"={bal[-1][1]}"
            ws[refs[tot.id]].number_format = number_format(amt)


def _xl_row(ref: str) -> int:
    return int(re.search(r"\d+", ref).group())


def _signed(amt: Amount) -> float:
    return -amt.value if amt.suffix.lower() == "dr" else amt.value


def _cell_in_row(lst: list[tuple[int, str, Amount]], row: int) -> tuple[str, Amount] | None:
    for r, ref, amt in lst:
        if r == row:
            return ref, amt
    return None


def _segment_in_row(segs: list[Segment], header: Segment, row: int, tol: float) -> Segment | None:
    for s in segs:
        if s.row == row and _overlaps(s, header, tol) and parse_amount(s.text):
            return s
    return None


# --------------------------------------------------------------------------- #
# Evaluating our own formulas (export without Excel recalculation)
# --------------------------------------------------------------------------- #
_REF = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
_SAFE = re.compile(r"^[\d\.\+\-\*/\(\) ]+$")


def evaluate(formula: str, ws_values, ws_formulas, _depth: int = 0) -> float | None:
    """Evaluate ``=SUM(A1:A9)`` / ``=N23-J26+L26`` style formulas."""
    if _depth > 50:
        return None
    expr = formula.lstrip("=")

    def cell(ref: str) -> float:
        v = ws_values[ref].value
        if v is None or (isinstance(v, str) and v.startswith("=")):
            f = ws_formulas[ref].value
            if isinstance(f, str) and f.startswith("="):
                v = evaluate(f, ws_values, ws_formulas, _depth + 1)
        if v is None or v == "":
            return 0.0
        if isinstance(v, (int, float)):
            return float(v)
        amt = parse_amount(str(v))
        return amt.value if amt else 0.0

    def sum_range(m: re.Match) -> str:
        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        total = 0.0
        for col in range(column_index_from_string(c1), column_index_from_string(c2) + 1):
            for row in range(r1, r2 + 1):
                total += cell(f"{get_column_letter(col)}{row}")
        return repr(total)

    expr = re.sub(r"SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)", sum_range, expr, flags=re.I)
    expr = _REF.sub(lambda m: repr(cell(f"{m.group(1)}{m.group(2)}")), expr)
    if not _SAFE.match(expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))  # noqa: S307 - arithmetic only, validated above
    except Exception:  # noqa: BLE001
        return None
