"""FastAPI application: PDF statement <-> Excel round-trip."""
from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response

from .excel_writer import write_workbook
from .extract import configure_tesseract, extract_pages
from .pdf_writer import excel_to_pdf

STATIC_DIR = Path(__file__).parent / "static"
OUTPUT_DIR = Path(__file__).resolve().parents[1] / "output"
AUTO_OPEN = os.environ.get("AUTO_OPEN", "1") != "0"
MAX_UPLOAD = 40 * 1024 * 1024

app = FastAPI(title="Statement PDF <-> Excel")


def _safe_stem(name: str | None, fallback: str) -> str:
    stem = Path(name or fallback).stem
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", stem).strip("_")
    return stem or fallback


async def _read_upload(file: UploadFile, exts: tuple[str, ...]) -> bytes:
    if not file.filename or not file.filename.lower().endswith(exts):
        raise HTTPException(400, f"Expected a {' / '.join(exts)} file.")
    data = await file.read()
    if not data:
        raise HTTPException(400, "Empty upload.")
    if len(data) > MAX_UPLOAD:
        raise HTTPException(413, "File too large (max 40 MB).")
    return data


def _save_and_open(data: bytes, name: str) -> Path:
    """Save the result under output/ and open it with the default app."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    path = OUTPUT_DIR / name
    try:
        path.write_bytes(data)
    except PermissionError:  # file is open in Excel / a viewer
        stem, i = path.stem, 2
        while (path := OUTPUT_DIR / f"{stem} ({i}){path.suffix}").exists():
            i += 1
        path.write_bytes(data)
    if AUTO_OPEN:
        _launch(path)
    return path


def _launch(path: Path, reveal: bool = False) -> None:
    if os.name == "nt":
        if reveal:
            subprocess.Popen(["explorer", "/select,", str(path)])
        else:
            os.startfile(path)  # noqa: S606
    else:
        subprocess.Popen(["xdg-open", str(path.parent if reveal else path)])


def _output_file(name: str) -> Path:
    path = (OUTPUT_DIR / name).resolve()
    if path.parent != OUTPUT_DIR.resolve() or not path.is_file():
        raise HTTPException(404, "File not found.")
    return path


@app.post("/api/open/{name}")
def open_file(name: str) -> dict:
    _launch(_output_file(name))
    return {"ok": True}


@app.post("/api/reveal/{name}")
def reveal_file(name: str) -> dict:
    _launch(_output_file(name), reveal=True)
    return {"ok": True}


@app.get("/api/files/{name}")
def download_file(name: str) -> FileResponse:
    return FileResponse(_output_file(name), filename=name)


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
def health() -> dict:
    return {"ok": True, "tesseract": configure_tesseract()}


@app.post("/api/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)) -> Response:
    pdf_bytes = await _read_upload(file, (".pdf",))
    try:
        pages = extract_pages(pdf_bytes)
        xlsx = write_workbook(pages, pdf_bytes, file.filename or "statement.pdf")
    except RuntimeError as e:  # e.g. Tesseract missing
        raise HTTPException(422, str(e)) from e
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"Conversion failed: {e}") from e
    name = _safe_stem(file.filename, "statement") + ".xlsx"
    saved = _save_and_open(xlsx, name)
    return Response(
        xlsx,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Access-Control-Expose-Headers": "X-Pages, X-Sources, X-Changes, X-Saved-As",
            "Content-Disposition": f'attachment; filename="{name}"',
            "X-Pages": str(len(pages)),
            "X-Saved-As": saved.name,
            "X-Sources": ",".join(p.source for p in pages),
        },
    )


@app.post("/api/excel-to-pdf")
async def excel_to_pdf_route(file: UploadFile = File(...)) -> Response:
    xlsx_bytes = await _read_upload(file, (".xlsx", ".xlsm"))
    try:
        pdf, n_changes = excel_to_pdf(xlsx_bytes)
    except ValueError as e:
        raise HTTPException(422, str(e)) from e
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"Conversion failed: {e}") from e
    name = _safe_stem(file.filename, "statement").removesuffix("_edited") + "_edited.pdf"
    saved = _save_and_open(pdf, name)
    return Response(
        pdf,
        media_type="application/pdf",
        headers={
            "Access-Control-Expose-Headers": "X-Pages, X-Sources, X-Changes, X-Saved-As",
            "Content-Disposition": f'attachment; filename="{name}"',
            "X-Changes": str(n_changes),
            "X-Saved-As": saved.name,
        },
    )


@app.exception_handler(HTTPException)
async def _http_exc(_, exc: HTTPException) -> JSONResponse:
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("backend.app:app", host="127.0.0.1", port=int(os.environ.get("PORT", 8000)), reload=True)
