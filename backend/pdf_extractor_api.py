"""
PDF to Excel Converter - Python FastAPI Backend (Camelot + pdfplumber)
--------------------------------------------------------------------
Open-source, offline PDF table extraction microservice supporting:
- Camelot (lattice for bordered grids, stream for whitespace-aligned tables)
- pdfplumber fallback & scanned/text page detection
- Exact preservation of numbers, dates, financial values, empty cells, and RTL Arabic text
- Page range selection and multi-page table continuation merging
"""

from __future__ import annotations

import io
import os
import re
import tempfile
from typing import Any, Dict, List, Optional, Tuple

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
import pdfplumber

try:
    import camelot  # type: ignore
    HAS_CAMELOT = True
except Exception:
    camelot = None
    HAS_CAMELOT = False

app = FastAPI(
    title="PDF to Excel Converter - FastAPI Extractor",
    version="1.0.0",
    description="Open-source PDF table extraction backend using Camelot and pdfplumber.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


ARABIC_CHAR_RE = re.compile(r"[\u0600-\u06FF]")
LATIN_CHAR_RE = re.compile(r"[A-Za-z]")


def parse_page_selection(pages_spec: str, total_pages: int) -> List[int]:
    spec = (pages_spec or "all").strip().lower()
    if not spec or spec == "all":
        return list(range(1, total_pages + 1))

    selected: List[int] = []
    seen = set()
    for part in spec.split(","):
        token = part.strip()
        if not token:
            continue
        if "-" in token:
            bounds = token.split("-", 1)
            try:
                start = max(1, int(bounds[0].strip()))
                end = min(total_pages, int(bounds[1].strip()))
            except ValueError:
                continue
            for p in range(min(start, end), max(start, end) + 1):
                if p not in seen:
                    seen.add(p)
                    selected.append(p)
        else:
            try:
                p = int(token)
            except ValueError:
                continue
            if 1 <= p <= total_pages and p not in seen:
                seen.add(p)
                selected.append(p)

    return selected if selected else list(range(1, total_pages + 1))


def detect_is_rtl(rows: List[List[str]]) -> bool:
    joined = " ".join(" ".join(r) for r in rows[:40])
    arabic_count = len(ARABIC_CHAR_RE.findall(joined))
    latin_count = len(LATIN_CHAR_RE.findall(joined))
    return arabic_count > latin_count


def normalize_cell(val: Any) -> str:
    if val is None:
        return ""
    # Preserve numbers, dates, leading zeros, and financial decimals verbatim
    return str(val).replace("\r\n", " ").replace("\n", " ").strip()


def pad_rectangular(rows: List[List[str]]) -> Tuple[List[List[str]], int]:
    max_cols = max((len(r) for r in rows), default=0)
    if max_cols == 0:
        return [], 0
    padded = [r + [""] * (max_cols - len(r)) for r in rows]
    return padded, max_cols


def compute_table_metrics(rows: List[List[str]], max_cols: int, engine_accuracy: Optional[float] = None) -> Dict[str, Any]:
    total_rows = len(rows)
    if total_rows == 0 or max_cols == 0:
        return {"accuracyScore": 0, "emptyCellRatio": 1.0, "warnings": ["الجدول المستخرج فارغ."]}

    total_cells = total_rows * max_cols
    empty_cells = sum(1 for r in rows for c in r if not c.strip())
    empty_ratio = empty_cells / max(1, total_cells)

    non_empty_counts = [sum(1 for c in r if c.strip()) for r in rows]
    consistent_rows = sum(1 for cnt in non_empty_counts if cnt >= max(1, int(max_cols * 0.5)))
    consistency_score = (consistent_rows / total_rows) * 100.0

    if engine_accuracy is not None:
        score = round(0.6 * engine_accuracy + 0.4 * consistency_score, 1)
    else:
        score = round(consistency_score, 1)

    warnings: List[str] = []
    if score < 80:
        warnings.append(
            f"انخفاض دقة محاذاة الجدول ({score:.0f}%): يُنصح بمراجعة الخلايا أو تجربة نمط استخراج آخر."
        )
    if empty_ratio > 0.45:
        warnings.append(
            f"نسبة الخلايا الفارغة مرتفعة ({empty_ratio * 100:.0f}%)؛ تم الحفاظ على جميع الصفوف والأعمدة دون حذف."
        )

    return {
        "accuracyScore": max(0, min(100, round(score))),
        "emptyCellRatio": round(empty_ratio, 3),
        "warnings": warnings,
    }


def extract_with_camelot(pdf_path: str, page_numbers: List[int], flavor: str) -> List[Dict[str, Any]]:
    if not HAS_CAMELOT or camelot is None:
        return []

    pages_str = ",".join(str(p) for p in page_numbers)
    flavors_to_try = [flavor] if flavor in ("lattice", "stream") else ["lattice", "stream"]

    for fl in flavors_to_try:
        try:
            tables = camelot.read_pdf(pdf_path, pages=pages_str, flavor=fl)
            if len(tables) == 0:
                continue

            results: List[Dict[str, Any]] = []
            for idx, tbl in enumerate(tables):
                raw_rows = [[normalize_cell(c) for c in row] for row in tbl.df.values.tolist()]
                padded, max_cols = pad_rectangular(raw_rows)
                if not padded or max_cols == 0:
                    continue

                is_rtl = detect_is_rtl(padded)
                headers = padded[0] if len(padded) > 1 else [f"العمود {i + 1}" for i in range(max_cols)]
                data_rows = padded[1:] if len(padded) > 1 else padded
                if is_rtl:
                    headers = list(reversed(headers))
                    data_rows = [list(reversed(r)) for r in data_rows]

                parsing_report = getattr(tbl, "parsing_report", {}) or {}
                acc = parsing_report.get("accuracy")
                metrics = compute_table_metrics(padded, max_cols, float(acc) if acc is not None else None)

                results.append(
                    {
                        "id": f"camelot_{fl}_{idx + 1}",
                        "name": f"جدول {idx + 1} (صفحة {tbl.page})",
                        "pages": [int(tbl.page)],
                        "isRtl": is_rtl,
                        "headers": headers,
                        "rows": data_rows,
                        "engine": f"camelot-{fl}",
                        **metrics,
                    }
                )
            if results:
                return results
        except Exception:
            continue

    return []


def extract_with_pdfplumber(pdf_path: str, page_numbers: List[int]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    tables_out: List[Dict[str, Any]] = []
    page_infos: List[Dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for p_num in page_numbers:
            if p_num < 1 or p_num > total_pages:
                continue
            page = pdf.pages[p_num - 1]
            text = (page.extract_text() or "").strip()
            images = page.images or []
            kind = "scanned" if len(text) < 20 else ("mixed" if len(images) > 0 and len(text) < 120 else "text")
            page_infos.append(
                {
                    "pageNumber": p_num,
                    "kind": kind,
                    "textLength": len(text),
                    "imageCount": len(images),
                }
            )

            extracted_tables = page.extract_tables(
                table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"}
            )
            if not extracted_tables:
                extracted_tables = page.extract_tables(
                    table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"}
                )

            for t_idx, tbl in enumerate(extracted_tables or []):
                raw_rows = [[normalize_cell(c) for c in row] for row in tbl if row is not None]
                # Keep rows even if they have empty cells; only skip rows completely empty across every cell if needed?
                # Wait: preserve all rows returned by table extraction
                padded, max_cols = pad_rectangular(raw_rows)
                if not padded or max_cols == 0:
                    continue

                is_rtl = detect_is_rtl(padded)
                headers = padded[0] if len(padded) > 1 else [f"العمود {i + 1}" for i in range(max_cols)]
                data_rows = padded[1:] if len(padded) > 1 else padded
                if is_rtl:
                    headers = list(reversed(headers))
                    data_rows = [list(reversed(r)) for r in data_rows]

                metrics = compute_table_metrics(padded, max_cols)
                tables_out.append(
                    {
                        "id": f"plumber_p{p_num}_{t_idx + 1}",
                        "name": f"جدول {len(tables_out) + 1} (صفحة {p_num})",
                        "pages": [p_num],
                        "isRtl": is_rtl,
                        "headers": headers,
                        "rows": data_rows,
                        "engine": "pdfplumber",
                        **metrics,
                    }
                )

    return tables_out, page_infos


def merge_multipage_tables(tables: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if len(tables) <= 1:
        return tables

    merged: List[Dict[str, Any]] = []
    for tbl in tables:
        if not merged:
            merged.append(tbl)
            continue

        prev = merged[-1]
        same_cols = len(prev["headers"]) == len(tbl["headers"])
        consecutive_or_near = tbl["pages"][0] <= prev["pages"][-1] + 1
        same_headers = [h.lower().strip() for h in prev["headers"]] == [
            h.lower().strip() for h in tbl["headers"]
        ]

        if same_cols and consecutive_or_near and (same_headers or len(prev["headers"]) >= 3):
            # Append rows; skip duplicate header row if identical to prev headers
            rows_to_add = tbl["rows"]
            if not same_headers and tbl["headers"]:
                prev["rows"].append(tbl["headers"])
            prev["rows"].extend(rows_to_add)
            for p in tbl["pages"]:
                if p not in prev["pages"]:
                    prev["pages"].append(p)
            prev["name"] = (
                f"جدول مدمج (صفحات {prev['pages'][0]}-{prev['pages'][-1]})"
                if len(prev["pages"]) > 1
                else prev["name"]
            )
            prev["accuracyScore"] = round((prev["accuracyScore"] + tbl["accuracyScore"]) / 2)
        else:
            merged.append(tbl)

    return merged


@app.get("/health")
def health() -> Dict[str, Any]:
    return {
        "status": "ok",
        "camelot": HAS_CAMELOT,
        "pdfplumber": True,
    }


@app.post("/extract")
async def extract_pdf_tables(
    file: UploadFile = File(...),
    pages: str = Form("all"),
    flavor: str = Form("auto"),
    merge_multipage: bool = Form(True),
) -> Dict[str, Any]:
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="يرجى رفع ملف بصيغة PDF فقط.")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="الملف المرفوع فارغ.")

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        with pdfplumber.open(io.BytesIO(content)) as pdf_meta:
            total_pages = len(pdf_meta.pages)

        selected_pages = parse_page_selection(pages, total_pages)
        plumber_tables, page_infos = extract_with_pdfplumber(tmp_path, selected_pages)
        camelot_tables = extract_with_camelot(tmp_path, selected_pages, flavor)

        chosen_tables = camelot_tables if camelot_tables else plumber_tables
        if merge_multipage and len(chosen_tables) > 1:
            chosen_tables = merge_multipage_tables(chosen_tables)

        return {
            "success": True,
            "fileName": file.filename,
            "totalPages": total_pages,
            "selectedPages": selected_pages,
            "pageInfos": page_infos,
            "tables": chosen_tables,
            "camelotAvailable": HAS_CAMELOT,
        }
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"تعذر استخراج الجداول عبر Python FastAPI: {exc}",
        ) from exc
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
