"""
title: Edit Office Files
author: giofsp
author_url: https://github.com/sergiofspedro
description: Unified tool to read, edit, and create Office files (.xlsx, .xls, .docx, .pptx) preserving original formatting and styles. Detects highlights, bold, italic formatting. Detects legacy .doc and .ppt. Note: Track changes are not supported.
version: 1.1.0
requirements: openpyxl, python-docx, python-pptx, xlrd
"""
import json
import io
import os
import re
import sqlite3
import sys
import traceback
from copy import copy
from typing import Optional, List, Dict, Any


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
_DB_PATH = os.path.join(
    os.environ.get("OPEN_WEBUI_DATA_DIR", ""),
    "data", "webui.db",
)
if not os.path.isfile(_DB_PATH):
    _DB_PATH = os.path.join(
        os.path.expanduser("~"),
        "AppData", "Roaming", "open-webui", "data", "webui.db",
    )

_UPLOAD_DIR = os.path.join(
    os.environ.get("OPEN_WEBUI_DATA_DIR", ""),
    "data", "uploads",
)
if not os.path.isdir(_UPLOAD_DIR):
    _UPLOAD_DIR = os.path.join(
        os.path.expanduser("~"),
        "AppData", "Roaming", "open-webui", "data", "uploads",
    )

_EXPORT_DIR = r"C:\Users\Administrator\AppData\Local\open-webui\exports"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _resolve_file_path(file_id: str) -> Optional[str]:
    """Resolve an Open WebUI file UUID to an absolute disk path."""
    try:
        conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
        row = conn.execute(
            "SELECT path FROM file WHERE id = ?", (file_id,)
        ).fetchone()
        conn.close()
    except Exception as exc:
        print(f"[office] DB lookup failed for {file_id}: {exc}", file=sys.stderr)
        return None

    if not row or not row[0]:
        # Fallback: try by filename
        try:
            conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
            row = conn.execute(
                "SELECT path FROM file WHERE filename LIKE ?",
                (f"%{file_id}%",),
            ).fetchone()
            conn.close()
        except Exception:
            pass

    if not row or not row[0]:
        print(f"[office] No path for file_id {file_id}", file=sys.stderr)
        return None

    path = row[0]
    if os.path.isfile(path):
        return path

    # Fallback: uploads directory
    candidate = os.path.join(_UPLOAD_DIR, os.path.basename(path))
    if os.path.isfile(candidate):
        return candidate

    # Last resort: UUID prefix match in uploads
    prefix = file_id.split("-")[0] if "-" in file_id else file_id[:8]
    if os.path.isdir(_UPLOAD_DIR):
        for name in os.listdir(_UPLOAD_DIR):
            if name.startswith(prefix):
                candidate = os.path.join(_UPLOAD_DIR, name)
                if os.path.isfile(candidate):
                    return candidate

    print(f"[office] File not found on disk: {path}", file=sys.stderr)
    return None


def _read_file_bytes(file_id: str) -> Optional[bytes]:
    """Return raw bytes for an Open WebUI file."""
    path = _resolve_file_path(file_id)
    if path is None:
        return None
    try:
        with open(path, "rb") as fh:
            return fh.read()
    except Exception as exc:
        print(f"[office] Read failed for {path}: {exc}", file=sys.stderr)
        return None


def _detect_type(filename: str) -> str:
    """Detect Office file type from filename."""
    lower = filename.lower()
    if lower.endswith(".xlsx"):
        return "xlsx"
    if lower.endswith(".xls"):
        return "xls"
    if lower.endswith(".docx"):
        return "docx"
    if lower.endswith(".doc"):
        return "doc"
    if lower.endswith(".pptx"):
        return "pptx"
    if lower.endswith(".ppt"):
        return "ppt"
    return "unknown"


def _save_file_sync(file_bytes: bytes, filename: str) -> Optional[str]:
    """Save file to exports dir and return HTTP download URL."""
    try:
        os.makedirs(_EXPORT_DIR, exist_ok=True)
        safe = "".join(c for c in filename if c.isalnum() or c in "._- ")
        if not safe:
            safe = "export"
        base_name, ext = os.path.splitext(safe)
        fname = safe
        counter = 2
        while os.path.exists(os.path.join(_EXPORT_DIR, fname)):
            fname = f"{base_name} ({counter}){ext}"
            counter += 1
        fpath = os.path.join(_EXPORT_DIR, fname)
        with open(fpath, "wb") as f:
            f.write(file_bytes)
        return f"http://localhost:9000/{fname}"
    except Exception as e:
        print(f"[office] Export save failed: {e}", file=sys.stderr)
        return None


def _cell_value(cell):
    """Extract a JSON-safe value from an openpyxl cell."""
    import datetime
    v = cell.value
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    if isinstance(v, (int, float, str, bool)):
        return v
    return str(v)


def _xls_to_xlsx(xls_data: bytes) -> bytes:
    """Convert .xls bytes to .xlsx bytes using xlrd + openpyxl.

    Returns an in-memory .xlsx workbook as bytes.
    """
    import xlrd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    xls_book = xlrd.open_workbook(file_contents=xls_data)
    xlsx_wb = openpyxl.Workbook()
    # Remove the default sheet; we'll add one per xls sheet
    xlsx_wb.remove(xlsx_wb.active)

    for sheet_idx in range(xls_book.nsheets):
        xls_sheet = xls_book.sheet_by_index(sheet_idx)
        ws = xlsx_wb.create_sheet(title=xls_sheet.name[:31])  # Excel 31-char limit

        for rx in range(xls_sheet.nrows):
            for cx in range(xls_sheet.ncols):
                cell = xls_sheet.cell(rx, cx)
                value = cell.value

                # xlrd date handling: if cell type is XL_CELL_DATE, convert
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        dt_tuple = xls_book.datemode, int(value)
                        import datetime as _dt
                        value = _dt.datetime(*xlrd.xldate_as_tuple(value, xls_book.datemode))
                    except Exception:
                        pass
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = bool(value)

                ws.cell(row=rx + 1, column=cx + 1, value=value)

        # Auto-fit column widths (rough estimate)
        for col_cells in ws.columns:
            max_len = 0
            for cell in col_cells:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            letter = openpyxl.utils.get_column_letter(col_cells[0].column)
            ws.column_dimensions[letter].width = min(max_len + 2, 50)

    out = io.BytesIO()
    xlsx_wb.save(out)
    xlsx_wb.close()
    xls_book.release_resources()
    out.seek(0)
    return out.read()


# =========================================================================
class Tools:
    class Valves:
        pass

    def __init__(self):
        pass

    # -----------------------------------------------------------------
    # Internal: save and return markdown link
    # -----------------------------------------------------------------
    async def _save_and_link(self, file_bytes: bytes, filename: str) -> tuple:
        """Save file and return (url, filename) or (None, None)."""
        url = _save_file_sync(file_bytes, filename)
        return (url, filename) if url else (None, None)

    # -----------------------------------------------------------------
    # READ
    # -----------------------------------------------------------------
    async def read_file(
        self,
        file_id: str,
        max_rows: int = 500,
        __user__=None,
        __request__=None,
    ) -> str:
        """Read any Office file (.xlsx, .xls, .docx, .pptx) and return its contents as structured JSON.

        Auto-detects the file type from the file ID or filename.
        For xlsx/xls: returns sheets with headers and rows.
        For docx: returns paragraphs with styles and tables.
        For pptx: returns slides with shapes and text.
        Legacy .doc and .ppt formats return a helpful error message.

        Args:
            file_id: The Open WebUI file ID (UUID) or filename
            max_rows: Maximum rows to return for xlsx (default 500)
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({
                    "error": (
                        f"Could not read file {file_id}. "
                        "Make sure the file was uploaded via the chat."
                    )
                })

            # Detect type from filename in DB
            try:
                conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row = conn.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn.close()
                filename = row[0] if row else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            result: Dict[str, Any] = {
                "file_id": file_id,
                "filename": filename,
                "type": file_type,
            }

            if file_type == "xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
                result["sheets"] = []
                for sn in wb.sheetnames:
                    ws = wb[sn]
                    sheet: Dict[str, Any] = {
                        "name": sn,
                        "headers": [],
                        "rows": [],
                        "total_rows": ws.max_row or 0,
                        "total_cols": ws.max_column or 0,
                    }
                    max_r = min(ws.max_row or 0, max_rows)
                    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=max_r), 1):
                        rd = [_cell_value(c) for c in row]
                        if ri == 1:
                            sheet["headers"] = [str(v) if v is not None else "" for v in rd]
                        else:
                            sheet["rows"].append(rd)
                    result["sheets"].append(sheet)
                wb.close()

            elif file_type == "xls":
                import xlrd
                xls_book = xlrd.open_workbook(file_contents=file_data)
                result["sheets"] = []
                for sheet_idx in range(xls_book.nsheets):
                    xls_sheet = xls_book.sheet_by_index(sheet_idx)
                    sheet: Dict[str, Any] = {
                        "name": xls_sheet.name,
                        "headers": [],
                        "rows": [],
                        "total_rows": xls_sheet.nrows,
                        "total_cols": xls_sheet.ncols,
                    }
                    max_r = min(xls_sheet.nrows, max_rows)
                    for rx in range(max_r):
                        row_values = []
                        for cx in range(xls_sheet.ncols):
                            cell = xls_sheet.cell(rx, cx)
                            value = cell.value
                            if cell.ctype == xlrd.XL_CELL_DATE:
                                try:
                                    import datetime as _dt
                                    value = _dt.datetime(*xlrd.xldate_as_tuple(value, xls_book.datemode))
                                except Exception:
                                    pass
                            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                                value = bool(value)
                            row_values.append(value)
                        if rx == 0:
                            sheet["headers"] = [str(v) if v is not None else "" for v in row_values]
                        else:
                            sheet["rows"].append(row_values)
                    result["sheets"].append(sheet)
                xls_book.release_resources()

            elif file_type == "docx":
                from docx import Document
                from docx.enum.text import WD_COLOR_INDEX
                doc = Document(io.BytesIO(file_data))
                paragraphs = []
                for p in doc.paragraphs:
                    if p.text.strip():
                        style = p.style.name if p.style else "Normal"
                        runs_info = []
                        for run in p.runs:
                            run_data = {"text": run.text}
                            if run.font.highlight_color and run.font.highlight_color != WD_COLOR_INDEX.AUTO:
                                run_data["highlighted"] = True
                                run_data["highlight_color"] = str(run.font.highlight_color)
                            if run.font.bold:
                                run_data["bold"] = True
                            if run.font.italic:
                                run_data["italic"] = True
                            runs_info.append(run_data)
                        paragraphs.append({"style": style, "text": p.text, "runs": runs_info})
                tables = []
                for t in doc.tables:
                    tbl = {"rows": []}
                    for row in t.rows:
                        tbl["rows"].append([cell.text for cell in row.cells])
                    tables.append(tbl)
                result["paragraphs"] = paragraphs
                result["tables"] = tables

            elif file_type == "pptx":
                from pptx import Presentation
                prs = Presentation(io.BytesIO(file_data))
                slides = []
                for si, slide in enumerate(prs.slides, 1):
                    sdata: Dict[str, Any] = {"number": si, "shapes": []}
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and shape.text.strip():
                            sdata["shapes"].append({
                                "type": str(shape.shape_type),
                                "name": shape.name,
                                "text": shape.text[:500],
                            })
                        if shape.has_table:
                            tbl = {"rows": []}
                            for row in shape.table.rows:
                                tbl["rows"].append([cell.text for cell in row.cells])
                            sdata["tables"] = tbl
                    slides.append(sdata)
                result["slides"] = slides

            elif file_type == "doc":
                result["error"] = "Legacy .doc format is not supported. Please convert to .docx first."

            elif file_type == "ppt":
                result["error"] = "Legacy .ppt format is not supported. Please convert to .pptx first."

            else:
                result["error"] = f"Unsupported file type. Detected: {file_type}. Supported: xlsx, xls, docx, pptx"

            return json.dumps(result, indent=2, default=str, ensure_ascii=False)

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # ADD CONTENT
    # -----------------------------------------------------------------
    async def add_content(
        self,
        file_id: str,
        content: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Add new content to an Office file while preserving original formatting.

        For spreadsheets (xlsx): content is CSV text with rows to add.
        For documents (docx): content is text to append at the end.
        For presentations (pptx): each line defines a new slide. Use "---" as separator between slides.

        Args:
            file_id: File ID to edit
            content: Content to add (CSV for xlsx, text for docx/pptx)
            output_filename: Optional output filename
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({"error": f"Could not read file {file_id}"})

            # Detect type
            try:
                conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row = conn.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn.close()
                filename = row[0] if row else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            out = io.BytesIO()
            out_name = output_filename

            if file_type == "xlsx":
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                wb = openpyxl.load_workbook(io.BytesIO(file_data))
                ws = wb.active
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"

                # Parse CSV content
                import csv as _csv_mod
                reader = _csv_mod.reader(io.StringIO(content))
                parsed_rows = []
                for csv_row in reader:
                    converted = []
                    for v in csv_row:
                        v = v.strip()
                        if v == '':
                            converted.append(None)
                        else:
                            try:
                                converted.append(int(v))
                            except ValueError:
                                try:
                                    converted.append(float(v))
                                except ValueError:
                                    if v.lower() == 'true':
                                        converted.append(True)
                                    elif v.lower() == 'false':
                                        converted.append(False)
                                    else:
                                        converted.append(v)
                    parsed_rows.append(converted)

                if not parsed_rows:
                    return json.dumps({"error": "No rows provided in CSV content"})

                # Get reference styles from last row
                ref = {}
                if ws.max_row and ws.max_row >= 1:
                    for cell in ws[ws.max_row]:
                        if cell.has_style:
                            ref[cell.column] = {
                                "font": copy(cell.font),
                                "fill": copy(cell.fill),
                                "border": copy(cell.border),
                                "alignment": copy(cell.alignment),
                                "number_format": cell.number_format,
                            }

                start = (ws.max_row or 0) + 1
                for i, rd in enumerate(parsed_rows):
                    for j, v in enumerate(rd, 1):
                        cell = ws.cell(row=start + i, column=j)
                        if j in ref:
                            try:
                                cell.font = copy(ref[j]["font"])
                                cell.fill = copy(ref[j]["fill"])
                                cell.border = copy(ref[j]["border"])
                                cell.alignment = copy(ref[j]["alignment"])
                                cell.number_format = ref[j]["number_format"]
                            except Exception:
                                pass
                        cell.value = v

                wb.save(out)
                wb.close()

            elif file_type == "xls":
                # Convert .xls to .xlsx, then apply same add logic
                file_data = _xls_to_xlsx(file_data)
                file_type = "xlsx"
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                wb = openpyxl.load_workbook(io.BytesIO(file_data))
                ws = wb.active

                # Parse CSV content
                import csv as _csv_mod
                reader = _csv_mod.reader(io.StringIO(content))
                parsed_rows = []
                for csv_row in reader:
                    converted = []
                    for v in csv_row:
                        v = v.strip()
                        if v == '':
                            converted.append(None)
                        else:
                            try:
                                converted.append(int(v))
                            except ValueError:
                                try:
                                    converted.append(float(v))
                                except ValueError:
                                    if v.lower() == 'true':
                                        converted.append(True)
                                    elif v.lower() == 'false':
                                        converted.append(False)
                                    else:
                                        converted.append(v)
                    parsed_rows.append(converted)

                if not parsed_rows:
                    return json.dumps({"error": "No rows provided in CSV content"})

                ref = {}
                if ws.max_row and ws.max_row >= 1:
                    for cell in ws[ws.max_row]:
                        if cell.has_style:
                            ref[cell.column] = {
                                "font": copy(cell.font),
                                "fill": copy(cell.fill),
                                "border": copy(cell.border),
                                "alignment": copy(cell.alignment),
                                "number_format": cell.number_format,
                            }

                start = (ws.max_row or 0) + 1
                for i, rd in enumerate(parsed_rows):
                    for j, v in enumerate(rd, 1):
                        cell = ws.cell(row=start + i, column=j)
                        if j in ref:
                            try:
                                cell.font = copy(ref[j]["font"])
                                cell.fill = copy(ref[j]["fill"])
                                cell.border = copy(ref[j]["border"])
                                cell.alignment = copy(ref[j]["alignment"])
                                cell.number_format = ref[j]["number_format"]
                            except Exception:
                                pass
                        cell.value = v

                wb.save(out)
                wb.close()

            elif file_type == "docx":
                from docx import Document
                doc = Document(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.docx"

                # Append paragraphs, preserving last paragraph's style
                last_style = "Normal"
                if doc.paragraphs:
                    last_style = doc.paragraphs[-1].style.name if doc.paragraphs[-1].style else "Normal"

                for line in content.split("\n"):
                    doc.add_paragraph(line, style=last_style)

                out = io.BytesIO()
                doc.save(out)

            elif file_type == "pptx":
                from pptx import Presentation
                from pptx.util import Inches
                prs = Presentation(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.pptx"

                # Split content by "---" into slides
                slide_specs = re.split(r'\n---\n|\r\n---\r\n|\n---\n', content)
                blank_layout = prs.slide_layouts[6]  # Blank layout

                for spec in slide_specs:
                    spec = spec.strip()
                    if not spec:
                        continue
                    lines = spec.split("\n")
                    title = lines[0].strip()
                    body = "\n".join(lines[1:]).strip() if len(lines) > 1 else ""

                    slide = prs.slides.add_slide(blank_layout)
                    # Add title
                    txBox = slide.shapes.add_textbox(
                        Inches(0.5), Inches(0.3), Inches(9), Inches(1)
                    )
                    tf = txBox.text_frame
                    tf.text = title
                    p = tf.paragraphs[0]
                    p.font.size = Inches(0.6)
                    p.font.bold = True

                    # Add body text
                    if body:
                        txBox2 = slide.shapes.add_textbox(
                            Inches(0.5), Inches(1.5), Inches(9), Inches(5.5)
                        )
                        tf2 = txBox2.text_frame
                        tf2.text = body
                        for para in tf2.paragraphs:
                            para.font.size = Inches(0.3)

                out = io.BytesIO()
                prs.save(out)

            else:
                return json.dumps({"error": f"Unsupported type: {file_type}"})

            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name)
            if url:
                return f"[{name}]({url})\n\nAdded content to {file_type.upper()} file, preserving original formatting."
            return json.dumps({"error": "Could not save file"})

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # REPLACE TEXT
    # -----------------------------------------------------------------
    async def replace_text(
        self,
        file_id: str,
        find_text: str,
        replace_with: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Find and replace text in any Office file while preserving original formatting.

        Works on cell values in xlsx, paragraph text in docx, and shape text in pptx.

        Args:
            file_id: File ID to edit
            find_text: Text to find
            replace_with: Text to replace with
            output_filename: Optional output filename
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({"error": f"Could not read file {file_id}"})

            # Detect type
            try:
                conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
                row = conn.execute(
                    "SELECT filename FROM file WHERE id = ?", (file_id,)
                ).fetchone()
                conn.close()
                filename = row[0] if row else file_id
            except Exception:
                filename = file_id

            file_type = _detect_type(filename)
            out = io.BytesIO()
            out_name = output_filename
            count = 0

            if file_type == "xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"

                for sn in wb.sheetnames:
                    ws = wb[sn]
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value is None:
                                continue
                            if isinstance(cell.value, str) and find_text in cell.value:
                                cell.value = cell.value.replace(find_text, replace_with)
                                count += 1
                            elif not isinstance(cell.value, str):
                                sval = str(cell.value)
                                if find_text in sval:
                                    cell.value = replace_with
                                    count += 1

                wb.save(out)
                wb.close()

            elif file_type == "xls":
                # Convert .xls to .xlsx, then apply same replace logic
                file_data = _xls_to_xlsx(file_data)
                file_type = "xlsx"
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.xlsx"
                import openpyxl
                wb = openpyxl.load_workbook(io.BytesIO(file_data))

                for sn in wb.sheetnames:
                    ws = wb[sn]
                    for row in ws.iter_rows():
                        for cell in row:
                            if cell.value is None:
                                continue
                            if isinstance(cell.value, str) and find_text in cell.value:
                                cell.value = cell.value.replace(find_text, replace_with)
                                count += 1
                            elif not isinstance(cell.value, str):
                                sval = str(cell.value)
                                if find_text in sval:
                                    cell.value = replace_with
                                    count += 1

                wb.save(out)
                wb.close()

            elif file_type == "docx":
                from docx import Document
                doc = Document(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.docx"

                for para in doc.paragraphs:
                    if find_text in para.text:
                        # Preserve formatting: replace in runs
                        full_text = para.text
                        if find_text in full_text:
                            new_text = full_text.replace(find_text, replace_with)
                            # Clear all runs and set new text in first run
                            if para.runs:
                                para.runs[0].text = new_text
                                for run in para.runs[1:]:
                                    run.text = ""
                                count += 1
                            else:
                                para.text = new_text
                                count += 1

                # Also replace in tables
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if find_text in cell.text:
                                for para in cell.paragraphs:
                                    if find_text in para.text:
                                        if para.runs:
                                            para.runs[0].text = para.text.replace(find_text, replace_with)
                                            for run in para.runs[1:]:
                                                run.text = ""
                                        else:
                                            para.text = para.text.replace(find_text, replace_with)
                                        count += 1

                doc.save(out)

            elif file_type == "pptx":
                from pptx import Presentation
                prs = Presentation(io.BytesIO(file_data))
                if not out_name:
                    out_name = os.path.splitext(filename)[0] + "_edited.pptx"

                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame"):
                            for para in shape.text_frame.paragraphs:
                                if find_text in para.text:
                                    if para.runs:
                                        para.runs[0].text = para.text.replace(find_text, replace_with)
                                        for run in para.runs[1:]:
                                            run.text = ""
                                    else:
                                        para.text = para.text.replace(find_text, replace_with)
                                    count += 1

                prs.save(out)

            else:
                return json.dumps({"error": f"Unsupported type: {file_type}"})

            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name)
            if url:
                return f"[{name}]({url})\n\nReplaced '{find_text}' with '{replace_with}' in {count} place(s), preserving all formatting."
            return json.dumps({"error": "Could not save file"})

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

    # -----------------------------------------------------------------
    # CREATE NEW FILE
    # -----------------------------------------------------------------
    async def create_file(
        self,
        file_type: str,
        content: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Create a new Office file from scratch.

        For xlsx: content is CSV with headers on first line.
        For docx: content is plain text (one paragraph per line).
        For pptx: each line defines a slide. Use "---" as separator between slides.

        Args:
            file_type: 'xlsx', 'docx', or 'pptx'
            content: Content specification
            output_filename: Output filename
        """
        try:
            ftype = file_type.lower().replace(".", "")
            if ftype not in ("xlsx", "docx", "pptx"):
                return json.dumps({"error": f"Unsupported type: {file_type}. Use xlsx, docx, or pptx."})

            out_name = output_filename or f"document.{ftype}"
            out = io.BytesIO()

            if ftype == "xlsx":
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                wb = openpyxl.Workbook()
                ws = wb.active

                import csv as _csv_mod
                reader = _csv_mod.reader(io.StringIO(content))
                rows = list(reader)

                if rows:
                    # First row = headers (styled)
                    for j, h in enumerate(rows[0], 1):
                        c = ws.cell(row=1, column=j, value=h.strip())
                        c.font = Font(bold=True, color="FFFFFF", size=11)
                        c.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                        c.alignment = Alignment(horizontal="center")

                    # Data rows
                    for i, rd in enumerate(rows[1:], 2):
                        for j, v in enumerate(rd, 1):
                            v = v.strip()
                            c = ws.cell(row=i, column=j)
                            if v == '':
                                c.value = None
                            else:
                                try:
                                    c.value = int(v)
                                except ValueError:
                                    try:
                                        c.value = float(v)
                                    except ValueError:
                                        if v.lower() == 'true':
                                            c.value = True
                                        elif v.lower() == 'false':
                                            c.value = False
                                        else:
                                            c.value = v
                            c.alignment = Alignment(
                                horizontal="center" if isinstance(c.value, (int, float)) else "left"
                            )

                    # Auto-fit columns
                    for col in ws.columns:
                        mx = max((len(str(c.value or "")) for c in col), default=5)
                        ws.column_dimensions[col[0].column_letter].width = min(mx + 3, 50)

                wb.save(out)
                wb.close()

            elif ftype == "docx":
                from docx import Document
                from docx.shared import Pt
                doc = Document()
                for line in content.split("\n"):
                    doc.add_paragraph(line)
                doc.save(out)

            elif ftype == "pptx":
                from pptx import Presentation
                from pptx.util import Inches
                prs = Presentation()
                slide_specs = re.split(r'\n---\n|\r\n---\r\n|\n---\n', content)
                blank_layout = prs.slide_layouts[6]

                for spec in slide_specs:
                    spec = spec.strip()
                    if not spec:
                        continue
                    lines = spec.split("\n")
                    title = lines[0].strip()
                    body = "\n".join(lines[1:]).strip() if len(lines) > 1 else ""

                    slide = prs.slides.add_slide(blank_layout)
                    txBox = slide.shapes.add_textbox(
                        Inches(0.5), Inches(0.3), Inches(9), Inches(1)
                    )
                    tf = txBox.text_frame
                    tf.text = title
                    p = tf.paragraphs[0]
                    p.font.size = Inches(0.6)
                    p.font.bold = True

                    if body:
                        txBox2 = slide.shapes.add_textbox(
                            Inches(0.5), Inches(1.5), Inches(9), Inches(5.5)
                        )
                        tf2 = txBox2.text_frame
                        tf2.text = body
                        for para in tf2.paragraphs:
                            para.font.size = Inches(0.3)

                prs.save(out)

            out.seek(0)
            url, name = await self._save_and_link(out.read(), out_name)
            if url:
                return f"[{name}]({url})\n\nCreated new {ftype.upper()} file."
            return json.dumps({"error": "Could not save file"})

        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})


    async def tracked_change(self, file_id: str, change_type: str, content: str, author: str = "Reviewer", paragraph_index: int = -1, output_filename: str = "", __user__=None, __request__=None) -> str:
        """Apply tracked changes (redlines) to a Word document with custom author name.
    
        change_type: replace (use old_text|||new_text), insert (append text with redline), delete (mark paragraph for deletion)
        author: Name shown in Word's Track Changes (e.g., "Sergio Pedro")
        """
        try:
            import sqlite3 as s3
            conn2 = s3.connect(r"C:\\Users\\Administrator\\AppData\\Roaming\\open-webui\\data\\webui.db")
            row = conn2.execute("SELECT filename, meta FROM file WHERE id=?", (file_id,)).fetchone()
            if not row:
                row = conn2.execute("SELECT filename, meta FROM file WHERE filename LIKE ?", (f"%{file_id}%",)).fetchone()
            if not row:
                conn2.close()
                return json.dumps({"error": "File not found"})
            filename = row[0]
            meta = json.loads(row[1]) if row[1] else {}
            fp = meta.get("path", file_id)
            if not os.path.exists(fp):
                fp = os.path.join(os.environ.get("APPDATA",""), "open-webui", "data", "uploads", os.path.basename(fp))
            if not os.path.exists(fp):
                conn2.close()
                return json.dumps({"error": "File not found on disk"})
            with open(fp, "rb") as f:
                data = f.read()
            conn2.close()
    
            from docx import Document
            from docx_revisions import RevisionParagraph
            doc = Document(io.BytesIO(data))
            out_name = output_filename or filename
            results = []
    
            if change_type == "replace":
                parts = content.split("|||", 1)
                if len(parts) != 2:
                    return json.dumps({"error": "Format: old_text|||new_text"})
                find_t, replace_t = parts
                for i, p in enumerate(doc.paragraphs):
                    if paragraph_index >= 0 and i != paragraph_index:
                        continue
                    if find_t in p.text:
                        rp = RevisionParagraph.from_paragraph(p)
                        cnt = rp.replace_tracked(find_t, replace_t, author=author)
                        results.append(f"Para {i}: {cnt} replacements")
            elif change_type == "insert":
                p = doc.add_paragraph()
                rp = RevisionParagraph.from_paragraph(p)
                rp.add_tracked_insertion(content, author=author)
                results.append("Inserted tracked text")
            elif change_type == "delete":
                idx = int(content) if content.isdigit() else paragraph_index
                if idx >= 0 and idx < len(doc.paragraphs):
                    p = doc.paragraphs[idx]
                    rp = RevisionParagraph.from_paragraph(p)
                    rp.add_tracked_deletion(0, len(p.text), author=author)
                    results.append(f"Marked para {idx} for deletion")
    
            out = io.BytesIO()
            doc.save(out)
            out.seek(0)
            url, name = self._save_and_link(out.read(), out_name)
            if url:
                return f"[{name}]({url})\n\nTracked changes by '{author}':\n" + "\n".join(results)
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})
    
    async def manage_revisions(self, file_id: str, action: str, output_filename: str = "", __user__=None, __request__=None) -> str:
        """List, accept_all or reject_all tracked changes in a Word document."""
        try:
            import sqlite3 as s3
            conn2 = s3.connect(r"C:\\Users\\Administrator\\AppData\\Roaming\\open-webui\\data\\webui.db")
            row = conn2.execute("SELECT filename, meta FROM file WHERE id=?", (file_id,)).fetchone()
            if not row:
                row = conn2.execute("SELECT filename, meta FROM file WHERE filename LIKE ?", (f"%{file_id}%",)).fetchone()
            if not row:
                conn2.close()
                return json.dumps({"error": "File not found"})
            filename = row[0]
            meta = json.loads(row[1]) if row[1] else {}
            fp = meta.get("path", file_id)
            if not os.path.exists(fp):
                fp = os.path.join(os.environ.get("APPDATA",""), "open-webui", "data", "uploads", os.path.basename(fp))
            with open(fp, "rb") as f:
                data = f.read()
            conn2.close()
    
            from docx_revisions import RevisionDocument
    
            if action == "list":
                rdoc = RevisionDocument(io.BytesIO(data))
                revs = []
                for para in rdoc.paragraphs:
                    try:
                        rp = RevisionParagraph.from_paragraph(para)
                        if rp.has_track_changes:
                            for ins in rp.insertions:
                                revs.append({"type": "insertion", "author": ins.author, "text": ins.text[:100]})
                            for d in rp.deletions:
                                revs.append({"type": "deletion", "author": d.author, "text": d.text[:100]})
                    except Exception:
                        pass
                return json.dumps({"revisions": revs, "count": len(revs)}, indent=2)
    
            out_name = output_filename or filename
            rdoc = RevisionDocument(io.BytesIO(data))
            if action == "accept_all":
                rdoc.accept_all()
                msg = "All track changes accepted"
            elif action == "reject_all":
                rdoc.reject_all()
                msg = "All track changes rejected"
            else:
                return json.dumps({"error": f"Unknown action: {action}"})
    
            out = io.BytesIO()
            rdoc.save(out)
            out.seek(0)
            url, name = self._save_and_link(out.read(), out_name)
            if url:
                return f"[{name}]({url})\n\n{msg}."
            return json.dumps({"error": "Could not save file"})
        except Exception as e:
            return json.dumps({"error": str(e), "traceback": traceback.format_exc()})

