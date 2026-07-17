"""
title: Edit Spreadsheet
author: giofsp
author_url: https://github.com/sergiofspedro
description: Create, read, and edit Excel (.xlsx) files with full cell-level editing, style preservation, and value replacement.
version: 1.0.0
requirements: openpyxl
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy
import json, io, traceback, datetime, sys, os, sqlite3, re
from typing import Optional, List, Dict, Any
from io import BytesIO

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


def _resolve_file_path(file_id: str) -> Optional[str]:
    """Read the absolute path for *file_id* directly from the SQLite DB.
    
    This bypasses the broken ``Storage.get_file()`` and
    ``Files.get_file_by_id()`` APIs by querying the database that
    Open WebUI itself writes to.
    """
    try:
        conn = sqlite3.connect(f"file:{_DB_PATH}?mode=ro", uri=True)
        row = conn.execute(
            "SELECT path FROM file WHERE id = ?", (file_id,)
        ).fetchone()
        conn.close()
    except Exception as exc:
        print(f"[xlsx] DB lookup failed for {file_id}: {exc}", file=sys.stderr)
        return None

    if not row or not row[0]:
        print(f"[xlsx] No path for file_id {file_id}", file=sys.stderr)
        return None

    path = row[0]
    # If the stored path is a real local file, return it directly.
    if os.path.isfile(path):
        return path

    # Otherwise try the uploads directory as a fallback.
    # The DB sometimes stores just the basename.
    candidate = os.path.join(_UPLOAD_DIR, os.path.basename(path))
    if os.path.isfile(candidate):
        return candidate

    # Last resort: glob the uploads dir for a matching UUID prefix.
    prefix = file_id.split("-")[0] if "-" in file_id else file_id[:8]
    for name in os.listdir(_UPLOAD_DIR):
        if name.startswith(prefix):
            candidate = os.path.join(_UPLOAD_DIR, name)
            if os.path.isfile(candidate):
                return candidate

    print(f"[xlsx] File not found on disk: {path}", file=sys.stderr)
    return None


def _read_file_bytes(file_id: str) -> Optional[bytes]:
    """Return the raw bytes for an Open WebUI file, or ``None``."""
    path = _resolve_file_path(file_id)
    if path is None:
        return None
    try:
        with open(path, "rb") as fh:
            return fh.read()
    except Exception as exc:
        print(f"[xlsx] Read failed for {path}: {exc}", file=sys.stderr)
        return None


def _cell_value(cell):
    """Extract a JSON-safe value from an openpyxl cell."""
    v = cell.value
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    if isinstance(v, (int, float, str, bool)):
        return v
    return str(v)


def _apply_style(target_cell, ref_cell):
    """Copy style attributes from *ref_cell* to *target_cell*."""
    try:
        target_cell.font = copy(ref_cell.font)
    except Exception:
        pass
    try:
        target_cell.fill = copy(ref_cell.fill)
    except Exception:
        pass
    try:
        target_cell.border = copy(ref_cell.border)
    except Exception:
        pass
    try:
        target_cell.alignment = copy(ref_cell.alignment)
    except Exception:
        pass
    try:
        target_cell.number_format = ref_cell.number_format
    except Exception:
        pass


# =========================================================================
class Tools:
    class Valves:
        pass

    def __init__(self):
        pass

    # -----------------------------------------------------------------
    # Internal: save bytes as an Open WebUI file and return the link.
    # -----------------------------------------------------------------
    async def _save_file(self, file_bytes, filename, __user__=None, __request__=None):
        """Save file to exports directory and return an HTTP download URL. Returns (url, filename) or (None, None)."""
        try:
            import os as _os
            export_dir = r"C:\Users\Administrator\AppData\Local\open-webui\exports"
            _os.makedirs(export_dir, exist_ok=True)
            # Sanitize filename: keep alphanumeric, dots, hyphens, underscores, spaces
            safe = "".join(c for c in filename if c.isalnum() or c in "._- ")
            if not safe:
                safe = "export"
            base_name, ext = _os.path.splitext(safe)
            fname = safe
            counter = 2
            while _os.path.exists(_os.path.join(export_dir, fname)):
                fname = f"{base_name} ({counter}){ext}"
                counter += 1
            fpath = _os.path.join(export_dir, fname)
            with open(fpath, "wb") as f:
                f.write(file_bytes)
            url = f"http://localhost:9000/{fname}"
            return url, filename
        except Exception as e:
            print(f"[xlsx] Export save failed: {e}", file=sys.stderr)
            return None, None
    async def read_spreadsheet(
        self,
        file_id: str,
        max_rows: int = 500,
        __user__=None,
        __request__=None,
    ) -> str:
        """Read an Excel file and return all sheets, headers, and rows as JSON.

        The *file_id* is the Open WebUI file UUID visible in the upload URL
        (e.g. ``/api/v1/files/<id>/content``).  Returns a JSON object with
        ``sheets[*].headers`` and ``sheets[*].rows``.
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

            wb = openpyxl.load_workbook(io.BytesIO(file_data), data_only=True)
            result: Dict[str, Any] = {
                "file_id": file_id,
                "sheets": [],
            }

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
                for ri, row in enumerate(
                    ws.iter_rows(min_row=1, max_row=max_r), 1
                ):
                    rd = [_cell_value(c) for c in row]
                    if ri == 1:
                        sheet["headers"] = [
                            str(v) if v is not None else "" for v in rd
                        ]
                    else:
                        sheet["rows"].append(rd)

                result["sheets"].append(sheet)

            wb.close()
            return json.dumps(result, indent=2, default=str)
        except Exception as exc:
            return json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })

    # -------------------- CREATE --------------------------------------
    async def create_spreadsheet(
        self,
        headers: List[str],
        rows: List[list],
        sheet_name: str = "Sheet1",
        output_filename: str = "spreadsheet.xlsx",
        __user__=None,
        __request__=None,
    ) -> str:
        """Create a brand-new Excel file with styled headers and data rows."""
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet_name

            # -- Header row
            for j, h in enumerate(headers, 1):
                c = ws.cell(row=1, column=j, value=h)
                c.font = Font(bold=True, color="FFFFFF", size=11)
                c.fill = PatternFill(
                    start_color="4472C4",
                    end_color="4472C4",
                    fill_type="solid",
                )
                c.alignment = Alignment(horizontal="center")

            # -- Data rows
            for i, rd in enumerate(rows, 2):
                for j, v in enumerate(rd, 1):
                    c = ws.cell(row=i, column=j, value=v)
                    c.alignment = Alignment(
                        horizontal=(
                            "center" if isinstance(v, (int, float)) else "left"
                        )
                    )

            # -- Auto-fit columns
            for col in ws.columns:
                mx = max(
                    (len(str(c.value or "")) for c in col), default=5
                )
                ws.column_dimensions[col[0].column_letter].width = min(
                    mx + 3, 50
                )

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            url, name = await self._save_file(
                out.read(), output_filename, __user__, __request__
            )
            if url:
                return (
                    f"[{name}]({url})\n\n"
                    f"Created spreadsheet with {len(headers)} columns "
                    f"and {len(rows)} rows."
                )
            return json.dumps({"error": "File save failed — browser download could not be triggered. Check server logs for [xlsx] entries."})
        except Exception as exc:
            return json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })

    # -------------------- ADD ROWS ------------------------------------
    async def add_rows_to_sheet(
        self,
        file_id: str,
        sheet_name: str,
        data: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Add rows from CSV text to a sheet preserving formatting. Each CSV line becomes one data row. Returns download URL or error. Use the download URL to get the updated file."""
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({
                    "error": f"Could not read file {file_id}"
                })

            wb = openpyxl.load_workbook(io.BytesIO(file_data))
            if sheet_name not in wb.sheetnames:
                avail = wb.sheetnames
                wb.close()
                return json.dumps({
                    "error": (
                        f"Sheet '{sheet_name}' not found. "
                        f"Available: {avail}"
                    )
                })

            ws = wb[sheet_name]

            # Grab reference styles from the last populated row
            ref: Dict[int, Any] = {}
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

            # Parse CSV data string into rows
            import csv as _csv_mod
            reader = _csv_mod.reader(io.StringIO(data))
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
                return json.dumps({"error": "No rows provided in CSV data"})

            start = (ws.max_row or 0) + 1
            for i, rd in enumerate(parsed_rows):
                for j, v in enumerate(rd, 1):
                    cell = ws.cell(row=start + i, column=j)
                    if j in ref:
                        _apply_style(cell, ws.cell(row=ws.max_row, column=j))
                    cell.value = v

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            fname = output_filename or "edited_spreadsheet.xlsx"
            url, name = await self._save_file(out.read(), fname, __user__, __request__)

            if url:
                return (
                    f"[{name}]({url})\n\n"
                    f"Added {len(parsed_rows)} rows to '{sheet_name}' "
                    f"preserving original formatting."
                )
            return json.dumps({"error": "File save failed — browser download could not be triggered. Check server logs for [xlsx] entries."})
        except Exception as exc:
            return json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })

    # -------------------- EDIT CELLS ----------------------------------
    async def edit_cells(
        self,
        file_id: str,
        sheet_name: str,
        edits: List[Dict[str, Any]],
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Edit specific cells in an existing spreadsheet.

        Each entry in *edits* is a dict with keys:
          - ``row`` (int, 1-based): row number
          - ``col`` (int, 1-based): column number
          - ``value`` (any): new cell value

        Example edits::

            [
                {"row": 2, "col": 3, "value": 42},
                {"row": 5, "col": 1, "value": "Updated text"},
            ]
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({
                    "error": f"Could not read file {file_id}"
                })

            wb = openpyxl.load_workbook(io.BytesIO(file_data))
            if sheet_name not in wb.sheetnames:
                avail = wb.sheetnames
                wb.close()
                return json.dumps({
                    "error": (
                        f"Sheet '{sheet_name}' not found. "
                        f"Available: {avail}"
                    )
                })

            ws = wb[sheet_name]
            applied = 0
            errors = []

            for edit in edits:
                try:
                    r = int(edit["row"])
                    c = int(edit["col"])
                    v = edit["value"]
                    ws.cell(row=r, column=c).value = v
                    applied += 1
                except (KeyError, ValueError, TypeError) as exc:
                    errors.append(f"edit {edit}: {exc}")

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            fname = output_filename or "edited_spreadsheet.xlsx"
            url, name = await self._save_file(out.read(), fname, __user__, __request__)

            if url:
                msg = (
                    f"[{name}]({url})\n\n"
                    f"Edited {applied} cell(s) in '{sheet_name}'."
                )
                if errors:
                    msg += f"\n\nWarnings: {'; '.join(errors)}"
                return msg
            return json.dumps({"error": "File save failed — browser download could not be triggered. Check server logs for [xlsx] entries."})
        except Exception as exc:
            return json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })

    # -------------------- REPLACE VALUES ------------------------------
    async def replace_values(
        self,
        file_id: str,
        sheet_name: str,
        find_value: str,
        replace_value: str,
        match_case: bool = False,
        search_all_sheets: bool = False,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Find and replace values across a sheet (or all sheets).

        Works on string cell values. Numeric cells are compared after
        ``str()`` conversion.  Returns the number of replacements made.
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({
                    "error": f"Could not read file {file_id}"
                })

            wb = openpyxl.load_workbook(io.BytesIO(file_data))
            sheets = (
                wb.sheetnames
                if search_all_sheets
                else [sheet_name]
            )
            total_replaced = 0

            for sn in sheets:
                if sn not in wb.sheetnames:
                    continue
                ws = wb[sn]
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is None:
                            continue
                        current = str(cell.value)
                        target = find_value if match_case else find_value.lower()
                        source = current if match_case else current.lower()
                        if isinstance(cell.value, str):
                            if target in source:
                                if match_case:
                                    new = current.replace(
                                        find_value, replace_value
                                    )
                                else:
                                    new = re.sub(
                                        re.escape(find_value),
                                        replace_value,
                                        current,
                                        flags=re.IGNORECASE,
                                    )
                                cell.value = new
                                total_replaced += 1
                        else:
                            # Check numeric / other types via str() cast
                            if target in source:
                                cell.value = replace_value
                                total_replaced += 1

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            fname = output_filename or "edited_spreadsheet.xlsx"
            url, name = await self._save_file(out.read(), fname, __user__, __request__)

            if url:
                scope = (
                    f"{len(sheets)} sheet(s)"
                    if search_all_sheets
                    else f"'{sheet_name}'"
                )
                return (
                    f"[{name}]({url})\n\n"
                    f"Replaced {total_replaced} occurrence(s) of "
                    f"'{find_value}' with '{replace_value}' in {scope}."
                )
            return json.dumps({"error": "File save failed — browser download could not be triggered. Check server logs for [xlsx] entries."})
        except Exception as exc:
            return json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })

    # -------------------- DELETE ROWS ---------------------------------
    async def delete_rows(
        self,
        file_id: str,
        sheet_name: str,
        row_numbers: List[int],
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Delete specific rows (1-based) from a sheet.

        Rows are deleted from the bottom up so that indices remain valid.
        """
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({
                    "error": f"Could not read file {file_id}"
                })

            wb = openpyxl.load_workbook(io.BytesIO(file_data))
            if sheet_name not in wb.sheetnames:
                avail = wb.sheetnames
                wb.close()
                return json.dumps({
                    "error": (
                        f"Sheet '{sheet_name}' not found. "
                        f"Available: {avail}"
                    )
                })

            ws = wb[sheet_name]
            sorted_rows = sorted(set(row_numbers), reverse=True)
            deleted = 0
            for rn in sorted_rows:
                if 1 <= rn <= (ws.max_row or 0):
                    ws.delete_rows(rn, 1)
                    deleted += 1

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            fname = output_filename or "edited_spreadsheet.xlsx"
            url, name = await self._save_file(out.read(), fname, __user__, __request__)

            if url:
                return (
                    f"[{name}]({url})\n\n"
                    f"Deleted {deleted} row(s) from '{sheet_name}'."
                )
            return json.dumps({"error": "File save failed — browser download could not be triggered. Check server logs for [xlsx] entries."})
        except Exception as exc:
            return json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })

    # -------------------- COPY SHEET ----------------------------------
    async def copy_sheet(
        self,
        file_id: str,
        source_sheet: str,
        new_sheet_name: str,
        output_filename: str = "",
        __user__=None,
        __request__=None,
    ) -> str:
        """Duplicate a sheet within the same workbook."""
        try:
            file_data = _read_file_bytes(file_id)
            if file_data is None:
                return json.dumps({
                    "error": f"Could not read file {file_id}"
                })

            wb = openpyxl.load_workbook(io.BytesIO(file_data))
            if source_sheet not in wb.sheetnames:
                avail = wb.sheetnames
                wb.close()
                return json.dumps({
                    "error": (
                        f"Sheet '{source_sheet}' not found. "
                        f"Available: {avail}"
                    )
                })

            wb.copy_worksheet(wb[source_sheet]).title = new_sheet_name

            out = io.BytesIO()
            wb.save(out)
            wb.close()
            out.seek(0)

            fname = output_filename or "edited_spreadsheet.xlsx"
            url, name = await self._save_file(out.read(), fname, __user__, __request__)

            if url:
                return (
                    f"[{name}]({url})\n\n"
                    f"Copied '{source_sheet}' → '{new_sheet_name}'."
                )
            return json.dumps({"error": "File save failed — browser download could not be triggered. Check server logs for [xlsx] entries."})
        except Exception as exc:
            return json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            })
