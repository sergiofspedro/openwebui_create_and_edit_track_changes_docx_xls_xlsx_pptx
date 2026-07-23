# Edit Office Files — Open WebUI Tool

Create, read, edit and export Office files (.docx, .xlsx, .xls, .pptx) directly from Open WebUI chats. Preserves original formatting and styles. Supports Word track changes (redlines) with custom author names.

## Features

| # | Function | Formats | Description |
|---|---|---|---|
| 1 | `read_file` | .xlsx .xls .docx .pptx | Read any Office file and return contents as structured JSON. Detects highlights, bold, italic in DOCX. |
| 2 | `add_content` | .xlsx .xls .docx .pptx | Add new content while preserving ALL original formatting. CSV rows for Excel, text for Word, slides for PowerPoint. |
| 3 | `replace_text` | .xlsx .xls .docx .pptx | Find and replace text across the entire file preserving fonts, styles, and cell formatting. |
| 4 | `create_file` | .xlsx .docx .pptx | Create a brand new Office file from scratch with professional styling. |
| 5 | `tracked_change` 🆕 | .docx | Apply Word track changes (redlines) with custom author name. Supports replace, insert, and delete modes. |
| 6 | `manage_revisions` 🆕 | .docx | List all tracked changes, accept all, or reject all revisions in a Word document. |
| 7 | `merge_pdfs` 🆕 | .pdf | Merge multiple PDFs into one using PyMuPDF. |
| 8 | `split_pdf` 🆕 | .pdf | Split PDF into parts by page count. |
| 9 | `merge_sheets` 🆕 | .xlsx | Merge Excel files preserving styles. |
| 10 | `batch_process` 🆕 | All | Apply operation to multiple files at once. |
| 11 | `auto_backup` 🆕 | - | Timestamped database backup for safety. |
| 12 | `tool_stats` 🆕 | - | Show tool usage dashboard with counts. |

### Track Changes (v1.2.0)

Built on [docx-revisions](https://github.com/balalofernandez/docx-revisions) library. Writes standard OOXML `w:ins` / `w:del` elements — 100% compatible with Microsoft Word.

```python
# Replace text with track changes
tracked_change(file_id, change_type="replace", content="old_text|||new_text", author="Sergio Pedro")

# Insert new text as tracked change
tracked_change(file_id, change_type="insert", content="New paragraph text", author="Reviewer")

# Mark paragraph for deletion
tracked_change(file_id, change_type="delete", content="3", author="Editor")

# List all revisions
manage_revisions(file_id, action="list")

# Accept all changes
manage_revisions(file_id, action="accept_all")
```

**Track changes visibility by program:**

| Program | Track changes visible? |
|---|---|
| Microsoft Word | Yes — 100% native support |
| LibreOffice | Yes — good OOXML revision support |
| Google Docs | Partial — opens .docx but may drop metadata |

### Format Support

| Format | Read | Edit (preserve style) | Create | Notes |
|---|---|---|---|---|
| .xlsx | Yes | Yes | Yes | Full support via openpyxl |
| .xls | Yes | Yes (saves as .xlsx) | — | Legacy format via xlrd |
| .docx | Yes | Yes + Track Changes | Yes | Full support via python-docx + docx-revisions |
| .pptx | Yes | Yes | Yes | Full support via python-pptx |
| .doc | No | No | No | Suggest converting to .docx |
| .ppt | No | No | No | Suggest converting to .pptx |

## Installation

### Method 1: Open WebUI Community
Search for "Edit Office Files" in the Open WebUI Community tools.

### Method 2: Manual Install
1. Download `tool.py` from this repo
2. In Open WebUI: Workspace > Tools > New Tool
3. Paste the code and save
4. Install dependencies:
```bash
pip install openpyxl python-docx python-pptx xlrd docx-revisions
```
5. Start the file server: `python file_server.py` (serves files on port 9000)

### Method 3: Batch Install
Use the Batch Install Plugins tool in Open WebUI pointing to this repo.

## Usage Examples

**Read a file:**
```
"Read this Excel file and show me the data"
"Show me what's in this Word document"
"What slides are in this presentation?"
```

**Add content:**
```
"Add these rows to the Excel keeping the same style:
Name,Age,City
Ana,30,Lisbon"
"Add this paragraph to the end of the Word document"
```

**Replace text:**
```
"Replace 'N/A' with 'Not Available' in this file"
"Change all '2025' to '2026' in the spreadsheet"
```

**Track changes (Word only):**
```
"Replace 'old contract' with 'new contract' in this Word doc as a tracked change, author=Sergio"
"Insert this clause as a redline, author=Legal Team"
"List all tracked changes in this document"
"Accept all revisions"
```

**Create new file:**
```
"Create an Excel with columns Name, Age, City and 5 rows of data"
"Create a PowerPoint with 3 slides about Q3 results"
```

## File Server
The tool saves generated files to a local `exports/` directory and serves them via a simple HTTP server on port 9000. Run:
```bash
python file_server.py
```
Files will be downloadable at `http://localhost:9000/filename.xlsx`

## Dependencies
- `openpyxl` — Excel .xlsx read/write
- `python-docx` — Word .docx read/write
- `python-pptx` — PowerPoint .pptx read/write
- `xlrd` — Legacy Excel .xls read
- `docx-revisions` — Word track changes (redlines)
- `lxml` — XML processing (dependency of docx-revisions)

## License
MIT

## Author
giofsp — [GitHub](https://github.com/sergiofspedro)
