# Edit Spreadsheet — Open WebUI Tool

Create, read, edit and export Excel (.xlsx) files directly from Open WebUI chats. Preserves original formatting and styles.

## Features

| Function | Description |
|---|---|
| `read_spreadsheet` | Read an Excel file and return all sheets, headers and rows as JSON |
| `create_spreadsheet` | Create a new Excel file from scratch with styled headers and data |
| `add_rows_to_sheet` | Add new rows to a sheet via CSV text, preserving original formatting |
| `edit_cells` | Modify specific cells by (row, col) coordinates |
| `replace_values` | Find-and-replace across a sheet or all sheets |
| `delete_rows` | Remove specific rows by number |
| `copy_sheet` | Duplicate a worksheet within the same workbook |

## Installation

### Method 1: Install from Open WebUI Community
Search for "Edit Spreadsheet" in the Open WebUI Community tools.

### Method 2: Manual Install
1. Download `tool.py` from this repo
2. In Open WebUI: Workspace → Tools → New Tool
3. Paste the code and save
4. Start the file server: `python file_server.py`
5. Files will be available at `http://localhost:9000/`

### Method 3: Batch Install
Use the Batch Install Plugins tool in Open WebUI and point to this repo.

## Usage Examples

**Read an Excel file:**
```
Upload your Excel file and ask:
"Reve este ficheiro Excel e mostra-me os dados"
```

**Add data preserving styles:**
```
"Adiciona estas linhas ao Excel mantendo o mesmo formato:
Nome,Idade,Cidade
Ana,30,Lisboa
Pedro,25,Porto"
```

**Edit specific cells:**
```
"Na sheet 'Vendas', muda a célula B5 para 1500"
```

**Find and replace:**
```
"Substitui todos os 'N/A' por 'Em falta' no Excel"
```

## Requirements
- Open WebUI Desktop or Server
- Python 3.10+
- Packages: `openpyxl`

## File Server
The tool saves generated files to a local `exports/` directory and serves them via a simple HTTP server on port 9000. Run:
```bash
python file_server.py
```
Files will be downloadable at `http://localhost:9000/filename.xlsx`

## License
MIT
