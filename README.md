# Bulk Folder Creator

A small desktop app (Python + Tkinter) for creating many folders and nested
folder structures at once. Type paths by hand or import them from a CSV/Excel
file, point it at a base directory, and click **Create Folders**.

## Features

- Create deeply nested folders in one go (`Client_A/Documents/Invoices`)
- One path per line, or import in bulk from **CSV** / **Excel**
- Smart import: single-column (full paths) **or** multi-column with
  fill-down, so you only type each parent once
- Export your current list back out to CSV/Excel
- Clear, friendly status messages and error reporting

## Run from source

Requires Python 3.

```bash
pip install openpyxl   # optional, enables Excel import/export
python folder_creator.py
```

> `openpyxl` is optional — without it the app still works with CSV files.

## Build a Windows .exe

On a Windows PC with Python installed, just double-click `build_exe.bat`,
or run:

```bash
pip install pyinstaller openpyxl
pyinstaller --onefile --windowed --name "BulkFolderCreator" folder_creator.py
```

The executable is written to `dist/BulkFolderCreator.exe`.

## Import file format

**Single column** — each cell is a complete path:

```
Path
01-Native/Documents/Invoices
02-Exchange/Imports
```

**Multiple columns** — each column is one level deeper; blank cells inherit
the value above them:

| Level 1     | Level 2   | Level 3  |
|-------------|-----------|----------|
| 01-Native   | Documents | Invoices |
|             |           | Receipts |
|             | Contracts |          |
| 02-Exchange | Imports   |          |

## License

Released under the [MIT License](LICENSE).
