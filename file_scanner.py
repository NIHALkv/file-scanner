"""
File Scanner
Scans a folder (or a whole drive) and saves an Excel list of every file:
File Name | Folder Path | Full Path | Type | Size (KB) | Last Modified

Usage:
    Double-click file_scanner.exe (Windows)   -> scans ALL drives on this computer
    python file_scanner.py                    -> scans ALL drives on this computer
    python file_scanner.py "C:\\Users\\me"     -> scan only that folder
"""

import os
import sys
import time
from datetime import datetime

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

EXCEL_ROW_LIMIT = 1_000_000  # Excel max is 1,048,576 rows per sheet; new sheet after this
HEADERS = ["File Name", "Folder Path", "Full Path", "Type", "Size (KB)", "Last Modified"]

# Skipped wherever they appear: system junk, not useful to list
SKIP_ANYWHERE = {
    "$recycle.bin", "system volume information", "appdata",
    "node_modules", ".git", "__pycache__",
}
# Skipped only at the top of a drive: operating-system folders
SKIP_AT_DRIVE_ROOT = {
    "windows", "programdata", "$windows.~bt", "$windows.~ws", "recovery",
    "proc", "sys", "dev", "run", "snap", "system", "library", "private", "volumes",
}


def folders_to_scan():
    """Default: every drive on this computer. A folder passed on the command line overrides it."""
    if len(sys.argv) > 1:
        return [sys.argv[1]]
    if sys.platform.startswith("win"):
        import string
        return [f"{letter}:\\" for letter in string.ascii_uppercase if os.path.exists(f"{letter}:\\")]
    return ["/"]  # macOS / Linux: whole file system


def output_path():
    """Save to Desktop if it exists, else the current folder."""
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    folder = desktop if os.path.isdir(desktop) else os.getcwd()
    return os.path.join(folder, f"File_List_{stamp}.xlsx")


def scan(root_folder):
    """Yield one row per file. Skips folders it cannot open instead of crashing."""
    root_norm = os.path.normpath(root_folder)
    is_drive_root = root_norm == os.path.normpath(os.path.splitdrive(root_norm)[0] + os.sep)
    for current, dirs, files in os.walk(root_folder, onerror=lambda e: None):
        at_root = is_drive_root and os.path.normpath(current) == root_norm
        dirs[:] = [
            d for d in dirs
            if d.lower() not in SKIP_ANYWHERE
            and not d.startswith(".")
            and not (at_root and d.lower() in SKIP_AT_DRIVE_ROOT)
        ]
        for name in files:
            full = os.path.join(current, name)
            try:
                st = os.stat(full)
                size_kb = round(st.st_size / 1024, 1)
                modified = datetime.fromtimestamp(st.st_mtime).replace(microsecond=0)
            except OSError:
                size_kb, modified = None, None
            ext = os.path.splitext(name)[1].lower().lstrip(".") or "(none)"
            yield [name, current, full, ext, size_kb, modified]


def new_sheet(wb, index):
    ws = wb.create_sheet(title="Files" if index == 1 else f"Files {index}")
    bold, fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
    header = []
    for h in HEADERS:
        c = WriteOnlyCell(ws, value=h)
        c.font, c.fill = bold, fill
        header.append(c)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 80
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 20
    ws.freeze_panes = "A2"
    ws.append(header)
    return ws


def main():
    folders = [f for f in folders_to_scan() if os.path.isdir(f)]
    if not folders:
        print("Nothing to scan. Exiting.")
        if getattr(sys, "frozen", False):
            input("Press Enter to close...")
        return

    out = output_path()
    print("Scanning all files on: " + ", ".join(folders))
    print("This can take several minutes for a full computer...\n")

    wb = Workbook(write_only=True)
    sheet_no, rows_in_sheet, total = 1, 0, 0
    ws = new_sheet(wb, sheet_no)
    start = time.time()

    for folder in folders:
        print(f"-> {folder}")
        for row in scan(folder):
            if rows_in_sheet >= EXCEL_ROW_LIMIT:
                sheet_no += 1
                ws, rows_in_sheet = new_sheet(wb, sheet_no), 0
            ws.append(row)
            rows_in_sheet += 1
            total += 1
            if total % 5000 == 0:
                print(f"  {total:,} files found...")

    wb.save(out)
    secs = round(time.time() - start, 1)
    print(f"\nDone. {total:,} files listed in {secs}s.")
    print(f"Excel saved to: {out}")

    # Open the Excel automatically on Windows/macOS
    try:
        if sys.platform.startswith("win"):
            os.startfile(out)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f'open "{out}"')
    except Exception:
        pass
    if getattr(sys, "frozen", False):  # keep the window open when run as .exe
        input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
