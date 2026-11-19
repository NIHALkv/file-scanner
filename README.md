# 📂 File Scanner

Scan any folder (or a whole drive) and get an **Excel list of every file** with its path.

| File Name | Folder Path | Full Path | Type | Size (KB) | Last Modified |
|---|---|---|---|---|---|

## Use it (Windows, no install)
1. Go to **[Releases](../../releases/latest)** and download `file_scanner.exe`.
2. Double-click it. If Windows shows "Windows protected your PC", click **More info → Run anyway** (the app is unsigned, not harmful).
3. Pick a folder or drive (e.g. `C:\`, or `Documents`).
4. Wait. The Excel file opens by itself and is saved on your **Desktop** as `File_List_<date>.xlsx`.

## Use it with Python (Windows / Mac / Linux)
```bash
pip install openpyxl
python file_scanner.py                 # opens a folder picker
python file_scanner.py "D:\Projects"   # or scan a folder directly
```

## Good to know
- Skips system and noisy folders (Windows, AppData, Recycle Bin, node_modules, .git, hidden folders).
- Folders it has no permission to open are skipped silently.
- More than 1 million files? They continue on extra sheets (`Files 2`, `Files 3`, ...).
- Nothing leaves your computer. The tool only reads file names, not file contents.

## Ideas for next versions
- [ ] Filter by file type (e.g. only PDFs / Excel files)
- [ ] Find duplicate files
- [ ] Find the largest files taking up space

Built by [Nihal K V](https://www.linkedin.com/in/nihalkvind/) with Claude.
