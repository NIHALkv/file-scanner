# 📂 File Scanner

Scans **every drive on your computer** and gives you an **Excel list of every file** with its path.

| File Name | Folder Path | Full Path | Type | Size (KB) | Last Modified |
|---|---|---|---|---|---|

## Use it (Windows, no install)
1. Go to **[Releases](../../releases/latest)** and download `file_scanner.exe`.
   - If Chrome says *"isn't commonly downloaded"*: click **⋯ → Keep → Show more → Keep anyway**.
2. Double-click it. If Windows shows "Windows protected your PC", click **More info → Run anyway** (the app is unsigned, not harmful).
3. It scans all drives (C:, D:, ...) automatically. A full computer can take a few minutes.
4. The Excel file opens by itself and is saved on your **Desktop** as `File_List_<date>.xlsx`.

## Use it with Python (Windows / Mac / Linux)
```bash
pip install openpyxl
python file_scanner.py                 # scans all drives
python file_scanner.py "D:\Projects"   # or scan only one folder
```

## Good to know
- Skips operating-system and junk folders (C:\Windows, ProgramData, AppData, Recycle Bin, node_modules, .git, hidden folders), so the list shows *your* files.
- Folders it has no permission to open are skipped silently.
- More than 1 million files? They continue on extra sheets (`Files 2`, `Files 3`, ...).
- Nothing leaves your computer. The tool only reads file names, not file contents.

## Ideas for next versions
- [ ] Filter by file type (e.g. only PDFs / Excel files)
- [ ] Find duplicate files
- [ ] Find the largest files taking up space

Built by [Nihal K V](https://www.linkedin.com/in/nihalkvind/) with Claude.
