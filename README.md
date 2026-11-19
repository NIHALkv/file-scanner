# 📂 File Scanner

Scans **every drive on your computer** and gives you an **Excel list of every file** with its name and path.

**[⬇️ Download file_scanner.exe (Windows)](../../releases/latest/download/file_scanner.exe)**: no install, no Python needed.

---

## What you get
An Excel file on your Desktop called `File_List_<date>_<time>.xlsx`:

| File Name | Folder Path | Full Path | Type | Size (KB) | Last Modified |
|---|---|---|---|---|---|
| Budget_2026.xlsx | C:\Users\Nihal\Documents | C:\Users\Nihal\Documents\Budget_2026.xlsx | xlsx | 48.2 | 2026-11-02 14:10 |
| Resume.pdf | D:\Jobs | D:\Jobs\Resume.pdf | pdf | 210.5 | 2026-10-28 09:41 |

The header row stays frozen and the columns are sized to fit, so you can use **Filter** straight away to sort or search.

## Useful for
- Finding where a file went ("which folder did I save that in?")
- Listing all PDFs / Excel / photos on a laptop
- Spotting old or large files to clean up
- Handing over or auditing the files on a work machine

---

## How to use (Windows)
1. **Download** [file_scanner.exe](../../releases/latest/download/file_scanner.exe).
   - If Chrome says *"isn't commonly downloaded"*, click **⋯ → Keep → Show more → Keep anyway**.
2. **Double-click** `file_scanner.exe`.
   - If Windows shows *"Windows protected your PC"*, click **More info → Run anyway**.
3. A black window opens and shows progress: `25,000 files found...`
4. When it finishes, the **Excel file opens by itself**. It is also saved on your **Desktop**.
5. Press **Enter** to close the black window.

⏱️ A full laptop usually takes **1–5 minutes**, depending on how many files you have.

### Why the security warnings?
The app is new and not "code-signed" (signing costs money), so Chrome and Windows warn about it the way they do for any new app.
The app is built automatically by GitHub from the code in this repo, and you can read every line in [`file_scanner.py`](file_scanner.py).

---

## Privacy
- **Nothing leaves your computer.** No internet connection is used.
- It reads only file **names, sizes and dates**, never what is inside the files.
- The Excel file is saved only on your Desktop.

## What it skips
To keep the list about *your* files, it skips:
- Windows system folders at the top of a drive: `C:\Windows`, `ProgramData`, `Recovery`
- App and junk folders anywhere: `AppData`, `$Recycle.Bin`, `System Volume Information`, `node_modules`, `.git`
- Hidden folders (names starting with `.`)
- Folders Windows won't let it open (skipped silently)

More than ~1 million files? The list continues on extra sheets: `Files 2`, `Files 3`, ...

---

## Troubleshooting
| Problem | Fix |
|---|---|
| Window closes instantly | Right-click the exe → **Run as administrator**, or check your antivirus didn't block it. |
| Excel didn't open | Look on your Desktop for `File_List_...xlsx`. |
| Some folders are missing | They are system or protected folders (see *What it skips*). |
| It takes very long | Normal for big or external drives. Let it run. |

---

## For developers / Python users
Works on Windows, macOS and Linux.
```bash
pip install -r requirements.txt
python file_scanner.py                 # scan all drives
python file_scanner.py "D:\Projects"   # scan only one folder
```

### How the .exe is made
Every push to `main` triggers [`.github/workflows/build-exe.yml`](.github/workflows/build-exe.yml), which:
1. builds `file_scanner.exe` on a Windows machine with PyInstaller, and
2. publishes it, with `README.txt`, to the **[latest release](../../releases/latest)**.

So the download link always gives the newest version.

### Files in this repo
| File | What it is |
|---|---|
| `file_scanner.py` | The tool (about 150 lines of Python) |
| `requirements.txt` | Python library needed (`openpyxl`) |
| `.github/workflows/build-exe.yml` | Auto-builds the Windows .exe |
| `README.txt` | Short plain-text guide shipped next to the .exe |

---

## Roadmap
- [ ] Filter by file type (e.g. only PDFs / Excel files)
- [ ] Find duplicate files
- [ ] Show the largest files taking up space

## Feedback
Found a bug or want a feature? Open an [Issue](../../issues) or message me on [LinkedIn](https://www.linkedin.com/in/nihalkvind/).

Built by **Nihal K V** with Claude.
