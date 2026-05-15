# GHG QA Tool — Deployment Guide

For the person packaging and distributing the tool to analysts.
Analysts receive `README.txt` instead — this document is for you.

---

## Prerequisites (your Mac/dev machine)

- Python 3.9+ with the project venv set up (`python -m venv .venv`)
- All packages installed: `pip install -r requirements.txt`
- `zip` available (standard on macOS/Linux)

To verify everything works before packaging:

```bash
cd /Users/bear/src/excel
.venv/bin/python tests/test_phase3_integration.py
```

All 35 tests should pass. If any fail, do not ship.

---

## Building the distributable ZIP

```bash
bash build_package.sh          # produces dist/ghg_qa_tool_v1.0.zip
bash build_package.sh 1.1      # bump the version number
```

The ZIP contains exactly what analysts need — no dev files, no tests, no venv:

```
ghg_qa_tool_v1.0/
  run_app.bat         ← analyst double-clicks this every time
  setup_windows.bat   ← analyst runs this once on first install
  README.txt          ← instructions for the analyst
  requirements.txt    ← package list (for IT reference)
  app/
    streamlit_app.py
    column_matcher.py
    diff_engine.py
    annotation_store.py
    export.py
```

---

## Deployment options

### Option A — IT-managed Python (simplest, ~24 KB ZIP)

Requires IT to have already installed Python 3.9+ on the analyst's machine.

1. Build: `bash build_package.sh 1.0`
2. Send `dist/ghg_qa_tool_v1.0.zip` to the analyst
3. Analyst extracts it (Desktop, `C:\Tools\`, anywhere)
4. Analyst runs `setup_windows.bat` once → installs streamlit, pandas, openpyxl
5. Analyst double-clicks `run_app.bat` to start the tool

**Suitable when:** IT controls Python installs, machines have internet access.

---

### Option B — Embedded WinPython (zero-dependency, ~80 MB ZIP)

No admin rights needed. Python is bundled inside the ZIP.

**One-time setup on your machine:**

1. Download WinPython "dot" build (no IDE) from https://winpython.github.io/
   - Choose the `python-3.X.X.amd64` variant (not the full WinPython installer)
   - Tested with Python 3.11 or 3.12 — do not use 3.14 (too new for some packages)

2. Extract the download. Inside you'll find a folder named `python-3.X.X.amd64`.
   Rename it to `python` and place it in the project root:

   ```
   /Users/bear/src/excel/python/   ← must contain python.exe at the top level
   ```

3. Install packages into the embedded Python:

   ```bash
   ./python/python.exe -m pip install streamlit==1.57.0 pandas==3.0.2 openpyxl==3.1.5
   ```

4. Build with the embed flag:

   ```bash
   bash build_package.sh 1.0 --embed
   ```

5. Send the ~80 MB ZIP. Analyst extracts and double-clicks `run_app.bat` — no setup step needed.

**Suitable when:** Analyst laptops are locked down (no admin, no internet, no IT Python).

---

## What happens when the analyst runs the tool

1. `run_app.bat` detects Python (embedded bundle first, then system)
2. Runs: `python -m streamlit run app\streamlit_app.py --server.port 8510`
3. A black command window stays open (the server) — analyst should leave it open
4. The browser opens automatically at `http://localhost:8510`
5. All data stays local — nothing is sent to any server

Port 8510 is used to avoid conflicts with other Streamlit tools. If it conflicts,
edit `run_app.bat` and `run_app.sh` and change `8510` to any free port (e.g. 8520).

---

## Updating the tool

1. Make and test your code changes
2. Run the full test suite: `python tests/test_phase3_integration.py`
3. Bump the version: `bash build_package.sh 1.1`
4. Send the new ZIP to analysts
5. Analysts replace their old folder with the new one
   (annotations from previous sessions are in a temp folder and will not carry over —
   see annotation persistence note below)

---

## Annotation persistence note

Annotations are saved as a sidecar JSON file in a system temp directory
(`%TEMP%\ghg_qa_XXXXXX\` on Windows). This means:

- Annotations **persist** as long as the command window stays open in a session
- Annotations are **lost** if the analyst closes the command window and reopens `run_app.bat`

If persistent annotations across sessions are needed (future enhancement),
the sidecar path should be written next to the Excel files or to a fixed
project folder rather than a temp directory.

---

## File inventory

| File | Purpose |
|---|---|
| `app/streamlit_app.py` | Streamlit UI — all three steps |
| `app/column_matcher.py` | Column header matching engine |
| `app/diff_engine.py` | Core diff logic, severity classification |
| `app/annotation_store.py` | Sidecar JSON annotation persistence |
| `app/export.py` | QA Log Excel writer |
| `run_app.bat` | Windows launcher |
| `run_app.sh` | Mac/Linux launcher |
| `setup_windows.bat` | First-time Windows installer |
| `requirements.txt` | Pinned Python dependencies |
| `build_package.sh` | Builds the distributable ZIP |
| `tests/make_test_data.py` | Generates synthetic test Excel files |
| `tests/test_phase3_integration.py` | 35-test integration suite |
| `README.txt` | Analyst-facing instructions (ships in ZIP) |
