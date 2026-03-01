# Student Screenshot Tool (Windows)

A simple snipping tool for students to capture part of the screen and save as PNG.

## Run from Python

1. Open terminal in this folder.
2. Install dependencies:
   ```bash
   py -m pip install -r requirements.txt
   ```
3. Start app:
   ```bash
   py app.py
   ```

## Build .exe

Run:
```bat
build_exe.bat
```

The script creates a local `.venv` automatically, installs dependencies, generates an app icon, and builds in that isolated environment.

After build finishes, the executable is at:

`%LOCALAPPDATA%\\StudentSnipBuild\\dist\\StudentSnip.exe`

A copy is also placed in the project folder:

`dist\\StudentSnip.exe`

## How to use

1. Launch the app.
2. Click **Start Snip** (or press **Ctrl+N**).
3. Drag to select an area.
4. In preview, annotate with **Draw**, **Rectangle**, or **Text** (with **Undo** and **Clear** available).
5. Click **Save** or **Cancel**.
6. Saved screenshots go to `Pictures\\StudentSnips` as step files like `step_001_YYYYMMDD_HHMMSS.png`.

Each save is also recorded in `Pictures\\StudentSnips\\capture_log.csv` with step number and timestamp.

Use **Open Picture Folder** in the app to open the screenshot folder quickly.

Press **Esc** during selection to cancel.
