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

## GitHub auto build

This repo includes a GitHub Actions workflow that builds a Windows `.exe` automatically on every push to `main`.

To download the built file:

1. Open your repo on GitHub.
2. Go to **Actions**.
3. Open the latest **Build Windows EXE** run.
4. Download the `StudentSnip-windows-exe` artifact.

## GitHub release build

This repo also includes a release workflow that publishes the `.exe` to a GitHub Release whenever you push a version tag.

Example:

```bash
git tag v1.0.0
git push origin v1.0.0
```

After that, open **Releases** in GitHub and download `StudentSnip.exe` directly from the new release.

## App version display

The app title bar and main window show the current version from `version.txt`.

To bump version for a new release:

1. Update `version.txt` (for example `1.0.1`).
2. Commit and push to `main`.
3. Create/push matching git tag (for example `v1.0.1`).

## How to use

1. Launch the app.
2. Click **Start Snip** (or press **Ctrl+N**).
3. Open saved screenshots folder with **Open Picture Folder** (or press **Ctrl+O**).
4. Drag to select an area.
5. In preview, annotate with **Draw**, **Rectangle**, or **Text** (with **Undo** and **Clear** available).
6. Click **Save** or **Cancel**.
7. Saved screenshots go to `Pictures\\StudentSnips` as step files like `step_001_YYYYMMDD_HHMMSS.png`.

Each save is also recorded in `Pictures\\StudentSnips\\capture_log.csv` with step number and timestamp.

## Progress GIF mode (for process documentation)

Use this when students want to show progression for a drawing or 3D model:

1. Click **Place Camera Frame**.
2. Move/resize that frame over the exact area to track (the center is transparent so you can see your work behind it).
3. After each work step, click **Capture Frame**.
4. Click **Export Progress GIF** to open an animated preview.
5. Under the preview, select exactly which frames to include (thumbnail + frame number).
6. Adjust the frame-delay slider to test speed, then click **Save GIF**, **Save MP4**, or **Save WebP**.
7. A loading/progress dialog appears while saving.
8. GIF success popup includes total animation duration in seconds.
9. GIF export is capped at **50MB** (Google Slides-friendly). If needed, the app auto-compresses by reducing scale/colors/frames.

Use **Clear Frames** to reset captured progress frames.

### Short video capture mode

If students want automatic frame capture:

1. Click **Capture Short Video**.
2. Choose FPS (frames per second, max **15 FPS**).
3. Click **Capture Frame** once to start timed capture.
4. Click **Stop Capture** to finish.
5. The same GIF preview opens; its default ms/frame matches the selected FPS timing.
6. Students can then fine-tune ms/frame and save GIF or MP4.

Use **Open Picture Folder** in the app to open the screenshot folder quickly.

Press **Esc** during selection to cancel.
