# Pro Headshot Cutout Studio

Photoshop-first batch tool for creating consistent `250x250` transparent PNG headshots with a dark-mode UI.

## What it does

- Scans source folders recursively.
- Preserves input folder structure in output.
- Optimizes huge source images before processing for speed (outside Photoshop, in parallel).
- Removes background with **Photoshop Remove Background** (recommended baseline), with optional fallback modes.
- Frames/crops headshot using face-first framing with extra top hair headroom to reduce clipping while keeping 250x250 composition.
- Uses dlib HOG face detection when available (falls back to OpenCV Haar) for better centering robustness.
- Shows large previews on white / grey / black backgrounds.
- Lets user open both source + output in Photoshop for manual touchups.
- Emits issue warnings (possible clipping, weak detections, odd coverage).

## Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open: `http://localhost:7860`

## Photoshop setup (Windows)

You can configure Photoshop in either way:

1. Put the path in the UI field **Photoshop Executable Path** (recommended).
2. Or set env var `PHOTOSHOP_EXE`.

Example PowerShell:

```powershell
$env:PHOTOSHOP_EXE = "C:\Program Files\Adobe\Adobe Photoshop 2024\Photoshop.exe"
python app.py
```

In the UI, set **Background Removal Engine** to:

- `photoshop` → strict Photoshop Remove Background only (best quality baseline).
- `auto` → Photoshop first, then rembg fallback if Photoshop fails.
- `rembg` → local AI only (no Photoshop).

## Notes

- For `photoshop` mode, only the Photoshop/background-removal step runs single-threaded; preprocessing (resize/optimize) runs in parallel outside Photoshop.
- If Photoshop appears stuck on one image, check for hidden modal dialogs in Photoshop and disable startup/compatibility prompts.
- Photoshop processing now stages per-image files under `output/.ps_jobs` to avoid temp-file lock/deletion races on Windows.
- Photoshop automation now runs COM-only by default (`pywin32`) for stability; this avoids Photoshop CLI SPL memory-manager failures seen on recent builds (including 2026).
- Photoshop remove-background now uses an expanded compatibility fallback chain (`removeBackground`/`autoCutout` IDs → `Select Subject` IDs → `selection mask`) for versions where certain menu items are unavailable.
- Batch processing now pre-selects the first output and auto-renders white/grey/black previews.
- Clicking a gallery item now also sets the preview picker so white/grey/black previews update immediately from top-right selection.
- Batch report now includes simple processing logs and progress is shown while preprocessing + Photoshop/crop stages run.
- UI preview/gallery files are mirrored into a local `.ui_cache` under the app directory to avoid Gradio external-path cache errors.
- UI selection now maps cached preview paths back to the real output path (`Matched Output`) so Photoshop opens the actual exported file.
- Added **Browse Input Folder** / **Browse Output Folder** buttons to open native folder picker dialogs (no copy/paste path needed).
- If `photoshop` mode is selected without an exe path, the report returns a single clear configuration error instead of 1 failure per file.

## Supported input file types

`jpg`, `jpeg`, `png`, `webp`, `tif`, `tiff`, `bmp`, `heic`
