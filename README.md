# Pro Headshot Cutout Studio

Photoshop-first batch tool for creating consistent `250x250` transparent PNG headshots with a dark-mode UI.

## What it does

- Scans source folders recursively.
- Preserves input folder structure in output.
- Optimizes huge source images before processing for speed.
- Removes background with **Photoshop Remove Background** (recommended baseline), with optional fallback modes.
- Frames/crops headshot to keep head centered with slight under-chin visibility.
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

- For `photoshop` mode, worker count is forced to 1 because Photoshop automation is single-instance and parallel jobs can conflict.
- If Photoshop appears stuck on one image, check for hidden modal dialogs in Photoshop and disable startup/compatibility prompts.
- Batch processing now pre-selects the first output and auto-renders white/grey/black previews.
- Added **Browse Input Folder** / **Browse Output Folder** buttons to open native folder picker dialogs (no copy/paste path needed).
- If `photoshop` mode is selected without an exe path, the report returns a single clear configuration error instead of 1 failure per file.

## Supported input file types

`jpg`, `jpeg`, `png`, `webp`, `tif`, `tiff`, `bmp`, `heic`
