# Pro Headshot Cutout Studio

A production-oriented batch headshot tool with a dark-mode UI that:

- Ingests large source folders recursively.
- Optimizes oversized files before heavy processing for speed.
- Removes backgrounds (Photoshop hook when configured, local AI fallback via `rembg`).
- Crops to a 250x250 transparent PNG headshot focused on head + slight under-chin.
- Preserves input folder structure in output.
- Flags potential quality issues.
- Shows large final previews on white/grey/black.
- Lets users open both source + output in Photoshop for manual fixes.

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open: `http://localhost:7860`

## Photoshop integration

Set environment vars to enable direct launch/edit:

- `PHOTOSHOP_EXE` → full path to Photoshop executable.
- `PHOTOSHOP_BG_JSX` → optional JSX script path for custom automation.

If not configured, the app uses local AI background removal (`rembg`) automatically.

## Output behavior

- Final format: transparent PNG.
- Final size: `250x250`.
- Folder structure is preserved from input root to output root.

## Supported input file types

`jpg`, `jpeg`, `png`, `webp`, `tif`, `tiff`, `bmp`, `heic`
