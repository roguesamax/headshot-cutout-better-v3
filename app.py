from __future__ import annotations

import json
import os
import subprocess
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable

import cv2
import gradio as gr
import numpy as np
from PIL import Image

try:
    from rembg import remove as rembg_remove
except Exception:
    rembg_remove = None


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp", ".heic"}


@dataclass
class ProcessConfig:
    max_side_for_processing: int = 2200
    output_size: int = 250
    workers: int = max(1, (os.cpu_count() or 4) - 1)
    removal_engine: str = "photoshop"


@dataclass
class HeadshotIssue:
    relative_path: str
    warnings: list[str]


CASCADE = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")


def iter_images(root: Path) -> Iterable[Path]:
    for path in root.rglob("*"):
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            yield path


def optimize_input_image(src_path: Path, config: ProcessConfig) -> Image.Image:
    image = Image.open(src_path).convert("RGB")
    width, height = image.size
    max_side = max(width, height)
    if max_side > config.max_side_for_processing:
        scale = config.max_side_for_processing / max_side
        image = image.resize((int(width * scale), int(height * scale)), Image.Resampling.LANCZOS)
    return image


def _escape_jsx_path(path: str) -> str:
    return path.replace("\\", "\\\\")


def _remove_background_photoshop(rgb_image: Image.Image) -> Image.Image:
    photoshop_exe = os.environ.get("PHOTOSHOP_EXE")
    if not photoshop_exe:
        raise RuntimeError("PHOTOSHOP_EXE is not set.")

    with tempfile.TemporaryDirectory() as td:
        tmp_dir = Path(td)
        src_file = tmp_dir / "source.png"
        out_file = tmp_dir / "result.png"
        jsx_file = tmp_dir / "remove_bg.jsx"

        rgb_image.save(src_file, format="PNG")

        jsx = f"""
#target photoshop
app.displayDialogs = DialogModes.NO;
var inFile = new File('{_escape_jsx_path(str(src_file))}');
var outFile = new File('{_escape_jsx_path(str(out_file))}');
app.open(inFile);
var doc = app.activeDocument;

try {{
    app.runMenuItem(stringIDToTypeID('autoCutout')); // Remove Background
}} catch (e) {{
    try {{
        app.runMenuItem(stringIDToTypeID('autoCutoutSubject')); // fallback for older ids
    }} catch (e2) {{
        throw e2;
    }}
}}

var opts = new PNGSaveOptions();
opts.compression = 6;
doc.saveAs(outFile, opts, true, Extension.LOWERCASE);
doc.close(SaveOptions.DONOTSAVECHANGES);
"""
        jsx_file.write_text(jsx, encoding="utf-8")

        proc = subprocess.run(
            [photoshop_exe, "-r", str(jsx_file)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        if proc.returncode != 0:
            raise RuntimeError(f"Photoshop failed: {proc.stderr.strip() or proc.stdout.strip()}")
        if not out_file.exists():
            raise RuntimeError("Photoshop ran but did not produce an output PNG.")

        return Image.open(out_file).convert("RGBA")


def _remove_background_rembg(rgb_image: Image.Image) -> Image.Image:
    if rembg_remove is None:
        raise RuntimeError("rembg is not available. Install rembg + onnxruntime.")
    raw = np.array(rgb_image)
    out = rembg_remove(raw)
    return Image.fromarray(out).convert("RGBA")


def remove_background(rgb_image: Image.Image, engine: str) -> tuple[Image.Image, list[str]]:
    warnings: list[str] = []
    engine = (engine or "photoshop").lower().strip()

    if engine == "photoshop":
        rgba = _remove_background_photoshop(rgb_image)
        return rgba, warnings

    if engine == "rembg":
        rgba = _remove_background_rembg(rgb_image)
        return rgba, warnings

    if engine == "auto":
        try:
            rgba = _remove_background_photoshop(rgb_image)
            warnings.append("Used Photoshop background removal.")
            return rgba, warnings
        except Exception:
            rgba = _remove_background_rembg(rgb_image)
            warnings.append("Photoshop unavailable, used rembg fallback.")
            return rgba, warnings

    raise RuntimeError("Invalid removal engine. Use photoshop, auto, or rembg.")


def detect_primary_face_bbox(rgba_image: Image.Image) -> tuple[int, int, int, int] | None:
    rgb = np.array(rgba_image.convert("RGB"))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    faces = CASCADE.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(36, 36))
    if len(faces) == 0:
        return None
    x, y, w, h = sorted(faces, key=lambda f: f[2] * f[3], reverse=True)[0]
    return int(x), int(y), int(w), int(h)


def alpha_bbox(rgba_image: Image.Image, threshold: int = 12) -> tuple[int, int, int, int] | None:
    alpha = np.array(rgba_image.split()[-1])
    ys, xs = np.where(alpha > threshold)
    if len(xs) == 0 or len(ys) == 0:
        return None
    return int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())


def crop_headshot(rgba_image: Image.Image, config: ProcessConfig) -> tuple[Image.Image, list[str]]:
    warnings: list[str] = []
    width, height = rgba_image.size

    a_bbox = alpha_bbox(rgba_image)
    face = detect_primary_face_bbox(rgba_image)

    if a_bbox is None and face is None:
        warnings.append("No subject/face reliably detected; used center square fallback.")
        side = min(width, height)
        left = (width - side) // 2
        top = (height - side) // 2
        crop = rgba_image.crop((left, top, left + side, top + side))
        return crop.resize((config.output_size, config.output_size), Image.Resampling.LANCZOS), warnings

    if a_bbox is not None:
        ax1, ay1, ax2, ay2 = a_bbox
        subj_w = ax2 - ax1 + 1
        subj_h = ay2 - ay1 + 1
        cx = (ax1 + ax2) / 2.0

        # place chin a bit below center for the requested style
        if face is not None:
            fx, fy, fw, fh = face
            cy = fy + fh * 0.53
            side = max(subj_w * 1.28, subj_h * 1.12)
        else:
            cy = ay1 + subj_h * 0.46
            side = max(subj_w * 1.30, subj_h * 1.16)
            warnings.append("Face detection failed; used alpha-mask framing.")
    else:
        fx, fy, fw, fh = face
        cx = fx + fw / 2.0
        cy = fy + fh * 0.53
        side = max(fw * 2.0, fh * 1.9)
        warnings.append("Alpha mask was weak; used face-based framing.")

    side = int(max(32, min(side, max(width, height))))
    left = int(round(cx - side / 2))
    top = int(round(cy - side / 2))
    right = left + side
    bottom = top + side

    if left < 0:
        right -= left
        left = 0
    if top < 0:
        bottom -= top
        top = 0
    if right > width:
        shift = right - width
        left -= shift
        right = width
    if bottom > height:
        shift = bottom - height
        top -= shift
        bottom = height

    left = max(0, left)
    top = max(0, top)

    crop = rgba_image.crop((left, top, right, bottom)).resize((config.output_size, config.output_size), Image.Resampling.LANCZOS)

    alpha = np.array(crop.split()[-1])
    visible_ratio = float(np.count_nonzero(alpha > 10)) / alpha.size

    if visible_ratio < 0.14:
        warnings.append("Visible subject area is too small; likely bad detect or tiny source subject.")
    if visible_ratio > 0.88:
        warnings.append("Visible subject area is very high; possible background remnants.")

    cols = np.count_nonzero(alpha > 12, axis=0)
    rows = np.count_nonzero(alpha > 12, axis=1)
    if cols.size > 0 and (cols[0] > 12 or cols[-1] > 12):
        warnings.append("Subject touches side boundary; ear clipping risk.")
    if rows.size > 0 and rows[0] > 12:
        warnings.append("Subject touches top boundary; hair/forehead clipping risk.")

    return crop, warnings


def process_one(src_path: Path, input_root: Path, output_root: Path, config: ProcessConfig) -> dict:
    rel = src_path.relative_to(input_root)
    out_path = output_root / rel.with_suffix(".png")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    optimized = optimize_input_image(src_path, config)
    rgba, bg_warnings = remove_background(optimized, config.removal_engine)
    headshot, crop_warnings = crop_headshot(rgba, config)
    headshot.save(out_path, format="PNG", optimize=True)

    return {
        "source": str(src_path),
        "output": str(out_path),
        "relative": str(rel),
        "warnings": [*bg_warnings, *crop_warnings],
    }


def compose_preview(headshot_path: str, bg: tuple[int, int, int]) -> np.ndarray:
    img = Image.open(headshot_path).convert("RGBA")
    canvas = Image.new("RGBA", img.size, (*bg, 255))
    canvas.alpha_composite(img)
    return np.array(canvas.convert("RGB"))


def open_in_photoshop(source_path: str, output_path: str) -> str:
    photoshop_exe = os.environ.get("PHOTOSHOP_EXE")
    if not photoshop_exe:
        return "Set PHOTOSHOP_EXE to enable Photoshop launch."

    if not source_path or not output_path:
        return "No selected output/source mapping yet. Select an output first."

    try:
        subprocess.Popen([photoshop_exe, source_path, output_path])
        return "Opened source and output in Photoshop."
    except Exception as exc:
        return f"Could not open Photoshop: {exc}"


def run_batch(
    input_folder: str,
    output_folder: str,
    workers: int,
    max_side: int,
    removal_engine: str,
) -> tuple[str, list[str], str, list[str]]:
    input_root = Path(input_folder)
    output_root = Path(output_folder)

    if not input_root.exists():
        return "Input folder does not exist.", [], "[]", []

    images = list(iter_images(input_root))
    if not images:
        return "No supported images found.", [], "[]", []

    output_root.mkdir(parents=True, exist_ok=True)

    config = ProcessConfig(
        workers=max(1, int(workers)),
        max_side_for_processing=int(max_side),
        removal_engine=(removal_engine or "photoshop").lower(),
    )

    # Photoshop is effectively single-process; parallel workers can cause contention.
    if config.removal_engine == "photoshop":
        config.workers = 1

    results: list[dict] = []
    failures: list[str] = []

    with ThreadPoolExecutor(max_workers=config.workers) as executor:
        futures = {executor.submit(process_one, img, input_root, output_root, config): img for img in images}
        for future in as_completed(futures):
            image = futures[future]
            try:
                results.append(future.result())
            except Exception as exc:
                failures.append(f"{image}: {exc}")

    results.sort(key=lambda r: r["relative"])

    report = {
        "total": len(images),
        "processed": len(results),
        "failed": len(failures),
        "removal_engine": config.removal_engine,
        "workers_used": config.workers,
        "failures": failures,
        "issues": [asdict(HeadshotIssue(r["relative"], r["warnings"])) for r in results if r["warnings"]],
    }

    preview_paths = [r["output"] for r in results]
    report_text = json.dumps(report, indent=2)
    return report_text, preview_paths, json.dumps(results), preview_paths


def show_preview(image_path: str) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    if not image_path:
        blank = np.full((250, 250, 3), 35, dtype=np.uint8)
        return blank, blank, blank

    if not Path(image_path).exists():
        blank = np.full((250, 250, 3), 35, dtype=np.uint8)
        return blank, blank, blank

    return (
        compose_preview(image_path, (255, 255, 255)),
        compose_preview(image_path, (127, 127, 127)),
        compose_preview(image_path, (0, 0, 0)),
    )


def find_source_from_results(image_path: str, raw_results: str) -> str:
    if not image_path or not raw_results:
        return ""
    rows = json.loads(raw_results)
    for row in rows:
        if row["output"] == image_path:
            return row["source"]
    return ""


def extract_gallery_path(evt: gr.SelectData) -> str:
    value = evt.value
    if isinstance(value, str):
        return value
    if isinstance(value, (list, tuple)) and len(value) > 0:
        return str(value[0])
    if isinstance(value, dict):
        if "image" in value and isinstance(value["image"], (list, tuple)) and value["image"]:
            return str(value["image"][0])
        if "name" in value:
            return str(value["name"])
    return ""


def build_ui() -> gr.Blocks:
    theme = gr.themes.Soft(
        primary_hue="indigo",
        secondary_hue="slate",
        neutral_hue="slate",
        radius_size=gr.themes.sizes.radius_lg,
    )

    css = """
    .gradio-container {background: #0f172a !important; color: #e2e8f0 !important;}
    .panel {background: #111827 !important; border: 1px solid #334155 !important; border-radius: 16px !important; padding: 12px;}
    #hero-preview img {min-height: 460px; object-fit: contain; background: #000;}
    """

    with gr.Blocks(theme=theme, css=css, title="Pro Headshot Cutout Studio") as demo:
        gr.Markdown("## Pro Headshot Cutout Studio")
        gr.Markdown(
            "Photoshop-first batch headshot extraction (250x250 PNG), with structure-preserving output and QA checks."
        )

        with gr.Row():
            with gr.Column(scale=1, elem_classes=["panel"]):
                input_folder = gr.Textbox(label="Input Folder", placeholder="C:/path/to/source/root")
                output_folder = gr.Textbox(label="Output Folder", placeholder="C:/path/to/output/root")
                removal_engine = gr.Radio(
                    choices=["photoshop", "auto", "rembg"],
                    value="photoshop",
                    label="Background Removal Engine",
                    info="photoshop = strict Photoshop Remove Background, auto = Photoshop then rembg fallback",
                )
                workers = gr.Slider(1, max(2, os.cpu_count() or 8), value=max(1, (os.cpu_count() or 4) - 1), step=1, label="Parallel Workers")
                max_side = gr.Slider(800, 4500, value=2200, step=100, label="Pre-resize Max Side (speed control)")
                run_btn = gr.Button("Process Batch", variant="primary")
                report = gr.Code(label="Batch Report / Issue Detection", language="json")

            with gr.Column(scale=2, elem_classes=["panel"]):
                gallery = gr.Gallery(label="Processed Outputs", columns=5, rows=2, height=320, object_fit="contain")
                selected = gr.Textbox(label="Selected Output PNG", interactive=False)
                picker = gr.Dropdown(label="Pick Output (stable preview)", choices=[], interactive=True)
                state_results = gr.State("[]")
                source_for_open = gr.Textbox(label="Matched Source", interactive=False)

                with gr.Row():
                    open_ps = gr.Button("Open Source + Output in Photoshop")
                    open_status = gr.Textbox(label="Photoshop Launch Status", interactive=False)

                with gr.Row(elem_id="hero-preview"):
                    preview_white = gr.Image(label="White", interactive=False, height=460)
                    preview_grey = gr.Image(label="Grey", interactive=False, height=460)
                    preview_black = gr.Image(label="Black", interactive=False, height=460)

        run_btn.click(
            fn=run_batch,
            inputs=[input_folder, output_folder, workers, max_side, removal_engine],
            outputs=[report, gallery, state_results, picker],
        )

        gallery.select(fn=extract_gallery_path, outputs=selected)
        picker.change(fn=lambda v: v or "", inputs=picker, outputs=selected)
        selected.change(fn=show_preview, inputs=selected, outputs=[preview_white, preview_grey, preview_black])
        selected.change(fn=find_source_from_results, inputs=[selected, state_results], outputs=source_for_open)
        open_ps.click(fn=open_in_photoshop, inputs=[source_for_open, selected], outputs=open_status)

    return demo


if __name__ == "__main__":
    app = build_ui()
    app.launch(server_name="0.0.0.0", server_port=7860)
