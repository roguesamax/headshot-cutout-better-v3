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
except Exception:  # rembg is optional at runtime
    rembg_remove = None


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp", ".heic"}


@dataclass
class ProcessConfig:
    max_side_for_processing: int = 2200
    output_size: int = 250
    workers: int = max(1, (os.cpu_count() or 4) - 1)


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


def remove_background_with_photoshop_or_fallback(rgb_image: Image.Image) -> Image.Image:
    """
    If Photoshop automation variables are configured, attempts to run Photoshop JSX.
    Otherwise falls back to rembg for local AI background removal.
    """
    if os.name == "nt" and os.environ.get("PHOTOSHOP_EXE") and os.environ.get("PHOTOSHOP_BG_JSX"):
        with tempfile.TemporaryDirectory() as td:
            tmp_in = Path(td) / "input.png"
            tmp_out = Path(td) / "output.png"
            rgb_image.save(tmp_in)
            try:
                subprocess.run(
                    [
                        os.environ["PHOTOSHOP_EXE"],
                        "-r",
                        os.environ["PHOTOSHOP_BG_JSX"],
                        str(tmp_in),
                        str(tmp_out),
                    ],
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                )
                if tmp_out.exists():
                    return Image.open(tmp_out).convert("RGBA")
            except Exception:
                pass

    if rembg_remove is None:
        raise RuntimeError(
            "Background removal dependency not available. Install rembg/onnxruntime or configure Photoshop automation."
        )

    raw = np.array(rgb_image)
    out = rembg_remove(raw)
    return Image.fromarray(out).convert("RGBA")


def detect_primary_face_bbox(rgba_image: Image.Image) -> tuple[int, int, int, int] | None:
    rgb = np.array(rgba_image.convert("RGB"))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)
    faces = CASCADE.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(40, 40))
    if len(faces) == 0:
        return None
    x, y, w, h = sorted(faces, key=lambda f: f[2] * f[3], reverse=True)[0]
    return int(x), int(y), int(w), int(h)


def crop_headshot(rgba_image: Image.Image, config: ProcessConfig) -> tuple[Image.Image, list[str]]:
    warnings: list[str] = []
    width, height = rgba_image.size
    face = detect_primary_face_bbox(rgba_image)

    if face is None:
        warnings.append("No face detected reliably; used center crop fallback.")
        side = min(width, height)
        x0 = (width - side) // 2
        y0 = (height - side) // 2
        crop = rgba_image.crop((x0, y0, x0 + side, y0 + side))
        return crop.resize((config.output_size, config.output_size), Image.Resampling.LANCZOS), warnings

    x, y, w, h = face

    # Expand to include top hair and slightly under chin.
    left = x - int(w * 0.45)
    right = x + w + int(w * 0.45)
    top = y - int(h * 0.40)
    bottom = y + h + int(h * 0.18)

    cx = (left + right) // 2
    cy = (top + bottom) // 2
    side = max(right - left, bottom - top)

    left = cx - side // 2
    top = cy - side // 2
    right = left + side
    bottom = top + side

    # clamp to image bounds
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

    if x < width * 0.05 or (x + w) > width * 0.95:
        warnings.append("Face is close to the side edge; ear clipping risk.")
    if y < height * 0.03:
        warnings.append("Forehead is close to top edge; crop risk.")

    crop = rgba_image.crop((left, top, right, bottom))
    crop = crop.resize((config.output_size, config.output_size), Image.Resampling.LANCZOS)

    alpha = np.array(crop.split()[-1])
    visible_ratio = float(np.count_nonzero(alpha > 10)) / alpha.size
    if visible_ratio < 0.18:
        warnings.append("Visible subject area is too small; likely mis-detection.")
    if visible_ratio > 0.82:
        warnings.append("Visible subject area is very high; possible background remnants.")

    return crop, warnings


def process_one(src_path: Path, input_root: Path, output_root: Path, config: ProcessConfig) -> dict:
    rel = src_path.relative_to(input_root)
    out_path = output_root / rel.with_suffix(".png")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    optimized = optimize_input_image(src_path, config)
    rgba = remove_background_with_photoshop_or_fallback(optimized)
    headshot, warnings = crop_headshot(rgba, config)
    headshot.save(out_path, format="PNG", optimize=True)

    return {
        "source": str(src_path),
        "output": str(out_path),
        "relative": str(rel),
        "warnings": warnings,
    }


def compose_preview(headshot_path: str, bg: tuple[int, int, int]) -> np.ndarray:
    img = Image.open(headshot_path).convert("RGBA")
    canvas = Image.new("RGBA", img.size, (*bg, 255))
    canvas.alpha_composite(img)
    return np.array(canvas.convert("RGB"))


def open_in_photoshop(source_path: str, output_path: str) -> str:
    photoshop_exe = os.environ.get("PHOTOSHOP_EXE")
    if not photoshop_exe:
        return "Set PHOTOSHOP_EXE environment variable to enable one-click Photoshop launch."

    try:
        subprocess.Popen([photoshop_exe, source_path, output_path])
        return "Opened source and output in Photoshop."
    except Exception as exc:
        return f"Could not open Photoshop: {exc}"


def run_batch(input_folder: str, output_folder: str, workers: int, max_side: int) -> tuple[str, list[str], str]:
    input_root = Path(input_folder)
    output_root = Path(output_folder)

    if not input_root.exists():
        return "Input folder does not exist.", [], "[]"

    images = list(iter_images(input_root))
    if not images:
        return "No supported images found.", [], "[]"

    output_root.mkdir(parents=True, exist_ok=True)

    config = ProcessConfig(workers=max(1, workers), max_side_for_processing=max_side)
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
        "failures": failures,
        "issues": [asdict(HeadshotIssue(r["relative"], r["warnings"])) for r in results if r["warnings"]],
    }

    preview_paths = [r["output"] for r in results]
    report_text = json.dumps(report, indent=2)
    return report_text, preview_paths, json.dumps(results)


def show_preview(image_path: str) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    if not image_path:
        blank = np.zeros((250, 250, 3), dtype=np.uint8)
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


def build_ui() -> gr.Blocks:
    theme = gr.themes.Soft(
        primary_hue="indigo",
        secondary_hue="slate",
        neutral_hue="slate",
        radius_size=gr.themes.sizes.radius_lg,
    )

    css = """
    .gradio-container {background: #0f172a !important; color: #e2e8f0 !important;}
    .panel {background: #111827 !important; border: 1px solid #334155 !important; border-radius: 16px !important;}
    #hero-preview img {min-height: 540px; object-fit: contain; background: #000;}
    """

    with gr.Blocks(theme=theme, css=css, title="Pro Headshot Cutout Studio") as demo:
        gr.Markdown("## Pro Headshot Cutout Studio")
        gr.Markdown("Fast batch Photoshop-style headshot extraction with structure-preserving output and QA checks.")

        with gr.Row():
            with gr.Column(scale=1, elem_classes=["panel"]):
                input_folder = gr.Textbox(label="Input Folder", placeholder="/path/to/source/root")
                output_folder = gr.Textbox(label="Output Folder", placeholder="/path/to/output/root")
                workers = gr.Slider(1, max(2, os.cpu_count() or 8), value=max(1, (os.cpu_count() or 4) - 1), step=1, label="Parallel Workers")
                max_side = gr.Slider(800, 4000, value=2200, step=100, label="Pre-resize Max Side (speed control)")
                run_btn = gr.Button("Process Batch", variant="primary")
                report = gr.Code(label="Batch Report / Issue Detection", language="json")

            with gr.Column(scale=2, elem_classes=["panel"]):
                gallery = gr.Gallery(label="Processed Outputs", columns=4, rows=2, height=300, object_fit="contain")
                selected = gr.Textbox(label="Selected Output PNG", interactive=False)
                state_results = gr.State("[]")
                source_for_open = gr.Textbox(label="Matched Source", interactive=False)

                with gr.Row():
                    open_ps = gr.Button("Open Source + Output in Photoshop")
                    open_status = gr.Textbox(label="Photoshop Launch Status", interactive=False)

                with gr.Row(elem_id="hero-preview"):
                    preview_white = gr.Image(label="White", interactive=False, height=540)
                    preview_grey = gr.Image(label="Grey", interactive=False, height=540)
                    preview_black = gr.Image(label="Black", interactive=False, height=540)

        run_btn.click(
            fn=run_batch,
            inputs=[input_folder, output_folder, workers, max_side],
            outputs=[report, gallery, state_results],
        )

        gallery.select(
            fn=lambda evt: evt.value["image"][0] if isinstance(evt.value, dict) else evt.value,
            outputs=selected,
        )

        selected.change(fn=show_preview, inputs=selected, outputs=[preview_white, preview_grey, preview_black])
        selected.change(fn=find_source_from_results, inputs=[selected, state_results], outputs=source_for_open)
        open_ps.click(fn=open_in_photoshop, inputs=[source_for_open, selected], outputs=open_status)

    return demo


if __name__ == "__main__":
    app = build_ui()
    app.launch(server_name="0.0.0.0", server_port=7860)
