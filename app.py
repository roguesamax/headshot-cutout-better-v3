from __future__ import annotations

import importlib
import importlib.util
import json
import shutil
import os
import subprocess
import time
import uuid
import tkinter as tk
from tkinter import filedialog
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable

import cv2
import gradio as gr
import numpy as np
from PIL import Image


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp", ".heic"}


def _resolve_rembg_remove():
    if importlib.util.find_spec("rembg") is None:
        return None
    rembg_module = importlib.import_module("rembg")
    return getattr(rembg_module, "remove", None)


REMBG_REMOVE = _resolve_rembg_remove()



def _resolve_win32com_client():
    if os.name != "nt":
        return None
    if importlib.util.find_spec("win32com.client") is None:
        return None
    return importlib.import_module("win32com.client")


WIN32COM_CLIENT = _resolve_win32com_client()


def _resolve_dlib_module():
    if importlib.util.find_spec("dlib") is None:
        return None
    return importlib.import_module("dlib")


DLIB = _resolve_dlib_module()
DLIB_HOG_DETECTOR = DLIB.get_frontal_face_detector() if DLIB is not None else None


@dataclass
class ProcessConfig:
    max_side_for_processing: int = 2200
    output_size: int = 250
    workers: int = max(1, (os.cpu_count() or 4) - 1)
    removal_engine: str = "photoshop"
    photoshop_exe: str = ""
    job_root: str = ""
    prep_root: str = ""
    photoshop_automation_mode: str = "com_only"


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


def preprocess_image_to_stage(src_path: Path, input_root: Path, config: ProcessConfig) -> dict:
    rel = src_path.relative_to(input_root)
    prep_root = Path(config.prep_root)
    staged_path = prep_root / rel.with_suffix(".jpg")
    staged_path.parent.mkdir(parents=True, exist_ok=True)

    optimized = optimize_input_image(src_path, config)
    optimized.save(staged_path, format="JPEG", quality=90, optimize=True, progressive=True)

    return {"source": str(src_path), "relative": str(rel), "staged": str(staged_path)}


def _escape_jsx_path(path: str) -> str:
    return path.replace("\\", "\\\\")


def browse_for_directory(current_value: str) -> str:
    """Open native folder picker (desktop/local runs)."""
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        initial = current_value if current_value and Path(current_value).exists() else str(Path.home())
        selected = filedialog.askdirectory(initialdir=initial)
        root.destroy()
        return selected or current_value or ""
    except Exception:
        return current_value or ""


def discover_photoshop_exe() -> str:
    env_val = os.environ.get("PHOTOSHOP_EXE", "").strip()
    if env_val and Path(env_val).exists():
        return env_val

    if os.name != "nt":
        return ""

    candidates = [
        r"C:\Program Files\Adobe\Adobe Photoshop 2025\Photoshop.exe",
        r"C:\Program Files\Adobe\Adobe Photoshop 2024\Photoshop.exe",
        r"C:\Program Files\Adobe\Adobe Photoshop 2023\Photoshop.exe",
        r"C:\Program Files\Adobe\Adobe Photoshop 2022\Photoshop.exe",
        r"C:\Program Files\Adobe\Adobe Photoshop 2021\Photoshop.exe",
        r"C:\Program Files\Adobe\Adobe Photoshop 2020\Photoshop.exe",
    ]
    for path in candidates:
        if Path(path).exists():
            return path

    return ""


def _run_jsx_via_com(jsx: str, photoshop_exe: str) -> None:
    if WIN32COM_CLIENT is None:
        raise RuntimeError("pywin32 is not installed; COM automation unavailable.")

    # Try to connect to existing Photoshop instance.
    app = None
    try:
        app = WIN32COM_CLIENT.GetActiveObject("Photoshop.Application")
    except Exception:
        if photoshop_exe and Path(photoshop_exe).exists():
            subprocess.Popen([photoshop_exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(2.5)
        app = WIN32COM_CLIENT.Dispatch("Photoshop.Application")

    app.DoJavaScript(jsx)


def _remove_background_photoshop_cli(job_dir: Path, photoshop_exe: str, jsx_file: Path, out_file: Path) -> None:
    if not photoshop_exe:
        raise RuntimeError("Photoshop executable is not configured for CLI fallback.")

    proc = subprocess.Popen(
        [photoshop_exe, "-r", str(jsx_file)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    deadline = time.time() + 120
    while time.time() < deadline:
        if out_file.exists():
            return

        if proc.poll() is not None:
            stdout, stderr = proc.communicate()
            if out_file.exists():
                return
            details = (stderr or stdout).strip()
            raise RuntimeError(f"Photoshop CLI failed: {details or 'No output file created.'}")

        time.sleep(0.25)

    raise RuntimeError("Timed out waiting for Photoshop output in CLI mode.")


def _remove_background_photoshop(staged_input_path: Path, photoshop_exe: str, job_root: Path, automation_mode: str) -> Image.Image:
    if WIN32COM_CLIENT is None and automation_mode in ("com_only", "com_then_cli"):
        raise RuntimeError(
            "pywin32 is required for Photoshop automation on your setup. Install pywin32 and restart the app."
        )
    if not photoshop_exe and automation_mode in ("cli_only", "com_then_cli"):
        raise RuntimeError("Photoshop executable is not configured.")

    job_id = uuid.uuid4().hex
    job_dir = job_root / f"job_{job_id}"
    job_dir.mkdir(parents=True, exist_ok=True)

    out_file = job_dir / f"result_{job_id}.png"
    status_file = job_dir / f"status_{job_id}.txt"
    jsx_file = job_dir / f"remove_bg_{job_id}.jsx"

    jsx = f"""
#target photoshop
app.displayDialogs = DialogModes.NO;
var inFile = new File('{_escape_jsx_path(str(staged_input_path))}');
var outFile = new File('{_escape_jsx_path(str(out_file))}');
var statusFile = new File('{_escape_jsx_path(str(status_file))}');

function unlockLayerIfNeeded(doc) {{
    try {{
        if (doc.activeLayer && doc.activeLayer.isBackgroundLayer) {{
            doc.activeLayer.isBackgroundLayer = false;
        }}
    }} catch (e) {{}}
}}

function tryRunMenu(idName) {{
    try {{
        app.runMenuItem(stringIDToTypeID(idName));
        return true;
    }} catch (e) {{
        return false;
    }}
}}

function tryExecAction(idName) {{
    try {{
        executeAction(stringIDToTypeID(idName), undefined, DialogModes.NO);
        return true;
    }} catch (e) {{
        return false;
    }}
}}

function removeBackgroundCompat() {{
    // Newer Photoshop builds may expose different command IDs.
    if (tryExecAction('removeBackground')) return true;
    if (tryRunMenu('removeBackground')) return true;
    if (tryRunMenu('autoCutout')) return true;
    if (tryRunMenu('autoCutoutSubject')) return true;
    return false;
}}

function selectSubjectCompat() {{
    if (tryExecAction('selectSubject')) return true;
    if (tryRunMenu('selectSubject')) return true;
    if (tryRunMenu('autoCutoutSubject')) return true;
    return false;
}}

function applySelectionMask(doc) {{
    var idMk = charIDToTypeID('Mk  ');
    var desc = new ActionDescriptor();
    var idNw = charIDToTypeID('Nw  ');
    var idChnl = charIDToTypeID('Chnl');
    desc.putClass(idNw, idChnl);

    var idAt = charIDToTypeID('At  ');
    var ref = new ActionReference();
    ref.putEnumerated(charIDToTypeID('Chnl'), charIDToTypeID('Chnl'), charIDToTypeID('Msk '));
    desc.putReference(idAt, ref);

    var idUsng = charIDToTypeID('Usng');
    var idUsrM = charIDToTypeID('UsrM');
    var idRvlS = charIDToTypeID('RvlS');
    desc.putEnumerated(idUsng, idUsrM, idRvlS);

    executeAction(idMk, desc, DialogModes.NO);
    try {{ doc.selection.deselect(); }} catch (e) {{}}
}}

try {{
    app.open(inFile);
    var doc = app.activeDocument;
    unlockLayerIfNeeded(doc);

    var removed = removeBackgroundCompat();
    if (!removed) {{
        var selected = selectSubjectCompat();
        if (!selected) {{
            throw new Error('No compatible Remove Background/Select Subject action found in this Photoshop build.');
        }}
        applySelectionMask(doc);
    }}

    var opts = new PNGSaveOptions();
    opts.compression = 6;
    doc.saveAs(outFile, opts, true, Extension.LOWERCASE);
    doc.close(SaveOptions.DONOTSAVECHANGES);

    statusFile.open('w');
    statusFile.write('OK');
    statusFile.close();
}} catch (err) {{
    try {{
        statusFile.open('w');
        statusFile.write('ERROR: ' + err.toString());
        statusFile.close();
    }} catch (_) {{}}
    throw err;
}}
"""
    jsx_file.write_text(jsx, encoding="utf-8")

    com_error: Exception | None = None
    cli_error: Exception | None = None

    if automation_mode in ("com_only", "com_then_cli"):
        try:
            _run_jsx_via_com(jsx, photoshop_exe)
        except Exception as exc:
            com_error = exc

    if (not out_file.exists()) and automation_mode in ("cli_only", "com_then_cli"):
        try:
            _remove_background_photoshop_cli(job_dir, photoshop_exe, jsx_file, out_file)
        except Exception as exc:
            cli_error = exc

    if not out_file.exists():
        if status_file.exists():
            raise RuntimeError(status_file.read_text(encoding="utf-8", errors="ignore").strip())
        if com_error and cli_error:
            raise RuntimeError(f"Photoshop COM error: {com_error} | CLI error: {cli_error}")
        if com_error:
            raise RuntimeError(f"Photoshop COM failed: {com_error}")
        if cli_error:
            raise RuntimeError(f"Photoshop CLI failed: {cli_error}")
        raise RuntimeError("Photoshop did not create output PNG.")

    return Image.open(out_file).convert("RGBA")


def _remove_background_rembg(staged_input_path: Path) -> Image.Image:
    if REMBG_REMOVE is None:
        raise RuntimeError("rembg is not available. Install rembg + onnxruntime.")
    rgb_image = Image.open(staged_input_path).convert("RGB")
    raw = np.array(rgb_image)
    out = REMBG_REMOVE(raw)
    return Image.fromarray(out).convert("RGBA")


def remove_background(staged_input_path: Path, config: ProcessConfig) -> tuple[Image.Image, list[str]]:
    warnings: list[str] = []
    engine = (config.removal_engine or "photoshop").lower().strip()

    if engine == "photoshop":
        rgba = _remove_background_photoshop(staged_input_path, config.photoshop_exe, Path(config.job_root), config.photoshop_automation_mode)
        return rgba, warnings

    if engine == "rembg":
        rgba = _remove_background_rembg(staged_input_path)
        return rgba, warnings

    if engine == "auto":
        try:
            rgba = _remove_background_photoshop(staged_input_path, config.photoshop_exe, Path(config.job_root), config.photoshop_automation_mode)
            warnings.append("Used Photoshop background removal.")
            return rgba, warnings
        except Exception:
            rgba = _remove_background_rembg(staged_input_path)
            warnings.append("Photoshop unavailable, used rembg fallback.")
            return rgba, warnings

    raise RuntimeError("Invalid removal engine. Use photoshop, auto, or rembg.")


def detect_primary_face_bbox(image: Image.Image) -> tuple[int, int, int, int] | None:
    rgb = np.array(image.convert("RGB"))
    gray = cv2.cvtColor(rgb, cv2.COLOR_RGB2GRAY)

    # 1) dlib HOG (if available) is generally more stable for portrait framing.
    if DLIB_HOG_DETECTOR is not None:
        try:
            rects = DLIB_HOG_DETECTOR(gray, 1)
            if rects:
                rect = sorted(rects, key=lambda r: (r.right() - r.left()) * (r.bottom() - r.top()), reverse=True)[0]
                x = max(0, int(rect.left()))
                y = max(0, int(rect.top()))
                w = max(1, int(rect.right() - rect.left()))
                h = max(1, int(rect.bottom() - rect.top()))
                return x, y, w, h
        except Exception:
            pass

    # 2) OpenCV Haar fallback.
    faces = CASCADE.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))
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


def crop_headshot(
    rgba_image: Image.Image,
    config: ProcessConfig,
    preferred_face: tuple[int, int, int, int] | None = None,
) -> tuple[Image.Image, list[str]]:
    warnings: list[str] = []
    width, height = rgba_image.size

    face = preferred_face or detect_primary_face_bbox(rgba_image)
    a_bbox = alpha_bbox(rgba_image)

    def clamp_square(cx: float, cy: float, side: float) -> tuple[int, int, int, int, float]:
        side = float(max(40, min(side, max(width, height))))
        left = int(round(cx - side / 2))
        top = int(round(cy - side / 2))
        right = int(round(left + side))
        bottom = int(round(top + side))

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
        side = float(min(right - left, bottom - top))
        right = int(left + side)
        bottom = int(top + side)
        return left, top, right, bottom, side

    if face is not None:
        fx, fy, fw, fh = face
        cx = fx + fw / 2.0

        # Target composition: chin ~3-5 px above bottom in 250x250 output.
        target_chin_px = config.output_size - 4
        side = max(fw * 1.58, fh * 1.72)

        jaw_y = fy + fh
        hair_top_est = fy - fh * 0.34
        left_ear_est = fx - fw * 0.22
        right_ear_est = fx + fw * 1.22

        chosen = None
        for _ in range(14):
            cy = jaw_y - ((target_chin_px / config.output_size) - 0.5) * side
            left, top, right, bottom, side_used = clamp_square(cx, cy, side)

            # Hard constraints to avoid clipping.
            ok_hair = top <= hair_top_est + 1
            ok_side_face = left <= left_ear_est + 1 and right >= right_ear_est - 1

            ok_alpha = True
            if a_bbox is not None:
                ax1, ay1, ax2, ay2 = a_bbox
                ok_alpha = top <= ay1 + 1 and left <= ax1 + 1 and right >= ax2 - 1

            if ok_hair and ok_side_face and ok_alpha:
                chosen = (left, top, right, bottom)
                break

            side *= 1.08

        if chosen is None:
            warnings.append("Adaptive framing hit bounds; used safest unclipped crop.")
            cy = jaw_y - ((target_chin_px / config.output_size) - 0.5) * side
            chosen = clamp_square(cx, cy, side)[:4]

        left, top, right, bottom = chosen

    elif a_bbox is not None:
        warnings.append("Face detection failed; used alpha-mask fallback framing.")
        ax1, ay1, ax2, ay2 = a_bbox
        subj_w = ax2 - ax1 + 1
        subj_h = ay2 - ay1 + 1
        cx = (ax1 + ax2) / 2.0
        cy = ay1 + subj_h * 0.32
        side = max(subj_w * 0.98, subj_h * 0.62)
        left, top, right, bottom, _ = clamp_square(cx, cy, side)
    else:
        warnings.append("No subject/face reliably detected; used center square fallback.")
        side = min(width, height)
        left = (width - side) // 2
        top = (height - side) // 2
        right = left + side
        bottom = top + side

    crop = rgba_image.crop((left, top, right, bottom)).resize((config.output_size, config.output_size), Image.Resampling.LANCZOS)

    alpha = np.array(crop.split()[-1])
    visible_ratio = float(np.count_nonzero(alpha > 10)) / alpha.size

    if visible_ratio < 0.10:
        warnings.append("Visible subject area is too small; likely bad detection.")
    if visible_ratio > 0.90:
        warnings.append("Visible subject area is very high; possible background remnants.")

    cols = np.count_nonzero(alpha > 12, axis=0)
    rows = np.count_nonzero(alpha > 12, axis=1)
    if cols.size > 0 and (cols[0] > 10 or cols[-1] > 10):
        warnings.append("Subject touches side boundary; ear clipping risk.")
    if rows.size > 0 and rows[0] > 10:
        warnings.append("Subject touches top boundary; hair clipping risk.")

    return crop, warnings


def process_one(prepped: dict, output_root: Path, config: ProcessConfig) -> dict:
    rel = Path(prepped["relative"])
    src_path = Path(prepped["source"])
    staged_path = Path(prepped["staged"])

    out_path = output_root / rel.with_suffix(".png")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    staged_rgb = Image.open(staged_path).convert("RGB")
    source_face = detect_primary_face_bbox(staged_rgb)

    rgba, bg_warnings = remove_background(staged_path, config)
    headshot, crop_warnings = crop_headshot(rgba, config, preferred_face=source_face)
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


def open_in_photoshop(source_path: str, output_path: str, photoshop_exe_input: str) -> str:
    photoshop_exe = (photoshop_exe_input or "").strip() or discover_photoshop_exe()
    if not photoshop_exe:
        return "Photoshop EXE not found. Set it in the UI field (or PHOTOSHOP_EXE env var)."

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
    photoshop_exe_input: str,
    progress=gr.Progress(),
) -> tuple[str, list[tuple[str, str]], str, gr.update, str, np.ndarray, np.ndarray, np.ndarray, str, str]:
    input_root = Path(input_folder)
    output_root = Path(output_folder)

    blank = np.full((250, 250, 3), 35, dtype=np.uint8)

    if not input_root.exists():
        return (
            "Input folder does not exist.",
            [],
            "[]",
            gr.update(choices=[], value=None),
            "",
            blank,
            blank,
            blank,
            "",
            "",
        )

    images = list(iter_images(input_root))
    if not images:
        return (
            "No supported images found.",
            [],
            "[]",
            gr.update(choices=[], value=None),
            "",
            blank,
            blank,
            blank,
            "",
            "",
        )

    output_root.mkdir(parents=True, exist_ok=True)

    photoshop_exe = (photoshop_exe_input or "").strip() or discover_photoshop_exe()
    job_root = output_root / ".ps_jobs"
    prep_root = output_root / ".prep"
    job_root.mkdir(parents=True, exist_ok=True)
    prep_root.mkdir(parents=True, exist_ok=True)

    config = ProcessConfig(
        workers=max(1, int(workers)),
        max_side_for_processing=int(max_side),
        removal_engine=(removal_engine or "photoshop").lower(),
        photoshop_exe=photoshop_exe,
        job_root=str(job_root),
        prep_root=str(prep_root),
        photoshop_automation_mode="com_only",
    )

    if config.removal_engine == "photoshop":
        config.workers = 1
        if not config.photoshop_exe:
            msg = {
                "total": len(images),
                "processed": 0,
                "failed": len(images),
                "removal_engine": config.removal_engine,
                "workers_used": 1,
                "preprocessed": 0,
                "failures": [
                    "Photoshop executable not configured. Fill 'Photoshop Executable Path' in UI or set PHOTOSHOP_EXE."
                ],
                "issues": [],
            }
            return (
                json.dumps(msg, indent=2),
                [],
                "[]",
                gr.update(choices=[], value=None),
                "",
                blank,
                blank,
                blank,
                "",
                "",
            )

    results: list[dict] = []
    failures: list[str] = []

    progress(0, desc="Starting batch…")

    prepped_items: list[dict] = []
    with ThreadPoolExecutor(max_workers=max(1, int(workers))) as prep_executor:
        prep_futures = {prep_executor.submit(preprocess_image_to_stage, img, input_root, config): img for img in images}
        prep_done = 0
        prep_total = max(1, len(prep_futures))
        for future in as_completed(prep_futures):
            image = prep_futures[future]
            prep_done += 1
            progress(prep_done / (prep_total * 2), desc=f"Preprocessing {prep_done}/{prep_total}")
            try:
                prepped_items.append(future.result())
            except Exception as exc:
                failures.append(f"{image}: preprocess failed: {exc}")

    prepped_items.sort(key=lambda r: r["relative"])

    processing_workers = 1 if config.removal_engine == "photoshop" else config.workers
    with ThreadPoolExecutor(max_workers=processing_workers) as executor:
        futures = {executor.submit(process_one, item, output_root, config): item for item in prepped_items}
        proc_done = 0
        proc_total = max(1, len(futures))
        for future in as_completed(futures):
            item = futures[future]
            proc_done += 1
            progress(0.5 + (proc_done / proc_total) * 0.5, desc=f"Photoshop/Crop {proc_done}/{proc_total}")
            try:
                results.append(future.result())
            except Exception as exc:
                failures.append(f"{item['source']}: {exc}")

    results.sort(key=lambda r: r["relative"])

    report = {
        "total": len(images),
        "processed": len(results),
        "failed": len(failures),
        "removal_engine": config.removal_engine,
        "workers_used": 1 if config.removal_engine == "photoshop" else config.workers,
        "preprocessed": len(prepped_items),
        "photoshop_exe": config.photoshop_exe or "",
        "photoshop_automation_mode": config.photoshop_automation_mode,
        "failures": failures,
        "issues": [asdict(HeadshotIssue(r["relative"], r["warnings"])) for r in results if r["warnings"]],
        "log": [
            f"Preprocessed {len(prepped_items)} files outside Photoshop.",
            f"Processed {len(results)} files through removal/crop stage.",
        ],
    }

    preview_paths = [r["output"] for r in results]
    ui_cache = Path.cwd() / ".ui_cache"
    ui_cache.mkdir(parents=True, exist_ok=True)

    gallery_items: list[tuple[str, str]] = []
    for i, row in enumerate(results):
        out = Path(row["output"])
        ui_name = f"{i:05d}_{out.name}"
        ui_path = ui_cache / ui_name
        try:
            shutil.copy2(out, ui_path)
            row["ui_output"] = str(ui_path)
            gallery_items.append((str(ui_path), out.name))
        except Exception:
            row["ui_output"] = row["output"]
            gallery_items.append((row["output"], out.name))

    report_text = json.dumps(report, indent=2)

    if not preview_paths:
        return (
            report_text,
            [],
            json.dumps(results),
            gr.update(choices=[], value=None),
            "",
            blank,
            blank,
            blank,
            "",
            "",
        )

    first = results[0].get("ui_output", preview_paths[0])
    src, actual_out = find_paths_from_results(first, json.dumps(results))
    w, g, b = show_preview(first)
    return (
        report_text,
        gallery_items,
        json.dumps(results),
        gr.update(choices=[r.get("ui_output", r["output"]) for r in results], value=first),
        first,
        w,
        g,
        b,
        src,
        actual_out,
    )


def show_preview(image_path: str) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    blank = np.full((250, 250, 3), 35, dtype=np.uint8)
    if not image_path:
        return blank, blank, blank

    if not Path(image_path).exists():
        return blank, blank, blank

    return (
        compose_preview(image_path, (255, 255, 255)),
        compose_preview(image_path, (127, 127, 127)),
        compose_preview(image_path, (0, 0, 0)),
    )


def find_paths_from_results(image_path: str, raw_results: str) -> tuple[str, str]:
    if not image_path or not raw_results:
        return "", ""
    rows = json.loads(raw_results)
    for row in rows:
        if row.get("ui_output") == image_path or row.get("output") == image_path:
            return row.get("source", ""), row.get("output", "")
    return "", ""


def extract_gallery_path(evt: gr.SelectData) -> str:
    value = evt.value
    if isinstance(value, str):
        return value
    if isinstance(value, (list, tuple)) and len(value) > 0:
        return str(value[0])
    if isinstance(value, dict):
        image_value = value.get("image")
        if isinstance(image_value, str):
            return image_value
        if isinstance(image_value, (list, tuple)) and image_value:
            return str(image_value[0])
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

    default_photoshop = discover_photoshop_exe()

    with gr.Blocks(title="Pro Headshot Cutout Studio") as demo:
        gr.Markdown("## Pro Headshot Cutout Studio")
        gr.Markdown(
            "Photoshop-first batch headshot extraction (250x250 PNG), with structure-preserving output and QA checks."
        )

        with gr.Row():
            with gr.Column(scale=1, elem_classes=["panel"]):
                input_folder = gr.Textbox(label="Input Folder", placeholder="C:/path/to/source/root")
                with gr.Row():
                    browse_input_btn = gr.Button("Browse Input Folder")
                output_folder = gr.Textbox(label="Output Folder", placeholder="C:/path/to/output/root")
                with gr.Row():
                    browse_output_btn = gr.Button("Browse Output Folder")
                photoshop_exe_input = gr.Textbox(
                    label="Photoshop Executable Path",
                    value=default_photoshop,
                    placeholder="C:/Program Files/Adobe/Adobe Photoshop 2024/Photoshop.exe",
                )
                removal_engine = gr.Radio(
                    choices=["photoshop", "auto", "rembg"],
                    value="photoshop",
                    label="Background Removal Engine",
                    info="photoshop = strict Photoshop Remove Background, auto = Photoshop then rembg fallback",
                )
                workers = gr.Slider(1, max(2, os.cpu_count() or 8), value=max(1, (os.cpu_count() or 4) - 1), step=1, label="Parallel Preprocess Workers")
                max_side = gr.Slider(800, 4500, value=1800, step=100, label="Pre-resize Max Side (outside Photoshop)")
                run_btn = gr.Button("Process Batch", variant="primary")
                report = gr.Code(label="Batch Report / Issue Detection", language="json")

            with gr.Column(scale=2, elem_classes=["panel"]):
                gallery = gr.Gallery(label="Processed Outputs", columns=5, rows=2, height=320, object_fit="contain")
                selected = gr.Textbox(label="Selected Output PNG", interactive=False)
                picker = gr.Dropdown(label="Pick Output (stable preview)", choices=[], interactive=True)
                state_results = gr.State("[]")
                source_for_open = gr.Textbox(label="Matched Source", interactive=False)
                actual_output_for_open = gr.Textbox(label="Matched Output", interactive=False, visible=False)

                with gr.Row():
                    open_ps = gr.Button("Open Source + Output in Photoshop")
                    open_status = gr.Textbox(label="Photoshop Launch Status", interactive=False)

                with gr.Row(elem_id="hero-preview"):
                    preview_white = gr.Image(label="White", interactive=False, height=460)
                    preview_grey = gr.Image(label="Grey", interactive=False, height=460)
                    preview_black = gr.Image(label="Black", interactive=False, height=460)

        browse_input_btn.click(fn=browse_for_directory, inputs=input_folder, outputs=input_folder)
        browse_output_btn.click(fn=browse_for_directory, inputs=output_folder, outputs=output_folder)

        run_btn.click(
            fn=run_batch,
            inputs=[input_folder, output_folder, workers, max_side, removal_engine, photoshop_exe_input],
            outputs=[report, gallery, state_results, picker, selected, preview_white, preview_grey, preview_black, source_for_open, actual_output_for_open],
        )

        gallery.select(fn=extract_gallery_path, outputs=selected)
        picker.change(fn=lambda v: v or "", inputs=picker, outputs=selected)
        selected.change(fn=show_preview, inputs=selected, outputs=[preview_white, preview_grey, preview_black])
        selected.change(fn=find_paths_from_results, inputs=[selected, state_results], outputs=[source_for_open, actual_output_for_open])
        open_ps.click(fn=open_in_photoshop, inputs=[source_for_open, actual_output_for_open, photoshop_exe_input], outputs=open_status)

    return demo


if __name__ == "__main__":
    app = build_ui()
    app.launch(server_name="0.0.0.0", server_port=7860, theme=gr.themes.Soft(primary_hue="indigo", secondary_hue="slate", neutral_hue="slate", radius_size=gr.themes.sizes.radius_lg), css=""".gradio-container {background: #0f172a !important; color: #e2e8f0 !important;} .panel {background: #111827 !important; border: 1px solid #334155 !important; border-radius: 16px !important; padding: 12px;} #hero-preview img {min-height: 460px; object-fit: contain; background: #000;}""", allowed_paths=[str(Path.cwd())])
