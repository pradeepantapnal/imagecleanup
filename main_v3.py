"""
Intelligent Photo Cleanup Engine v3.2
======================================
Single unified solution — one script, one command, one output.

  python main_v3.py --folder "D:\\Photos"
  python main_v3.py --folder "D:\\Photos" --enable-vision --vision-limit 200
  python main_v3.py --folder "D:\\Photos" --dry-run
  python main_v3.py --folder "D:\\Photos" --limit 50 --enable-vision

All metrics computed inline during Stage 1 (blur, BRISQUE-approx, composite,
colorfulness, brightness, contrast, resolution). No intermediate files.

Output: 6-column Excel (Filename, Score/100, Decision, Category, Caption, Reason)
sorted worst-first, color-coded. Also CSV.

Optional:
  --enable-vision     Vision LLM on ambiguous photos (score 35-65 band)
  --enable-faces      Face detection/clustering
  --limit N           Process N images only
  --dry-run           Score distribution preview, no files written
  --vision-limit N    Max images to vision LLM (default: 20)
  --vision-model M    Primary Ollama model (default: llava)
  --metrics-excel X   Use external metrics workbook
  --generate-metrics  Generate metrics workbook and exit
  --output-dir DIR    Output directory (default: output)
  --clear-cache       Wipe cache before running
"""

import os
import re
import sys
import csv
import json
import math
import base64
import hashlib
import sqlite3
import argparse
import logging
import requests
from time import perf_counter
from io import BytesIO
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager
from collections import defaultdict, Counter
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple

# ── Required ───────────────────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("pip install openpyxl"); sys.exit(1)

try:
    from PIL import Image
except ImportError:
    print("pip install pillow"); sys.exit(1)

try:
    import imagehash
except ImportError:
    print("pip install imagehash"); sys.exit(1)

# ── Optional ───────────────────────────────────────────────────────────────────
try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False
    def tqdm(it, **kw):
        return it

try:
    import clip
    import torch
    HAS_CLIP = True
except ImportError:
    HAS_CLIP = False

try:
    import face_recognition
    HAS_FACES = True
except ImportError:
    HAS_FACES = False

# ── Constants ──────────────────────────────────────────────────────────────────
VERSION          = "3.2"
SUPPORTED_EXTS   = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif", ".heic"}
DUPLICATE_THRESH = 8
BLUR_THRESH      = 80.0
EVENT_GAP_HOURS  = 6
OLLAMA_URL       = "http://localhost:11434/api/generate"
VISION_TIMEOUT   = 300
VISION_BAND_LOW  = 35
VISION_BAND_HIGH = 65

# ── Data model ─────────────────────────────────────────────────────────────────
@dataclass
class Photo:
    path: str
    filename: str = ""
    file_size: int = 0
    file_exists: bool = True
    # Inline metrics
    blur: float = 0.0
    brisque: float = 0.0
    composite_score: float = 0.0
    colorfulness: float = 0.0
    brightness: float = 0.0
    contrast: float = 0.0
    resolution: float = 0.0
    width: int = 0
    height: int = 0
    label: str = ""
    # Hashes
    md5: str = ""
    phash: str = ""
    # Groups
    dup_group: str = ""
    dup_rank: int = 0
    burst_group: str = ""
    burst_rank: int = 0
    is_burst_winner: bool = False
    # CLIP
    clip_tags: str = ""
    clip_confidence: float = 0.0
    # Event
    event_id: str = ""
    event_date: str = ""
    # Faces
    has_faces: bool = False
    face_count: int = 0
    person_ids: str = ""
    # Vision LLM
    caption: str = ""
    vision_category: str = ""
    vision_quality: str = ""
    vision_keep: str = ""
    vision_delete: str = ""
    vision_model: str = ""
    # Score & decision
    score: int = -1
    decision: str = ""
    reason: str = ""
    # Meta
    cache_key: str = ""
    error: str = ""


# ── Database ───────────────────────────────────────────────────────────────────
class DB:
    def __init__(self, path: str):
        self.conn = sqlite3.connect(path)
        cols = []
        for f in Photo.__dataclass_fields__.values():
            t = "BOOLEAN" if f.type in (bool, "bool") else \
                "INTEGER"  if f.type in (int, "int")   else \
                "REAL"     if f.type in (float, "float") else "TEXT"
            cols.append(f"{f.name} {t}")
        cols[0] = "path TEXT PRIMARY KEY"
        self.conn.execute(f"CREATE TABLE IF NOT EXISTS photos ({', '.join(cols)})")
        self.conn.commit()

    def save(self, p: Photo):
        d = asdict(p)
        cols = list(d.keys())
        sql = (f"INSERT INTO photos ({','.join(cols)}) VALUES ({','.join(['?']*len(cols))})"
               f" ON CONFLICT(path) DO UPDATE SET "
               f"{', '.join(f'{c}=excluded.{c}' for c in cols if c != 'path')}")
        self.conn.execute(sql, list(d.values()))
        self.conn.commit()

    def load(self, path: str) -> Optional[Photo]:
        cur = self.conn.execute("SELECT * FROM photos WHERE path=?", (path,))
        row = cur.fetchone()
        if row:
            cols = [d[0] for d in cur.description]
            try:
                return Photo(**dict(zip(cols, row)))
            except TypeError:
                return None
        return None

    def close(self):
        self.conn.close()


# ── Logging ────────────────────────────────────────────────────────────────────
def setup_log(output_dir: str) -> logging.Logger:
    log_path = Path(output_dir) / "engine.log"
    fmt = "%(asctime)s.%(msecs)03d %(levelname)s [%(name)s] %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S%z"
    logging.basicConfig(level=logging.INFO, format=fmt, handlers=[
        logging.FileHandler(log_path, encoding="utf-8"),
    ])
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter(fmt=fmt, datefmt=datefmt))
    for handler in logging.getLogger().handlers:
        if isinstance(handler, logging.FileHandler):
            handler.setFormatter(logging.Formatter(fmt=fmt, datefmt=datefmt))
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    logging.getLogger().addHandler(console)
    return logging.getLogger("photo")

log = logging.getLogger("photo")


@contextmanager
def stage_timer(label: str):
    started_at = datetime.now().astimezone()
    tick = perf_counter()
    log.info(f"[START] {label} @ {started_at.isoformat(timespec='milliseconds')}")
    try:
        yield
    finally:
        finished_at = datetime.now().astimezone()
        elapsed = perf_counter() - tick
        log.info(
            f"[END]   {label} @ {finished_at.isoformat(timespec='milliseconds')} "
            f"(elapsed={elapsed:.2f}s)"
        )


# ── Image metrics ──────────────────────────────────────────────────────────────
def compute_blur(path: Path) -> float:
    """Laplacian variance — higher = sharper."""
    try:
        if HAS_CV2:
            img = cv2.imread(str(path))
            if img is None:
                return -1.0
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            return float(cv2.Laplacian(gray, cv2.CV_64F).var())
        with Image.open(path) as img:
            px = list(img.convert("L").resize((256, 256)).getdata())
            m = sum(px) / len(px)
            return float(sum((p - m) ** 2 for p in px) / len(px))
    except Exception:
        return -1.0


def compute_brisque_approx(path: Path) -> float:
    """
    BRISQUE-like no-reference quality estimate using local normalized
    luminance coefficients (MSCN). Lower = better quality.
    Approximates OpenCV's BRISQUE without contrib module.
    Range: roughly 0-100, where <15 is excellent, >50 is poor.
    """
    try:
        if HAS_CV2:
            img = cv2.imread(str(path), cv2.IMREAD_GRAYSCALE)
            if img is None:
                return 50.0
            img = cv2.resize(img, (256, 256)).astype(np.float64)
            mu = cv2.GaussianBlur(img, (7, 7), 7 / 6)
            mu_sq = mu * mu
            sigma = cv2.GaussianBlur(img * img, (7, 7), 7 / 6)
            sigma = np.sqrt(np.abs(sigma - mu_sq)) + 1e-7
            mscn = (img - mu) / sigma
            # Fit generalized Gaussian — use shape and variance as proxies
            mscn_flat = mscn.flatten()
            var = float(np.var(mscn_flat))
            kurt = float(np.mean(mscn_flat ** 4) / (var ** 2 + 1e-7)) - 3.0
            # Natural images: kurtosis ~0-3, variance ~0.5-1.5
            # Distorted: higher kurtosis, extreme variance
            quality = 20.0 + abs(kurt) * 8.0 + abs(var - 1.0) * 15.0
            return max(0.0, min(100.0, quality))
        else:
            # PIL fallback — cruder estimate from local variance stats
            with Image.open(path) as img:
                px = list(img.convert("L").resize((256, 256)).getdata())
            mean = sum(px) / len(px)
            var = sum((p - mean) ** 2 for p in px) / len(px)
            # Low variance = flat/blurry = high BRISQUE
            # Very high variance = noisy = high BRISQUE
            if var < 500:
                return min(100.0, 80.0 - var * 0.1)
            elif var > 4000:
                return min(100.0, 20.0 + (var - 4000) * 0.01)
            else:
                return max(5.0, 50.0 - (var - 500) * 0.012)
    except Exception:
        return 50.0


def compute_colorfulness(img: Image.Image) -> float:
    """Hasler & Susstrunk colorfulness metric."""
    try:
        data = list(img.convert("RGB").resize((256, 256)).getdata())
        r = [p[0] for p in data]
        g = [p[1] for p in data]
        b = [p[2] for p in data]
        n = len(r)
        rg = [r[i] - g[i] for i in range(n)]
        yb = [0.5 * (r[i] + g[i]) - b[i] for i in range(n)]
        rg_m = sum(rg) / n
        yb_m = sum(yb) / n
        rg_s = math.sqrt(sum((v - rg_m) ** 2 for v in rg) / n)
        yb_s = math.sqrt(sum((v - yb_m) ** 2 for v in yb) / n)
        return math.sqrt(rg_s ** 2 + yb_s ** 2) + 0.3 * math.sqrt(rg_m ** 2 + yb_m ** 2)
    except Exception:
        return 0.0


def compute_brightness_contrast(img: Image.Image) -> Tuple[float, float]:
    """Luminance mean (brightness) and std-dev (contrast)."""
    try:
        px = list(img.convert("L").resize((256, 256)).getdata())
        mean = sum(px) / len(px)
        std = math.sqrt(sum((p - mean) ** 2 for p in px) / len(px))
        return mean, std
    except Exception:
        return 0.0, 0.0


def compute_composite(blur: float, brisque: float, resolution: float,
                       colorfulness: float, contrast: float) -> Tuple[float, str]:
    """
    Composite quality score (0-100k range for spread) and label.
    Combines all inline metrics into a single figure for relative ranking.
    """
    blur_norm = min(100.0, max(0.0, blur / 10.0)) if blur > 0 else 0.0
    brisque_norm = max(0.0, 100.0 - brisque)  # invert: lower brisque = higher score
    res_norm = min(100.0, resolution / 20000.0)
    color_norm = min(100.0, colorfulness)
    contrast_norm = min(100.0, contrast * 1.5)

    composite = (blur_norm * 0.35 +
                 brisque_norm * 0.25 +
                 res_norm * 0.15 +
                 color_norm * 0.15 +
                 contrast_norm * 0.10) * 1000.0

    if composite > 70000 and blur > 300:
        label = "GOOD"
    elif composite > 40000 and blur > 80:
        label = "AVERAGE"
    else:
        label = "POOR"

    return composite, label


# ── Helpers ────────────────────────────────────────────────────────────────────
def find_images(folder: str, limit: Optional[int]) -> List[Path]:
    imgs = set()
    for ext in SUPPORTED_EXTS:
        imgs.update(Path(folder).rglob(f"*{ext}"))
        imgs.update(Path(folder).rglob(f"*{ext.upper()}"))
    result = sorted(imgs)
    return result[:limit] if limit else result


def file_md5(path: Path) -> str:
    try:
        h = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return ""


def file_phash(path: Path) -> str:
    try:
        with Image.open(path) as img:
            return str(imagehash.phash(img.convert("RGB")))
    except Exception:
        return ""


def exif_ts(path: Path) -> Optional[datetime]:
    try:
        from PIL.ExifTags import TAGS
        with Image.open(path) as img:
            exif = img._getexif()
            if exif:
                for tag, val in exif.items():
                    if TAGS.get(tag) == "DateTimeOriginal":
                        return datetime.strptime(val, "%Y:%m:%d %H:%M:%S")
    except Exception:
        pass
    return None


def filename_ts(path: Path) -> Optional[datetime]:
    for pat in [r'(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})', r'(\d{4})(\d{2})(\d{2})']:
        m = re.search(pat, path.name)
        if m:
            try:
                g = list(map(int, m.groups()))
                return datetime(*g) if len(g) == 6 else datetime(g[0], g[1], g[2])
            except ValueError:
                pass
    return None


def get_ts(path: Path) -> datetime:
    return exif_ts(path) or filename_ts(path) or datetime.fromtimestamp(path.stat().st_mtime)


def to_b64(path: Path, max_size=(800, 800)) -> str:
    try:
        resample = getattr(Image, 'LANCZOS', Image.Resampling.LANCZOS)
        with Image.open(path) as img:
            img = img.convert("RGB")
            img.thumbnail(max_size, resample)
            buf = BytesIO()
            img.save(buf, format="JPEG", quality=80)
            return base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return ""


def parse_json(raw: str) -> dict:
    raw = re.sub(r'^```(?:json)?\s*', '', raw.strip(), flags=re.MULTILINE)
    raw = re.sub(r'```\s*$', '', raw.strip(), flags=re.MULTILINE)
    raw = raw.strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if m:
            try:
                return json.loads(m.group())
            except Exception:
                pass
    raise ValueError(f"Cannot parse JSON: {raw[:100]}")


def cache_key(folder: str, metrics_excel: Optional[str] = None) -> str:
    folder_abs = str(Path(folder).resolve())
    metrics_abs = str(Path(metrics_excel).resolve()) if metrics_excel else "inline_metrics"
    return hashlib.md5(f"{folder_abs}|{metrics_abs}".encode()).hexdigest()[:12]


def _to_float(val, default: float = 0.0) -> float:
    try:
        if val is None or val == "":
            return default
        return float(val)
    except Exception:
        return default


def _to_int(val, default: int = 0) -> int:
    try:
        if val is None or val == "":
            return default
        return int(val)
    except Exception:
        return default


def _norm_path(s: str) -> str:
    return str(s).strip().replace("\\", "/").lower()


def load_metrics_excel(path: str) -> Dict[str, dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(min_row=1, max_row=1, values_only=True)
    header_row = next(rows, None) or []
    headers = {str(h).strip().lower(): idx for idx, h in enumerate(header_row) if h is not None}

    by_path: Dict[str, dict] = {}
    by_filename: Dict[str, dict] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        path_val = row[headers["path"]] if "path" in headers else None
        file_val = row[headers["filename"]] if "filename" in headers else None
        if path_val is None and file_val is None:
            continue

        data = {
            "blur": _to_float(row[headers["blur"]], 0.0) if "blur" in headers else 0.0,
            "brisque": _to_float(row[headers["brisque"]], 0.0) if "brisque" in headers else 0.0,
            "composite_score": _to_float(row[headers["composite_score"]], 0.0) if "composite_score" in headers else 0.0,
            "label": str(row[headers["label"]]).strip().upper() if "label" in headers and row[headers["label"]] is not None else "",
            "resolution": _to_float(row[headers["resolution"]], 0.0) if "resolution" in headers else 0.0,
            "width": _to_int(row[headers["width"]], 0) if "width" in headers else 0,
            "height": _to_int(row[headers["height"]], 0) if "height" in headers else 0,
            "colorfulness": _to_float(row[headers["colorfulness"]], 0.0) if "colorfulness" in headers else 0.0,
            "brightness": _to_float(row[headers["brightness"]], 0.0) if "brightness" in headers else 0.0,
            "contrast": _to_float(row[headers["contrast"]], 0.0) if "contrast" in headers else 0.0,
        }

        if path_val:
            by_path[_norm_path(str(path_val))] = data
        if file_val:
            by_filename[str(file_val).strip().lower()] = data

    return {"by_path": by_path, "by_filename": by_filename}


# ── Stage 1: Load & compute all metrics inline ────────────────────────────────
def _match_metrics_row(img: Path, folder: str, metrics: Dict[str, dict]) -> Optional[dict]:
    by_path = metrics.get("by_path", {})
    by_filename = metrics.get("by_filename", {})

    abs_key = _norm_path(str(img))
    if abs_key in by_path:
        return by_path[abs_key]

    try:
        rel_key = _norm_path(str(img.relative_to(Path(folder))))
        if rel_key in by_path:
            return by_path[rel_key]
    except Exception:
        pass

    return by_filename.get(img.name.lower())


def stage_load(images: List[Path], db: DB, ck: str,
               folder: str, metrics: Optional[Dict[str, dict]] = None) -> List[Photo]:
    log.info(f"S1 load+metrics: {len(images)} images")
    photos = []
    for img in tqdm(images, desc="S1 metrics"):
        cached = db.load(str(img))
        if cached and cached.cache_key == ck and cached.md5:
            photos.append(cached)
            continue

        p = Photo(path=str(img), filename=img.name,
                  file_size=img.stat().st_size, cache_key=ck)
        try:
            with Image.open(img) as pil_img:
                rgb = pil_img.convert("RGB")
                p.width, p.height = pil_img.size
                p.resolution = p.width * p.height
                p.colorfulness = round(compute_colorfulness(rgb), 2)
                brt, ctr = compute_brightness_contrast(rgb)
                p.brightness = round(brt, 2)
                p.contrast = round(ctr, 2)
        except Exception as e:
            p.error = f"PIL:{e}"

        p.blur = round(compute_blur(img), 2)
        p.brisque = round(compute_brisque_approx(img), 2)
        p.composite_score, p.label = compute_composite(
            p.blur, p.brisque, p.resolution, p.colorfulness, p.contrast
        )
        p.composite_score = round(p.composite_score, 2)

        if metrics:
            m = _match_metrics_row(img, folder, metrics)
            if m:
                p.blur = round(m.get("blur", p.blur), 2)
                p.brisque = round(m.get("brisque", p.brisque), 2)
                p.composite_score = round(m.get("composite_score", p.composite_score), 2)
                p.label = m.get("label") or p.label
                if m.get("resolution", 0) > 0:
                    p.resolution = m.get("resolution", p.resolution)
                if m.get("width", 0) > 0:
                    p.width = m.get("width", p.width)
                if m.get("height", 0) > 0:
                    p.height = m.get("height", p.height)
                if m.get("colorfulness", 0) > 0:
                    p.colorfulness = m.get("colorfulness", p.colorfulness)
                if m.get("brightness", 0) > 0:
                    p.brightness = m.get("brightness", p.brightness)
                if m.get("contrast", 0) > 0:
                    p.contrast = m.get("contrast", p.contrast)

        photos.append(p)
        db.save(p)

    log.info(f"             : {len(photos)} loaded with inline metrics")
    return photos


# ── Stage 2: Duplicates (MD5) ──────────────────────────────────────────────────
def stage_duplicates(photos: List[Photo], db: DB) -> List[Photo]:
    log.info("S2 duplicates: scanning")
    md5_groups = defaultdict(list)
    for p in tqdm(photos, desc="S2 MD5"):
        if not p.md5:
            p.md5 = file_md5(Path(p.path))
            db.save(p)
        if p.md5:
            md5_groups[p.md5].append(p)

    dup_count = 0
    for h, group in md5_groups.items():
        if len(group) > 1:
            group.sort(key=lambda x: (-x.composite_score, -x.blur, -x.file_size))
            for rank, p in enumerate(group):
                p.dup_group = f"DUP_{h[:8]}"
                p.dup_rank = rank + 1
                db.save(p)
            dup_count += 1
    log.info(f"             : {dup_count} duplicate groups")
    return photos


# ── Stage 3: Burst detection (pHash + timestamp) ──────────────────────────────
def stage_burst(photos: List[Photo], db: DB) -> List[Photo]:
    log.info("S3 burst     : scanning")
    valid = [p for p in photos if p.file_exists and not p.dup_group]
    for p in tqdm(valid, desc="S3 pHash"):
        if not p.phash:
            p.phash = file_phash(Path(p.path))
            db.save(p)

    # Pre-compute timestamps once (avoids millions of EXIF re-reads)
    log.info("             : computing timestamps")
    ts_map = {}
    for p in valid:
        ts_map[p.path] = get_ts(Path(p.path))

    # Sort by timestamp so we only compare nearby photos
    valid_sorted = sorted(valid, key=lambda p: ts_map[p.path])

    # Pre-parse hashes once
    hash_map = {}
    for p in valid_sorted:
        if p.phash:
            try:
                hash_map[p.path] = imagehash.hex_to_hash(p.phash)
            except Exception:
                pass

    groups, processed = [], set()
    for i, pa in enumerate(valid_sorted):
        if pa.path in processed or pa.path not in hash_map:
            continue
        ts_a = ts_map[pa.path]
        ha = hash_map[pa.path]
        group = [pa]
        processed.add(pa.path)
        # Only look forward in time-sorted list; stop when gap > 5 min
        for pb in valid_sorted[i + 1:]:
            if (ts_map[pb.path] - ts_a).total_seconds() > 300:
                break  # sorted by time, so all remaining are also >5 min away
            if pb.path in processed or pb.path not in hash_map:
                continue
            if ha - hash_map[pb.path] <= DUPLICATE_THRESH:
                group.append(pb)
                processed.add(pb.path)
        if len(group) > 1:
            groups.append(group)

    for gid, group in enumerate(groups):
        group.sort(key=lambda x: (-x.composite_score, -x.blur, -x.file_size))
        for rank, p in enumerate(group):
            p.burst_group = f"BURST_{gid:03d}"
            p.burst_rank = rank + 1
            p.is_burst_winner = (rank == 0)
            db.save(p)
    log.info(f"             : {len(groups)} burst groups")
    return photos


# ── Stage 4: CLIP tagging ─────────────────────────────────────────────────────
def stage_clip(photos: List[Photo], db: DB) -> List[Photo]:
    if not HAS_CLIP:
        log.info("S4 CLIP      : skipped (not installed)")
        return photos

    try:
        device = "cuda" if torch.cuda.is_available() else "cpu"
        model, preprocess = clip.load("ViT-B/32", device=device)
        categories = ["indoor", "outdoor", "daytime", "nighttime", "people", "landscape",
                      "document", "food", "animal", "architecture", "screenshot", "event"]
        labels = ["indoor", "outdoor", "day", "night", "people", "landscape",
                  "document", "food", "animal", "architecture", "screenshot", "event"]
        tokens = clip.tokenize([f"a photo of {c}" for c in categories]).to(device)
        valid = [p for p in photos if p.file_exists and not p.clip_tags]
        log.info(f"S4 CLIP      : {len(valid)} images ({device})")
        for p in tqdm(valid, desc="S4 CLIP"):
            try:
                with Image.open(p.path) as img:
                    t = preprocess(img.convert("RGB")).unsqueeze(0).to(device)
                with torch.no_grad():
                    lp = (100.0 * model.encode_image(t) @ model.encode_text(tokens).T).softmax(dim=-1)
                probs = lp.cpu().numpy()[0]
                top = [labels[i] for i in probs.argsort()[-3:][::-1] if probs[i] > 0.2]
                p.clip_tags = ", ".join(top)
                p.clip_confidence = float(max(probs))
                db.save(p)
            except Exception as e:
                p.error = f"CLIP:{e}"
                db.save(p)
    except Exception as e:
        log.error(f"S4 CLIP failed: {e}")
    return photos


# ── Stage 5: Event grouping ───────────────────────────────────────────────────
def stage_events(photos: List[Photo], db: DB) -> List[Photo]:
    valid = sorted(
        [(p, get_ts(Path(p.path))) for p in photos if p.file_exists],
        key=lambda x: x[1]
    )
    groups, current, last_ts = [], [], None
    for p, ts in valid:
        if last_ts is None or (ts - last_ts).total_seconds() / 3600 > EVENT_GAP_HOURS:
            if current:
                groups.append(current)
            current = []
        current.append((p, ts))
        last_ts = ts
    if current:
        groups.append(current)

    for gid, group in enumerate(groups):
        date_str = group[0][1].strftime("%Y-%m-%d")
        eid = f"{date_str}_{chr(65 + gid % 26)}"
        for p, ts in group:
            p.event_id = eid
            p.event_date = date_str
            db.save(p)
    log.info(f"S5 events    : {len(groups)} groups")
    return photos


# ── Stage 6: Face detection ───────────────────────────────────────────────────
def stage_faces(photos: List[Photo], db: DB) -> List[Photo]:
    if not HAS_FACES:
        log.info("S6 faces     : skipped (not installed)")
        return photos
    valid = [p for p in photos if p.file_exists and not p.has_faces and p.face_count == 0]
    log.info(f"S6 faces     : {len(valid)} images")
    all_encs = []
    for p in tqdm(valid, desc="S6 faces"):
        try:
            img = face_recognition.load_image_file(p.path)
            locs = face_recognition.face_locations(img)
            encs = face_recognition.face_encodings(img, locs)
            p.has_faces = len(locs) > 0
            p.face_count = len(locs)
            for e in encs:
                all_encs.append((e, p))
            db.save(p)
        except Exception as e:
            p.error = f"face:{e}"
            db.save(p)

    clusters = []
    for enc, p in all_encs:
        assigned = False
        for cid, cencs in clusters:
            if min(face_recognition.face_distance(cencs, enc)) < 0.55:
                cencs.append(enc)
                lbl = f"Person_{chr(65 + cid % 26)}"
                existing = p.person_ids.split(", ") if p.person_ids else []
                if lbl not in existing:
                    p.person_ids = ", ".join(existing + [lbl])
                assigned = True
                break
        if not assigned:
            cid = len(clusters)
            clusters.append((cid, [enc]))
            lbl = f"Person_{chr(65 + cid % 26)}"
            existing = p.person_ids.split(", ") if p.person_ids else []
            if lbl not in existing:
                p.person_ids = ", ".join(existing + [lbl])
        db.save(p)
    log.info(f"             : {len(clusters)} persons found")
    return photos


# ── Scoring engine ─────────────────────────────────────────────────────────────
def compute_scores(photos: List[Photo]) -> List[Photo]:
    """
    Score 0-100. Relative (folder percentiles) + absolute signals.

    Point budget:
      Blur         25 pts  (relative to folder p75)
      BRISQUE      15 pts  (absolute — lower is better)
      Composite    15 pts  (relative to folder p75)
      Label        10 pts  (GOOD/AVERAGE/POOR from composite)
      Burst winner 10 pts
      Event unique 10 pts
      Vision       10 pts  (x2.5 weight when no vision, to compensate)
                          (excellent=+20, good=+12, avg=0, poor=-15)
    Total possible: ~100+ before clamping

    Hard overrides:
      Exact duplicate non-winner -> score=10, REMOVE
      Burst non-winner           -> score=20, REMOVE
    """
    # Folder percentiles for relative scoring
    blurs = sorted([p.blur for p in photos if p.blur > 0])
    composites = sorted([p.composite_score for p in photos if p.composite_score > 0])
    blur_p75 = blurs[int(len(blurs) * 0.75)] if blurs else 200.0
    comp_p75 = composites[int(len(composites) * 0.75)] if composites else 50000.0

    # Pre-compute event counts once
    event_counts = Counter(p.event_id for p in photos if p.event_id)

    QUALITY_MAP = {"excellent": 20, "good": 12, "average": 0, "poor": -15}
    LABEL_MAP   = {"GOOD": 10, "AVERAGE": 0, "POOR": -10}

    # Vision weight: higher when no vision data available to compensate
    has_any_vision = any(p.vision_model for p in photos)
    vision_weight = 1.0 if has_any_vision else 2.5

    for p in photos:
        s = 50.0

        # ── Hard overrides ────────────────────────────────────────────────
        if p.dup_group and p.dup_rank > 1:
            p.score = 10
            p.decision = "REMOVE"
            p.reason = f"Exact duplicate (copy #{p.dup_rank})"
            continue

        if p.burst_group and not p.is_burst_winner:
            p.score = 20
            p.decision = "REMOVE"
            p.reason = f"Burst duplicate (rank #{p.burst_rank} of group)"
            continue

        # ── Blur: relative to folder p75 (25 pts) ────────────────────────
        if p.blur > 0:
            ratio = p.blur / blur_p75
            if ratio >= 1.5:    s += 25
            elif ratio >= 1.0:  s += 15
            elif ratio >= 0.5:  s += 5
            elif ratio >= 0.2:  s -= 10
            else:               s -= 20

        # ── BRISQUE: absolute, lower is better (15 pts) ──────────────────
        if p.brisque > 0:
            if p.brisque < 15:    s += 15
            elif p.brisque < 30:  s += 8
            elif p.brisque < 50:  s += 2
            else:                 s -= 10

        # ── Composite: relative to folder p75 (15 pts) ───────────────────
        if p.composite_score > 0:
            ratio = p.composite_score / comp_p75
            if ratio >= 1.5:    s += 15
            elif ratio >= 1.0:  s += 8
            elif ratio >= 0.5:  s += 2
            else:               s -= 5

        # ── Label (10 pts) ───────────────────────────────────────────────
        s += LABEL_MAP.get(p.label.upper() if p.label else "", 0)

        # ── Burst winner bonus (10 pts) ──────────────────────────────────
        if p.is_burst_winner:
            s += 10

        # ── Event uniqueness (10 pts) ────────────────────────────────────
        if p.event_id and event_counts[p.event_id] == 1:
            s += 10

        # ── Vision quality (10 pts base, weighted) ───────────────────────
        if p.vision_quality:
            s += QUALITY_MAP.get(p.vision_quality.lower(), 0) * vision_weight

        p.score = max(0, min(100, int(s)))

    # ── Assign decisions ─────────────────────────────────────────────────
    # Adaptive thresholds: only tighten when vision covered >50% of review band
    review_band = [p for p in photos if not p.decision and VISION_BAND_LOW <= p.score <= VISION_BAND_HIGH]
    vision_count = sum(1 for p in review_band if p.vision_model)
    vision_coverage = vision_count / max(len(review_band), 1)
    if vision_coverage > 0.5:
        keep_thresh, remove_thresh = 70, 30
    else:
        keep_thresh, remove_thresh = 65, 35

    for p in photos:
        if p.decision:
            continue

        if p.score >= keep_thresh:
            p.decision = "KEEP"
            parts = []
            if p.is_burst_winner:                    parts.append("best in burst")
            if p.blur > blur_p75:                    parts.append("sharp")
            if p.brisque > 0 and p.brisque < 20:    parts.append("low noise")
            if p.label == "GOOD":                    parts.append("good quality")
            if p.vision_quality in ("excellent", "good"):
                parts.append(f"vision:{p.vision_quality}")
            p.reason = ", ".join(parts) if parts else "high quality score"

        elif p.score <= remove_thresh:
            p.decision = "REMOVE"
            parts = []
            if p.blur > 0 and p.blur < BLUR_THRESH: parts.append("blurry")
            if p.brisque > 50:                       parts.append("high noise")
            if p.label == "POOR":                    parts.append("poor quality")
            if p.vision_quality == "poor":           parts.append("vision:poor")
            p.reason = ", ".join(parts) if parts else "low quality score"

        else:
            p.decision = "REVIEW"
            p.reason = "ambiguous -- manual review suggested"

    scored = [p for p in photos if p.score >= 0]
    if scored:
        log.info(f"Scores       : min={min(p.score for p in scored)} "
                 f"max={max(p.score for p in scored)} "
                 f"avg={sum(p.score for p in scored) // len(scored)} "
                 f"| keep>={keep_thresh} remove<={remove_thresh}")
    return photos


# ── Stage 7: Vision LLM (selective, score band 35-65) ─────────────────────────
def stage_vision(photos: List[Photo], db: DB,
                 primary: str, fallback: str, limit: int) -> List[Photo]:
    candidates = [p for p in photos
                  if p.file_exists and not p.vision_model
                  and VISION_BAND_LOW <= p.score <= VISION_BAND_HIGH][:limit]

    if not candidates:
        log.info("S7 vision    : no candidates in review band")
        return photos

    log.info(f"S7 vision    : {len(candidates)} candidates (score {VISION_BAND_LOW}-{VISION_BAND_HIGH})")

    for p in tqdm(candidates, desc="S7 vision"):
        b64 = to_b64(Path(p.path))
        if not b64:
            continue

        for model in [primary, fallback]:
            try:
                # Model-aware prompts: llava returns JSON, llama3.2-vision returns text
                if "llava" in model:
                    prompt = (
                        f'Analyze this photo. Blur={p.blur:.0f}, '
                        f'brisque={p.brisque:.0f}, composite={p.composite_score:.0f}. '
                        f'Reply ONLY with JSON, no fences: '
                        f'{{"caption":"one sentence","category":"people/landscape/food/'
                        f'document/animal/architecture/event/other","quality":"excellent/'
                        f'good/average/poor","keep_reason":"...","delete_reason":"..."}}'
                    )
                else:
                    # llama3.2-vision ignores JSON instructions, use structured text
                    prompt = (
                        "Describe this photo in one sentence.\n"
                        "Then on separate lines write:\n"
                        "QUALITY: excellent/good/average/poor\n"
                        "CATEGORY: people/landscape/food/document/animal/architecture/event/other"
                    )

                resp = requests.post(OLLAMA_URL, json={
                    "model": model, "prompt": prompt,
                    "images": [b64], "stream": False,
                    "options": {"temperature": 0.1, "num_predict": 256}
                }, timeout=VISION_TIMEOUT)
                resp.raise_for_status()
                raw = resp.json().get("response", "").strip()

                if "llava" in model:
                    data = parse_json(raw)
                    p.caption         = data.get("caption", "")
                    p.vision_category = data.get("category", "")
                    p.vision_quality  = data.get("quality", "")
                    p.vision_keep     = data.get("keep_reason", "")
                    p.vision_delete   = data.get("delete_reason", "")
                else:
                    # Regex parser for llama3.2-vision structured text output
                    qm = re.search(r'QUALITY:\s*(excellent|good|average|poor)', raw, re.I)
                    cm = re.search(r'CATEGORY:\s*(\w+)', raw, re.I)
                    sentences = [s.strip() for s in raw.replace('\n', ' ').split('.')
                                 if len(s.strip()) > 20]
                    p.caption        = sentences[0] if sentences else ""
                    p.vision_quality = qm.group(1).lower() if qm else "average"
                    p.vision_category = cm.group(1).lower() if cm else "other"

                p.vision_model = model
                db.save(p)
                break  # success, skip fallback

            except (ValueError, json.JSONDecodeError) as e:
                log.warning(f"Vision JSON fail {Path(p.path).name} ({model}): {e}")
                p.error = f"vision_json:{model}"
            except requests.exceptions.Timeout:
                log.warning(f"Vision timeout {Path(p.path).name} ({model})")
                p.error = f"vision_timeout:{model}"
            except Exception as e:
                log.warning(f"Vision error {Path(p.path).name} ({model}): {type(e).__name__}")
                p.error = f"vision_err:{model}"

    # Re-score with vision data incorporated
    photos = compute_scores(photos)
    log.info(f"             : {sum(1 for p in photos if p.vision_model)} processed")
    return photos


# ── Output: Excel ──────────────────────────────────────────────────────────────
def write_excel(photos: List[Photo], path: str):
    COLS = [
        ("filename",        "Filename",   28, "left"),
        ("score",           "Score /100",  10, "center"),
        ("decision",        "Decision",   12, "center"),
        ("vision_category", "Category",   14, "center"),
        ("caption",         "Caption",    50, "left"),
        ("reason",          "Reason",     35, "left"),
    ]

    DECISION_STYLE = {
        "KEEP":   ("006100", "C6EFCE"),
        "REMOVE": ("9C0006", "FFC7CE"),
        "REVIEW": ("7D6608", "FFEB9C"),
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Photo Cleanup"

    thin = Border(
        bottom=Side(style="thin", color="DDDDDD"),
        right=Side(style="thin", color="DDDDDD")
    )

    for ci, (_, label, width, _) in enumerate(COLS, 1):
        c = ws.cell(row=1, column=ci, value=label)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill(start_color="1A3C6E", end_color="1A3C6E", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

    # Sort strictly by score ascending (worst first)
    sorted_photos = sorted(
        [p for p in photos if p.file_exists],
        key=lambda p: (p.score, p.filename.lower())
    )

    alt_fill = [
        PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
        PatternFill(start_color="F5F7FA", end_color="F5F7FA", fill_type="solid"),
    ]

    for ri, p in enumerate(sorted_photos, 2):
        bg = alt_fill[ri % 2]
        for ci, (fld, _, _, align) in enumerate(COLS, 1):
            val = getattr(p, fld, "") or ""
            if fld == "score":
                val = int(val) if val != "" else ""
            # CLIP tags as fallback category when vision hasn't run
            if fld == "vision_category" and not val and p.clip_tags:
                val = p.clip_tags
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = thin
            c.alignment = Alignment(
                horizontal=align, vertical="center",
                wrap_text=(fld in ("caption", "reason"))
            )

            if fld == "decision" and val in DECISION_STYLE:
                fg, bg_col = DECISION_STYLE[val]
                c.font = Font(name="Arial", bold=True, color=fg, size=10)
                c.fill = PatternFill(start_color=bg_col, end_color=bg_col, fill_type="solid")
            elif fld == "score":
                score = int(val) if val != "" else 50
                if score >= 70:     col = "C6EFCE"
                elif score >= 36:   col = "FFEB9C"
                else:               col = "FFC7CE"
                c.font = Font(name="Arial", size=10, bold=True)
                c.fill = PatternFill(start_color=col, end_color=col, fill_type="solid")
            else:
                c.font = Font(name="Arial", size=10)
                c.fill = bg

        ws.row_dimensions[ri].height = 32 if p.caption else 18

    # Summary footer
    ws.append([])
    sr = ws.max_row + 1
    keep = sum(1 for p in sorted_photos if p.decision == "KEEP")
    remove = sum(1 for p in sorted_photos if p.decision == "REMOVE")
    review = sum(1 for p in sorted_photos if p.decision == "REVIEW")
    for i, (lbl, val) in enumerate([
        ("Total", len(sorted_photos)),
        ("KEEP", keep), ("REMOVE", remove), ("REVIEW", review)
    ]):
        ws.cell(row=sr + i, column=1, value=lbl).font = Font(name="Arial", bold=True)
        ws.cell(row=sr + i, column=2, value=val).font = Font(name="Arial", bold=True)

    wb.save(path)
    log.info(f"Excel        : KEEP={keep} REMOVE={remove} REVIEW={review}")


# ── Output: CSV ────────────────────────────────────────────────────────────────
def write_csv(photos: List[Photo], path: str):
    order = {"REMOVE": 0, "REVIEW": 1, "KEEP": 2}
    rows = sorted(
        [p for p in photos if p.file_exists],
        key=lambda p: (order.get(p.decision, 3), p.score)
    )
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["filename", "score", "decision", "category", "caption", "reason", "path"])
        for p in rows:
            w.writerow([p.filename, p.score, p.decision,
                        p.vision_category or p.clip_tags,
                        p.caption, p.reason, p.path])
    log.info(f"CSV          : {len(rows)} rows")


def write_metrics_excel(photos: List[Photo], path: str, folder: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Metrics"
    headers = [
        "path", "filename", "blur", "brisque", "composite_score", "label",
        "resolution", "width", "height", "colorfulness", "brightness", "contrast"
    ]
    ws.append(headers)
    root = Path(folder)
    for p in photos:
        full = Path(p.path)
        try:
            rel = str(full.relative_to(root)).replace("/", "\\")
        except Exception:
            rel = p.path
        ws.append([
            rel, p.filename, p.blur, p.brisque, p.composite_score, p.label,
            p.resolution, p.width, p.height, p.colorfulness, p.brightness, p.contrast
        ])
    wb.save(path)
    log.info(f"Metrics XLSX : {len(photos)} rows -> {path}")


# ── Dry run ────────────────────────────────────────────────────────────────────
def print_dry_run(photos: List[Photo]):
    buckets = {"0-20": 0, "21-35": 0, "36-50": 0, "51-65": 0, "66-80": 0, "81-100": 0}
    for p in photos:
        s = p.score
        if s <= 20:     buckets["0-20"] += 1
        elif s <= 35:   buckets["21-35"] += 1
        elif s <= 50:   buckets["36-50"] += 1
        elif s <= 65:   buckets["51-65"] += 1
        elif s <= 80:   buckets["66-80"] += 1
        else:           buckets["81-100"] += 1

    keep = sum(1 for p in photos if p.decision == "KEEP")
    remove = sum(1 for p in photos if p.decision == "REMOVE")
    review = sum(1 for p in photos if p.decision == "REVIEW")

    print("\n-- DRY RUN: Score Distribution -----------------------------")
    for band, count in buckets.items():
        bar = "#" * (count * 40 // max(len(photos), 1))
        print(f"  {band:8s} |{bar:<40s}| {count}")
    print(f"\n  KEEP: {keep}  REMOVE: {remove}  REVIEW: {review}  TOTAL: {len(photos)}")
    print("  (no files written in dry-run mode)\n")


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description=f"Photo Cleanup Engine v{VERSION}")
    parser.add_argument("--folder", required=True,
                        help="Path to photo folder")
    parser.add_argument("--output-dir", default="output",
                        help="Output directory (default: output)")
    parser.add_argument("--limit", type=int, default=None,
                        help="Process only N images (default: all)")
    parser.add_argument("--enable-vision", action="store_true",
                        help="Use Ollama vision LLM on ambiguous photos")
    parser.add_argument("--enable-faces", action="store_true",
                        help="Enable face detection and clustering")
    parser.add_argument("--vision-model", default="llava",
                        help="Primary Ollama vision model (default: llava)")
    parser.add_argument("--vision-fallback", default="llama3.2-vision",
                        help="Fallback vision model")
    parser.add_argument("--vision-limit", type=int, default=20,
                        help="Max images to send to vision LLM (default: 20)")
    parser.add_argument("--metrics-excel", default=None,
                        help="Path to metrics Excel file (optional)")
    parser.add_argument("--generate-metrics", action="store_true",
                        help="Generate metrics Excel from folder scan and exit")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview score distribution, write no files")
    parser.add_argument("--clear-cache", action="store_true",
                        help="Wipe cached results before running")
    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    global log
    log = setup_log(str(out_dir))

    if args.clear_cache:
        db_path = out_dir / "cache.db"
        if db_path.exists():
            db_path.unlink()
        log.info("Cache        : cleared")

    start = datetime.now()
    ck = cache_key(args.folder, args.metrics_excel)

    log.info(f"Photo Cleanup Engine v{VERSION}")
    log.info(f"Folder       : {args.folder}")
    log.info(f"Config       : limit={args.limit or 'all'} vision={args.enable_vision} faces={args.enable_faces}")

    metrics = None
    if args.metrics_excel:
        mpath = Path(args.metrics_excel)
        if not mpath.exists():
            raise FileNotFoundError(f"--metrics-excel not found: {args.metrics_excel}")
        metrics = load_metrics_excel(str(mpath))
        log.info(f"Metrics      : loaded from {args.metrics_excel}")

    db = DB(str(out_dir / "cache.db"))

    try:
        with stage_timer("Scan images"):
            images = find_images(args.folder, args.limit)
        log.info(f"Found        : {len(images)} images")

        with stage_timer("S1 load+metrics"):
            photos = stage_load(images, db, ck, args.folder, metrics)

        if args.generate_metrics:
            metrics_name = f"{Path(args.folder).name}_metrics.xlsx"
            with stage_timer("Write metrics workbook"):
                write_metrics_excel(photos, str(out_dir / metrics_name), args.folder)
            log.info("Done         : metrics generated")
            return

        with stage_timer("S2 duplicates"):
            photos = stage_duplicates(photos, db)
        with stage_timer("S3 burst detection"):
            photos = stage_burst(photos, db)
        with stage_timer("S4 CLIP tagging"):
            photos = stage_clip(photos, db)
        with stage_timer("S5 event grouping"):
            photos = stage_events(photos, db)
        if args.enable_faces:
            with stage_timer("S6 face detection"):
                photos = stage_faces(photos, db)

        # Initial scoring before vision
        with stage_timer("Compute scores"):
            photos = compute_scores(photos)

        if args.enable_vision:
            with stage_timer("S7 vision analysis"):
                photos = stage_vision(photos, db,
                                      args.vision_model, args.vision_fallback,
                                      args.vision_limit)

        if args.dry_run:
            with stage_timer("Dry-run summary"):
                print_dry_run(photos)
        else:
            with stage_timer("Write Excel"):
                write_excel(photos, str(out_dir / "photo_cleanup.xlsx"))
            with stage_timer("Write CSV"):
                write_csv(photos, str(out_dir / "photo_cleanup.csv"))

        secs = (datetime.now() - start).total_seconds()
        log.info(f"Done         : {len(photos)} photos in {secs:.0f}s")

    finally:
        db.close()


if __name__ == "__main__":
    main()
