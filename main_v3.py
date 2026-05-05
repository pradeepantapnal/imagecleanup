"""
Intelligent Photo Cleanup Engine v3.5 - Architect Edition
Optimized for: Local VLM Fidelity, Multiprocessing, and Execution Discipline.
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
import platform
import shutil
import subprocess
from concurrent.futures import ProcessPoolExecutor
from time import perf_counter, process_time
from io import BytesIO
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager
from collections import defaultdict, Counter
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple
from PIL import Image, ExifTags

# ── Dependencies ──────────────────────────────────────────────────────────────
def check_deps():
    deps = ["openpyxl", "PIL", "imagehash", "cv2", "numpy", "tqdm"]
    # Internal check for key libs
    try:
        import openpyxl, PIL, imagehash, cv2, numpy, tqdm
    except ImportError as e:
        print(f"Missing dependency: {e}. Run: pip install openpyxl pillow imagehash opencv-python numpy tqdm")
        sys.exit(1)

# ── Constants ──────────────────────────────────────────────────────────────────
VERSION          = "3.5"
SUPPORTED_EXTS   = {".jpg", ".jpeg", ".png", ".webp", ".heic"}
BLUR_ABS_FLOOR   = 30.0  # Absolute minimum for 'sharp'
BRISQUE_ABS_CEIL = 60.0  # Absolute maximum for 'clean'
OLLAMA_URL       = "http://localhost:11434/api/generate"
VISION_BAND_LOW  = 35
VISION_BAND_HIGH = 65

@dataclass
class Photo:
    path: str
    filename: str = ""
    file_size: int = 0
    # Technical Metadata
    iso: int = 0
    exposure_time: str = ""
    width: int = 0
    height: int = 0
    resolution: float = 0.0
    # Metrics
    blur: float = 0.0
    brisque: float = 0.0
    composite_score: float = 0.0
    colorfulness: float = 0.0
    brightness: float = 0.0
    contrast: float = 0.0
    label: str = ""
    # Grouping/Logic
    md5: str = ""
    phash: str = ""
    dup_group: str = ""
    dup_rank: int = 0
    is_burst_winner: bool = False
    # Vision Data
    caption: str = ""
    vision_quality: str = ""
    vision_memorability: int = 0
    vision_model: str = ""
    # Decision
    score: int = -1
    decision: str = ""
    reason: str = ""
    cache_key: str = ""
    error: str = ""

# ── Core Metric Engine (Independent for Pickling) ──────────────────────────────

def get_technical_meta(path: Path) -> dict:
    meta = {"iso": 0, "exposure": "", "w": 0, "h": 0}
    try:
        with Image.open(path) as img:
            meta["w"], meta["h"] = img.size
            exif = {ExifTags.TAGS[k]: v for k, v in img._getexif().items() if k in ExifTags.TAGS} if img._getexif() else {}
            meta["iso"] = exif.get("ISOSpeedRatings", 0)
            meta["exposure"] = str(exif.get("ExposureTime", ""))
    except: pass
    return meta

def compute_metrics_worker(img_path_str: str) -> dict:
    """Worker function for multiprocessing."""
    path = Path(img_path_str)
    try:
        import cv2
        import numpy as np
        img = cv2.imread(img_path_str)
        if img is None: return {"error": "CV2 read fail"}
        
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        blur = float(cv2.Laplacian(gray, cv2.CV_64F).var())
        
        # BRISQUE-approx
        img_small = cv2.resize(gray, (256, 256)).astype(np.float64)
        mu = cv2.GaussianBlur(img_small, (7, 7), 7/6)
        sigma = np.sqrt(np.abs(cv2.GaussianBlur(img_small**2, (7, 7), 7/6) - mu**2)) + 1e-7
        mscn = (img_small - mu) / sigma
        brisque = float(20.0 + abs(np.mean(mscn**4)/(np.var(mscn)**2+1e-7)-3.0)*8.0)
        
        # Color/Brightness
        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        brightness = float(np.mean(hsv[:,:,2]))
        contrast = float(np.std(gray))
        
        meta = get_technical_meta(path)
        
        return {
            "path": img_path_str,
            "blur": round(blur, 2),
            "brisque": round(brisque, 2),
            "brightness": round(brightness, 2),
            "contrast": round(contrast, 2),
            "iso": meta["iso"],
            "exposure": meta["exposure"],
            "width": meta["w"],
            "height": meta["h"]
        }
    except Exception as e:
        return {"path": img_path_str, "error": str(e)}

# ── VLM Fidelity Enhancement (ROI Crop) ───────────────────────────────────────

def create_vision_tiled_b64(path: Path) -> str:
    """Creates a composite of full image + 100% detail crop to bypass VLM downsampling."""
    try:
        with Image.open(path) as img:
            img = img.convert("RGB")
            w, h = img.size
            
            # 1. Downsampled full view
            full_view = img.copy()
            full_view.thumbnail((600, 600))
            
            # 2. 100% Detail Crop (Center)
            crop_size = 300
            left = (w - crop_size) // 2
            top = (h - crop_size) // 2
            detail_crop = img.crop((left, top, left + crop_size, top + crop_size))
            
            # 3. Tile them side-by-side
            combined = Image.new('RGB', (full_view.width + detail_crop.width, max(full_view.height, detail_crop.height)))
            combined.paste(full_view, (0, 0))
            combined.paste(detail_crop, (full_view.width, 0))
            
            buf = BytesIO()
            combined.save(buf, format="JPEG", quality=85)
            return base64.b64encode(buf.getvalue()).decode()
    except: return ""

# ── Optimized DB ──────────────────────────────────────────────────────────────

class DB:
    def __init__(self, path: str):
        self.conn = sqlite3.connect(path)
        self.conn.execute("PRAGMA journal_mode=WAL") # High-performance logging
        self.conn.execute("PRAGMA synchronous=NORMAL")
        # Initialize tables using dataclass fields...
        # (Simplified for brevity, same logic as your v3.2)
        self.conn.execute("CREATE TABLE IF NOT EXISTS photos (path TEXT PRIMARY KEY, json_data TEXT)")
    
    def save_batch(self, photos: List[Photo]):
        with self.conn:
            for p in photos:
                self.conn.execute("INSERT OR REPLACE INTO photos VALUES (?, ?)", (p.path, json.dumps(asdict(p))))

    def load(self, path: str) -> Optional[Photo]:
        row = self.conn.execute("SELECT json_data FROM photos WHERE path=?", (path,)).fetchone()
        return Photo(**json.loads(row[0])) if row else None

# ── Stages ────────────────────────────────────────────────────────────────────

def stage_1_parallel(images: List[Path], db: DB, ck: str) -> List[Photo]:
    """Parallelized metric extraction."""
    photos = []
    to_process = []
    
    for img in images:
        cached = db.load(str(img))
        if cached and cached.cache_key == ck:
            photos.append(cached)
        else:
            to_process.append(str(img))
    
    if to_process:
        from tqdm import tqdm
        print(f"S1: Processing {len(to_process)} new images using {os.cpu_count()-1} cores...")
        with ProcessPoolExecutor(max_workers=os.cpu_count()-1) as executor:
            results = list(tqdm(executor.map(compute_metrics_worker, to_process), total=len(to_process)))
            
        for r in results:
            if "error" in r: continue
            p = Photo(path=r["path"], cache_key=ck, filename=Path(r["path"]).name)
            p.blur = r["blur"]
            p.brisque = r["brisque"]
            p.brightness = r["brightness"]
            p.contrast = r["contrast"]
            p.iso = r["iso"]
            p.exposure_time = r["exposure"]
            p.width = r["width"]
            p.height = r["height"]
            p.resolution = p.width * p.height
            photos.append(p)
            
        db.save_batch(photos)
    return photos

def stage_vision_v35(photos: List[Photo], db: DB, limit: int):
    """VLM analysis with detail-aware cropping."""
    candidates = [p for p in photos if VISION_BAND_LOW <= p.score <= VISION_BAND_HIGH][:limit]
    if not candidates: return
    
    from tqdm import tqdm
    for p in tqdm(candidates, desc="S7 Vision (Detail-Aware)"):
        b64 = create_vision_tiled_b64(Path(p.path))
        if not b64: continue
        
        prompt = (
            f"AUDIT MODE. Left image: full scene. Right image: 100% crop of center.\n"
            f"Metrics: Blur={p.blur}, BRISQUE={p.brisque}, ISO={p.iso}.\n"
            f"Task: Evaluate if focus is sharp in the detail crop. Ignore noise if ISO > 1600.\n"
            f"Return JSON: {{\"quality\": \"excellent/good/average/poor\", \"memorability\": 1-5, \"reason\": \"str\"}}"
        )
        
        try:
            resp = requests.post(OLLAMA_URL, json={
                "model": "llava", "prompt": prompt, "images": [b64], "stream": False,
                "options": {"temperature": 0, "num_predict": 128}
            }, timeout=60)
            data = json.loads(re.search(r'\{.*\}', resp.json()['response'], re.DOTALL).group())
            p.vision_quality = data.get("quality", "average")
            p.vision_memorability = int(data.get("memorability", 2))
            p.reason = data.get("reason", "")
            p.vision_model = "llava-v3.5-roi"
            db.save_batch([p])
        except Exception as e:
            p.error = f"VisionFail: {str(e)}"

# ── Scoring ───────────────────────────────────────────────────────────────────

def apply_architect_scoring(photos: List[Photo]):
    """Decision logic with absolute floors."""
    # Relative thresholds
    blurs = sorted([p.blur for p in photos if p.blur > 0])
    p75_blur = blurs[int(len(blurs)*0.75)] if blurs else 100
    
    for p in photos:
        s = 50
        
        # Absolute Failure Checks (Execution Discipline)
        if p.blur < BLUR_ABS_FLOOR: s -= 40
        if p.brisque > BRISQUE_ABS_CEIL: s -= 20
        
        # ISO-aware noise forgiveness
        if p.iso > 1600 and p.brisque > 40:
            s += 10 # Forgive noise if high ISO was necessary
            
        # Relative boost
        if p.blur > p75_blur: s += 20
        
        # Vision weights
        if p.vision_quality == "excellent": s += 25
        if p.vision_quality == "poor": s -= 30
        
        p.score = max(0, min(100, int(s)))
        p.decision = "KEEP" if p.score > 70 else "REMOVE" if p.score < 35 else "REVIEW"

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    check_deps()
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", required=True)
    parser.add_argument("--vision", action="store_true")
    args = parser.parse_args()
    
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    db = DB(str(out_dir / "engine_v35.db"))
    
    # Execution
    images = sorted(list(Path(args.folder).rglob("*.*"))) 
    images = [i for i in images if i.suffix.lower() in SUPPORTED_EXTS]
    
    ck = hashlib.md5(args.folder.encode()).hexdigest()[:8]
    
    photos = stage_1_parallel(images, db, ck)
    apply_architect_scoring(photos) # Pre-vision score
    
    if args.vision:
        stage_vision_v35(photos, db, limit=50)
        apply_architect_scoring(photos) # Final score
        
    # Final save and reporting...
    print(f"Processed {len(photos)} images. See output/engine_v35.db or exports.")

if __name__ == "__main__":
    main()
