# Intelligent Photo Cleanup Engine v3.2

A fully local, single-script photo cleanup pipeline. Point it at a folder, get a scored Excel telling you what to keep, review, and remove. No cloud APIs, no file deletion — recommendations only.

```
python main_v3.py --folder "D:\Photos"
```

One command. All metrics computed inline. Output: a 6-column Excel sorted worst-first.

## What it does

1. **Computes quality metrics** inline for every photo — blur (Laplacian variance), BRISQUE-approximation (MSCN-based), colorfulness, brightness, contrast, resolution
2. **Detects exact duplicates** via MD5 hashing
3. **Detects burst sequences** via pHash similarity + timestamp proximity (5-minute window)
4. **Tags scenes** with CLIP (indoor/outdoor, people/landscape, food, etc.) — GPU-accelerated when available
5. **Groups into events** by 6-hour timestamp gaps
6. **Scores each photo 0–100** using relative (folder percentiles) + absolute signals
7. **Optionally runs vision LLM** (Ollama) on ambiguous photos (score 35–65) for caption, category, quality, and memorability assessment
8. **Outputs a slim Excel** with 6 columns: Filename, Score/100, Decision, Category, Caption, Reason

## Output

### Excel (photo_cleanup.xlsx)

Sorted by score ascending (worst first), color-coded:

| Column | Description |
|--------|-------------|
| Filename | Photo name |
| Score /100 | Quality score — green (65+), yellow (36–64), red (≤35) |
| Decision | **KEEP** (green), **REMOVE** (red), or **REVIEW** (yellow) |
| Category | Scene tag from CLIP or vision LLM |
| Caption | One-line description (populated when vision runs) |
| Reason | Why — e.g. "sharp, low noise", "burst duplicate (rank #3)", "blurry" |

Summary footer shows KEEP/REMOVE/REVIEW totals.

### How to use the Excel

Start from the top (lowest scores). The REMOVE block is mostly duplicates and burst copies — confirm and delete. The REVIEW block needs manual judgment — photos near 65 are probably keepers, near 35 probably not. The KEEP block can be ignored unless you want to trim further.

## Scoring formula

Baseline: 50 points. Signals add or subtract:

| Signal | Points | Type |
|--------|--------|------|
| Blur sharpness | 25 | Relative to folder 75th percentile |
| BRISQUE quality | 15 | Absolute (lower BRISQUE = better) |
| Composite score | 15 | Relative to folder 75th percentile |
| Quality label | 10 | GOOD/AVERAGE/POOR from composite |
| Burst winner | 10 | Best photo in burst group |
| Event uniqueness | 10 | Only photo in its event |
| Vision quality | 10 | excellent/good/average/poor from LLM |
| Memorability | 15 | 1–5 semantic score from LLM |

Hard overrides bypass scoring: exact duplicate copies get score=10 REMOVE, burst non-winners get score=20 REMOVE.

Decision thresholds: KEEP ≥ 65, REMOVE ≤ 35. Thresholds tighten to 70/30 when vision has covered >50% of the ambiguous band.

## Installation

```
pip install openpyxl pillow imagehash opencv-python tqdm numpy
pip install git+https://github.com/openai/CLIP.git torch torchvision
```

For GPU-accelerated CLIP (recommended with NVIDIA GPU):
```
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu121
```

For vision LLM features, install [Ollama](https://ollama.com) and pull models:
```
ollama pull llava
ollama pull llama3.2-vision
```

## Usage

### Basic — compute everything, no LLM
```
python main_v3.py --folder "D:\Photos\Naples"
```

### With vision LLM on ambiguous photos
```
python main_v3.py --folder "D:\Photos" --enable-vision --vision-limit 200
```

### Preview score distribution without writing files
```
python main_v3.py --folder "D:\Photos" --dry-run
```

### Process a subset
```
python main_v3.py --folder "D:\Photos" --limit 100
```

### Use config file for defaults
```
python main_v3.py --folder "D:\Photos" --config config.json
```

### Generate metrics workbook (for external tools)
```
python main_v3.py --folder "D:\Photos" --generate-metrics
```

### Fresh run (clear cached results)
```
python main_v3.py --folder "D:\Photos" --clear-cache
```

## Command-line options

| Option | Default | Description |
|--------|---------|-------------|
| `--folder` | (required) | Path to photo folder |
| `--output-dir` | `output` | Where to write Excel, CSV, logs, cache |
| `--limit` | all | Process only N images |
| `--enable-vision` | off | Run Ollama vision LLM on score 35–65 band |
| `--vision-model` | `llava` | Primary Ollama model (returns JSON) |
| `--vision-fallback` | `llama3.2-vision` | Fallback model (returns structured text) |
| `--vision-limit` | 20 | Max images to send to vision |
| `--metrics-excel` | none | Optional external metrics workbook |
| `--generate-metrics` | off | Write metrics workbook and exit |
| `--enable-faces` | off | Face detection and clustering |
| `--dry-run` | off | Preview only, no files written |
| `--clear-cache` | off | Delete cache.db before running |
| `--config` | `config.json` | JSON config file (CLI overrides config) |

## Configuration file (config.json)

All CLI options can be set as defaults in `config.json`. CLI arguments override config values.

```json
{
  "output_dir": "output",
  "limit": null,
  "metrics_excel": null,
  "enable_vision": false,
  "vision_model": "llava",
  "vision_fallback": "llama3.2-vision",
  "vision_limit": 75,
  "enable_faces": false,
  "dry_run": false,
  "clear_cache": false,
  "duplicate_threshold": 8,
  "event_gap_hours": 6,
  "blur_threshold": 80.0,
  "ollama_url": "http://localhost:11434/api/generate",
  "vision_timeout": 300
}
```

## Architecture

Seven-stage pipeline with SQLite caching for crash-safe resume:

```
S1  Load + inline metrics  →  blur, BRISQUE, composite, colorfulness, brightness, contrast
S2  Duplicate detection    →  MD5 exact match grouping
S3  Burst detection        →  pHash + timestamp proximity (pre-computed, O(n) not O(n²))
S4  CLIP tagging           →  scene classification (GPU when available)
S5  Event grouping         →  6-hour timestamp gap clustering
S6  Face detection         →  (optional) face_recognition clustering
    Scoring                →  0–100 score, KEEP/REMOVE/REVIEW decision
S7  Vision LLM             →  (optional) Ollama on ambiguous band, adds memorability
    Re-score               →  incorporate vision signals, write Excel + CSV
```

Cache persists across runs. Reruns skip completed stages. Cache auto-invalidates when folder or metrics path changes.

## Vision LLM details

Vision fires only on photos scoring 35–65 (the ambiguous band). Two model-aware prompt strategies:

**llava** — gets a JSON prompt, returns structured data including memorability 1–5. Memorability scale: 1=mundane (parking lot, accidental shot), 3=decent moment, 5=exceptional (once-in-a-lifetime). Key instruction: "A blurry photo of a meaningful moment scores higher than a sharp photo of nothing."

**llama3.2-vision** — ignores JSON instructions, gets structured text prompt (QUALITY/CATEGORY/MEMORABILITY tags), parsed with regex. Used as fallback when llava fails.

## Performance

Tested on 2287 photos (Camera folder) with RTX 2000 Ada 8GB:

| Stage | Time |
|-------|------|
| S1 metrics (fresh) | ~17 min |
| S1 metrics (cached) | <1 sec |
| S2 duplicates | ~45 sec |
| S3 burst detection | ~5 min (optimized) |
| S4 CLIP (GPU) | ~2 min |
| S4 CLIP (CPU) | ~21 min |
| S7 vision (400 photos) | ~1h 44min |

## Dependencies

### Required
- `openpyxl` — Excel output
- `pillow` — image processing
- `imagehash` — perceptual hashing

### Recommended
- `opencv-python` — faster blur/BRISQUE computation
- `tqdm` — progress bars
- `numpy` — numeric operations

### Optional
- `clip` + `torch` — scene tagging (GPU recommended)
- `face_recognition` — face detection/clustering
- Ollama with `llava` / `llama3.2-vision` — vision LLM

## License

Personal use. 100% local processing, no cloud APIs.
