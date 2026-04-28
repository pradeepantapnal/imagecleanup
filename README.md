# Intelligent Photo Cleanup Engine

A fully local Windows Python tool that analyzes photo folders using existing quality metrics to generate intelligent cleanup recommendations (KEEP/REVIEW/CANDIDATE_REMOVE). No cloud calls, no file deletion - recommendations only.

## Features

- **Existing Metrics Integration**: Uses your existing Excel quality metrics as the primary signal
- **Exact Duplicate Detection**: MD5-based identification of identical files
- **Near-Duplicate/Burst Detection**: pHash similarity with temporal proximity for burst photo groups
- **CLIP Scene Tagging**: Local CLIP embeddings for broad scene categorization (indoor/outdoor, people/landscape, etc.)
- **Event Grouping**: Conservative event detection based on timestamps and optional GPS
- **Deterministic Decision Engine**: Rule-based KEEP/REVIEW/CANDIDATE_REMOVE recommendations
- **Crash-Safe Resume**: SQLite persistence with automatic checkpointing and resume capability
- **Multiple Output Formats**: Excel (enriched), CSV (decisions), Markdown (report), and logs

## Architecture

The engine follows a 6-stage pipeline:

1. **Stage 1 - Load and Reconcile**: Load Excel metrics and reconcile with file system
2. **Stage 2 - Exact Duplicate Detection**: MD5 hashing and grouping
3. **Stage 3 - Burst Detection**: Near-duplicate detection with temporal proximity
4. **Stage 4 - CLIP Enrichment**: Scene tagging using local CLIP embeddings
5. **Stage 5 - Event Grouping**: Time-based event clustering
6. **Stage 6 - Decision Engine**: Deterministic KEEP/REVIEW/CANDIDATE_REMOVE policy

## Requirements

### Python Dependencies
```bash
pip install -r requirements.txt
```

Required packages:
- `openpyxl` - Excel file handling
- `pillow` - Image processing
- `imagehash` - Perceptual hashing
- `opencv-python` - Computer vision operations
- `tqdm` - Progress bars
- `clip-by-openai` - CLIP embeddings
- `torch` - PyTorch for CLIP
- `numpy` - Numerical operations

### Optional Dependencies
- **GPU**: CLIP can use GPU acceleration with CUDA, but works on CPU
- **Existing Metrics Excel**: Required for v1 operation

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Prepare your metrics Excel file with existing quality scores

## Usage

### Basic Usage
```bash
python main_v3.py --folder "D:\Photos"
```

### With Custom Config File
```bash
python main_v3.py --folder "D:\Photos" --config "my_config.json"
```

### With CLI Overrides
```bash
python main_v3.py --folder "D:\Photos" --enable-vision --vision-model llava
```

### V1 (Basic) Usage
```bash
python main.py --folder "D:\Photos"
```

### Command-Line Options

#### Required
| Option | Description |
|--------|-------------|
| `--folder` | Path to photo folder to scan |

#### Optional
| Option | Description |
|--------|-------------|
| `--config` | Path to JSON config file (default: config.json) |
| `--metrics-excel` | Override metrics Excel path from config |
| `--output-dir` | Override output directory from config |
| `--limit` | Override image limit from config |
| `--enable-vision` | Override: enable vision LLM enrichment |
| `--vision-model` | Override vision model from config |

## Configuration File

The engine uses a JSON configuration file (`config.json`) for most parameters:

```json
{
  "output_dir": "output",
  "limit": null,
  "metrics_excel": null,

  "enable_vision": true,
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

`main_v3.py` also accepts legacy nested keys (`v2_features`, `processing`, `models`) for backward compatibility.

## Input Requirements

### Metrics Excel Format
The Excel file must contain quality metrics for your photos. Expected columns include:
- `path` or `filename` - Image file path
- `composite_score` - Overall quality score
- `blur` - Blur detection score
- `sharpness` - Sharpness metric
- `resolution` - Image resolution

## Output Files

The engine generates four outputs in the specified directory:

### 1. `photo_report_enriched.xlsx`
Enriched Excel file with original metrics plus new analysis columns:
- Reconciliation status (file/metrics existence)
- Duplicate detection results
- Burst detection results
- CLIP scene tags
- Event grouping
- Final decisions with reasoning

### 2. `photo_decisions.csv`
Machine-readable CSV with cleanup decisions:
- File path and existence status
- Final decision (KEEP/REVIEW/CANDIDATE_REMOVE)
- Decision confidence and reasoning
- Key metrics for sorting/filtering

### 3. `photo_report.md`
Human-readable Markdown report with:
- Summary statistics
- Decision breakdown
- Recommendations
- Processing summary

### 4. `engine.log`
Detailed processing log with:
- Stage-by-stage progress
- Error messages
- Performance metrics

### 5. `photo_cache.db`
SQLite database for crash-safe resume:
- Cached computation results
- Pipeline state
- Configuration fingerprint

## Decision Engine Logic

The deterministic decision engine applies rules in this order:

1. **Missing Files** → REVIEW (file in Excel but not on disk)
2. **Missing Metrics** → REVIEW (file exists but no quality metrics)
3. **Exact Duplicate Losers** → CANDIDATE_REMOVE (lower-ranked duplicates)
4. **Burst Non-Winners** → CANDIDATE_REMOVE (lower quality burst photos)
5. **Burst Winners** → KEEP (best quality in burst group)
6. **Only Image in Event** → KEEP (sole representative of event)
7. **Low Quality Unique** → REVIEW (low composite score)
8. **Default** → KEEP (event representative)

## Crash-Safe Resume

The engine automatically resumes from interruptions:
- SQLite database caches all computation results
- Configuration fingerprint detects parameter changes
- Checkpoint after every N images and stage boundaries
- Skips already-completed work when config unchanged

## Configuration

Edit constants in `main.py` to customize behavior:

```python
SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif", ".heic"}
DUPLICATE_THRESH = 8      # Hamming distance for near-duplicates
BLUR_THRESH = 80.0         # Laplacian variance threshold
EVENT_GAP_HOURS = 6        # Time gap for event grouping
CHECKPOINT_INTERVAL = 10  # Checkpoint every N images
```

## Performance Notes

- **Exact duplicates**: Fast, MD5 hashing
- **Burst detection**: Medium, pHash + temporal analysis
- **CLIP embeddings**: Slower, benefits from GPU acceleration
- **Event grouping**: Fast, timestamp-based
- **Decision engine**: Very fast, rule-based

## Troubleshooting

### "Metrics Excel not found"
- Ensure the Excel file path is correct
- Verify the file contains the expected columns

### "CLIP not available"
- Install with: `pip install clip-by-openai torch`
- CLIP is optional; engine works without it (scene tagging will be limited)

### "Out of memory"
- Process folders in smaller batches using `--limit`
- CLIP embeddings can be memory-intensive

### "Resume not working"
- Check that configuration hasn't changed
- Verify the cache database isn't corrupted
- Delete `photo_cache.db` to start fresh

## v2 Roadmap

Features planned for future versions:
- Face detection and clustering
- Selective vision LLM review
- Semantic caption search
- Smarter event naming
- Advanced deduplication algorithms

## License

This project is provided as-is for personal photo organization.

## Design Philosophy

**v1 focuses on stability and determinism:**
- Existing metrics first, AI enrichment second
- Rule-based decisions over probabilistic approaches
- Conservative event grouping
- Crash-safe operation
- Fully local processing

This ensures reliable, repeatable results for photo cleanup tasks.
