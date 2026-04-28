# Handoff Prompt — Photo Cleanup Engine v3.2

## Who I am
Pradeep — embedded firmware architect (EDK2/UEFI), also working on a personal photo library
cleanup tool as a side project. Running on Windows with NVIDIA RTX 2000 Ada (8GB VRAM),
Ollama installed locally with these models: llava:latest, llama3.2-vision:latest,
nomic-embed-text:latest, hermes3:latest.

---

## What we built: main_v3.py

A single-script photo cleanup pipeline. Key design decisions (do not change these):

- **One script only** — main_v3.py does everything. No helper scripts.
- **Single command** — `python main_v3.py --folder "D:\Photos"` computes all metrics inline,
  scores, and outputs Excel. No intermediate files or two-step workflows.
- **6-column slim Excel output** — Filename, Score/100, Decision, Category, Caption, Reason.
  Sorted by score ascending (worst first). Color-coded decisions.
- **0-100 aggregated score per photo** — relative to folder (blur/composite normalized to
  folder p75) + absolute (BRISQUE, duplicate penalties, memorability).
- **Decisions: KEEP / REMOVE / REVIEW** — not configurable labels.
- **Inline metrics** — blur, BRISQUE-approx (MSCN-based), composite, colorfulness, brightness,
  contrast, resolution — all computed during Stage 1 load. No external metrics file required.
- **Optional external metrics** — `--metrics-excel` can override inline values when available.
- **Vision LLM fires only on score band 35-65** — not all photos. Capped by --vision-limit.
- **Memorability scoring** — vision prompt asks for 1-5 memorability (mundane→exceptional).
  Worth 15 pts. "A blurry photo of a meaningful moment scores higher than a sharp photo of nothing."
- **SQLite cache** — crash-safe resume. Auto schema migration (ALTER TABLE for new columns).
  Cache key = folder + metrics path hash.
- **config.json support** — layered config system. CLI overrides config, config overrides defaults.
  Supports legacy nested keys (v2_features, processing, models).
- **llava is primary vision model, llama3.2-vision is fallback** — llama3.2-vision ignores
  JSON instructions and returns narrative text, so it gets a different prompt and regex parser.
  llava returns proper JSON.

---

## Known bugs fixed (do not reintroduce)

1. **Path matching for metrics Excel** — Excel has relative paths like `PHOTOS\filename.jpg`,
   file scan has absolute paths. Fixed with filename-only fallback match via `_match_metrics_row`.

2. **Decision engine all-REVIEW bug** — original code had `if not metrics_exists: continue`
   which skipped all logic for unmatched photos. Removed. Now all photos go through
   duplicate/burst/event/vision logic regardless.

3. **Vision JSON parsing** — llama3.2-vision returns narrative, not JSON. Fixed with
   model-aware prompts: llava gets JSON prompt, llama3.2-vision gets structured text prompt
   (QUALITY:/CATEGORY:/MEMORABILITY: tags) parsed with regex.

4. **CLIP wrong package** — `pip install clip` installs wrong package.
   Correct: `pip install git+https://github.com/openai/CLIP.git torch`

5. **Unicode logging on Windows** — log file uses UTF-8 encoding, console handler safe.

6. **Stage ordering** — initial scoring pass runs BEFORE vision so vision triggers
   (based on score band 35-65) actually fire.

7. **O(n²) burst detection** — pre-compute timestamps and hashes once, sort by time,
   break when gap > 5 min. Was hanging on 2287 photos (millions of EXIF re-reads).

8. **Cache schema mismatch crash** — DB.load() catches TypeError from column mismatches.
   DB.__init__() uses ALTER TABLE to add new columns. No --clear-cache needed for schema changes.

9. **Threshold shift cancelling vision gains** — thresholds only tighten to 70/30 when
   vision covered >50% of the review band. Otherwise stays at 65/35.

10. **Pillow deprecation** — `getdata()` replaced with `get_flattened_data()` where available,
    with fallback for Pillow < 11.

---

## Scoring formula

```
Baseline: 50 pts

Blur (25 pts, relative to folder p75):
  ratio >= 1.5  → +25
  ratio >= 1.0  → +15
  ratio >= 0.5  → +5
  ratio >= 0.2  → -10
  else          → -20

BRISQUE (15 pts, absolute — lower is better):
  < 15  → +15
  < 30  → +8
  < 50  → +2
  else  → -10

Composite score (15 pts, relative to folder p75):
  ratio >= 1.5 → +15
  ratio >= 1.0 → +8
  ratio >= 0.5 → +2
  else         → -5

Label (10 pts): GOOD=+10, AVERAGE=0, POOR=-10
Burst winner: +10
Only image in event: +10
Vision quality (10 pts): excellent=+20, good=+12, average=0, poor=-15
Memorability (15 pts): 1=-5, 2=0, 3=+5, 4=+10, 5=+15

Hard overrides (bypass scoring):
  Exact duplicate non-winner → score=10, REMOVE
  Burst non-winner           → score=20, REMOVE
```

Thresholds:
  Default:                  KEEP >= 65, REMOVE <= 35
  Vision coverage > 50%:    KEEP >= 70, REMOVE <= 30

---

## Current state

- Script runs end-to-end on 2287 Camera photos and 499 Naples photos
- 305 burst groups detected (Camera), 58 (Naples)
- 150 event groups (Camera), 1 (Naples)
- Vision working: llava ~15-18s/image, JSON parsing correct, memorability returned
- CLIP running on GPU (cuda) — RTX 2000 Ada, ~2 min for 2287 images
- BRISQUE approximation computed inline via MSCN coefficients (no opencv-contrib needed)
- Composite score derived from blur/BRISQUE/resolution/colorfulness/contrast
- DB auto-migrates schema on new columns
- config.json drives defaults, CLI overrides
- Results on Camera (2287 photos, 400 vision): KEEP=838, REMOVE=903, REVIEW=546

---

## Config file (config.json)

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

---

## Things Pradeep cares about (preferences)

- Crisp, aligned log output — stage_timer with wall+CPU time, no verbose per-image debug lines
- Minimal Excel columns — only what drives action
- No separate helper scripts — one script does everything
- Score should be interpretable as "confidence this is worth keeping"
- Privacy — 100% local, no cloud APIs
- Does not want explanations of things he already knows
- Pushes back when responses drift or repeat mistakes
- Single unified solution — no two-step workflows, no intermediate files

---

## The script (main_v3.py)

[PASTE CONTENTS OF main_v3.py HERE BEFORE SENDING]
