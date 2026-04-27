# Handoff Prompt — Photo Cleanup Engine

## Who I am
Pradeep — working on a personal photo library
cleanup tool as a side project. Running on Windows with NVIDIA RTX 2000 Ada (8GB VRAM),
Ollama installed locally with these models: llava:latest, llama3.2-vision:latest,
nomic-embed-text:latest, hermes3:latest.

---

## What we built: main_v3.py

A single-script photo cleanup pipeline. Key design decisions (do not change these):

- **One script only** — main_v3.py does everything. No helper scripts.
- **6-column slim Excel output** — Filename, Score/100, Decision, Category, Caption, Reason.
  Sorted by score ascending (worst first). Color-coded decisions.
- **0-100 aggregated score per photo** — relative to folder (blur/composite normalized to
  folder p75) + absolute (BRISQUE, duplicate penalties).
- **Decisions: KEEP / REMOVE / REVIEW** — not configurable labels.
- **Vision LLM fires only on score band 35-65** — not all photos. Capped by --vision-limit.
- **SQLite cache** — crash-safe resume. Cache key = folder+metrics path hash.
- **llava is primary vision model, llama3.2-vision is fallback** — llama3.2-vision ignores
  JSON instructions and returns narrative text, so it gets a different prompt and regex parser.
  llava returns proper JSON.

---

## Known bugs fixed (do not reintroduce)

1. **Path matching for metrics Excel** — Excel has relative paths like `PHOTOS\filename.jpg`,
   file scan has absolute paths. Fixed with filename-only fallback match.

2. **Decision engine all-REVIEW bug** — original code had `if not metrics_exists: continue`
   which skipped all logic for unmatched photos. Removed. Now all photos go through
   duplicate/burst/event/vision logic regardless.

3. **Vision JSON parsing** — llama3.2-vision returns narrative, not JSON. Fixed with
   model-aware prompts: llava gets JSON prompt, llama3.2-vision gets structured text prompt
   (QUALITY:/CATEGORY: tags) parsed with regex.

4. **CLIP wrong package** — `pip install clip` installs wrong package.
   Correct: `pip install git+https://github.com/openai/CLIP.git torch`

5. **Unicode logging on Windows** — log file uses UTF-8 encoding, console handler safe.

6. **Stage ordering** — initial scoring pass runs BEFORE vision so vision triggers
   (based on decision_confidence) actually fire.

7. **Metrics not matching** — sample_metrics.xlsx had empty path/filename columns (dummy
   data from create_sample_metrics.py). Fixed by adding --generate-metrics flag to
   main_v3.py that generates real metrics from the folder itself.

---

## Scoring formula (do not change without discussion)

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
Vision quality (10 pts × weight):
  weight=1.0 with metrics, weight=2.5 without metrics
  excellent=+20, good=+12, average=0, poor=-15

Hard overrides (bypass scoring):
  Exact duplicate non-winner → score=10, REMOVE
  Burst non-winner           → score=20, REMOVE
```

Thresholds:
  With metrics:    KEEP >= 70, REMOVE <= 30
  Without metrics: KEEP >= 52, REMOVE <= 35

---

## Current state

- Script runs end-to-end on Windows and produces both Excel + CSV outputs.
- Latest validated run (2026-04-27) on 102 images completed in 252s.
- Duplicate/burst/event stages are working (2 duplicate groups, 1 burst group, 69 event groups in that run).
- CLIP is running on CUDA in current setup (log shows `S4 CLIP : ... (cuda)`).
- Vision stage is active and expensive (~13.1s/image; 15 images took ~196s with `--vision-limit 15`).
- Output distribution in latest run: KEEP=32, REMOVE=24, REVIEW=46.
- Pillow deprecation warning around `Image.getdata()` has now been addressed in `main_v3.py` using
  `get_flattened_data()` with a backward-compatible fallback.

---

## Immediate next steps

1. Re-run on a small folder (20-50 photos) and confirm there are no Pillow deprecation warnings:
   ```
   python main_v3.py --folder "C:\...\input" --enable-vision --vision-limit 5
   ```

2. Run `--generate-metrics` to produce real blur/composite scores for Naples:
   ```
   python main_v3.py --folder "C:\...\Naples" --generate-metrics
   ```

3. Run full pipeline with those metrics + vision overnight:
   ```
   python main_v3.py --folder "C:\...\Naples" --metrics-excel "output\Naples_metrics.xlsx" --enable-vision --vision-limit 200
   ```

4. If runtime is high, tune vision workload first (`--vision-limit`) before touching score thresholds.

5. Install/verify GPU-accelerated torch only if CLIP falls back to CPU again:
   ```
   pip install torch torchvision --index-url https://download.pytorch.org/whl/cu121
   ```

6. Interpret Excel results: REMOVE first (safe deletes), then sort REVIEW by score
   ascending and work upward.

---

## Things Pradeep cares about (preferences)

- Crisp, aligned log output — no verbose per-image debug lines
- Minimal Excel columns — only what drives action
- No separate helper scripts — one script does everything
- Score should be interpretable as "confidence this is worth keeping"
- Privacy — 100% local, no cloud APIs
- Does not want explanations of things he already knows
- Pushes back when responses drift or repeat mistakes

---

## The script (main_v3.py)

[PASTE CONTENTS OF main_v3.py HERE BEFORE SENDING]
