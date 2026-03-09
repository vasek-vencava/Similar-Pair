# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script project for **product name similarity matching** in Google Sheets. It pairs product names from a query sheet against a reference catalog using text similarity (cosine + Jaccard), size/unit normalization, and brand matching. Managed via [clasp](https://github.com/nicholaschiang/clasp) for local development.

## Deployment

```bash
# Push local changes to Google Apps Script
clasp push

# Pull remote changes
clasp pull

# Open the script in browser
clasp open
```

No build step, no tests, no linter. Code runs directly in the Apps Script V8 runtime.

## Architecture

### Files and Responsibilities

- **`Podobnost Párování názvu.js`** — Core pairing engine. Contains all scoring logic, index management, and the `onOpen()` menu. Entry points: `RunAllBatches()`, `RunNextBatch()`, `DebugRow(row)`.
- **`SizeSort.js`** — Standalone size-based sorting utility with inline HTML dialog. Entry: `SizeSortV1OpenSortDialog()`.
- **`Stahování dat.js`** — CSV import from external product feeds (DE/EN). Downloads, parses (Windows-1250), and writes to sheets. Entry: `AA_importAll()`, `AA_importDE()`, `AA_importEN()`.
- **`URLs.js`** — XML feed importer for product URLs. Entry: `urlsimport()`.

### Pairing Pipeline (main flow)

1. **Index build** (`buildOrLoadIndex`) — tokenizes reference names, builds inverted token/number indexes, computes IDF weights. Index is cached as gzip JSON on Google Drive (`PAIRING_INDEX_v8_SPECS.json.gz`).
2. **Candidate selection** — IDF-weighted token lookup + numeric key lookup, with fallback sampling.
3. **Coarse scoring** — Jaccard + size match bonus to narrow to `MAX_CANDIDATES` (250).
4. **Fine scoring** (`similarityScore`) — Cosine similarity on character trigrams (60%) + Jaccard on tokens (40%). Spec matching via `extractSpecs()` and `compareSpecs()`: matching specs add `SPEC_MATCH_BONUS` (+15%), mismatching specs multiply by `SPEC_MISMATCH_MULT` (×0.2). Score is NOT capped at 100% — can exceed it when specs confirm the match. Brand weight added separately (10%).
5. **Output** — Top K (3) matches written to columns O+ of `Nazvy_k_Parovani` sheet.

### Sheet Layout

| Sheet | Purpose |
|---|---|
| `Nase_nazvy` | Reference catalog. Col B=brand, C-R=name variants |
| `Nazvy_k_Parovani` | Query names (col B) + brand (col C). Output starts at col O |
| `skladDE` / `skladEN` | Imported product data from CSV feeds |
| `URLs` | Imported product URLs from XML feed |

### Key Configuration Constants

All config is at the top of `Podobnost Párování názvu.js`. Important ones:
- `BATCH_SIZE` (1000), `MAX_RUN_MS` (5.5 min) — controls auto-continuation to stay within Apps Script execution limits.
- `MAX_CANDIDATES` (250), `TOKEN_CANDIDATE_MULT` (80) — candidate pool sizing.
- `BRAND_WEIGHT` (0.10) — brand influence on final score.
- `SPEC_MATCH_BONUS` (0.15) — +15% per matching spec (volume, weight, flow, power, length).
- `SPEC_MISMATCH_MULT` (0.2) — ×0.2 per mismatching spec (heavy penalty).
- Specs are extracted by `extractSpecs()` which recognizes: volume (ml/l), weight (g/kg), length (mm/cm/m), flow (m3/h), power (W/kW).

## Conventions

- Language: Czech for UI strings, variable names, and comments. Function names mix Czech and English.
- All files share a single Apps Script project (no modules/imports). Functions must have globally unique names.
- `SizeSort.js` prefixes all identifiers with `SizeSortV1` to avoid collisions.
- `Stahování dat.js` prefixes with `AA_`.
- Helper functions like `safeStr`, `cleanCell`, `normalize` are in the main pairing file and used across scripts.
- The `safeCell` function in `Stahování dat.js` is separate from `cleanCell` in the pairing file — they are not interchangeable.
