# MLBB Counter Draft — v2.0

A counter-pick recommendation engine for **Mobile Legends: Bang Bang**. It migrates hero matchup data from an Excel workbook into **Supabase (PostgreSQL)**, precomputes a full counter-score matrix using a multi-factor scoring formula, and serves the results through a real-time interactive draft board.

**Live demo** — open `index.html` in any browser (see [Frontend](#frontend-usage) below).

---

## Table of Contents

- [Features](#features)
- [Tech Stack](#tech-stack)
- [Architecture & Data Flow](#architecture--data-flow)
- [Scoring Formula](#scoring-formula)
- [Database Schema](#database-schema)
- [Frontend Usage](#frontend-usage)
- [Project Structure](#project-structure)
- [Setup](#setup)
- [Running the Migration](#running-the-migration)
- [Running Score Computation](#running-score-computation)
- [Frontend Deployment](#frontend-deployment)
- [GitHub Actions](#github-actions)
- [Manual Overrides](#manual-overrides)
- [Hard Counter Rules](#hard-counter-rules)
- [Notes & Gotchas](#notes--gotchas)

---

## Features

- **133 MLBB heroes** with full stats, roles, damage types, power spikes, style tags, and lane assignments
- **Multi-factor scoring** — role matchup, burst potential, difficulty gap, damage-type advantage, power-spike timing, style interactions, hard-counter rules, range-type advantage, and anti-heal advantage
- **Precomputed matrix** — all 133 × 132 = ~17,500 attacker→defender pairs computed offline and stored in Postgres with component breakdown
- **Interactive draft board** — select enemy picks in hexagonal slots and see the top 10 counter picks update in real time
- **Sum / Average mode** — toggle between total score and per-pick average aggregate
- **Per-matchup explanation** — each recommendation shows matched hard-counter rules with human-readable descriptions (e.g. "Strong against Dash/Blink heroes — Bouncing Ball cancels dashes...")
- **Final Line-up modal** — one-click team summary with one hero per lane, role-prioritized selection (Jungle→Assassin, Roam→Tank/Support, etc.), no duplicate heroes, and per-enemy reason breakdowns
- **Lane labels** — picked heroes show their lane(s) in bold accent color on the slot
- **Quick-take summary** — top overall counter displayed with a natural-language explanation of why it fits
- **Manual overrides** — force specific scores for any attacker→defender pair, bypassing the formula
- **Hard counter rules** — 102 special-case rules (Tag, Hero, Role, DamageType, Resource conditions) with notes; case-insensitive matching
- **CI/CD** — two GitHub Actions workflows auto-run migration and score recomputation on push

---

## Tech Stack

| Layer | Technology |
|---|---|
| **Frontend** | Vanilla HTML + CSS + JavaScript (no framework) |
| **Backend** | [Supabase](https://supabase.com/) (BaaS — Postgres DB, REST API, JS/Python SDK) |
| **Database** | PostgreSQL (via Supabase) |
| **Data Source** | Microsoft Excel workbook (`.xlsx`) — single source of truth |
| **Migration** | Python 3.12 + `openpyxl` + `supabase` SDK |
| **Computation** | Python 3.12 + `supabase` SDK |
| **CI/CD** | GitHub Actions (2 workflows) |

---

## Architecture & Data Flow

```
┌─────────────────────────────────────────────────────────────────────┐
│                    Excel Workbook (source of truth)                 │
│  mobile_legends_heroes_updated.xlsx                                 │
│    ┌──────────┐  ┌────────────┐  ┌──────────┐  ┌────────────────┐  │
│    │  Heroes  │  │ Data-Input │  │Document- │  │CounterCalc/    │  │
│    │  (stats) │  │ (weights,  │  │ation -   │  │Lookup/         │  │
│    │          │  │  matrices, │  │Comput-   │  │LookupCalc      │  │
│    │          │  │  rules,    │  │ations    │  │(reference      │  │
│    │          │  │  overrides)│  │(formula) │  │ only)          │  │
│    └────┬─────┘  └─────┬──────┘  └──────────┘  └────────────────┘  │
└─────────┼──────────────┼────────────────────────────────────────────┘
          │              │
          ▼              ▼
┌─────────────────────────────────────────────────────────────────────┐
│                   migrate_to_supabase.py                            │
│  Reads Heroes → heroes table                                        │
│  Reads Data-Input → global_weights, role_matrix, style_matrix,      │
│                     hard_counter_rules, manual_overrides            │
│  Deduplicates hard_counter_rules by (attacker, type, value) key     │
└─────────────────────────────────┬───────────────────────────────────┘
                                  │ upsert
                                  ▼
┌─────────────────────────────────────────────────────────────────────┐
│                     Supabase (PostgreSQL)                           │
│  ┌────────┐ ┌───────────────┐ ┌────────────┐ ┌───────────────┐     │
│  │ heroes │ │global_weights │ │role_matrix │ │style_matrix   │     │
│  ├────────┤ ├───────────────┤ ├────────────┤ ├───────────────┤     │
│  │133 rows│ │10 coefficients│ │ 6×6 grid   │ │22×22 grid     │     │
│  └────────┘ └───────────────┘ └────────────┘ └───────────────┘     │
│  ┌───────────────────┐ ┌────────────────┐                          │
│  │hard_counter_rules │ │manual_overrides│                          │
│  ├───────────────────┤ ├────────────────┤                          │
│  │~102 rules         │ │~1000 overrides │                          │
│  └───────────────────┘ └────────────────┘                          │
└──────────────────────────────┬─────────────────────────────────────┘
                               │ read
                               ▼
┌─────────────────────────────────────────────────────────────────────┐
│                   compute_counters.py                               │
│  Reads all reference tables, computes formula for every             │
│  (attacker, defender) pair, stores matched rule notes in           │
│  matched_rules JSONB column, upserts into counter_scores.          │
│  Supports --dry-run for preview without writes.                    │
└─────────────────────────────────┬───────────────────────────────────┘
                                  │ upsert
                                  ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      counter_scores table                           │
│  133 × 132 = ~17,556 rows, each with component breakdown            │
│  + matched_rules JSONB (which hard-counter rules fired)            │
└─────────────────────────────────┬───────────────────────────────────┘
                                  │ SELECT (read-only, anon key)
                                  ▼
┌─────────────────────────────────────────────────────────────────────┐
│                     index.html (browser)                            │
│  - Fetches all heroes + scores on load                              │
│  - Interactive draft board with 5 hexagonal enemy slots             │
│  - Real-time top-10 counter recommendations                         │
│  - Per-matchup explanation (matched rule notes)                     │
│  - Sum / Average mode toggle                                        │
│  - Final Line-up modal (role-prioritized, duplicate-free)          │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Scoring Formula

For every (attacker, defender) hero pair, the total counter score is:

```
TOTAL = Role Advantage
      + Stat (Burst)
      + Difficulty Gap
      + Damage Type Advantage
      + Power Spike Timing
      + Style Matchup
      + Hard Counter Bonus
      + Range Type Advantage
      + Anti-Heal Advantage
```

If a `manual_overrides` row exists for that exact pair, **TOTAL is replaced entirely** by the override score.

### Component Details

#### 1. Role Advantage
```
MAX(role_matrix[atk.role → def.role],
    role_matrix[atk.role2 → def.role])   (skip if role2 is null)
× role_mult
```
The 6×6 role grid (Tank, Fighter, Assassin, Mage, Marksman, Support) defines base advantage points. If the attacker has a secondary role (`role2`), the better of the two matchups is used.

#### 2. Stat (Burst Potential)
```
((atk.offense + atk.ability_effects) / 2 - def.durability) × burst_mult
```
Measures whether the attacker's offensive stats overcome the defender's durability.

#### 3. Difficulty Gap
```
gap = atk.difficulty - def.difficulty
gap × diff_mult   if gap > 20
0                 otherwise
```
A significant difficulty gap (attacker harder to play than defender) adds a bonus.

#### 4. Damage Type Advantage
```
dmgtype_mixed   if attacker damage_type is "Mixed"
dmgtype_same    elif attacker == defender damage_type
dmgtype_diff    else
```
Mixed-damage attackers are hardest to itemize against, then same-type, then different-type.

#### 5. Power Spike Timing
```
(atk.spike_order - def.spike_order) × spike_mult
```
Heroes that spike earlier get an advantage against late-game heroes.

#### 6. Style Matchup
```
MAX(style_matrix[atk.style1/2 → def.style1/2]) × style_mult
```
Evaluates up to 4 combinations and takes the maximum.

#### 7. Hard Counter Bonus
```
SUM(bonus_to_attacker - penalty_to_defender) for every matching hard_counter_rule
```
Rules match when `rule.attacker == atk.name` (case-insensitive) and the defender satisfies the condition:
- **Tag** — defender's `style1` or `style2` matches `condition_value`
- **Hero** — defender's `name` matches `condition_value`
- **Role** — defender's `role` or `role2` matches `condition_value`
- **DamageType** — defender's `damage_type` matches `condition_value`
- **Resource** — defender's `resource` matches `condition_value`

Matched rules are stored in the `counter_scores.matched_rules` JSONB column and used by the frontend to display explanations.

#### 8. Range Type Advantage
```
rangetype_mult   if attacker has a mobility tag AND defender.range_type == "Ranged"
0                otherwise
```
Mobility-tagged attackers get a bonus against ranged defenders.

#### 9. Anti-Heal Advantage
```
antiheal_mult    if attacker.has_antiheal AND defender has Sustain/Healing tag
0                otherwise
```
Heroes with innate anti-heal counter sustain/healing-tagged opponents.

---

## Database Schema

Seven tables defined in `schema.sql`:

| Table | Purpose | Key |
|---|---|---|
| `heroes` | Per-hero base stats, roles, tags, lanes, damage type, power spike | `id` (PK), `name` (UNIQUE) |
| `global_weights` | Tunable formula coefficients (10 rows) | `coefficient` (PK) |
| `role_matrix` | 6×6 role matchup grid | `(attacker_role, defender_role)` (PK) |
| `style_matrix` | 22×22 style/tag interaction grid | `(attacker_tag, defender_tag)` (PK) |
| `hard_counter_rules` | Special-case counter rule triggers with notes | `id` (serial PK), unique index on `(attacker, condition_type, condition_value)` |
| `manual_overrides` | Forced exact scores for specific pairs | `(attacker, defender)` (PK) |
| `counter_scores` | Precomputed per-pair scores with component breakdown and matched rule explanations | `(attacker, defender)` (PK) |

`counter_scores` includes:
- `matched_rules jsonb` — array of matched hard-counter rule objects (type, value, bonus, penalty, note)
- Indexes on `defender` and `score DESC` for fast lookup

---

## Frontend Usage

`index.html` is a fully self-contained single-page application:

1. **Connects** to Supabase via the JS client (CDN-loaded `@supabase/supabase-js@2`)
2. **Loads** all 133 hero names and all ~17,500 counter scores (with `matched_rules`) into an in-memory `Map`
3. **Renders** 5 hexagonal enemy-pick slots in a row with lane suggestions
4. **User taps** an empty slot → a fuzzy-search popover appears with hero names filtered as you type
5. **Selecting** a hero locks it into the slot (shows name, role, and lane in bold)
6. **Results panel** immediately shows the top 10 counter picks:
   - Quick-take summary of the #1 overall counter
   - Top 5 overall + top 5 per lane (tabbed by EXP/Mid/Jungle/Gold/Roam)
   - Each card shows hero name, role(s), lane, aggregate score
   - Per-enemy breakdown with horizontal bar + matched rule explanations
7. **Mode toggle** switches between **Sum** (add all matchup scores) and **Average** (divide by number of enemy picks)
8. **Final Line-up button** opens a modal with one hero per lane:
   - Role-prioritized selection (Jungle→Assassin, Roam→Tank/Support, etc.)
   - No duplicate heroes
   - Per-enemy explanation bullets with score contributions
   - Role coverage indicator + overall summary
   - Close and Reselect buttons

The frontend is **read-only** — it uses the public anon key with no write access.

---

## Project Structure

```
.
├── .github/workflows/
│   ├── migrate.yml                 # Auto-run migration on workbook/schema changes
│   └── recompute_counters.yml      # Auto-recompute scores on every push to main
├── archive/                        # Historical workbook backups (not used)
├── index.html                      # Frontend draft board (single-page app)
├── schema.sql                      # PostgreSQL table definitions
├── migrate_to_supabase.py          # Workbook → Supabase migration script
├── compute_counters.py             # Score matrix computation script
├── mobile_legends_heroes_updated.xlsx  # Source workbook (single source of truth)
├── requirements.txt                # Python dependencies
├── .gitignore                      # Ignores .env and __pycache__
└── README.md                       # This file
```

---

## Setup

### Prerequisites
```bash
pip install -r requirements.txt
```

`requirements.txt`:
```
openpyxl        # Read Excel workbooks
supabase        # Supabase Python client (supabase-py)
```

### Supabase Project

1. Create a project at [supabase.com](https://supabase.com)
2. Run `schema.sql` in the SQL Editor to create all 7 tables
3. Get your credentials from **Project Settings → API**:
   - **Project URL** → `SUPABASE_URL`
   - **`service_role` key** → `SUPABASE_SERVICE_KEY` (for migration/computation scripts)
   - **`anon` public key** → `SUPABASE_ANON_KEY` (for the frontend)

### Environment Variables

| Variable | Used By | Where to Find It |
|---|---|---|
| `SUPABASE_URL` | Python scripts | Supabase Dashboard → Project Settings → API → Project URL |
| `SUPABASE_SERVICE_KEY` | Python scripts | Supabase Dashboard → Project Settings → API → `service_role` secret |

Never commit secrets. Export them locally or use a `.env` file. In CI, set as GitHub Secrets.

---

## Running the Migration

Migrates the workbook into Supabase reference tables (`heroes`, `global_weights`, `role_matrix`, `style_matrix`, `hard_counter_rules`, `manual_overrides`).

```bash
export SUPABASE_URL="https://xxxx.supabase.co"
export SUPABASE_SERVICE_KEY="eyJ..."
python migrate_to_supabase.py mobile_legends_heroes_updated.xlsx
```

The script:
1. Extracts **heroes** from the `Heroes` sheet
2. Extracts **weights, matrices, rules, overrides** from `Data-Input` sheet at documented row ranges
3. **Upserts** into `heroes`, `global_weights`, `role_matrix`, `style_matrix`, `manual_overrides`
4. **Upserts** into `hard_counter_rules` — uses a unique index on `(attacker, condition_type, condition_value)` so re-running safely updates existing rules without creating duplicates; duplicate rows in the workbook are merged (bonus/penalty summed)
5. Prints row counts and validates hero count is exactly 133

---

## Running Score Computation

Computes all ~17,500 attacker→defender scores and upserts into `counter_scores`:

```bash
export SUPABASE_URL="https://xxxx.supabase.co"
export SUPABASE_SERVICE_KEY="eyJ..."
python compute_counters.py
```

### Dry-Run Mode

Preview computed rows without writing to the database:

```bash
python compute_counters.py --dry-run
```

Output:
```
--- DRY RUN: first 10 of 17556 computed rows ---
{'attacker': 'Alice', 'defender': 'Alpha', 'score': 12.345, 'role_advantage': 2.0, ...}
...
Dry run complete. No rows written to counter_scores.
```

---

## Frontend Deployment

`index.html` is a static file — no build step required. Host it via:

- **GitHub Pages** (recommended — push to `gh-pages` branch or use `docs/` folder)
- **Any static file server**: `npx serve`, Python `http.server`, Nginx, etc.
- **Direct file open** may not work due to CORS restrictions on Supabase

**Important**: Update the `SUPABASE_URL` and `SUPABASE_ANON_KEY` constants at the top of the script block in `index.html` before deploying:

```js
const SUPABASE_URL = "https://your-project.supabase.co";
const SUPABASE_ANON_KEY = "eyJ...";
```

---

## GitHub Actions

Two CI workflows are included:

### 1. `migrate.yml` — Data Migration

- **Triggers on**: Push to `main` touching `mobile_legends_heroes_updated.xlsx`, `migrate_to_supabase.py`, or `schema.sql`; also manual (`workflow_dispatch`)
- **Runs**: Checkout → Python 3.12 → `pip install` → pre-flight diagnostic (file & secrets check) → `python migrate_to_supabase.py`

### 2. `recompute_counters.yml` — Score Recalculation

- **Triggers on**: Every push to `main`; also manual
- **Runs**: Checkout → Python 3.12 → `pip install` → pre-flight diagnostic → `python compute_counters.py`

### Required GitHub Secrets

| Secret | Value |
|---|---|
| `SUPABASE_URL` | Your Supabase project URL |
| `SUPABASE_SERVICE_KEY` | Your `service_role` key |

Set them at **Settings → Secrets and variables → Actions** in your repo.

---

## Manual Overrides

The `manual_overrides` table lets you force an exact score for any (attacker, defender) pair, completely bypassing the formula. This is useful for matchups the formula gets wrong.

- **One-directional**: Override for Alice→Balthazar does not apply to Balthazar→Alice
- **Source**: Edited in the `Data-Input` sheet of the workbook (Block 6)
- **When present**: `compute_counters.py` uses the override score directly instead of computing the 9 components

---

## Hard Counter Rules

The `hard_counter_rules` table defines special-case bonuses. Each rule has:

| Field | Description |
|---|---|
| `attacker` | The hero who gets the bonus |
| `condition_type` | `Tag`, `Hero`, `Role`, `DamageType`, or `Resource` |
| `condition_value` | The value to match against the defender |
| `bonus_to_attacker` | Points added to the attacker's score |
| `penalty_to_defender` | Points subtracted from the defender's score (effectively added to the attacker) |
| `note` | Human-readable explanation used by the frontend |

Example: A rule like `("Karrie", "Tag", "Sustain", 10, -10, "True damage passive shreds lifesteal/regen-heavy sustain heroes.")` means Karrie gets +10 against Sustain-tagged heroes, and those defenders get -10 (net +20).

The workbook has dropdown validations on columns A-C (hero names, condition types, condition values) for easy editing. The extraction script deduplicates rows with the same `(attacker, condition_type, condition_value)` by summing bonus/penalty values.

---

## Notes & Gotchas

1. **`hard_counter_rules` upsert**: The table has a unique index on `(attacker, condition_type, condition_value)`. The migration script upserts, so re-running safely updates existing rules without creating duplicates. Duplicate rows in the workbook with the same key are merged automatically.

2. **`Data-Input` is the only hand-edit sheet**: The `Heroes` sheet is raw data; `CounterCalculator`, `Lookup`, and `LookupCalc` are formula-driven convenience views inside Excel — these are not migrated.

3. **Counter scores are precomputed**: The frontend does not compute scores on the fly. After changing hero data or weights, you must re-run `compute_counters.py` to update `counter_scores`.

4. **Anon key is public**: The Supabase anon key in `index.html` is meant to be visible to clients. It should be restricted via Row-Level Security (RLS) policies. For this read-only use case, the `heroes` and `counter_scores` tables can be set to `SELECT`-only for the anon role.

5. **133 heroes**: The sanity check in `migrate_to_supabase.py` expects exactly 133 hero rows. If the workbook is updated with new heroes, update the `EXPECTED_HERO_COUNT` constant.

6. **Case-insensitive matching**: The `hard_counter_bonus()` function in `compute_counters.py` uses case-insensitive comparisons (`.lower()`) for all rule conditions, so workbook data with inconsistent casing still matches correctly.

7. **No testing framework**: The project has no automated tests. Use `--dry-run` on `compute_counters.py` for manual verification of scores.

---

## License

MIT — see license file if included. This project is not affiliated with Moonton or Mobile Legends: Bang Bang.
