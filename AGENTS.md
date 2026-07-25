# MLBB Drafter v2.0

## Architecture

Single-page application (Vanilla JS) backed by Supabase. Two main modes:

- **Counter Draft** — enter 1–5 enemy picks, get ranked counter-picks (top 5 overall + top 5 per lane)
- **Team Fight** — build two 5-hero lineups, compare aggregate counter scores

All counter scores are **pre-computed** in `counter_scores` table by `compute_counters.py`. The front-end fetches and caches the entire matrix (~17.5k rows for 133 heroes) on load.

## Key files

| File | Purpose |
|------|---------|
| `index.html` | Single-page app — all HTML, CSS, and JS |
| `compute_counters.py` | Python script to recompute the full counter score matrix and upsert to Supabase |
| `AGENTS.md` | This file |

## Data flow

```
Supabase heroes table  ─┐
Supabase counter_scores ─┤──> index.html (fetches on load, caches in Maps)
Supabase synergy_scores ─┤
Supabase role_matrix   ─┘

compute_counters.py ──> counter_scores (upsert)
    reads: heroes, global_weights, role_matrix, style_matrix,
           hard_counter_rules, manual_overrides
    writes: counter_scores
```

## Build / run

No build step. Open `index.html` in a browser (or serve with any static file server).

```powershell
# local dev
npx serve .
```

## Python score computation

```powershell
$env:SUPABASE_URL = "https://xxxx.supabase.co"
$env:SUPABASE_SERVICE_KEY = "sb_secret_..."
python compute_counters.py

# dry-run (no writes):
python compute_counters.py --dry-run
```

The script loads all reference tables from Supabase, computes scores for every (attacker, defender) pair, and upserts into `counter_scores`. See the docstring in `compute_counters.py` for the full formula.

**2026-07-25:** Difficulty Gap removed from formula. Assumption is all players are equally skilled — difficulty no longer contributes to counter scores. All 17,556 pairs recomputed. Re-run `compute_counters.py` after any weight/rule/data change.

## Supabase schema notes

- `counter_scores` has a composite PK on `(attacker, defender)`
- `heroes` table must include: `name, role, role2, lane1, lane2, damage_type, style1, style2, has_antiheal, spike_order, has_true_damage, offense, durability, ability_effects, difficulty, resource, teamfight_contribution`
- `global_weights` uses column `coefficient` for key names and `value` for the float
- `hard_counter_rules` links an attacker hero to a condition (Tag/Hero/Role/DamageType/Resource) on the defender

## JS state invariants

- `slots[0..4]` — enemy picks (null if empty)
- `banSlots[0..4]` — banned heroes (null if empty)
- `teamA[0..4]`, `teamB[0..4]` — fight mode rosters
- `scoreMap`: key `"attacker|defender"` ➝ `{ score, matched_rules[] }`
- `synergyMap`: key `"hero_a|hero_b"` (bidirectional) ➝ number
- `roleMatrix`: key `"atk_role|def_role"` ➝ number

## Conventions

- Front-end: hero names compared case-sensitively (as stored in Supabase). Back-end (`compute_counters.py`): string comparisons use `_norm()` (`.strip().lower()`) for robust rule matching.
- `escapeHtml()` must be used for any user-visible interpolated string that could contain hero names or user input
- Lane detection: `computeLanes()` reads `lane1`/`lane2` from all heroes, prioritizes exp/mid/jungle/gold/roam by substring match
- Mode toggle: "sum" (default) sums scores across enemies; "avg" divides by enemy count

## compute_counters.py design notes

- **Performance:** Hard counter rules are pre-grouped by attacker name (`rules_by_attacker` dict) at load time. This avoids scanning all rules for every (attacker, defender) pair — 17,556 iterations × ~100 rules → O(n² × m) without grouping.
- **Validation fails fast:** Missing `global_weights` coefficients trigger an immediate `sys.exit` with the missing keys listed. Duplicate hero names are detected by building a `name → [ids]` map and reported with the conflicting IDs.
- **`_norm()` helper:** All string comparisons in rule matching use `_norm(s) = (s or "").strip().lower()` to survive whitespace/case inconsistencies in Supabase data entry.
- **`DEFAULT_SPIKE_ORDER = 3`:** Null `spike_order` falls back to the "Mid" band to avoid a `TypeError` crash on arithmetic.
- **`manual_overrides` lookup key:** Uses `(atk["name"], defn["name"])` — names, not IDs. Works because the duplicate-name validation ensures names are unique.

## Excel workbook conventions

The single source of truth is `mobile_legends_heroes_updated.xlsx`:

| Sheet | Content |
|-------|---------|
| Heroes | Hero stats (name, role, lane, damage_type, style, stats, etc.) |
| Data-Input | Global weights (rows 7-16), role matrix (18-24), hard counter rules (29-130), style matrix (164-186), manual overrides (191-1189) |
| Synergy | Hero synergy pairs (optional) |

**Hard counter rule conventions:**
- Default to symmetric values (`bonus_to_attacker = -penalty_to_defender`). Asymmetric values are allowed when a specific matchup requires the attacker to benefit more (or less) than the defender is penalized, but this should be the exception.
- Each `(attacker, condition_type, condition_value)` triple must appear at most once. Duplicates cause silent double-stacking in `hard_counter_bonus()`.
- The `_norm()` helper in `compute_counters.py` normalizes all string comparisons to lowercase. Entries with inconsistent casing (e.g. `"Angela"` vs `"angela"`) will be grouped together and stack. Run a case-insensitive dedup check after editing.

**Workflow for edits:**
1. Edit the Excel
2. Run `migrate_to_supabase.py` to push to Supabase
3. Run `compute_counters.py` to recompute all 17,556 pairs

## Troubleshooting

- **"row count mismatch" after compute** — the hero count may have changed. Run `compute_counters.py` again.
- **score rows not loading** — check Supabase project's "Max Rows" setting (default 1000). The front-end paginates with `.range()` but if it's set below 1000 the fetch may need smaller pages.
- **computation crashes with `KeyError` on weight lookup** — a required coefficient is missing from `global_weights`. Add the missing key and re-run.
- **rule count differs between Excel and DB after migration** — the migrate script upserts by key but doesn't delete stale rows. If you remove rules from the Excel, delete the corresponding rows in Supabase manually or truncate the table before re-migrating.
