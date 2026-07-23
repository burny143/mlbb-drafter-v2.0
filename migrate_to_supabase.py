#!/usr/bin/env python3
"""
migrate_to_supabase.py

One-shot migration of the MLBB Hero Counter workbook into Supabase.

Reads:
    Heroes                          -> heroes
    Data-Input  rows 7-16   (Block 1: Global Weights)          -> global_weights
    Data-Input  rows 18-24  (Block 2: Role Matchup Matrix)     -> role_matrix
    Data-Input  rows 29-128 (Block 3: Hard Counter Rules)      -> hard_counter_rules
    Data-Input  rows 163-185(Block 5: Style/Tag Interaction)   -> style_matrix
    Data-Input  rows 190-1189(Block 6: Manual Overrides)       -> manual_overrides

Does NOT touch `counter_scores` — that table is populated later by a separate
batch job that actually computes the formula from Documentation - Computations.

Usage:
    export SUPABASE_URL="https://xxxx.supabase.co"
    export SUPABASE_SERVICE_KEY="eyJ..."
    python migrate_to_supabase.py path/to/mobile_legends_heroes_updated.xlsx
"""

import os
import sys
from pathlib import Path
from typing import Any, Optional

import openpyxl
from supabase import create_client, Client

# Load .env file if present
_dotenv = Path(__file__).parent / ".env"
if _dotenv.exists():
    for _line in _dotenv.read_text().splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip())


# ----------------------------------------------------------------------------
# Config: row ranges from the workbook (1-indexed, inclusive, as documented
# in Data-Input's section headers / the task spec).
# ----------------------------------------------------------------------------
GLOBAL_WEIGHTS_ROWS = (7, 16)          # Block 1
ROLE_MATRIX_HEADER_ROW = 18            # Block 2 header (defender roles across)
ROLE_MATRIX_DATA_ROWS = (19, 24)       # Block 2 data rows (attacker roles down)
HARD_COUNTER_ROWS = (29, 128)          # Block 3
STYLE_MATRIX_HEADER_ROW = 163          # Block 5 header (defender tags across)
STYLE_MATRIX_DATA_ROWS = (164, 185)    # Block 5 data rows (attacker tags down)
MANUAL_OVERRIDES_ROWS = (190, 1189)    # Block 6

EXPECTED_HERO_COUNT = 133


def get_supabase_client() -> Client:
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_KEY")
    if not url or not key:
        sys.exit(
            "ERROR: SUPABASE_URL and SUPABASE_SERVICE_KEY must be set as "
            "environment variables. Never hardcode credentials in this script."
        )
    return create_client(url, key)


def to_bool_yes_no(value: Optional[str]) -> bool:
    """Heroes.has_antiheal is stored in the sheet as the string 'Yes'/'No'."""
    if value is None:
        return False
    return str(value).strip().lower() == "yes"


def clean(value: Any) -> Any:
    """Trim strings, pass through everything else (including None/numbers)."""
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value


def row_values(ws, row: int, min_col: int, max_col: int) -> list:
    return [ws.cell(row=row, column=c).value for c in range(min_col, max_col + 1)]


# ----------------------------------------------------------------------------
# Extraction functions — each returns a list of dicts ready to upsert.
# ----------------------------------------------------------------------------
def extract_heroes(wb) -> list[dict]:
    ws = wb["Heroes"]
    headers = [c.value for c in ws[1]]
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:  # blank trailing row
            continue
        rec = dict(zip(headers, row))
        rec["id"] = int(rec["id"])
        for col in ("offense", "ability_effects", "durability", "difficulty", "spike_order"):
            rec[col] = int(rec[col])
        rec["has_antiheal"] = to_bool_yes_no(rec.get("has_antiheal"))
        for col in ("role2", "style2", "lane2"):
            rec[col] = clean(rec.get(col))
        for col in ("name", "role", "style1", "damage_type", "range_type",
                    "lane1", "power_spike", "resource"):
            rec[col] = clean(rec.get(col))
        records.append(rec)
    return records


def extract_global_weights(wb) -> list[dict]:
    ws = wb["Data-Input"]
    start, end = GLOBAL_WEIGHTS_ROWS
    records = []
    for row in range(start, end + 1):
        coefficient, value, description = row_values(ws, row, 1, 3)
        coefficient = clean(coefficient)
        if coefficient is None:
            continue
        records.append({
            "coefficient": coefficient,
            "value": float(value),
            "description": clean(description),
        })
    return records


def extract_role_matrix(wb) -> list[dict]:
    ws = wb["Data-Input"]
    defender_roles = [clean(v) for v in row_values(ws, ROLE_MATRIX_HEADER_ROW, 2, 7)]
    start, end = ROLE_MATRIX_DATA_ROWS
    records = []
    for row in range(start, end + 1):
        vals = row_values(ws, row, 1, 7)
        attacker_role = clean(vals[0])
        if attacker_role is None:
            continue
        for defender_role, points in zip(defender_roles, vals[1:]):
            if defender_role is None or points is None:
                continue
            records.append({
                "attacker_role": attacker_role,
                "defender_role": defender_role,
                "points": float(points),
            })
    return records


def extract_hard_counter_rules(wb) -> list[dict]:
    ws = wb["Data-Input"]
    start, end = HARD_COUNTER_ROWS
    seen = {}
    for row in range(start, end + 1):
        attacker, condition_type, condition_value, bonus, penalty, note = row_values(ws, row, 1, 6)
        attacker = clean(attacker)
        if attacker is None:
            continue  # skip blank rows
        key = (attacker, clean(condition_type), clean(condition_value))
        if key in seen:
            seen[key]["bonus_to_attacker"] += float(bonus)
            seen[key]["penalty_to_defender"] += float(penalty)
        else:
            seen[key] = {
                "attacker": attacker,
                "condition_type": clean(condition_type),
                "condition_value": clean(condition_value),
                "bonus_to_attacker": float(bonus),
                "penalty_to_defender": float(penalty),
                "note": clean(note),
            }
    return list(seen.values())


def extract_style_matrix(wb) -> list[dict]:
    ws = wb["Data-Input"]
    # Header row has an extra leading label cell ("Attacker\Defender") in col A;
    # tags run from column B onward, however many are populated.
    header_row = [c.value for c in ws[STYLE_MATRIX_HEADER_ROW]]
    defender_tags = [clean(v) for v in header_row[1:] if clean(v) is not None]
    n_tags = len(defender_tags)

    start, end = STYLE_MATRIX_DATA_ROWS
    records = []
    for row in range(start, end + 1):
        vals = row_values(ws, row, 1, 1 + n_tags)
        attacker_tag = clean(vals[0])
        if attacker_tag is None:
            continue
        for defender_tag, points in zip(defender_tags, vals[1:]):
            if points is None:
                continue
            records.append({
                "attacker_tag": attacker_tag,
                "defender_tag": defender_tag,
                "points": float(points),
            })
    return records


def extract_manual_overrides(wb) -> list[dict]:
    ws = wb["Data-Input"]
    start, end = MANUAL_OVERRIDES_ROWS
    records = []
    for row in range(start, end + 1):
        attacker, defender, score, note = row_values(ws, row, 1, 4)
        attacker = clean(attacker)
        defender = clean(defender)
        if attacker is None or defender is None:
            continue  # skip blank rows (the "Key" formula column is ignored)
        records.append({
            "attacker": attacker,
            "defender": defender,
            "score": float(score) if score is not None else None,
            "note": clean(note),
        })
    return records


# ----------------------------------------------------------------------------
# Upsert helper — supabase-py batches, chunked to stay well under request
# size limits for the larger tables (style_matrix ~ 484 rows, heroes ~ 133).
# ----------------------------------------------------------------------------
def upsert_in_chunks(client: Client, table: str, records: list[dict],
                      on_conflict: str, chunk_size: int = 500) -> int:
    if not records:
        print(f"  {table}: nothing to upsert (0 rows found)")
        return 0
    total = 0
    for i in range(0, len(records), chunk_size):
        chunk = records[i:i + chunk_size]
        client.table(table).upsert(chunk, on_conflict=on_conflict).execute()
        total += len(chunk)
    print(f"  {table}: upserted {total} rows")
    return total


def main():
    if len(sys.argv) != 2:
        sys.exit("Usage: python migrate_to_supabase.py <path_to_workbook.xlsx>")
    xlsx_path = sys.argv[1]

    print(f"Loading workbook: {xlsx_path}")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    client = get_supabase_client()

    print("Extracting data from workbook...")
    heroes = extract_heroes(wb)
    weights = extract_global_weights(wb)
    role_matrix = extract_role_matrix(wb)
    hard_counters = extract_hard_counter_rules(wb)
    style_matrix = extract_style_matrix(wb)
    overrides = extract_manual_overrides(wb)

    print("\nUpserting into Supabase...")
    counts = {}
    counts["heroes"] = upsert_in_chunks(client, "heroes", heroes, on_conflict="id")
    counts["global_weights"] = upsert_in_chunks(
        client, "global_weights", weights, on_conflict="coefficient"
    )
    counts["role_matrix"] = upsert_in_chunks(
        client, "role_matrix", role_matrix, on_conflict="attacker_role,defender_role"
    )

    # hard_counter_rules uses a unique index on (attacker, condition_type, condition_value).
    # This is an upsert so re-running the migration safely updates existing rules.
    if hard_counters:
        counts["hard_counter_rules"] = upsert_in_chunks(
            client, 
            "hard_counter_rules", 
            hard_counters, 
            on_conflict="attacker,condition_type,condition_value"
        )
    else:
        counts["hard_counter_rules"] = 0
        print(" hard_counter_rules: nothing to insert (0 rows found)")

    counts["style_matrix"] = upsert_in_chunks(
        client, "style_matrix", style_matrix, on_conflict="attacker_tag,defender_tag"
    )
    counts["manual_overrides"] = upsert_in_chunks(
        client, "manual_overrides", overrides, on_conflict="attacker,defender"
    )

    print("\n--- Row counts written ---")
    for table, n in counts.items():
        print(f"  {table:20s}: {n}")

    # ------------------------------------------------------------------
    # Sanity check
    # ------------------------------------------------------------------
    print("\n--- Sanity check ---")
    heroes_count_resp = client.table("heroes").select("id", count="exact").execute()
    actual_hero_count = heroes_count_resp.count
    if actual_hero_count == EXPECTED_HERO_COUNT:
        print(f"  OK: heroes table has {actual_hero_count} rows (expected {EXPECTED_HERO_COUNT})")
    else:
        print(
            f"  WARNING: heroes table has {actual_hero_count} rows, "
            f"expected {EXPECTED_HERO_COUNT}. Check the workbook / migration."
        )
        sys.exit(1)

    print("\nMigration complete.")


if __name__ == "__main__":
    main()
