#!/usr/bin/env python3
"""
compute_counters.py

Recomputes the FULL attacker -> defender counter score matrix for every
hero pair and upserts the results into `counter_scores`.

Formula (mirrors the Excel model's "Documentation - Computations" sheet):

    TOTAL = Role Advantage
          + Stat
          + Difficulty Gap
          + Damage Type Advantage
          + Power Spike Timing
          + Style Matchup
          + Hard Counter Bonus

    ...then TOTAL is replaced entirely by a Manual Override score if one
    exists for that exact (attacker, defender) pair.

Component definitions:

    Role Advantage
        MAX(role_matrix[atk.role -> def.role],
            role_matrix[atk.role2 -> def.role])   (skip atk.role2 if null)
        * role_mult

    Stat
        ((atk.offense + atk.ability_effects) / 2 - def.durability) * burst_mult

    Difficulty Gap
        gap = atk.difficulty - def.difficulty
        gap * diff_mult   if gap > 20
        0                 otherwise

    Damage Type Advantage
        dmgtype_mixed   if atk.damage_type == "Mixed"
        dmgtype_same    elif atk.damage_type == def.damage_type
        dmgtype_diff    else

    Power Spike Timing
        (atk.spike_order - def.spike_order) * spike_mult

    Style Matchup
        MAX over up to 4 lookups of
            style_matrix[(atk.style1 or atk.style2)][(def.style1 or def.style2)]
        * style_mult
        (only combinations where both tags are non-null are considered)

    Hard Counter Bonus
        sum of (bonus_to_attacker - penalty_to_defender) for every row in
        hard_counter_rules where rule.attacker == atk.name AND:
          - condition_type == "Tag"        AND def.style1/style2 == condition_value
          - condition_type == "Hero"       AND def.name == condition_value
          - condition_type == "Role"       AND def.role/role2 == condition_value
          - condition_type == "DamageType" AND def.damage_type == condition_value
          - condition_type == "Resource"   AND def.resource == condition_value
        (multiple matching rules stack)

Usage:
    export SUPABASE_URL="https://xxxx.supabase.co"
    export SUPABASE_SERVICE_KEY="sb_secret_..."
    python compute_counters.py              # computes and upserts everything
    python compute_counters.py --dry-run    # prints first 10 rows, no writes
"""

import argparse
import os
import sys
from pathlib import Path
from typing import Any, Optional

from supabase import create_client, Client

# Load .env file if present
_dotenv = Path(__file__).parent / ".env"
if _dotenv.exists():
    for _line in _dotenv.read_text().splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip())

CHUNK_SIZE = 500
PAGE_SIZE = 1000  # supabase-py / PostgREST default row cap per request


# ----------------------------------------------------------------------------
# Client + generic paginated fetch (reference tables are small, but this is
# safe regardless of how large manual_overrides or hard_counter_rules grow).
# ----------------------------------------------------------------------------
def get_supabase_client() -> Client:
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_KEY")
    if not url or not key:
        sys.exit(
            "ERROR: SUPABASE_URL and SUPABASE_SERVICE_KEY must be set as "
            "environment variables. Never hardcode credentials in this script."
        )
    return create_client(url, key)


def fetch_all(client: Client, table: str) -> list[dict]:
    rows: list[dict] = []
    start = 0
    while True:
        end = start + PAGE_SIZE - 1
        resp = client.table(table).select("*").range(start, end).execute()
        batch = resp.data or []
        rows.extend(batch)
        if len(batch) < PAGE_SIZE:
            break
        start += PAGE_SIZE
    return rows


# ----------------------------------------------------------------------------
# Load reference data into memory, in the shapes the formula needs.
# ----------------------------------------------------------------------------
def load_reference_data(client: Client) -> dict[str, Any]:
    print("Loading reference tables from Supabase...")

    heroes = fetch_all(client, "heroes")
    weights_rows = fetch_all(client, "global_weights")
    role_rows = fetch_all(client, "role_matrix")
    style_rows = fetch_all(client, "style_matrix")
    hard_counter_rules = fetch_all(client, "hard_counter_rules")
    override_rows = fetch_all(client, "manual_overrides")

    weights = {r["coefficient"]: float(r["value"]) for r in weights_rows}
    role_matrix = {
        (r["attacker_role"], r["defender_role"]): float(r["points"])
        for r in role_rows
    }
    style_matrix = {
        (r["attacker_tag"], r["defender_tag"]): float(r["points"])
        for r in style_rows
    }
    manual_overrides = {
        (r["attacker"], r["defender"]): (
            float(r["score"]) if r["score"] is not None else None
        )
        for r in override_rows
    }

    print(f"  heroes: {len(heroes)}")
    print(f"  global_weights: {len(weights)}")
    print(f"  role_matrix: {len(role_matrix)}")
    print(f"  style_matrix: {len(style_matrix)}")
    print(f"  hard_counter_rules: {len(hard_counter_rules)}")
    print(f"  manual_overrides: {len(manual_overrides)}")

    return {
        "heroes": heroes,
        "weights": weights,
        "role_matrix": role_matrix,
        "style_matrix": style_matrix,
        "hard_counter_rules": hard_counter_rules,
        "manual_overrides": manual_overrides,
    }


# ----------------------------------------------------------------------------
# Formula components
# ----------------------------------------------------------------------------
def role_advantage(atk: dict, defn: dict, role_matrix: dict, role_mult: float) -> float:
    candidates = []
    v1 = role_matrix.get((atk["role"], defn["role"]))
    if v1 is not None:
        candidates.append(v1)
    if atk.get("role2"):
        v2 = role_matrix.get((atk["role2"], defn["role"]))
        if v2 is not None:
            candidates.append(v2)
    if not candidates:
        return 0.0
    return max(candidates) * role_mult


def stat_component(atk: dict, defn: dict, burst_mult: float) -> float:
    return ((atk["offense"] + atk["ability_effects"]) / 2 - defn["durability"]) * burst_mult


def difficulty_gap(atk: dict, defn: dict, diff_mult: float) -> float:
    gap = atk["difficulty"] - defn["difficulty"]
    return gap * diff_mult if gap > 20 else 0.0


def damage_type_advantage(atk: dict, defn: dict, weights: dict) -> float:
    if atk["damage_type"] == "Mixed":
        return weights["dmgtype_mixed"]
    if atk["damage_type"] == defn["damage_type"]:
        return weights["dmgtype_same"]
    return weights["dmgtype_diff"]


def power_spike_timing(atk: dict, defn: dict, spike_mult: float) -> float:
    return (atk["spike_order"] - defn["spike_order"]) * spike_mult


def style_matchup(atk: dict, defn: dict, style_matrix: dict, style_mult: float) -> float:
    atk_tags = [t for t in (atk.get("style1"), atk.get("style2")) if t]
    def_tags = [t for t in (defn.get("style1"), defn.get("style2")) if t]
    candidates = []
    for at in atk_tags:
        for dt in def_tags:
            v = style_matrix.get((at, dt))
            if v is not None:
                candidates.append(v)
    if not candidates:
        return 0.0
    return max(candidates) * style_mult


def hard_counter_bonus(atk: dict, defn: dict, rules: list[dict]) -> float:
    total = 0.0
    atk_name_lower = atk["name"].lower()
    for rule in rules:
        if rule["attacker"].lower() != atk_name_lower:
            continue
        ctype = rule["condition_type"]
        cval = rule["condition_value"]
        matched = False
        if ctype == "Tag":
            matched = cval.lower() in (defn.get("style1", "").lower(), defn.get("style2", "").lower())
        elif ctype == "Hero":
            matched = defn["name"].lower() == cval.lower()
        elif ctype == "Role":
            matched = cval.lower() in (defn.get("role", "").lower(), defn.get("role2", "").lower())
        elif ctype == "DamageType":
            matched = defn["damage_type"].lower() == cval.lower()
        elif ctype == "Resource":
            matched = defn["resource"].lower() == cval.lower()
        if matched:
            total += float(rule["bonus_to_attacker"]) - float(rule["penalty_to_defender"])
    return total


def compute_score(atk: dict, defn: dict, ref: dict) -> dict:
    weights = ref["weights"]
    override = ref["manual_overrides"].get((atk["name"], defn["name"]))

    ra = role_advantage(atk, defn, ref["role_matrix"], weights["role_mult"])
    st = stat_component(atk, defn, weights["burst_mult"])
    dg = difficulty_gap(atk, defn, weights["diff_mult"])
    dta = damage_type_advantage(atk, defn, weights)
    pst = power_spike_timing(atk, defn, weights["spike_mult"])
    sm = style_matchup(atk, defn, ref["style_matrix"], weights["style_mult"])
    hcb = hard_counter_bonus(atk, defn, ref["hard_counter_rules"])
    rta = range_type_advantage(atk, defn, weights["rangetype_mult"])
    aha = antiheal_advantage(atk, defn, weights["antiheal_mult"])

    total = ra + st + dg + dta + pst + sm + hcb + rta + aha
    if override is not None:
        total = override

    return {
        "attacker": atk["name"],
        "defender": defn["name"],
        "score": total,
        "role_advantage": ra,
        "stat": st,
        "difficulty_gap": dg,
        "damage_type_adv": dta,
        "power_spike_timing": pst,
        "style_matchup": sm,
        "hard_counter_bonus": hcb,
    }


def compute_all(ref: dict) -> list[dict]:
    heroes = ref["heroes"]
    results = []
    for atk in heroes:
        for defn in heroes:
            if atk["id"] == defn["id"]:
                continue
            results.append(compute_score(atk, defn, ref))
    return results


# ----------------------------------------------------------------------------
# Upsert helper
# ----------------------------------------------------------------------------
def upsert_in_chunks(client: Client, table: str, records: list[dict],
                      on_conflict: str, chunk_size: int = CHUNK_SIZE) -> int:
    if not records:
        print(f"  {table}: nothing to upsert (0 rows)")
        return 0
    total = 0
    for i in range(0, len(records), chunk_size):
        chunk = records[i:i + chunk_size]
        client.table(table).upsert(chunk, on_conflict=on_conflict).execute()
        total += len(chunk)
        print(f"  {table}: upserted {total}/{len(records)} rows", end="\r")
    print(f"  {table}: upserted {total} rows" + " " * 10)
    return total


def print_preview(records: list[dict], n: int = 10) -> None:
    print(f"\n--- DRY RUN: first {min(n, len(records))} of {len(records)} computed rows ---")
    cols = ["attacker", "defender", "score", "role_advantage", "stat",
            "difficulty_gap", "damage_type_adv", "power_spike_timing",
            "style_matchup", "hard_counter_bonus"]
    for r in records[:n]:
        print({c: round(r[c], 3) if isinstance(r[c], float) else r[c] for c in cols})

MOBILITY_TAGS = {"Dash/Blink", "High Mobility", "Global Dive", "Mobility"}
SUSTAIN_HEAL_TAGS = {"Sustain", "Healing"}

def range_type_advantage(atk: dict, defn: dict, rangetype_mult: float) -> float:
    atk_tags = {atk.get("style1"), atk.get("style2")}
    if atk_tags & MOBILITY_TAGS and defn.get("range_type") == "Ranged":
        return rangetype_mult
    return 0.0

def antiheal_advantage(atk: dict, defn: dict, antiheal_mult: float) -> float:
    def_tags = {defn.get("style1"), defn.get("style2")}
    if atk.get("has_antiheal") and (def_tags & SUSTAIN_HEAL_TAGS):
        return antiheal_mult
    return 0.0


def main():
    parser = argparse.ArgumentParser(description="Recompute the MLBB counter score matrix.")
    parser.add_argument("--dry-run", action="store_true",
                         help="Compute everything and print the first 10 rows, but do not write to the DB.")
    args = parser.parse_args()

    client = get_supabase_client()
    ref = load_reference_data(client)

    hero_count = len(ref["heroes"])
    expected_pairs = hero_count * (hero_count - 1)
    print(f"\nComputing scores for {hero_count} heroes "
          f"({expected_pairs} attacker->defender pairs, self-pairs skipped)...")

    results = compute_all(ref)

    if len(results) != expected_pairs:
        print(f"  WARNING: computed {len(results)} rows, expected {expected_pairs}. "
              f"Check for duplicate hero names or IDs.")

    if args.dry_run:
        print_preview(results)
        print("\nDry run complete. No rows written to counter_scores.")
        return

    print("\nUpserting into Supabase...")
    n = upsert_in_chunks(client, "counter_scores", results, on_conflict="attacker,defender")

    print("\n--- Sanity check ---")
    count_resp = client.table("counter_scores").select("attacker", count="exact").execute()
    actual = count_resp.count
    if actual == expected_pairs:
        print(f"  OK: counter_scores has {actual} rows (expected {expected_pairs})")
    else:
        print(f"  WARNING: counter_scores has {actual} rows, expected {expected_pairs}.")
        sys.exit(1)

    print(f"\nMigration complete. {n} rows upserted.")


if __name__ == "__main__":
    main()
