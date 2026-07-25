#!/usr/bin/env python3
"""
compute_counters.py

Recomputes the FULL attacker -> defender counter score matrix for every
hero pair and upserts the results into `counter_scores`.

Formula (mirrors the Excel model's "Documentation - Computations" sheet):

TOTAL = Role Advantage
      + Stat
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

    Stat (damage + survivability)
        damage_score  = ((atk.offense + atk.ability_effects) / 2 - def.durability) * burst_mult
        surv_score    = (atk.durability - (def.offense + def.ability_effects) / 2) * burst_mult * 0.3

    Damage Type Advantage
        dmgtype_mixed   if atk.damage_type == "Mixed"
        dmgtype_same    elif atk.damage_type == def.damage_type
        dmgtype_diff    else
        + TRUE_DAMAGE_BONUS (2.0) if atk.has_true_damage

    Power Spike Timing
        Categorical 3×3 matrix (Early/Mid/Late bands):
          Early vs Late = +8, Late vs Early = -6
          Mid vs Late = +4, Late vs Mid = -2
        + fine_tune = (spike_order_diff) * 0.5

    Style Matchup
        (best * 0.6 + avg(rest) * 0.4) over up to 4 lookups of
            style_matrix[(atk.style1 or atk.style2)][(def.style1 or def.style2)]
        * style_mult

    Hard Counter Bonus
        sum of (bonus_to_attacker - penalty_to_defender) for every matching
        hard_counter_rule (case-insensitive match via .lower())

    Range Type Advantage
        rangetype_mult   if atk has mobility tag AND def.range_type == "Ranged"

    Anti-Heal Advantage
        antiheal_mult    if atk.has_antiheal AND def has Sustain/Heal tag

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

# Power Spike categorical bands
SPIKE_BANDS = {1: "Early", 2: "Early", 3: "Mid", 4: "Mid", 5: "Late"}

SPIKE_MATRIX = {
    ("Early","Early"): 0,   ("Early","Mid"):   4,
    ("Early","Late"):  8,   ("Mid","Early"):  -2,
    ("Mid","Mid"):     0,   ("Mid","Late"):    4,
    ("Late","Early"): -6,   ("Late","Mid"):   -2,
    ("Late","Late"):   0,
}

TRUE_DAMAGE_BONUS = 2.0
SURVIVABILITY_WEIGHT = 0.3

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
        (r["attacker"].strip(), r["defender"].strip()): (
            float(r["score"]) if r["score"] is not None else None
        )
        for r in override_rows
    }

    # --- Duplicate hero name validation ---------------------------------
    # counter_scores is keyed by (attacker name, defender name), not by
    # hero id, and the upsert's on_conflict target is "attacker,defender".
    # If two heroes share the same name but different ids, they are NOT
    # skipped as a self-pair (that check is id-based) but their rows WILL
    # collide and silently overwrite each other during upsert, producing
    # a row count lower than expected with no error raised. Fail loudly
    # here instead so it isn't mistaken for the unrelated "hero count
    # changed" case described in AGENTS.md's troubleshooting section.
    name_to_ids: dict[str, list] = {}
    for h in heroes:
        name_to_ids.setdefault(h["name"], []).append(h["id"])
    duplicates = {name: ids for name, ids in name_to_ids.items() if len(ids) > 1}
    if duplicates:
        details = "; ".join(f"{name!r} (ids: {ids})" for name, ids in duplicates.items())
        sys.exit(
            "ERROR: duplicate hero names found in `heroes` table — this will "
            "cause counter_scores rows to silently collide and overwrite each "
            f"other. Fix the duplicate name(s) before recomputing: {details}"
        )

    # --- Pre-group hard counter rules by attacker ------------------------
    # hard_counter_bonus() is called once per (attacker, defender) pair
    # (~heroes^2 times). Scanning the full rule list every call is
    # O(heroes^2 * rules); grouping once here makes each call O(rules
    # matching this attacker) instead.
    rules_by_attacker: dict[str, list[dict]] = {}
    for rule in hard_counter_rules:
        key = rule["attacker"].strip().lower()
        rules_by_attacker.setdefault(key, []).append(rule)

    required_weight_keys = {
        "role_mult", "burst_mult", "dmgtype_mixed", "dmgtype_same",
        "dmgtype_diff", "spike_mult", "style_mult", "rangetype_mult",
        "antiheal_mult",
    }
    missing_keys = required_weight_keys - weights.keys()
    if missing_keys:
        sys.exit(
            "ERROR: global_weights table is missing required coefficient(s): "
            f"{sorted(missing_keys)}. Add them to global_weights before running "
            "this script."
        )

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
        "rules_by_attacker": rules_by_attacker,
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
    damage_score = ((atk["offense"] + atk["ability_effects"]) / 2 - defn["durability"]) * burst_mult
    surv_score = (atk["durability"] - (defn["offense"] + defn["ability_effects"]) / 2) * burst_mult * SURVIVABILITY_WEIGHT
    return damage_score + surv_score


def damage_type_advantage(atk: dict, defn: dict, weights: dict) -> float:
    if atk["damage_type"] == "Mixed":
        base = weights["dmgtype_mixed"]
    elif atk["damage_type"] == defn["damage_type"]:
        base = weights["dmgtype_same"]
    else:
        base = weights["dmgtype_diff"]
    if atk.get("has_true_damage"):
        base += TRUE_DAMAGE_BONUS
    return base


DEFAULT_SPIKE_ORDER = 3  # "Mid" band fallback for null/missing spike_order

def power_spike_timing(atk: dict, defn: dict, spike_mult: float) -> float:
    # spike_order can be NULL in the heroes table. SPIKE_BANDS.get(..., "Mid")
    # already tolerates None for the band lookup, but the fine_tune subtraction
    # below does raw arithmetic on the same field and previously crashed the
    # whole batch (TypeError: unsupported operand type(s) for -: 'NoneType'
    # and 'int') the first time it hit a hero with no spike_order set.
    a_spike = atk["spike_order"] if atk["spike_order"] is not None else DEFAULT_SPIKE_ORDER
    d_spike = defn["spike_order"] if defn["spike_order"] is not None else DEFAULT_SPIKE_ORDER
    a_band = SPIKE_BANDS.get(a_spike, "Mid")
    d_band = SPIKE_BANDS.get(d_spike, "Mid")
    base = SPIKE_MATRIX.get((a_band, d_band), 0)
    fine_tune = (a_spike - d_spike) * 0.5
    return (base + fine_tune) * spike_mult


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
    candidates.sort(reverse=True)
    if len(candidates) <= 1:
        return candidates[0] * style_mult
    best = candidates[0]
    rest_avg = sum(candidates[1:]) / len(candidates[1:])
    return (best * 0.6 + rest_avg * 0.4) * style_mult


def _norm(s: Optional[str]) -> str:
    """Whitespace- and case-insensitive normalization for rule matching."""
    return (s or "").strip().lower()


def hard_counter_bonus(atk: dict, defn: dict, rules_by_attacker: dict[str, list[dict]]) -> tuple[float, list]:
    total = 0.0
    matched_rules = []
    # Pre-grouped by normalized attacker name: O(rules for this attacker)
    # instead of scanning every rule for every (attacker, defender) pair.
    candidate_rules = rules_by_attacker.get(_norm(atk["name"]), [])
    defn_name = _norm(defn["name"])
    defn_style1 = _norm(defn.get("style1"))
    defn_style2 = _norm(defn.get("style2"))
    defn_role = _norm(defn.get("role"))
    defn_role2 = _norm(defn.get("role2"))
    defn_damage_type = _norm(defn.get("damage_type"))
    defn_resource = _norm(defn.get("resource"))

    for rule in candidate_rules:
        ctype = rule["condition_type"]
        cval_raw = rule["condition_value"]
        cval = _norm(cval_raw)
        matched = False
        if ctype == "Tag":
            matched = cval in (defn_style1, defn_style2)
        elif ctype == "Hero":
            matched = defn_name == cval
        elif ctype == "Role":
            matched = cval in (defn_role, defn_role2)
        elif ctype == "DamageType":
            matched = defn_damage_type == cval
        elif ctype == "Resource":
            matched = defn_resource == cval
        if matched:
            bonus = float(rule["bonus_to_attacker"])
            penalty = float(rule["penalty_to_defender"])
            total += bonus - penalty
            matched_rules.append({
                "type": ctype,
                "value": cval_raw,
                "bonus": bonus,
                "penalty": penalty,
                "note": rule.get("note") or "",
            })
    return total, matched_rules


def compute_score(atk: dict, defn: dict, ref: dict) -> dict:
    weights = ref["weights"]
    override = ref["manual_overrides"].get((atk["name"], defn["name"]))

    ra = role_advantage(atk, defn, ref["role_matrix"], weights["role_mult"])
    st = stat_component(atk, defn, weights["burst_mult"])
    dta = damage_type_advantage(atk, defn, weights)
    pst = power_spike_timing(atk, defn, weights["spike_mult"])
    sm = style_matchup(atk, defn, ref["style_matrix"], weights["style_mult"])
    hcb, matched_rules = hard_counter_bonus(atk, defn, ref["rules_by_attacker"])
    rta = range_type_advantage(atk, defn, weights["rangetype_mult"])
    aha = antiheal_advantage(atk, defn, weights["antiheal_mult"])

    total = ra + st + dta + pst + sm + hcb + rta + aha
    if override is not None:
        total = override

    return {
        "attacker": atk["name"],
        "defender": defn["name"],
        "score": total,
        "role_advantage": ra,
        "stat": st,
        "difficulty_gap": 0.0,
        "damage_type_adv": dta,
        "power_spike_timing": pst,
        "style_matchup": sm,
        "hard_counter_bonus": hcb,
        "range_type_adv": rta,
        "antiheal_adv": aha,
        "matched_rules": matched_rules,
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
            "style_matchup", "hard_counter_bonus", "range_type_adv",
            "antiheal_adv"]
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