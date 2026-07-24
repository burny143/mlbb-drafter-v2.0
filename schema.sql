-- ============================================================================
-- MLBB Hero Counter Model — Supabase (Postgres) schema
-- ============================================================================
-- Run this once against your Supabase project (SQL Editor, or via
-- `supabase db push` / psql). Safe to re-run: uses IF NOT EXISTS / drops
-- guarded by CASCADE only where noted.
-- ============================================================================

-- ----------------------------------------------------------------------------
-- 1) heroes — raw per-hero data (Heroes sheet)
-- ----------------------------------------------------------------------------
create table if not exists heroes (
    id               int primary key,
    name             text not null unique,
    role             text not null,
    role2            text,
    offense          int not null,
    ability_effects  int not null,
    durability       int not null,
    difficulty       int not null,
    style1           text not null,
    style2           text,
    damage_type      text not null,
    range_type       text not null,
    lane1            text not null,
    lane2            text,
    power_spike      text not null,
    resource         text not null,
    spike_order      int not null,
    has_antiheal     boolean not null default false
);

comment on table heroes is 'One row per MLBB hero: base stats, roles, tags/lanes. Source: Heroes sheet.';

-- ----------------------------------------------------------------------------
-- 2) global_weights — tunable coefficients (Data-Input Block 1, rows 7-16)
-- ----------------------------------------------------------------------------
create table if not exists global_weights (
    coefficient text primary key,
    value       numeric not null,
    description text
);

comment on table global_weights is 'Named multipliers/coefficients used by the score formula (role_mult, burst_mult, diff_mult, style_mult, spike_mult, dmgtype_same, dmgtype_diff, dmgtype_mixed, rangetype_mult, antiheal_mult).';

-- ----------------------------------------------------------------------------
-- 3) role_matrix — role vs role points, BEFORE role_mult (Block 2, rows 18-24)
-- ----------------------------------------------------------------------------
create table if not exists role_matrix (
    attacker_role text not null,
    defender_role text not null,
    points        numeric not null,
    primary key (attacker_role, defender_role)
);

comment on table role_matrix is '6x6 role matchup grid (Tank/Fighter/Assassin/Mage/Marksman/Support), points before role_mult.';

-- ----------------------------------------------------------------------------
-- 4) style_matrix — tag vs tag points, BEFORE style_mult (Block 5, rows 163-185)
-- ----------------------------------------------------------------------------
create table if not exists style_matrix (
    attacker_tag text not null,
    defender_tag text not null,
    points       numeric not null,
    primary key (attacker_tag, defender_tag)
);

comment on table style_matrix is '22x22 style/tag interaction grid (the 22 "In-use" tags from the Tag Glossary), points before style_mult.';

-- ----------------------------------------------------------------------------
-- 5) hard_counter_rules — named-hero / tag special cases (Block 3, rows 29-128)
-- ----------------------------------------------------------------------------
create table if not exists hard_counter_rules (
    id                   serial primary key,
    attacker             text not null,
    condition_type       text not null check (condition_type in ('Tag', 'Hero', 'Role', 'DamageType', 'Resource')),
    condition_value      text not null,
    bonus_to_attacker    numeric not null,
    penalty_to_defender  numeric not null,
    note                 text
);

comment on table hard_counter_rules is 'One row per hard-counter rule. condition_type is one of: Tag, Hero, Role, DamageType, Resource. Contribution = bonus_to_attacker - penalty_to_defender.';

create index if not exists idx_hard_counter_rules_attacker on hard_counter_rules (attacker);
create unique index if not exists idx_hard_counter_rules_unique on hard_counter_rules (attacker, condition_type, condition_value);

-- ----------------------------------------------------------------------------
-- 6) manual_overrides — force an exact score for one pair (Block 6, rows 190-1189)
-- ----------------------------------------------------------------------------
create table if not exists manual_overrides (
    attacker text not null,
    defender text not null,
    score    numeric not null,
    note     text,
    primary key (attacker, defender)
);

comment on table manual_overrides is 'One-directional (attacker -> defender) forced score that bypasses the formula entirely when present.';

-- ----------------------------------------------------------------------------
-- 7) counter_scores — precomputed output (populated later by a Python batch
--    job, NOT by this migration). Table is created here so the batch job has
--    somewhere to write to.
-- ----------------------------------------------------------------------------
create table if not exists counter_scores (
    attacker              text not null,
    defender              text not null,
    score                 numeric not null,
    role_advantage        numeric not null,
    stat                  numeric not null,
    difficulty_gap        numeric not null,
    damage_type_adv       numeric not null,
    power_spike_timing    numeric not null,
    style_matchup         numeric not null,
    hard_counter_bonus    numeric not null,
    matched_rules         jsonb default '[]'::jsonb,
    computed_at           timestamptz not null default now(),
    primary key (attacker, defender)
);

comment on table counter_scores is 'Precomputed component-by-component counter score for every attacker/defender pair. Populated by a separate batch job, not this migration.';

create index if not exists idx_counter_scores_defender on counter_scores (defender);
create index if not exists idx_counter_scores_score on counter_scores (score desc);

-- ----------------------------------------------------------------------------
-- Optional but recommended: foreign keys tying matchup tables back to heroes.
-- Left as NOT VALID / commented out by default because role_matrix and
-- style_matrix key on role/tag names, not hero names, and hard_counter_rules /
-- manual_overrides reference hero names that must already exist in `heroes`.
-- Uncomment once you've verified there are no orphaned name references.
-- ----------------------------------------------------------------------------
-- alter table hard_counter_rules
--     add constraint hard_counter_rules_attacker_fkey
--     foreign key (attacker) references heroes(name);
--
-- alter table manual_overrides
--     add constraint manual_overrides_attacker_fkey
--     foreign key (attacker) references heroes(name),
--     add constraint manual_overrides_defender_fkey
--     foreign key (defender) references heroes(name);
--
-- alter table counter_scores
--     add constraint counter_scores_attacker_fkey
--     foreign key (attacker) references heroes(name),
--     add constraint counter_scores_defender_fkey
--     foreign key (defender) references heroes(name);
