"""
enrich.py — Browser edition.

Enrichment with two browser-safe providers only:
  1. custom_file  — emission_factors.xlsx uploaded by user (parsed in-memory)
  2. lifetime_defaults — built-in lookup table

All HTTP providers (Rejoose, Climatiq, EPD, Resilio, Ecoinvent, tscircuit) are removed.
No disk I/O. ef_table is passed in as a dict (parsed once from BytesIO by the caller).
"""

import math
import statistics as _stats
from typing import Optional

import openpyxl

# ── Provider hierarchies (browser-only subset) ─────────────────────────────────

EMISSION_HIERARCHY = ["custom_file"]
POWER_HIERARCHY    = ["custom_file"]
LIFETIME_HIERARCHY = ["lifetime_defaults"]

# ── Schema definitions ─────────────────────────────────────────────────────────

ENRICHABLE_FIELDS = {
    "cell_site": [
        {"field": "electricity_emission_factor", "unit_field": "electricity_emission_factor_unit", "default_unit": "kgCO2eq/kWh", "precondition": "electricity_source", "provider_group": "emission"},
        {"field": "fuel_emission_factor",        "unit_field": "fuel_emission_factor_unit",        "default_unit": "kgCO2eq/L",   "precondition": "fuel_type",          "provider_group": "emission"},
        {"field": "refrigerant_emission_factor", "unit_field": "refrigerant_emission_factor_unit", "default_unit": "kgCO2eq/m3",  "precondition": "refrigerant_type",   "provider_group": "emission"},
    ],
    "active": [
        {"field": "life_time",                    "unit_field": None,                                "default_unit": None,           "precondition": None,                    "provider_group": "lifetime"},
        {"field": "production_emissions",          "unit_field": "production_emissions_unit",         "default_unit": "kgCO2eq/unit", "precondition": None,                    "provider_group": "emission"},
        {"field": "endoflife_emissions",           "unit_field": "endoflife_emissions_unit",          "default_unit": "kgCO2eq/unit", "precondition": None,                    "provider_group": "emission"},
        {"field": "power_source_emission_factor",  "unit_field": "power_source_emission_factor_unit", "default_unit": "kgCO2eq/kWh",  "precondition": "power_source",          "provider_group": "emission"},
        {"field": "installation_emission_factor",  "unit_field": "installation_emission_factor_unit", "default_unit": None,           "precondition": "installation_quantity", "provider_group": "emission"},
        {"field": "maintenance_emission_factor",   "unit_field": "maintenance_emission_factor_unit",  "default_unit": None,           "precondition": "maintenance_quantity",  "provider_group": "emission"},
        {"field": "power_idle",                    "unit_field": "power_idle_unit",                   "default_unit": "W",            "precondition": None,                    "provider_group": "power", "skip_if": "power_quantity"},
        {"field": "power_max",                     "unit_field": "power_max_unit",                    "default_unit": "W",            "precondition": None,                    "provider_group": "power", "skip_if": "power_quantity"},
    ],
    "passive": [
        {"field": "life_time",                    "unit_field": None,                                "default_unit": None,           "precondition": None,                    "provider_group": "lifetime"},
        {"field": "production_emissions",          "unit_field": "production_emissions_unit",         "default_unit": "kgCO2eq/unit", "precondition": None,                    "provider_group": "emission"},
        {"field": "endoflife_emissions",           "unit_field": "endoflife_emissions_unit",          "default_unit": "kgCO2eq/unit", "precondition": None,                    "provider_group": "emission"},
        {"field": "installation_emission_factor",  "unit_field": "installation_emission_factor_unit", "default_unit": None,           "precondition": "installation_quantity", "provider_group": "emission"},
        {"field": "maintenance_emission_factor",   "unit_field": "maintenance_emission_factor_unit",  "default_unit": None,           "precondition": "maintenance_quantity",  "provider_group": "emission"},
    ],
    "infrastructure": [
        {"field": "life_time",                    "unit_field": None,                                "default_unit": None,           "precondition": None,                    "provider_group": "lifetime"},
        {"field": "production_emissions",          "unit_field": "production_emissions_unit",         "default_unit": "kgCO2eq/unit", "precondition": None,                    "provider_group": "emission"},
        {"field": "endoflife_emissions",           "unit_field": "endoflife_emissions_unit",          "default_unit": "kgCO2eq/unit", "precondition": None,                    "provider_group": "emission"},
        {"field": "installation_emission_factor",  "unit_field": "installation_emission_factor_unit", "default_unit": None,           "precondition": "installation_quantity", "provider_group": "emission"},
        {"field": "maintenance_emission_factor",   "unit_field": "maintenance_emission_factor_unit",  "default_unit": None,           "precondition": "maintenance_quantity",  "provider_group": "emission"},
    ],
}

CONFIDENCE = {
    "custom_file":       1.0,
    "cell_site":         1.0,
    "lifetime_defaults": 0.5,
}

_POWER_SOURCE_TO_SITE_EF = {
    "electricity": ("electricity_emission_factor", "electricity_emission_factor_unit"),
    "battery":     ("electricity_emission_factor", "electricity_emission_factor_unit"),
    "fuel":        ("fuel_emission_factor",         "fuel_emission_factor_unit"),
    "refrigerant": ("refrigerant_emission_factor",  "refrigerant_emission_factor_unit"),
}

_KNOWN_ENRICHABLE_FIELDS = frozenset({
    "production_emissions", "endoflife_emissions",
    "power_idle", "power_max", "power_source_emission_factor",
    "installation_emission_factor", "maintenance_emission_factor", "life_time",
    "electricity_emission_factor", "fuel_emission_factor", "refrigerant_emission_factor",
})

# ── Group uncertainty ──────────────────────────────────────────────────────────

_GROUP_TYPE_FIELD = {
    "active":         "active_subtype",
    "passive":        "passive_type",
    "infrastructure": "infrastructure_type",
}

_EMBODIED_SD_FIELDS = [
    "production_emissions",
    "endoflife_emissions",
    "installation_emission_factor",
    "maintenance_emission_factor",
]

_POWER_LOAD_LOW  = 0.3
_POWER_LOAD_HIGH = 1.0


def _group_stdev(values: list) -> float:
    nums = [v for v in values if v is not None]
    if len(nums) < 2:
        return 0.0
    try:
        return _stats.stdev(nums)
    except Exception:
        return 0.0


def compute_group_uncertainty(rows: list, schema: str) -> tuple:
    type_field = _GROUP_TYPE_FIELD.get(schema, "")

    groups: dict = {}
    for row in rows:
        t = str(row.get(type_field, "") or "").strip() or "__unknown__"
        if t not in groups:
            groups[t] = {f: [] for f in _EMBODIED_SD_FIELDS}
        for f in _EMBODIED_SD_FIELDS:
            v = row.get(f)
            if v is not None:
                try:
                    groups[t][f].append(float(v))
                except (ValueError, TypeError):
                    pass

    group_sds = {t: {f: _group_stdev(vals) for f, vals in fv.items()}
                 for t, fv in groups.items()}

    pq_groups: dict = {}
    if schema == "active":
        for row in rows:
            if row.get("power_quantity") is None:
                continue
            t    = str(row.get(type_field, "") or "").strip() or "__unknown__"
            unit = str(row.get("power_unit", "") or "").strip().lower()
            v = row.get("power_quantity")
            if v is not None:
                try:
                    pq_groups.setdefault((t, unit), []).append(float(v))
                except (ValueError, TypeError):
                    pass
    pq_sds = {k: _group_stdev(vals) for k, vals in pq_groups.items()}

    for row in rows:
        t   = str(row.get(type_field, "") or "").strip() or "__unknown__"
        sds = group_sds.get(t, {})
        for f in _EMBODIED_SD_FIELDS:
            row[f"{f}_sd"] = sds.get(f, 0.0)

        if schema == "active":
            unit = str(row.get("power_unit", "") or "").strip().lower()
            row["power_quantity_sd"] = (
                pq_sds.get((t, unit), 0.0)
                if row.get("power_quantity") is not None else None
            )
            p_idle_v = row.get("power_idle")
            p_max_v  = row.get("power_max")
            if p_idle_v is not None and p_max_v is not None:
                try:
                    p_idle = float(p_idle_v)
                    p_max  = float(p_max_v)
                    if str(row.get("power_idle_unit", "") or "").strip().lower() == "kw":
                        p_idle *= 1000.0
                    if str(row.get("power_max_unit",  "") or "").strip().lower() == "kw":
                        p_max  *= 1000.0
                    row["power_estimated_low_w"]  = round(p_idle + _POWER_LOAD_LOW  * (p_max - p_idle), 4)
                    row["power_estimated_high_w"] = round(p_idle + _POWER_LOAD_HIGH * (p_max - p_idle), 4)
                except (ValueError, TypeError):
                    row["power_estimated_low_w"]  = None
                    row["power_estimated_high_w"] = None
            else:
                row["power_estimated_low_w"]  = None
                row["power_estimated_high_w"] = None
        else:
            row["power_quantity_sd"]      = None
            row["power_estimated_low_w"]  = None
            row["power_estimated_high_w"] = None

    new_cols = [f"{f}_sd" for f in _EMBODIED_SD_FIELDS]
    if schema == "active":
        new_cols += ["power_quantity_sd", "power_estimated_low_w", "power_estimated_high_w"]

    return rows, new_cols


# ── Lifetime defaults ──────────────────────────────────────────────────────────

LIFETIME_DEFAULTS = {
    "WLS": 8, "SWITCH": 8, "switch": 8, "KRAFT": 12, "ROUTER": 8, "router": 8,
    "DC-HÅLLARE": 15, "aggregation_router": 8, "AIR": 8, "cellular_modem": 8,
    "chassis": 10, "edge_platform": 8, "firewall": 8, "gateway": 8,
    "radio_unit": 8, "baseband_unit": 8, "base_station": 10, "ISAM": 10,
    "memory_card": 5, "antenna": 15, "camera": 7, "sensor": 7, "SFP": 7,
    "light": 10, "BCI": 10, "DIN/EN": 15, "SLA/VRLA": 5,
    "standby_power_generator": 15, "prime_power_generator": 15,
    "portable_industrial_generator": 10, "inverter_generator": 12,
    "container_sized_generator": 15, "water_cooled_systems": 15,
    "air_cooled_systems": 15, "industrial_chillers": 15,
    "evaporative_cooling_systems": 12, "hybrid_system": 12, "specialized_cooling": 15,
    "generator": 15, "cooling": 15, "fire_suppression": 15, "electrical_equipment": 10,
    "fiber_cable": 25, "electrical_cables": 25, "COAX": 20, "splitters": 20,
    "shelters": 25, "cabinets": 20, "plugs": 10, "fencing": 25,
    "steel": 30, "aluminum": 30, "plastic": 15,
    "tower": 40, "mast": 35, "rooftop_mount": 25, "pole": 25,
    "building": 40, "underground": 40, "container": 20,
    "real estate": 40, "manhole": 40, "concrete": 40, "ducts & pipes": 30,
}

# ── In-memory cache ────────────────────────────────────────────────────────────

_cache: dict = {}

def _cache_get(provider, field, search_key):
    return _cache.get((provider, field, search_key))

def _cache_set(provider, field, search_key, result):
    _cache[(provider, field, search_key)] = result

# ── Search key derivation ──────────────────────────────────────────────────────

def get_search_keys(row: dict, schema: str, field: str) -> list:
    def _vals(*keys):
        return [str(row[k]).strip() for k in keys
                if row.get(k) and str(row.get(k, "")).strip()]

    if field == "life_time":
        if schema == "active":
            return _vals("active_subtype", "active_type")
        elif schema == "passive":
            return _vals("passive_type")
        elif schema == "infrastructure":
            return _vals("infrastructure_type")
    elif field in ("production_emissions", "endoflife_emissions"):
        if schema == "active":
            return _vals("manufacture_part_number", "active_subtype", "active_type", "brand")
        elif schema == "passive":
            return _vals("manufacture_part_number", "passive_type", "brand")
        elif schema == "infrastructure":
            return _vals("manufacture_part_number", "infrastructure_type", "brand")
    elif field in ("power_idle", "power_max"):
        return _vals("manufacture_part_number", "chip_id", "active_subtype", "active_type", "brand")
    elif field == "power_source_emission_factor":
        ps = str(row.get("power_source", "")).strip()
        keys = []
        if ps:
            keys += [ps, f"M-{ps}"]
        keys.append("power_source_emission_factor")
        return keys
    elif field == "electricity_emission_factor":
        src = str(row.get("electricity_source", "")).strip()
        return ([src, f"M-{src}"] if src else []) + ["electricity_emission_factor"]
    elif field == "fuel_emission_factor":
        ft = str(row.get("fuel_type", "")).strip()
        return ([ft, f"M-{ft}"] if ft else []) + ["fuel_emission_factor"]
    elif field == "refrigerant_emission_factor":
        rt = str(row.get("refrigerant_type", "")).strip()
        return ([rt, f"M-{rt}"] if rt else []) + ["refrigerant_emission_factor"]
    elif field == "installation_emission_factor":
        return _vals("installation_method")
    elif field == "maintenance_emission_factor":
        return _vals("maintenance_method")
    return []

# ── Precondition ───────────────────────────────────────────────────────────────

def precondition_met(row: dict, field_cfg: dict) -> bool:
    precond = field_cfg.get("precondition")
    if precond is not None:
        v = row.get(precond)
        if v is None or str(v).strip() == "":
            return False
    skip_if = field_cfg.get("skip_if")
    if skip_if is not None:
        v = row.get(skip_if)
        if v is not None and str(v).strip() != "":
            return False
    return True

# ── Unit validation ────────────────────────────────────────────────────────────

_PREFIX_CANON = {
    "kgco2eq": "kgCO2eq", "kgco2e": "kgCO2eq",
    "kg co2e": "kgCO2eq", "kg co2eq": "kgCO2eq",
    "gco2eq":  "gCO2eq",  "gco2e":   "gCO2eq",
    "g co2e":  "gCO2eq",  "g co2eq": "gCO2eq",
    "tco2eq":  "tCO2eq",  "tco2e":   "tCO2eq",
    "t co2e":  "tCO2eq",  "t co2eq": "tCO2eq",
}

_SCALE = {
    ("gCO2eq", "kgCO2eq"): 0.001,
    ("tCO2eq", "kgCO2eq"): 1000.0,
}


def _normalise_unit(unit_str: str) -> str:
    s = str(unit_str).strip()
    if "/" in s:
        prefix, _, suffix = s.partition("/")
        normed = _PREFIX_CANON.get(prefix.strip().lower(), prefix.strip())
        return f"{normed}/{suffix}"
    return _PREFIX_CANON.get(s.lower(), s)


def validate_unit(value: float, returned_unit: str, expected_unit: str) -> Optional[float]:
    nr = _normalise_unit(returned_unit)
    ne = _normalise_unit(expected_unit)
    if nr == ne:
        return float(value)
    r_parts = nr.split("/", 1)
    e_parts = ne.split("/", 1)
    if len(r_parts) == 2 and len(e_parts) == 2 and r_parts[1] == e_parts[1]:
        scale = _SCALE.get((r_parts[0], e_parts[0]))
        if scale is not None:
            return float(value) * scale
    return None

# ── Custom file provider ───────────────────────────────────────────────────────

def parse_ef_table(wb: openpyxl.Workbook) -> dict:
    """Parse an emission_factors.xlsx workbook into the ef_table lookup dict."""
    ef_table = {}
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))

    header_row_idx = None
    for i, row in enumerate(all_rows):
        if row and str(row[0]).strip().lower() == "id":
            header_row_idx = i
            break
    if header_row_idx is None:
        return ef_table

    headers = [str(h).strip().lower() if h else "" for h in all_rows[header_row_idx]]
    for row in all_rows[header_row_idx + 1:]:
        if not row or not row[0]:
            continue
        d = dict(zip(headers, row))
        row_id      = str(d.get("id", "")).strip()
        row_field   = str(d.get("field", "")).strip() if d.get("field") else ""
        if not row_field:
            row_notes = str(d.get("notes", "")).strip() if d.get("notes") else ""
            if row_notes in _KNOWN_ENRICHABLE_FIELDS:
                row_field = row_notes
        row_country = str(d.get("country", "")).strip() if d.get("country") else ""
        ef          = d.get("emission_factor")
        unit        = str(d.get("unit", "")).strip() if d.get("unit") else ""
        if not row_id or ef is None:
            continue
        try:
            ef_table[(row_id, row_field, row_country)] = {
                "value": float(ef), "unit": unit, "source": "custom_file"
            }
        except (ValueError, TypeError):
            pass
    return ef_table


def _fetch_custom_file(search_key: str, field: str, ef_table: dict, country: str = "") -> dict:
    """Look up in the pre-parsed ef_table dict."""
    c = country or ""
    return (
        ef_table.get((search_key, field, c)) or
        ef_table.get((search_key, field, "")) or
        ef_table.get((search_key, "", c)) or
        ef_table.get((search_key, "", "")) or
        {}
    )


def _fetch_lifetime_defaults(search_key: str, field: str) -> dict:
    years = LIFETIME_DEFAULTS.get(search_key)
    if years is not None:
        return {"value": years, "unit": "years", "source": "lifetime_defaults"}
    return {}

# ── Row enrichment ─────────────────────────────────────────────────────────────

def enrich_row(row: dict, schema: str, provider: str, field_cfgs: list,
               ef_table: dict = None, filled_log: list = None) -> int:
    filled = 0
    for fc in field_cfgs:
        fname      = fc["field"]
        unit_field = fc.get("unit_field")

        v = row.get(fname)
        if v is not None and str(v).strip():
            continue
        if not precondition_met(row, fc):
            continue
        if unit_field is not None:
            expected_unit = row.get(unit_field) or fc.get("default_unit")
            if not expected_unit:
                continue
        else:
            expected_unit = None

        country = str(row.get("country", "")).strip()
        for sk in get_search_keys(row, schema, fname):
            cached = _cache_get(provider, fname, sk)
            if cached is not None:
                result = cached
            elif provider == "custom_file" and ef_table is not None:
                result = _fetch_custom_file(sk, fname, ef_table, country)
                if result:
                    _cache_set(provider, fname, sk, result)
            elif provider == "lifetime_defaults":
                result = _fetch_lifetime_defaults(sk, fname)
                if result:
                    _cache_set(provider, fname, sk, result)
            else:
                result = {}

            if not result:
                continue

            if expected_unit is not None:
                converted = validate_unit(result["value"], result.get("unit", ""), expected_unit)
                if converted is None:
                    continue
                row[fname]      = converted
                row[unit_field] = expected_unit
            else:
                row[fname] = result["value"]

            row[f"{fname}_source"]     = result["source"]
            row[f"{fname}_confidence"] = CONFIDENCE.get(result["source"], 0.5)
            if filled_log is not None:
                filled_log.append((fname, sk, row[fname]))
            filled += 1
            break
    return filled

# ── Cell site EF inheritance ───────────────────────────────────────────────────

def _enrich_from_cell_site(data_rows: list, cell_site_lookup: dict) -> int:
    filled = 0
    for row in data_rows:
        if row.get("power_source_emission_factor") is not None and \
                str(row.get("power_source_emission_factor", "")).strip():
            continue
        cs_id = str(row.get("cell_site_id", "") or "").strip()
        if not cs_id:
            continue
        site = cell_site_lookup.get(cs_id)
        if not site:
            continue
        ps = str(row.get("power_source", "") or "").strip().lower()
        ef_info = _POWER_SOURCE_TO_SITE_EF.get(ps)
        if not ef_info:
            continue
        ef_field, unit_field = ef_info
        ef_val = None
        raw = site.get(ef_field)
        if raw is not None:
            try:
                ef_val = float(raw)
            except (ValueError, TypeError):
                pass
        if ef_val is None:
            continue
        row["power_source_emission_factor"]            = ef_val
        row["power_source_emission_factor_unit"]       = site.get(unit_field) or ""
        row["power_source_emission_factor_source"]     = "cell_site"
        row["power_source_emission_factor_confidence"] = CONFIDENCE["cell_site"]
        filled += 1
    return filled

# ── Main in-memory enrichment entry point ─────────────────────────────────────

def enrich_in_memory(rows: list, schema: str,
                     ef_table: dict = None,
                     cell_site_lookup: dict = None) -> tuple:
    """
    Enrich a list of row dicts in-memory.

    Provider hierarchy: cell_site EF inheritance → custom_file → lifetime_defaults.
    Returns (enriched_rows, enrichment_summary_dict, unresolved_list).
    enriched_rows is a new list of copies (originals are not mutated).
    unresolved_list contains (search_key, field, unit) for fields that could not be filled.
    """
    data_rows = [dict(r) for r in rows]
    field_cfgs    = ENRICHABLE_FIELDS.get(schema, [])
    emission_cfgs = [fc for fc in field_cfgs if fc["provider_group"] == "emission"]
    power_cfgs    = [fc for fc in field_cfgs if fc["provider_group"] == "power"]
    lifetime_cfgs = [fc for fc in field_cfgs if fc["provider_group"] == "lifetime"]

    all_fills = {}

    if cell_site_lookup:
        cs_filled = _enrich_from_cell_site(data_rows, cell_site_lookup)
        if cs_filled:
            all_fills["cell_site"] = cs_filled

    for provider in LIFETIME_HIERARCHY:
        log = []
        for row in data_rows:
            enrich_row(row, schema, provider, lifetime_cfgs, ef_table=ef_table, filled_log=log)
        if log:
            all_fills[provider] = log

    for provider in EMISSION_HIERARCHY:
        log = []
        for row in data_rows:
            enrich_row(row, schema, provider, emission_cfgs, ef_table=ef_table, filled_log=log)
        if log:
            all_fills[provider] = log

    for provider in POWER_HIERARCHY:
        log = []
        for row in data_rows:
            enrich_row(row, schema, provider, power_cfgs, ef_table=ef_table, filled_log=log)
        if log:
            all_fills[provider] = log

    # Collect unresolved fields
    unresolved_seen = set()
    unresolved = []
    for row in data_rows:
        for fc in field_cfgs:
            if fc.get("provider_group") == "lifetime":
                continue
            fname = fc["field"]
            v = row.get(fname)
            if v is not None and str(v).strip():
                continue
            if not precondition_met(row, fc):
                continue
            search_keys = get_search_keys(row, schema, fname)
            if not search_keys:
                continue
            sk         = search_keys[0]
            unit_field = fc.get("unit_field")
            exp_unit   = (row.get(unit_field) if unit_field else None) or fc.get("default_unit") or ""
            key        = (sk, fname)
            if key not in unresolved_seen:
                unresolved_seen.add(key)
                unresolved.append((sk, fname, exp_unit))

    data_rows, _ = compute_group_uncertainty(data_rows, schema)

    summary = {
        src: len(log) if isinstance(log, list) else log
        for src, log in all_fills.items()
    }
    return data_rows, summary, unresolved
