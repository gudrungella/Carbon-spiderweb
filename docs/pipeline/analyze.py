"""
analyze.py — Browser edition.

Same logic as step8_analyze.py but runs entirely in-memory.
No disk I/O; data is passed in as pre-enriched list[dict].
Entry point: analyze_in_memory().
"""

import math
from collections import defaultdict

import model_embodied
import model_operational
from validate import normalise_network_type


def _f(v):
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


# ── Site-level modifiers ───────────────────────────────────────────────────────

_EMB_EMISSION_FIELDS = frozenset({
    "production_emissions_annual", "install_total", "install_total_annual",
    "cradle_to_site_annual", "eol_emissions_annual", "embodied_emissions_annual",
})
_OP_EMISSION_FIELDS = frozenset({
    "op_energy", "op_maintenance", "op_total",
})


def _site_scale_factor(site: dict):
    status = str(site.get("operational_status", "") or "").strip().lower()
    if status == "decommissioned":
        return True, 0.0
    raw = site.get("per_rented")
    if raw is not None and str(raw).strip() != "":
        try:
            per_rented = float(raw)
            if 0 < per_rented <= 100:
                return False, round(1.0 - per_rented / 100.0, 6)
        except (ValueError, TypeError):
            pass
    return False, 1.0


def _apply_site_modifiers(emb: dict, op: dict, site: dict) -> tuple:
    _, factor = _site_scale_factor(site)
    if factor == 1.0:
        return emb, op

    def _scale(d, fields):
        result = dict(d)
        for f in fields:
            v = result.get(f)
            if v is not None:
                result[f] = round(v * factor, 4)
        return result

    adj_emb = _scale(emb, _EMB_EMISSION_FIELDS)
    adj_op  = _scale(op,  _OP_EMISSION_FIELDS)
    is_decom = (factor == 0.0)
    if is_decom:
        note = "site decommissioned — emissions zeroed"
    else:
        note = f"per_rented={round((1.0-factor)*100,2)}% — emissions scaled by {factor}"
    for d in (adj_emb, adj_op):
        existing = d.get("flags", "") or ""
        d["flags"] = f"{existing}; {note}".lstrip("; ")
    return adj_emb, adj_op


# ── Aggregation ────────────────────────────────────────────────────────────────

def _aggregate(results: list, cs_lookup: dict, site_op: dict) -> tuple:
    by_schema  = defaultdict(lambda: {"emb": 0.0, "op": 0.0, "eol": 0.0, "t": 0.0, "n": 0})
    by_network = defaultdict(lambda: {"emb": 0.0, "op": 0.0, "eol": 0.0, "t": 0.0})
    by_site    = defaultdict(lambda: {"emb": 0.0, "op": 0.0, "eol": 0.0, "t": 0.0, "nt": ""})

    for schema, row, emb, op in results:
        cs_id   = str(row.get("cell_site_id", "")).strip()
        site    = cs_lookup.get(cs_id, {})
        _dnt    = "3rd party site" if cs_id not in cs_lookup else "unknown"
        _raw_nt = str(site.get("network_type") or row.get("network_type") or _dnt).strip()
        nt      = normalise_network_type(_raw_nt) if _raw_nt not in ("3rd party site", "unknown") else _raw_nt
        emb_val = emb["cradle_to_site_annual"] or 0.0
        eol_val = emb["eol_emissions_annual"] or 0.0
        op_val  = (op["op_total"] or 0.0) if op["power_path"] != "site_measured" else 0.0
        total   = emb_val + op_val + eol_val
        for agg in (by_schema[schema], by_network[nt], by_site[cs_id]):
            agg["emb"] += emb_val
            agg["op"]  += op_val
            agg["eol"] += eol_val
            agg["t"]   += total
        by_schema[schema]["n"] += 1
        by_site[cs_id]["nt"] = nt

    for cs_id, sources in site_op.items():
        site    = cs_lookup.get(cs_id, {})
        _raw_nt = str(site.get("network_type") or "unknown").strip()
        nt      = normalise_network_type(_raw_nt) if _raw_nt != "unknown" else "unknown"
        for source_data in sources.values():
            op_val = source_data["op_energy"] or 0.0
            for agg in (by_network[nt], by_site[cs_id]):
                agg["op"] += op_val
                agg["t"]  += op_val

    return by_schema, by_network, by_site


# ── Site-level op emissions from measured consumption ─────────────────────────

def _build_site_op(cs_lookup: dict, ef_table: dict) -> dict:
    site_op = {}
    _sources = [
        ("measured_electricity", "electricity_emission_factor", "electricity_source"),
        ("measured_fuel",        "fuel_emission_factor",        "fuel_type"),
        ("measured_refrigerant", "refrigerant_emission_factor", "refrigerant_type"),
    ]
    for cs_id, site in cs_lookup.items():
        for measured_field, ef_field, source_type_field in _sources:
            qty = _f(site.get(measured_field))
            if qty is None:
                continue
            ef = _f(site.get(ef_field))
            if ef is None:
                source = str(site.get(source_type_field, "") or "").strip()
                for key in ([source, f"M-{source}"] if source else []):
                    result = (ef_table.get((key, "power_source_emission_factor"))
                              or ef_table.get((key, "")))
                    if result:
                        ef = _f(result.get("value"))
                        if ef is not None:
                            break
            op_energy = round(qty * ef, 4) if ef is not None else None
            site_op.setdefault(cs_id, {})[measured_field] = {
                "qty": qty, "ef": ef, "op_energy": op_energy,
            }

    for cs_id, sources in site_op.items():
        site = cs_lookup.get(cs_id, {})
        _, factor = _site_scale_factor(site)
        for source_data in sources.values():
            if source_data["op_energy"] is not None:
                source_data["op_energy"] = round(source_data["op_energy"] * factor, 4)

    return site_op


# ── Sensitivity analysis ───────────────────────────────────────────────────────

def _perturb_row(row: dict, field: str, factor: float) -> dict:
    v = row.get(field)
    if v is None:
        return row
    try:
        fv = float(v)
    except (ValueError, TypeError):
        return row
    result = dict(row)
    result[field] = round(fv * factor, 6)
    return result


def _compute_grand_total(cell_sites, active, passive, infra, ef_table) -> float:
    cs_lookup = {str(s.get("cell_site_id", "")).strip(): s for s in cell_sites}
    site_op   = _build_site_op(cs_lookup, ef_table)
    _na_op    = {f: None for f in model_operational.FIELDS}
    _na_op["power_path"] = "not_applicable"
    _na_op["flags"]      = "operational emissions not applicable for this component type"
    results = []
    for row in active:
        cs_id = str(row.get("cell_site_id", "")).strip()
        site  = cs_lookup.get(cs_id, {})
        emb   = model_embodied.compute(row)
        op    = model_operational.compute(row, "active", ef_table, site=site)
        emb, op = _apply_site_modifiers(emb, op, site)
        results.append(("active", row, emb, op))
    for row in passive:
        cs_id = str(row.get("cell_site_id", "")).strip()
        site  = cs_lookup.get(cs_id, {})
        emb   = model_embodied.compute(row)
        emb, _ = _apply_site_modifiers(emb, dict(_na_op), site)
        results.append(("passive", row, emb, dict(_na_op)))
    for row in infra:
        cs_id = str(row.get("cell_site_id", "")).strip()
        site  = cs_lookup.get(cs_id, {})
        emb   = model_embodied.compute(row)
        emb, _ = _apply_site_modifiers(emb, dict(_na_op), site)
        results.append(("infrastructure", row, emb, dict(_na_op)))
    by_schema, _, _ = _aggregate(results, cs_lookup, site_op)
    grand_t = sum(v["t"] for v in by_schema.values())
    for sources in site_op.values():
        for source_data in sources.values():
            grand_t += source_data["op_energy"] or 0.0
    return grand_t


def run_sensitivity(cell_sites, active, passive, infra, ef_table) -> list:
    _A_ALL = ["production_emissions", "endoflife_emissions",
              "power_quantity", "power_idle", "power_max", "maintenance_quantity"]
    _P_ALL = ["production_emissions", "endoflife_emissions", "installation_quantity"]
    _I_ALL = ["production_emissions", "endoflife_emissions", "installation_quantity"]
    _PWR   = ["power_quantity", "power_idle", "power_max"]

    _PARAMS = [
        ("Equipment lifetime",         ["life_time"],                      ["life_time"],                      ["life_time"]),
        ("Active component lifetime",  ["life_time"],                      [],                                 []),
        ("Infrastructure lifetime",    [],                                 [],                                 ["life_time"]),
        ("Production emissions EF",    ["production_emissions"],           ["production_emissions"],           ["production_emissions"]),
        ("Installation EF",            ["installation_emission_factor"],   ["installation_emission_factor"],   ["installation_emission_factor"]),
        ("Installation quantity",      ["installation_quantity"],          ["installation_quantity"],          ["installation_quantity"]),
        ("Grid emission factor",       ["power_source_emission_factor"],   [],                                 []),
        ("End-of-life emissions EF",   ["endoflife_emissions"],            ["endoflife_emissions"],            ["endoflife_emissions"]),
        ("Maintenance EF",             ["maintenance_emission_factor"],    [],                                 []),
        ("Maintenance frequency",      ["maintenance_quantity"],           [],                                 []),
        ("Network power draw",         _PWR,                              [],                                 []),
        ("Active component count",     _A_ALL,                            [],                                 []),
        ("Passive component count",    [],                                _P_ALL,                             []),
        ("Infrastructure count",       [],                                [],                                 _I_ALL),
        ("Fleet size",                 _A_ALL,                            _P_ALL,                             _I_ALL),
    ]

    baseline = _compute_grand_total(cell_sites, active, passive, infra, ef_table)
    if not baseline:
        return []

    def _perturb(rows, fields, factor):
        result = []
        for row in rows:
            r = row
            for f in fields:
                r = _perturb_row(r, f, factor)
            result.append(r)
        return result

    results = []
    for label, a_fields, p_fields, i_fields in _PARAMS:
        row_results = {}
        for factor in (0.8, 1.2):
            pa = _perturb(active,  a_fields, factor) if a_fields else active
            pp = _perturb(passive, p_fields, factor) if p_fields else passive
            pi = _perturb(infra,   i_fields, factor) if i_fields else infra
            total = _compute_grand_total(cell_sites, pa, pp, pi, ef_table)
            row_results[factor] = round(100.0 * total / baseline, 2)
        low   = row_results[0.8]
        high  = row_results[1.2]
        swing = round(abs(high - low), 2)
        results.append({"label": label, "low": low, "high": high, "swing": swing})

    results.sort(key=lambda x: x["swing"], reverse=True)
    return results


# ── Uncertainty analysis ───────────────────────────────────────────────────────

def _fsd(v) -> float:
    if v is None:
        return 0.0
    try:
        return max(0.0, float(v))
    except (ValueError, TypeError):
        return 0.0


def _compute_row_uncertainty(row: dict, emb: dict, op: dict) -> dict:
    lifetime = max(emb.get("lifetime_used") or 1.0, 1e-9)

    sd_prod   = _fsd(row.get("production_emissions_sd"))
    sd_ef_ins = _fsd(row.get("installation_emission_factor_sd"))
    inst_qty  = _f(row.get("installation_quantity")) or 0.0

    sigma_prod_annual = sd_prod / lifetime
    sigma_inst_annual = sd_ef_ins * inst_qty / lifetime
    sigma_emb = math.sqrt(sigma_prod_annual ** 2 + sigma_inst_annual ** 2)

    sigma_eol = _fsd(row.get("endoflife_emissions_sd")) / lifetime

    sigma_op_e = 0.0
    power_path = op.get("power_path") or ""

    if power_path == "quantity":
        pq    = _f(row.get("power_quantity"))
        pq_sd = _fsd(row.get("power_quantity_sd"))
        op_e  = _f(op.get("op_energy")) or 0.0
        if pq and pq_sd:
            sigma_op_e = op_e * (pq_sd / pq)
    elif power_path == "estimated":
        low_w  = _f(row.get("power_estimated_low_w"))
        high_w = _f(row.get("power_estimated_high_w"))
        op_e   = _f(op.get("op_energy")) or 0.0
        if low_w is not None and high_w is not None:
            p_idle = _f(row.get("power_idle")) or 0.0
            p_max  = _f(row.get("power_max"))  or 0.0
            if str(row.get("power_idle_unit", "") or "").strip().lower() == "kw":
                p_idle *= 1000.0
            if str(row.get("power_max_unit",  "") or "").strip().lower() == "kw":
                p_max  *= 1000.0
            p_central = p_idle + 0.8 * (p_max - p_idle)
            if p_central > 0:
                op_low_e  = op_e * (low_w  / p_central)
                op_high_e = op_e * (high_w / p_central)
                sigma_op_e = (op_high_e - op_low_e) / 2.0

    sd_maint    = _fsd(row.get("maintenance_emission_factor_sd"))
    maint_qty   = _f(row.get("maintenance_quantity")) or 0.0
    sigma_maint = sd_maint * maint_qty
    sigma_op    = math.sqrt(sigma_op_e ** 2 + sigma_maint ** 2)
    sigma_total = math.sqrt(sigma_emb ** 2 + sigma_eol ** 2 + sigma_op ** 2)

    return {
        "sigma_emb":   round(sigma_emb,   4),
        "sigma_eol":   round(sigma_eol,   4),
        "sigma_op":    round(sigma_op,    4),
        "sigma_total": round(sigma_total, 4),
    }


def compute_uncertainty(results: list, cs_lookup: dict) -> dict:
    def _zero():
        return {"var_emb": 0.0, "var_eol": 0.0, "var_op": 0.0, "var_t": 0.0}

    grand      = _zero()
    by_schema  = defaultdict(_zero)
    by_network = defaultdict(_zero)
    by_site    = defaultdict(_zero)
    sigma2_elec = 0.0

    for schema, row, emb, op in results:
        if op.get("power_path") == "site_measured":
            continue
        cs_id = str(row.get("cell_site_id", "")).strip()
        site  = cs_lookup.get(cs_id, {})
        _raw_nt = str(site.get("network_type") or row.get("network_type") or "unknown").strip()
        nt = normalise_network_type(_raw_nt) if _raw_nt not in ("3rd party site", "unknown") else _raw_nt

        u = _compute_row_uncertainty(row, emb, op)
        for agg in (grand, by_schema[schema], by_network[nt], by_site[cs_id]):
            agg["var_emb"] += u["sigma_emb"]   ** 2
            agg["var_eol"] += u["sigma_eol"]   ** 2
            agg["var_op"]  += u["sigma_op"]    ** 2
            agg["var_t"]   += u["sigma_total"] ** 2

        if schema == "active":
            ps = str(row.get("power_source", "") or "").strip().lower()
            if ps in ("electricity", "battery"):
                power_path = op.get("power_path") or ""
                if power_path == "quantity":
                    pq    = _f(row.get("power_quantity"))
                    pq_sd = _fsd(row.get("power_quantity_sd"))
                    ac    = _f(op.get("annual_consumption")) or 0.0
                    if pq and pq_sd:
                        sigma2_elec += (ac * (pq_sd / pq)) ** 2
                elif power_path == "estimated":
                    low_w  = _f(row.get("power_estimated_low_w"))
                    high_w = _f(row.get("power_estimated_high_w"))
                    if low_w is not None and high_w is not None:
                        sigma2_elec += ((high_w - low_w) / 2 * 8760 / 1000) ** 2

    def _to_sd(d):
        return {k.replace("var_", "") + "_sd": round(math.sqrt(v), 1)
                for k, v in d.items()}

    return {
        "grand":              _to_sd(grand),
        "by_schema":          {k: _to_sd(v) for k, v in by_schema.items()},
        "by_network":         {k: _to_sd(v) for k, v in by_network.items()},
        "by_site":            {k: _to_sd(v) for k, v in by_site.items()},
        "electricity_kwh_sd": round(math.sqrt(sigma2_elec), 1),
    }


def compute_high_variability_groups(active: list, passive: list, infra: list,
                                    top_n: int = 15) -> list:
    _FIELDS = {
        "active":         ["production_emissions", "endoflife_emissions",
                           "installation_emission_factor", "maintenance_emission_factor"],
        "passive":        ["production_emissions", "endoflife_emissions",
                           "installation_emission_factor"],
        "infrastructure": ["production_emissions", "endoflife_emissions",
                           "installation_emission_factor"],
    }
    _TYPE_FIELD = {
        "active": "active_subtype", "passive": "passive_type",
        "infrastructure": "infrastructure_type",
    }
    _FIELD_LABELS = {
        "production_emissions":         "Production emissions",
        "endoflife_emissions":          "End-of-life emissions",
        "installation_emission_factor": "Installation EF",
        "maintenance_emission_factor":  "Maintenance EF",
    }

    entries = []
    for schema, rows in [("active", active), ("passive", passive), ("infrastructure", infra)]:
        tf = _TYPE_FIELD[schema]
        groups: dict = {}
        for row in rows:
            t = str(row.get(tf, "") or "").strip() or "__unknown__"
            groups.setdefault(t, []).append(row)
        for type_name, grp in groups.items():
            n = len(grp)
            if n < 2:
                continue
            for field in _FIELDS[schema]:
                vals = [_f(r.get(field)) for r in grp]
                vals = [v for v in vals if v is not None]
                if not vals:
                    continue
                mean = sum(vals) / len(vals)
                if abs(mean) < 1e-9:
                    continue
                sd = _fsd(grp[0].get(f"{field}_sd"))
                if sd == 0:
                    continue
                entries.append({
                    "schema": schema, "type": type_name,
                    "field": _FIELD_LABELS.get(field, field),
                    "n": n, "mean": round(mean, 2), "sd": round(sd, 2),
                    "relative_sd_pct": round(100.0 * sd / abs(mean), 1),
                })

    entries.sort(key=lambda x: x["relative_sd_pct"], reverse=True)
    return entries[:top_n]


# ── Main in-memory entry point ─────────────────────────────────────────────────

def analyze_in_memory(active_rows: list, passive_rows: list, infra_rows: list,
                      cell_site_rows: list, ef_table: dict = None) -> tuple:
    """
    Run the full analysis pipeline in-memory.

    Returns (summary, op_rows, emb_rows, sensitivity, uncertainty)
    where summary is a dict suitable for visualise.build_data().
    """
    ef_table = ef_table or {}
    cs_lookup = {str(s.get("cell_site_id", "")).strip(): s for s in cell_site_rows}
    site_op   = _build_site_op(cs_lookup, ef_table)

    _na_op = {f: None for f in model_operational.FIELDS}
    _na_op["power_path"] = "not_applicable"
    _na_op["flags"]      = "operational emissions not applicable for this component type"

    results = []
    for row in active_rows:
        cs_id = str(row.get("cell_site_id", "")).strip()
        site  = cs_lookup.get(cs_id, {})
        emb   = model_embodied.compute(row)
        op    = model_operational.compute(row, "active", ef_table, site=site)
        emb, op = _apply_site_modifiers(emb, op, site)
        results.append(("active", row, emb, op))
    for row in passive_rows:
        cs_id = str(row.get("cell_site_id", "")).strip()
        site  = cs_lookup.get(cs_id, {})
        emb   = model_embodied.compute(row)
        emb, _ = _apply_site_modifiers(emb, dict(_na_op), site)
        results.append(("passive", row, emb, dict(_na_op)))
    for row in infra_rows:
        cs_id = str(row.get("cell_site_id", "")).strip()
        site  = cs_lookup.get(cs_id, {})
        emb   = model_embodied.compute(row)
        emb, _ = _apply_site_modifiers(emb, dict(_na_op), site)
        results.append(("infrastructure", row, emb, dict(_na_op)))

    by_schema, by_network, by_site = _aggregate(results, cs_lookup, site_op)

    grand = {k: sum(v[k] for v in by_schema.values()) for k in ("emb", "op", "eol", "t")}
    for sources in site_op.values():
        for source_data in sources.values():
            grand["op"] += source_data["op_energy"] or 0.0
            grand["t"]  += source_data["op_energy"] or 0.0

    # Electricity consumption total
    elec_kwh = 0.0
    for row in active_rows:
        ps = str(row.get("power_source", "") or "").strip().lower()
        if ps in ("electricity", "battery"):
            ac = _f(row.get("annual_consumption"))
            if ac is not None:
                elec_kwh += ac
    for cs_id, sources in site_op.items():
        src_data = sources.get("measured_electricity")
        if src_data:
            elec_kwh += src_data["qty"] or 0.0

    # Format summary in the structure expected by visualise.build_data()
    summary = {
        "By Component Type": [
            {"Type": schema, "Count": v["n"],
             "Embodied (kgCO2e/yr)": round(v["emb"], 2),
             "Operational (kgCO2e/yr)": round(v["op"], 2),
             "End-of-life (kgCO2e/yr)": round(v["eol"], 2),
             "Total (kgCO2e/yr)": round(v["t"], 2)}
            for schema, v in sorted(by_schema.items())
        ],
        "By Network Type": [
            {"Network": nt,
             "Embodied (kgCO2e/yr)": round(v["emb"], 2),
             "Operational (kgCO2e/yr)": round(v["op"], 2),
             "End-of-life (kgCO2e/yr)": round(v["eol"], 2),
             "Total (kgCO2e/yr)": round(v["t"], 2)}
            for nt, v in sorted(by_network.items())
        ],
        "By Cell Site": [
            {"Cell Site": cs_id, "Network": v["nt"],
             "Embodied (kgCO2e/yr)": round(v["emb"], 2),
             "Operational (kgCO2e/yr)": round(v["op"], 2),
             "End-of-life (kgCO2e/yr)": round(v["eol"], 2),
             "Total (kgCO2e/yr)": round(v["t"], 2)}
            for cs_id, v in sorted(by_site.items())
        ],
    }

    emb_rows = [{"schema": s, **r, **e} for s, r, e, _ in results]
    op_rows  = [{"schema": s, **r, **o} for s, r, _, o in results]

    uncertainty = compute_uncertainty(results, cs_lookup)
    uncertainty["high_variability"] = compute_high_variability_groups(
        active_rows, passive_rows, infra_rows)

    sensitivity = run_sensitivity(
        cell_site_rows, active_rows, passive_rows, infra_rows, ef_table)

    return summary, op_rows, emb_rows, sensitivity, uncertainty
