"""
validate.py — Browser edition.

Pure in-memory validation: no telecom_api server, no HTTP calls, no disk I/O.
Callers open workbooks from BytesIO and pass them to load_sheet().
build_cell_site_lookup() replaces the server-side fetch_cell_site_lookup().
"""

from dataclasses import dataclass, field as dc_field

import openpyxl

# ── Enum definitions (mirrors swagger.yaml) ────────────────────────────────────

SITE_TYPES            = ["macro","micro","pico","femto","rooftop","indoor","greenfield","small","medium","large"]
NETWORK_TYPES         = ["Access","metro","aggregation","Backbone","core"]

_NETWORK_TYPE_NORM: dict = {
    "metro":    "metro",
    "Metro":    "metro",
    "access":   "Access",
    "Access":   "Access",
    "backbone": "Backbone",
    "Backbone": "Backbone",
    "backcobe": "Backbone",
    "#N/A":     "unknown",
    "NA":       "unknown",
}

def normalise_network_type(nt) -> str:
    s = str(nt or "").strip()
    if not s or s in ("#N/A", "NA"):
        return "unknown"
    return _NETWORK_TYPE_NORM.get(s, s)

OPERATIONAL_STATUSES  = ["active","planned","under_construction","decommissioned","maintenance"]
ELECTRICITY_SOURCES   = ["grid","renewable","solar","wind","none"]
FUEL_TYPES            = ["bensin","diesel","alternative","none","other"]
REFRIGERANT_TYPES     = ["none","R-11","R-12","R-22","R-123","R-134a","R-410A",
                         "R-245fa","R-32","R-1233zd(E)","R-1234yf","R-513A","other"]
ACTIVE_TYPES          = ["generator","cooling","fire_suppression",
                         "network_equipment","radio_equipment","power_equipment"]
ACTIVE_SUBTYPES       = [
    "standby_power_generator","prime_power_generator","portable_industrial_generator",
    "inverter_generator","container_sized_generator",
    "water_cooled_systems","air_cooled_systems","industrial_chillers",
    "evaporative_cooling_systems","hybrid_system","specialized_cooling",
    "gas_suppression","sprinkler_system","foam_system",
    "aggregation_router","chassis","edge_platform","firewall","gateway",
    "isam","memory_card","router","switch","sfp",
    "air","cellular_modem","radio_unit","baseband_unit","base_station","antenna","wls",
    "sla_vrla","bci","din_en","dc_holder","power_distribution","light","camera","sensor",
]
POWER_SOURCES         = ["battery","electricity","fuel","refrigerant","other"]
PASSIVE_TYPES         = ["fiber_cable","electrical_cables","COAX","splitters","shelters",
                         "cabinets","plugs","fencing","steel","aluminum","plastic"]
INFRA_TYPES           = ["tower","mast","rooftop_mount","pole","building","underground","container",
                         "real estate","manhole","concrete","ducts & pipes"]
TECHNOLOGY_TYPES      = ["2G","3G","4G","5G","6G","ADSL","DSL","FTTH","FTTC","GPON","HFC",
                         "LoRa","MPLS","MEC","Network slicing","PON","POTS/PSTN","RFID",
                         "SD-WAN","VoIP","WLAN","XPON"]
INSTALL_METHODS       = ["traditional_excavation_green","microtrenching_green","milling_green",
                         "plowing_green","traditional_excavation_urban","microtrenching_urban",
                         "milling_urban","plowing_urban","truck","van","ship","airplane","crane","other"]
MAINT_METHODS         = ["physical_security_services","decommissioning","replacements",
                         "environmental_monitoring","failure_monitoring","software_updates","other"]
INSTALL_UNITS         = ["km","kW","kWh","L","m3","m2"]
POWER_UNITS           = ["kW","W","kWh","L","m3"]
INSTALL_EMIT_UNITS    = ["kgCO2eq/km","kgCO2eq/kW","kgCO2eq/kWh","kgCO2eq/L","kgCO2eq/m3","kgCO2eq/m2"]
POWER_EMIT_UNITS      = ["kgCO2eq/kW","kgCO2eq/W","kgCO2eq/kWh","kgCO2eq/L","kgCO2eq/m3"]
EMISSION_UNIT         = ["kgCO2eq/unit"]

MEASURED_ELECTRICITY_UNITS = ["kWh"]
MEASURED_FUEL_UNITS        = ["L", "m3", "kg"]
MEASURED_REFRIGERANT_UNITS = ["kg", "m3"]
ELECTRICITY_EF_UNITS       = ["kgCO2eq/kWh"]
FUEL_EF_UNITS              = ["kgCO2eq/L", "kgCO2eq/m3", "kgCO2eq/kg"]
REFRIGERANT_EF_UNITS       = ["kgCO2eq/m3", "kgCO2eq/kg"]

_MEASURED_EF_UNIT_COMPAT = [
    ("measured_electricity_unit", "electricity_emission_factor_unit",
     {"kWh": "kgCO2eq/kWh"}),
    ("measured_fuel_unit",        "fuel_emission_factor_unit",
     {"L": "kgCO2eq/L", "m3": "kgCO2eq/m3", "kg": "kgCO2eq/kg"}),
    ("measured_refrigerant_unit", "refrigerant_emission_factor_unit",
     {"kg": "kgCO2eq/kg", "m3": "kgCO2eq/m3"}),
]

# ── Data classes ───────────────────────────────────────────────────────────────

@dataclass
class Issue:
    level:   str
    field:   str
    message: str

@dataclass
class RowResult:
    row_num:           int
    status:            str
    issues:            list  = dc_field(default_factory=list)
    quality:           int   = 100
    api_ids:           dict  = dc_field(default_factory=dict)
    enrichment_fields: list  = dc_field(default_factory=list)

# ── Helpers ────────────────────────────────────────────────────────────────────

def present(v):
    return v is not None and str(v).strip() != ""

def to_num(v):
    try:    return float(v)
    except: return None

def compute_quality(issues):
    return max(0, 100 - 10 * sum(1 for i in issues if i.level == "WARNING"))

def _cs_id_key(raw):
    n = to_num(raw)
    return int(n) if n is not None else str(raw).strip()

# ── Generic check functions ────────────────────────────────────────────────────

def chk_mandatory(row, fields):
    return [Issue("ERROR", f, f"Required field '{f}' is missing")
            for f in fields if not present(row.get(f))]

def chk_enums(row, enum_map):
    issues = []
    for fname, allowed in enum_map.items():
        v = row.get(fname)
        if present(v) and str(v) not in allowed:
            issues.append(Issue("WARNING", fname, f"'{v}' is not a valid value for '{fname}'"))
    return issues

def chk_num_unit_pairs(row, pairs):
    issues = []
    for nf, uf in pairs:
        has_n = present(row.get(nf))
        has_u = present(row.get(uf))
        if has_n and not has_u:
            issues.append(Issue("ERROR", uf, f"'{uf}' is required when '{nf}' is provided"))
        elif has_u and not has_n:
            issues.append(Issue("INFO", nf, f"'{uf}' is provided but '{nf}' is empty"))
    return issues

def chk_positive(row, fields):
    issues = []
    for f in fields:
        v = row.get(f)
        if present(v):
            n = to_num(v)
            if n is None:
                issues.append(Issue("ERROR", f, f"'{f}' must be a number, got '{v}'"))
            elif n < 0:
                issues.append(Issue("ERROR", f, f"'{f}' must be >= 0, got {n}"))
    return issues

def chk_life_time(row):
    v = row.get("life_time")
    if not present(v):
        return [Issue("WARNING", "life_time", "life_time is missing — will be filled by enrichment")]
    n = to_num(v)
    if n is None:
        return [Issue("ERROR", "life_time", f"life_time must be a number, got '{v}'")]
    if not (0 <= n <= 200):
        return [Issue("ERROR", "life_time", f"life_time must be between 0 and 200, got {n}")]
    if n < 1:
        return [Issue("WARNING", "life_time", f"life_time is unusually short ({n} years)")]
    return []

def chk_power_fields(row):
    has_qty  = present(row.get("power_quantity"))
    has_idle = present(row.get("power_idle"))
    has_max  = present(row.get("power_max"))
    if not has_qty and not (has_idle and has_max):
        return [Issue("WARNING", "power_quantity",
                      "Active component should have either power_quantity or both power_idle and power_max")]
    return []

def chk_install_maint_complete(row):
    keys = ["installation_method","installation_quantity",
            "maintenance_method","maintenance_quantity"]
    if all(not present(row.get(k)) for k in keys):
        return [Issue("WARNING", "installation_method",
                      "All installation and maintenance fields are empty")]
    return []

def chk_embodied_endoflife(row):
    he = present(row.get("production_emissions"))
    hl = present(row.get("endoflife_emissions"))
    if he and not hl:
        return [Issue("INFO", "endoflife_emissions",
                      "production_emissions provided but endoflife_emissions is missing")]
    if hl and not he:
        return [Issue("INFO", "production_emissions",
                      "endoflife_emissions provided but production_emissions is missing")]
    return []

def chk_ef_without_qty(row):
    issues = []
    for ef, qty in [("installation_emission_factor", "installation_quantity"),
                    ("maintenance_emission_factor",  "maintenance_quantity")]:
        if present(row.get(ef)) and not present(row.get(qty)):
            issues.append(Issue("INFO", qty, f"'{ef}' provided but '{qty}' is missing"))
    return issues

# ── Shared enum/pair config ────────────────────────────────────────────────────

_COMMON_ENUM = {
    "installation_method":              INSTALL_METHODS,
    "installation_unit":                INSTALL_UNITS,
    "installation_emission_factor_unit":INSTALL_EMIT_UNITS,
    "maintenance_method":               MAINT_METHODS,
    "maintenance_unit":                 INSTALL_UNITS,
    "maintenance_emission_factor_unit": INSTALL_EMIT_UNITS,
    "production_emissions_unit":        EMISSION_UNIT,
    "endoflife_emissions_unit":         EMISSION_UNIT,
}

_COMMON_PAIRS = [
    ("installation_quantity",        "installation_unit"),
    ("installation_emission_factor", "installation_emission_factor_unit"),
    ("maintenance_quantity",         "maintenance_unit"),
    ("maintenance_emission_factor",  "maintenance_emission_factor_unit"),
    ("production_emissions",         "production_emissions_unit"),
    ("endoflife_emissions",          "endoflife_emissions_unit"),
]

# ── Schema validators ──────────────────────────────────────────────────────────

def validate_cell_site(row):
    row = dict(row)
    if "network_type" in row:
        row["network_type"] = normalise_network_type(row["network_type"])
    issues = (
        chk_mandatory(row, ["site_type", "network_type", "country"]) +
        chk_enums(row, {
            "site_type":                        SITE_TYPES,
            "network_type":                     NETWORK_TYPES,
            "operational_status":               OPERATIONAL_STATUSES,
            "electricity_source":               ELECTRICITY_SOURCES,
            "fuel_type":                        FUEL_TYPES,
            "refrigerant_type":                 REFRIGERANT_TYPES,
            "measured_electricity_unit":        MEASURED_ELECTRICITY_UNITS,
            "measured_fuel_unit":               MEASURED_FUEL_UNITS,
            "measured_refrigerant_unit":        MEASURED_REFRIGERANT_UNITS,
            "electricity_emission_factor_unit": ELECTRICITY_EF_UNITS,
            "fuel_emission_factor_unit":        FUEL_EF_UNITS,
            "refrigerant_emission_factor_unit": REFRIGERANT_EF_UNITS,
        })
    )
    pr = to_num(row.get("per_rented"))
    if pr is not None and not (0 <= pr <= 100):
        issues.append(Issue("ERROR", "per_rented",
                            f"per_rented must be between 0 and 100, got {pr}"))
    for meas_unit_field, ef_unit_field, compat in _MEASURED_EF_UNIT_COMPAT:
        meas_unit = str(row.get(meas_unit_field) or "").strip()
        ef_unit   = str(row.get(ef_unit_field)   or "").strip()
        if meas_unit and ef_unit:
            expected = compat.get(meas_unit)
            if expected and ef_unit != expected:
                issues.append(Issue("ERROR", ef_unit_field,
                    f"{ef_unit_field} '{ef_unit}' is inconsistent with "
                    f"{meas_unit_field} '{meas_unit}' — expected '{expected}'"))
    return issues

def validate_active_component(row, cell_site_lookup):
    issues = (
        chk_mandatory(row, ["power_source"]) +
        chk_enums(row, {
            **_COMMON_ENUM,
            "active_type":                      ACTIVE_TYPES,
            "active_subtype":                   ACTIVE_SUBTYPES,
            "technology_type":                  TECHNOLOGY_TYPES,
            "power_source":                     POWER_SOURCES,
            "power_unit":                       POWER_UNITS,
            "power_idle_unit":                  POWER_UNITS,
            "power_max_unit":                   POWER_UNITS,
            "power_source_emission_factor_unit":POWER_EMIT_UNITS,
        }) +
        chk_life_time(row) +
        chk_positive(row, ["power_quantity", "power_idle", "power_max",
                            "production_emissions", "endoflife_emissions"]) +
        chk_num_unit_pairs(row, _COMMON_PAIRS + [
            ("power_quantity",              "power_unit"),
            ("power_idle",                  "power_idle_unit"),
            ("power_max",                   "power_max_unit"),
            ("power_source_emission_factor","power_source_emission_factor_unit"),
        ]) +
        chk_power_fields(row) +
        chk_install_maint_complete(row) +
        chk_embodied_endoflife(row) +
        chk_ef_without_qty(row)
    )

    if not present(row.get("active_type")) and not present(row.get("active_subtype")):
        issues.append(Issue("WARNING", "active_type",
                            "Neither active_type nor active_subtype is provided"))

    idle = to_num(row.get("power_idle"))
    pmax = to_num(row.get("power_max"))
    if idle is not None and pmax is not None and pmax < idle:
        issues.append(Issue("ERROR", "power_max",
                            f"power_max ({pmax}) must be >= power_idle ({idle})"))

    cs_id_raw = row.get("cell_site_id")
    if present(cs_id_raw) and str(row.get("power_source", "")).strip() == "fuel":
        site = cell_site_lookup.get(_cs_id_key(cs_id_raw))
        if site and (not present(site.get("fuel_type")) or site.get("fuel_type") == "none"):
            issues.append(Issue("WARNING", "power_source",
                                "power_source is 'fuel' but the linked cell site has no fuel_type"))

    return issues

def validate_passive_component(row):
    return (
        chk_mandatory(row, ["passive_type"]) +
        chk_enums(row, {
            **_COMMON_ENUM,
            "passive_type":    PASSIVE_TYPES,
            "technology_type": TECHNOLOGY_TYPES,
        }) +
        chk_life_time(row) +
        chk_positive(row, ["production_emissions", "endoflife_emissions"]) +
        chk_num_unit_pairs(row, _COMMON_PAIRS) +
        chk_install_maint_complete(row) +
        chk_embodied_endoflife(row) +
        chk_ef_without_qty(row)
    )

def validate_infrastructure(row):
    row = dict(row)
    if "network_type" in row:
        row["network_type"] = normalise_network_type(row["network_type"])
    return (
        chk_mandatory(row, ["infrastructure_type", "network_type"]) +
        chk_enums(row, {
            **_COMMON_ENUM,
            "infrastructure_type": INFRA_TYPES,
            "network_type":        NETWORK_TYPES,
        }) +
        chk_life_time(row) +
        chk_positive(row, ["production_emissions", "endoflife_emissions"]) +
        chk_num_unit_pairs(row, _COMMON_PAIRS) +
        chk_install_maint_complete(row) +
        chk_embodied_endoflife(row) +
        chk_ef_without_qty(row)
    )

# ── Workbook helpers ───────────────────────────────────────────────────────────

def load_sheet(wb: openpyxl.Workbook):
    """Return (headers, data_rows) from the 'Data' sheet of an already-opened workbook."""
    if "Data" not in wb.sheetnames:
        return None, None
    ws = wb["Data"]
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < 2:
        return None, None
    headers   = list(all_rows[0])
    data_rows = [r for r in all_rows[1:]
                 if any(v is not None and str(v).strip() for v in r)]
    return headers, data_rows


def to_dict(headers, row):
    """Convert a flat tuple row to a dict; normalise cell_site_id to str."""
    d = {}
    for h, v in zip(headers, row):
        if h == "cell_site_id" and v is not None:
            v = str(int(v)) if isinstance(v, float) and v == int(v) else str(v)
        d[h] = v.strip() if isinstance(v, str) else v
    return d


def build_cell_site_lookup(cs_rows: list) -> dict:
    """Build {cell_site_id_str: row_dict} from a list of validated cell site row dicts."""
    lookup = {}
    for row in cs_rows:
        cs_id = str(row.get("cell_site_id", "") or "").strip()
        if cs_id:
            lookup[cs_id] = row
    return lookup


def validate_all(cs_wb, active_wb, passive_wb, infra_wb) -> tuple:
    """
    Validate all four sheets in-memory.

    Returns:
        (cell_site_rows, active_rows, passive_rows, infra_rows, summary)
        where *_rows are list[dict] for OK/WARNED rows only,
        and summary is a dict suitable for display.
    """
    summary = {}

    # ── Cell sites ─────────────────────────────────────────────────────────────
    cs_headers, cs_raw = load_sheet(cs_wb)
    cs_dicts, cs_results = [], []
    if cs_headers is not None:
        for i, row in enumerate(cs_raw, 1):
            rd     = to_dict(cs_headers, row)
            issues = validate_cell_site(rd)
            errors = [x for x in issues if x.level == "ERROR"]
            warns  = [x for x in issues if x.level == "WARNING"]
            status = "REJECTED" if errors else ("WARNED" if warns else "OK")
            quality = compute_quality(issues)
            rr = RowResult(i, status, issues, quality,
                           {"cell_site_id": rd.get("cell_site_id")},
                           [x.field for x in warns])
            cs_results.append(rr)
            if status in ("OK", "WARNED"):
                cs_dicts.append(rd)
    summary["cell_site"] = _count(cs_results)

    cell_site_lookup = build_cell_site_lookup(cs_dicts)
    known_ids = set(cell_site_lookup.keys())

    # ── Active components ──────────────────────────────────────────────────────
    ac_headers, ac_raw = load_sheet(active_wb)
    ac_dicts, ac_results = [], []
    if ac_headers is not None:
        for i, row in enumerate(ac_raw, 1):
            rd     = to_dict(ac_headers, row)
            cs_id  = str(rd.get("cell_site_id", "") or "").strip()
            issues = []
            if cs_id and cs_id not in known_ids:
                issues.append(Issue("WARNING", "cell_site_id",
                                    f"cell_site_id {cs_id} not found — treated as 3rd party"))
            issues += validate_active_component(rd, cell_site_lookup)
            errors  = [x for x in issues if x.level == "ERROR"]
            warns   = [x for x in issues if x.level == "WARNING"]
            status  = "REJECTED" if errors else ("WARNED" if warns else "OK")
            quality = compute_quality(issues)
            rr = RowResult(i, status, issues, quality, {}, [x.field for x in warns])
            ac_results.append(rr)
            if status in ("OK", "WARNED"):
                ac_dicts.append(rd)
    summary["active"] = _count(ac_results)

    # ── Passive components ─────────────────────────────────────────────────────
    pa_headers, pa_raw = load_sheet(passive_wb)
    pa_dicts, pa_results = [], []
    if pa_headers is not None:
        for i, row in enumerate(pa_raw, 1):
            rd     = to_dict(pa_headers, row)
            cs_id  = str(rd.get("cell_site_id", "") or "").strip()
            issues = []
            if cs_id and cs_id not in known_ids:
                issues.append(Issue("WARNING", "cell_site_id",
                                    f"cell_site_id {cs_id} not found — treated as 3rd party"))
            issues += validate_passive_component(rd)
            errors  = [x for x in issues if x.level == "ERROR"]
            warns   = [x for x in issues if x.level == "WARNING"]
            status  = "REJECTED" if errors else ("WARNED" if warns else "OK")
            quality = compute_quality(issues)
            rr = RowResult(i, status, issues, quality, {}, [x.field for x in warns])
            pa_results.append(rr)
            if status in ("OK", "WARNED"):
                pa_dicts.append(rd)
    summary["passive"] = _count(pa_results)

    # ── Infrastructure ─────────────────────────────────────────────────────────
    in_headers, in_raw = load_sheet(infra_wb)
    in_dicts, in_results = [], []
    if in_headers is not None:
        for i, row in enumerate(in_raw, 1):
            rd     = to_dict(in_headers, row)
            cs_id  = str(rd.get("cell_site_id", "") or "").strip()
            issues = []
            if cs_id and cs_id not in known_ids:
                issues.append(Issue("WARNING", "cell_site_id",
                                    f"cell_site_id {cs_id} not found — treated as 3rd party"))
            issues += validate_infrastructure(rd)
            errors  = [x for x in issues if x.level == "ERROR"]
            warns   = [x for x in issues if x.level == "WARNING"]
            status  = "REJECTED" if errors else ("WARNED" if warns else "OK")
            quality = compute_quality(issues)
            rr = RowResult(i, status, issues, quality, {}, [x.field for x in warns])
            in_results.append(rr)
            if status in ("OK", "WARNED"):
                in_dicts.append(rd)
    summary["infrastructure"] = _count(in_results)

    return cs_dicts, ac_dicts, pa_dicts, in_dicts, cell_site_lookup, summary


def _count(results: list) -> dict:
    total    = len(results)
    ok       = sum(1 for r in results if r.status == "OK")
    warned   = sum(1 for r in results if r.status == "WARNED")
    rejected = sum(1 for r in results if r.status == "REJECTED")
    return {"total": total, "ok": ok, "warned": warned, "rejected": rejected}
