"""
model_operational.py — Operational emissions model.

Computes annualised operational emissions for a single active component row.
None in the output = input was not provided or not applicable; 0.0 = explicitly zero.
Flags collect any issues found during computation.

Power source hierarchy (applied independently per energy source type):
  1. site_measured  — measured_electricity / measured_fuel / measured_refrigerant
                      from cell_site (highest confidence); op_energy computed at site
                      level in step8, not per component
  2. quantity       — power_quantity from active component (direct consumption)
  3. estimated      — power_idle + 0.8 × (power_max − power_idle) × 8760 / 1000
                      (electricity only)

Unit conversion applied to power_quantity and power_idle/max:
  W   → kWh/yr  : × 8760 / 1000
  kW  → kWh/yr  : × 8760
  kWh → kWh/yr  : no conversion (assumed already annual)
  L   → L/yr    : no conversion
  m3  → m3/yr   : no conversion

  op_energy           = annual_consumption × power_source_emission_factor  (annual, kgCO2e/yr)
  op_maintenance      = maintenance_quantity × maintenance_emission_factor  (one-time, kgCO2e)
  op_total            = op_energy + op_maintenance
"""

from typing import Optional

HOURS_PER_YEAR = 8760

FIELDS = [
    "op_energy",
    "op_maintenance",
    "op_total",
    "power_estimated_w",
    "annual_consumption",
    "power_path",
    "flags",
]

_ELEC_SOURCES   = {"electricity", "battery"}
_FUEL_SOURCES   = {"fuel"}
_REFRIG_SOURCES = {"refrigerant"}

_MEASURED_MAP = {
    "electricity": "measured_electricity",
    "battery":     "measured_electricity",
    "fuel":        "measured_fuel",
    "refrigerant": "measured_refrigerant",
}

# Maps power_unit value → (multiplier to get annual kWh, canonical unit label)
_POWER_UNIT_CONVERSION = {
    "w":   (HOURS_PER_YEAR / 1000.0, "kWh/yr"),
    "kw":  (HOURS_PER_YEAR,          "kWh/yr"),
    "kwh": (1.0,                     "kWh/yr"),
    "l":   (1.0,                     "L/yr"),
    "m3":  (1.0,                     "m3/yr"),
}


def _f(v) -> Optional[float]:
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _convert_power(value: float, unit: str) -> tuple:
    """Return (converted_value, flag_str_or_None).
    Converts power/consumption value to annual units based on unit string.
    """
    u = unit.strip().lower() if unit else ""
    conversion = _POWER_UNIT_CONVERSION.get(u)
    if conversion:
        multiplier, _ = conversion
        return value * multiplier, None
    if u:
        return value, f"power_unit '{unit}' not recognised — assuming quantity is already in EF denominator unit"
    return value, None


def _convert_power_w(value: float, unit: str) -> float:
    """Convert a power value (idle or max) to watts."""
    u = unit.strip().lower() if unit else ""
    if u == "kw":
        return value * 1000.0
    return value  # assume W or empty → already watts


def _get_ef(row: dict, ef_table: dict) -> Optional[float]:
    """Return power_source_emission_factor for this row from the row itself or ef_table."""
    v = _f(row.get("power_source_emission_factor"))
    if v is not None:
        return v
    ps = str(row.get("power_source", "") or "").strip()
    for key in ([ps, f"M-{ps}"] if ps else []) + ["power_source_emission_factor"]:
        result = ef_table.get((key, "power_source_emission_factor")) or ef_table.get((key, ""))
        if result:
            v = _f(result.get("value"))
            if v is not None:
                return v
    return None


def compute(row: dict, schema: str, ef_table: dict, site: dict = None) -> dict:
    """Return operational emission figures for one row.

    None values indicate missing or not-applicable data.
    'flags' is a semicolon-separated string of issues found.
    op_maintenance is a one-time cost (not annualised).
    When power_path = 'site_measured', op_energy is None — site total is
    computed in step8 and must not be summed from individual components.
    """
    # ── Schema guard ──────────────────────────────────────────────────────────
    if schema != "active":
        return {
            "op_energy":         None,
            "op_maintenance":    None,
            "op_total":          None,
            "power_estimated_w": None,
            "annual_consumption": None,
            "power_path":        "not_applicable",
            "flags":             "operational emissions not applicable for this component type",
        }

    flags            = []
    op_energy        = None
    annual_consumption = None
    power_estimated  = None
    power_path       = None
    site             = site or {}

    # ── Power source and consumption hierarchy ────────────────────────────────
    power_source   = str(row.get("power_source", "") or "").strip().lower()
    measured_field = _MEASURED_MAP.get(power_source)
    measured_qty   = _f(site.get(measured_field)) if measured_field else None

    if measured_qty is not None:
        # Level 1 — site-level measured value
        power_path         = "site_measured"
        annual_consumption = measured_qty
        flags.append("op_energy: calculated at site level from measured consumption")
        op_energy = None  # step8 computes site total

    elif _f(row.get("power_quantity")) is not None:
        # Level 2 — component direct quantity
        power_path  = "quantity"
        raw_qty     = _f(row.get("power_quantity"))
        pwr_unit    = str(row.get("power_unit", "") or "")
        converted, unit_flag = _convert_power(raw_qty, pwr_unit)
        annual_consumption = converted
        if unit_flag:
            flags.append(unit_flag)
        ef = _get_ef(row, ef_table)
        if ef is None:
            flags.append("power_source_emission_factor: not provided")
        else:
            op_energy = round(annual_consumption * ef, 4)

    elif power_source in _ELEC_SOURCES:
        # Level 3 — estimate from idle/max (electricity only)
        power_path = "estimated"
        p_idle     = _f(row.get("power_idle"))
        p_max      = _f(row.get("power_max"))
        if p_idle is None:
            flags.append("power_idle: not provided")
        if p_max is None:
            flags.append("power_max: not provided")

        if p_idle is not None and p_max is not None:
            idle_w = _convert_power_w(p_idle, str(row.get("power_idle_unit", "") or ""))
            max_w  = _convert_power_w(p_max,  str(row.get("power_max_unit",  "") or ""))
            if max_w < idle_w:
                flags.append(f"power_max < power_idle — estimated power may be unreliable "
                             f"(idle={idle_w:.1f} W, max={max_w:.1f} W)")
            pe_w            = idle_w + 0.8 * (max_w - idle_w)
            power_estimated = pe_w
            annual_consumption = pe_w * HOURS_PER_YEAR / 1000.0
            ef = _get_ef(row, ef_table)
            if ef is None:
                flags.append("power_source_emission_factor: not provided")
            else:
                op_energy = round(annual_consumption * ef, 4)
        else:
            flags.append("power_quantity: not provided — could not estimate from idle/max")

    else:
        flags.append("power_quantity: not provided")
        if not power_source:
            flags.append("power_source: not provided")

    # Top-level flag when no operational energy can be calculated
    if op_energy is None and power_path != "site_measured":
        flags.append("op_energy: could not be calculated — no usable power data")

    # ── Maintenance (one-time cost, never divided by lifetime) ────────────────
    maint_qty = _f(row.get("maintenance_quantity"))
    maint_ef  = _f(row.get("maintenance_emission_factor"))

    if maint_qty is None:
        flags.append("maintenance_quantity: not provided")
    if maint_ef is None:
        flags.append("maintenance_emission_factor: not provided")

    if maint_qty is not None and maint_ef is not None:
        op_maintenance = round(maint_qty * maint_ef, 4)
    else:
        op_maintenance = None

    # ── op_total ──────────────────────────────────────────────────────────────
    available = [x for x in (op_energy, op_maintenance) if x is not None]
    if available:
        op_total = round(sum(available), 4)
        if len(available) < 2:
            flags.append("op_total: partial — some components not provided")
    else:
        op_total = None

    return {
        "op_energy":          op_energy,
        "op_maintenance":     op_maintenance,
        "op_total":           op_total,
        "power_estimated_w":  round(power_estimated, 2) if power_estimated is not None else None,
        "annual_consumption": round(annual_consumption, 2) if annual_consumption is not None else None,
        "power_path":         power_path,
        "flags":              "; ".join(flags),
    }
