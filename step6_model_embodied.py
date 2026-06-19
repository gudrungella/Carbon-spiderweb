"""
model_embodied.py — Embodied emissions model.

Computes annualised embodied and end-of-life emissions for a single component row.
None in the output = input was not provided; 0.0 = explicitly zero.
Flags collect any issues found during computation.

  production_emissions_annual = production_emissions / lifetime
  install_total               = installation_quantity × installation_emission_factor
  install_total_annual        = install_total / lifetime
  cradle_to_site_annual       = production_emissions_annual + install_total_annual  (partial if either is None)
  eol_emissions_annual        = endoflife_emissions / lifetime
  embodied_emissions_annual   = cradle_to_site_annual + eol_emissions_annual        (partial if either is None)
"""


def _f(v):
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _check_unit_match(qty_unit: str, ef_unit: str, label: str):
    """Return a flag string if the quantity unit doesn't match the EF unit denominator, else None."""
    if not qty_unit or not ef_unit:
        return None
    # Extract denominator from EF unit (e.g. "kgCO2eq/km" → "km")
    parts = ef_unit.split("/", 1)
    if len(parts) != 2:
        return None
    denominator = parts[1].strip().lower()
    if qty_unit.strip().lower() != denominator:
        return (f"{label}: unit mismatch — quantity unit '{qty_unit}' does not match "
                f"EF denominator '{denominator}' (from '{ef_unit}')")
    return None


FIELDS = [
    "production_emissions_annual",
    "install_total",
    "install_total_annual",
    "cradle_to_site_annual",
    "eol_emissions_annual",
    "embodied_emissions_annual",
    "lifetime_used",
    "flags",
]


def compute(row: dict) -> dict:
    """Return annualised embodied and end-of-life figures for one row.

    None values indicate missing input data.
    'flags' is a semicolon-separated string of issues found during computation.
    Totals are computed from available components and flagged as partial if any input is missing.
    """
    flags = []

    # ── Lifetime ──────────────────────────────────────────────────────────────
    life_time_raw = _f(row.get("life_time"))
    if life_time_raw is None or life_time_raw == 0.0:
        life_time = 1.0
        flags.append("life_time not provided — defaulted to 1.0 year")
    else:
        life_time = life_time_raw

    # ── Check whether ALL primary inputs are absent ───────────────────────────
    prod_raw  = _f(row.get("production_emissions"))
    inst_qty  = _f(row.get("installation_quantity"))
    inst_ef   = _f(row.get("installation_emission_factor"))
    eol_raw   = _f(row.get("endoflife_emissions"))

    all_absent = (prod_raw is None and inst_qty is None
                  and inst_ef is None and eol_raw is None)

    if all_absent:
        flags.append("no embodied input data provided")
        return {
            "production_emissions_annual": None,
            "install_total":               None,
            "install_total_annual":        None,
            "cradle_to_site_annual":       None,
            "eol_emissions_annual":        None,
            "embodied_emissions_annual":   None,
            "lifetime_used":               life_time,
            "flags":                       "; ".join(flags),
        }

    # ── Production emissions ───────────────────────────────────────────────────
    if prod_raw is None:
        flags.append("production_emissions: not provided")
        production_emissions_annual = None
    else:
        production_emissions_annual = round(prod_raw / life_time, 4)

    # ── Installation ───────────────────────────────────────────────────────────
    if inst_qty is None:
        flags.append("installation_quantity: not provided")
    if inst_ef is None:
        flags.append("installation_emission_factor: not provided")

    if inst_qty is not None and inst_ef is not None:
        # Unit compatibility check
        inst_unit    = str(row.get("installation_unit", "") or "")
        inst_ef_unit = str(row.get("installation_emission_factor_unit", "") or "")
        unit_flag = _check_unit_match(inst_unit, inst_ef_unit, "installation")
        if unit_flag:
            flags.append(unit_flag)
        install_total        = round(inst_qty * inst_ef, 4)
        install_total_annual = round(install_total / life_time, 4)
    else:
        install_total        = None
        install_total_annual = None

    # ── Cradle-to-site (production + installation) ─────────────────────────────
    if production_emissions_annual is not None and install_total_annual is not None:
        cradle_to_site_annual = round(production_emissions_annual + install_total_annual, 4)
    elif production_emissions_annual is not None:
        cradle_to_site_annual = production_emissions_annual
        flags.append("cradle_to_site_annual: partial — installation not provided")
    elif install_total_annual is not None:
        cradle_to_site_annual = install_total_annual
        flags.append("cradle_to_site_annual: partial — production not provided")
    else:
        cradle_to_site_annual = None

    # ── End of life ────────────────────────────────────────────────────────────
    if eol_raw is None:
        flags.append("endoflife_emissions: not provided")
        eol_emissions_annual = None
    else:
        eol_emissions_annual = round(eol_raw / life_time, 4)

    # ── Maintenance unit check (if present) ───────────────────────────────────
    maint_qty     = _f(row.get("maintenance_quantity"))
    maint_ef_unit = str(row.get("maintenance_emission_factor_unit", "") or "")
    maint_unit    = str(row.get("maintenance_unit", "") or "")
    if maint_qty is not None and maint_ef_unit:
        unit_flag = _check_unit_match(maint_unit, maint_ef_unit, "maintenance")
        if unit_flag:
            flags.append(unit_flag)

    # ── Total embodied ─────────────────────────────────────────────────────────
    available = [x for x in (cradle_to_site_annual, eol_emissions_annual) if x is not None]
    if available:
        embodied_emissions_annual = round(sum(available), 4)
        if len(available) < 2:
            flags.append("embodied_emissions_annual: partial — some components not provided")
    else:
        embodied_emissions_annual = None

    return {
        "production_emissions_annual": production_emissions_annual,
        "install_total":               install_total,
        "install_total_annual":        install_total_annual,
        "cradle_to_site_annual":       cradle_to_site_annual,
        "eol_emissions_annual":        eol_emissions_annual,
        "embodied_emissions_annual":   embodied_emissions_annual,
        "lifetime_used":               life_time,
        "flags":                       "; ".join(flags),
    }
