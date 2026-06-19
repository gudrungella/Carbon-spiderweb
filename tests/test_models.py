"""Tests for step6_model_embodied.py and step7_model_operational.py."""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))

import step6_model_embodied as s6
import step7_model_operational as s7

HOURS_PER_YEAR = 8760
EF_TABLE = {}  # tests pass EF directly in the row


# ─── Step 6 helpers ────────────────────────────────────────────────────────────

def base_row_6(**overrides):
    """Full row with all primary inputs present."""
    row = {
        "life_time": 10.0,
        "production_emissions": 100.0,
        "installation_quantity": 5.0,
        "installation_emission_factor": 2.0,
        "installation_unit": "km",
        "installation_emission_factor_unit": "kgCO2eq/km",
        "endoflife_emissions": 50.0,
    }
    row.update(overrides)
    return row


# ─── Step 6 tests ──────────────────────────────────────────────────────────────

class TestStep6AllInputsPresent:
    def setup_method(self):
        self.result = s6.compute(base_row_6())

    def test_production_annual(self):
        assert self.result["production_emissions_annual"] == 10.0

    def test_install_total(self):
        assert self.result["install_total"] == 10.0

    def test_install_total_annual(self):
        assert self.result["install_total_annual"] == 1.0

    def test_cradle_to_site(self):
        assert self.result["cradle_to_site_annual"] == 11.0

    def test_eol_annual(self):
        assert self.result["eol_emissions_annual"] == 5.0

    def test_embodied_total(self):
        assert self.result["embodied_emissions_annual"] == 16.0

    def test_no_flags(self):
        assert self.result["flags"] == ""


class TestStep6AllAbsent:
    def setup_method(self):
        self.result = s6.compute({})

    def test_all_outputs_none(self):
        for key in [
            "production_emissions_annual", "install_total", "install_total_annual",
            "cradle_to_site_annual", "eol_emissions_annual", "embodied_emissions_annual",
        ]:
            assert self.result[key] is None, f"{key} should be None"

    def test_flag_present(self):
        assert "no embodied input data provided" in self.result["flags"]


class TestStep6OnlyProductionProvided:
    def setup_method(self):
        self.result = s6.compute({"life_time": 10.0, "production_emissions": 100.0})

    def test_production_annual(self):
        assert self.result["production_emissions_annual"] == 10.0

    def test_install_none(self):
        assert self.result["install_total"] is None
        assert self.result["install_total_annual"] is None

    def test_cradle_to_site_partial(self):
        assert self.result["cradle_to_site_annual"] == 10.0
        assert "cradle_to_site_annual: partial" in self.result["flags"]

    def test_eol_none(self):
        assert self.result["eol_emissions_annual"] is None

    def test_embodied_partial(self):
        assert self.result["embodied_emissions_annual"] == 10.0
        assert "embodied_emissions_annual: partial" in self.result["flags"]


class TestStep6LifetimeZero:
    def setup_method(self):
        self.result = s6.compute({"production_emissions": 100.0, "life_time": 0})

    def test_lifetime_defaulted(self):
        assert self.result["lifetime_used"] == 1.0

    def test_flag_present(self):
        assert "life_time not provided — defaulted to 1.0 year" in self.result["flags"]

    def test_production_uses_default_lifetime(self):
        assert self.result["production_emissions_annual"] == 100.0


class TestStep6InstallationUnitMismatch:
    def setup_method(self):
        self.result = s6.compute({
            "production_emissions": 100.0,
            "life_time": 10.0,
            "installation_quantity": 5.0,
            "installation_emission_factor": 2.0,
            "installation_unit": "trip",
            "installation_emission_factor_unit": "kgCO2eq/km",
        })

    def test_flag_fired(self):
        assert "installation: unit mismatch" in self.result["flags"]

    def test_computation_still_proceeds(self):
        assert self.result["install_total"] == 10.0


class TestStep6MaintenanceUnitMismatch:
    def setup_method(self):
        self.result = s6.compute({
            "production_emissions": 100.0,
            "life_time": 10.0,
            "maintenance_quantity": 3.0,
            "maintenance_unit": "trip",
            "maintenance_emission_factor_unit": "kgCO2eq/km",
        })

    def test_flag_fired(self):
        assert "maintenance: unit mismatch" in self.result["flags"]


# ─── Step 7 helpers ────────────────────────────────────────────────────────────

def base_row_7(**overrides):
    """Active row with electricity source, EF in row, and complete maintenance."""
    row = {
        "power_source": "electricity",
        "power_source_emission_factor": 0.5,
        "maintenance_quantity": 1.0,
        "maintenance_emission_factor": 10.0,
    }
    row.update(overrides)
    return row


# ─── Step 7 tests ──────────────────────────────────────────────────────────────

class TestStep7NotActive:
    def setup_method(self):
        self.result = s7.compute({}, schema="passive", ef_table=EF_TABLE)

    def test_passive_schema_path(self):
        assert self.result["power_path"] == "not_applicable"

    def test_passive_all_none(self):
        for key in ["op_energy", "op_maintenance", "op_total", "annual_consumption"]:
            assert self.result[key] is None, f"{key} should be None for passive schema"


@pytest.mark.parametrize("unit,qty,expected_consumption", [
    ("W",   100.0,  round(100.0 * HOURS_PER_YEAR / 1000.0, 2)),
    ("kW",    1.0,  round(1.0   * HOURS_PER_YEAR,           2)),
    ("kWh", 1000.0, 1000.0),
])
def test_step7_unit_conversions(unit, qty, expected_consumption):
    r = s7.compute(base_row_7(power_quantity=qty, power_unit=unit), "active", EF_TABLE)
    assert r["annual_consumption"] == expected_consumption
    assert r["power_path"] == "quantity"


class TestStep7EstimatedPath:
    def test_normal_idle_max(self):
        r = s7.compute(
            base_row_7(power_idle=100.0, power_max=500.0),
            "active", EF_TABLE,
        )
        assert r["power_path"] == "estimated"
        assert r["power_estimated_w"] == 420.0
        assert r["annual_consumption"] == round(420.0 * HOURS_PER_YEAR / 1000.0, 2)

    def test_power_max_less_than_idle_flag(self):
        r = s7.compute(
            base_row_7(power_idle=500.0, power_max=100.0),
            "active", EF_TABLE,
        )
        assert "power_max < power_idle" in r["flags"]

    def test_power_max_less_than_idle_still_computes(self):
        r = s7.compute(
            base_row_7(power_idle=500.0, power_max=100.0),
            "active", EF_TABLE,
        )
        assert r["power_estimated_w"] == 180.0

    def test_idle_missing_flags(self):
        r = s7.compute(base_row_7(power_max=500.0), "active", EF_TABLE)
        assert "power_idle: not provided" in r["flags"]
        assert r["annual_consumption"] is None

    def test_both_idle_and_max_missing(self):
        r = s7.compute(
            {"power_source": "electricity", "power_source_emission_factor": 0.5},
            "active", EF_TABLE,
        )
        assert "power_idle: not provided" in r["flags"]
        assert "power_max: not provided" in r["flags"]


class TestStep7NoPowerData:
    def test_missing_source_and_quantity(self):
        r = s7.compute({}, "active", EF_TABLE)
        assert "power_source: not provided" in r["flags"]
        assert "power_quantity: not provided" in r["flags"]
        assert "op_energy: could not be calculated" in r["flags"]
        assert r["op_energy"] is None
