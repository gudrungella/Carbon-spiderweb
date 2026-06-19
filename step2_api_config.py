"""
API configuration for emission factor and power data enrichment.

Secrets (API keys, client credentials) belong in .env — not here.
Copy .env.example to .env and fill in your values.

Two provider hierarchies:
  EMISSION_HIERARCHY — tried in order for emission factor fields
  POWER_HIERARCHY    — tried in order for power_idle and power_max
"""
from pathlib import Path

# ---------------------------------------------------------------------------
# Provider hierarchies
# ---------------------------------------------------------------------------

# Emission factor providers: custom file → Rejoose → Climatiq → EPD → Resilio → Ecoinvent
EMISSION_HIERARCHY = ["custom_file", "rejoose", "climatiq", "epd", "resilio", "ecoinvent"]

# Power data providers: custom_file first, then tscircuit
POWER_HIERARCHY = ["custom_file", "tscircuit"]

# Lifetime defaults: fills life_time from a built-in type/subtype lookup table
LIFETIME_HIERARCHY = ["lifetime_defaults"]

# ---------------------------------------------------------------------------
# Custom file (priority 1 for emission factors)
# Local Excel lookup keyed by id. Set emission_factor values to fill fields.
# Rows with blank emission_factor are written by the manual fallback step.
# ---------------------------------------------------------------------------
CUSTOM_FILE_PATH = Path(__file__).parent / "data" / "emission_factors.xlsx"

# ---------------------------------------------------------------------------
# Provider base URLs (change only if the services move)
# ---------------------------------------------------------------------------

REJOOSE_BASE_URL      = "https://app.rejoose.com/"

CLIMATIQ_BASE_URL     = "https://api.climatiq.io"
CLIMATIQ_DATA_VERSION = "^6"   # "latest v6.x"; pin to e.g. "6.1" if needed

EPD_BASE_URL          = "https://epd.apim.developer.azure-api.net"

RESILIO_BASE_URL      = "https://db.resilio.tech"

ECOINVENT_TOKEN_URL   = "https://sso.ecoinvent.org/realms/ecoinvent/protocol/openid-connect/token"
ECOINVENT_BASE_URL    = "https://api.ecoinvent.org"

TSCIRCUIT_BASE_URL    = "https://api.tscircuit.com"
