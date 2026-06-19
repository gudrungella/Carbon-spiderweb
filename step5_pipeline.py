"""
pipeline.py — Telecom inventory pipeline: validate → enrich.

Usage:
    python3 pipeline.py

Steps:
  1. Validate all four Excel files (cell_site, active, passive, infrastructure)
     against schema rules and POST valid rows to telecom_api.
  2. Print validation summary.
  3. If ERROR rows exist, confirm before continuing.
  4. Enrich active, passive, and infrastructure files.
     OK and WARNED rows are enriched; REJECTED/API_ERROR rows are skipped.
  5. Print enrichment summary.

The API server must be running before executing this script:
    uvicorn main:app --host 0.0.0.0 --port 3000 --reload   (from inside telecom_api/)
"""

import argparse
import sys
from functools import partial

from step3_validate import (
    check_server,
    load_sheet, to_dict, present, to_num,
    validate_cell_site, validate_active_component,
    validate_passive_component, validate_infrastructure,
    fetch_cell_site_lookup, collect_cell_site_ids, post_row,
    compute_quality, RowResult, Issue,
    print_row, print_summary, process_file,
    fetch_all_from_api, validate_from_api,
    BASE_URL, DATA_DIR,
)
import step4_enrich as enrich


# ---------------------------------------------------------------------------
# Validation pass (mirrors validate.main() but returns results)
# ---------------------------------------------------------------------------

def run_validation() -> dict:
    """Run full validation pass and return all_results dict."""
    check_server()

    passive_cell_site_ids = collect_cell_site_ids("passive_components.xlsx")
    infra_cell_site_ids   = collect_cell_site_ids("infrastructure.xlsx")
    cell_site_lookup      = fetch_cell_site_lookup()
    known_cell_site_ids   = set(cell_site_lookup.keys())

    all_results = {}

    # ── Cell sites ────────────────────────────────────────────────────────────
    cs_headers, cs_rows = load_sheet("cell_site.xlsx")
    cs_results = []

    if cs_headers is None:
        print("\nSKIP: cell_site.xlsx — not found or empty.")
    else:
        print(f"\n--- cell_site.xlsx ({len(cs_rows)} rows) ---")
        for row in cs_rows:
            row_num = row[0].row
            rd      = to_dict(cs_headers, row)
            issues  = validate_cell_site(rd)

            errors = [i for i in issues if i.level == "ERROR"]
            if errors:
                rr = RowResult(row_num, "REJECTED", issues)
                print_row(rr)
                cs_results.append(rr)
                continue

            body = {k: v for k, v in rd.items() if present(v)}
            resp = post_row(f"{BASE_URL}/cell-sites", body)

            if isinstance(resp, Exception):
                rr = RowResult(row_num, "API_ERROR",
                               issues + [Issue("ERROR", "api", str(resp))])
                print_row(rr)
                cs_results.append(rr)
                continue

            if resp.status_code in (200, 201):
                returned = resp.json()
                cs_id    = returned.get("cell_site_id")
                cell_site_lookup[cs_id]    = returned
                known_cell_site_ids.add(cs_id)

                if cs_id not in passive_cell_site_ids:
                    issues.append(Issue("WARNING", "cell_site_id",
                                        f"Cell site {cs_id} has no linked passive component"))
                if cs_id not in infra_cell_site_ids:
                    issues.append(Issue("WARNING", "cell_site_id",
                                        f"Cell site {cs_id} has no linked infrastructure"))

                warnings = [i for i in issues if i.level == "WARNING"]
                quality  = compute_quality(issues)
                status   = "WARNED" if warnings else "OK"
                ef       = [i.field for i in warnings]
                rr = RowResult(row_num, status, issues, quality,
                               {"cell_site_id": cs_id}, ef)
            else:
                try:    msg = resp.json().get("message", resp.text)
                except: msg = resp.text
                rr = RowResult(row_num, "API_ERROR",
                               issues + [Issue("ERROR", "api", f"HTTP {resp.status_code}: {msg}")])

            print_row(rr)
            cs_results.append(rr)

    all_results["cell_site.xlsx"] = cs_results

    # ── Active components ─────────────────────────────────────────────────────
    all_results["active_components.xlsx"] = process_file(
        "active_components.xlsx",
        validator_fn=partial(validate_active_component, cell_site_lookup=cell_site_lookup),
        url_fn=lambda rd: f"{BASE_URL}/cell-sites/{rd['cell_site_id']}/active-components",
        skip_keys={"cell_site_id"},
        known_cell_site_ids=known_cell_site_ids,
    )

    # ── Passive components ────────────────────────────────────────────────────
    all_results["passive_components.xlsx"] = process_file(
        "passive_components.xlsx",
        validator_fn=validate_passive_component,
        url_fn=lambda rd: f"{BASE_URL}/cell-sites/{rd['cell_site_id']}/passive-components",
        skip_keys={"cell_site_id"},
        known_cell_site_ids=known_cell_site_ids,
    )

    # ── Infrastructure ────────────────────────────────────────────────────────
    all_results["infrastructure.xlsx"] = process_file(
        "infrastructure.xlsx",
        validator_fn=validate_infrastructure,
        url_fn=lambda rd: f"{BASE_URL}/cell-sites/{rd['cell_site_id']}/infrastructure",
        skip_keys={"cell_site_id"},
        known_cell_site_ids=known_cell_site_ids,
        allow_missing_cs=True,
    )

    return all_results


# ---------------------------------------------------------------------------
# Enrichment pass
# ---------------------------------------------------------------------------

def run_enrichment(all_results: dict) -> dict:
    """Enrich each component file. Returns dict of xlsx_name → enriched rows."""
    print("\n" + "=" * 62)
    print("ENRICHMENT")
    print("=" * 62)

    targets = [
        ("active_components.xlsx",  "active"),
        ("passive_components.xlsx", "passive"),
        ("infrastructure.xlsx",     "infrastructure"),
    ]

    enrichment_summary = {}
    for xlsx_name, schema in targets:
        results  = all_results.get(xlsx_name, [])
        eligible = [r for r in results if r.status in ("OK", "WARNED")]
        total    = len(results)

        if not eligible:
            print(f"\n  {xlsx_name} — no eligible rows (0 of {total} passed validation).")
            enrichment_summary[xlsx_name] = []
            continue

        skipped = total - len(eligible)
        skipped_note = f"  ({skipped} ERROR/API_ERROR row(s) skipped)" if skipped else ""
        print(f"\n--- {xlsx_name} ({len(eligible)} of {total} rows eligible){skipped_note} ---")
        filepath = DATA_DIR / xlsx_name
        enriched_rows = enrich.enrich_file(str(filepath), schema)
        enrichment_summary[xlsx_name] = enriched_rows

    return enrichment_summary


# ---------------------------------------------------------------------------
# Enrichment summary
# ---------------------------------------------------------------------------

def print_enrichment_summary(enrichment_summary: dict) -> None:
    print("\n" + "=" * 62)
    print("ENRICHMENT SUMMARY")
    print("=" * 62)

    targets = [
        ("active_components.xlsx",  "active"),
        ("passive_components.xlsx", "passive"),
        ("infrastructure.xlsx",     "infrastructure"),
    ]

    for xlsx_name, schema in targets:
        rows = enrichment_summary.get(xlsx_name)
        if rows is None:
            print(f"\n  {xlsx_name}  —  skipped")
            continue

        field_cfgs = enrich.ENRICHABLE_FIELDS.get(schema, [])

        # Collect per-source stats and per-field detail
        # filled_detail: {source: {field: [(search_key, value), ...]}}
        filled_detail: dict = {}
        still_empty: list   = []

        for row in rows:
            for fc in field_cfgs:
                fname = fc["field"]
                src   = row.get(f"{fname}_source")
                val   = row.get(fname)
                if src and val is not None:
                    filled_detail.setdefault(src, {}).setdefault(fname, [])
                    # derive the search key used (best available identifier)
                    keys = enrich.get_search_keys(row, schema, fname)
                    sk   = keys[0] if keys else "?"
                    filled_detail[src][fname].append((sk, val))
                elif val is None or not str(val).strip():
                    if enrich.precondition_met(row, fc) and enrich.get_search_keys(row, schema, fname):
                        still_empty.append(fname)

        print(f"\n  {xlsx_name}")
        if not rows:
            print(f"    No rows processed.")
            continue

        if filled_detail:
            for src in sorted(filled_detail.keys()):
                conf        = enrich.CONFIDENCE.get(src, 0.5)
                total_fills = sum(len(v) for v in filled_detail[src].values())
                print(f"    [{src}]  confidence {conf:.0%}  — {total_fills} fill(s)")
                for fname, entries in sorted(filled_detail[src].items()):
                    unit_label = " yr" if fname == "life_time" else ""
                    # Group entries by (sk, val) for compact display
                    from collections import Counter
                    counts = Counter(entries)
                    parts = []
                    for (sk, val), cnt in sorted(counts.items(), key=lambda x: str(x[0][0])):
                        suffix = f" ×{cnt}" if cnt > 1 else ""
                        parts.append(f"{sk} → {val}{unit_label}{suffix}")
                    print(f"      {fname}: {',  '.join(parts)}")
        else:
            print(f"    No fields were enriched.")

        if still_empty:
            unique_empty = sorted(set(still_empty))
            print(f"    Still empty ({len(still_empty)} instance(s)): {', '.join(unique_empty)}")
            print(f"    -> Fill in data/emission_factors.xlsx and re-run.")
        elif filled_detail:
            print(f"    All enrichable fields filled.")


# ---------------------------------------------------------------------------
# API source — validation and enrichment
# ---------------------------------------------------------------------------

def run_validation_from_api():
    """Validate data already in telecom_api (no POSTing). Returns (all_results, rows_by_schema)."""
    return validate_from_api()


def run_enrichment_api(rows_by_schema: dict, all_results: dict,
                       cell_sites: list = None) -> dict:
    """Enrich rows fetched from the API. Writes *_enriched.xlsx for each component type."""
    print("\n" + "=" * 62)
    print("ENRICHMENT")
    print("=" * 62)

    targets = [
        ("active_components.xlsx",  "active",        "active_components"),
        ("passive_components.xlsx", "passive",        "passive_components"),
        ("infrastructure.xlsx",     "infrastructure", "infrastructure"),
    ]

    enrichment_summary = {}
    for xlsx_name, schema, out_name in targets:
        results  = all_results.get(xlsx_name, [])
        eligible = [r for r in results if r.status in ("OK", "WARNED")]
        total    = len(results)
        rows     = rows_by_schema.get(schema, [])

        if not eligible:
            print(f"\n  {xlsx_name} — no eligible rows (0 of {total} passed validation).")
            enrichment_summary[xlsx_name] = []
            continue

        skipped      = total - len(eligible)
        skipped_note = f"  ({skipped} row(s) skipped)" if skipped else ""
        print(f"\n--- {xlsx_name} ({len(eligible)} of {total} rows eligible){skipped_note} ---")
        enriched_rows = enrich.enrich_rows(rows, schema, out_name,
                                            cell_sites=cell_sites)
        enrichment_summary[xlsx_name] = enriched_rows

    return enrichment_summary


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Telecom inventory pipeline")
    parser.add_argument(
        "--source", choices=["excel", "api"], default="excel",
        help="Input source: 'excel' reads from data/*.xlsx (default); "
             "'api' reads from telecom_api (data entered directly via the API)",
    )
    parser.add_argument(
        "--yes", "-y", action="store_true",
        help="Skip confirmation prompt when errors are found and continue to enrichment",
    )
    args = parser.parse_args()

    print("=" * 62)
    print("TELECOM INVENTORY PIPELINE")
    print(f"Source: {args.source.upper()}")
    print("=" * 62)

    print("\nStep 1: Validation")
    print("-" * 62)

    if args.source == "api":
        all_results, rows_by_schema = run_validation_from_api()
        cell_sites = rows_by_schema.get("cell_sites", [])
    else:
        all_results = run_validation()
        cell_sites  = []

    print_summary(all_results)

    has_errors = any(
        r.status in ("REJECTED", "API_ERROR")
        for results in all_results.values()
        for r in results
    )
    if has_errors:
        print("\n  Some rows have validation errors and will be skipped during enrichment.")
        if args.yes:
            print("  Continuing (--yes flag set).")
        else:
            ans = input("  Continue to enrichment? [y/N] ").strip().lower()
            if ans not in ("y", "yes"):
                print("  Aborted.")
                sys.exit(0)

    print("\n\nStep 2: Enrichment")
    print("-" * 62)

    if args.source == "api":
        enrichment_summary = run_enrichment_api(rows_by_schema, all_results,
                                                cell_sites=cell_sites)
    else:
        enrichment_summary = run_enrichment(all_results)

    print_enrichment_summary(enrichment_summary)


if __name__ == "__main__":
    main()
