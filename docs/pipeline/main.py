"""
main.py — Browser pipeline entry point.

Called from JavaScript via Pyodide. Accepts file bytes, runs the full
validate → enrich → analyze → visualise pipeline in-memory, and returns
a dict with the generated HTML report and status summaries.
"""

import io
import sys

import openpyxl

import validate as _validate
import enrich as _enrich
import analyze as _analyze
import visualise as _visualise


def _open_wb(file_bytes) -> openpyxl.Workbook:
    """Open an openpyxl Workbook from bytes (bytes, bytearray, or memoryview)."""
    return openpyxl.load_workbook(io.BytesIO(bytes(file_bytes)), data_only=True)


def run_analysis(
    cell_site_bytes,
    active_bytes,
    passive_bytes,
    infra_bytes,
    ef_bytes=None,
) -> dict:
    """
    Full browser pipeline.

    Parameters
    ----------
    cell_site_bytes  : bytes — cell_site.xlsx content
    active_bytes     : bytes — active_components.xlsx content
    passive_bytes    : bytes — passive_components.xlsx content
    infra_bytes      : bytes — infrastructure.xlsx content
    ef_bytes         : bytes or None — emission_factors.xlsx (optional)

    Returns
    -------
    dict with keys:
      "html"         : str   — complete self-contained HTML report
      "validation"   : dict  — {schema: {total, ok, warned, rejected}}
      "enrichment"   : dict  — {schema: {provider: fill_count}}
      "unresolved"   : list  — [(search_key, field, unit)] not filled
      "errors"       : list  — error strings, empty on success
    """
    errors = []

    # ── 1. Open workbooks ──────────────────────────────────────────────────────
    try:
        cs_wb     = _open_wb(cell_site_bytes)
        active_wb = _open_wb(active_bytes)
        passive_wb = _open_wb(passive_bytes)
        infra_wb  = _open_wb(infra_bytes)
    except Exception as exc:
        return {"html": "", "validation": {}, "enrichment": {},
                "unresolved": [], "errors": [f"Failed to open Excel files: {exc}"]}

    ef_table = {}
    if ef_bytes is not None:
        try:
            ef_wb    = _open_wb(ef_bytes)
            ef_table = _enrich.parse_ef_table(ef_wb)
        except Exception as exc:
            errors.append(f"Could not read emission_factors.xlsx: {exc}")

    # ── 2. Validate ────────────────────────────────────────────────────────────
    try:
        (cs_rows, active_rows, passive_rows, infra_rows,
         cell_site_lookup, val_summary) = _validate.validate_all(
             cs_wb, active_wb, passive_wb, infra_wb)
    except Exception as exc:
        return {"html": "", "validation": {}, "enrichment": {},
                "unresolved": [], "errors": [f"Validation failed: {exc}"]}

    if not (cs_rows or active_rows or passive_rows or infra_rows):
        return {"html": "", "validation": val_summary, "enrichment": {},
                "unresolved": [], "errors": ["No valid rows found in any file."]}

    # ── 3. Enrich ──────────────────────────────────────────────────────────────
    enrich_summary = {}
    all_unresolved = []

    active_enriched, active_enrich_s, active_unres = _enrich.enrich_in_memory(
        active_rows, "active", ef_table=ef_table, cell_site_lookup=cell_site_lookup)
    enrich_summary["active"] = active_enrich_s
    all_unresolved.extend(active_unres)

    passive_enriched, passive_enrich_s, passive_unres = _enrich.enrich_in_memory(
        passive_rows, "passive", ef_table=ef_table, cell_site_lookup=cell_site_lookup)
    enrich_summary["passive"] = passive_enrich_s
    all_unresolved.extend(passive_unres)

    infra_enriched, infra_enrich_s, infra_unres = _enrich.enrich_in_memory(
        infra_rows, "infrastructure", ef_table=ef_table, cell_site_lookup=cell_site_lookup)
    enrich_summary["infrastructure"] = infra_enrich_s
    all_unresolved.extend(infra_unres)

    # Enrich cell site EFs (electricity/fuel/refrigerant) for site_op computation
    cs_enriched, _, _ = _enrich.enrich_in_memory(
        cs_rows, "cell_site", ef_table=ef_table, cell_site_lookup={})
    # Rebuild lookup with enriched cell sites
    cell_site_lookup_enriched = _validate.build_cell_site_lookup(cs_enriched)

    # ── 4. Analyze ─────────────────────────────────────────────────────────────
    try:
        summary, op_rows, emb_rows, sensitivity, uncertainty = _analyze.analyze_in_memory(
            active_enriched, passive_enriched, infra_enriched,
            cs_enriched, ef_table=ef_table)
    except Exception as exc:
        errors.append(f"Analysis failed: {exc}")
        return {"html": "", "validation": val_summary, "enrichment": enrich_summary,
                "unresolved": all_unresolved, "errors": errors}

    # ── 5. Visualise ───────────────────────────────────────────────────────────
    enriched_for_viz = {
        "active":         active_enriched,
        "passive":        passive_enriched,
        "infrastructure": infra_enriched,
    }
    try:
        data = _visualise.build_data(
            summary=summary,
            op_rows=op_rows,
            emb_rows=emb_rows,
            cell_sites=cs_enriched,
            sensitivity=sensitivity,
            uncertainty=uncertainty,
            enriched=enriched_for_viz,
        )
        html = _visualise.generate_html(data)
    except Exception as exc:
        errors.append(f"Report generation failed: {exc}")
        return {"html": "", "validation": val_summary, "enrichment": enrich_summary,
                "unresolved": all_unresolved, "errors": errors}

    return {
        "html":       html,
        "validation": val_summary,
        "enrichment": enrich_summary,
        "unresolved": all_unresolved,
        "errors":     errors,
    }
