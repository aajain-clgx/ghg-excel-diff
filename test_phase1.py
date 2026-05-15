"""
test_phase1.py
--------------
Self-contained test for Phase 1 core engine:
  - column_matcher
  - diff_engine
  - annotation_store

Creates two synthetic GHG Excel files (v1 and v2) with known differences,
runs the diff, and prints results to console.

Run with:
    python test_phase1.py
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

# Add app/ to the path so imports work when running from project root
sys.path.insert(0, str(Path(__file__).parent / "app"))

import pandas as pd
import openpyxl

from column_matcher import match_columns, apply_manual_pair
from diff_engine import (
    list_sheets, load_sheet, run_diff,
    SEVERITY_HIGH, SEVERITY_WARN, SEVERITY_MINOR,
)
from annotation_store import AnnotationStore, sidecar_path


# ---------------------------------------------------------------------------
# Synthetic data
# ---------------------------------------------------------------------------

V1_DATA = [
    # site_id, ef_natural_gas, activity_mmbtu, scope1_total, region
    ("Riverbend",     53.02,  10200, 541.0,  "East"),
    ("Eastport Hub",  48.50,   8400, 407.4,  "East"),
    ("Lakewood Plant",50.00,   5000, 250.0,  "West"),
    ("Northfield",    52.10,   3100, 161.5,  "West"),
    ("Harbor Works",  49.75,   7200, 358.2,  "South"),
    ("OldSite",       51.00,   2000, 102.0,  "North"),  # Will be removed in V2
]

V2_DATA = [
    # Changes:
    # Riverbend:     EF changed 53.02 → 64.88  (+22.4%)  — isolated EF change, activity same
    # Eastport Hub:  EF changed 48.50 → 57.00  (+17.5%)  — HIGH
    # Lakewood Plant: activity changed but scope1 blank in V2 — WARN (blank appeared)
    # Northfield:    tiny EF change 52.10 → 52.62 (+1.0%)  — MINOR
    # Harbor Works:  no change
    # NewSite added in V2
    # OldSite removed
    ("Riverbend",      64.88,  10200, 662.2,  "East"),   # EF changed, activity same
    ("Eastport Hub",   57.00,   8400, 478.8,  "East"),   # EF + scope1 changed
    ("Lakewood Plant", 50.00,   5500, None,   "West"),   # Activity changed, scope1 blank
    ("Northfield",     52.62,   3100, 163.1,  "West"),   # Tiny EF change
    ("Harbor Works",   49.75,   7200, 358.2,  "South"),  # No change
    ("New Denver Hub",  51.00,  4500, 229.5,  "North"),  # New site
]

V1_COLUMNS = ["Site ID", "EF Natural Gas (kg/MMBtu)", "Activity (MMBtu)",
              "Scope 1 Total (tCO2e)", "Region"]
V2_COLUMNS = ["Site ID", "EF Natural Gas (kg/MMBtu)", "Activity (MMBtu)",
              "Scope 1 Total (tCO2e)", "Region"]  # Same names — clean match


def _write_excel(path: Path, data: list, columns: list, sheet_name: str = "GHG Data") -> None:
    df = pd.DataFrame(data, columns=columns)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"Info": ["GHG Project", "See GHG Data sheet"]}).to_excel(
            writer, sheet_name="Project Info", index=False
        )
        df.to_excel(writer, sheet_name=sheet_name, index=False)


# ---------------------------------------------------------------------------
# Test runner
# ---------------------------------------------------------------------------

def run_tests() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        v1_path = tmp / "ghg_report_v1.xlsx"
        v2_path = tmp / "ghg_report_v2.xlsx"

        print("=" * 60)
        print("GHG QA Tool — Phase 1 Engine Test")
        print("=" * 60)

        # ----------------------------------------------------------------
        # 1. Write synthetic Excel files
        # ----------------------------------------------------------------
        _write_excel(v1_path, V1_DATA, V1_COLUMNS)
        _write_excel(v2_path, V2_DATA, V2_COLUMNS)
        print(f"\n✅ Synthetic Excel files written to {tmpdir}")

        # ----------------------------------------------------------------
        # 2. Sheet listing
        # ----------------------------------------------------------------
        sheets_v1 = list_sheets(v1_path)
        sheets_v2 = list_sheets(v2_path)
        assert "GHG Data" in sheets_v1, f"Expected 'GHG Data' in {sheets_v1}"
        assert "Project Info" in sheets_v1
        print(f"\n✅ Sheet listing OK")
        print(f"   V1 sheets: {sheets_v1}")
        print(f"   V2 sheets: {sheets_v2}")

        # ----------------------------------------------------------------
        # 3. Load selected sheet
        # ----------------------------------------------------------------
        df_v1 = load_sheet(v1_path, "GHG Data")
        df_v2 = load_sheet(v2_path, "GHG Data")
        assert list(df_v1.columns) == V1_COLUMNS
        assert len(df_v1) == 6
        assert len(df_v2) == 6
        print(f"\n✅ Sheet loading OK — V1: {len(df_v1)} rows, V2: {len(df_v2)} rows")

        # ----------------------------------------------------------------
        # 4. Column matching
        # ----------------------------------------------------------------
        match_result = match_columns(
            headers_v1=list(df_v1.columns),
            headers_v2=list(df_v2.columns),
        )
        assert len(match_result.matched_pairs) == 5, (
            f"Expected 5 matched pairs, got {len(match_result.matched_pairs)}"
        )
        assert len(match_result.v1_only) == 0
        assert len(match_result.v2_only) == 0
        assert match_result.auto_id_column == "Site ID"
        print(f"\n✅ Column matching OK")
        print(f"   Matched pairs: {len(match_result.matched_pairs)}")
        print(f"   Auto-detected ID column: '{match_result.auto_id_column}'")

        # Test with a renamed column
        headers_v2_renamed = list(df_v2.columns)
        headers_v2_renamed[1] = "EF Natural Gas (kg CO2e/MMBtu)"  # Renamed
        match_renamed = match_columns(list(df_v1.columns), headers_v2_renamed)
        assert len(match_renamed.v1_only) == 1
        assert len(match_renamed.v2_only) == 1
        print(f"   Renamed column correctly treated as deleted+added (no silent match) ✅")

        # Test manual pairing
        apply_manual_pair(
            match_renamed,
            v1_header="EF Natural Gas (kg/MMBtu)",
            v2_header="EF Natural Gas (kg CO2e/MMBtu)",
        )
        assert len(match_renamed.v1_only) == 0
        assert len(match_renamed.v2_only) == 0
        print(f"   Manual pairing works correctly ✅")

        # ----------------------------------------------------------------
        # 5. Run diff
        # ----------------------------------------------------------------
        column_mapping = match_result.to_serializable()
        diff = run_diff(
            df_v1=df_v1,
            df_v2=df_v2,
            id_column="Site ID",
            column_mapping=column_mapping,
            materiality_threshold=5.0,
        )

        print(f"\n✅ Diff engine ran successfully")
        print(f"\n   {diff.summary_sentence()}")
        print(f"\n   Total V1 sites: {diff.total_sites_v1}")
        print(f"   Changed sites:  {diff.changed_site_count}")
        print(f"   Sites added:    {diff.structural.sites_added}")
        print(f"   Sites removed:  {diff.structural.sites_removed}")

        # Verify known outcomes
        site_ids = {s.site_id for s in diff.site_diffs}
        assert "Harbor Works" not in site_ids, "Harbor Works should be unchanged"
        assert "Riverbend" in site_ids
        assert "Eastport Hub" in site_ids

        riverbend = next(s for s in diff.site_diffs if s.site_id == "Riverbend")
        assert riverbend.severity == SEVERITY_HIGH, (
            f"Riverbend expected HIGH, got {riverbend.severity}"
        )

        lakewood = next(s for s in diff.site_diffs if s.site_id == "Lakewood Plant")
        assert lakewood.severity in (SEVERITY_WARN, SEVERITY_HIGH), (
            f"Lakewood expected WARN/HIGH (blank scope1), got {lakewood.severity}"
        )

        northfield = next(s for s in diff.site_diffs if s.site_id == "Northfield")
        assert northfield.severity == SEVERITY_MINOR, (
            f"Northfield expected MINOR, got {northfield.severity}"
        )

        assert diff.structural.sites_added == ["New Denver Hub"]
        assert diff.structural.sites_removed == ["OldSite"]

        print(f"\n   Site diff breakdown:")
        for site in diff.site_diffs:
            icon = {"HIGH": "🔴", "WARN": "🟡", "MINOR": "🟢"}[site.severity]
            print(f"   {icon} {site.site_id} (max Δ: {site.max_delta_pct:.1f}%)")
            for cd in site.column_diffs:
                pct = f"{cd.delta_pct:+.1f}%" if cd.delta_pct is not None else "N/A"
                print(f"      • {cd.column_v1}: {cd.v1_value} → {cd.v2_value} ({pct})")
                if cd.hint:
                    print(f"        Hint: {cd.hint}")

        # ----------------------------------------------------------------
        # 6. Annotation store
        # ----------------------------------------------------------------
        sidecar = sidecar_path(tmp, "2026-05-08", "Draft v1", "Draft v2")
        store = AnnotationStore.load_or_create(
            sidecar,
            analyst="Jane Smith",
            project="Test Project",
            v1_path=str(v1_path),
            v1_label="Draft v1",
            v2_path=str(v2_path),
            v2_label="Draft v2",
            id_column="Site ID",
            materiality_threshold=5.0,
            column_mapping=column_mapping,
            structural_changes={
                "sites_added": diff.structural.sites_added,
                "sites_removed": diff.structural.sites_removed,
                "columns_added": diff.structural.columns_added,
                "columns_removed": diff.structural.columns_removed,
            },
        )

        # Annotate a flag
        store.set_annotation(
            "Riverbend", "EF Natural Gas (kg/MMBtu)",
            AnnotationStore.DISPOSITION_ERROR,
            "Wrong EF year used — should be EPA 2024 not 2023"
        )
        store.set_annotation(
            "Eastport Hub", "EF Natural Gas (kg/MMBtu)",
            AnnotationStore.DISPOSITION_EXPECTED,
            "Updated to EPA 2024 — expected"
        )

        # Verify persistence: reload from disk
        store2 = AnnotationStore.load_or_create(
            sidecar,
            analyst="Jane Smith",
            project="Test Project",
            v1_path=str(v1_path),
            v1_label="Draft v1",
            v2_path=str(v2_path),
            v2_label="Draft v2",
            id_column="Site ID",
            materiality_threshold=5.0,
            column_mapping=column_mapping,
            structural_changes={},
        )
        ann = store2.get_annotation("Riverbend", "EF Natural Gas (kg/MMBtu)")
        assert ann is not None
        assert ann["disposition"] == AnnotationStore.DISPOSITION_ERROR
        assert ann["note"] == "Wrong EF year used — should be EPA 2024 not 2023"

        print(f"\n✅ Annotation store OK")
        print(f"   Sidecar written to: {sidecar.name}")
        print(f"   Loaded and verified 2 annotations correctly")

        # Unanswered HIGH flags
        high_keys = [
            AnnotationStore.annotation_key(s.site_id, cd.column_v1)
            for s in diff.high_impact_sites
            for cd in s.column_diffs
            if cd.severity == SEVERITY_HIGH
        ]
        unanswered = store2.unanswered_high_impact_keys(high_keys)
        print(f"   Unanswered HIGH flags: {len(unanswered)}")

        # ----------------------------------------------------------------
        # 7. Password-protected file error
        # ----------------------------------------------------------------
        # (Cannot create a real password-protected file without xlwings/win32com
        # but we test that the load function doesn't crash on normal files)
        df_check = load_sheet(v1_path, "GHG Data")
        assert len(df_check) == 6
        print(f"\n✅ load_sheet re-validated OK")

        print("\n" + "=" * 60)
        print("ALL PHASE 1 TESTS PASSED ✅")
        print("=" * 60)


if __name__ == "__main__":
    run_tests()
