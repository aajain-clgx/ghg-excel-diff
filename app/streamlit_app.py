"""
streamlit_app.py
----------------
GHG Excel QA Tool — Streamlit web UI entry point.

Run with:
    streamlit run app/streamlit_app.py

Three logical steps rendered on a single page via st.session_state:
  Step 1 — Project Setup    (file upload, labels, sheet selection)
  Step 2 — Column Matching  (review / adjust matched columns, set ID + threshold)
  Step 3 — Diff Results     (annotate per-site, export QA Log)

Core engine modules (column_matcher, diff_engine, annotation_store, export)
are unchanged from Phase 1.
"""

from __future__ import annotations

import io
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# ── Make sure app/ is importable when launched from project root ──────────
sys.path.insert(0, str(Path(__file__).parent))

from column_matcher import match_columns, apply_manual_pair, ColumnMatchResult
from diff_engine import (
    DiffResult, SiteDiff, ColumnDiff,
    run_diff, load_template_headers,
    SEVERITY_HIGH, SEVERITY_WARN, SEVERITY_MINOR,
)
from annotation_store import AnnotationStore, sidecar_path
from export import write_qa_log

# ── Constants ─────────────────────────────────────────────────────────────
SEVERITY_ICON = {SEVERITY_HIGH: "🔴", SEVERITY_WARN: "🟡", SEVERITY_MINOR: "🟢"}
STEP_NAMES = ["📁 Project Setup", "🔗 Column Matching", "🔍 Diff Results"]

st.set_page_config(
    page_title="GHG QA Tool",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ---------------------------------------------------------------------------
# Session state helpers
# ---------------------------------------------------------------------------

def _init_state() -> None:
    defaults = {
        "step": 1,
        "setup": None,           # dict from step 1
        "col_setup": None,       # dict from step 2
        "df_v1": None,
        "df_v2": None,
        "match_result": None,
        "diff_result": None,
        "store": None,
        "export_path": None,     # temp file path after export
        "filter_severity": "All",
        "selected_site": None,
        "sheet_select_data": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def _go(step: int) -> None:
    st.session_state["step"] = step


def _reset() -> None:
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _read_excel_bytes(uploaded_file) -> dict[str, pd.DataFrame]:
    """Return {sheet_name: DataFrame} from an uploaded file object."""
    buf = io.BytesIO(uploaded_file.read())
    return pd.read_excel(buf, sheet_name=None, dtype=str, keep_default_na=False)


def _sheets_from_upload(uploaded_file) -> list[str]:
    buf = io.BytesIO(uploaded_file.read())
    xl = pd.ExcelFile(buf)
    return xl.sheet_names


def _nav_bar() -> None:
    step = st.session_state["step"]
    cols = st.columns(len(STEP_NAMES))
    for i, name in enumerate(STEP_NAMES, start=1):
        with cols[i - 1]:
            if i < step:
                st.markdown(f"~~{name}~~ ✓", help="Completed")
            elif i == step:
                st.markdown(f"**{name}**")
            else:
                st.markdown(f"<span style='opacity:0.4'>{name}</span>",
                            unsafe_allow_html=True)
    st.divider()


# ---------------------------------------------------------------------------
# Step 1 — Project Setup
# ---------------------------------------------------------------------------

def _show_sheet_selector() -> None:
    """Second-pass UI when uploaded files contain multiple sheets."""
    data = st.session_state["sheet_select_data"]
    v1_sheets = list(data["all_v1"].keys())
    v2_sheets = list(data["all_v2"].keys())

    st.header("📁 Select Sheets")
    st.info(
        f"**{data['v1_label']}** has {len(v1_sheets)} sheet(s).  "
        f"**{data['v2_label']}** has {len(v2_sheets)} sheet(s).  "
        "Select the sheet to compare in each file."
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{data['v1_label']}**")
        v1_sheet = st.selectbox("Sheet", v1_sheets, key="v1_sheet_sel")
    with c2:
        st.markdown(f"**{data['v2_label']}**")
        v2_sheet = st.selectbox("Sheet", v2_sheets, key="v2_sheet_sel")

    bc1, bc2 = st.columns([1, 4])
    with bc1:
        if st.button("← Back"):
            st.session_state["sheet_select_data"] = None
            st.rerun()
    with bc2:
        if st.button("Continue →", type="primary", use_container_width=True):
            df_v1 = data["all_v1"][v1_sheet]
            df_v2 = data["all_v2"][v2_sheet]
            match_result = match_columns(
                headers_v1=list(df_v1.columns),
                headers_v2=list(df_v2.columns),
                template_headers=data.get("template_headers"),
            )
            st.session_state["setup"] = {
                "analyst":  data["analyst"],
                "project":  data["project"],
                "v1_label": data["v1_label"],
                "v2_label": data["v2_label"],
                "v1_sheet": v1_sheet,
                "v2_sheet": v2_sheet,
            }
            st.session_state["df_v1"] = df_v1
            st.session_state["df_v2"] = df_v2
            st.session_state["match_result"] = match_result
            st.session_state["sheet_select_data"] = None
            _go(2)
            st.rerun()


def _step1_project_setup() -> None:
    # Sheet-selection pass (triggered when files have multiple sheets)
    if st.session_state.get("sheet_select_data"):
        _show_sheet_selector()
        return

    st.header("📁 Project Setup")

    with st.form("project_form"):
        c1, c2 = st.columns(2)
        with c1:
            analyst = st.text_input("Analyst Name *", placeholder="Your name")
        with c2:
            project = st.text_input("Project / Report Name *",
                                    placeholder="e.g. GHG Inventory 2025 Q4")

        st.markdown("#### Version 1 (Baseline)")
        v1_file = st.file_uploader("Upload V1 Excel file *", type=["xlsx"],
                                   key="v1_upload")
        v1_label = st.text_input("V1 Label", value="V1",
                                 placeholder="e.g. Draft 1, Prior Year")

        st.markdown("#### Version 2 (Revised)")
        v2_file = st.file_uploader("Upload V2 Excel file *", type=["xlsx"],
                                   key="v2_upload")
        v2_label = st.text_input("V2 Label", value="V2",
                                 placeholder="e.g. Draft 2, Current Year")

        st.markdown("#### Template (optional)")
        tmpl_file = st.file_uploader("Upload column template .xlsx",
                                     type=["xlsx"], key="tmpl_upload")

        submitted = st.form_submit_button("Load Files →", type="primary",
                                          use_container_width=True)

    if not submitted:
        return

    # Validation
    errors = []
    if not analyst.strip():
        errors.append("Analyst Name is required.")
    if not project.strip():
        errors.append("Project Name is required.")
    if v1_file is None:
        errors.append("V1 file is required.")
    if v2_file is None:
        errors.append("V2 file is required.")
    if errors:
        for e in errors:
            st.error(e)
        return

    # Read files (template too — upload objects only available during form submission)
    with st.spinner("Reading Excel files…"):
        try:
            all_v1 = _read_excel_bytes(v1_file)
            all_v2 = _read_excel_bytes(v2_file)
        except Exception as exc:
            st.error(f"Could not read files: {exc}")
            return

    template_headers = None
    if tmpl_file is not None:
        try:
            template_headers = load_template_headers(io.BytesIO(tmpl_file.read()))
        except Exception:
            st.warning("Template file could not be read — template check skipped.")

    v1_sheets = list(all_v1.keys())
    v2_sheets = list(all_v2.keys())

    # Multi-sheet: store data and hand off to _show_sheet_selector
    if not (len(v1_sheets) == 1 and len(v2_sheets) == 1):
        st.session_state["sheet_select_data"] = {
            "analyst":          analyst.strip(),
            "project":          project.strip(),
            "v1_label":         v1_label.strip() or "V1",
            "v2_label":         v2_label.strip() or "V2",
            "all_v1":           all_v1,
            "all_v2":           all_v2,
            "template_headers": template_headers,
        }
        st.rerun()
        return

    # Single-sheet path — proceed directly
    v1_sheet = v1_sheets[0]
    v2_sheet = v2_sheets[0]
    df_v1 = all_v1[v1_sheet]
    df_v2 = all_v2[v2_sheet]

    match_result = match_columns(
        headers_v1=list(df_v1.columns),
        headers_v2=list(df_v2.columns),
        template_headers=template_headers,
    )

    st.session_state["setup"] = {
        "analyst": analyst.strip(),
        "project": project.strip(),
        "v1_label": v1_label.strip() or "V1",
        "v2_label": v2_label.strip() or "V2",
        "v1_sheet": v1_sheet,
        "v2_sheet": v2_sheet,
    }
    st.session_state["df_v1"] = df_v1
    st.session_state["df_v2"] = df_v2
    st.session_state["match_result"] = match_result
    _go(2)
    st.rerun()


# ---------------------------------------------------------------------------
# Step 2 — Column Matching
# ---------------------------------------------------------------------------

def _step2_column_matching() -> None:
    setup: dict = st.session_state["setup"]
    match: ColumnMatchResult = st.session_state["match_result"]
    df_v1: pd.DataFrame = st.session_state["df_v1"]
    df_v2: pd.DataFrame = st.session_state["df_v2"]

    v1_label = setup["v1_label"]
    v2_label = setup["v2_label"]

    st.header("🔗 Column Matching")
    st.caption(
        f"Comparing **{v1_label}** (sheet: {setup['v1_sheet']}) vs "
        f"**{v2_label}** (sheet: {setup['v2_sheet']})"
    )

    # ── Auto-matched columns ──────────────────────────────────────────────
    st.subheader("Auto-matched Columns")
    matched_pairs = [(m.v1_header, m.v2_header) for m in match.matched_pairs]
    if matched_pairs:
        tbl = pd.DataFrame(matched_pairs, columns=[v1_label, v2_label])
        st.dataframe(tbl, use_container_width=True, hide_index=True)
    else:
        st.warning("No columns matched automatically.")

    # ── Unmatched columns ─────────────────────────────────────────────────
    all_unmatched_v1 = [m.v1_header for m in match.v1_only]
    all_unmatched_v2 = [m.v2_header for m in match.v2_only]
    if all_unmatched_v1 or all_unmatched_v2:
        st.subheader("Unmatched Columns")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**Only in {v1_label}**")
            for h in all_unmatched_v1:
                st.markdown(f"- `{h}`")
        with c2:
            st.markdown(f"**Only in {v2_label}**")
            for h in all_unmatched_v2:
                st.markdown(f"- `{h}`")

    # ── Manual pairing ────────────────────────────────────────────────────
    if all_unmatched_v1 and all_unmatched_v2:
        with st.expander("➕ Manually pair unmatched columns"):
            c1, c2 = st.columns(2)
            with c1:
                sel_v1 = st.selectbox(f"{v1_label} column", ["—"] + all_unmatched_v1,
                                       key="manual_v1")
            with c2:
                sel_v2 = st.selectbox(f"{v2_label} column", ["—"] + all_unmatched_v2,
                                       key="manual_v2")
            if st.button("Add Pair") and sel_v1 != "—" and sel_v2 != "—":
                apply_manual_pair(match, sel_v1, sel_v2)
                st.rerun()

    # ── Structural changes preview ────────────────────────────────────────
    v1_ids = set(df_v1.iloc[:, 0].astype(str))
    v2_ids = set(df_v2.iloc[:, 0].astype(str))
    added = v2_ids - v1_ids
    removed = v1_ids - v2_ids
    if added or removed:
        with st.expander(f"🏗 Structural Changes — {len(added)} added, {len(removed)} removed"):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"**Sites added in {v2_label}**")
                for s in sorted(added):
                    st.markdown(f"- {s}")
            with c2:
                st.markdown(f"**Sites removed in {v2_label}**")
                for s in sorted(removed):
                    st.markdown(f"- {s}")

    # ── Configuration ─────────────────────────────────────────────────────
    st.subheader("Diff Configuration")
    all_cols_v1 = list(df_v1.columns)
    # Pre-select the auto-detected site ID if available
    default_id_idx = 0
    if match.auto_id_column and match.auto_id_column in all_cols_v1:
        default_id_idx = all_cols_v1.index(match.auto_id_column)

    c1, c2 = st.columns(2)
    with c1:
        id_column = st.selectbox(
            "Site ID Column *",
            all_cols_v1,
            index=default_id_idx,
            help="The column used to join rows between V1 and V2. "
                 "Must contain unique identifiers (site codes, facility IDs, etc.)",
        )
    with c2:
        materiality_threshold = st.number_input(
            "Materiality Threshold (%)",
            min_value=0.0,
            max_value=100.0,
            value=5.0,
            step=0.5,
            help="Changes below this % are flagged 🟢 Minor. Above = 🔴 High Impact.",
        )
    # ── Threshold preview ─────────────────────────────────────────────────
    try:
        _prev = run_diff(df_v1, df_v2, id_column,
                         match.to_serializable(), materiality_threshold)
        p1, p2, p3 = st.columns(3)
        p1.metric("🔴 High Impact", len(_prev.high_impact_sites),
                  help=f"Changes above {materiality_threshold:.0f}%")
        p2.metric("🟡 Possible Error", len(_prev.warn_sites))
        p3.metric("🟢 Minor", len(_prev.minor_sites),
                  help=f"Changes at or below {materiality_threshold:.0f}%")
    except Exception:
        pass
    # ── Navigation ────────────────────────────────────────────────────────
    c1, c2 = st.columns([1, 4])
    with c1:
        if st.button("← Back", use_container_width=True):
            _go(1)
            st.rerun()
    with c2:
        if st.button("Run Diff →", type="primary", use_container_width=True):
            column_mapping = match.to_serializable()
            if not column_mapping:
                st.error("No column pairs to diff. Please match at least one column.")
                return

            with st.spinner("Running diff…"):
                try:
                    diff_result = run_diff(
                        df_v1=df_v1,
                        df_v2=df_v2,
                        id_column=id_column,
                        column_mapping=column_mapping,
                        materiality_threshold=materiality_threshold,
                    )
                except Exception as exc:
                    st.error(f"Diff failed: {exc}")
                    return

            # Build annotation store (in-memory sidecar; writes to temp dir)
            run_date = datetime.now().strftime("%Y-%m-%d")
            tmp_dir = tempfile.mkdtemp(prefix="ghg_qa_")
            sidecar = sidecar_path(tmp_dir, run_date,
                                   setup["v1_label"], setup["v2_label"])
            structural = {
                "sites_added":     list(diff_result.structural.sites_added),
                "sites_removed":   list(diff_result.structural.sites_removed),
                "columns_added":   list(diff_result.structural.columns_added),
                "columns_removed": list(diff_result.structural.columns_removed),
            }
            store = AnnotationStore.load_or_create(
                sidecar,
                analyst=setup["analyst"],
                project=setup["project"],
                v1_path="[uploaded]",
                v1_label=setup["v1_label"],
                v2_path="[uploaded]",
                v2_label=setup["v2_label"],
                id_column=id_column,
                materiality_threshold=materiality_threshold,
                column_mapping=column_mapping,
                structural_changes=structural,
            )

            st.session_state["col_setup"] = {
                "id_column": id_column,
                "column_mapping": column_mapping,
                "materiality_threshold": materiality_threshold,
            }
            st.session_state["diff_result"] = diff_result
            st.session_state["store"] = store
            st.session_state["selected_site"] = None
            _go(3)
            st.rerun()


# ---------------------------------------------------------------------------
# Step 3 — Diff Results
# ---------------------------------------------------------------------------

def _severity_order(sd: SiteDiff) -> int:
    if sd.severity == SEVERITY_HIGH:
        return 0
    if sd.severity == SEVERITY_WARN:
        return 1
    return 2


def _site_rows(diff: DiffResult, filt: str) -> list[SiteDiff]:
    rows = list(diff.site_diffs)
    if filt == "🔴 High Impact":
        rows = [r for r in rows if r.severity == SEVERITY_HIGH]
    elif filt == "🟡 Possible Error":
        rows = [r for r in rows if r.severity == SEVERITY_WARN]
    elif filt == "🟢 Minor":
        rows = [r for r in rows if r.severity == SEVERITY_MINOR]
    elif filt == "⬜ Unchanged":
        rows = [r for r in rows if not r.column_diffs]
    return sorted(rows, key=_severity_order)


def _disposition_badge(disposition: str | None) -> str:
    icons = {
        AnnotationStore.DISPOSITION_EXPECTED: "✅",
        AnnotationStore.DISPOSITION_ERROR: "🔧",
        AnnotationStore.DISPOSITION_REVIEW: "❓",
    }
    return icons.get(disposition, "")


def _step3_diff_results() -> None:
    setup: dict = st.session_state["setup"]
    diff: DiffResult = st.session_state["diff_result"]
    store: AnnotationStore = st.session_state["store"]
    v1_label = setup["v1_label"]
    v2_label = setup["v2_label"]

    st.header("🔍 Diff Results")

    # ── Summary banner ────────────────────────────────────────────────────
    summary = diff.summary_sentence()
    st.info(f"**{summary}**")

    # ── Structural changes ────────────────────────────────────────────────
    s = diff.structural
    if s.sites_added or s.sites_removed or s.columns_added or s.columns_removed:
        with st.expander("🏗 Structural Changes"):
            c1, c2 = st.columns(2)
            with c1:
                if s.sites_added:
                    st.markdown(f"**Sites added ({len(s.sites_added)})**")
                    for x in sorted(s.sites_added):
                        st.markdown(f"- {x}")
                if s.columns_added:
                    st.markdown(f"**Columns added ({len(s.columns_added)})**")
                    for x in sorted(s.columns_added):
                        st.markdown(f"- `{x}`")
            with c2:
                if s.sites_removed:
                    st.markdown(f"**Sites removed ({len(s.sites_removed)})**")
                    for x in sorted(s.sites_removed):
                        st.markdown(f"- {x}")
                if s.columns_removed:
                    st.markdown(f"**Columns removed ({len(s.columns_removed)})**")
                    for x in sorted(s.columns_removed):
                        st.markdown(f"- `{x}`")

    # ── Filter ────────────────────────────────────────────────────────────
    filt = st.radio(
        "Filter",
        ["All", "🔴 High Impact", "🟡 Possible Error", "🟢 Minor", "⬜ Unchanged"],
        horizontal=True,
        key="filter_severity",
        label_visibility="collapsed",
    )

    visible_sites = _site_rows(diff, filt)

    if not visible_sites:
        st.info("No sites match the current filter.")
    else:
        # ── Site table ────────────────────────────────────────────────────
        site_data = []
        for sd in visible_sites:
            # Per-column badge: reflect worst disposition across all diffs
            col_disps = [
                (store.get_annotation(sd.site_id, cd.column_v1) or {}).get("disposition")
                for cd in sd.column_diffs
            ]
            if any(d == AnnotationStore.DISPOSITION_ERROR for d in col_disps):
                badge = "🔧"
            elif any(d == AnnotationStore.DISPOSITION_REVIEW for d in col_disps):
                badge = "❓"
            elif col_disps and all(d == AnnotationStore.DISPOSITION_EXPECTED for d in col_disps):
                badge = "✅"
            elif any(d is not None for d in col_disps):
                badge = "···"  # partially annotated
            else:
                badge = ""
            max_pct = max((abs(d.delta_pct) for d in sd.column_diffs if d.delta_pct is not None), default=None)
            pct_str = f"{max_pct:+.1f}%" if max_pct is not None else "—"
            site_data.append({
                "": SEVERITY_ICON.get(sd.severity, ""), # icon
                "Site": sd.site_id,
                "Max Δ%": pct_str,
                "Changes": len(sd.column_diffs),
                "Status": badge,
            })

        df_sites = pd.DataFrame(site_data)
        selected = st.dataframe(
            df_sites,
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="site_table",
        )

        # ── Detail panel ──────────────────────────────────────────────────
        sel_rows = selected.selection.rows if selected.selection else []
        if sel_rows:
            idx = sel_rows[0]
            sd: SiteDiff = visible_sites[idx]
            st.session_state["selected_site"] = sd.site_id

            st.divider()
            st.subheader(f"Site: {sd.site_id}")

            if not sd.column_diffs:
                st.success("No column differences found for this site.")
            else:
                for cd in sd.column_diffs:
                    sev_icon = SEVERITY_ICON.get(cd.severity, "")
                    pct_str = (f"{cd.delta_pct:+.1f}%" if cd.delta_pct is not None
                               else "n/a")
                    with st.container(border=True):
                        h1, h2, h3 = st.columns([3, 2, 2])
                        with h1:
                            st.markdown(f"{sev_icon} **{cd.column_v1}**")
                        with h2:
                            st.markdown(f"`{cd.v1_value}` → `{cd.v2_value}`")
                        with h3:
                            st.markdown(f"Δ {pct_str}")

                # ── Annotation ────────────────────────────────────────────
                st.markdown("**Annotate this site**")
                ann = store.get_annotation(sd.site_id, "__all__") or {}
                current_disp = ann.get("disposition")
                current_note = ann.get("note", "")

                disp_options = [
                    AnnotationStore.DISPOSITION_EXPECTED,
                    AnnotationStore.DISPOSITION_ERROR,
                    AnnotationStore.DISPOSITION_REVIEW,
                ]
                disp_labels = {
                    AnnotationStore.DISPOSITION_EXPECTED: "✅ Expected",
                    AnnotationStore.DISPOSITION_ERROR: "🔧 Error",
                    AnnotationStore.DISPOSITION_REVIEW: "❓ Needs Review",
                }
                c1, c2 = st.columns([2, 3])
                with c1:
                    disp_choice = st.radio(
                        "Disposition",
                        disp_options,
                        format_func=lambda x: disp_labels[x],
                        index=(disp_options.index(current_disp)
                               if current_disp in disp_options else 0),
                        key=f"disp_{sd.site_id}",
                    )
                with c2:
                    note_text = st.text_area(
                        "Note",
                        value=current_note,
                        key=f"note_{sd.site_id}",
                        placeholder="Explain the change…",
                        height=100,
                    )

                if st.button("💾 Save Annotation", key=f"save_{sd.site_id}",
                             type="primary"):
                    store.set_annotation(sd.site_id, "__all__",
                                         disp_choice, note_text)
                    st.success("Annotation saved.")
                    st.rerun()

    st.divider()

    # ── Export ────────────────────────────────────────────────────────────
    st.subheader("Export QA Log")

    # Check for unanswered HIGH column diffs
    high_keys = [
        f"{sd.site_id}::{cd.column_v1}"
        for sd in diff.site_diffs
        for cd in sd.column_diffs
        if cd.severity == SEVERITY_HIGH
    ]
    unanswered = store.unanswered_high_impact_keys(high_keys)
    if unanswered:
        st.warning(
            f"⚠️ {len(unanswered)} High Impact site(s) have no annotation yet. "
            "You can still export, but the QA Log will flag them."
        )

    if st.button("📥 Generate & Download QA Log", type="primary"):
        with st.spinner("Generating QA Log…"):
            buf = io.BytesIO()
            write_qa_log(buf, diff, store, setup["project"],
                         v1_label, v2_label)
            buf.seek(0)

        fname = (
            f"QA_Log_{setup['project'].replace(' ', '_')}_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        st.download_button(
            label="⬇️ Download QA Log Excel",
            data=buf,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    c1, c2 = st.columns([1, 4])
    with c1:
        if st.button("← Back to Column Setup"):
            _go(2)
            st.rerun()
    with c2:
        if st.button("🔄 Start New Comparison"):
            _reset()


# ---------------------------------------------------------------------------
# App shell
# ---------------------------------------------------------------------------

def main() -> None:
    _init_state()
    _nav_bar()

    step = st.session_state["step"]
    if step == 1:
        _step1_project_setup()
    elif step == 2:
        _step2_column_matching()
    elif step == 3:
        _step3_diff_results()


main()
