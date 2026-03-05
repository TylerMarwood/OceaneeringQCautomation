"""
Oceaneering QC Comparison Tool
================================
Streamlit application that compares a client-supplied Excel sheet against an
INFORM spreadsheet, performing:
  1. Tag column validation (duplicates, missing tags)
  2. Header matching (exact string)
  3. Cell-level data comparison (with type normalisation)

Run with:
    python -m streamlit run app.py
"""

from __future__ import annotations

import traceback

import pandas as pd
import streamlit as st

from utils.comparison import run_full_comparison
from utils.parser import get_sheet_names, parse_client_sheet, parse_inform_sheet, preview_rows
from utils.report import generate_comparison_report, generate_tag_issues_report

# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Oceaneering QC Tool",
    page_icon=":mag:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# CSS  (Font Awesome 6 loaded via @import)
# ---------------------------------------------------------------------------

st.markdown(
    """
    <style>
    @import url('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css');

    /* ---- Global ---- */
    [data-testid="stAppViewContainer"] { background-color: #f5f7fa; }
    h1 { color: #1F4E79; }
    h2 { color: #2E75B6; border-bottom: 2px solid #2E75B6; padding-bottom: 4px; }
    h3 { color: #1F4E79; }

    /* ---- Cards ---- */
    .qc-card {
        background: white;
        border-radius: 8px;
        padding: 20px 24px;
        margin-bottom: 16px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        border-left: 4px solid #2E75B6;
        color: #1a1a1a;
    }
    .qc-card-red   { border-left-color: #dc3545; }
    .qc-card-green { border-left-color: #28a745; }
    .qc-card-amber { border-left-color: #ffc107; }

    /* ---- Instruction callouts ---- */
    .instruction-box {
        background: #eef4fb;
        border: 1px solid #b8d0ea;
        border-left: 4px solid #2E75B6;
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 18px;
        color: #1a2a3a;
        font-size: 14px;
        line-height: 1.6;
    }
    .instruction-box .inst-title {
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 6px;
        font-size: 14px;
    }
    .instruction-box ul {
        margin: 6px 0 0 0;
        padding-left: 20px;
    }
    .instruction-box li { margin-bottom: 3px; }

    /* ---- Header chips ---- */
    .chip-green {
        display: inline-block;
        background: #d4edda; color: #155724;
        border: 1px solid #c3e6cb;
        border-radius: 16px; padding: 3px 12px;
        font-size: 13px; font-weight: 600;
        margin: 3px 4px;
    }
    .chip-red {
        display: inline-block;
        background: #f8d7da; color: #721c24;
        border: 1px solid #f5c6cb;
        border-radius: 16px; padding: 3px 12px;
        font-size: 13px; font-weight: 600;
        margin: 3px 4px;
    }
    .chip-amber {
        display: inline-block;
        background: #fff3cd; color: #5a4000;
        border: 1px solid #ffeeba;
        border-radius: 16px; padding: 3px 12px;
        font-size: 13px; font-weight: 600;
        margin: 3px 4px;
    }

    /* ---- Stat boxes ---- */
    .stat-row { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 16px; }
    .stat-box {
        flex: 1 1 120px;
        background: white;
        border-radius: 8px;
        padding: 14px 18px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .stat-number { font-size: 2em; font-weight: 700; line-height: 1.1; }
    .stat-label  { font-size: 12px; color: #444; margin-top: 4px; }
    .stat-green  .stat-number { color: #28a745; }
    .stat-red    .stat-number { color: #dc3545; }
    .stat-blue   .stat-number { color: #2E75B6; }
    .stat-amber  .stat-number { color: #7a5800; }

    /* ---- Step badges ---- */
    .step-badge {
        background: #2E75B6; color: white;
        border-radius: 50%; width: 28px; height: 28px;
        display: inline-flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 14px; margin-right: 8px;
    }

    /* ---- Status labels (replaces colour-dot emojis) ---- */
    .status-label {
        display: inline-block;
        font-size: 12px;
        font-weight: 600;
        padding: 2px 8px;
        border-radius: 4px;
        margin-right: 6px;
    }
    .status-label-green  { background: #d4edda; color: #155724; }
    .status-label-red    { background: #f8d7da; color: #721c24; }
    .status-label-amber  { background: #fff3cd; color: #5a4000; }
    .status-label-blue   { background: #d1ecf1; color: #0c5460; }

    /* ---- Dividers ---- */
    .section-divider {
        border: none; border-top: 2px solid #e0e7ef;
        margin: 24px 0;
    }

    /* ---- Upload zones ---- */
    [data-testid="stFileUploader"] {
        border: 2px dashed #2E75B6 !important;
        border-radius: 8px !important;
    }

    /* ---- Footer ---- */
    .qc-footer {
        text-align: center;
        color: #555;
        font-size: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Session state initialisation
# ---------------------------------------------------------------------------

def _init_state() -> None:
    defaults = {
        "client_bytes":      None,
        "client_filename":   None,
        "inform_bytes":      None,
        "inform_filename":   None,
        "client_df":         None,
        "inform_df":         None,
        "client_headers":    [],
        "inform_headers":    [],
        "comparison_result": None,
        "comparison_ran":    False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _safe_read(uploaded_file) -> bytes | None:
    try:
        return uploaded_file.read()
    except Exception:
        return None


def _chip(label: str, colour: str) -> str:
    return f'<span class="chip-{colour}">{label}</span>'


def _stat_box(number, label: str, colour: str) -> str:
    return (
        f'<div class="stat-box stat-{colour}">'
        f'<div class="stat-number">{number}</div>'
        f'<div class="stat-label">{label}</div>'
        f'</div>'
    )


def _instruction(title: str, body: str) -> None:
    """Render a styled instruction/explanation callout."""
    st.markdown(
        f'<div class="instruction-box">'
        f'<div class="inst-title"><i class="fas fa-circle-info"></i>&nbsp;&nbsp;{title}</div>'
        f'{body}'
        f'</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# App header
# ---------------------------------------------------------------------------

st.markdown(
    """
    <div style="background: linear-gradient(135deg,#1F4E79,#2E75B6);
                border-radius:10px; padding:24px 32px; margin-bottom:24px;">
      <h1 style="color:white; margin:0; font-size:1.8em;">
        <i class="fas fa-magnifying-glass-chart" style="margin-right:12px;"></i>
        Oceaneering QC Comparison Tool
      </h1>
      <p style="color:#c8ddf0; margin:8px 0 0; font-size:14px; line-height:1.6;">
        Upload a client Excel sheet and an INFORM spreadsheet to validate tags,
        match column headers, and compare data cell-by-cell. Work through each
        numbered step in sequence to produce a full QC report.
      </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ===========================================================================
# SECTION 1 — File upload  (always rendered)
# ===========================================================================

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
st.markdown(
    '<h2><span class="step-badge">1</span>Upload Files</h2>',
    unsafe_allow_html=True,
)

_instruction(
    "How to use this section",
    """<ul>
      <li><strong>Client Sheet</strong> — the Excel file supplied by the client containing instrument or equipment data to be quality-checked.</li>
      <li><strong>INFORM Sheet</strong> — the reference spreadsheet exported from the INFORM system. This is the source of truth against which the client data is compared.</li>
      <li>Both files must be in <strong>.xlsx</strong> or <strong>.xls</strong> format. Once both files are loaded, the next steps will become available below.</li>
    </ul>""",
)

col_up_client, col_up_inform = st.columns(2, gap="large")

with col_up_client:
    st.markdown("#### Client Sheet")
    client_file = st.file_uploader(
        "Upload client Excel file",
        type=["xlsx", "xls"],
        key="upload_client",
        label_visibility="collapsed",
    )
    if client_file:
        if st.session_state.client_filename != client_file.name:
            st.session_state.client_bytes    = _safe_read(client_file)
            st.session_state.client_filename = client_file.name
            st.session_state.client_df         = None
            st.session_state.comparison_result = None
            st.session_state.comparison_ran    = False
        st.success(f"Loaded: **{client_file.name}**")

with col_up_inform:
    st.markdown("#### INFORM Sheet")
    inform_file = st.file_uploader(
        "Upload INFORM Excel file",
        type=["xlsx", "xls"],
        key="upload_inform",
        label_visibility="collapsed",
    )
    if inform_file:
        if st.session_state.inform_filename != inform_file.name:
            st.session_state.inform_bytes    = _safe_read(inform_file)
            st.session_state.inform_filename = inform_file.name
            st.session_state.inform_df         = None
            st.session_state.comparison_result = None
            st.session_state.comparison_ran    = False
        st.success(f"Loaded: **{inform_file.name}**")

# ---------------------------------------------------------------------------
# Everything below this line only renders when both files are present
# ---------------------------------------------------------------------------

_files_ready = bool(
    st.session_state.client_bytes and st.session_state.inform_bytes
)

if not _files_ready:
    st.info("Upload both files above to continue.", icon=None)

else:
    # =======================================================================
    # SECTION 2 — Sheet selection & configuration
    # =======================================================================

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown(
        '<h2><span class="step-badge">2</span>Configure Sheets</h2>',
        unsafe_allow_html=True,
    )

    _instruction(
        "How to use this section",
        """<ul>
          <li><strong>Sheet tab</strong> — select which worksheet within the uploaded file should be used for comparison. Most files contain only one sheet; if yours has multiple, choose the correct one from the dropdown.</li>
          <li><strong>Header row</strong> — the row number (counting from 1) that contains the column names. For example, if column names appear on row 3 of the spreadsheet, enter 3.</li>
          <li><strong>Data start row</strong> — the row number where the first actual data record begins. This must be greater than the header row.</li>
          <li>Use the <strong>Preview</strong> panel to visually verify your row selections before proceeding. The highlighted rows show which row has been identified as the header (green) and where data begins (blue).</li>
        </ul>""",
    )

    # Read sheet names — show errors inline and abort this section if failed
    client_sheets = None
    inform_sheets = None

    try:
        client_sheets = get_sheet_names(st.session_state.client_bytes)
    except Exception as exc:
        st.error(f"Could not read client file sheets: {exc}")

    try:
        inform_sheets = get_sheet_names(st.session_state.inform_bytes)
    except Exception as exc:
        st.error(f"Could not read INFORM file sheets: {exc}")

    if client_sheets is not None and inform_sheets is not None:

        col_cfg_client, col_cfg_inform = st.columns(2, gap="large")

        with col_cfg_client:
            st.markdown(
                '<div class="qc-card"><b>Client Sheet Configuration</b></div>',
                unsafe_allow_html=True,
            )

            client_sheet = st.selectbox(
                "Sheet to compare", client_sheets, key="sel_client_sheet"
            )

            c1, c2 = st.columns(2)
            with c1:
                client_header_row = st.number_input(
                    "Header row", min_value=1, max_value=500, value=1,
                    key="client_hrow",
                    help="Row number (1-indexed) that contains column headers.",
                )
            with c2:
                client_data_row = st.number_input(
                    "Data start row", min_value=2, max_value=500, value=2,
                    key="client_drow",
                    help="Row number (1-indexed) where data begins.",
                )

            if client_data_row <= client_header_row:
                st.warning("Data start row must be greater than header row.")

            with st.expander("Preview raw rows (first 8 rows from row 1)", expanded=False):
                try:
                    preview_df = preview_rows(
                        st.session_state.client_bytes, client_sheet, 1, 8
                    )
                    def _hl_client(row):
                        if row.name == client_header_row:
                            return ["background-color:#d4edda"] * len(row)
                        if row.name == client_data_row:
                            return ["background-color:#cce5ff"] * len(row)
                        return [""] * len(row)
                    st.dataframe(
                        preview_df.style.apply(_hl_client, axis=1),
                        use_container_width=True,
                        height=250,
                    )
                    st.markdown(
                        '<span class="status-label status-label-green">Green</span> Header row &nbsp;&nbsp;'
                        '<span class="status-label status-label-blue">Blue</span> Data start row',
                        unsafe_allow_html=True,
                    )
                except Exception as exc:
                    st.warning(f"Preview unavailable: {exc}")

        with col_cfg_inform:
            st.markdown(
                '<div class="qc-card"><b>INFORM Sheet Configuration</b></div>',
                unsafe_allow_html=True,
            )

            inform_sheet = st.selectbox(
                "Sheet to compare", inform_sheets, key="sel_inform_sheet"
            )

            st.info(
                "**INFORM fixed format detected.**  \n"
                "Row 1 — Column headers  |  Row 2 — Metadata (skipped)  |  "
                "Row 3 — Original / New column markers  |  Row 4 onward — Data records  \n\n"
                "Only **Original** columns will be extracted for comparison. "
                "No row configuration is required for the INFORM sheet.",
            )

            with st.expander("Preview raw rows (first 8 rows from row 1)", expanded=False):
                try:
                    preview_df = preview_rows(
                        st.session_state.inform_bytes, inform_sheet, 1, 8
                    )
                    st.dataframe(preview_df, use_container_width=True, height=250)
                    st.caption(
                        "Row 1 = column headers  |  Row 2 = metadata  |  "
                        "Row 3 = Original/New markers  |  Row 4 onward = data records"
                    )
                except Exception as exc:
                    st.warning(f"Preview unavailable: {exc}")

        # ===================================================================
        # SECTION 3 — Parse sheets & select tag columns
        # ===================================================================

        st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        st.markdown(
            '<h2><span class="step-badge">3</span>Select Tag Columns</h2>',
            unsafe_allow_html=True,
        )

        _instruction(
            "How to use this section",
            """<ul>
              <li>The <strong>tag column</strong> is the column that contains unique identifiers for each instrument or piece of equipment — for example, tag numbers such as <code>FT-1001</code> or <code>PT-2045</code>.</li>
              <li>Select the tag column from the <strong>Client sheet</strong> and the corresponding tag column from the <strong>INFORM sheet</strong>. These columns are used to align rows from both sheets before comparing their values.</li>
              <li>Rows whose tag values exist in only one sheet will be flagged as issues and excluded from the cell-level comparison. Only tags present in <em>both</em> sheets will be compared.</li>
              <li>Use the <strong>Preview</strong> panels below to confirm the parsed data looks correct before running the comparison.</li>
            </ul>""",
        )

        parse_error = False
        client_df = client_headers = inform_df = inform_headers = None

        try:
            client_df, client_headers = parse_client_sheet(
                st.session_state.client_bytes,
                client_sheet,
                int(client_header_row),
                int(client_data_row),
            )
            st.session_state.client_df      = client_df
            st.session_state.client_headers = client_headers
        except Exception as exc:
            st.error(f"Error parsing client sheet: {exc}")
            with st.expander("Traceback"):
                st.code(traceback.format_exc())
            parse_error = True

        try:
            inform_df, inform_headers = parse_inform_sheet(
                st.session_state.inform_bytes,
                inform_sheet,
            )
            st.session_state.inform_df      = inform_df
            st.session_state.inform_headers = inform_headers
        except Exception as exc:
            st.error(f"Error parsing INFORM sheet: {exc}")
            with st.expander("Traceback"):
                st.code(traceback.format_exc())
            parse_error = True

        if parse_error:
            st.error("Fix the parsing errors above before continuing.")

        else:
            col_tag_client, col_tag_inform = st.columns(2, gap="large")

            with col_tag_client:
                client_tag_col = st.selectbox(
                    "Tag column — Client sheet",
                    client_headers,
                    key="client_tag_col",
                    help="The column whose values uniquely identify each instrument or tag. Used to match rows between the two sheets.",
                )

            with col_tag_inform:
                inform_tag_col = st.selectbox(
                    "Tag column — INFORM sheet (Original columns only)",
                    inform_headers,
                    key="inform_tag_col",
                    help="The corresponding tag column in the INFORM sheet. Must contain the same tag identifiers as the client sheet.",
                )

            col_prev_client, col_prev_inform = st.columns(2, gap="large")

            with col_prev_client:
                with st.expander(
                    f"Preview client data  ({len(client_df):,} rows x {len(client_headers)} columns)",
                    expanded=False,
                ):
                    st.dataframe(client_df.head(20), use_container_width=True, height=300)

            with col_prev_inform:
                with st.expander(
                    f"Preview INFORM data  ({len(inform_df):,} rows x {len(inform_headers)} columns)",
                    expanded=False,
                ):
                    st.dataframe(inform_df.head(20), use_container_width=True, height=300)

            # ===============================================================
            # SECTION 4 — Run comparison
            # ===============================================================

            st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
            st.markdown(
                '<h2><span class="step-badge">4</span>Run Comparison</h2>',
                unsafe_allow_html=True,
            )

            _instruction(
                "What the comparison does",
                """<ul>
                  <li><strong>Tag validation</strong> — checks that all tag values are unique within each sheet and that every tag is present in both sheets. Duplicates and one-sided tags are reported as issues.</li>
                  <li><strong>Header matching</strong> — compares column names between the two sheets. Only columns with an exact name match in both sheets will be included in the cell-level comparison.</li>
                  <li><strong>Cell-level comparison</strong> — for every row/column combination where the tag and header match, the tool compares the values from both sheets. Values are normalised (whitespace trimmed, numeric types reconciled) before comparison to reduce false positives.</li>
                  <li>Results are displayed in the tabs that appear below. You can also download formatted Excel reports for distribution.</li>
                </ul>""",
            )

            run_btn = st.button(
                "Run QC Comparison",
                type="primary",
                use_container_width=False,
            )

            if run_btn:
                st.session_state.comparison_ran    = False
                st.session_state.comparison_result = None

                with st.spinner("Running comparison — please wait ..."):
                    try:
                        _result = run_full_comparison(
                            client_df      = client_df,
                            inform_df      = inform_df,
                            client_tag_col = client_tag_col,
                            inform_tag_col = inform_tag_col,
                            client_headers = client_headers,
                            inform_headers = inform_headers,
                        )
                        st.session_state.comparison_result = _result
                        st.session_state.comparison_ran    = True
                    except Exception as exc:
                        st.error(f"Comparison failed: {exc}")
                        with st.expander("Traceback"):
                            st.code(traceback.format_exc())

            # ===============================================================
            # SECTION 5 — Results  (only when a comparison has been run)
            # ===============================================================

            if st.session_state.comparison_ran and st.session_state.comparison_result is not None:

                result = st.session_state.comparison_result
                tv     = result.tag_validation
                hm     = result.header_match
                df_all = result.diffs_df()

                st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
                st.markdown(
                    '<h2><span class="step-badge">5</span>Results</h2>',
                    unsafe_allow_html=True,
                )

                total_cells = len(df_all)
                match_cells = int(result.match_count)
                mm_cells    = int(result.mismatch_count)
                rate_pct    = f"{match_cells / total_cells * 100:.1f}%" if total_cells else "—"

                st.markdown(
                    f"""
                    <div class="stat-row">
                      {_stat_box(len(tv.comparable_tags), "Comparable Tags",   "blue")}
                      {_stat_box(len(tv.issues),          "Tag Issues",        "red" if tv.issues else "green")}
                      {_stat_box(len(hm.matched),         "Matched Headers",   "blue")}
                      {_stat_box(len(hm.client_only) + len(hm.inform_only), "Unmatched Headers",
                                 "red" if (hm.client_only or hm.inform_only) else "green")}
                      {_stat_box(total_cells,             "Cell Comparisons",  "blue")}
                      {_stat_box(match_cells,             "Cells Match",       "green")}
                      {_stat_box(mm_cells,                "Cells Differ",      "red" if mm_cells else "green")}
                      {_stat_box(rate_pct,               "Match Rate",        "green" if mm_cells == 0 else "amber")}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                tab_tags, tab_headers, tab_diff, tab_downloads = st.tabs([
                    "Tag Validation",
                    "Header Matching",
                    "Cell Comparison",
                    "Download Reports",
                ])

                # -----------------------------------------------------------
                # TAB 1 — Tag validation
                # -----------------------------------------------------------

                with tab_tags:

                    _instruction(
                        "About Tag Validation",
                        """<ul>
                          <li><strong>Comparable tags</strong> — tag values found in both sheets. Only these tags are included in the cell-level comparison.</li>
                          <li><strong>Client-only tags</strong> — tags present in the client sheet but absent from INFORM. These rows cannot be compared and are excluded from cell-level results.</li>
                          <li><strong>INFORM-only tags</strong> — tags present in INFORM but absent from the client sheet. Also excluded from cell-level results.</li>
                          <li><strong>Duplicate tags</strong> — tag values that appear more than once within a single sheet. These are flagged because row matching would be ambiguous.</li>
                        </ul>""",
                    )

                    tv_col1, tv_col2, tv_col3 = st.columns(3)
                    with tv_col1:
                        st.metric("Comparable tags", len(tv.comparable_tags))
                    with tv_col2:
                        st.metric(
                            "Client-only tags", len(tv.client_only_tags),
                            delta=f"{len(tv.client_only_tags)} not compared"
                                  if tv.client_only_tags else None,
                            delta_color="off",
                        )
                    with tv_col3:
                        st.metric(
                            "INFORM-only tags", len(tv.inform_only_tags),
                            delta=f"{len(tv.inform_only_tags)} not compared"
                                  if tv.inform_only_tags else None,
                            delta_color="off",
                        )

                    if not tv.issues:
                        st.success(
                            "All tags are unique and present in both sheets. "
                            "Full comparison is available.",
                        )
                    else:
                        dup_issues  = [i for i in tv.issues if "Duplicate" in i.reason]
                        miss_issues = [i for i in tv.issues
                                       if "Duplicate" not in i.reason and "(empty)" not in i.tag]
                        null_issues = [i for i in tv.issues if "(empty)" in i.tag]

                        if null_issues:
                            with st.expander(
                                f"Empty / Null Tags — {len(null_issues)} issue(s)",
                                expanded=True,
                            ):
                                null_df = pd.DataFrame(
                                    [{"Sheet": i.sheet, "Reason": i.reason} for i in null_issues]
                                )
                                st.dataframe(null_df, use_container_width=True, hide_index=True)

                        if dup_issues:
                            with st.expander(
                                f"Duplicate Tags — {len(dup_issues)} tag(s)", expanded=True
                            ):
                                dup_df = pd.DataFrame(
                                    [{"Tag": i.tag, "Sheet": i.sheet, "Reason": i.reason}
                                     for i in dup_issues]
                                )
                                st.dataframe(
                                    dup_df.style.apply(
                                        lambda _: ["background-color:#ffc7ce"] * len(dup_df.columns),
                                        axis=1,
                                    ),
                                    use_container_width=True,
                                    hide_index=True,
                                )

                        if miss_issues:
                            client_miss = [i for i in miss_issues if i.sheet == "Client"]
                            inform_miss = [i for i in miss_issues if i.sheet == "INFORM"]

                            if client_miss:
                                with st.expander(
                                    f"Client-Only Tags — {len(client_miss)} tag(s), no INFORM comparison",
                                    expanded=len(client_miss) <= 50,
                                ):
                                    cm_df = pd.DataFrame(
                                        [{"Tag": i.tag, "Reason": i.reason} for i in client_miss]
                                    )
                                    st.dataframe(
                                        cm_df.style.apply(
                                            lambda _: ["background-color:#fff3cd"] * len(cm_df.columns),
                                            axis=1,
                                        ),
                                        use_container_width=True,
                                        hide_index=True,
                                    )

                            if inform_miss:
                                with st.expander(
                                    f"INFORM-Only Tags — {len(inform_miss)} tag(s), no client comparison",
                                    expanded=len(inform_miss) <= 50,
                                ):
                                    im_df = pd.DataFrame(
                                        [{"Tag": i.tag, "Reason": i.reason} for i in inform_miss]
                                    )
                                    st.dataframe(
                                        im_df.style.apply(
                                            lambda _: ["background-color:#fff3cd"] * len(im_df.columns),
                                            axis=1,
                                        ),
                                        use_container_width=True,
                                        hide_index=True,
                                    )

                        if tv.comparable_tags:
                            st.info(
                                f"{len(tv.comparable_tags):,} tag(s) are comparable and have "
                                f"been included in the data comparison below.",
                            )
                        else:
                            st.error(
                                "No comparable tags found. Check that the correct tag columns "
                                "are selected for both sheets.",
                            )

                # -----------------------------------------------------------
                # TAB 2 — Header matching
                # -----------------------------------------------------------

                with tab_headers:

                    _instruction(
                        "About Header Matching",
                        """<ul>
                          <li>Column headers are compared using an <strong>exact string match</strong>. Column names must be identical in both sheets (including spacing and capitalisation) to be included in cell-level comparison.</li>
                          <li><strong>Matched headers</strong> — columns present in both sheets that will be compared cell-by-cell.</li>
                          <li><strong>Client-only headers</strong> — columns present in the client sheet but not in INFORM. Their data will not be compared.</li>
                          <li><strong>INFORM-only headers</strong> — columns present in INFORM but not in the client sheet. Their data will not be compared.</li>
                        </ul>""",
                    )

                    h_col1, h_col2, h_col3 = st.columns(3)
                    with h_col1:
                        st.metric("Matched headers", len(hm.matched))
                    with h_col2:
                        st.metric("Client-only headers", len(hm.client_only))
                    with h_col3:
                        st.metric("INFORM-only headers", len(hm.inform_only))

                    if hm.all_matched:
                        st.success("All headers match exactly between the two sheets.")
                    else:
                        if hm.client_only:
                            st.warning(
                                f"**{len(hm.client_only)} header(s)** exist in the client sheet "
                                f"but have no matching INFORM column — these will not be compared."
                            )
                        if hm.inform_only:
                            st.warning(
                                f"**{len(hm.inform_only)} header(s)** exist in the INFORM sheet "
                                f"but have no matching client column — these will not be compared."
                            )

                    st.markdown("#### All Headers at a Glance")

                    hg_col1, hg_col2, hg_col3 = st.columns(3)

                    with hg_col1:
                        st.markdown(
                            '<span class="status-label status-label-green">Matched</span> '
                            '<strong>Present in both sheets</strong>',
                            unsafe_allow_html=True,
                        )
                        chips = (
                            "".join(_chip(h, "green") for h in hm.matched)
                            if hm.matched else "<i>None</i>"
                        )
                        st.markdown(chips, unsafe_allow_html=True)

                    with hg_col2:
                        st.markdown(
                            '<span class="status-label status-label-amber">Client only</span> '
                            '<strong>Not in INFORM</strong>',
                            unsafe_allow_html=True,
                        )
                        chips = (
                            "".join(_chip(h, "amber") for h in hm.client_only)
                            if hm.client_only else "<i>None</i>"
                        )
                        st.markdown(chips, unsafe_allow_html=True)

                    with hg_col3:
                        st.markdown(
                            '<span class="status-label status-label-amber">INFORM only</span> '
                            '<strong>Not in client sheet</strong>',
                            unsafe_allow_html=True,
                        )
                        chips = (
                            "".join(_chip(h, "amber") for h in hm.inform_only)
                            if hm.inform_only else "<i>None</i>"
                        )
                        st.markdown(chips, unsafe_allow_html=True)

                    with st.expander("View as table", expanded=False):
                        max_len = max(len(hm.matched), len(hm.client_only), len(hm.inform_only), 1)
                        hdr_table = pd.DataFrame({
                            "Matched Headers":     hm.matched     + [""] * (max_len - len(hm.matched)),
                            "Client-Only Headers": hm.client_only + [""] * (max_len - len(hm.client_only)),
                            "INFORM-Only Headers": hm.inform_only + [""] * (max_len - len(hm.inform_only)),
                        })

                        def _colour_header_table(col: pd.Series) -> list[str]:
                            if col.name == "Matched Headers":
                                return ["background-color:#c6efce" if v else "" for v in col]
                            return ["background-color:#ffeb9c" if v else "" for v in col]

                        st.dataframe(
                            hdr_table.style.apply(_colour_header_table, axis=0),
                            use_container_width=True,
                            hide_index=True,
                        )

                # -----------------------------------------------------------
                # TAB 3 — Cell comparison
                # -----------------------------------------------------------

                with tab_diff:

                    _instruction(
                        "About Cell Comparison",
                        """<ul>
                          <li>Each row in this table represents one cell compared between the two sheets, identified by its <strong>Tag</strong> (row identifier) and <strong>Header</strong> (column identifier).</li>
                          <li><strong>Client Value / INFORM Value</strong> — the raw values as they appear in each file.</li>
                          <li><strong>Client (Normalised) / INFORM (Normalised)</strong> — values after normalisation (whitespace trimmed, numbers standardised). Comparison is performed on normalised values to reduce false mismatches caused by formatting differences.</li>
                          <li>Use the <strong>filter controls</strong> to focus on mismatches, a specific column header, or a specific tag. The bar chart at the bottom shows which columns have the most discrepancies.</li>
                        </ul>""",
                    )

                    if df_all.empty:
                        st.info(
                            "No cell-level comparison data. This can happen if no tags are "
                            "comparable or no headers matched between the sheets."
                        )
                    else:
                        mm_df = df_all[~df_all["Match"]]

                        diff_col1, diff_col2, diff_col3 = st.columns(3)
                        with diff_col1:
                            st.metric("Total comparisons", f"{total_cells:,}")
                        with diff_col2:
                            st.metric("Matches", f"{match_cells:,}", delta=None, delta_color="off")
                        with diff_col3:
                            st.metric("Mismatches", f"{mm_cells:,}", delta=None, delta_color="off")

                        filter_col1, filter_col2, filter_col3 = st.columns(3)
                        with filter_col1:
                            show_mode = st.radio(
                                "Show",
                                ["All results", "Mismatches only", "Matches only"],
                                horizontal=True,
                                key="show_mode",
                            )
                        with filter_col2:
                            header_filter = st.multiselect(
                                "Filter by Header",
                                options=sorted(df_all["Header"].unique()),
                                default=[],
                                key="header_filter",
                            )
                        with filter_col3:
                            tag_search = st.text_input(
                                "Search by Tag", value="", key="tag_search",
                                placeholder="e.g. FT-1001",
                            )

                        view_df = df_all.copy()
                        if show_mode == "Mismatches only":
                            view_df = view_df[~view_df["Match"]]
                        elif show_mode == "Matches only":
                            view_df = view_df[view_df["Match"]]
                        if header_filter:
                            view_df = view_df[view_df["Header"].isin(header_filter)]
                        if tag_search.strip():
                            view_df = view_df[
                                view_df["Tag"].str.contains(
                                    tag_search.strip(), case=False, na=False
                                )
                            ]

                        st.caption(f"Showing {len(view_df):,} of {total_cells:,} comparisons")

                        if view_df.empty:
                            st.info("No rows match the current filters.")
                        else:
                            display_cols = [
                                "Tag", "Header",
                                "Client Value", "INFORM Value",
                                "Client (Normalised)", "INFORM (Normalised)",
                                "Match",
                            ]

                            def _style_diff_table(df: pd.DataFrame):
                                def row_style(row):
                                    colour = "#c6efce" if row["Match"] else "#ffc7ce"
                                    return [f"background-color:{colour}"] * len(row)
                                return df.style.apply(row_style, axis=1)

                            st.dataframe(
                                _style_diff_table(view_df[display_cols]),
                                use_container_width=True,
                                height=500,
                                hide_index=True,
                            )

                        if not mm_df.empty:
                            st.markdown("#### Mismatch Count by Header")
                            st.caption(
                                "The chart below shows how many mismatches were found in each "
                                "column. Columns with the highest mismatch count appear first."
                            )
                            mm_by_hdr = (
                                mm_df.groupby("Header")
                                .size()
                                .reset_index(name="Mismatch Count")
                                .sort_values("Mismatch Count", ascending=False)
                            )
                            st.bar_chart(mm_by_hdr.set_index("Header")["Mismatch Count"])

                # -----------------------------------------------------------
                # TAB 4 — Downloads
                # -----------------------------------------------------------

                with tab_downloads:

                    _instruction(
                        "About the downloadable reports",
                        """<ul>
                          <li><strong>Tag Issues Report</strong> — a standalone Excel file listing all tag-level problems: duplicates, tags present in only one sheet, and empty tag values. Use this to identify data quality issues before re-running the comparison.</li>
                          <li><strong>Full Comparison Report</strong> — a multi-sheet Excel workbook containing: a Summary, Tag Issues, Header Info, Mismatches Only, and All Results. Cells are colour-coded green (match) and red (mismatch) for fast visual review. Share this report with the relevant team for sign-off.</li>
                          <li><strong>Raw CSV</strong> — the complete comparison dataset in CSV format, suitable for further analysis in Excel, Power BI, or other tools.</li>
                        </ul>""",
                    )

                    st.markdown("### Export Results")

                    dl_col1, dl_col2 = st.columns(2)

                    with dl_col1:
                        st.markdown("#### Tag Issues Report")
                        st.markdown(
                            "Lists all duplicate tags, tags present in only one sheet, "
                            "and any empty tag values found during validation."
                        )
                        try:
                            issues_xlsx = generate_tag_issues_report(tv.issues_df())
                            st.download_button(
                                label="Download Tag Issues (.xlsx)",
                                data=issues_xlsx,
                                file_name="QC_Tag_Issues.xlsx",
                                mime=(
                                    "application/vnd.openxmlformats-officedocument"
                                    ".spreadsheetml.sheet"
                                ),
                                use_container_width=True,
                            )
                        except Exception as exc:
                            st.error(f"Could not generate tag issues report: {exc}")

                    with dl_col2:
                        st.markdown("#### Full Comparison Report")
                        st.markdown(
                            "Multi-sheet workbook: Summary, Tag Issues, Header Info, "
                            "Mismatches Only, and All Results with green/red cell colouring."
                        )
                        try:
                            full_xlsx = generate_comparison_report(result)
                            st.download_button(
                                label="Download Full Report (.xlsx)",
                                data=full_xlsx,
                                file_name="QC_Full_Comparison_Report.xlsx",
                                mime=(
                                    "application/vnd.openxmlformats-officedocument"
                                    ".spreadsheetml.sheet"
                                ),
                                use_container_width=True,
                            )
                        except Exception as exc:
                            st.error(f"Could not generate full report: {exc}")

                    if not df_all.empty:
                        st.markdown("---")
                        st.markdown("#### Raw Comparison Data (CSV)")
                        st.markdown(
                            "Download the complete comparison dataset as a CSV file for use "
                            "in Excel, Power BI, or other analysis tools."
                        )
                        csv_data = df_all.to_csv(index=False).encode("utf-8")
                        st.download_button(
                            label="Download as CSV",
                            data=csv_data,
                            file_name="QC_Comparison_Data.csv",
                            mime="text/csv",
                        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
st.markdown(
    '<p class="qc-footer">Oceaneering QC Comparison Tool — Internal Use Only</p>',
    unsafe_allow_html=True,
)
