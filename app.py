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
# Global CSS — forces light-mode colours regardless of OS preference,
# fixes all text contrast issues, and styles section containers.
# Font Awesome 6 is loaded via @import for professional icons.
# ---------------------------------------------------------------------------

st.markdown(
    """
    <style>
    @import url('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css');

    /* ================================================================
       GLOBAL COLOUR RESET — override any dark-mode or theme defaults
       ================================================================ */
    html, body,
    [data-testid="stApp"],
    [data-testid="stAppViewContainer"],
    [data-testid="stMainBlocksContainer"],
    [data-testid="stMain"],
    [data-testid="block-container"] {
        background-color: #f2f4f8 !important;
        color: #1a1a1a !important;
    }

    /* All text elements — prevent white-on-white */
    p, span, div, li, td, th, label,
    [data-testid="stMarkdownContainer"],
    [data-testid="stMarkdownContainer"] * {
        color: #1a1a1a;
    }

    /* Streamlit headings */
    h1, h2, h3, h4,
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3 {
        color: #1F4E79 !important;
    }

    /* Metric labels and values */
    [data-testid="stMetricLabel"] > div,
    [data-testid="stMetricLabel"] p,
    [data-testid="stMetricValue"] > div,
    [data-testid="stMetricDelta"] > div {
        color: #1a1a1a !important;
    }

    /* Selectbox, number input, text input labels */
    [data-testid="stWidgetLabel"] p,
    .stSelectbox label,
    .stNumberInput label,
    .stTextInput label,
    .stRadio label,
    .stCheckbox label,
    .stMultiSelect label {
        color: #1a1a1a !important;
        font-weight: 600 !important;
    }

    /* Radio button option text */
    .stRadio [data-testid="stMarkdownContainer"] p {
        color: #1a1a1a !important;
    }

    /* Captions */
    [data-testid="stCaptionContainer"] p,
    .stCaption p {
        color: #444 !important;
    }

    /* Expander header text */
    [data-testid="stExpander"] summary span p,
    [data-testid="stExpander"] summary p {
        color: #1a1a1a !important;
    }

    /* Tab labels */
    [data-testid="stTabs"] [data-testid="stMarkdownContainer"] p,
    button[data-baseweb="tab"] p,
    button[data-baseweb="tab"] span {
        color: #1a1a1a !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] p,
    button[data-baseweb="tab"][aria-selected="true"] span {
        color: #1F4E79 !important;
        font-weight: 700 !important;
    }

    /* Info / warning / error / success boxes */
    [data-testid="stAlert"] p,
    [data-testid="stAlert"] span,
    [data-testid="stAlert"] li {
        color: #1a1a1a !important;
    }

    /* Dataframe text */
    [data-testid="stDataFrame"] * {
        color: #1a1a1a !important;
    }

    /* ================================================================
       SECTION CONTAINERS
       Bordered containers get a white background and a coloured
       top-accent stripe injected via .sect-header-* divs.
       ================================================================ */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #ffffff !important;
        border: 1px solid #c8d6e8 !important;
        border-radius: 10px !important;
        box-shadow: 0 2px 8px rgba(30, 60, 100, 0.07) !important;
        overflow: hidden !important;
    }

    /* Section header stripe — rendered as the first child of each container */
    .sect-header {
        margin: -16px -16px 16px -16px;
        padding: 12px 20px 11px 20px;
        border-radius: 0;
        display: flex;
        align-items: center;
        gap: 10px;
        font-size: 1.05em;
        font-weight: 700;
        letter-spacing: 0.01em;
    }
    .sect-header .step-badge {
        background: rgba(255,255,255,0.25);
        color: white;
        border-radius: 50%;
        width: 26px; height: 26px;
        display: inline-flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 13px;
        flex-shrink: 0;
    }
    .sect-header i { font-size: 1em; }

    /* Per-section stripe colours */
    .sect-upload   { background: #1F4E79; color: white !important; }
    .sect-config   { background: #245B8A; color: white !important; }
    .sect-tags     { background: #1a6b5c; color: white !important; }
    .sect-run      { background: #5a3e8a; color: white !important; }
    .sect-results  { background: #8a3e3e; color: white !important; }

    /* Ensure stripe text stays white */
    .sect-header,
    .sect-header p,
    .sect-header span,
    .sect-header i {
        color: white !important;
    }

    /* ================================================================
       INSTRUCTION CALLOUTS
       ================================================================ */
    .instruction-box {
        background: #eef4fb;
        border: 1px solid #b8d0ea;
        border-left: 4px solid #2E75B6;
        border-radius: 6px;
        padding: 14px 18px;
        margin-bottom: 16px;
        color: #1a2a3a !important;
        font-size: 14px;
        line-height: 1.6;
    }
    .instruction-box .inst-title {
        font-weight: 700;
        color: #1F4E79 !important;
        margin-bottom: 6px;
        font-size: 14px;
    }
    .instruction-box ul {
        margin: 6px 0 0 0;
        padding-left: 20px;
        color: #1a2a3a !important;
    }
    .instruction-box li { margin-bottom: 3px; color: #1a2a3a !important; }
    .instruction-box code {
        background: #dce8f5;
        color: #1a2a3a !important;
        padding: 1px 5px;
        border-radius: 3px;
        font-size: 0.9em;
    }

    /* ================================================================
       CHIPS & STATUS LABELS
       ================================================================ */
    .chip-green {
        display: inline-block;
        background: #d4edda; color: #155724 !important;
        border: 1px solid #c3e6cb;
        border-radius: 16px; padding: 3px 12px;
        font-size: 13px; font-weight: 600;
        margin: 3px 4px;
    }
    .chip-red {
        display: inline-block;
        background: #f8d7da; color: #721c24 !important;
        border: 1px solid #f5c6cb;
        border-radius: 16px; padding: 3px 12px;
        font-size: 13px; font-weight: 600;
        margin: 3px 4px;
    }
    .chip-amber {
        display: inline-block;
        background: #fff3cd; color: #5a4000 !important;
        border: 1px solid #ffeeba;
        border-radius: 16px; padding: 3px 12px;
        font-size: 13px; font-weight: 600;
        margin: 3px 4px;
    }
    .status-label {
        display: inline-block;
        font-size: 12px; font-weight: 600;
        padding: 2px 8px; border-radius: 4px; margin-right: 6px;
    }
    .status-label-green  { background: #d4edda; color: #155724 !important; }
    .status-label-red    { background: #f8d7da; color: #721c24 !important; }
    .status-label-amber  { background: #fff3cd; color: #5a4000 !important; }
    .status-label-blue   { background: #d1ecf1; color: #0c5460 !important; }

    /* ================================================================
       STAT BOXES
       ================================================================ */
    .stat-row {
        display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px;
    }
    .stat-box {
        flex: 1 1 110px;
        background: white;
        border: 1px solid #d0dcea;
        border-radius: 8px;
        padding: 14px 16px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .stat-number { font-size: 1.9em; font-weight: 700; line-height: 1.1; }
    .stat-label  { font-size: 11px; color: #444 !important; margin-top: 4px; }
    .stat-green  .stat-number { color: #1a7a3c !important; }
    .stat-red    .stat-number { color: #c0392b !important; }
    .stat-blue   .stat-number { color: #1F4E79 !important; }
    .stat-amber  .stat-number { color: #7a5800 !important; }

    /* ================================================================
       MISC
       ================================================================ */
    [data-testid="stFileUploader"] {
        border: 2px dashed #2E75B6 !important;
        border-radius: 8px !important;
    }

    .qc-footer {
        text-align: center;
        color: #555 !important;
        font-size: 12px;
        padding: 8px 0;
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
        # Configuration gate — only True after the user clicks "Apply Configuration"
        "config_applied":    False,
        # Hash of the widget values at last Apply; used to detect stale config
        "config_key":        "",
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


def _styled_dataframe(df: pd.DataFrame, style_fn, **kwargs) -> None:
    """
    Render a styled dataframe, raising the Pandas Styler cell limit as needed.
    Falls back to an unstyled table with an explanatory caption when the dataset
    is too large to colour-code efficiently (> 1,000,000 cells).
    """
    n_cells = df.shape[0] * df.shape[1]
    if n_cells > 1_000_000:
        st.dataframe(df, **kwargs)
        st.caption(
            f"Row colouring is disabled for this table ({n_cells:,} cells). "
            "Export the Full Comparison Report for colour-coded results in Excel."
        )
        return
    if n_cells > pd.get_option("styler.render.max_elements"):
        pd.set_option("styler.render.max_elements", n_cells)
    st.dataframe(style_fn(df), **kwargs)


def _instruction(title: str, body: str) -> None:
    """Render a blue instruction/explanation callout."""
    st.markdown(
        f'<div class="instruction-box">'
        f'<div class="inst-title">'
        f'<i class="fas fa-circle-info"></i>&nbsp;&nbsp;{title}'
        f'</div>{body}</div>',
        unsafe_allow_html=True,
    )


def _sect_header(step: int, icon_class: str, title: str, colour_class: str) -> None:
    """Render the coloured stripe at the top of a section container."""
    st.markdown(
        f'<div class="sect-header {colour_class}">'
        f'<span class="step-badge">{step}</span>'
        f'<i class="{icon_class}"></i>'
        f'&nbsp;{title}'
        f'</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# App header banner
# ---------------------------------------------------------------------------

st.markdown(
    """
    <div style="background:linear-gradient(135deg,#1F4E79,#2E75B6);
                border-radius:10px; padding:24px 32px; margin-bottom:20px;">
      <h1 style="color:#ffffff !important; margin:0; font-size:1.8em;">
        <i class="fas fa-magnifying-glass-chart" style="margin-right:12px;"></i>
        Oceaneering QC Comparison Tool
      </h1>
      <p style="color:#d4e6f5 !important; margin:8px 0 0; font-size:14px; line-height:1.6;">
        Upload a client Excel sheet and an INFORM spreadsheet to validate tags,
        match column headers, and compare data cell-by-cell. Work through each
        numbered step in sequence to produce a full QC report.
      </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ===========================================================================
# SECTION 1 — File upload
# ===========================================================================

with st.container(border=True):
    _sect_header(1, "fas fa-file-arrow-up", "Upload Files", "sect-upload")

    _instruction(
        "How to use this section",
        """<ul>
          <li><strong>Client Sheet</strong> — the Excel file supplied by the client
              containing instrument or equipment data to be quality-checked.</li>
          <li><strong>INFORM Sheet</strong> — the reference spreadsheet exported from
              the INFORM system. This is the source of truth against which the client
              data is compared.</li>
          <li>Both files must be in <strong>.xlsx</strong> , <strong>.xls</strong> or
               <strong>.xlsm</strong> format. Once both files are uploaded the remaining steps will appear
              below.</li>
        </ul>""",
    )

    col_up_client, col_up_inform = st.columns(2, gap="large")

    with col_up_client:
        st.markdown("#### Client Sheet")
        client_file = st.file_uploader(
            "Upload client Excel file",
            type=["xlsx", "xls", "xlsm"],
            key="upload_client",
            label_visibility="collapsed",
        )
        if client_file:
            if st.session_state.client_filename != client_file.name:
                st.session_state.client_bytes      = _safe_read(client_file)
                st.session_state.client_filename   = client_file.name
                st.session_state.client_df         = None
                st.session_state.comparison_result = None
                st.session_state.comparison_ran    = False
                st.session_state.config_applied    = False
                st.session_state.config_key        = ""
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
                st.session_state.inform_bytes      = _safe_read(inform_file)
                st.session_state.inform_filename   = inform_file.name
                st.session_state.inform_df         = None
                st.session_state.comparison_result = None
                st.session_state.comparison_ran    = False
                st.session_state.config_applied    = False
                st.session_state.config_key        = ""
            st.success(f"Loaded: **{inform_file.name}**")

# ---------------------------------------------------------------------------
# Guard — everything below requires both files
# ---------------------------------------------------------------------------

_files_ready = bool(
    st.session_state.client_bytes and st.session_state.inform_bytes
)

if not _files_ready:
    st.info("Upload both files above to continue with the comparison.")

else:
    # =======================================================================
    # SECTION 2 — Sheet selection & row configuration
    # =======================================================================

    # Safe defaults — overwritten below if sheet names load successfully
    _apply_clicked = False
    _current_key   = ""

    with st.container(border=True):
        _sect_header(2, "fas fa-table-cells-large", "Configure Sheets", "sect-config")

        _instruction(
            "How to use this section",
            """<ul>
              <li><strong>Sheet tab</strong> — select which worksheet within the
                  uploaded file should be used. If your file has only one sheet it
                  will be pre-selected.</li>
              <li><strong>Header row</strong> — the row number (counting from 1) that
                  contains column names. For example, if column names appear on row 3,
                  enter 3.</li>
              <li><strong>Data start row</strong> — the row number where the first
                  actual data record begins. Must be greater than the header row.</li>
              <li><strong>Original / New marker row</strong> (INFORM only) — the row
                  that labels each column as "Original" or "New". Only columns labelled
                  "Original" are extracted for comparison. Set this to the same value as
                  the header row to include all columns.</li>
              <li>Use the <strong>Preview</strong> panel inside each column to visually
                  confirm your row selections.
                  <span class="status-label status-label-green">Green</span> = Header row,
                  <span class="status-label status-label-amber">Amber</span> = Original/New marker row,
                  <span class="status-label status-label-blue">Blue</span> = Data start row.</li>
              <li>When satisfied with the configuration, click
                  <strong>Apply Configuration</strong> at the bottom of this section to
                  load the sheets. Steps 3 onwards will not appear until configuration
                  has been applied.</li>
            </ul>""",
        )

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
                with st.container(border=True):
                    st.markdown("##### Client Sheet Configuration")

                    client_sheet = st.selectbox(
                        "Sheet to use", client_sheets, key="sel_client_sheet"
                    )

                    c1, c2 = st.columns(2)
                    with c1:
                        client_header_row = st.number_input(
                            "Header row",
                            min_value=1, max_value=500, value=1,
                            key="client_hrow",
                            help="Row number (1-indexed) that contains column headers.",
                        )
                    with c2:
                        client_data_row = st.number_input(
                            "Data start row",
                            min_value=2, max_value=500, value=2,
                            key="client_drow",
                            help="Row number (1-indexed) where data begins.",
                        )

                    if client_data_row <= client_header_row:
                        st.warning("Data start row must be greater than the header row.")

                    with st.expander("Preview raw rows (first 8 rows)", expanded=False):
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
                                '<span class="status-label status-label-green">Green</span>'
                                ' = Header row&nbsp;&nbsp;'
                                '<span class="status-label status-label-blue">Blue</span>'
                                ' = Data start row',
                                unsafe_allow_html=True,
                            )
                        except Exception as exc:
                            st.warning(f"Preview unavailable: {exc}")

            with col_cfg_inform:
                with st.container(border=True):
                    st.markdown("##### INFORM Sheet Configuration")

                    inform_sheet = st.selectbox(
                        "Sheet to use", inform_sheets, key="sel_inform_sheet"
                    )

                    i1, i2 = st.columns(2)
                    with i1:
                        inform_header_row = st.number_input(
                            "Header row",
                            min_value=1, max_value=500, value=1,
                            key="inform_hrow",
                            help="Row number (1-indexed) that contains column headers.",
                        )
                    with i2:
                        inform_data_row = st.number_input(
                            "Data start row",
                            min_value=2, max_value=500, value=4,
                            key="inform_drow",
                            help="Row number (1-indexed) where data records begin.",
                        )

                    inform_marker_row = st.number_input(
                        "Original / New marker row",
                        min_value=1, max_value=500, value=3,
                        key="inform_mrow",
                        help=(
                            "Row that contains the 'Original' / 'New' label for each "
                            "column. Only columns labelled 'Original' are extracted for "
                            "comparison. Set to the same value as the header row to "
                            "disable this filtering and include all columns."
                        ),
                    )

                    if inform_data_row <= inform_header_row:
                        st.warning("Data start row must be greater than the header row.")

                    with st.expander("Preview raw rows (first 8 rows)", expanded=False):
                        try:
                            preview_df = preview_rows(
                                st.session_state.inform_bytes, inform_sheet, 1, 8
                            )

                            def _hl_inform(row):
                                if row.name == inform_header_row:
                                    return ["background-color:#d4edda"] * len(row)
                                if row.name == inform_marker_row:
                                    return ["background-color:#fff3cd"] * len(row)
                                if row.name == inform_data_row:
                                    return ["background-color:#cce5ff"] * len(row)
                                return [""] * len(row)

                            st.dataframe(
                                preview_df.style.apply(_hl_inform, axis=1),
                                use_container_width=True,
                                height=250,
                            )
                            st.markdown(
                                '<span class="status-label status-label-green">Green</span>'
                                ' = Header row&nbsp;&nbsp;'
                                '<span class="status-label status-label-amber">Amber</span>'
                                ' = Original/New marker row&nbsp;&nbsp;'
                                '<span class="status-label status-label-blue">Blue</span>'
                                ' = Data start row',
                                unsafe_allow_html=True,
                            )
                        except Exception as exc:
                            st.warning(f"Preview unavailable: {exc}")

            # ------------------------------------------------------------------
            # Detect whether the widget values have changed since the last apply
            # ------------------------------------------------------------------
            _current_key = (
                f"{client_sheet}|{client_header_row}|{client_data_row}|"
                f"{inform_sheet}|{inform_header_row}|{inform_marker_row}|{inform_data_row}"
            )
            _config_stale = (
                st.session_state.config_applied
                and st.session_state.config_key != _current_key
            )

            if _config_stale:
                st.warning(
                    "Configuration has changed since the last apply. "
                    "Click **Apply Configuration** to reload the sheets with the new settings."
                )

            # ------------------------------------------------------------------
            # Apply Configuration button — right-aligned
            # ------------------------------------------------------------------
            st.markdown("---")
            _spacer, _btn_col = st.columns([4, 1])
            with _btn_col:
                _apply_clicked = st.button(
                    "Apply Configuration",
                    type="primary",
                    use_container_width=True,
                    key="btn_apply_config",
                )

    # -----------------------------------------------------------------------
    # Parse sheets when Apply is clicked — runs outside the Section 2 border
    # container so error messages appear below it, not inside it.
    # -----------------------------------------------------------------------
    if client_sheets is not None and inform_sheets is not None and _apply_clicked:
        _parse_error = False
        with st.spinner("Loading sheets — please wait ..."):
            try:
                _client_df, _client_headers = parse_client_sheet(
                    st.session_state.client_bytes,
                    client_sheet,
                    int(client_header_row),
                    int(client_data_row),
                )
                st.session_state.client_df      = _client_df
                st.session_state.client_headers = _client_headers
            except Exception as exc:
                st.error(f"Error parsing client sheet: {exc}")
                with st.expander("Traceback"):
                    st.code(traceback.format_exc())
                _parse_error = True

            try:
                _inform_df, _inform_headers = parse_inform_sheet(
                    st.session_state.inform_bytes,
                    inform_sheet,
                    header_row=int(inform_header_row),
                    marker_row=int(inform_marker_row),
                    data_start_row=int(inform_data_row),
                )
                st.session_state.inform_df      = _inform_df
                st.session_state.inform_headers = _inform_headers
            except Exception as exc:
                st.error(f"Error parsing INFORM sheet: {exc}")
                with st.expander("Traceback"):
                    st.code(traceback.format_exc())
                _parse_error = True

        if not _parse_error:
            st.session_state.config_applied    = True
            st.session_state.config_key        = _current_key
            st.session_state.comparison_result = None
            st.session_state.comparison_ran    = False
            st.success("Configuration applied. Continue with Step 3 below.")

    # -----------------------------------------------------------------------
    # Sections 3, 4, 5 — only visible after a successful Apply
    # -----------------------------------------------------------------------
    _show_downstream = (
        client_sheets is not None
        and inform_sheets is not None
        and st.session_state.config_applied
        and st.session_state.config_key == _current_key
    )

    if _show_downstream:
        client_df      = st.session_state.client_df
        client_headers = st.session_state.client_headers
        inform_df      = st.session_state.inform_df
        inform_headers = st.session_state.inform_headers

        if client_df is None or inform_df is None:
            st.error("Sheet data is unavailable. Re-apply configuration.")
        else:
            with st.container(border=True):
                _sect_header(3, "fas fa-tag", "Select Tag Columns", "sect-tags")

                _instruction(
                    "How to use this section",
                    """<ul>
                      <li>The <strong>tag column</strong> contains the unique identifier
                          for each instrument or piece of equipment — for example tag
                          numbers such as <code>FT-1001</code> or <code>PT-2045</code>.</li>
                      <li>Select the tag column from both the Client sheet and the INFORM
                          sheet. These values are used to align rows between the two files
                          before comparing their data.</li>
                      <li>Rows whose tag appears in only one sheet are flagged as issues
                          and excluded from the cell-level comparison. Only tags present in
                          <em>both</em> sheets are compared.</li>
                      <li>Use the <strong>Preview</strong> panels to confirm the parsed
                          data looks correct before proceeding.</li>
                    </ul>""",
                )

                col_tag_client, col_tag_inform = st.columns(2, gap="large")

                with col_tag_client:
                    client_tag_col = st.selectbox(
                        "Tag column — Client sheet",
                        client_headers,
                        key="client_tag_col",
                        help=(
                            "Column whose values uniquely identify each instrument or tag. "
                            "Used to match rows between the two sheets."
                        ),
                    )

                with col_tag_inform:
                    inform_tag_col = st.selectbox(
                        "Tag column — INFORM sheet (Original columns only)",
                        inform_headers,
                        key="inform_tag_col",
                        help=(
                            "The corresponding tag column in the INFORM sheet. "
                            "Must contain the same tag identifiers as the client sheet."
                        ),
                    )

                col_prev_client, col_prev_inform = st.columns(2, gap="large")

                with col_prev_client:
                    with st.expander(
                        f"Preview client data  ({len(client_df):,} rows x "
                        f"{len(client_headers)} columns)",
                        expanded=False,
                    ):
                        st.dataframe(
                            client_df.head(20), use_container_width=True, height=300
                        )

                with col_prev_inform:
                    with st.expander(
                        f"Preview INFORM data  ({len(inform_df):,} rows x "
                        f"{len(inform_headers)} columns)",
                        expanded=False,
                    ):
                        st.dataframe(
                            inform_df.head(20), use_container_width=True, height=300
                        )

            # ===============================================================
            # SECTION 4 — Run comparison
            # ===============================================================

            with st.container(border=True):
                _sect_header(4, "fas fa-play-circle", "Run Comparison", "sect-run")

                _instruction(
                    "What the comparison does",
                    """<ul>
                      <li><strong>Tag validation</strong> — checks that all tag values are
                          unique within each sheet and that every tag is present in both
                          sheets. Duplicates and one-sided tags are reported as issues.</li>
                      <li><strong>Header matching</strong> — compares column names between
                          the two sheets. Only columns with an exact name match in both
                          sheets are included in the cell-level comparison.</li>
                      <li><strong>Cell-level comparison</strong> — for every row/column
                          combination where both the tag and the header match, the values
                          from both sheets are compared. Values are normalised (whitespace
                          trimmed, numeric types reconciled) before comparison to reduce
                          false positives caused by formatting differences.</li>
                      <li>Results appear in the tabs below after the comparison completes.
                          Formatted Excel reports can be downloaded from the
                          <strong>Download Reports</strong> tab.</li>
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
            # SECTION 5 — Results
            # ===============================================================

            if st.session_state.comparison_ran and st.session_state.comparison_result:

                result = st.session_state.comparison_result
                tv     = result.tag_validation
                hm     = result.header_match
                df_all = result.diffs_df()

                total_cells  = len(df_all)
                match_cells  = int(result.match_count)
                mm_cells     = int(result.mismatch_count)
                rate_pct     = f"{match_cells / total_cells * 100:.1f}%" if total_cells else "—"

                with st.container(border=True):
                    _sect_header(5, "fas fa-chart-bar", "Results", "sect-results")

                    st.markdown(
                        f"""
                        <div class="stat-row">
                          {_stat_box(len(tv.comparable_tags), "Comparable Tags", "blue")}
                          {_stat_box(len(tv.issues), "Tag Issues",
                                     "red" if tv.issues else "green")}
                          {_stat_box(len(hm.matched), "Matched Headers", "blue")}
                          {_stat_box(len(hm.client_only) + len(hm.inform_only),
                                     "Unmatched Headers",
                                     "red" if (hm.client_only or hm.inform_only) else "green")}
                          {_stat_box(total_cells, "Cell Comparisons", "blue")}
                          {_stat_box(match_cells, "Cells Match", "green")}
                          {_stat_box(mm_cells, "Cells Differ",
                                     "red" if mm_cells else "green")}
                          {_stat_box(rate_pct, "Match Rate",
                                     "green" if mm_cells == 0 else "amber")}
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
                              <li><strong>Comparable tags</strong> — tag values found in
                                  both sheets. Only these tags are included in the
                                  cell-level comparison.</li>
                              <li><strong>Client-only tags</strong> — tags present in the
                                  client sheet but absent from INFORM. These rows cannot
                                  be compared and are excluded from cell-level results.</li>
                              <li><strong>INFORM-only tags</strong> — tags present in
                                  INFORM but absent from the client sheet. Also excluded
                                  from cell-level results.</li>
                              <li><strong>Duplicate tags</strong> — tag values that appear
                                  more than once within a single sheet. These are flagged
                                  because row matching would be ambiguous.</li>
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
                                "Full comparison is available."
                            )
                        else:
                            dup_issues  = [i for i in tv.issues if "Duplicate" in i.reason]
                            miss_issues = [
                                i for i in tv.issues
                                if "Duplicate" not in i.reason and "(empty)" not in i.tag
                            ]
                            null_issues = [i for i in tv.issues if "(empty)" in i.tag]

                            if null_issues:
                                with st.expander(
                                    f"Empty / Null Tags — {len(null_issues)} issue(s)",
                                    expanded=True,
                                ):
                                    null_df = pd.DataFrame(
                                        [{"Sheet": i.sheet, "Reason": i.reason}
                                         for i in null_issues]
                                    )
                                    st.dataframe(
                                        null_df, use_container_width=True, hide_index=True
                                    )

                            if dup_issues:
                                with st.expander(
                                    f"Duplicate Tags — {len(dup_issues)} tag(s)",
                                    expanded=True,
                                ):
                                    dup_df = pd.DataFrame(
                                        [{"Tag": i.tag, "Sheet": i.sheet, "Reason": i.reason}
                                         for i in dup_issues]
                                    )
                                    st.dataframe(
                                        dup_df.style.apply(
                                            lambda _: ["background-color:#ffc7ce"]
                                                      * len(dup_df.columns),
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
                                        f"Client-Only Tags — {len(client_miss)} tag(s),"
                                        f" no INFORM match",
                                        expanded=len(client_miss) <= 50,
                                    ):
                                        cm_df = pd.DataFrame(
                                            [{"Tag": i.tag, "Reason": i.reason}
                                             for i in client_miss]
                                        )
                                        st.dataframe(
                                            cm_df.style.apply(
                                                lambda _: ["background-color:#fff3cd"]
                                                          * len(cm_df.columns),
                                                axis=1,
                                            ),
                                            use_container_width=True,
                                            hide_index=True,
                                        )

                                if inform_miss:
                                    with st.expander(
                                        f"INFORM-Only Tags — {len(inform_miss)} tag(s),"
                                        f" no client match",
                                        expanded=len(inform_miss) <= 50,
                                    ):
                                        im_df = pd.DataFrame(
                                            [{"Tag": i.tag, "Reason": i.reason}
                                             for i in inform_miss]
                                        )
                                        st.dataframe(
                                            im_df.style.apply(
                                                lambda _: ["background-color:#fff3cd"]
                                                          * len(im_df.columns),
                                                axis=1,
                                            ),
                                            use_container_width=True,
                                            hide_index=True,
                                        )

                            if tv.comparable_tags:
                                st.info(
                                    f"{len(tv.comparable_tags):,} tag(s) are comparable and "
                                    f"have been included in the data comparison."
                                )
                            else:
                                st.error(
                                    "No comparable tags found. Check that the correct tag "
                                    "columns are selected for both sheets."
                                )

                    # -----------------------------------------------------------
                    # TAB 2 — Header matching
                    # -----------------------------------------------------------

                    with tab_headers:
                        _instruction(
                            "About Header Matching",
                            """<ul>
                              <li>Column headers are compared using an
                                  <strong>exact string match</strong>. Names must be
                                  identical in both sheets (including spacing and
                                  capitalisation) to be included in cell-level
                                  comparison.</li>
                              <li><strong>Matched headers</strong> — columns present in
                                  both sheets that will be compared cell-by-cell.</li>
                              <li><strong>Client-only headers</strong> — columns present
                                  in the client sheet but not in INFORM. Their data will
                                  not be compared.</li>
                              <li><strong>INFORM-only headers</strong> — columns present
                                  in INFORM but not in the client sheet. Their data will
                                  not be compared.</li>
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
                                    f"**{len(hm.client_only)} header(s)** exist in the client "
                                    f"sheet but have no matching INFORM column — these will not "
                                    f"be compared."
                                )
                            if hm.inform_only:
                                st.warning(
                                    f"**{len(hm.inform_only)} header(s)** exist in the INFORM "
                                    f"sheet but have no matching client column — these will not "
                                    f"be compared."
                                )

                        st.markdown("#### All Headers at a Glance")

                        hg_col1, hg_col2, hg_col3 = st.columns(3)

                        with hg_col1:
                            st.markdown(
                                '<span class="status-label status-label-green">Matched</span>'
                                ' <strong>Present in both sheets</strong>',
                                unsafe_allow_html=True,
                            )
                            chips = (
                                "".join(_chip(h, "green") for h in hm.matched)
                                if hm.matched else "<i>None</i>"
                            )
                            st.markdown(chips, unsafe_allow_html=True)

                        with hg_col2:
                            st.markdown(
                                '<span class="status-label status-label-amber">Client only</span>'
                                ' <strong>Not in INFORM</strong>',
                                unsafe_allow_html=True,
                            )
                            chips = (
                                "".join(_chip(h, "amber") for h in hm.client_only)
                                if hm.client_only else "<i>None</i>"
                            )
                            st.markdown(chips, unsafe_allow_html=True)

                        with hg_col3:
                            st.markdown(
                                '<span class="status-label status-label-amber">INFORM only</span>'
                                ' <strong>Not in client sheet</strong>',
                                unsafe_allow_html=True,
                            )
                            chips = (
                                "".join(_chip(h, "amber") for h in hm.inform_only)
                                if hm.inform_only else "<i>None</i>"
                            )
                            st.markdown(chips, unsafe_allow_html=True)

                        with st.expander("View as table", expanded=False):
                            max_len = max(
                                len(hm.matched), len(hm.client_only), len(hm.inform_only), 1
                            )
                            hdr_table = pd.DataFrame({
                                "Matched Headers":     (
                                    hm.matched + [""] * (max_len - len(hm.matched))
                                ),
                                "Client-Only Headers": (
                                    hm.client_only + [""] * (max_len - len(hm.client_only))
                                ),
                                "INFORM-Only Headers": (
                                    hm.inform_only + [""] * (max_len - len(hm.inform_only))
                                ),
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
                              <li>Each row represents one cell compared between the two
                                  sheets, identified by its <strong>Tag</strong> (row
                                  identifier) and <strong>Header</strong> (column
                                  identifier).</li>
                              <li><strong>Client Value / INFORM Value</strong> — the raw
                                  values as they appear in each file.</li>
                              <li><strong>Normalised columns</strong> — values after
                                  whitespace trimming and numeric standardisation.
                                  Comparison is performed on normalised values to reduce
                                  false mismatches caused by formatting differences.</li>
                              <li>Use the <strong>filter controls</strong> to focus on
                                  mismatches, a specific column header, or a specific tag.
                                  The bar chart at the bottom shows which columns contain
                                  the most discrepancies.</li>
                            </ul>""",
                        )

                        if df_all.empty:
                            st.info(
                                "No cell-level comparison data available. This occurs when "
                                "no tags are comparable or no headers matched."
                            )
                        else:
                            mm_df = df_all[~df_all["Match"]]

                            diff_col1, diff_col2, diff_col3 = st.columns(3)
                            with diff_col1:
                                st.metric("Total comparisons", f"{total_cells:,}")
                            with diff_col2:
                                st.metric("Matches", f"{match_cells:,}")
                            with diff_col3:
                                st.metric("Mismatches", f"{mm_cells:,}")

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
                                    "Search by Tag",
                                    value="",
                                    key="tag_search",
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

                            st.caption(
                                f"Showing {len(view_df):,} of {total_cells:,} comparisons"
                            )

                            if view_df.empty:
                                st.info("No rows match the current filters.")
                            else:
                                display_cols = [
                                    "Tag", "Header",
                                    "Client Value", "INFORM Value",
                                    "Client (Normalised)", "INFORM (Normalised)",
                                    "Match",
                                ]

                                def _diff_style_fn(df: pd.DataFrame):
                                    def row_style(row):
                                        colour = "#c6efce" if row["Match"] else "#ffc7ce"
                                        return [f"background-color:{colour}"] * len(row)
                                    return df.style.apply(row_style, axis=1)

                                _styled_dataframe(
                                    view_df[display_cols],
                                    _diff_style_fn,
                                    use_container_width=True,
                                    height=500,
                                    hide_index=True,
                                )

                            if not mm_df.empty:
                                st.markdown("#### Mismatch Count by Header")
                                st.caption(
                                    "Columns with the highest mismatch count are shown first. "
                                    "Use this chart to prioritise which columns to investigate."
                                )
                                mm_by_hdr = (
                                    mm_df.groupby("Header")
                                    .size()
                                    .reset_index(name="Mismatch Count")
                                    .sort_values("Mismatch Count", ascending=False)
                                )
                                st.bar_chart(
                                    mm_by_hdr.set_index("Header")["Mismatch Count"]
                                )

                    # -----------------------------------------------------------
                    # TAB 4 — Downloads
                    # -----------------------------------------------------------

                    with tab_downloads:
                        _instruction(
                            "About the downloadable reports",
                            """<ul>
                              <li><strong>Tag Issues Report</strong> — a standalone Excel
                                  file listing all tag-level problems: duplicates, tags
                                  present in only one sheet, and empty tag values. Use
                                  this to identify data quality issues before re-running
                                  the comparison.</li>
                              <li><strong>Full Comparison Report</strong> — a multi-sheet
                                  Excel workbook containing a Summary, Tag Issues, Header
                                  Info, Mismatches Only, and All Results sheets. Cells are
                                  colour-coded green (match) and red (mismatch) for fast
                                  visual review.</li>
                              <li><strong>Raw CSV</strong> — the complete comparison
                                  dataset in CSV format, suitable for further analysis in
                                  Excel, Power BI, or other tools.</li>
                            </ul>""",
                        )

                        st.markdown("### Export Results")

                        dl_col1, dl_col2 = st.columns(2)

                        with dl_col1:
                            with st.container(border=True):
                                st.markdown("##### Tag Issues Report")
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
                            with st.container(border=True):
                                st.markdown("##### Full Comparison Report")
                                st.markdown(
                                    "Multi-sheet workbook: Summary, Tag Issues, Header Info, "
                                    "Mismatches Only, and All Results with green/red colouring."
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
                            with st.container(border=True):
                                st.markdown("##### Raw Comparison Data (CSV)")
                                st.markdown(
                                    "Download the complete comparison dataset as a CSV file "
                                    "for use in Excel, Power BI, or other analysis tools."
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

st.markdown(
    '<p class="qc-footer">Oceaneering QC Comparison Tool — Internal Use Only</p>',
    unsafe_allow_html=True,
)
