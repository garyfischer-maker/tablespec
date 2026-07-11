"""Streamlit UI — Synaptiq Data Quality Platform.

Two top-level modes:
    Compare   — profile two tables side-by-side, compute schema diff + drift metrics
    Profile   — profile one or more tables independently (no comparison)

Each mode has its own validate / run flow. Outputs (HTML profiles, Excel
workbook, Mermaid diagrams, metamodel JSON) are written to a UC Volume and
displayed inline after a successful run.
"""

from __future__ import annotations

import os
import time
import traceback
from datetime import datetime, timezone
from typing import List, Optional

import streamlit as st
import streamlit.components.v1 as components

from profiler.catalog import (
    Connection,
    TableRef,
    VolumeRef,
    describe_table,
    list_catalogs,
    list_schemas,
    list_tables,
    list_volumes,
    load_connections,
    load_env_labels,
)
from profiler.compare import compare_tables, schema_change_counts
from profiler.row_diff import RowDiffResult, compute_row_diff, diff_pct, summarise
from profiler.excel import write_workbook
from profiler.manifest import ComparisonParams, SideSpec, new_manifest
from profiler.metamodel import DatasetProfile, Lineage, ProfilerRun, new_run_id
from profiler.profile import profile_table
from profiler.storage import (
    ensure_run_folder,
    list_runs,
    make_run_folder,
    read_text,
    write_json,
    write_json_schema,
    write_metamodel,
    write_mermaid_diagrams,
    write_text,
)


# ---------------------------------------------------------------------------
# Page config

st.set_page_config(
    page_title="Tablespec Guidebook and Profiling",
    page_icon="🔬",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Synaptiq brand styling
# Palette extracted from AIQ deck (April 2022):
#   Synaptiq Blue  #8BA4BD  — steel blue, dominant brand colour
#   Synaptiq Amber #C8956A  — warm copper-orange accent
#   Charcoal       #2D3748  — primary text
#   Off-white      #EEF3F8  — secondary backgrounds

st.markdown(
    """
<style>
/* ── Trim default top whitespace, but keep a cushion above the header bar ── */
.block-container,
div[data-testid="stMainBlockContainer"] {
    padding-top: 2.5rem !important;
}

/* ── Brand header bar ─────────────────────────────────────────── */
.synaptiq-header {
    background: linear-gradient(135deg, #8BA4BD 0%, #6B8EAD 100%);
    padding: 1.1rem 2rem 0.9rem 2rem;
    border-radius: 8px;
    margin-bottom: 1.2rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
/* .synaptiq-logo-mark removed — replaced by st.image() above the header bar */
.synaptiq-wordmark {
    color: #FFFFFF;
    font-size: 1.35rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    line-height: 1;
}
.synaptiq-tagline {
    color: rgba(255,255,255,0.72);
    font-size: 0.72rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-top: 2px;
}
/* Three-section header: wordmark left, product name centered, logo right. */
.hdr-left   { flex: 1; text-align: left; }
.hdr-center { flex: 1; text-align: center; }
.hdr-right  { flex: 1; display: flex; justify-content: flex-end; align-items: center; }
.synaptiq-product-name {
    color: #FFFFFF;
    font-size: 1.15rem;
    font-weight: 700;
    letter-spacing: 0.09em;
    text-transform: uppercase;
}

/* ── Tab styling ──────────────────────────────────────────────── */
div[data-testid="stTabs"] button[role="tab"] {
    font-weight: 700;
    font-size: 1.18rem;
    letter-spacing: 0.03em;
    color: #6B8EAD;
    padding: 0.4rem 0.2rem;
}
div[data-testid="stTabs"] button[role="tab"] p {
    font-size: 1.18rem;
    font-weight: 700;
}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
    color: #C8956A !important;
    border-bottom: 3px solid #C8956A;
}

/* ── Metric tiles ─────────────────────────────────────────────── */
div[data-testid="metric-container"] {
    background: #EEF3F8;
    border-left: 4px solid #8BA4BD;
    border-radius: 6px;
    padding: 0.6rem 0.8rem;
}
div[data-testid="metric-container"] label {
    color: #6B8EAD;
    font-size: 0.78rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
div[data-testid="metric-container"] div[data-testid="metric-value"] {
    color: #2D3748;
    font-weight: 700;
}

/* ── Sidebar ──────────────────────────────────────────────────── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #8BA4BD 0%, #7A96B0 100%);
}
section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] .streamlit-expanderHeader {
    color: rgba(255,255,255,0.85) !important;
}
/* Input fields and buttons inside sidebar need dark text to be readable.
   The global sidebar * { color: white } rule must be overridden for these
   light-background elements — add any new sidebar widget types here. */
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea {
    color: #2D3748 !important;
    background: #FFFFFF !important;
}
section[data-testid="stSidebar"] input::placeholder,
section[data-testid="stSidebar"] textarea::placeholder {
    color: #9AA5B4 !important;
}
section[data-testid="stSidebar"] button {
    background: rgba(200,149,106,0.75) !important;
    border: none !important;
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] button:hover {
    background: #C8956A !important;
}
section[data-testid="stSidebar"] button p,
section[data-testid="stSidebar"] button span {
    color: #FFFFFF !important;
}

/* ── Buttons ──────────────────────────────────────────────────── */
div[data-testid="stButton"] > button[kind="primary"] {
    background: #C8956A;
    border: none;
    color: white;
    font-weight: 600;
    border-radius: 6px;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    background: #B8845A;
    border: none;
}

/* ── Section subheaders ───────────────────────────────────────── */
h3 { color: #8BA4BD; }
h4 { color: #6B8EAD; }

/* ── Divider accent ───────────────────────────────────────────── */
hr { border-top: 1px solid #C8956A33; }

/* ── Success / info boxes ─────────────────────────────────────── */
div[data-testid="stAlert"][data-type="success"] {
    border-left: 4px solid #C8956A;
    background: #FDF5EE;
}
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Cached data loaders


@st.cache_data(ttl=60)
def _connections() -> List[Connection]:
    return load_connections()


@st.cache_data(ttl=300)
def _env_labels() -> List[str]:
    return load_env_labels()


# Cache catalog/schema/table lookups so the warehouse is only hit once per
# TTL window, not on every Streamlit re-run (which happens on every widget
# interaction). This prevents the UI from blocking while the warehouse wakes.
@st.cache_data(ttl=120, show_spinner=False)
def _table_has_meta_load(catalog: str, schema: str, table: str) -> bool:
    """Check if table has META_Load_DTTM column (via UC REST — no warehouse needed)."""
    try:
        from profiler.catalog import _workspace_client

        info = _workspace_client().tables.get(full_name=f"{catalog}.{schema}.{table}")
        return any(
            (c.name or "").lower() == "meta_load_dttm" for c in (info.columns or [])
        )
    except Exception:
        return False


@st.cache_data(ttl=120, show_spinner=False)
def _load_dates_for_table(catalog: str, schema: str, table: str) -> List[str]:
    """Return distinct load dates from META_Load_DTTM, newest first."""
    from profiler.catalog import _runtime

    if _runtime() != "databricks":
        return []
    try:
        result = _exec_suggestion_sql(
            f"SELECT DISTINCT DATE(META_Load_DTTM) AS d "
            f"FROM `{catalog}`.`{schema}`.`{table}` "
            f"WHERE META_Load_DTTM IS NOT NULL "
            f"ORDER BY d DESC LIMIT 60"
        )
        if result.result and result.result.data_array:
            return [str(row[0]) for row in result.result.data_array if row[0]]
        return []
    except Exception:
        return []


@st.cache_data(ttl=120)
def _cached_schemas(catalog: str) -> List[str]:
    conn = _connections()[0] if _connections() else None
    if conn is None:
        return []
    return list_schemas(conn, catalog)


@st.cache_data(ttl=120)
def _cached_tables(catalog: str, schema: str) -> List[str]:
    conn = _connections()[0] if _connections() else None
    if conn is None:
        return []
    return list_tables(conn, catalog, schema)


def _connection_by_name(name: str) -> Optional[Connection]:
    for c in _connections():
        if c.name == name:
            return c
    return None


# ---------------------------------------------------------------------------
# Reusable widget sections


def _table_picker(key: str, title: str = "", default_env_idx: int = 0) -> dict:
    """Cascading catalog → schema → table picker. Returns a side dict.

    Every SQL call is wrapped in try/except so a warehouse error in one
    picker never prevents the other side from rendering.
    """
    if title:
        st.markdown(f"**{title}**")

    conn_names = [c.name for c in _connections()]
    if len(conn_names) == 1:
        conn_name = conn_names[0]
        conn = _connection_by_name(conn_name)
    else:
        conn_name = st.selectbox("Connection", options=conn_names, key=f"{key}_conn")
        conn = _connection_by_name(conn_name)

    env_label = st.selectbox(
        "Env label",
        options=_env_labels(),
        index=default_env_idx,
        key=f"{key}_env",
        help="Label used in report titles and run folder names.",
    )

    # ── Catalog ──────────────────────────────────────────────────
    try:
        catalogs = list_catalogs(conn) if conn else []
    except Exception as exc:  # noqa: BLE001
        st.error(f"Cannot load catalogs: {exc}")
        catalogs = []

    catalog = st.selectbox(
        "Catalog",
        options=catalogs,
        key=f"{key}_catalog",
        index=0 if catalogs else None,
    )

    # ── Schema ───────────────────────────────────────────────────
    schemas: List[str] = []
    if catalog:
        try:
            schemas = _cached_schemas(catalog)
        except Exception as exc:  # noqa: BLE001
            st.warning(
                f"Cannot load schemas for `{catalog}` — {exc}\n\n"
                f"Run in Databricks SQL:  "
                f"`GRANT USE CATALOG ON CATALOG {catalog} TO <app-sp>;`  \n"
                f"`GRANT USE SCHEMA, SELECT ON ALL SCHEMAS IN CATALOG {catalog} TO <app-sp>;`"
            )

    if catalog and not schemas:
        st.caption(
            f"No schemas visible in `{catalog}` — check service principal grants."
        )

    schema = st.selectbox(
        "Schema",
        options=schemas,
        key=f"{key}_schema",
        index=0 if schemas else None,
        placeholder="— pick a schema —",
    )

    # ── Table ────────────────────────────────────────────────────
    tables: List[str] = []
    if catalog and schema:
        try:
            tables = _cached_tables(catalog, schema)
        except Exception as exc:  # noqa: BLE001
            st.warning(f"Cannot load tables for `{catalog}.{schema}` — {exc}")

    if schema and not tables:
        st.caption(f"No tables in `{catalog}.{schema}`.")

    table = st.selectbox(
        "Table",
        options=tables,
        key=f"{key}_table",
        index=0 if tables else None,
        placeholder="— pick a table —",
    )

    # ── Load date filter (META_Load_DTTM) ────────────────────────
    load_date: Optional[str] = None
    if catalog and schema and table:
        has_meta = _table_has_meta_load(catalog, schema, table)
        if has_meta:
            filter_on = st.checkbox(
                "Filter by load date (META_Load_DTTM)",
                key=f"{key}_filter_date",
                help="Profile only rows from a specific nightly load batch.",
            )
            if filter_on:
                dates = _load_dates_for_table(catalog, schema, table)
                if dates:
                    chosen = st.selectbox(
                        "Load date",
                        options=["All dates (full table)"] + dates,
                        key=f"{key}_load_date",
                        help="Newest load first.",
                    )
                    if chosen != "All dates (full table)":
                        load_date = chosen
                else:
                    st.caption("No load dates found — warehouse may be offline.")

    return {
        "env_label": env_label,
        "connection": conn,
        "catalog": catalog,
        "schema": schema,
        "table": table,
        "load_date": load_date,
    }


def _sampling_section(key: str) -> tuple[str, int, str]:
    """Sampling controls. Returns (sampling_mode, sample_n, stratify_by)."""
    with st.expander("Sampling options", expanded=False):
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            sampling_mode = st.selectbox(
                "Mode",
                options=["Full table", "Sample N rows", "Stratified by column"],
                index=0,
                key=f"{key}_sampling_mode",
            )
        with col2:
            sample_n = st.number_input(
                "N rows",
                min_value=1_000,
                max_value=100_000_000,
                value=1_000_000,
                step=1_000,
                disabled=(sampling_mode != "Sample N rows"),
                key=f"{key}_sample_n",
            )
        with col3:
            stratify_by = st.text_input(
                "Stratify column",
                value="",
                placeholder="e.g. region",
                disabled=(sampling_mode != "Stratified by column"),
                key=f"{key}_stratify",
            )
    return sampling_mode, int(sample_n), stratify_by


def _output_section(key: str) -> tuple[Optional[VolumeRef], str, str]:
    """Output volume picker. Returns (VolumeRef | None, catalog, schema)."""
    st.markdown("**Output destination**")
    local_conns = [c for c in _connections() if c.type == "native"]
    local_conn = local_conns[0] if local_conns else None
    local_cats = list_catalogs(local_conn) if local_conn else []

    # Pre-select the first catalog/schema that has a volume.
    def _first_with_vol():
        if not local_conn:
            return None, None
        for cat in list_catalogs(local_conn):
            for sch in list_schemas(local_conn, cat):
                if list_volumes(cat, sch):
                    return cat, sch
        cats = list_catalogs(local_conn)
        if cats:
            schs = list_schemas(local_conn, cats[0])
            return cats[0], (schs[0] if schs else None)
        return None, None

    def _idx(opts, val):
        return opts.index(val) if (val is not None and val in opts) else 0

    def_cat, def_sch = _first_with_vol()

    c1, c2, c3, c4 = st.columns([2, 2, 2, 3])
    with c1:
        out_cat = st.selectbox(
            "Catalog",
            options=local_cats,
            index=_idx(local_cats, def_cat) if local_cats else None,
            key=f"{key}_out_cat",
        )
    with c2:
        out_schs = list_schemas(local_conn, out_cat) if (local_conn and out_cat) else []
        out_sch = st.selectbox(
            "Schema",
            options=out_schs,
            index=_idx(out_schs, def_sch) if out_schs else None,
            key=f"{key}_out_sch",
        )
    with c3:
        out_vols = list_volumes(out_cat, out_sch) if (out_cat and out_sch) else []
        out_vol = st.selectbox(
            "Volume",
            options=out_vols,
            index=0 if out_vols else None,
            key=f"{key}_out_vol",
        )
    with c4:
        run_label = st.text_input(
            "Run label (optional)",
            value="",
            key=f"{key}_run_label",
            help="Appended to run folder name.",
        )

    if out_cat and out_sch and not out_vols:
        st.warning(
            f"No volumes in `{out_cat}.{out_sch}`. "
            f"Create one: `CREATE VOLUME {out_cat}.{out_sch}.ab_runs;`"
        )

    if out_cat and out_sch and out_vol:
        vol_ref = VolumeRef(catalog=out_cat, schema=out_sch, volume=out_vol)
        st.session_state["output_volume"] = vol_ref
    else:
        vol_ref = None

    return vol_ref, out_cat, out_sch, run_label


def _render_run_outputs(
    folder,
    profiler_run: ProfilerRun,
    mode: str,
    row_diff: Optional[RowDiffResult] = None,
) -> None:
    """Display run results inline: summary stats, Mermaid diagrams, HTML iframes."""
    st.divider()
    st.subheader("Profile results" if mode == "profile" else "Comparison results")

    # Summary metrics
    a = profiler_run.side_a
    b = profiler_run.side_b if mode == "compare" else None
    is_compare = mode == "compare"

    col1, col2, col3, col4 = st.columns(4)

    # Use neutral labels in profile mode; "Side A / Side B" only in compare mode
    col1.metric("Side A rows" if is_compare else "Rows", f"{a.row_count:,}")
    col1.metric("Side A columns" if is_compare else "Columns", a.column_count)
    if b:
        col2.metric("Side B rows", f"{b.row_count:,}")
        col2.metric("Side B columns", b.column_count)

    total_alerts_a = sum(len(c.alerts) for c in a.columns)
    col3.metric("Side A alerts" if is_compare else "Alerts", total_alerts_a)
    if b:
        total_alerts_b = sum(len(c.alerts) for c in b.columns)
        col3.metric("Side B alerts", total_alerts_b)

    if is_compare and profiler_run.comparisons:
        changes = schema_change_counts(profiler_run.comparisons)
        n_drifted = sum(
            1
            for c in profiler_run.comparisons
            if c.verdict in ("moderate", "significant")
        )
        n_schema_changed = sum(v for k, v in changes.items() if k != "unchanged")
        col4.metric("Schema changes", n_schema_changed)
        col4.metric("Drifted columns", n_drifted)

    # Mermaid diagrams
    st.markdown("#### Schema diagrams")
    mmd_tab_labels = (
        ["Side A schema", "Side B schema", "Drift view"]
        if mode == "compare"
        else ["Schema"]
    )
    mmd_files = (
        ["schema_a.mmd", "schema_b.mmd", "drift.mmd"]
        if mode == "compare"
        else ["schema_a.mmd"]
    )

    import html as _html

    mmd_tabs = st.tabs(mmd_tab_labels)
    for tab, fname in zip(mmd_tabs, mmd_files):
        with tab:
            path = f"{folder.path}/{fname}"
            try:
                mmd_src = read_text(path)
                mmd_escaped = _html.escape(mmd_src)
                components.html(
                    f"""<!DOCTYPE html>
<html><head>
<script src="https://cdn.jsdelivr.net/npm/mermaid@10.9.0/dist/mermaid.min.js"></script>
<style>
  body{{margin:0;background:white;overflow:hidden}}
  #ctrl{{position:sticky;top:0;z-index:100;background:rgba(255,255,255,.95);
         padding:4px 8px;display:flex;gap:4px;align-items:center;
         border-bottom:1px solid #e0e8f0}}
  #ctrl button{{background:#8BA4BD;color:white;border:none;border-radius:4px;
                padding:3px 10px;font-size:12px;cursor:pointer;font-weight:600}}
  #ctrl button:hover{{background:#6B8EAD}}
  #ctrl span{{font-size:12px;color:#666;min-width:38px;text-align:center}}
  #outer{{overflow:auto;width:100%;height:calc(100vh - 36px);padding:8px;box-sizing:border-box}}
  #inner{{transform-origin:top left;display:inline-block;min-width:100%}}
  #err{{color:#c0392b;padding:8px;font-size:12px}}
  pre.src{{font-size:10px;color:#888;white-space:pre-wrap;word-break:break-all}}
</style>
</head><body>
<div id="ctrl">
  <button onclick="z(0.25)">＋ Zoom in</button>
  <button onclick="z(-0.25)">－ Zoom out</button>
  <button onclick="fit()">⊡ Fit</button>
  <button onclick="reset()">1:1</button>
  <span id="pct">100%</span>
</div>
<div id="outer"><div id="inner">
  <pre class="mermaid">{mmd_escaped}</pre>
</div></div>
<script>
var sc=1;
function applyScale(){{
  var inner=document.getElementById('inner');
  inner.style.transform='scale('+sc+')';
  inner.style.width=(100/sc)+'%';
  document.getElementById('pct').textContent=Math.round(sc*100)+'%';
}}
function z(d){{sc=Math.max(0.1,Math.min(5,sc+d));applyScale();}}
function reset(){{sc=1;applyScale();}}
function fit(){{
  var outer=document.getElementById('outer');
  var svg=outer.querySelector('svg');
  if(svg){{sc=Math.min(1,(outer.clientWidth-20)/svg.getBoundingClientRect().width);applyScale();}}
}}
mermaid.initialize({{startOnLoad:false,theme:'default',securityLevel:'loose',
  er:{{useMaxWidth:false}},flowchart:{{useMaxWidth:false}}}});
document.addEventListener('DOMContentLoaded',async function(){{
  try{{
    await mermaid.run({{querySelector:'.mermaid'}});
    setTimeout(fit,200);
  }}catch(e){{
    document.getElementById('inner').innerHTML=
      '<div id="err">Render error: '+e.message+'</div>'+
      '<details><summary style="cursor:pointer;color:#666;font-size:11px">Show source</summary>'+
      '<pre class="src">'+{repr(mmd_escaped)}+'</pre></details>';
  }}
}});
</script>
</body></html>""",
                    height=580,
                    scrolling=False,
                )
            except Exception as exc:
                st.warning(f"`{fname}` — {exc}")

    # HTML profile reports (iframe)
    st.markdown("#### Profile reports")
    html_files = (
        [("Side A", "profile_a.html"), ("Side B", "profile_b.html")]
        if mode == "compare"
        else [("Profile", "profile_a.html")]
    )
    html_tabs = st.tabs([label for label, _ in html_files])
    for tab, (label, fname) in zip(html_tabs, html_files):
        with tab:
            path = f"{folder.path}/{fname}"
            try:
                html_content = read_text(path)
                components.html(html_content, height=800, scrolling=True)
            except Exception:
                st.info(
                    f"`{fname}` not yet available — profile may still be generating."
                )

    # Row diff results
    if row_diff is not None and mode == "compare":
        st.markdown("#### Row-level diff")
        if row_diff.error:
            st.warning(f"Row diff error: {row_diff.error}")
        else:
            r1, r2, r3, r4 = st.columns(4)
            r1.metric(
                "Removed",
                f"{row_diff.rows_only_in_a:,}",
                help="Rows in Side A with no matching key in Side B",
            )
            r2.metric(
                "Added",
                f"{row_diff.rows_only_in_b:,}",
                help="Rows in Side B with no matching key in Side A",
            )
            r3.metric(
                "Changed",
                f"{row_diff.rows_changed:,}",
                help="Rows with same key but different values",
            )
            r4.metric(
                "Identical",
                f"{row_diff.rows_identical:,}",
                help="Rows with same key and identical values",
            )

            if row_diff.has_differences:
                import pandas as _pd

                cols = row_diff.col_names or []
                rdiff_tabs = st.tabs(["Removed", "Added", "Changed"])
                with rdiff_tabs[0]:
                    if row_diff.sample_removed:
                        st.dataframe(
                            _pd.DataFrame(row_diff.sample_removed, columns=cols),
                            use_container_width=True,
                        )
                    else:
                        st.caption("No removed rows.")
                with rdiff_tabs[1]:
                    if row_diff.sample_added:
                        st.dataframe(
                            _pd.DataFrame(row_diff.sample_added, columns=cols),
                            use_container_width=True,
                        )
                    else:
                        st.caption("No added rows.")
                with rdiff_tabs[2]:
                    if row_diff.sample_changed:
                        st.dataframe(
                            _pd.DataFrame(row_diff.sample_changed),
                            use_container_width=True,
                        )
                    else:
                        st.caption("No changed rows.")
            else:
                st.success(
                    "All matched rows are identical — no row-level differences found."
                )

    # Artifact paths
    with st.expander("Output file locations", expanded=False):
        st.code(folder.path, language="text")
        for art in [
            "manifest.json",
            "metamodel.json",
            "ab_summary.xlsx",
            "profile_a.html",
            "profile_b.html",
            "schema_a.mmd",
            "schema_b.mmd",
            "drift.mmd",
        ]:
            st.caption(f"`{folder.path}/{art}`")


# ---------------------------------------------------------------------------
# Sidebar

# ---------------------------------------------------------------------------
# Suggestions helpers


def _exec_suggestion_sql(statement: str) -> any:
    """Run a SQL statement for suggestions via Statement Execution API."""
    import os as _os, time as _time
    from profiler.catalog import _workspace_client

    w = _workspace_client()
    wid = _os.environ.get("DATABRICKS_WAREHOUSE_ID", "")
    result = w.statement_execution.execute_statement(
        warehouse_id=wid, statement=statement, wait_timeout="50s"
    )
    deadline = _time.time() + 120
    while True:
        state = str(result.status.state).upper() if result.status else "UNKNOWN"
        if "SUCCEEDED" in state:
            return result
        if any(s in state for s in ("FAILED", "CANCELLED", "CLOSED")):
            err = (
                result.status.error.message
                if (result.status and result.status.error)
                else state
            )
            raise RuntimeError(err)
        if _time.time() > deadline:
            raise TimeoutError("Timed out after 2 minutes")
        _time.sleep(2)
        result = w.statement_execution.get_statement(result.statement_id)


@st.cache_data(ttl=60, show_spinner=False)
def _query_df(statement: str):
    """Run a SELECT via the Statement Execution API and return a pandas DataFrame.

    Column names come from the result manifest; all values arrive as strings
    (callers cast numeric columns as needed). Cached briefly so tab re-renders
    don't re-hit the warehouse on every widget interaction.
    """
    import pandas as pd

    res = _exec_suggestion_sql(statement)
    schema = res.manifest.schema if (res.manifest and res.manifest.schema) else None
    cols = [c.name for c in schema.columns] if (schema and schema.columns) else []
    data = res.result.data_array if (res.result and res.result.data_array) else []
    return pd.DataFrame(data, columns=cols)


def _submit_suggestion(name: str, suggestion: str) -> str:
    """Insert a suggestion row. Returns '' on success, error string on failure."""
    import uuid
    from datetime import datetime, timezone
    from profiler.catalog import _runtime

    if _runtime() != "databricks":
        return "Suggestions are only stored in Databricks mode."
    sid = str(uuid.uuid4())
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    name_e = (name or "Anonymous").replace("'", "''")
    sugg_e = suggestion.replace("'", "''")
    try:
        _exec_suggestion_sql("""
            CREATE TABLE IF NOT EXISTS `dev`.`test_main_profiler`.`user_suggestions` (
                suggestion_id STRING NOT NULL, submitted_by STRING,
                suggestion STRING NOT NULL, submitted_at TIMESTAMP NOT NULL
            ) USING DELTA TBLPROPERTIES ('delta.autoOptimize.optimizeWrite' = 'true')
        """)
        _exec_suggestion_sql(
            f"INSERT INTO `dev`.`test_main_profiler`.`user_suggestions` "
            f"(suggestion_id, submitted_by, suggestion, submitted_at) "
            f"VALUES ('{sid}', '{name_e}', '{sugg_e}', TIMESTAMP '{now}')"
        )
        return ""
    except Exception as exc:  # noqa: BLE001
        return str(exc)


def _load_suggestions() -> list:
    """Return up to 20 most-recent suggestions."""
    from profiler.catalog import _runtime

    if _runtime() != "databricks":
        return []
    try:
        result = _exec_suggestion_sql("""
            SELECT submitted_by, suggestion, submitted_at
            FROM `dev`.`test_main_profiler`.`user_suggestions`
            ORDER BY submitted_at DESC LIMIT 20
        """)
        if result.result and result.result.data_array:
            return result.result.data_array
        return []
    except Exception:  # noqa: BLE001
        return []


# ---------------------------------------------------------------------------
# Sidebar


def _load_run_history(limit: int = 20) -> list[dict]:
    """Query the governance Delta tables for recent run summaries."""
    from profiler.catalog import _runtime

    if _runtime() != "databricks":
        return []
    try:
        result = _exec_suggestion_sql(f"""
            SELECT
                pr.run_id,
                pr.run_label,
                pr.created_utc,
                pr.side_a_fqn,
                pr.side_b_fqn,
                COALESCE(da.row_count, 0)    AS rows_a,
                COALESCE(da.column_count, 0) AS cols_a,
                COALESCE(db.row_count, 0)    AS rows_b,
                COALESCE(cc.drifted, 0)       AS drifted_cols,
                COALESCE(cc.schema_chg, 0)    AS schema_changes,
                COALESCE(al.alerts_a, 0)      AS alerts_a,
                COALESCE(al.alerts_b, 0)      AS alerts_b
            FROM dev.test_main_profiler.profiler_runs pr
            LEFT JOIN dev.test_main_profiler.dataset_profiles da
                ON pr.run_id = da.run_id AND da.side = 'A'
            LEFT JOIN dev.test_main_profiler.dataset_profiles db
                ON pr.run_id = db.run_id AND db.side = 'B'
            LEFT JOIN (
                SELECT run_id,
                    COUNT(CASE WHEN verdict IN ('moderate','significant') THEN 1 END) AS drifted,
                    COUNT(CASE WHEN schema_change != 'unchanged' THEN 1 END)          AS schema_chg
                FROM dev.test_main_profiler.column_comparisons
                GROUP BY run_id
            ) cc ON pr.run_id = cc.run_id
            LEFT JOIN (
                SELECT run_id,
                    COUNT(CASE WHEN side = 'A' THEN 1 END) AS alerts_a,
                    COUNT(CASE WHEN side = 'B' THEN 1 END) AS alerts_b
                FROM dev.test_main_profiler.column_alerts
                GROUP BY run_id
            ) al ON pr.run_id = al.run_id
            ORDER BY pr.created_utc DESC
            LIMIT {limit}
        """)
        if not (result.result and result.result.data_array):
            return []
        cols = [c.name for c in result.manifest.schema.columns]
        return [dict(zip(cols, row)) for row in result.result.data_array]
    except Exception:  # noqa: BLE001
        return []


def _render_run_card(run: dict) -> None:
    """Render a compact run summary card in the sidebar."""
    created = str(run.get("created_utc", ""))[:16].replace("T", " ")
    label = run.get("run_label") or ""
    a_fqn = str(run.get("side_a_fqn", ""))
    b_fqn = str(run.get("side_b_fqn", ""))
    # Show just table name for brevity
    a_short = a_fqn.split(".")[-1] if a_fqn else "?"
    b_short = b_fqn.split(".")[-1] if b_fqn else "?"
    rows_a = int(run.get("rows_a", 0))
    rows_b = int(run.get("rows_b", 0))
    drifted = int(run.get("drifted_cols", 0))
    schema = int(run.get("schema_changes", 0))
    alerts = int(run.get("alerts_a", 0)) + int(run.get("alerts_b", 0))

    drift_colour = "#e74c3c" if drifted > 0 else "#27ae60"
    title = f"{a_short} vs {b_short}" + (f" — {label}" if label else "")

    st.sidebar.markdown(
        f"""<div style='background:rgba(255,255,255,0.1);border-radius:6px;
            padding:7px 9px;margin-bottom:6px;font-size:0.78rem;color:white;'>
          <div style='font-weight:600;margin-bottom:2px;white-space:nowrap;
               overflow:hidden;text-overflow:ellipsis;' title='{title}'>{title}</div>
          <div style='opacity:0.7;font-size:0.7rem;margin-bottom:4px;'>{created}</div>
          <div style='display:flex;gap:6px;flex-wrap:wrap;'>
            <span title='Side A rows'>A: {rows_a:,}</span>
            <span title='Side B rows'>B: {rows_b:,}</span>
            <span style='color:{drift_colour};font-weight:600;'
                  title='Drifted columns'>⟳ {drifted}</span>
            <span title='Schema changes'>Δ {schema}</span>
            <span title='Total alerts'>⚠ {alerts}</span>
          </div>
        </div>""",
        unsafe_allow_html=True,
    )


def _render_sidebar_compute():
    """Initialize-compute control in the sidebar (databricks runtime only)."""
    if os.environ.get("PROFILER_RUNTIME", "mock").lower() != "databricks":
        return
    with st.sidebar:
        st.markdown(
            "<div style='font-size:0.72rem;text-transform:uppercase;"
            "letter-spacing:0.08em;opacity:0.8;margin-bottom:4px;'>Compute</div>",
            unsafe_allow_html=True,
        )
        warm_clicked = st.button(
            "⚡ Initialize Compute",
            type="secondary",
            help="Start the SQL warehouse before running a profile.",
            use_container_width=True,
        )
        if warm_clicked:
            wid = os.environ.get("DATABRICKS_WAREHOUSE_ID", "")
            if not wid:
                st.error("DATABRICKS_WAREHOUSE_ID not set.")
            else:
                with st.spinner("Starting warehouse …"):
                    import time as _time
                    from datetime import timedelta as _td

                    _t0 = _time.time()
                    try:
                        from profiler.catalog import _workspace_client

                        _w = _workspace_client()
                        try:
                            _wh = _w.warehouses.get(id=wid)
                            _state = str(_wh.state).upper() if _wh.state else "UNKNOWN"
                        except Exception:
                            _state = "UNKNOWN"
                        if "RUNNING" in _state:
                            st.success("✅ Warehouse already RUNNING.")
                            st.session_state["compute_warmed"] = True
                        else:
                            st.caption(f"State: **{_state}**. Starting …")
                            try:
                                _w.warehouses.start(id=wid)
                                _w.warehouses.wait_get_warehouse_running(
                                    id=wid, timeout=_td(minutes=10)
                                )
                                st.success(f"✅ RUNNING — {_time.time() - _t0:.1f}s.")
                                st.session_state["compute_warmed"] = True
                            except Exception as _start_exc:  # noqa: BLE001
                                st.warning(
                                    f"Could not start automatically: {_start_exc}\n\n"
                                    "Ask an admin to grant **Can use** on the warehouse "
                                    "to the app's service principal, or start it in "
                                    "**SQL → Warehouses**."
                                )
                    except Exception as _exc:  # noqa: BLE001
                        st.error(f"Cannot reach warehouse `{wid}`: {_exc}")
        elif "compute_warmed" not in st.session_state:
            st.caption("💡 Start the warehouse before profiling.")


def _render_sidebar_dashboard():
    """Go-To-Dashboard button pinned at the bottom of the sidebar."""
    st.sidebar.markdown(
        """<div style='text-align:center;padding:10px 0 6px 0;'>
        <a href="https://adb-7405619521761591.11.azuredatabricks.net/dashboardsv3/01f15f7d18171dac85cdc247e47d48a5/published?o=7405619521761591"
           target="_blank" style="text-decoration:none;">
          <button style="background:#C8956A;color:white;border:none;border-radius:6px;
                         padding:9px 18px;font-weight:600;font-size:0.85rem;
                         letter-spacing:0.04em;cursor:pointer;width:100%;">
            Go To Dashboard ↗
          </button>
        </a></div>""",
        unsafe_allow_html=True,
    )


def _sidebar():
    st.sidebar.markdown(
        """
<div style='text-align:center; padding: 0.5rem 0 0.8rem 0;'>
  <div style='font-size:1.1rem; font-weight:700; letter-spacing:0.05em;
              color:#FFFFFF;'>Synaptiq</div>
  <div style='font-size:0.62rem; letter-spacing:0.14em; text-transform:uppercase;
              color:rgba(255,255,255,0.65); margin-top:2px;'>Tablespec Guidebook and Profiling</div>
</div>
""",
        unsafe_allow_html=True,
    )
    st.sidebar.divider()

    _render_sidebar_compute()
    _render_sidebar_dashboard()
    st.sidebar.divider()

    with st.sidebar.expander("Run history", expanded=False):
        if st.button("Load run history", key="load_run_history_btn"):
            st.session_state["show_run_history"] = True

        if st.session_state.get("show_run_history"):
            runs = _load_run_history()
            if not runs:
                st.caption("_No runs yet — complete a profile to populate._")
            else:
                for run in runs:
                    _render_run_card(run)
        else:
            st.caption("Click above to load from the governance tables.")

    st.sidebar.divider()

    # ── Suggestions ──────────────────────────────────────────────────────────
    with st.sidebar.expander("💡 Suggestions", expanded=False):
        st.markdown(
            "<div style='color:white;font-size:0.85rem;margin-bottom:10px;'>"
            "Have an idea or improvement? Let us know!</div>",
            unsafe_allow_html=True,
        )
        with st.form("suggestion_form", clear_on_submit=True):
            s_name = st.text_input("Your name", placeholder="e.g. Gary Fischer")
            s_text = st.text_area(
                "Suggestion",
                placeholder="Describe the feature or improvement you'd like to see…",
                height=120,
            )
            s_submitted = st.form_submit_button("Submit", type="primary")

        if s_submitted:
            if not s_text.strip():
                st.warning("Please enter a suggestion.")
            else:
                err = _submit_suggestion(s_name.strip(), s_text.strip())
                if err:
                    st.error(f"Could not save: {err}")
                else:
                    st.success("Thanks! Suggestion saved.")

        if st.button("Show suggestion history", key="load_suggestions_btn"):
            st.session_state["show_suggestions"] = True

        if st.session_state.get("show_suggestions"):
            history = _load_suggestions()
            if not history:
                st.caption("_No suggestions yet._")
            else:
                st.markdown(
                    "<div style='color:rgba(255,255,255,0.5);font-size:0.75rem;"
                    "text-transform:uppercase;letter-spacing:0.08em;"
                    "margin:6px 0 4px 0;'>Recent suggestions</div>",
                    unsafe_allow_html=True,
                )
                for row in history:
                    who = row[0] or "Anonymous"
                    text = row[1] or ""
                    when = str(row[2])[:16] if row[2] else ""
                    st.markdown(
                        f"<div style='background:rgba(255,255,255,0.1);"
                        f"border-radius:4px;padding:6px 8px;margin-bottom:6px;"
                        f"font-size:0.8rem;color:white;'>"
                        f"<b>{who}</b>&nbsp;"
                        f"<span style='opacity:0.6;font-size:0.72rem;'>{when}</span>"
                        f"<br>{text}</div>",
                        unsafe_allow_html=True,
                    )


_sidebar()

# ---------------------------------------------------------------------------
# Main layout — logo row + dashboard link, then blue header bar

import os as _os

_logo_path = _os.path.join(_os.path.dirname(__file__), "assets", "Synaptiq_001.png")

# Blue brand header bar: wordmark (left) · product name (center). Empty right
# section keeps the centered title balanced.
st.markdown(
    """
<div class="synaptiq-header">
  <div class="hdr-left">
    <div class="synaptiq-wordmark">Synaptiq</div>
    <div class="synaptiq-tagline">The Humankind of AI</div>
  </div>
  <div class="hdr-center">
    <div class="synaptiq-product-name">Tablespec Guidebook and Profiling</div>
  </div>
  <div class="hdr-right"></div>
</div>
""",
    unsafe_allow_html=True,
)

# Subtitle (left) + logo (far right), on the row directly under the bar.
_sub_col, _sub_logo_col = st.columns([5, 1], vertical_alignment="center")
with _sub_col:
    st.markdown(
        "<div style='font-size:1.18rem;color:#4A5568;font-weight:500;"
        "padding-top:4px;'>Your spec driven ingestion framework: "
        "Guidebook and Profiling tools</div>",
        unsafe_allow_html=True,
    )
with _sub_logo_col:
    if _os.path.exists(_logo_path):
        st.image(_logo_path, width=150)

st.divider()
_GENIE_SPACE_ID = os.environ.get("GENIE_SPACE_ID", "")

tab_guidebook, tab_compare, tab_profile, tab_load, tab_genie = st.tabs(
    [
        "📖  Guidebook",
        "⚖️  Compare two tables",
        "🔍  Profile table(s)",
        "📥  Load Results",
        "🤖  Ask Genie",
    ]
)


# ===========================================================================
# TAB 1 — COMPARE
# ===========================================================================

with tab_compare:
    st.markdown(
        "Profile two tables side-by-side and compute schema diff + "
        "drift metrics (PSI, KS, Chi-square, JS divergence)."
    )
    st.divider()

    # ---- Table pickers ----
    st.subheader("1. Pick the two tables")
    col_a, col_b = st.columns(2, gap="large")

    _cmp_labels = _env_labels()
    _default_a: dict = {}
    _default_b: dict = {}

    with col_a:
        try:
            idx_a = _cmp_labels.index("PROD") if "PROD" in _cmp_labels else 0
            _default_a = _table_picker(
                "cmp_a", title="Side A — baseline", default_env_idx=idx_a
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"Side A error: {exc}")
            st.code(traceback.format_exc(), language="python")

    with col_b:
        try:
            idx_b = (
                _cmp_labels.index("TEST")
                if "TEST" in _cmp_labels
                else min(1, len(_cmp_labels) - 1)
            )
            _default_b = _table_picker(
                "cmp_b", title="Side B — candidate / TEST", default_env_idx=idx_b
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"Side B error: {exc}")
            st.code(traceback.format_exc(), language="python")

    side_a = _default_a
    side_b = _default_b

    st.divider()

    # ---- Settings ----
    st.subheader("2. Settings")
    with st.expander("Comparison depth", expanded=False):
        depth = st.radio(
            "Depth",
            options=[
                "Aggregate + distributions + schema diff",
                "Include row-level diff (requires row key)",
            ],
            index=0,
            horizontal=False,
            key="cmp_depth",
        )
        with_row_level = depth.startswith("Include row-level")
        row_keys: List[str] = []
        max_mismatches = 100
        if with_row_level:
            ck1, ck2 = st.columns([2, 1])
            with ck1:
                keys_str = st.text_input(
                    "Row key column(s) — comma-separated",
                    placeholder="e.g. claim_id,claim_line_number",
                    key="cmp_row_keys",
                )
                row_keys = [k.strip() for k in keys_str.split(",") if k.strip()]
            with ck2:
                max_mismatches = st.number_input(
                    "Max sample mismatches",
                    min_value=10,
                    max_value=10_000,
                    value=100,
                    step=10,
                    key="cmp_max_mm",
                )

    cmp_sampling_mode, cmp_sample_n, cmp_stratify = _sampling_section("cmp")

    st.divider()

    # ---- Output ----
    st.subheader("3. Output destination")
    cmp_vol_ref, cmp_out_cat, cmp_out_sch, cmp_run_label = _output_section("cmp")

    st.divider()

    # ---- Validate + Run ----
    st.subheader("4. Run")
    cmp_col_v, cmp_col_r, _ = st.columns([1, 1, 3])
    cmp_validate = cmp_col_v.button(
        "Validate", type="secondary", key="cmp_validate_btn"
    )
    cmp_run = cmp_col_r.button(
        "Run compare",
        type="primary",
        disabled=not st.session_state.get("cmp_validated", False),
        key="cmp_run_btn",
    )
    cmp_status = st.container()

    # -- Validation --
    def _validate_compare() -> tuple[bool, List[str]]:
        msgs: List[str] = []
        ok = True
        for side, label in [(side_a, "A"), (side_b, "B")]:
            if not all(
                [side["connection"], side["catalog"], side["schema"], side["table"]]
            ):
                msgs.append(
                    f"❌ Side {label}: catalog/schema/table not fully selected."
                )
                ok = False
        if cmp_vol_ref is None:
            msgs.append("❌ Output volume not selected.")
            ok = False
        if with_row_level and not row_keys:
            msgs.append("❌ Row-level diff selected but no key columns provided.")
            ok = False
        if not ok:
            return ok, msgs
        for side, label in [(side_a, "A"), (side_b, "B")]:
            ref = TableRef(
                connection=side["connection"].name,
                catalog=side["catalog"],
                schema=side["schema"],
                table=side["table"],
            )
            try:
                rc = describe_table(ref)
                msgs.append(
                    f"✅ Side {label}: `{ref.fqn}` reachable"
                    + (f" — {rc:,} rows" if rc is not None else "")
                )
                side["_row_count"] = rc
            except Exception as exc:  # noqa: BLE001
                ok = False
                msgs.append(f"❌ Side {label}: `{ref.fqn}` — {exc}")
        return ok, msgs

    if cmp_validate:
        with cmp_status:
            with st.spinner("Validating…"):
                ok, msgs = _validate_compare()
            for m in msgs:
                st.markdown(m)
            st.session_state["cmp_validated"] = ok
            if ok:
                st.success("Ready to run.")
            else:
                st.error("Fix issues above.")

    # -- Run --
    if cmp_run and st.session_state.get("cmp_validated", False):
        with cmp_status:
            t0 = time.time()
            try:
                folder = make_run_folder(
                    output=cmp_vol_ref,
                    side_a_env=side_a["env_label"],
                    side_b_env=side_b["env_label"],
                    table_name_a=side_a["table"],
                    table_name_b=side_b["table"],
                    run_label=cmp_run_label or None,
                )
                ensure_run_folder(folder)

                # Immediate write test — confirms volume is writable before
                # spending minutes on profiling.
                try:
                    write_text(folder, "_write_test.txt", "ok")
                    st.caption(f"✅ Volume write OK → `{folder.path}`")
                except Exception as exc:
                    st.error(f"❌ Cannot write to volume: {exc}")
                    raise

                _sample_n = (
                    cmp_sample_n if cmp_sampling_mode == "Sample N rows" else None
                )

                t_profile = time.time()
                st.caption("⏳ Step 1/5: profiling Side A …")
                try:
                    dataset_a = profile_table(
                        ref=TableRef(
                            connection=side_a["connection"].name,
                            catalog=side_a["catalog"],
                            schema=side_a["schema"],
                            table=side_a["table"],
                        ),
                        env_label=side_a["env_label"],
                        folder=folder,
                        html_filename="profile_a.html",
                        sample_n=_sample_n,
                        load_date=side_a.get("load_date"),
                    )
                    st.caption(
                        f"✅ Step 1/5: Side A — {dataset_a.row_count:,} rows, {dataset_a.column_count} cols"
                    )
                except Exception as exc:
                    st.error(f"❌ Step 1/5 failed — {exc}")
                    st.code(traceback.format_exc(), language="python")
                    raise

                st.caption("⏳ Step 2/5: profiling Side B …")
                try:
                    dataset_b = profile_table(
                        ref=TableRef(
                            connection=side_b["connection"].name,
                            catalog=side_b["catalog"],
                            schema=side_b["schema"],
                            table=side_b["table"],
                        ),
                        env_label=side_b["env_label"],
                        folder=folder,
                        html_filename="profile_b.html",
                        sample_n=_sample_n,
                        load_date=side_b.get("load_date"),
                    )
                    st.caption(
                        f"✅ Step 2/5: Side B — {dataset_b.row_count:,} rows, {dataset_b.column_count} cols"
                    )
                except Exception as exc:
                    st.error(f"❌ Step 2/5 failed — {exc}")
                    st.code(traceback.format_exc(), language="python")
                    raise

                t_profiled = time.time() - t_profile

                st.caption("⏳ Step 3/5: schema diff + drift metrics …")
                comparisons = compare_tables(dataset_a, dataset_b)
                st.caption("✅ Step 3/5: schema diff + drift complete")

                # ── Optional row-level diff ───────────────────────────────
                row_diff_result: Optional[RowDiffResult] = None
                if with_row_level and row_keys:
                    st.caption(f"⏳ Step 3b/5: row-level diff on keys {row_keys} …")
                    ref_a = TableRef(
                        connection=side_a["connection"].name,
                        catalog=side_a["catalog"],
                        schema=side_a["schema"],
                        table=side_a["table"],
                    )
                    ref_b = TableRef(
                        connection=side_b["connection"].name,
                        catalog=side_b["catalog"],
                        schema=side_b["schema"],
                        table=side_b["table"],
                    )
                    row_diff_result = compute_row_diff(
                        ref_a, ref_b, row_keys, max_mismatches
                    )
                    if row_diff_result.error:
                        st.warning(f"Row diff warning: {row_diff_result.error}")
                    else:
                        st.caption(
                            f"✅ Step 3b/5: row diff — "
                            f"{row_diff_result.rows_only_in_a:,} removed, "
                            f"{row_diff_result.rows_only_in_b:,} added, "
                            f"{row_diff_result.rows_changed:,} changed"
                        )

                profiler_run = ProfilerRun(
                    run_id=new_run_id(),
                    run_label=cmp_run_label or None,
                    created_utc=datetime.now(timezone.utc),
                    side_a=dataset_a,
                    side_b=dataset_b,
                    comparisons=comparisons,
                    lineage=Lineage(
                        manifest="manifest.json",
                        html_profile_a="profile_a.html",
                        html_profile_b="profile_b.html",
                        excel_summary="ab_summary.xlsx",
                    ),
                )

                st.caption("⏳ Step 4/5: writing artifacts …")
                with st.spinner(
                    "Writing Excel workbook, metamodel, and Mermaid diagrams …"
                ):
                    write_workbook(folder, profiler_run, row_diff=row_diff_result)
                    write_metamodel(folder, profiler_run)
                    write_json_schema(folder)
                    write_mermaid_diagrams(folder, profiler_run)
                    st.caption(f"✅ Artifacts written to `{folder.path}`")

                manifest = new_manifest(
                    run_id=folder.run_id,
                    run_label=cmp_run_label or None,
                    side_a=SideSpec(
                        env_label=side_a["env_label"],
                        connection=side_a["connection"].name,
                        connection_type=side_a["connection"].type,
                        catalog=side_a["catalog"],
                        schema=side_a["schema"],
                        table=side_a["table"],
                        row_count=dataset_a.row_count,
                    ),
                    side_b=SideSpec(
                        env_label=side_b["env_label"],
                        connection=side_b["connection"].name,
                        connection_type=side_b["connection"].type,
                        catalog=side_b["catalog"],
                        schema=side_b["schema"],
                        table=side_b["table"],
                        row_count=dataset_b.row_count,
                    ),
                    comparison=ComparisonParams(
                        depth="with_row_level" if with_row_level else "aggregate_only",
                        row_keys=row_keys,
                        max_sample_mismatches=int(max_mismatches),
                        sampling_mode={
                            "Full table": "full",
                            "Sample N rows": "sample_n",
                            "Stratified by column": "stratified",
                        }[cmp_sampling_mode],
                        sample_n=_sample_n,
                        stratify_by=cmp_stratify or None,
                    ),
                    output_folder=folder.path,
                )
                manifest.add_timing("setup", time.time() - t0 - t_profiled)
                manifest.add_timing("profiling", t_profiled)
                for art in [
                    "profile_a.html",
                    "profile_b.html",
                    "ab_summary.xlsx",
                    "metamodel.json",
                    "schema_a.mmd",
                    "schema_b.mmd",
                    "drift.mmd",
                ]:
                    manifest.add_artifact(art)
                write_json(folder, "manifest.json", manifest.to_dict())

                if os.environ.get("PROFILER_RUNTIME", "mock").lower() == "databricks":
                    try:
                        from profiler import delta_repo

                        delta_repo.ensure_tables(cmp_out_cat, cmp_out_sch)
                        delta_repo.ingest(profiler_run, cmp_out_cat, cmp_out_sch)
                        st.caption(
                            f"✅ Governance tables updated in `{cmp_out_cat}.{cmp_out_sch}`"
                        )
                    except RuntimeError as exc:
                        # Grant errors surface here — ingest may have succeeded
                        msg = str(exc)
                        if "GRANT" in msg.upper() or "grant" in msg:
                            st.warning(
                                f"⚠️ Tables written but grants need admin help:\n{msg}"
                            )
                        else:
                            st.warning(f"Delta repo issue — {exc}")
                    except Exception as exc:  # noqa: BLE001
                        st.warning(f"Delta repo ingest skipped — {exc}")
                else:
                    st.caption(
                        f"⚠️ PROFILER_RUNTIME={os.environ.get('PROFILER_RUNTIME', 'NOT SET')} "
                        f"— Delta tables skipped (expected 'databricks')"
                    )

                st.success(
                    f"Compare run complete in {time.time() - t0:.1f}s — `{folder.folder_name}`"
                )
                _render_run_outputs(
                    folder, profiler_run, mode="compare", row_diff=row_diff_result
                )

            except Exception:  # noqa: BLE001
                st.error("Run failed.")
                st.code(traceback.format_exc(), language="python")


# ===========================================================================
# TAB 2 — PROFILE
# ===========================================================================

with tab_profile:
    st.markdown(
        "Profile one or more tables independently. "
        "Produces HTML reports, column-level DQ stats, alerts, and metamodel JSON. "
        "No cross-table comparison."
    )
    st.divider()

    # ---- Dynamic table list ----
    st.subheader("1. Pick tables to profile")

    if "n_profile_tables" not in st.session_state:
        st.session_state["n_profile_tables"] = 1

    n_tables = st.session_state["n_profile_tables"]
    profile_sides: List[dict] = []

    labels = _env_labels()
    for i in range(n_tables):
        default_idx = min(i, len(labels) - 1)
        with st.expander(f"Table {i + 1}", expanded=True):
            profile_sides.append(_table_picker(f"prf_{i}", default_env_idx=default_idx))

    btn_col1, btn_col2, _ = st.columns([1, 1, 4])
    if btn_col1.button("+ Add table", key="prf_add"):
        st.session_state["n_profile_tables"] += 1
        st.rerun()
    if btn_col2.button(
        "− Remove last",
        key="prf_rem",
        disabled=st.session_state["n_profile_tables"] <= 1,
    ):
        st.session_state["n_profile_tables"] -= 1
        st.rerun()

    st.divider()

    # ---- Settings ----
    st.subheader("2. Settings")
    prf_sampling_mode, prf_sample_n, prf_stratify = _sampling_section("prf")

    st.divider()

    # ---- Output ----
    st.subheader("3. Output destination")
    prf_vol_ref, prf_out_cat, prf_out_sch, prf_run_label = _output_section("prf")

    st.divider()

    # ---- Validate + Run ----
    st.subheader("4. Run")
    prf_col_v, prf_col_r, _ = st.columns([1, 1, 3])
    prf_validate = prf_col_v.button(
        "Validate", type="secondary", key="prf_validate_btn"
    )
    prf_run = prf_col_r.button(
        "Run profile",
        type="primary",
        disabled=not st.session_state.get("prf_validated", False),
        key="prf_run_btn",
    )
    prf_status = st.container()

    # -- Validation --
    def _validate_profile() -> tuple[bool, List[str]]:
        msgs: List[str] = []
        ok = True
        for i, side in enumerate(profile_sides):
            if not all(
                [side["connection"], side["catalog"], side["schema"], side["table"]]
            ):
                msgs.append(f"❌ Table {i + 1}: not fully selected.")
                ok = False
        if prf_vol_ref is None:
            msgs.append("❌ Output volume not selected.")
            ok = False
        if not ok:
            return ok, msgs
        for i, side in enumerate(profile_sides):
            ref = TableRef(
                connection=side["connection"].name,
                catalog=side["catalog"],
                schema=side["schema"],
                table=side["table"],
            )
            try:
                rc = describe_table(ref)
                msgs.append(
                    f"✅ Table {i + 1}: `{ref.fqn}` reachable"
                    + (f" — {rc:,} rows" if rc is not None else "")
                )
                side["_row_count"] = rc
            except Exception as exc:  # noqa: BLE001
                ok = False
                msgs.append(f"❌ Table {i + 1}: `{ref.fqn}` — {exc}")
        return ok, msgs

    if prf_validate:
        with prf_status:
            with st.spinner("Validating…"):
                ok, msgs = _validate_profile()
            for m in msgs:
                st.markdown(m)
            st.session_state["prf_validated"] = ok
            if ok:
                st.success("Ready to run.")
            else:
                st.error("Fix issues above.")

    # -- Run --
    if prf_run and st.session_state.get("prf_validated", False):
        with prf_status:
            t0 = time.time()
            try:
                _sample_n = (
                    prf_sample_n if prf_sampling_mode == "Sample N rows" else None
                )

                for i, side in enumerate(profile_sides):
                    tbl_label = f"{side['catalog']}.{side['schema']}.{side['table']}"
                    st.markdown(f"---\n**Table {i + 1} of {n_tables}: `{tbl_label}`**")

                    folder = make_run_folder(
                        output=prf_vol_ref,
                        side_a_env=side["env_label"],
                        side_b_env=side["env_label"],
                        table_name_a=side["table"],
                        table_name_b=side["table"],
                        run_label=prf_run_label or None,
                    )
                    ensure_run_folder(folder)

                    with st.spinner(f"Profiling `{tbl_label}` …"):
                        dataset = profile_table(
                            ref=TableRef(
                                connection=side["connection"].name,
                                catalog=side["catalog"],
                                schema=side["schema"],
                                table=side["table"],
                            ),
                            env_label=side["env_label"],
                            folder=folder,
                            html_filename="profile_a.html",
                            sample_n=_sample_n,
                            load_date=side.get("load_date"),
                        )

                    # Build a single-side ProfilerRun (side_b mirrors side_a,
                    # comparisons empty — no cross-table diff in profile mode).
                    profiler_run = ProfilerRun(
                        run_id=new_run_id(),
                        run_label=prf_run_label or None,
                        created_utc=datetime.now(timezone.utc),
                        side_a=dataset,
                        side_b=dataset,
                        comparisons=[],
                        lineage=Lineage(
                            manifest="manifest.json",
                            html_profile_a="profile_a.html",
                        ),
                    )

                    with st.spinner("Writing metamodel and schema diagram …"):
                        write_metamodel(folder, profiler_run)
                        write_json_schema(folder)
                        # Only the Side-A schema diagram is meaningful in profile mode
                        from profiler.mermaid import render_side_schema

                        schema_mmd = render_side_schema(dataset)
                        write_text(folder, "schema_a.mmd", schema_mmd)

                    if (
                        os.environ.get("PROFILER_RUNTIME", "mock").lower()
                        == "databricks"
                    ):
                        try:
                            from profiler import delta_repo

                            delta_repo.ensure_tables(prf_out_cat, prf_out_sch)
                            delta_repo.ingest(profiler_run, prf_out_cat, prf_out_sch)
                        except Exception as exc:  # noqa: BLE001
                            st.warning(f"Delta repo ingest skipped: {exc}")

                    st.success(
                        f"Table {i + 1} profiled — "
                        f"{dataset.row_count:,} rows, {dataset.column_count} columns, "
                        f"{sum(len(c.alerts) for c in dataset.columns)} alerts"
                    )
                    _render_run_outputs(folder, profiler_run, mode="profile")

                st.success(
                    f"All {n_tables} table(s) profiled in {time.time() - t0:.1f}s."
                )

            except Exception:  # noqa: BLE001
                st.error("Run failed.")
                st.code(traceback.format_exc(), language="python")


# ===========================================================================
# TAB 3 — ASK GENIE
# ===========================================================================

with tab_genie:
    st.markdown(
        "Ask natural-language questions about your profiling runs, drift metrics, "
        "and data quality alerts. Genie queries the governance tables and returns "
        "an answer, the SQL it ran, and the result set."
    )

    if not _GENIE_SPACE_ID:
        st.warning(
            "**Genie Space ID not configured.**  \n"
            "Add `GENIE_SPACE_ID` to `app.yaml` → `env` section, then redeploy:\n"
            '```yaml\n- name: GENIE_SPACE_ID\n  value: "<your-space-id>"\n```\n'
            "Find your Space ID in the Databricks UI: "
            "**AI/BI → Genie Spaces → your space → URL contains the ID.**"
        )
    else:
        from profiler.genie_chat import ask, GenieResult

        # Initialise conversation state
        if "genie_conv_id" not in st.session_state:
            st.session_state["genie_conv_id"] = None
        if "genie_messages" not in st.session_state:
            st.session_state["genie_messages"] = []

        # Controls row
        gcol1, gcol2 = st.columns([6, 1])
        with gcol2:
            if st.button("Clear chat", key="genie_clear"):
                st.session_state["genie_conv_id"] = None
                st.session_state["genie_messages"] = []
                st.rerun()

        with gcol1:
            st.caption(
                f"Space: `{_GENIE_SPACE_ID[:8]}...`  "
                + (
                    "| Conversation active"
                    if st.session_state["genie_conv_id"]
                    else "| New conversation"
                )
            )

        # Render conversation history
        for msg in st.session_state["genie_messages"]:
            with st.chat_message(
                msg["role"], avatar="🧑" if msg["role"] == "user" else "🤖"
            ):
                st.markdown(msg["content"])
                if msg.get("sql"):
                    with st.expander("Generated SQL", expanded=False):
                        st.code(msg["sql"], language="sql")
                if msg.get("rows") and msg.get("col_names"):
                    import pandas as _pd

                    st.dataframe(
                        _pd.DataFrame(msg["rows"], columns=msg["col_names"]),
                        use_container_width=True,
                        height=min(300, 38 + len(msg["rows"]) * 35),
                    )

        # Suggested starter questions
        if not st.session_state["genie_messages"]:
            st.markdown("**Try asking:**")
            suggestions = [
                "How many profiler runs have completed this month?",
                "Which columns have the highest PSI drift across all runs?",
                "Show me the top 10 columns with critical alerts.",
                "What schema changes were detected in the last 7 days?",
                "Compare average PSI for medical_claim vs lab_result tables.",
            ]
            for s in suggestions:
                if st.button(s, key=f"genie_suggestion_{s[:20]}"):
                    st.session_state["genie_pending_question"] = s
                    st.rerun()

        # Pick up a suggestion button click carried from previous render
        if "genie_pending_question" in st.session_state:
            prompt = st.session_state.pop("genie_pending_question")
        else:
            prompt = st.chat_input(
                "Ask Genie about your profiling data…",
                key="genie_input",
            )

        if prompt:
            # Render user bubble immediately
            with st.chat_message("user", avatar="🧑"):
                st.markdown(prompt)
            st.session_state["genie_messages"].append(
                {
                    "role": "user",
                    "content": prompt,
                }
            )

            # Render Genie response bubble
            with st.chat_message("assistant", avatar="🤖"):
                with st.spinner("Genie is thinking…"):
                    try:
                        result, conv_id, _ = ask(
                            _GENIE_SPACE_ID,
                            prompt,
                            conv_id=st.session_state["genie_conv_id"],
                        )
                        st.session_state["genie_conv_id"] = conv_id

                        if result.error:
                            response_text = f"⚠️ {result.error}"
                        else:
                            response_text = (
                                result.text_response
                                or "*(Genie returned no text — check the Generated SQL tab below for the query result)*"
                            )

                        st.markdown(response_text)

                        if result.sql:
                            with st.expander("Generated SQL", expanded=True):
                                st.code(result.sql, language="sql")

                        if result.has_data:
                            import pandas as _pd

                            st.dataframe(
                                _pd.DataFrame(result.rows, columns=result.col_names),
                                use_container_width=True,
                                height=min(300, 38 + len(result.rows) * 35),
                            )

                        # Debug: show raw response structure when no text found
                        if not result.text_response and not result.error:
                            with st.expander(
                                "Debug — raw Genie response", expanded=False
                            ):
                                st.json(
                                    {
                                        "text_response": result.text_response,
                                        "sql": result.sql,
                                        "col_names": result.col_names,
                                        "row_count": len(result.rows),
                                    }
                                )

                        st.session_state["genie_messages"].append(
                            {
                                "role": "assistant",
                                "content": response_text,
                                "sql": result.sql,
                                "rows": result.rows,
                                "col_names": result.col_names,
                            }
                        )

                    except Exception as exc:  # noqa: BLE001
                        err_msg = f"❌ Genie error: {exc}"
                        st.error(err_msg)
                        st.session_state["genie_messages"].append(
                            {
                                "role": "assistant",
                                "content": err_msg,
                            }
                        )


# ===========================================================================
# TAB 4 — LOAD RESULTS
# ===========================================================================
# Nightly incremental load + GX validation results, persisted to Delta by the
# tablespec pipeline (db_load_persist.py). Governance tables live in the same
# schema as the profiler tables.

_LOAD_GOV = "dev.test_main_profiler"


def _num(v, default=0):
    """Coerce a warehouse string cell to int (results arrive as strings)."""
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


with tab_load:
    st.markdown(
        "Results of the **nightly incremental load** (local CSV → `raw_<t>` → "
        "`ingested_<t>`) and its Great Expectations validation, persisted per run "
        "for trending."
    )
    st.divider()

    from profiler.catalog import _runtime as _load_runtime

    if _load_runtime() != "databricks":
        st.info("Load Results are only available in Databricks mode.")
    else:
        try:
            runs = _query_df(
                "SELECT run_id, run_ts, run_date, target_schema, table_count, "
                "total_file_rows, total_ingested_rows, tables_with_drift, "
                "validation_errors, validation_warnings, "
                "total_promoted_inserted, total_promoted_updated, promoted, status "
                f"FROM {_LOAD_GOV}.load_runs ORDER BY run_ts DESC LIMIT 50"
            )
        except Exception as exc:  # noqa: BLE001
            st.warning(
                "Could not read load results. The tables are created on the first "
                f"run of db_load_persist.py.\n\nDetails: {exc}"
            )
            runs = None

        if runs is None or runs.empty:
            st.info("No load runs recorded yet. Run `db_load_persist.py` to populate.")
        else:
            # ---- Run picker -------------------------------------------------
            def _run_label(i: int) -> str:
                r = runs.iloc[i]
                return (
                    f"{r['run_ts']}  •  {r['status']}  •  {_num(r['table_count'])} tables  "
                    f"(err {_num(r['validation_errors'])}, warn {_num(r['validation_warnings'])})"
                )

            idx = st.selectbox(
                "Run", options=list(range(len(runs))), format_func=_run_label, index=0
            )
            run = runs.iloc[idx]
            run_id = run["run_id"]

            # ---- Summary metrics -------------------------------------------
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Tables", _num(run["table_count"]))
            c2.metric("File rows", f"{_num(run['total_file_rows']):,}")
            c3.metric("Ingested rows", f"{_num(run['total_ingested_rows']):,}")
            c4.metric("Tables w/ drift", _num(run["tables_with_drift"]))
            c5.metric(
                "Validation err / warn",
                f"{_num(run['validation_errors'])} / {_num(run['validation_warnings'])}",
            )
            _promoted = str(run.get("promoted", "")).lower() in ("true", "1")
            if _promoted:
                c6.metric(
                    "Promoted ins / upd",
                    f"{_num(run['total_promoted_inserted']):,} / "
                    f"{_num(run['total_promoted_updated']):,}",
                )
            else:
                c6.metric("Promoted", "—", help="Promotion was skipped for this run.")

            file_rows = _num(run["total_file_rows"])
            ing_rows = _num(run["total_ingested_rows"])
            if file_rows and ing_rows != file_rows:
                st.warning(
                    f"Row-count mismatch: {file_rows:,} file rows vs "
                    f"{ing_rows:,} ingested — possible dropped/duplicated rows."
                )

            # ---- Per-table results -----------------------------------------
            st.subheader("Per-table results")
            tbl = _query_df(
                "SELECT table, status, file_rows, raw_rows, ingested_rows, "
                "dropped_cols, null_filled_cols, validation_errors, validation_warnings, "
                "target_before, promoted_inserted, promoted_updated, target_after "
                f"FROM {_LOAD_GOV}.load_table_results WHERE run_id = '{run_id}' "
                "ORDER BY table"
            )
            st.dataframe(tbl, use_container_width=True, hide_index=True)

            drift = (
                tbl[tbl["dropped_cols"].notna() | tbl["null_filled_cols"].notna()]
                if not tbl.empty
                else tbl
            )
            if not drift.empty:
                st.caption(
                    "⚠️ Schema drift detected on: " + ", ".join(drift["table"].tolist())
                )

            # ---- Validation issues -----------------------------------------
            st.subheader("Validation issues")
            issues = _query_df(
                "SELECT table, severity, error_type, column_name, rule_name, message "
                f"FROM {_LOAD_GOV}.load_validation_results WHERE run_id = '{run_id}' "
                "ORDER BY severity, table"
            )
            if issues.empty:
                st.success("No validation issues for this run.")
            else:
                sevs = sorted(issues["severity"].dropna().unique().tolist())
                chosen = st.multiselect("Severity", options=sevs, default=sevs)
                shown = issues[issues["severity"].isin(chosen)] if chosen else issues
                st.dataframe(shown, use_container_width=True, hide_index=True)

            # ---- Trend across runs -----------------------------------------
            st.subheader("Trend across runs")
            trend = runs[
                [
                    "run_ts",
                    "total_ingested_rows",
                    "validation_errors",
                    "validation_warnings",
                ]
            ].copy()
            for col in (
                "total_ingested_rows",
                "validation_errors",
                "validation_warnings",
            ):
                trend[col] = trend[col].apply(_num)
            trend = (
                trend.rename(
                    columns={
                        "total_ingested_rows": "Ingested rows",
                        "validation_errors": "Errors",
                        "validation_warnings": "Warnings",
                    }
                )
                .set_index("run_ts")
                .sort_index()
            )
            st.line_chart(trend[["Errors", "Warnings"]])
            st.line_chart(trend[["Ingested rows"]])

            prom_trend = runs[
                ["run_ts", "total_promoted_inserted", "total_promoted_updated"]
            ].copy()
            for col in ("total_promoted_inserted", "total_promoted_updated"):
                prom_trend[col] = prom_trend[col].apply(_num)
            prom_trend = (
                prom_trend.rename(
                    columns={
                        "total_promoted_inserted": "Promoted (inserted)",
                        "total_promoted_updated": "Promoted (updated)",
                    }
                )
                .set_index("run_ts")
                .sort_index()
            )
            st.line_chart(prom_trend)


# ===========================================================================
# TAB 5 - GUIDEBOOK
# ===========================================================================
# Renders tablespec's static UMF guidebook inside this app, from either:
#   1. a UC Volume directory of UMFs written by the tablespec pipeline, or
#   2. Spark-free reflection of a catalog/schema via INFORMATION_SCHEMA.
#
# Databricks Apps have a SQL warehouse and the workspace SDK but NO
# SparkSession, so reflection goes through tablespec's
# umf_from_information_schema rather than SparkToUmfMapper/JdbcToUmfMapper,
# both of which require Spark.
#
# The guidebook is a multi-page static site whose pages link to one another.
# st.components.v1.html renders a srcdoc iframe in which those relative links
# cannot resolve, so a page selector replaces cross-page navigation. The full
# site is still downloadable as a zip, where the links work normally.

_GUIDEBOOK_HEIGHT = 900

_INFO_SCHEMA_COLS = (
    "column_name, data_type, is_nullable, character_maximum_length, "
    "numeric_precision, numeric_scale, comment, ordinal_position"
)


def _tablespec_import_error() -> str:
    """Return '' when tablespec is importable, else the failure message."""
    try:
        import tablespec  # noqa: F401
    except Exception as exc:  # noqa: BLE001
        return str(exc)
    return ""


def _collect_site(out_dir) -> dict:
    """Read a generated guidebook into {relative_path: text}."""
    files = {}
    for path in sorted(out_dir.rglob("*")):
        if path.is_file():
            files[path.relative_to(out_dir).as_posix()] = path.read_text(
                encoding="utf-8"
            )
    return files


def _zip_site(files: dict) -> bytes:
    import io
    import zipfile

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in files.items():
            archive.writestr(name, text)
    return buf.getvalue()


@st.cache_data(ttl=300, show_spinner=False)
def _guidebook_from_reflection(catalog: str, schema: str, tables: tuple) -> dict:
    """Reflect tables via INFORMATION_SCHEMA, then render the guidebook."""
    import tempfile
    from pathlib import Path as _Path

    from tablespec import generate_guidebook, save_umf_to_yaml
    from tablespec.profiling import umf_from_information_schema

    with tempfile.TemporaryDirectory() as tmp:
        umf_dir = _Path(tmp) / "umf"
        umf_dir.mkdir()
        for table in tables:
            rows = _query_df(
                f"SELECT {_INFO_SCHEMA_COLS} "
                f"FROM `{catalog}`.information_schema.columns "
                f"WHERE table_schema = '{schema}' AND table_name = '{table}' "
                "ORDER BY ordinal_position"
            ).to_dict("records")
            if not rows:
                continue
            umf = umf_from_information_schema(table, rows)
            save_umf_to_yaml(umf, umf_dir / f"{table}.umf.yaml")

        out_dir = _Path(tmp) / "guidebook"
        generate_guidebook(root=umf_dir, output_dir=out_dir)
        return _collect_site(out_dir)


def _list_volume_umfs(volume_dir: str) -> list:
    """Return (filename, full_path) for UMF artifacts directly under volume_dir."""
    from profiler.catalog import _runtime, _workspace_client

    suffixes = (".umf.yaml", ".umf.json")
    if _runtime() != "databricks":
        from pathlib import Path as _Path

        local = _Path(volume_dir)
        if not local.is_dir():
            return []
        return [
            (p.name, str(p))
            for p in sorted(local.iterdir())
            if p.name.endswith(suffixes)
        ]

    entries = _workspace_client().files.list_directory_contents(
        directory_path=volume_dir
    )
    return [
        (entry.name, entry.path)
        for entry in entries
        if (entry.name or "").endswith(suffixes)
    ]


@st.cache_data(ttl=300, show_spinner=False)
def _guidebook_from_volume(volume_dir: str) -> dict:
    """Download UMFs from a UC Volume, then render the guidebook."""
    import tempfile
    from pathlib import Path as _Path

    from profiler.storage import read_text
    from tablespec import generate_guidebook

    found = _list_volume_umfs(volume_dir)
    if not found:
        return {}

    with tempfile.TemporaryDirectory() as tmp:
        umf_dir = _Path(tmp) / "umf"
        umf_dir.mkdir()
        for name, path in found:
            (umf_dir / name).write_text(read_text(path), encoding="utf-8")

        out_dir = _Path(tmp) / "guidebook"
        generate_guidebook(root=umf_dir, output_dir=out_dir)
        return _collect_site(out_dir)


def _render_guidebook_site(files: dict) -> None:
    """Page selector + inline render + zip download for a generated site."""
    pages = sorted(name for name in files if name.endswith(".html"))
    if not pages:
        st.warning("The guidebook generated no pages - no UMFs were discovered.")
        return

    # index.html first; then table pages alphabetically.
    ordered = [p for p in pages if p == "index.html"] + [
        p for p in pages if p != "index.html"
    ]

    left, right = st.columns([3, 1])
    with left:
        selected = st.selectbox(
            "Page",
            options=ordered,
            format_func=lambda p: "Index" if p == "index.html" else p[:-5],
            key="guidebook_page",
            help="Links between pages do not work inside the embedded frame; "
            "use this selector, or download the site.",
        )
    with right:
        st.download_button(
            "Download site (.zip)",
            data=_zip_site(files),
            file_name="guidebook.zip",
            mime="application/zip",
            use_container_width=True,
        )

    st.caption(f"{len(pages)} page(s) generated.")
    components.html(files[selected], height=_GUIDEBOOK_HEIGHT, scrolling=True)


with tab_guidebook:
    st.markdown(
        "Render the **tablespec guidebook** - one page per table with columns, "
        "types, lineage, and validation rules - from UMF specs."
    )
    st.divider()

    _ts_error = _tablespec_import_error()
    if _ts_error:
        st.error(
            "tablespec is not importable in this app environment, so the "
            "guidebook cannot be rendered.\n\n"
            f"Details: {_ts_error}\n\n"
            "Ensure the app's requirements install the tablespec package "
            "(see docs/guide/data-profiling-app.md)."
        )
    else:
        _source = st.radio(
            "UMF source",
            options=["Reflect from catalog", "From UC Volume"],
            horizontal=True,
            key="guidebook_source",
            help="Reflection reads INFORMATION_SCHEMA via the SQL warehouse "
            "(no Spark required). Volume reads UMFs written by the pipeline.",
        )

        _site = None

        if _source == "Reflect from catalog":
            _conn = _connections()[0] if _connections() else None
            try:
                _catalogs = list_catalogs(_conn) if _conn else []
            except Exception as exc:  # noqa: BLE001
                st.error(f"Cannot load catalogs: {exc}")
                _catalogs = []

            _c1, _c2 = st.columns(2)
            with _c1:
                _gb_catalog = st.selectbox(
                    "Catalog", options=_catalogs, key="guidebook_catalog"
                )
            with _c2:
                _gb_schemas = _cached_schemas(_gb_catalog) if _gb_catalog else []
                _gb_schema = st.selectbox(
                    "Schema", options=_gb_schemas, key="guidebook_schema"
                )

            _gb_tables = (
                _cached_tables(_gb_catalog, _gb_schema)
                if (_gb_catalog and _gb_schema)
                else []
            )
            _picked = st.multiselect(
                "Tables",
                options=_gb_tables,
                default=_gb_tables[:10],
                key="guidebook_tables",
                help="Each table is reflected from INFORMATION_SCHEMA into a UMF.",
            )

            if st.button("Build guidebook", type="primary", key="guidebook_build"):
                if not _picked:
                    st.warning("Pick at least one table.")
                else:
                    with st.spinner(f"Reflecting {len(_picked)} table(s)..."):
                        try:
                            _site = _guidebook_from_reflection(
                                _gb_catalog, _gb_schema, tuple(_picked)
                            )
                            st.session_state["guidebook_site"] = _site
                        except Exception as exc:  # noqa: BLE001
                            st.error(f"Guidebook generation failed: {exc}")

        else:
            _vol = st.text_input(
                "Volume directory",
                value=st.session_state.get(
                    "guidebook_volume", "/Volumes/dev/test_main_profiler/ab_runs/umf"
                ),
                key="guidebook_volume",
                help="Directory holding *.umf.yaml or *.umf.json artifacts.",
            )
            if st.button(
                "Build guidebook from Volume", type="primary", key="guidebook_build_vol"
            ):
                with st.spinner(f"Reading UMFs from {_vol}..."):
                    try:
                        _site = _guidebook_from_volume(_vol)
                        if not _site:
                            st.warning(
                                f"No *.umf.yaml or *.umf.json files found in {_vol}."
                            )
                        else:
                            st.session_state["guidebook_site"] = _site
                    except Exception as exc:  # noqa: BLE001
                        st.error(f"Could not read {_vol}: {exc}")

        _cached_site = st.session_state.get("guidebook_site")
        if _cached_site:
            st.divider()
            _render_guidebook_site(_cached_site)
