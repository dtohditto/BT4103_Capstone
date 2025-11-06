import os
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import CSVCuration
import io
from typing import List, Tuple
from datetime import datetime
from zoneinfo import ZoneInfo
import copy
import plotly.io as pio
import time
import traceback
from io import BytesIO

# Prefer the new defaults API; fall back silently if not available
try:
    pio.defaults.to_image.format = "png"
except Exception:
    try:
        # Legacy (still works for now)
        pio.kaleido.scope.default_format = "png"
    except Exception:
        pass

st.set_page_config(page_title="EE Analytics Dashboard", layout="wide")

# ──────────────────────────────────────────────────────────────────────────────
# Data loading & basic preparation
# ──────────────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data(src) -> pd.DataFrame:
    """Accepts a DataFrame or a file-like CSV and returns a cleaned DataFrame."""
    df = src.copy() if isinstance(src, pd.DataFrame) else pd.read_csv(src)

    # Dates
    for c in ["Programme Start Date", "Programme End Date", "Run_Month"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Month label for charts
    if "Run_Month" in df.columns:
        df["Run_Month_Label"] = df["Run_Month"].dt.strftime("%Y-%m")

    # Age buckets
    if "Age" in df.columns:
        age_num = pd.to_numeric(df["Age"], errors="coerce")
        bins    = [0, 34, 44, 54, 64, 200]
        labels  = ["<35", "35–44", "45–54", "55–64", "65+"]
        df["Age_Group"] = pd.cut(age_num, bins=bins, labels=labels, right=True).astype("string")

    # Light tidy-up for common filters
    cat_cols = [
        "Application Status", "Applicant Type", "Primary Category",
        "Secondary Category", "Seniority", "Gender", "Country Of Residence",
        "Truncated Programme Name", "Domain"
    ]
    for c in cat_cols:
        if c in df.columns:
            df[c] = df[c].astype("string").str.strip()

    return df


# ──────────────────────────────────────────────────────────────────────────────
# Uploads & curation (sidebar)
# ──────────────────────────────────────────────────────────────────────────────
st.sidebar.header("Data")

st.sidebar.info(
    "**Choose ONE path:**\n\n"
    "1) **Path A – Curated CSV ready**\n"
    "   Upload your curated CSV below (no need to click **Run curation**).\n\n"
    "2) **Path B – Start from raw files**\n"
    "   Upload the **Programme** file and the **Cost** file, then click **Run curation**.\n"
    "   When it finishes, **download the curated CSV and upload it** under **Curated CSV** below."
)

# Path B — raw files
st.sidebar.subheader("Path B — Upload raw files")
new_uploaded_programme = st.sidebar.file_uploader(
    "Programme file (CSV/XLSX/XLSM/XLS)", type=["csv", "xlsx", "xlsm", "xls"], key="prog_upload"
)
new_uploaded_cost = st.sidebar.file_uploader(
    "Cost file (CSV/XLSX/XLSM/XLS)", type=["csv", "xlsx", "xlsm", "xls"], key="cost_upload"
)

run = st.sidebar.button("▶️ Run curation", use_container_width=True)

df_curated = None
csv_bytes = None

if run:
    if not new_uploaded_programme or not new_uploaded_cost:
        st.sidebar.error("Please upload **both** Programme **and** Cost files before running curation.")
    else:
        res = CSVCuration.curate_programme_and_cost_data(
            new_uploaded_programme, new_uploaded_cost, return_csv_bytes=True
        )
        if res is None:
            st.sidebar.warning("No curated output produced. Please check your inputs.")
        elif isinstance(res, tuple):
            df_curated, csv_bytes = res
        elif isinstance(res, (bytes, bytearray)):
            csv_bytes = bytes(res)
        elif isinstance(res, pd.DataFrame):
            df_curated = res
            try:
                csv_bytes = df_curated.to_csv(index=False).encode("utf-8-sig")
            except Exception:
                csv_bytes = None

# Curated CSV download (after curation)
if csv_bytes is not None:
    st.sidebar.success(
        "Curation complete. **Step 2:** Download the curated CSV, then upload it under **Path A → Curated CSV**."
    )
    st.sidebar.download_button(
        "💾 Download curated CSV",
        data=csv_bytes,
        file_name="dashboard_curated.csv",
        mime="text/csv",
        use_container_width=True,
    )
else:
    st.sidebar.caption(
        "After curation completes, a download button will appear here. "
        "Then upload the curated CSV under **Path A** below."
    )

st.sidebar.divider()

# Path A — curated CSV
st.sidebar.subheader("Path A — Or upload a curated CSV")
uploaded_curated = st.sidebar.file_uploader("Curated CSV", type=["csv"], key="curated_upload")

# Pick data source (no local file fallback)
data_src = uploaded_curated if uploaded_curated is not None else None

if data_src is None:
    st.info(
        "**No data loaded.** Either upload a curated CSV (Path A), **or** upload Programme & Cost and run curation, "
        "then upload the curated CSV (Path B)."
    )
    st.stop()

# One-time load
df_full = load_data(data_src)
df = df_full.copy()

if df.empty:
    st.info("No data found. Please upload a curated CSV to continue.")
    st.stop()

st.title("Executive Education Analytics Dashboard")

# ──────────────────────────────────────────────────────────────────────────────
# Helpers & shared settings
# ──────────────────────────────────────────────────────────────────────────────
UNKNOWN_LIKE = {
    "unknown", "unspecified", "not specified", "not provided", "not available",
    "n/a", "na", "null", "none", "-", "", "Others"
}

def _safe_key(label: str, suffix: str) -> str:
    return f"{label}_{suffix}".replace(" ", "_").lower()

# UI label ↔ column map
COL_MAP = {
    "Pri Category": "Primary Category",
    "Sec Category": "Secondary Category",
    "Country": "Country Of Residence",
    "Application Status": "Application Status",
    "Applicant Type": "Applicant Type",
    "Seniority": "Seniority",
    "Domain": "Domain",
}
UI_FILTER_LABELS = [
    "Application Status", "Applicant Type", "Pri Category",
    "Sec Category", "Country", "Seniority", "Domain"
]

def _col_from_label(label: str) -> str:
    return COL_MAP.get(label, label)

def multiselect_with_all_button(label: str, df_source: pd.DataFrame, default_all: bool = True):
    """Multiselect with a quick 'Select all' button."""
    col = _col_from_label(label)
    raw = df_source.get(col, pd.Series([], dtype="object")).copy()
    s = raw.astype("string").fillna("Unknown") if pd.api.types.is_categorical_dtype(raw) else raw.fillna("Unknown")
    options = sorted(s.unique().tolist())

    ms_key = _safe_key(label, "multi")
    if ms_key not in st.session_state:
        st.session_state[ms_key] = options[:] if default_all else []

    current = [v for v in st.session_state[ms_key] if v in options]
    if current != st.session_state[ms_key]:
        st.session_state[ms_key] = current

    if st.button(f"Select all {label}", key=_safe_key(label, "btn_all")):
        st.session_state[ms_key] = options[:]

    st.multiselect(label, options, key=ms_key)
    return st.session_state[ms_key]

def apply_filter(series: pd.Series, selected: list[str]) -> pd.Series:
    """Return a boolean mask for the selected options (treat missing as 'Unknown')."""
    if selected is None:
        return pd.Series(True, index=series.index)
    s = series.astype("string").fillna("Unknown") if pd.api.types.is_categorical_dtype(series) else series.fillna("Unknown")
    all_opts = set(s.unique().tolist())
    sel_set  = set(selected or [])
    if not selected or sel_set == all_opts:
        return pd.Series(True, index=series.index)
    return s.isin(selected)

# “Unknown” helpers
def _norm_str(s: pd.Series) -> pd.Series:
    return s.astype("string").str.strip().str.lower()

def coalesce_unknown(series: pd.Series) -> pd.Series:
    s = series.astype("string").str.strip()
    s_norm = s.str.lower()
    mask_u = s.isna() | (s == "") | s_norm.isin(UNKNOWN_LIKE)
    return s.mask(mask_u, "Unknown")

def filter_unknown_no_ui(df_in: pd.DataFrame, column: str, include_unknown: bool) -> pd.DataFrame:
    if column not in df_in.columns:
        return df_in
    s_norm = df_in[column].astype("string").str.strip().str.lower()
    mask = df_in[column].isna() | s_norm.isin(UNKNOWN_LIKE)
    return df_in.copy() if include_unknown else df_in.loc[~mask].copy()

def dq_caption(df_in: pd.DataFrame, column: str, label: str):
    """Compact note showing Unknown + Missing share for a column."""
    if column not in df_in.columns:
        return
    s_norm = df_in[column].astype("string").str.strip().str.lower()
    mask = df_in[column].isna() | s_norm.isin(UNKNOWN_LIKE)
    total = len(df_in)
    unknown_combined = int(mask.sum())
    valid = int(total - unknown_combined)
    pct = (lambda n: (n / total * 100.0) if total > 0 else 0.0)
    st.caption(
        f"**Data quality — {label}:** Unknown + Missing **{unknown_combined}** ({pct(unknown_combined):.1f}%), "
        f"Valid **{valid}** ({pct(valid):.1f}%)."
    )

def dq_note_only(df_in: pd.DataFrame, column: str, label: str):
    """Same as above, without a checkbox control."""
    if column not in df_in.columns:
        return
    s = df_in[column]
    s_norm = s.astype("string").str.strip().str.lower()
    mask = s.isna() | s.isin(UNKNOWN_LIKE)
    total = len(df_in)
    unknown_combined = int(mask.sum())
    valid = int(total - unknown_combined)
    pct = lambda n: (n / total * 100.0) if total > 0 else 0.0
    st.caption(
        f"**Data quality — {label}:** Unknown + Missing **{unknown_combined}** ({pct(unknown_combined):.1f}%), "
        f"Valid **{valid}** ({pct(valid):.1f}%)."
    )

def add_unknown_checkbox_and_note(
    df_in: pd.DataFrame, column: str, *, label: str | None = None,
    key: str | None = None, note_style: str = "caption",
) -> pd.DataFrame:
    """Optional checkbox to include/exclude Unknown + Missing for a specific section."""
    label = label or column
    if column not in df_in.columns:
        st.warning(f"Column '{column}' not found; skipping.")
        return df_in

    s_norm = df_in[column].astype("string").str.strip().str.lower()
    mask = df_in[column].isna() | s_norm.isin(UNKNOWN_LIKE)

    total = len(df_in)
    unknown_combined = int(mask.sum())
    valid = int(total - unknown_combined)

    include_unknown = st.checkbox(
        f"Include 'Unknown' in {label}",
        value=False,
        key=key or f"include_unknown_{label}",
        help="Uncheck to hide rows where this value is missing or unknown.",
    )
    filtered = df_in.copy() if include_unknown else df_in.loc[~mask].copy()

    pct = (lambda n: (n / total * 100.0) if total > 0 else 0.0)
    note_text = (
        f"**Data quality — {label}:** Unknown + Missing **{unknown_combined}** ({pct(unknown_combined):.1f}%), "
        f"Valid **{valid}** ({pct(valid):.1f}%). The charts below **{'include' if include_unknown else 'exclude'}** Unknown + Missing."
    )
    {"warning": st.warning, "caption": st.caption}.get(note_style, st.info)(note_text)

    return filtered

# Plot helpers
if "plot_counter" not in st.session_state:
    st.session_state["plot_counter"] = 0

def _next_plot_key(prefix: str) -> str:
    st.session_state["plot_counter"] += 1
    return f"{prefix}_{st.session_state['plot_counter']}"

def plotly_show(fig, *, prefix: str, **kwargs):
    st.plotly_chart(fig, use_container_width=True, key=_next_plot_key(prefix), **kwargs)

def safe_plot(check_df: pd.DataFrame, plot_callable):
    if isinstance(check_df, (pd.DataFrame, pd.Series)) and check_df.empty:
        st.warning("No data to display after filtering. Try including 'Unknown'.")
        return
    plot_callable()

def create_html_report(charts: list[tuple[str, "px.Figure"]]) -> bytes:
    """
    Build a single self-contained HTML with all charts (interactive).
    Applies export styling per figure. Uses SGT for timestamp.
    """
    now_sgt = datetime.now(ZoneInfo("Asia/Singapore")).strftime("%Y-%m-%d %H:%M:%S")
    parts = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        "<title>EE Analytics Report</title>",
        ("<style>"
         "body{font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;margin:24px;}"
         "h1{font-size:22px;margin:0 0 8px 0;}"
         "h2{font-size:18px;margin:20px 0 8px 0;}"
         "hr{margin:24px 0;border:0;border-top:1px solid #ddd;}"
         "</style>"),
        "</head><body>",
        f"<h1>EE Analytics Report</h1><p>Generated: {now_sgt}</p><hr/>"
    ]

    for title, fig in charts:
        parts.append(f"<h2>{(title or 'Untitled Chart')}</h2>")
        styled = _style_for_export(fig)
        parts.append(styled.to_html(full_html=False, include_plotlyjs="inline"))
        parts.append("<hr/>")

    parts.append("</body></html>")
    return "\n".join(parts).encode("utf-8")

def _style_for_export(fig, *, kind: str | None = None):
    """
    Safe export styling:
    - plotly_white template, consistent fonts
    - Force BLUE for non-geo/heatmap/pie using update_traces (robust)
    - Fix heatmap/long y-label margins
    - Ensure scatter lines show markers for readability
    """
    import copy
    f = copy.deepcopy(fig)

    data = tuple(getattr(f, "data", ()) or ())
    trace_types = set()
    for t in data:
        ttype = getattr(t, "type", "") or ""
        trace_types.add(ttype)

    is_geo     = any(t in trace_types for t in ("scattergeo", "choropleth"))
    is_heatmap = any(t in trace_types for t in ("heatmap", "imshow"))
    is_pie     = "pie" in trace_types

    f.update_layout(
        template="plotly_white",
        font=dict(family="Inter, Segoe UI, Roboto, Arial, sans-serif", size=12),
        title=dict(font=dict(size=18)),
        legend=dict(font=dict(size=11)),
        margin=dict(l=80, r=20, t=60, b=60),
        xaxis_title_standoff=30,
        yaxis_title_standoff=40,
    )
    f.update_layout(
        hovermode="closest",
        hoverlabel=dict(
            bgcolor="white",
            font_color="black",
            font_size=12,
            bordercolor="black"
        )
    )
    f.update_xaxes(automargin=True, tickangle=0)
    f.update_yaxes(automargin=True)

    BLUE = "#1f77b4"

    if not (is_geo or is_heatmap or is_pie):
        f.update_traces(marker_color=BLUE, selector=dict(type="bar"))
        f.update_traces(marker_color=BLUE, selector=dict(type="histogram"))
        f.update_traces(marker_color=BLUE, selector=dict(type="barpolar"))

        f.update_traces(line=dict(color=BLUE), selector=dict(type="scatter"))
        f.update_traces(marker=dict(color=BLUE, line=dict(color=BLUE)), selector=dict(type="scatter"))
        f.update_traces(line=dict(color=BLUE), selector=dict(type="scattergl"))
        f.update_traces(marker=dict(color=BLUE, line=dict(color=BLUE)), selector=dict(type="scattergl"))

        f.update_traces(marker_color=BLUE, selector=dict(type="box"))
        f.update_traces(marker_color=BLUE, selector=dict(type="violin"))
        f.update_traces(marker_color=BLUE, selector=dict(type="funnel"))
        f.update_traces(marker_color=BLUE, selector=dict(type="waterfall"))

        for t in data:
            if getattr(t, "type", "") in ("scatter", "scattergl"):
                mode = getattr(t, "mode", "lines") or "lines"
                if "lines" in mode and "markers" not in mode:
                    t.mode = mode + "+markers"

    if is_heatmap:
        cur_l = int((f.layout.margin.l or 0))
        f.update_layout(margin=dict(l=max(cur_l, 120)))
        f.update_yaxes(tickangle=0)

    if "bar" in trace_types:
        any_h = any(getattr(t, "orientation", None) == "h" for t in data if getattr(t, "type", "") == "bar")
        if any_h:
            yvals = []
            for t in data:
                if getattr(t, "type", "") == "bar" and getattr(t, "orientation", None) == "h":
                    y = getattr(t, "y", None)
                    if y is not None:
                        yvals = [str(v) for v in list(y)]
                        break
            max_label_len = max((len(s) for s in yvals), default=12)
            left_padding = int(min(max(120, max_label_len * 7), 340))
            f.update_layout(margin=dict(l=left_padding))
            nrows = len(set(yvals)) or 6
            base_h = 28 * nrows + 140
            f.update_layout(height=min(max(base_h, 360), 1200))

    return f

# ──────────────────────────────────────────────────────────────────────────────
# Global date span (from df_full)
# ──────────────────────────────────────────────────────────────────────────────
if "Run_Month" in df_full.columns:
    full_min = pd.to_datetime(df_full["Run_Month"], errors="coerce").min()
    full_max = pd.to_datetime(df_full["Run_Month"], errors="coerce").max()
    if "run_month_full_span" not in st.session_state:
        st.session_state["run_month_full_span"] = (full_min.date(), full_max.date())
    if "run_month_range" not in st.session_state:
        st.session_state["run_month_range"] = st.session_state["run_month_full_span"]

# ──────────────────────────────────────────────────────────────────────────────
# Sidebar: global filters
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Filters")

    def select_all_filters():
        for label in UI_FILTER_LABELS:
            ms_key = _safe_key(label, "multi")
            col = _col_from_label(label)
            series = df.get(col, pd.Series([], dtype="object")).copy()
            series = series.astype("string").fillna("Unknown") if pd.api.types.is_categorical_dtype(series) else series.fillna("Unknown")
            st.session_state[ms_key] = sorted(series.unique().tolist())
        if "run_month_full_span" in st.session_state:
            st.session_state["run_month_range"] = st.session_state["run_month_full_span"]

    def clear_all_filters():
        for label in UI_FILTER_LABELS:
            st.session_state[_safe_key(label, "multi")] = []
        if "run_month_full_span" in st.session_state:
            st.session_state["run_month_range"] = st.session_state["run_month_full_span"]

    c1, c2 = st.columns(2)
    with c1:
        st.button("✅ Select all filters", key="btn_select_all_filters", on_click=select_all_filters, use_container_width=True)
    with c2:
        st.button("🧹 Clear all filters", key="btn_clear_all_filters", on_click=clear_all_filters, use_container_width=True)

    if "Run_Month" in df.columns:
        full_min_date, full_max_date = st.session_state["run_month_full_span"]
        st.date_input(
            "Run month range",
            key="run_month_range",
            value=st.session_state["run_month_range"],
            min_value=full_min_date,
            max_value=full_max_date,
        )
        start_d, end_d = map(pd.to_datetime, st.session_state["run_month_range"])
        df = df[(df["Run_Month"] >= start_d) & (df["Run_Month"] <= end_d)]

    # Per-filter multiselects
    sel_status   = multiselect_with_all_button("Application Status", df)
    sel_app_type = multiselect_with_all_button("Applicant Type", df)
    sel_primcat  = multiselect_with_all_button("Pri Category", df)
    sel_secncat  = multiselect_with_all_button("Sec Category", df)
    sel_country  = multiselect_with_all_button("Country", df)
    sel_senior   = multiselect_with_all_button("Seniority", df)
    sel_domain   = multiselect_with_all_button("Domain", df)
    top_k = st.number_input("Top K (for Top-X charts)", min_value=3, max_value=50, value=10, step=1)

# Apply filters
mask = (
    apply_filter(df.get(_col_from_label("Application Status"), pd.Series(index=df.index)), sel_status) &
    apply_filter(df.get(_col_from_label("Applicant Type"),     pd.Series(index=df.index)), sel_app_type) &
    apply_filter(df.get(_col_from_label("Pri Category"),       pd.Series(index=df.index)), sel_primcat) &
    apply_filter(df.get(_col_from_label("Sec Category"),       pd.Series(index=df.index)), sel_secncat) &
    apply_filter(df.get(_col_from_label("Country"),            pd.Series(index=df.index)), sel_country) &
    apply_filter(df.get(_col_from_label("Seniority"),          pd.Series(index=df.index)), sel_senior) &
    apply_filter(df.get(_col_from_label("Domain"),             pd.Series(index=df.index)), sel_domain)
)
df_f = df[mask].copy()
st.caption(f"Showing **{len(df_f):,}** of **{len(df):,}** rows after filters")

# ──────────────────────────────────────────────────────────────────────────────
# Chart Collection List (for HTML report)
# ──────────────────────────────────────────────────────────────────────────────
charts_for_html: list[tuple[str, "px.Figure"]] = []

# ──────────────────────────────────────────────────────────────────────────────
# Tabs & visuals
# ──────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab_6, tab_7, tab_8, tab_9 = st.tabs([
    "📈 Time Series",
    "🗺️ Geography",
    "🏷️ Programmes × Country",
    "👔 Titles & Orgs",
    "🧮 Age & Demographics",
    "🧭 Category Insights",
    "💰 Programme Cost",
    "🎯 Programme Deep Dive",
    "ℹ️ Data Preview",
])

# Tab 1: Time Series
with tab1:
    st.subheader("Participants over Time")
    if "Run_Month" in df_f.columns:
        ts = df_f.groupby("Run_Month").size().reset_index(name="Participants").sort_values("Run_Month")
        fig = px.line(ts, x="Run_Month", y="Participants", markers=True, title="Participants over Time")
        fig.update_layout(yaxis_title="Participants", xaxis_title="Run Month")

        charts_for_html.append(("Participants over Time", fig))
        plotly_show(fig, prefix="tab1_participants_over_time")

    if "Programme Start Date" in df_f.columns:
        tmp = df_f.copy()
        tmp["Start_Month_Num"] = tmp["Programme Start Date"].dt.month
        tmp["Start_Quarter"]   = tmp["Programme Start Date"].dt.quarter

        col_a, col_b = st.columns(2)
        with col_a:
            mon = tmp.groupby("Start_Month_Num").size().reset_index(name="Applications")
            mon["Month"] = mon["Start_Month_Num"].map({1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"})
            mon = mon.sort_values("Start_Month_Num")
            figm = px.bar(mon, x="Month", y="Applications", title="Applications by Start Month")

            charts_for_html.append(("Applications by Start Month", figm))
            plotly_show(figm, prefix="tab1_by_start_month")

        with col_b:
            q = tmp.groupby("Start_Quarter").size().reset_index(name="Applications")
            figq = px.bar(q, x="Start_Quarter", y="Applications", title="Applications by Start Quarter")

            charts_for_html.append(("Applications by Start Quarter", figq))
            plotly_show(figq, prefix="tab1_by_start_quarter")

# Tab 2: Geography
with tab2:
    st.subheader("Geospatial: Participants by Country")
    country_col = "Country Of Residence"
    if country_col in df_f.columns:
        dq_note_only(df_f, country_col, "Country")

        exclude_sg = st.checkbox("Exclude Singapore (reduce skew)", value=False, key="geo_exclude_sg")

        base = df_f.copy()
        if exclude_sg:
            base = base[base[country_col] != "Singapore"].copy()

        s = base[country_col].astype("string").str.strip()
        s_norm = s.str.lower()
        mask_unknown = s.isna() | (s == "") | s_norm.isin(UNKNOWN_LIKE)
        df_map = base.loc[~mask_unknown].copy()

        if df_map.empty:
            st.info("No data to display for current filters.")
        else:
            geo = df_map.groupby(country_col).size().reset_index(name="Participants")
            size = np.sqrt(geo["Participants"].astype(float).clip(lower=0))
            size = 10 + (size / size.max()) * 30 if size.max() > 0 else np.full_like(size, 10, dtype=float)
            geo["BubbleSize"] = size

            color_values = geo["Participants"].astype(float)
            cmin = float(color_values.min())
            cmax = float(np.quantile(color_values, 0.95))
            if cmax <= cmin:
                cmax = cmin + 1.0

            fig = px.scatter_geo(
                geo,
                locations=country_col,
                locationmode="country names",
                size="BubbleSize",
                color="Participants",
                hover_name=country_col,
                hover_data={"Participants": True, "BubbleSize": False},
                color_continuous_scale="Viridis",
                projection="natural earth",
                title="Geospatial: Participants by Country"
            )
            fig.update_traces(marker=dict(sizemode="area", line=dict(width=0.5, color="rgba(0,0,0,0.25)")))
            fig.update_layout(
                coloraxis_colorbar_title="Participants",
                coloraxis_cmin=cmin,
                coloraxis_cmax=cmax,
                margin=dict(l=0, r=0, t=40, b=0),
            )

            charts_for_html.append(("Geospatial: Participants by Country", fig))
            plotly_show(fig, prefix="tab2_geo_map")

        # Top-K countries bar
        st.markdown("**Pareto of Countries (Top K)**")
        if df_f.empty:
            st.info("No countries to display for the current filters.")
        else:
            pareto_base = df_f.copy()
            if exclude_sg:
                pareto_base = pareto_base[~_norm_str(pareto_base[country_col]).eq("singapore")].copy()

            s_norm = _norm_str(pareto_base[country_col])
            mask_unknown = pareto_base[country_col].isna() | s_norm.isin(UNKNOWN_LIKE)
            s = pareto_base.loc[~mask_unknown, country_col]

            if s.empty:
                st.info("No valid countries to display for the current filters.")
            else:
                counts_df = s.value_counts(dropna=False).reset_index()
                counts_df.columns = [country_col, "Participants"]
                total_universe = int(counts_df["Participants"].sum())

                final_df = counts_df.nlargest(int(top_k), "Participants").copy()
                final_df["Share_%"] = (final_df["Participants"] / total_universe * 100.0) if total_universe > 0 else 0.0

                chart_title = f"Top {int(top_k)} Countries by Participants"
                fig_bar = px.bar(
                    final_df,
                    x=country_col,
                    y="Participants",
                    title=chart_title,
                    text=final_df["Share_%"].round(1).astype(str) + "%",
                )
                fig_bar.update_traces(textposition="outside", cliponaxis=False)
                fig_bar.update_layout(xaxis_tickangle=-45, yaxis_title="Participants", xaxis_title="Country")

                charts_for_html.append((chart_title, fig_bar))
                plotly_show(fig_bar, prefix="tab2_geo_pareto")

                sg_note = " (Singapore excluded)" if exclude_sg else ""
                st.caption(
                    f"Total participants shown: {total_universe:,}{sg_note}. "
                    f"Top {int(top_k)} countries account for {final_df['Share_%'].sum():.1f}% of the shown total."
                )

# Tab 3: Programmes × Country (two heatmaps share the same %/raw toggle)
with tab3:
    st.subheader("Top Programmes & Country Breakdown")

    prog_col    = "Truncated Programme Name"
    country_col = "Country Of Residence"

    if (prog_col in df_f.columns) and (country_col in df_f.columns):
        dq_note_only(df_f, country_col, "Country (Heatmaps)")
        st.caption("Rows are ordered by total participants after filtering (Unknown removed; Singapore excluded if selected).")

        exclude_sg_tab3 = st.checkbox("Exclude Singapore (reduce skew)", value=False, key="pc_exclude_sg")
        show_pct_tab3 = st.toggle("Show % (vs raw counts)", value=True, key="tab3_show_pct")

        base = df_f.copy()
        if exclude_sg_tab3:
            norm_cty = base[country_col].astype("string").str.strip().str.casefold()
            base = base.loc[~norm_cty.eq("singapore")].copy()

        # Heatmap 1: Programme × Country
        st.markdown("### Heatmap 1 — Programme × Country")
        s = base[country_col].astype("string").str.strip()
        s_norm = s.str.lower()
        mask_unknown = s.isna() | (s == "") | s_norm.isin(UNKNOWN_LIKE)
        base_hm1 = base.loc[~mask_unknown].copy()

        if base_hm1.empty:
            st.info("No valid Country data for Heatmap 1 after filtering.")
        else:
            top_progs = base_hm1[prog_col].value_counts().nlargest(int(top_k)).index.tolist()
            agg = (base_hm1[base_hm1[prog_col].isin(top_progs)]
                   .groupby([prog_col, country_col]).size().reset_index(name="Participants"))
            top_countries = agg.groupby(country_col)["Participants"].sum().nlargest(int(top_k)).index.tolist()
            agg = agg[agg[country_col].isin(top_countries)]

            hm1_counts = agg.pivot(index=prog_col, columns=country_col, values="Participants").fillna(0)
            hm1_counts = hm1_counts.loc[hm1_counts.sum(axis=1).sort_values(ascending=False).index]

            if show_pct_tab3:
                total_sum = hm1_counts.values.sum()
                hm1_pct = (hm1_counts / total_sum * 100).round(2) if total_sum > 0 else hm1_counts * 0
                Z = hm1_pct.values
                x_labels, y_labels = hm1_pct.columns, hm1_pct.index
                color_label = "Share of total (%)"
                title_suffix = "(% of total)"
                text_fmt = ".2f"
                hover_tmpl = "Programme: %{y}<br>Country: %{x}<br>Share of total: %{z:.2f}%<extra></extra>"
            else:
                Z = hm1_counts.values
                x_labels, y_labels = hm1_counts.columns, hm1_counts.index
                color_label = "Participants"
                title_suffix = "(raw)"
                text_fmt = "d"
                hover_tmpl = "Programme: %{y}<br>Country: %{x}<br>Participants: %{z:.0f}<extra></extra>"

            chart_title = f"Programme × Country {title_suffix}"
            fig_hm1 = px.imshow(
                Z, x=x_labels, y=y_labels, color_continuous_scale="Viridis", aspect="auto",
                labels=dict(x="Country Of Residence", y="Programme", color=color_label),
                text_auto=text_fmt, title=chart_title,
            )
            fig_hm1.update_traces(hovertemplate=hover_tmpl)
            fig_hm1.update_layout(xaxis_title="Country Of Residence", yaxis_title="Programme (Anon)")

            charts_for_html.append((chart_title, fig_hm1))
            plotly_show(fig_hm1, prefix="tab3_prog_country_heatmap")

        # Heatmap 2: Top Countries × Primary Category
        st.markdown(f"### Heatmap 2 — Top {top_k} Countries × Primary Category")
        cat_col = "Primary Category"
        if cat_col not in df_f.columns:
            st.info("‘Primary Category’ column not found for Heatmap 2.")
        else:
            s2 = base[country_col].astype("string").str.strip()
            s2_norm = s2.str.lower()
            mask_unknown2 = s2.isna() | (s2 == "") | s2_norm.isin(UNKNOWN_LIKE)
            base_hm2 = base.loc[~mask_unknown2].copy()

            if base_hm2.empty:
                st.info("No valid Country data for Heatmap 2 after filtering.")
            else:
                top_countries_hm2 = base_hm2[country_col].value_counts().nlargest(int(top_k)).index.tolist()
                df_top_cty = base_hm2[base_hm2[country_col].isin(top_countries_hm2)].copy()

                agg_cat = df_top_cty.groupby([country_col, cat_col]).size().reset_index(name="Participants")
                hm2_counts = agg_cat.pivot(index=country_col, columns=cat_col, values="Participants").fillna(0)
                hm2_counts = hm2_counts.loc[hm2_counts.sum(axis=1).sort_values(ascending=False).index]

                if show_pct_tab3:
                    row_sums = hm2_counts.sum(axis=1).replace(0, np.nan)
                    hm2_pct = (hm2_counts.div(row_sums, axis=0) * 100).round(2).fillna(0)
                    Z2 = hm2_pct.values
                    x2, y2 = hm2_pct.columns, hm2_pct.index
                    color_label2 = "Row %"
                    title_suffix2 = "(row %)"
                    text_fmt2 = ".2f"
                    hover_tmpl2 = "Country: %{y}<br>Primary Category: %{x}<br>Row share: %{z:.2f}%<extra></extra>"
                else:
                    Z2 = hm2_counts.values
                    x2, y2 = hm2_counts.columns, hm2_counts.index
                    color_label2 = "Participants"
                    title_suffix2 = "(raw)"
                    text_fmt2 = "d"
                    hover_tmpl2 = "Country: %{y}<br>Primary Category: %{x}<br>Participants: %{z:.0f}<extra></extra>"

                chart_title = f"For each country: distribution {title_suffix2}"
                fig_cat = px.imshow(
                    Z2, x=x2, y=y2, color_continuous_scale="Viridis", aspect="auto",
                    labels=dict(x="Primary Category", y="Country Of Residence", color=color_label2),
                    text_auto=text_fmt2, title=chart_title,
                )
                fig_cat.update_traces(hovertemplate=hover_tmpl2)
                fig_cat.update_layout(xaxis_title="Primary Category", yaxis_title="Country Of Residence")

                charts_for_html.append((chart_title, fig_cat))
                plotly_show(fig_cat, prefix="tab3_country_primarycat_heatmap")
    else:
        st.info("Required columns not found: please ensure ‘Truncated Programme Name’ and ‘Country Of Residence’ are present.")

# Tab 4: Titles & Organisations
with tab4:
    st.subheader("Top Job Titles & Organisations")

    if "Job Title Clean" in df_f.columns:
        df_title = add_unknown_checkbox_and_note(df_f, "Job Title Clean", label="Job Title", key="job_title_tab4", note_style="caption")
        s_titles = df_title["Job Title Clean"]
        if "Unknown" in df_title["Job Title Clean"].astype("string").str.lower().unique() or df_title["Job Title Clean"].isna().any():
            s_titles = coalesce_unknown(s_titles)

        top_titles = s_titles.value_counts(dropna=False).head(top_k).reset_index()
        top_titles.columns = ["Job Title", "Participants"]

        def plot_top_titles():
            chart_title = f"Top {top_k} Job Titles"
            fig = px.bar(top_titles, x="Participants", y="Job Title", orientation="h", title=chart_title)
            charts_for_html.append((chart_title, fig))
            plotly_show(fig, prefix="tab4_top_titles")
        safe_plot(top_titles, plot_top_titles)

    org_col = "Organisation Name"
    if org_col in df_f.columns:
        df_org = add_unknown_checkbox_and_note(df_f, org_col, label="Organisation", key="orgs", note_style="caption")
        top_orgs = df_org[org_col].value_counts().nlargest(top_k).reset_index()
        top_orgs.columns = ["Organisation", "Participants"]

        def plot_top_orgs():
            chart_title = f"Top {top_k} Organisations"
            fig = px.bar(top_orgs, x="Participants", y="Organisation", orientation="h", title=chart_title)
            charts_for_html.append((chart_title, fig))
            plotly_show(fig, prefix="tab4_top_orgs")
        safe_plot(top_orgs, plot_top_orgs)

    if "Domain" in df_f.columns:
        df_dom = add_unknown_checkbox_and_note(df_f, "Domain", label="Domain", key="domain_tab4", note_style="caption")
        include_unknown_domain = bool(st.session_state.get("domain_tab4", False))

        if not include_unknown_domain:
            mask_others = (df_dom["Domain"].astype("string").str.strip().str.lower() == "others")
            df_dom = df_dom.loc[~mask_others].copy()

        top_domains = df_dom["Domain"].value_counts(dropna=False).nlargest(int(top_k)).reset_index()
        top_domains.columns = ["Domain", "Participants"]

        st.caption(
            "Note: **'Others'** is a valid cluster label from topic modeling. It's hidden by default for clarity "
            "but still counted as valid data in the note above."
        )

        def plot_top_domains():
            chart_title = f"Top {int(top_k)} Domains"
            fig = px.bar(top_domains, x="Participants", y="Domain", orientation="h", title=chart_title)
            charts_for_html.append((chart_title, fig))
            plotly_show(fig, prefix="tab4_top_domains")
        safe_plot(top_domains, plot_top_domains)

    if "Seniority" in df_f.columns:
        df_sen = add_unknown_checkbox_and_note(df_f, "Seniority", key="seniority", note_style="caption")
        sen = df_sen["Seniority"].value_counts().reset_index()
        sen.columns = ["Seniority", "Participants"]

        def plot_seniority():
            chart_title = "Participants by Seniority"
            fig = px.bar(sen, x="Seniority", y="Participants", title=chart_title)
            charts_for_html.append((chart_title, fig))
            plotly_show(fig, prefix="tab4_seniority")
        safe_plot(sen, plot_seniority)

# Tab 5: Age & Demographics
with tab5:
    st.subheader("Demographics")

    if "Age_Group" in df_f.columns:
        df_age = add_unknown_checkbox_and_note(df_f, "Age_Group", label="Age Group", key="agegroup_tab5", note_style="caption")
        include_unknown_age = bool(st.session_state.get("agegroup_tab5", False))
        s_raw = df_age["Age_Group"].astype("string")

        if include_unknown_age:
            s = coalesce_unknown(s_raw)
            order = ["<35", "35–44", "45–54", "55–64", "65+", "Unknown"]
        else:
            s = s_raw[s_raw.ne("Unknown")]
            order = ["<35", "35–44", "45–54", "55–64", "65+"]

        agec = (s.value_counts(dropna=False).reindex(order).fillna(0).rename_axis("Age Group").reset_index(name="Participants"))

        def plot_age_group():
            chart_title = "Participants by Age Group"
            fig = px.bar(agec, x="Age Group", y="Participants", title=chart_title)
            charts_for_html.append((chart_title, fig))
            plotly_show(fig, prefix="tab5_agegroup_bar")
        safe_plot(agec, plot_age_group)

    if "Gender" in df_f.columns:
        df_gender = add_unknown_checkbox_and_note(df_f, "Gender", key="gender_tab5", note_style="caption")
        gender = df_gender["Gender"].value_counts().reset_index()
        gender.columns = ["Gender", "Participants"]

        def plot_gender():
            chart_title = "Gender Split"
            fig = px.pie(gender, names="Gender", values="Participants", title=chart_title)
            charts_for_html.append((chart_title, fig))
            plotly_show(fig, prefix="tab5_gender_pie")
        safe_plot(gender, plot_gender)

# Tab 6: Category Insights
with tab_6:
    st.subheader("Category Insights")
    sub_age, sub_country = st.tabs(["📊 Age Distribution per Category", "🌍 Country Distribution per Category"])

    with sub_age:
        st.markdown("##### Age Distribution per Category")
        cat_type = st.radio("Choose category type:", ["Primary Category", "Secondary Category"], key="age_cat_type", horizontal=True)
        cat_col = cat_type

        if (cat_col in df_f.columns) and ("Age_Group" in df_f.columns):
            cat_values = df_f[cat_col].astype("string").fillna("Unknown").replace({"": "Unknown"}).unique().tolist()
            cat_values = [v for v in sorted(cat_values) if v != "Unknown"] + (["Unknown"] if "Unknown" in cat_values else [])
            selected_cat = st.selectbox(f"Select {cat_type}:", cat_values, key="age_cat_select")

            subset = df_f[df_f[cat_col].astype("string").fillna("Unknown") == selected_cat].copy()
            if subset.empty:
                st.info("No rows for this selection.")
            else:
                sub_age_df = add_unknown_checkbox_and_note(subset, "Age_Group", label="Age Group (this selection)", key="age_dist_sub", note_style="caption")
                s = coalesce_unknown(sub_age_df["Age_Group"])
                dist = (s.value_counts(normalize=True, dropna=False) * 100.0).reset_index()
                dist.columns = ["Age Group", "Percentage"]

                order_full = ["<35", "35–44", "45–54", "55–64", "65+", "Unknown"]
                present = [g for g in order_full if g in dist["Age Group"].values]
                dist = dist.set_index("Age Group").reindex(present).reset_index()

                if dist.empty:
                    st.info("No Age Group data to display for this selection.")
                else:
                    text_labels = dist["Percentage"].round(1).astype(str) + "%"
                    chart_title = f"Age Distribution (%) – {cat_type}: {selected_cat}"
                    fig = px.bar(dist, x="Age Group", y="Percentage", title=chart_title, text=text_labels)
                    fig.update_traces(textposition="outside", cliponaxis=False)
                    ymax = min(100.0, float(dist["Percentage"].max()) + 10.0)
                    fig.update_layout(yaxis_range=[0, ymax])

                    charts_for_html.append((chart_title, fig))
                    plotly_show(fig, prefix="tab6_age_dist_by_cat")
        else:
            st.info("Required columns not found: please include ‘Age_Group’ and the selected category column.")

    with sub_country:
        st.markdown("##### Country Distribution per Category")
        cat_type = st.radio("Choose category type:", ["Primary Category", "Secondary Category"], key="country_cat_type", horizontal=True)
        cat_col = cat_type
        country_col = "Country Of Residence"
        exclude_sg_tab6 = st.checkbox("Exclude Singapore (reduce skew)", value=False, key="tab6_exclude_sg")

        if (cat_col in df_f.columns) and (country_col in df_f.columns):
            cat_values = df_f[cat_col].astype("string").fillna("Unknown").replace({"": "Unknown"}).unique().tolist()
            cat_values = [v for v in sorted(cat_values) if v != "Unknown"] + (["Unknown"] if "Unknown" in cat_values else [])
            selected_cat = st.selectbox(f"Select {cat_type}:", cat_values, key="country_cat_select")

            subset = df_f[df_f[cat_col].astype("string").fillna("Unknown") == selected_cat].copy()
            if subset.empty:
                st.info("No rows for this selection.")
            else:
                dq_note_only(subset, country_col, "Country (this selection)")

                if exclude_sg_tab6:
                    subset = subset[~_norm_str(subset[country_col]).eq("singapore")].copy()

                s_norm = _norm_str(subset[country_col])
                mask_unknown = subset[country_col].isna() | s_norm.isin(UNKNOWN_LIKE)
                sub_valid = subset.loc[~mask_unknown].copy()

                if sub_valid.empty:
                    st.info("No valid countries to display after removing Unknown (and Singapore, if excluded).")
                else:
                    counts_df = sub_valid[country_col].value_counts(dropna=False).reset_index()
                    counts_df.columns = [country_col, "Participants"]
                    total_universe = int(counts_df["Participants"].sum())
                    final_df = counts_df.nlargest(int(top_k), "Participants").copy()
                    final_df["Share_%"] = (final_df["Participants"] / total_universe * 100.0) if total_universe > 0 else 0.0

                    chart_title = f"Top {int(top_k)} Countries by Participants — {cat_type}: {selected_cat}"
                    fig = px.bar(final_df, x=country_col, y="Participants", title=chart_title, text=final_df["Share_%"].round(1).astype(str) + "%")
                    fig.update_traces(textposition="outside", cliponaxis=False)
                    fig.update_layout(xaxis_tickangle=-45, yaxis_title="Participants", xaxis_title="Country")

                    charts_for_html.append((chart_title, fig))
                    plotly_show(fig, prefix="tab6_country_dist_by_cat_like_tab2")

                    sg_note = " (Singapore excluded)" if exclude_sg_tab6 else ""
                    st.caption(
                        f"Total participants shown: {total_universe:,}{sg_note}. "
                        f"Top {int(top_k)} countries account for {final_df['Share_%'].sum():.1f}% of the shown total."
                    )
        else:
            st.info("Required columns not found: please include ‘Country Of Residence’ and the chosen category column.")

# Tab 7: Programme Cost
with tab_7:
    st.subheader("Programme Cost Analysis")
    required_cols = ["Programme Cost", "Truncated Programme Name", "Run_Month"]
    if not all(col in df_f.columns for col in required_cols):
        st.warning("Missing columns: ‘Programme Cost’, ‘Truncated Programme Name’, or ‘Run_Month’.")
    else:
        df_cost = df_f.copy()
        df_cost['Programme Cost'] = pd.to_numeric(df_cost['Programme Cost'], errors='coerce')
        df_cost.dropna(subset=['Programme Cost'], inplace=True)

        if df_cost.empty:
            st.info("No rows with valid programme costs for the current filters.")
        else:
            st.markdown("##### Enrolment Volume vs. Programme Cost")
            grouped = df_cost.groupby('Truncated Programme Name').agg(
                enrolment_volume=('Truncated Programme Name', 'size'),
                programme_cost=('Programme Cost', 'first')
            ).reset_index()

            chart_title = "Enrolment Volume vs. Programme Cost"
            fig_scatter = px.scatter(
                grouped,
                x='programme_cost',
                y='enrolment_volume',
                title=chart_title,
                labels={'programme_cost': 'Programme Cost ($)', 'enrolment_volume': 'Total Enrolments'},
                hover_data=['Truncated Programme Name']
            )
            fig_scatter.update_traces(marker=dict(size=12, opacity=0.7, line=dict(width=1, color='DarkSlateGrey')))

            charts_for_html.append((chart_title, fig_scatter))
            plotly_show(fig_scatter, prefix="tab7_cost_vs_enrolment")

            st.divider()
            st.markdown("##### Monthly Revenue Trend")
            monthly_revenue = df_cost.groupby(df_cost['Run_Month'].dt.to_period('M'))['Programme Cost'].sum().reset_index()
            monthly_revenue.rename(columns={'Programme Cost': 'Total_Revenue'}, inplace=True)
            monthly_revenue['Run_Month'] = monthly_revenue['Run_Month'].dt.to_timestamp()
            monthly_revenue = monthly_revenue.sort_values("Run_Month")

            chart_title = "Monthly Revenue Trend"
            fig_trend = px.line(monthly_revenue, x='Run_Month', y='Total_Revenue', title=chart_title, labels={'Run_Month': 'Month', 'Total_Revenue': 'Total Revenue ($)'}, markers=True)
            fig_trend.update_layout(yaxis_title="Total Revenue ($)", xaxis_title="Month")

            charts_for_html.append((chart_title, fig_trend))
            plotly_show(fig_trend, prefix="tab7_monthly_revenue")

            st.divider()
            st.markdown("##### Top K Countries by Total Revenue")
            if "Country Of Residence" in df_cost.columns:
                top_countries = df_cost.groupby("Country Of Residence")["Programme Cost"].sum().nlargest(top_k).reset_index()
                top_countries.columns = ["Country Of Residence", "Total Revenue"]

                chart_title = f"Top {top_k} Countries by Total Revenue"
                fig2 = px.bar(top_countries, x="Total Revenue", y="Country Of Residence", orientation="h", title=chart_title)

                charts_for_html.append((chart_title, fig2))
                st.plotly_chart(fig2, use_container_width=True)

# Tab 8: Programme Deep Dive
with tab_8:
    st.subheader("Programme Deep Dive")
    include_unknown_deep = st.checkbox("Include 'Unknown' in Deep Dive visuals", value=False, key="dd_include_unknown")

    prog_col = "Truncated Programme Name"
    if prog_col not in df_f.columns:
        st.info("Programme column not found.")
    else:
        progs = (df_f[prog_col].dropna().astype(str).sort_values().unique().tolist()) if not df_f.empty else []
        if not progs:
            st.info("No programmes available under current filters.")
        else:
            sel_prog = st.selectbox("Select a programme", progs, index=0, key="prog_dd_select")
            p = df_f[df_f[prog_col] == sel_prog].copy()
            if p.empty:
                st.info("No rows for this programme with current filters.")
            else:
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    st.metric("Participants", f"{len(p):,}")
                with c2:
                    st.metric("Unique Runs", int(p.get("Truncated Programme Run", pd.Series()).nunique() if "Truncated Programme Run" in p.columns else 0))
                with c3:
                    st.metric("Countries", int(p.get("Country Of Residence", pd.Series()).nunique() if "Country Of Residence" in p.columns else 0))
                with c4:
                    st.metric("Median Age", f"{p['Age'].median():.0f}" if "Age" in p.columns and p["Age"].notna().any() else "—")
                with c5:
                    st.metric("Cost", p["Programme Cost"].unique() if p["Programme Cost"].notna().any() else "Unknown")

                date_min = pd.to_datetime(p.get("Programme Start Date")).min() if "Programme Start Date" in p.columns else pd.NaT
                date_max = pd.to_datetime(p.get("Programme End Date")).max() if "Programme End Date" in p.columns else pd.NaT
                if pd.notna(date_min) and pd.notna(date_max):
                    st.caption(f"**Programme Date Range:** {date_min:%A, %d %B %Y} → {date_max:%A, %d %B %Y}")
                    st.divider()

                if "Run_Month" in p.columns:
                    ts = p.groupby("Run_Month").size().reset_index(name="Participants").sort_values("Run_Month")
                    chart_title = "Participants over Time (by Run Month)"
                    fig_ts = px.line(ts, x="Run_Month", y="Participants", markers=True, title=chart_title)
                    fig_ts.update_layout(yaxis_title="Participants", xaxis_title="Run Month")
                    fig_ts.update_xaxes(tickformat="%b %Y")
                    fig_ts.update_traces(hovertemplate="Run Month=%{x|%b %Y}<br>Participants=%{y}<extra></extra>")

                    charts_for_html.append((f"{sel_prog}: {chart_title}", fig_ts))
                    plotly_show(fig_ts, prefix="tab8_prog_ts")

                colL, colR = st.columns(2)
                with colL:
                    if "Application Status" in p.columns:
                        p_status = filter_unknown_no_ui(p, "Application Status", include_unknown_deep)
                        dq_caption(p, "Application Status", "Application Status")
                        status = p_status["Application Status"].value_counts().reset_index()
                        status.columns = ["Application Status", "Count"]

                        def plot_app_status():
                            chart_title = "Application Status Breakdown"
                            fig = px.bar(status, x="Application Status", y="Count", title=chart_title, text="Count").update_traces(textposition="outside", cliponaxis=False).update_layout(xaxis_tickangle=-30)
                            charts_for_html.append((f"{sel_prog}: {chart_title}", fig))
                            plotly_show(fig, prefix="tab8_status_bar")
                        safe_plot(status, plot_app_status)

                    if "Gender" in p.columns:
                        p_gender = filter_unknown_no_ui(p, "Gender", include_unknown_deep)
                        dq_caption(p, "Gender", "Gender")
                        gender = p_gender["Gender"].value_counts().reset_index()
                        gender.columns = ["Gender", "Participants"]

                        def plot_gender_split():
                            chart_title = "Gender Split"
                            fig = px.pie(gender, names="Gender", values="Participants", title=chart_title)
                            charts_for_html.append((f"{sel_prog}: {chart_title}", fig))
                            plotly_show(fig, prefix="tab8_gender_pie")
                        safe_plot(gender, plot_gender_split)

                with colR:
                    if "Country Of Residence" in p.columns:
                        p_cty = filter_unknown_no_ui(p, "Country Of Residence", include_unknown_deep)
                        dq_caption(p, "Country Of Residence", "Country")
                        s = p_cty["Country Of Residence"]
                        if include_unknown_deep:
                            s = coalesce_unknown(s)

                        if s.empty:
                            st.info("No country data to display for this selection.")
                        else:
                            ctry_counts = s.value_counts(dropna=False).reset_index()
                            ctry_counts.columns = ["Country", "Participants"]
                            unknown_row = ctry_counts[ctry_counts["Country"] == "Unknown"]
                            non_unknown = ctry_counts[ctry_counts["Country"] != "Unknown"]
                            top_non_unknown = non_unknown.nlargest(top_k, "Participants")
                            ctry_top = pd.concat([top_non_unknown, unknown_row], ignore_index=True).drop_duplicates(subset=["Country"])

                            order = ctry_top.sort_values("Participants", ascending=False)["Country"].tolist()
                            if "Unknown" in order:
                                order = [c for c in order if c != "Unknown"] + ["Unknown"]
                                ctry_top["Country"] = pd.Categorical(ctry_top["Country"], categories=order, ordered=True)
                                ctry_top = ctry_top.sort_values("Country")

                            def plot_top_countries():
                                chart_title = f"Top {top_k} Countries (Participants)"
                                fig = px.bar(ctry_top, x="Participants", y="Country", orientation="h", title=chart_title)
                                charts_for_html.append((f"{sel_prog}: {chart_title}", fig))
                                plotly_show(fig, prefix="tab8_top_countries")
                            safe_plot(ctry_top, plot_top_countries)

                    org_col = "Organisation Name"
                    if org_col in p.columns:
                        p_org = filter_unknown_no_ui(p, org_col, include_unknown_deep)
                        dq_caption(p, org_col, "Organisation")
                        orgs = p_org[org_col].value_counts().reset_index()
                        orgs.columns = ["Organisation", "Participants"]
                        orgs_top = orgs.head(top_k)

                        def plot_prog_top_orgs():
                            chart_title = f"Top {top_k} Organisations (Participants)"
                            fig = px.bar(orgs_top, x="Participants", y="Organisation", orientation="h", title=chart_title)
                            charts_for_html.append((f"{sel_prog}: {chart_title}", fig))
                            plotly_show(fig, prefix="tab8_top_orgs")
                        safe_plot(orgs_top, plot_prog_top_orgs)

                colA, colB = st.columns(2)
                with colA:
                    if "Seniority" in p.columns:
                        p_sen = filter_unknown_no_ui(p, "Seniority", include_unknown_deep)
                        dq_caption(p, "Seniority", "Seniority")
                        sen = p_sen["Seniority"].value_counts().reset_index()
                        sen.columns = ["Seniority", "Participants"]

                        def plot_prog_seniority():
                            chart_title = "Seniority Mix"
                            fig = px.bar(sen, x="Seniority", y="Participants", title=chart_title)
                            charts_for_html.append((f"{sel_prog}: {chart_title}", fig))
                            plotly_show(fig, prefix="tab8_seniority")
                        safe_plot(sen, plot_prog_seniority)

                with colB:
                    if "Age_Group" in p.columns:
                        p_age = filter_unknown_no_ui(p, "Age_Group", include_unknown_deep)
                        dq_caption(p, "Age_Group", "Age Group")
                        s = p_age["Age_Group"].astype("string")
                        if include_unknown_deep:
                            s = coalesce_unknown(s)
                        age_counts = s.value_counts(dropna=False)

                        order_full = ["<35","35–44","45–54","55–64","65+","Unknown"]
                        present = [g for g in order_full if g in age_counts.index]

                        if not present:
                            st.info("No Age Group data to display for this programme.")
                        else:
                            age_df = (age_counts.reindex(present).fillna(0).rename_axis("Age Group").reset_index(name="Participants"))
                            chart_title = "Age Group Distribution"
                            fig = px.bar(age_df, x="Age Group", y="Participants", title=chart_title, text="Participants")
                            fig.update_traces(textposition="outside", cliponaxis=False)

                            charts_for_html.append((f"{sel_prog}: {chart_title}", fig))
                            plotly_show(fig, prefix="tab8_agegroup")

                st.markdown("##### Runs for this Programme")
                run_cols = [c for c in ["Truncated Programme Run", "Programme Start Date", "Programme End Date", "Country Of Residence", "Application Status"] if c in p.columns]
                st.dataframe(p.sort_values(["Run_Month","Programme Start Date"]).loc[:, run_cols].head(500), use_container_width=True, hide_index=True)

                st.download_button(
                    f"Download '{sel_prog}' rows (CSV)",
                    data=p.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"{sel_prog[:40].replace(' ','_')}_export.csv",
                    mime="text/csv"
                )

# Tab 9: Data Preview
with tab_9:
    st.subheader("Filtered Data Preview")
    preview_cols = [c for c in [
        "Application ID", "Contact ID", "Application Status", "Applicant Type",
        "Organisation Name", "Job Title Clean", "Seniority",
        "Truncated Programme Name", "Truncated Programme Run", "Primary Category", "Secondary Category",
        "Programme Start Date", "Programme End Date", "Run_Month", "Run_Month_Label",
        "Gender", "Age", "Country Of Residence", "Domain", "Programme Cost"
    ] if c in df_f.columns]

    st.dataframe(df_f.sort_values("Run_Month").loc[:, preview_cols].head(500), use_container_width=True, hide_index=True)
    st.download_button("Download filtered CSV", data=df_f.to_csv(index=False).encode("utf-8-sig"), file_name="filtered_export.csv", mime="text/csv")

# ──────────────────────────────────────────────────────────────────────────────
#  HTML Download Section (only)
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.header("📥 Download Report")

    if "html_bytes" not in st.session_state:
        st.session_state.html_bytes = None

    if st.button("Generate HTML Report", key="btn_generate_html", use_container_width=True):
        st.session_state.html_bytes = create_html_report(charts_for_html)

    if st.session_state.html_bytes:
        st.download_button(
            "Download HTML Report (Interactive)",
            data=st.session_state.html_bytes,
            file_name="Dashboard_Report.html",
            mime="text/html",
            use_container_width=True,
        )
