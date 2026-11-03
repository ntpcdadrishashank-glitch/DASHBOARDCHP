"""
NTPC Dadri — CHP Coal Feeding & BOXN/BOBR Rake Demurrage Dashboard
===================================================================
Single-file Streamlit app that mirrors the "Dashboard" sheet of
CHP DASHBOARDS.xlsx (Summary + RAKESTATUS sheets).

Install once:
    pip install streamlit pandas numpy plotly openpyxl requests

Run:
    streamlit run app.py

Data source
-----------
* On every fresh open of the app (new browser session) the workbook is
  downloaded automatically from the OneDrive / SharePoint share link below.
* If the download fails (link not anonymous, no internet, etc.) the app falls
  back to a local copy named CHP_DASHBOARDS.xlsx next to this file.
* The sidebar also has "Fetch latest from link" (manual refresh) and an
  "Upload workbook" option.

Theme: light. (Also put the included .streamlit/config.toml next to app.py so
Streamlit itself starts in light mode.)
"""

from __future__ import annotations

import base64
import io
import re
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

APP_DIR = Path(__file__).parent

# OneDrive / SharePoint share link of the live workbook
SHAREPOINT_SHARE_URL = (
    "https://ntpccoin-my.sharepoint.com/:x:/g/personal/ssharit_ntpc_co_in/"
    "IQBmPz9JVSoLQ6wFwopTIyhZAc2hbaAK3urloA4IMV9RYi4?e=bUYHvY"
)

# Local fallback copies (first one found is used if the online fetch fails)
DEFAULT_FILE_NAMES = [
    "CHP_DASHBOARDS.xlsx",
    "CHP DASHBOARDS.xlsx",
    "CHP_DASHBOARD.xlsx",
]

# Demurrage-amount tariff slabs (Rs per wagon per hour) — mirrors RAKESTATUS!AQ
RATE_SLAB_1 = 150.0
RATE_SLAB_2 = 165.0
RATE_SLAB_3 = 150.0 * 1.25

# ---------------------------------------------------------------------------
# Light theme
# ---------------------------------------------------------------------------

try:
    from streamlit import config as _st_config

    _st_config.set_option("theme.base", "light")
    _st_config.set_option("theme.backgroundColor", "#FFFFFF")
    _st_config.set_option("theme.secondaryBackgroundColor", "#F3F6FA")
    _st_config.set_option("theme.textColor", "#1F2933")
    _st_config.set_option("theme.primaryColor", "#1F6FEB")
except Exception:  # noqa: BLE001
    pass

pio.templates.default = "plotly_white"


def _light(fig):
    """Apply a clean, white-background style to a Plotly figure."""
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font_color="#1F2933",
    )
    return fig


st.set_page_config(
    page_title="CHP Dashboard | Coal Feeding & Rake Demurrage",
    page_icon="🚂",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    [data-testid="stMetric"] {
        background: #F6F8FA;
        border: 1px solid #D0D7DE;
        border-radius: 10px;
        padding: 12px 14px 6px 14px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    [data-testid="stMetricLabel"] { font-size: 0.74rem; color: #57606A; }
    [data-testid="stMetricValue"] { font-size: 1.30rem; color: #1F2933; }
    .section-title {
        font-size: 1.05rem; font-weight: 700; margin-top: 0.6rem; margin-bottom: 0.3rem;
        color: #1F2933;
    }
    .subsection-title {
        font-size: 0.95rem; font-weight: 700; margin-top: 0.9rem; margin-bottom: 0.2rem;
        color: #3B4652;
    }
    hr, [data-testid="stDivider"] { border-color: #D0D7DE !important; }
    @media (max-width: 640px) {
        [data-testid="stMetricValue"] { font-size: 1.05rem; }
        [data-testid="stMetricLabel"] { font-size: 0.68rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _col(df: pd.DataFrame, wanted: str) -> str:
    lookup = {_norm(c): c for c in df.columns}
    key = _norm(wanted)
    if key in lookup:
        return lookup[key]
    for k, v in lookup.items():
        if key in k:
            return v
    raise KeyError(f"Column matching '{wanted}' not found. Available: {list(df.columns)}")


def _hhmm_to_hours(val) -> float:
    if val is None:
        return np.nan
    if isinstance(val, timedelta):
        return val.total_seconds() / 3600.0
    if hasattr(val, "hour") and hasattr(val, "minute"):
        return val.hour + val.minute / 60.0
    if isinstance(val, str):
        parts = val.strip().split(":")
        if len(parts) >= 2:
            try:
                return int(parts[0]) + int(parts[1]) / 60.0
            except ValueError:
                return np.nan
    if isinstance(val, (int, float)) and not pd.isna(val):
        return float(val) * 24.0
    return np.nan


def _demurrage_amount(hrs: float, wagons: float) -> float:
    if pd.isna(hrs) or pd.isna(wagons):
        return np.nan
    if hrs <= 6:
        return hrs * RATE_SLAB_1 * wagons
    if hrs <= 12:
        return 6 * RATE_SLAB_1 * wagons + (hrs - 6) * RATE_SLAB_2 * wagons
    if hrs <= 24:
        return 6 * RATE_SLAB_1 * wagons + 6 * RATE_SLAB_2 * wagons + (hrs - 12) * RATE_SLAB_3 * wagons
    return np.nan


def _fmt_hm(hours: float) -> str:
    if pd.isna(hours):
        return "—"
    total_min = round(hours * 60)
    h, m = divmod(total_min, 60)
    return f"{h}:{m:02d}"


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Reading CHP workbook...")
def parse_workbook(file_bytes: bytes):
    bio = io.BytesIO(file_bytes)

    summary = pd.read_excel(bio, sheet_name="Summary", header=1, engine="openpyxl")
    date_col = _col(summary, "date")
    summary = summary.rename(columns={date_col: "date"})
    summary["date"] = pd.to_datetime(summary["date"], errors="coerce")
    summary = summary.dropna(subset=["date"]).reset_index(drop=True)
    for c in [c for c in summary.columns if c != "date"]:
        summary[c] = pd.to_numeric(summary[c], errors="coerce")

    bio.seek(0)
    rake = pd.read_excel(bio, sheet_name="RAKESTATUS", header=0, engine="openpyxl")
    rake = rake.rename(
        columns={
            _col(rake, "raketype"): "rake_type",
            _col(rake, "source"): "colliery_raw",
            _col(rake, "noofwagons"): "wagons",
            _col(rake, "rakeplacerly_date"): "placed_at",
            _col(rake, "rake_cycletime"): "cycle_raw",
            _col(rake, "total_unloadingtime"): "unload_raw",
            _col(rake, "coal_weight_mines"): "coal_mines_mt",
            _col(rake, "rakeroundno"): "round_no",
            _col(rake, "debit_hours"): "debit_min",
            _col(rake, "siding"): "siding",
            _col(rake, "mgr_remarks"): "mgr_remarks",   # RAKESTATUS column AC
            _col(rake, "chp_remark"): "chp_remarks",    # RAKESTATUS column AD
        }
    )
    for _c in ("mgr_remarks", "chp_remarks"):
        rake[_c] = rake[_c].fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})
    rake["date"] = pd.to_datetime(rake["placed_at"].astype(str).str.slice(0, 10), errors="coerce")
    rake["rake_type"] = rake["rake_type"].astype(str).str.strip()
    rake["colliery_raw"] = rake["colliery_raw"].astype(str).str.strip()
    rake["colliery"] = np.where(rake["colliery_raw"] == "PBCMP", "NTPC", rake["colliery_raw"])
    rake["siding"] = rake["siding"].astype(str).str.strip()
    rake["wagons"] = pd.to_numeric(rake["wagons"], errors="coerce")
    rake["debit_min"] = pd.to_numeric(rake["debit_min"], errors="coerce").fillna(0)
    rake["demurrage_hr"] = np.ceil(rake["debit_min"] / 60.0)
    rake["demurrage_amount"] = [
        _demurrage_amount(h, w) for h, w in zip(rake["demurrage_hr"], rake["wagons"])
    ]
    rake["cycle_hours"] = rake["cycle_raw"].apply(_hhmm_to_hours)
    rake["unload_hours"] = rake["unload_raw"].apply(_hhmm_to_hours)
    rake["coal_mines_mt"] = pd.to_numeric(rake["coal_mines_mt"], errors="coerce")
    rake["valid"] = rake["round_no"].notna() & rake["date"].notna()
    rake = rake[rake["valid"]].reset_index(drop=True)
    rake["month"] = rake["date"].dt.to_period("M")

    return summary, rake


def _share_token(url: str) -> str:
    """Microsoft 'shares' API encoding of a sharing URL: u!<urlsafe-base64, no padding>."""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return "u!" + b64


def _looks_like_xlsx(content: bytes) -> bool:
    return len(content) > 2000 and content[:2] == b"PK"  # .xlsx is a zip file


def fetch_from_onedrive(share_url: str, timeout: int = 25) -> bytes:
    """Download the workbook from an OneDrive/SharePoint share link without any
    sign-in. Tries several anonymous download strategies. Works when the link is
    shared as 'Anyone with the link'. Raises RuntimeError if all strategies fail."""
    import requests
    from urllib.parse import urlsplit

    parts = urlsplit(share_url)
    host = f"{parts.scheme}://{parts.netloc}"
    headers = {"User-Agent": "Mozilla/5.0 (CHP-Dashboard)"}
    sep = "&" if "?" in share_url else "?"

    candidates = [
        share_url + sep + "download=1",                                     # direct download flag
        f"{host}/_api/v2.0/shares/{_share_token(share_url)}/driveitem/content",  # shares API
        f"{host}/_api/v2.0/shares/{_share_token(share_url.split('?')[0])}/driveitem/content",
    ]

    errors = []
    for url in candidates:
        try:
            resp = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
            if resp.status_code == 200 and _looks_like_xlsx(resp.content):
                return resp.content
            errors.append(f"HTTP {resp.status_code} ({resp.headers.get('Content-Type', '?')})")
        except Exception as exc:  # noqa: BLE001
            errors.append(type(exc).__name__)

    raise RuntimeError(
        "Could not download the workbook anonymously from OneDrive "
        f"({'; '.join(errors)}). Make sure the link is shared as 'Anyone with the link' "
        "(not 'People in NTPC' / specific people). You can still upload the file manually."
    )


def _load_local_copy():
    for name in DEFAULT_FILE_NAMES:
        p = APP_DIR / name
        if p.exists():
            return p.read_bytes(), p.name
    return None, None


def _set_source(content: bytes, label: str):
    st.session_state["wb_bytes"] = content
    st.session_state["data_source"] = label
    st.session_state["loaded_at"] = datetime.now()


# ---------------------------------------------------------------------------
# Sidebar — data source (auto-refresh on every new session / app open)
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## 🚂 NTPC DADRI CHP Dashboard")
    st.caption("Coal Handling Plant — DADRI")
    st.markdown("### Data Source")

    # --- Auto-refresh when the app is opened (once per browser session) ---
    if "wb_bytes" not in st.session_state:
        try:
            with st.spinner("Fetching latest workbook from OneDrive..."):
                _set_source(fetch_from_onedrive(SHAREPOINT_SHARE_URL), "OneDrive (live link)")
        except Exception as exc:  # noqa: BLE001
            local_bytes, local_name = _load_local_copy()
            if local_bytes is not None:
                _set_source(local_bytes, f"Local fallback copy ({local_name})")
                st.session_state["fetch_error"] = str(exc)
            else:
                st.session_state["wb_bytes"] = None
                st.session_state["data_source"] = "No data loaded yet"
                st.session_state["fetch_error"] = str(exc)

    if st.session_state.get("fetch_error"):
        st.warning(st.session_state["fetch_error"])

    st.markdown(f"**Current source:** {st.session_state['data_source']}")
    if st.session_state.get("loaded_at"):
        st.caption(f"Last loaded: {st.session_state['loaded_at']:%d %b %Y, %H:%M:%S}")

    if st.button("🔄 Refresh from OneDrive now", width="stretch"):
        with st.spinner("Downloading latest workbook..."):
            try:
                _set_source(fetch_from_onedrive(SHAREPOINT_SHARE_URL), "OneDrive (live link)")
                st.session_state.pop("fetch_error", None)
                parse_workbook.clear()
                st.rerun()
            except Exception as exc:  # noqa: BLE001
                st.warning(str(exc))

    uploaded = st.file_uploader("…or upload the workbook (.xlsx)", type=["xlsx"])
    if uploaded is not None:
        content = uploaded.getvalue()
        if content != st.session_state.get("_last_uploaded"):
            st.session_state["_last_uploaded"] = content
            _set_source(content, f"Uploaded file ({uploaded.name})")
            st.session_state.pop("fetch_error", None)
            parse_workbook.clear()
            st.rerun()

    if st.session_state["wb_bytes"] is None:
        st.info("Could not load data automatically. Upload the CHP workbook to get started.")
        st.stop()

# ---------------------------------------------------------------------------
# Parse data
# ---------------------------------------------------------------------------

try:
    summary_df, rake_df = parse_workbook(st.session_state["wb_bytes"])
except Exception as exc:  # noqa: BLE001
    st.error(f"Couldn't read that workbook — make sure it has 'Summary' and "
             f"'RAKESTATUS' sheets in the expected layout.\n\nDetails: {exc}")
    st.stop()

if summary_df.empty and rake_df.empty:
    st.error("No usable rows found in the 'Summary' or 'RAKESTATUS' sheets.")
    st.stop()

# ---------------------------------------------------------------------------
# Sidebar — filters
# ---------------------------------------------------------------------------

all_dates = pd.concat([summary_df["date"], rake_df["date"]]).dropna()
min_date, max_date = all_dates.min().date(), all_dates.max().date()

with st.sidebar:
    st.markdown("### Filters")
    date_range = st.date_input(
        "Date range", value=(min_date, max_date), min_value=min_date, max_value=max_date,
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date, end_date = min_date, max_date

    rake_type_opts = ["All"] + sorted(rake_df["rake_type"].dropna().unique().tolist())
    sel_rake_type = st.selectbox("Rake Type", rake_type_opts, index=0)

    colliery_opts = ["All"] + sorted(rake_df["colliery"].dropna().unique().tolist())
    sel_colliery = st.selectbox("Colliery", colliery_opts, index=0)

    siding_opts = ["All"] + sorted(rake_df["siding"].dropna().unique().tolist())
    sel_siding = st.selectbox("Siding / Mine", siding_opts, index=0)

start_ts, end_ts = pd.Timestamp(start_date), pd.Timestamp(end_date)

mask_summary = (summary_df["date"] >= start_ts) & (summary_df["date"] <= end_ts)
sdf = summary_df.loc[mask_summary].copy()

mask_date = (rake_df["date"] >= start_ts) & (rake_df["date"] <= end_ts)
mask_type = (sel_rake_type == "All") | (rake_df["rake_type"] == sel_rake_type)
mask_colliery = (sel_colliery == "All") | (rake_df["colliery"] == sel_colliery)
mask_siding = (sel_siding == "All") | (rake_df["siding"] == sel_siding)

mask_full = mask_date & mask_type & mask_colliery & mask_siding
mask_excl_colliery = mask_date & mask_type & mask_siding
mask_excl_type = mask_date & mask_colliery & mask_siding
mask_excl_siding = mask_date & mask_type & mask_colliery

rdf = rake_df.loc[mask_full].copy()

with st.sidebar:
    if len(rake_df):
        completeness = rake_df["debit_min"].notna().mean()
        st.caption(f"Data completeness: {completeness:.0%} of rake rows have a demurrage entry.")

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

st.title("🚂 CHP Coal Feeding & Rake Demurrage Dashboard")
st.caption(
    f"NTPC Dadri CHP · {start_date:%d %b %Y} – {end_date:%d %b %Y} · "
    f"Rake Type: {sel_rake_type} · Colliery: {sel_colliery} · Siding: {sel_siding} · "
    f"{len(sdf)} operating day(s) · {len(rdf)} rake(s) in scope"
)
st.divider()

# ---------------------------------------------------------------------------
# KPI Row 1
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title">Headline KPIs</div>', unsafe_allow_html=True)

c_coal = _col(sdf, "coal bunkered (mt)")
c_stack = _col(sdf, "stacking in (mt)")
c_reclaim = _col(sdf, "total reclaiming")
c_biomass = _col(sdf, "biomass feeding total")
c_buf_total = _col(sdf, "buf total (average) (%)")

row1 = st.columns(4)
row1[0].metric("Coal Bunkered (MT)", f"{sdf[c_coal].sum():,.0f}")
row1[1].metric("Stacking In (MT)", f"{sdf[c_stack].sum():,.0f}")
row1[2].metric("Total Reclaiming (MT)", f"{sdf[c_reclaim].sum():,.0f}")
row1[3].metric("Biomass Blended (MT)", f"{sdf[c_biomass].sum():,.0f}")

row2 = st.columns(4)
avg_buf = sdf[c_buf_total].mean()
row2[0].metric("Avg BUF — All Conveyors (%)", f"{avg_buf:,.2f}%" if pd.notna(avg_buf) else "—")
row2[1].metric("Demurrage (HR)", f"{rdf['demurrage_hr'].sum():,.0f}")
row2[2].metric("No. of Rakes", f"{len(rdf):,}")
row2[3].metric("Demurrage Amount (Rs)", f"₹{rdf['demurrage_amount'].sum():,.0f}")

st.divider()

# ---------------------------------------------------------------------------
# KPI Row 2
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title">Rake & Demurrage Detail</div>', unsafe_allow_html=True)

total_rakes = len(rdf)
total_wagons = rdf["wagons"].sum()
avg_wagons_per_rake = total_wagons / total_rakes if total_rakes else np.nan
total_demurrage_hr = rdf["demurrage_hr"].sum()
avg_demurrage_hr = total_demurrage_hr / total_rakes if total_rakes else np.nan
pct_zero_demurrage = rdf["demurrage_hr"].eq(0).mean() if total_rakes else np.nan
avg_cycle_hours = rdf["cycle_hours"].mean()
avg_unload_hours = rdf["unload_hours"].mean()

boxn_split = rake_df.loc[mask_excl_type]
boxn_count = (boxn_split["rake_type"] == "BOXN").sum()
bobr_count = (boxn_split["rake_type"] == "BOBR").sum()

avg_coal_per_rake = rdf["coal_mines_mt"].sum() / total_rakes if total_rakes else np.nan
days_with_data = sdf["date"].notna().sum()
idle_days = (sdf[c_stack] == 0).sum()

if len(sdf) and sdf[c_coal].notna().any():
    peak_idx = sdf[c_coal].idxmax()
    peak_bunkering = sdf.loc[peak_idx, c_coal]
    peak_bunkering_date = sdf.loc[peak_idx, "date"]
else:
    peak_bunkering, peak_bunkering_date = np.nan, None

total_demurrage_amt = rdf["demurrage_amount"].sum()
avg_demurrage_amt_per_rake = total_demurrage_amt / total_rakes if total_rakes else np.nan

r2c1 = st.columns(6)
r2c1[0].metric("Total Rakes", f"{total_rakes:,}")
r2c1[1].metric("Total Wagons", f"{total_wagons:,.0f}")
r2c1[2].metric("Avg Wagons / Rake", f"{avg_wagons_per_rake:,.1f}" if pd.notna(avg_wagons_per_rake) else "—")
r2c1[3].metric("Total Demurrage (HR)", f"{total_demurrage_hr:,.0f}")
r2c1[4].metric("Avg Demurrage / Rake (HR)", f"{avg_demurrage_hr:,.2f}" if pd.notna(avg_demurrage_hr) else "—")
r2c1[5].metric("% Rakes, Zero Demurrage", f"{pct_zero_demurrage:.0%}" if pd.notna(pct_zero_demurrage) else "—")

r2c2 = st.columns(6)
r2c2[0].metric("Avg Cycle Time (H:MM)", _fmt_hm(avg_cycle_hours))
r2c2[1].metric("Avg Unloading Time (H:MM)", _fmt_hm(avg_unload_hours))
r2c2[2].metric("BOXN Rakes (of filtered)", f"{boxn_count:,}")
r2c2[3].metric("BOBR Rakes (of filtered)", f"{bobr_count:,}")
r2c2[4].metric("Avg Coal Unloaded / Rake (MT)", f"{avg_coal_per_rake:,.1f}" if pd.notna(avg_coal_per_rake) else "—")
r2c2[5].metric("Days with Data", f"{days_with_data:,}")

r2c3 = st.columns(6)
r2c3[0].metric("Idle Days (0 Stacking-In)", f"{idle_days:,}")
r2c3[1].metric("Peak Bunkering (MT)", f"{peak_bunkering:,.0f}" if pd.notna(peak_bunkering) else "—")
r2c3[2].metric("Peak Bunkering Date", f"{peak_bunkering_date:%d %b %Y}" if peak_bunkering_date is not None else "—")
r2c3[3].metric("Total Demurrage Amount (Rs)", f"₹{total_demurrage_amt:,.0f}")
r2c3[4].metric("Avg Demurrage Amount / Rake (Rs)", f"₹{avg_demurrage_amt_per_rake:,.0f}" if pd.notna(avg_demurrage_amt_per_rake) else "—")
r2c3[5].empty()

st.divider()

# ---------------------------------------------------------------------------
# Sector-wise Coal Movement & Stage-wise Split
# ---------------------------------------------------------------------------

sec_left, sec_right = st.columns(2)

sectors = list(range(1, 13))
stack_cols = {n: _col(sdf, f"stacking in sec {n}") for n in sectors}
reclaim_cols = {n: _col(sdf, f"total reclaiming in sec {n}") for n in sectors}

with sec_left:
    st.markdown('<div class="section-title">Sector-wise Coal Movement (MT)</div>', unsafe_allow_html=True)
    sec_df = pd.DataFrame(
        {
            "Sector": [f"Sec-{n}" for n in sectors],
            "Stacking In": [sdf[stack_cols[n]].sum() for n in sectors],
            "Reclaiming": [sdf[reclaim_cols[n]].sum() for n in sectors],
        }
    )
    fig_sec = px.bar(sec_df, x="Sector", y=["Stacking In", "Reclaiming"], barmode="group")
    fig_sec.update_layout(height=380, legend=dict(orientation="h", y=1.15), margin=dict(t=10, l=10, r=10, b=10),
                          yaxis_title="MT")
    st.plotly_chart(_light(fig_sec), width="stretch")

with sec_right:
    st.markdown('<div class="section-title">Stage-wise Split (MT)</div>', unsafe_allow_html=True)
    stage_df = pd.DataFrame(
        {
            "Stage": ["Stage-1 (Sec 1-8)", "Stage-2 (Sec 9-12)"],
            "Stacking In": [
                sum(sdf[stack_cols[n]].sum() for n in range(1, 9)),
                sum(sdf[stack_cols[n]].sum() for n in range(9, 13)),
            ],
            "Reclaiming": [
                sum(sdf[reclaim_cols[n]].sum() for n in range(1, 9)),
                sum(sdf[reclaim_cols[n]].sum() for n in range(9, 13)),
            ],
        }
    )
    fig_stage = px.bar(stage_df, x="Stage", y=["Stacking In", "Reclaiming"], barmode="group")
    fig_stage.update_layout(height=380, legend=dict(orientation="h", y=1.15), margin=dict(t=10, l=10, r=10, b=10),
                            yaxis_title="MT")
    st.plotly_chart(_light(fig_stage), width="stretch")

# ---------------------------------------------------------------------------
# Average BUF (%)
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title">Average BUF (%) — Filtered</div>', unsafe_allow_html=True)
buf_conveyors = ["1A", "1B", "3A", "3B", "9A", "9B", "14A", "14B"]
buf_cols = {name: _col(sdf, f"buf {name.lower()} (%)") for name in buf_conveyors}
buf_vals = {name: sdf[col].mean() for name, col in buf_cols.items()}
buf_vals["Total (Avg)"] = sdf[c_buf_total].mean()
buf_series = pd.Series(buf_vals).sort_values()
fig_buf = px.bar(buf_series, orientation="h", labels={"value": "Avg BUF (%)", "index": "Conveyor"})
fig_buf.update_layout(height=360, showlegend=False, margin=dict(t=10, l=10, r=10, b=10))
st.plotly_chart(_light(fig_buf), width="stretch")

st.divider()

# ---------------------------------------------------------------------------
# Monthly Rake Count vs Demurrage
# ---------------------------------------------------------------------------

mo_left, mo_right = st.columns(2)

with mo_left:
    st.markdown('<div class="section-title">Monthly Rake Count vs Demurrage (HR)</div>', unsafe_allow_html=True)
    if len(rdf):
        mo = rdf.groupby("month").agg(rakes=("round_no", "count"), demurrage=("demurrage_hr", "sum")).reset_index()
        mo["month_label"] = mo["month"].astype(str)
        fig_mo = go.Figure()
        fig_mo.add_trace(go.Bar(x=mo["month_label"], y=mo["rakes"], name="Rakes", marker_color="#5BA8D0", yaxis="y"))
        fig_mo.add_trace(go.Scatter(x=mo["month_label"], y=mo["demurrage"], name="Demurrage (HR)",
                                    mode="lines+markers", line=dict(color="#D62828", width=3), yaxis="y2"))
        fig_mo.update_layout(
            height=380, margin=dict(t=10, l=10, r=10, b=10),
            yaxis=dict(title="No. of Rakes"),
            yaxis2=dict(title="Demurrage (HR)", overlaying="y", side="right"),
            legend=dict(orientation="h", y=1.15),
        )
        st.plotly_chart(_light(fig_mo), width="stretch")
    else:
        st.info("No rake data in the selected filters.")

with mo_right:
    st.markdown('<div class="section-title">Demurrage Amount (Rs) by Month</div>', unsafe_allow_html=True)
    if len(rdf):
        amt_mo = rdf.groupby("month")["demurrage_amount"].sum().reset_index()
        amt_mo["month_label"] = amt_mo["month"].astype(str)
        fig_amt = px.bar(amt_mo, x="month_label", y="demurrage_amount",
                         labels={"demurrage_amount": "Amount (Rs)", "month_label": "Month"})
        fig_amt.update_layout(height=380, showlegend=False, margin=dict(t=10, l=10, r=10, b=10))
        st.plotly_chart(_light(fig_amt), width="stretch")
    else:
        st.info("No rake data in the selected filters.")

# ---------------------------------------------------------------------------
# Colliery-wise table
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title">Colliery-wise Rake & Demurrage Summary</div>', unsafe_allow_html=True)
colliery_scope = rake_df.loc[mask_excl_colliery]
grp = pd.DataFrame()
if len(colliery_scope):
    grp = colliery_scope.groupby("colliery").agg(
        Rakes=("round_no", "count"),
        Wagons=("wagons", "sum"),
        **{"Total Demurrage (HR)": ("demurrage_hr", "sum")},
    ).reset_index().rename(columns={"colliery": "Colliery"})
    grp["Avg Wagons/Rake"] = (grp["Wagons"] / grp["Rakes"]).round(2)
    grp["Avg Demurrage/Rake (HR)"] = (grp["Total Demurrage (HR)"] / grp["Rakes"]).round(2)
    grp["% Share of Rakes"] = (grp["Rakes"] / grp["Rakes"].sum() * 100).round(1)
    grp = grp.sort_values("Rakes", ascending=False)[
        ["Colliery", "Rakes", "Wagons", "Avg Wagons/Rake", "Total Demurrage (HR)",
         "Avg Demurrage/Rake (HR)", "% Share of Rakes"]
    ]
    st.dataframe(grp, width="stretch", hide_index=True)
else:
    st.info("No rake data in the selected filters.")

st.divider()

# ---------------------------------------------------------------------------
# Insights & Highlights
# ---------------------------------------------------------------------------

st.markdown('<div class="section-title">🔍 Insights & Highlights</div>', unsafe_allow_html=True)


def _build_highlights() -> list[str]:
    bullets: list[str] = []

    if len(sdf) and sdf[c_coal].notna().any():
        best_i, worst_i = sdf[c_coal].idxmax(), sdf[c_coal].idxmin()
        bullets.append(
            f"📈 Best bunkering day was **{sdf.loc[best_i,'date']:%d %b}** "
            f"({sdf.loc[best_i,c_coal]:,.0f} MT); lowest was **{sdf.loc[worst_i,'date']:%d %b}** "
            f"({sdf.loc[worst_i,c_coal]:,.0f} MT)."
        )

    stack_total = sdf[c_stack].sum()
    reclaim_total = sdf[c_reclaim].sum()
    if pd.notna(stack_total) and pd.notna(reclaim_total):
        net = reclaim_total - stack_total
        direction = "drew down" if net > 0 else "built up"
        bullets.append(
            f"📦 Yard stock **{direction}** by **{abs(net):,.0f} MT** over the period "
            f"(Stacking-In {stack_total:,.0f} MT vs Reclaiming {reclaim_total:,.0f} MT)."
        )

    if total_rakes:
        bullets.append(
            f"✅ **{pct_zero_demurrage:.0%}** of rakes in scope incurred **zero demurrage** "
            f"({total_rakes - int(rdf['demurrage_hr'].gt(0).sum())} of {total_rakes})."
        )

    if buf_vals:
        weak_conv = min((k for k in buf_vals if k != "Total (Avg)"),
                        key=lambda k: buf_vals[k] if pd.notna(buf_vals[k]) else np.inf)
        weak_val = buf_vals.get(weak_conv, np.nan)
        if pd.notna(weak_val):
            if weak_val < 0:
                bullets.append(
                    f"⚠️ Conveyor **{weak_conv}** shows an anomalous **negative average BUF ({weak_val:.1f}%)** "
                    f"— likely a bad sensor reading on one or more days; worth checking the source data."
                )
            else:
                bullets.append(f"⚙️ Conveyor **{weak_conv}** had the lowest average BUF at **{weak_val:.1f}%**.")

    if len(grp):
        top_coll = grp.sort_values("Total Demurrage (HR)", ascending=False).iloc[0]
        if top_coll["Total Demurrage (HR)"] > 0:
            bullets.append(
                f"⛏️ **{top_coll['Colliery']}** accounted for the most demurrage hours "
                f"(**{top_coll['Total Demurrage (HR)']:,.0f} HR** across {top_coll['Rakes']:,.0f} rakes)."
            )

    if total_rakes:
        by_type = rdf.groupby("rake_type")["demurrage_hr"].mean()
        if len(by_type) == 2:
            bullets.append(
                f"🚋 **{by_type.idxmax()}** rakes averaged more demurrage per rake "
                f"({by_type.max():.2f} HR) than the other type ({by_type.min():.2f} HR)."
            )

    if total_rakes and rdf["demurrage_amount"].notna().any():
        worst_rake = rdf.loc[rdf["demurrage_amount"].idxmax()]
        bullets.append(
            f"🔴 The single costliest rake was on **{worst_rake['date']:%d %b}** "
            f"({worst_rake['rake_type']} from {worst_rake['colliery']}) at "
            f"**₹{worst_rake['demurrage_amount']:,.0f}** ({worst_rake['demurrage_hr']:.0f} HR demurrage)."
        )

    return bullets


highlights = _build_highlights()
if highlights:
    st.markdown("\n".join(f"- {b}" for b in highlights))
else:
    st.info("Not enough data in the current filter selection to generate highlights.")

ins_left, ins_right = st.columns(2)

with ins_left:
    st.markdown('<div class="subsection-title">Daily Coal Bunkered vs Reclaiming (MT)</div>', unsafe_allow_html=True)
    if len(sdf):
        trend_df = sdf[["date", c_coal, c_reclaim]].rename(columns={c_coal: "Coal Bunkered", c_reclaim: "Reclaiming"})
        fig_trend1 = px.line(trend_df, x="date", y=["Coal Bunkered", "Reclaiming"], markers=True)
        fig_trend1.update_layout(height=340, margin=dict(t=10, l=10, r=10, b=10),
                                 legend=dict(orientation="h", y=1.15), yaxis_title="MT", xaxis_title="")
        st.plotly_chart(_light(fig_trend1), width="stretch")
    else:
        st.info("No data in range.")

with ins_right:
    st.markdown('<div class="subsection-title">Daily Reclaiming vs Biomass Blended (MT)</div>', unsafe_allow_html=True)
    if len(sdf):
        trend_df2 = sdf[["date", c_reclaim, c_biomass]].rename(columns={c_reclaim: "Reclaiming", c_biomass: "Biomass Blended"})
        fig_trend2 = px.line(trend_df2, x="date", y=["Reclaiming", "Biomass Blended"], markers=True)
        fig_trend2.update_layout(height=340, margin=dict(t=10, l=10, r=10, b=10),
                                 legend=dict(orientation="h", y=1.15), yaxis_title="MT", xaxis_title="")
        st.plotly_chart(_light(fig_trend2), width="stretch")
    else:
        st.info("No data in range.")

st.markdown('<div class="subsection-title">BUF (%) Trend by Conveyor</div>', unsafe_allow_html=True)
if len(sdf):
    buf_trend_df = sdf[["date"] + list(buf_cols.values())].copy()
    buf_trend_df.columns = ["date"] + list(buf_cols.keys())
    buf_trend_df["Total (Avg)"] = sdf[c_buf_total]
    fig_buf_trend = px.line(buf_trend_df, x="date", y=list(buf_trend_df.columns[1:]), markers=True)
    fig_buf_trend.update_layout(height=360, margin=dict(t=10, l=10, r=10, b=10),
                                legend=dict(orientation="h", y=1.2), yaxis_title="BUF (%)", xaxis_title="")
    st.plotly_chart(_light(fig_buf_trend), width="stretch")
else:
    st.info("No data in range.")

ins2_left, ins2_right = st.columns(2)

with ins2_left:
    st.markdown('<div class="subsection-title">Rake Type Comparison — BOXN vs BOBR</div>', unsafe_allow_html=True)
    if total_rakes:
        cmp = rdf.groupby("rake_type").agg(
            Rakes=("round_no", "count"),
            **{"Total Wagons": ("wagons", "sum")},
            **{"Avg Wagons/Rake": ("wagons", "mean")},
            **{"Total Demurrage (HR)": ("demurrage_hr", "sum")},
            **{"Avg Demurrage/Rake (HR)": ("demurrage_hr", "mean")},
            **{"Avg Cycle Time (Hrs)": ("cycle_hours", "mean")},
            **{"Avg Unloading Time (Hrs)": ("unload_hours", "mean")},
            **{"Total Demurrage Amount (Rs)": ("demurrage_amount", "sum")},
        ).round(2).reset_index().rename(columns={"rake_type": "Rake Type"})
        st.dataframe(cmp, width="stretch", hide_index=True)
    else:
        st.info("No rake data in the selected filters.")

with ins2_right:
    st.markdown('<div class="subsection-title">Demurrage Hours — Distribution</div>', unsafe_allow_html=True)
    if total_rakes:
        fig_hist = px.histogram(rdf, x="demurrage_hr", nbins=20, labels={"demurrage_hr": "Demurrage (HR) per rake"})
        fig_hist.update_layout(height=320, margin=dict(t=10, l=10, r=10, b=10), showlegend=False,
                               yaxis_title="No. of Rakes")
        st.plotly_chart(_light(fig_hist), width="stretch")
    else:
        st.info("No rake data in the selected filters.")

st.markdown('<div class="subsection-title">Siding-wise Rake & Demurrage Summary</div>', unsafe_allow_html=True)
siding_scope = rake_df.loc[mask_excl_siding]
if len(siding_scope):
    sgrp = siding_scope.groupby("siding").agg(
        Rakes=("round_no", "count"),
        Wagons=("wagons", "sum"),
        **{"Total Demurrage (HR)": ("demurrage_hr", "sum")},
        **{"Total Demurrage Amount (Rs)": ("demurrage_amount", "sum")},
    ).reset_index().rename(columns={"siding": "Siding"})
    sgrp["Avg Wagons/Rake"] = (sgrp["Wagons"] / sgrp["Rakes"]).round(2)
    sgrp["Avg Demurrage/Rake (HR)"] = (sgrp["Total Demurrage (HR)"] / sgrp["Rakes"]).round(2)
    sgrp = sgrp.sort_values("Rakes", ascending=False)[
        ["Siding", "Rakes", "Wagons", "Avg Wagons/Rake", "Total Demurrage (HR)",
         "Avg Demurrage/Rake (HR)", "Total Demurrage Amount (Rs)"]
    ]
    st.dataframe(sgrp, width="stretch", hide_index=True)
else:
    st.info("No rake data in the selected filters.")

st.markdown('<div class="subsection-title">Top 10 Rakes by Demurrage</div>', unsafe_allow_html=True)
if total_rakes and rdf["demurrage_hr"].gt(0).any():
    top10 = rdf[rdf["demurrage_hr"] > 0].sort_values("demurrage_amount", ascending=False).head(10)
    top10_show = top10[["date", "rake_type", "colliery", "siding", "wagons", "demurrage_hr", "demurrage_amount",
                        "cycle_hours", "unload_hours", "mgr_remarks", "chp_remarks"]].rename(columns={
        "date": "Date", "rake_type": "Rake Type", "colliery": "Colliery", "siding": "Siding",
        "wagons": "Wagons", "demurrage_hr": "Demurrage (HR)", "demurrage_amount": "Demurrage Amount (Rs)",
        "cycle_hours": "Cycle Time (Hrs)", "unload_hours": "Unloading Time (Hrs)",
        "mgr_remarks": "MGR Remarks", "chp_remarks": "CHP Remarks",
    })
    st.dataframe(
        top10_show, width="stretch", hide_index=True,
        column_config={
            "MGR Remarks": st.column_config.TextColumn("MGR Remarks", width="large"),
            "CHP Remarks": st.column_config.TextColumn("CHP Remarks", width="large"),
        },
    )
else:
    st.info("No rakes with demurrage in the selected filters.")

st.divider()

# ---------------------------------------------------------------------------
# Raw data / export
# ---------------------------------------------------------------------------

with st.expander("📄 Raw data (filtered) & downloads"):
    st.markdown('<div class="subsection-title">Summary sheet (filtered)</div>', unsafe_allow_html=True)
    st.dataframe(sdf, width="stretch", hide_index=True)
    st.download_button("Download Summary CSV", sdf.to_csv(index=False).encode(), "summary_filtered.csv")

    st.markdown('<div class="subsection-title">Rake-level data (filtered)</div>', unsafe_allow_html=True)
    show_cols = ["date", "rake_type", "colliery", "siding", "wagons", "cycle_hours", "unload_hours",
                 "demurrage_hr", "demurrage_amount", "coal_mines_mt", "mgr_remarks", "chp_remarks"]
    show_cols = [c for c in show_cols if c in rdf.columns]
    st.dataframe(rdf[show_cols], width="stretch", hide_index=True)
    st.download_button("Download Rake Data CSV", rdf[show_cols].to_csv(index=False).encode(), "rake_data_filtered.csv")

st.divider()
st.caption("Mirrors the 'Dashboard' sheet of the CHP workbook (Summary + RAKESTATUS sheets). "
           "Demurrage amount uses the same tariff slabs as the workbook: ₹150/hr (0-6h), "
           "₹165/hr (6-12h), ₹187.5/hr (12-24h), per wagon.")
