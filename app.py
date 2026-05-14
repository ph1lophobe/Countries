"""
Relocation Intelligence 2026
─────────────────────────────
Streamlit dashboard — 31 cities, 80+ metrics.

Requirements (put in requirements.txt):
    streamlit>=1.35.0
    pandas>=2.0.0
    numpy>=1.24.0
    plotly>=5.20.0
    openpyxl>=3.1.0
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ─── CONSTANTS ────────────────────────────────────────────────────────────────
DATA_PATH = Path(__file__).resolve().parent / "Country.xlsx - Лист1.csv"

# Maps city name → country label as it appears in the CSV country-header row
CITY_TO_COUNTRY: Dict[str, str] = {
    "Auckland": "New Zealand 🇳🇿",   "Wellington": "New Zealand 🇳🇿",
    "Christchurch": "New Zealand 🇳🇿",
    "Melbourne": "Australia 🇦🇺",    "Sydney": "Australia 🇦🇺",
    "Brisbane": "Australia 🇦🇺",     "Perth": "Australia 🇦🇺",
    "Adelaide": "Australia 🇦🇺",
    "Toronto": "Canada 🇨🇦",         "Montreal": "Canada 🇨🇦",
    "Ottawa": "Canada 🇨🇦",
    "Chicago": "USA 🇺🇸",            "Philadelphia": "USA 🇺🇸",
    "San Diego": "USA 🇺🇸",          "Dallas": "USA 🇺🇸",
    "Austin": "USA 🇺🇸",             "Seattle": "USA 🇺🇸",
    "Birmingham": "United Kingdom 🇬🇧", "Liverpool": "United Kingdom 🇬🇧",
    "Southhampton": "United Kingdom 🇬🇧", "Manchester": "United Kingdom 🇬🇧",
    "Amsterdam": "Netherlands 🇳🇱",  "Rotterdam": "Netherlands 🇳🇱",
    "Oslo": "Norway 🇳🇴",
    "Belgrade": "Serbia 🇷🇸",
    "Warsaw": "Poland 🇵🇱",          "Krakow": "Poland 🇵🇱",
    "Lodz": "Poland 🇵🇱",            "Gdansk": "Poland 🇵🇱",
    "Moscow": "Russia 🇷🇺",
    "Minsk": "Belarus",
}

PLOTLY_CFG = {"displayModeBar": True, "displaylogo": False,
              "responsive": True, "scrollZoom": False}

COLORS = px.colors.qualitative.Set2

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Relocation Intelligence 2026",
    layout="wide",
    initial_sidebar_state="collapsed",
    page_icon="🌍",
)

st.markdown("""
<style>
[data-testid="stMetricValue"]  { font-size:1.35rem; }
.block-container {
    padding-top:.75rem !important;
    padding-left:min(1.25rem,4vw) !important;
    padding-right:min(1.25rem,4vw) !important;
    max-width:1420px;
}
@media(max-width:768px){
    [data-testid="stMetricValue"] { font-size:1rem !important; }
    h1 { font-size:1.3rem !important; }
    h2,h3 { font-size:1rem !important; }
    [data-testid="stTabs"] button {
        padding-left:.3rem !important;
        padding-right:.3rem !important;
        font-size:.75rem !important;
    }
}
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING ─────────────────────────────────────────────────────────────
def _num(v: Any) -> float:
    """Parse a cell value to float, handling EU decimals, thousands sep, k-suffix, $, %."""
    if pd.isna(v):
        return np.nan
    if isinstance(v, (int, np.integer, float, np.floating)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace("\u00a0", "").replace("$", "").replace("%", "")
    if not s or s.lower() in ("nan", "na", "n/a", "-", ""):
        return np.nan
    mult = 1.0
    if s.lower().endswith("k"):
        mult, s = 1_000.0, s[:-1].strip()
    sc = s.replace(" ", "")
    # thousands with commas: 1,234 or 1,234,567
    if re.fullmatch(r"-?\d{1,3}(?:,\d{3})+", sc):
        return float(sc.replace(",", "")) * mult
    if "," in sc and "." in sc:
        sc = sc.replace(".", "").replace(",", ".")
    elif "," in sc:
        sc = sc.replace(",", ".")
    try:
        return float(sc) * mult
    except ValueError:
        return np.nan


@st.cache_data(show_spinner="Loading data…")
def load_data() -> Tuple[pd.DataFrame, List[str], Optional[str]]:
    if not DATA_PATH.is_file():
        return pd.DataFrame(), [], f"File not found: {DATA_PATH}"
    try:
        raw = pd.read_csv(DATA_PATH, header=None, encoding="utf-8-sig", engine="python")
    except Exception as exc:
        return pd.DataFrame(), [], str(exc)

    # ── locate the city-header row (col C == "Auckland") ──
    city_hdr = -1
    for i in range(len(raw)):
        if str(raw.iloc[i, 2]).strip() == "Auckland":
            city_hdr = i
            break
    if city_hdr == -1:
        return pd.DataFrame(), [], "City header row not found."

    # ── locate the country-header row (col B == "Country") ──
    country_hdr = -1
    for i in range(city_hdr):
        if str(raw.iloc[i, 1]).strip().lower() == "country":
            country_hdr = i
            break

    # build col-index → city-name map
    hrow = raw.iloc[city_hdr]
    skip_vals = {"nan", "city", "country", "category", "spendings", ""}
    city_map: Dict[int, str] = {
        j: str(hrow.iloc[j]).strip()
        for j in range(2, len(hrow))
        if str(hrow.iloc[j]).strip().lower() not in skip_vals
    }
    cities: List[str] = list(city_map.values())

    skip_metrics = {
        "city", "country", "category", "spendings", "nan", "",
        "population, kk", "land area, km^2", "last update",
    }

    def _parse_rows(row_start: int, row_end: int) -> List[Dict[str, Any]]:
        rows = []
        for ri in range(row_start, row_end):
            name = str(raw.iloc[ri, 1]).strip()
            if not name or name.lower() in skip_metrics:
                continue
            r: Dict[str, Any] = {"Metric": name}
            for j, city in city_map.items():
                r[city] = _num(raw.iloc[ri, j])
            rows.append(r)
        return rows

    # city-level rows (below city header)
    parsed = _parse_rows(city_hdr + 1, len(raw))
    city_metric_names = {r["Metric"] for r in parsed}

    # country-level rows (between country header and city header)
    # broadcast each country value to all its cities
    if country_hdr >= 0:
        ch = raw.iloc[country_hdr]
        country_col_map: List[Tuple[int, str]] = [
            (j, str(ch.iloc[j]).strip())
            for j in range(2, len(ch))
            if str(ch.iloc[j]).strip().lower() not in skip_vals
        ]
        for ri in range(country_hdr + 1, city_hdr):
            name = str(raw.iloc[ri, 1]).strip()
            if not name or name.lower() in skip_metrics or name in city_metric_names:
                continue
            country_vals = {lbl: _num(raw.iloc[ri, j]) for j, lbl in country_col_map}
            r: Dict[str, Any] = {"Metric": name}
            for city in cities:
                cc = CITY_TO_COUNTRY.get(city)
                r[city] = country_vals.get(cc, np.nan) if cc else np.nan
            parsed.append(r)

    df = pd.DataFrame(parsed)
    # drop exact-duplicate metric names (keep first)
    df = df[~df["Metric"].duplicated(keep="first")]
    return df, cities, None


df, CITIES, _err = load_data()
if _err:
    st.error(_err)
if df.empty:
    st.error("Could not load data — check that the CSV is in the same folder as app.py.")
    st.stop()

MINDEX = df.set_index("Metric")


# ─── HELPERS ──────────────────────────────────────────────────────────────────
def row(metric: str, cities: List[str]) -> pd.Series:
    if metric not in MINDEX.index:
        return pd.Series({c: np.nan for c in cities}, dtype=float)
    # .loc[metric] returns a Series when the index is unique (we deduplicated on load),
    # but Pylance sees the ambiguous overload.  Slicing with [metric] and .iloc[0]
    # guarantees a Series regardless, and satisfies the type checker.
    result = MINDEX.loc[[metric]].iloc[0].reindex(cities)
    return pd.to_numeric(result, errors="coerce")


def v(series: pd.Series, city: str) -> float:
    val = series.get(city, np.nan)
    return float(val) if pd.notna(val) else np.nan


def fmt_usd(val) -> str:
    return f"${val:,.0f}" if pd.notna(val) and not np.isnan(val) else "—"


def fmtp(val) -> str:
    return f"{val:.1f}%" if pd.notna(val) and not np.isnan(val) else "—"


def saving_metric_name(people: int, all_exp: bool, inc: str) -> str:
    """Return the exact CSV metric string for the chosen savings scenario."""
    pool = "all" if all_exp else "base"
    if people == 1:
        sal = "Avg Dev salary" if inc == "dev" else "Avg salary"
        return f"Saving {pool} 1 people on {sal}, USD/mo "
    sal = "Avg Dev + Des salary" if inc == "dev_des" else "Avg Dev salary"
    return f"Saving {pool} 2 people on {sal}, USD/mo "


def chart(fig: go.Figure, h: int = 420) -> None:
    fig.update_layout(height=h, margin=dict(l=4, r=4, t=44, b=8))
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CFG)


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Timeline & Savings")
    start_year  = st.number_input("Start year",       2024, 2035, 2026)
    prep_years  = st.slider(      "Years to prepare", 1, 15, 5)
    target_year = start_year + prep_years
    st.divider()
    monthly_save = st.number_input("Monthly savings ($)", 0, 100_000, 2_000, step=100)
    growth       = st.slider(      "Savings growth (% / yr)", 0, 50, 10)
    st.divider()
    st.caption(f"Data: 09.05.2026 · {len(CITIES)} cities · {len(MINDEX)} metrics")

# ─── TITLE + CITY SELECTOR ────────────────────────────────────────────────────
st.title("🌍 Global Relocation Intelligence 2026")
st.caption(f"{len(CITIES)} cities · {len(MINDEX)} metrics · Updated May 2026")

DEFAULTS = [c for c in CITIES
            if any(n in c for n in ["Auckland","Warsaw","Toronto","Amsterdam","Belgrade"])][:5]
selected = st.multiselect("Select cities to compare:", options=CITIES, default=DEFAULTS)
if not selected:
    st.info("Choose at least one city above to begin.")
    st.stop()

# ─── PROFILE SWITCHERS (main area) ────────────────────────────────────────────
st.subheader("Calculation profile")
pc1, pc2, pc3 = st.columns(3)

with pc1:
    family = int(st.segmented_control(
        "Family size", options=[1, 2], default=1,
        format_func=lambda x: "Solo" if x == 1 else "Couple", key="fam",
    ) or 1)

with pc2:
    exp_mode   = st.segmented_control("Expenses", ["Base", "All"], default="Base", key="exp") or "Base"
    all_exp    = exp_mode == "All"

with pc3:
    if family == 1:
        inc = str(st.segmented_control(
            "Salary", ["avg", "dev"], default="dev",
            format_func=lambda x: "Average" if x == "avg" else "IT / Dev", key="inc1",
        ) or "dev")
    else:
        inc = str(st.segmented_control(
            "Couple income", ["dev", "dev_des"], default="dev",
            format_func=lambda x: "Both IT" if x == "dev" else "IT + Designer", key="inc2",
        ) or "dev")

exp_label   = "All" if all_exp else "Base"
EXP_METRIC  = f"Total {exp_label} / {family} people, USD/mo"
SAV_METRIC  = saving_metric_name(family, all_exp, inc)
SAL_METRIC  = "Avg Salary Dev, USD/mo" if (family == 2 or inc == "dev") else "Avg Salary, USD/mo"

with st.expander("Active CSV rows", expanded=False):
    st.caption(f"**Salary:** `{SAL_METRIC}`")
    st.caption(f"**Expenses:** `{EXP_METRIC}`")
    st.caption(f"**Saving:** `{SAV_METRIC.strip()}`")

st.divider()

# ─── ACCUMULATION CURVE (precompute once) ─────────────────────────────────────
_yr_range: List[int] = list(range(start_year, target_year + 1))
_balance, _annual, _accum = 0.0, monthly_save * 12.0, []
for _ in _yr_range:
    _balance += _annual
    _accum.append(_balance)
    _annual *= 1.0 + growth / 100.0
FINAL_SAVED: float = _accum[-1] if _accum else 0.0

# ══════════════════════════════════════════════════════════════════════════════
TABS = st.tabs([
    "📊 Overview",
    "💸 Finances",
    "🏠 Housing",
    "🛒 Cost of Living",
    "🌿 Quality of Life",
    "🛂 Immigration",
    "📈 Savings Plan",
    "🏆 Final Score",
])


# ══════════════════════════════════════════════════════════════════════════════
# 0 · OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with TABS[0]:
    sal_s  = row(SAL_METRIC,  selected)
    exp_s  = row(EXP_METRIC,  selected)
    sav_s  = row(SAV_METRIC,  selected)
    rent_s = row("1 bed. ap. in city centre", selected)
    qol_s  = row("Quality of life", selected)
    safe_s = row("Safety", selected)

    for city in selected:
        sv   = v(sal_s,  city)
        ev   = v(exp_s,  city)
        savv = v(sav_s,  city)
        net  = savv if pd.notna(savv) else (sv - ev if pd.notna(sv) and pd.notna(ev) else np.nan)
        rat  = net / sv * 100 if pd.notna(net) and sv and sv > 0 else np.nan
        with st.expander(f"**{city}**", expanded=len(selected) <= 4):
            c1, c2, c3 = st.columns(3)
            c4, c5, c6 = st.columns(3)
            c1.metric("💰 Salary",       fmt_usd(sv))
            c2.metric(f"💳 {exp_label} Expenses", fmt_usd(ev))
            c3.metric("📦 Net Save / mo", fmt_usd(net),
                      delta=fmtp(rat) + " of income" if pd.notna(rat) else "")
            c4.metric("🏠 Rent 1BR (ctr)", fmt_usd(v(rent_s, city)))
            c5.metric("🌟 Quality of Life", f"{v(qol_s, city):.0f}/250" if pd.notna(v(qol_s,city)) else "—")
            c6.metric("🛡️ Safety",         f"{v(safe_s, city):.0f}/100" if pd.notna(v(safe_s,city)) else "—")

    # ── Radar
    st.markdown("### 🕸️ Multi-Dimension Radar")
    R_METRICS = ["Quality of life","Safety","Purch, power","Health care","Climate","Internet Speed (Mbps)"]
    R_LABELS  = ["Quality of Life","Safety","Purch. Power","Healthcare","Climate","Internet"]
    R_MAXES   = [250, 100, 200, 100, 100, 350]
    fig_r = go.Figure()
    for i, city in enumerate(selected):
        vals = [v(row(m, selected), city) or 0 for m in R_METRICS]
        norm = [min(x / mx * 100, 100) for x, mx in zip(vals, R_MAXES)]
        fig_r.add_trace(go.Scatterpolar(
            r=norm + [norm[0]], theta=R_LABELS + [R_LABELS[0]],
            fill="toself", name=city, line_color=COLORS[i % len(COLORS)],
        ))
    fig_r.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0,100])))
    chart(fig_r, 520)

    # ── Quick comparison table
    st.markdown("### 📋 Quick Comparison")
    q_rows = {
        "Salary ($/mo)"          : SAL_METRIC,
        "Base Expenses 1 ($/mo)" : "Total Base / 1 people, USD/mo",
        "Net Save/mo"            : SAV_METRIC,
        "Rent 1BR centre"        : "1 bed. ap. in city centre",
        "Quality of Life"        : "Quality of life",
        "Safety"                 : "Safety",
        "Crime Index"            : "Crime index",
        "Healthcare"             : "Health care",
        "Climate"                : "Climate",
        "Purchasing Power"       : "Purch, power",
        "Internet (Mbps)"        : "Internet Speed (Mbps)",
        "Pollution"              : "Pollution",
    }
    q_df = pd.DataFrame({lbl: row(m, selected) for lbl, m in q_rows.items()}, index=selected).T
    st.dataframe(q_df.round(1).style.background_gradient(cmap="RdYlGn", axis=1),
                 use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# 1 · FINANCES
# ══════════════════════════════════════════════════════════════════════════════
with TABS[1]:
    st.subheader("💰 Income, Expenses & Savings")

    sal_avg = row("Avg Salary, USD/mo",     selected)
    sal_dev = row("Avg Salary Dev, USD/mo", selected)
    exp_b1  = row("Total Base / 1 people, USD/mo", selected)
    exp_a1  = row("Total All / 1 people, USD/mo",  selected)
    exp_b2  = row("Total Base / 2 people, USD/mo", selected)
    exp_a2  = row("Total All / 2 people, USD/mo",  selected)

    # salary vs expenses
    fig_fe = go.Figure()
    fig_fe.add_bar(x=selected, y=[v(sal_avg,c) for c in selected], name="Avg Salary",        marker_color="#90caf9")
    fig_fe.add_bar(x=selected, y=[v(sal_dev,c) for c in selected], name="Dev Salary",        marker_color="#1565c0")
    fig_fe.add_bar(x=selected, y=[v(exp_b1, c) for c in selected], name="Base Exp (solo)",   marker_color="#ef9a9a")
    fig_fe.add_bar(x=selected, y=[v(exp_b2, c) for c in selected], name="Base Exp (couple)", marker_color="#b71c1c")
    fig_fe.update_layout(barmode="group", yaxis_tickformat="$,.0f",
                         title="Salary vs Base Monthly Expenses (USD/mo)")
    chart(fig_fe, 460)

    # ── All 8 savings scenarios from CSV
    st.markdown("### 💧 Net Monthly Savings — all scenarios")
    SAV_SCENARIOS: Dict[str, str] = {
        "1·Base·Avg"     : "Saving base 1 people on Avg salary, USD/mo ",
        "1·Base·Dev"     : "Saving base 1 people on Avg Dev salary, USD/mo ",
        "1·All·Avg"      : "Saving all 1 people on Avg salary, USD/mo ",
        "1·All·Dev"      : "Saving all 1 people on Avg Dev salary, USD/mo ",
        "2·Base·Dev"     : "Saving base 2 people on Avg Dev salary, USD/mo ",
        "2·Base·Dev+Des" : "Saving base 2 people on Avg Dev + Des salary, USD/mo ",
        "2·All·Dev"      : "Saving all 2 people on Avg Dev salary, USD/mo ",
        "2·All·Dev+Des"  : "Saving all 2 people on Avg Dev + Des salary, USD/mo ",
    }
    sav_long = (
        pd.DataFrame({lbl: row(m, selected) for lbl, m in SAV_SCENARIOS.items()}, index=selected)
        .reset_index(names="City")
        .melt(id_vars="City", var_name="Scenario", value_name="USD/mo")
    )
    fig_sav = px.bar(sav_long, x="City", y="USD/mo", color="Scenario",
                     barmode="group",
                     category_orders={"Scenario": list(SAV_SCENARIOS.keys())},
                     title="Net monthly savings — all 8 CSV scenarios",
                     color_discrete_sequence=px.colors.qualitative.Set2)
    fig_sav.add_hline(y=0, line_dash="dash", line_color="red", opacity=.6)
    fig_sav.update_layout(
        yaxis_tickformat="$,.0f",
        legend=dict(orientation="h", yanchor="top", y=-0.30,
                    xanchor="left", x=0, font=dict(size=11)),
        margin=dict(l=4, r=4, t=44, b=140),
        bargap=0.10, bargroupgap=0.05,
        xaxis=dict(tickangle=-35, automargin=True),
    )
    st.plotly_chart(fig_sav, use_container_width=True, config=PLOTLY_CFG)
    st.caption(f"Active profile: **{next((k for k,m in SAV_SCENARIOS.items() if m==SAV_METRIC), SAV_METRIC.strip())}**")

    # ── Purchasing power + tax
    st.markdown("### 💪 Purchasing Power & Tax Rates")
    col1, col2 = st.columns(2)
    with col1:
        pp = row("Purch, power", selected)
        fig_pp = px.bar(x=selected, y=[v(pp,c) for c in selected],
                        title="Purchasing Power Index (New York=100)",
                        color=[v(pp,c) for c in selected], color_continuous_scale="Blues",
                        labels={"x":"City","y":"Index"})
        fig_pp.update_layout(showlegend=False)
        chart(fig_pp, 380)
    with col2:
        tax_sub = df[df["Metric"].isin(["salary tax, %","VAT, %"])][["Metric"]+selected]
        if not tax_sub.empty:
            tax_long = tax_sub.melt(id_vars="Metric", var_name="City", value_name="Rate")
            tax_long["Rate"] = pd.to_numeric(tax_long["Rate"], errors="coerce")
            fig_tax = px.bar(tax_long, x="City", y="Rate", color="Metric", barmode="group",
                             title="Salary Tax & VAT (%)",
                             color_discrete_map={"salary tax, %":"#ef5350","VAT, %":"#ffa726"})
            chart(fig_tax, 380)

    # ── Full financial table
    st.markdown("### 📋 Financial Summary")
    fin = pd.DataFrame({
        "Avg Salary"       : sal_avg,
        "Dev Salary"       : sal_dev,
        "Base Exp (1p)"    : exp_b1,
        "All Exp (1p)"     : exp_a1,
        "Base Exp (2p)"    : exp_b2,
        "All Exp (2p)"     : exp_a2,
        "Net Save (profile)": row(SAV_METRIC, selected),
        "Purch. Power"     : row("Purch, power", selected),
    }, index=selected)
    mcols = [c for c in fin.columns if c != "Purch. Power"]
    st.dataframe(
        fin.round(0)
           .style
           .format("${:,.0f}", subset=list(mcols))
           .format("{:.0f}",   subset=["Purch. Power"])
           .background_gradient(cmap="RdYlGn", axis=0),
        use_container_width=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
# 2 · HOUSING
# ══════════════════════════════════════════════════════════════════════════════
with TABS[2]:
    st.subheader("🏠 Housing Market")

    r1c   = row("1 bed. ap. in city centre",   selected)
    r1o   = row("1 bed. ap. out. city centre", selected)
    r3c   = row("3 bed. ap. in city centre",   selected)
    r3o   = row("3 bed. ap. out. city centre", selected)
    buy_c = row("Buy Apartment Price in centre, 80m*m / USD",   selected)
    buy_o = row("Buy Apartment Price out. centre, 80m*m / USD", selected)
    sqm_c = row("USD/m*m in city centre",      selected)
    sqm_o = row("USD/m*m out. city centre",    selected)
    rrate = row("Interest Rate (20-Year Fixed, in %)", selected)
    pti   = row("Prop, to income", selected)

    col1, col2 = st.columns(2)
    with col1:
        fg = go.Figure()
        fg.add_bar(x=selected, y=[v(r1c,c) for c in selected], name="1BR Centre",  marker_color="#42a5f5")
        fg.add_bar(x=selected, y=[v(r1o,c) for c in selected], name="1BR Suburbs", marker_color="#90caf9")
        fg.add_bar(x=selected, y=[v(r3c,c) for c in selected], name="3BR Centre",  marker_color="#1565c0")
        fg.add_bar(x=selected, y=[v(r3o,c) for c in selected], name="3BR Suburbs", marker_color="#4fc3f7")
        fg.update_layout(barmode="group", yaxis_tickformat="$,.0f", title="Monthly Rent (USD/mo)")
        chart(fg, 420)
    with col2:
        fg2 = go.Figure()
        fg2.add_bar(x=selected, y=[v(buy_c,c) for c in selected], name="80m² Centre",  marker_color="#ef5350")
        fg2.add_bar(x=selected, y=[v(buy_o,c) for c in selected], name="80m² Suburbs", marker_color="#ffcdd2")
        fg2.update_layout(barmode="group", yaxis_tickformat="$,.0f", title="Purchase Price — 80 m²")
        chart(fg2, 420)

    # price per m²
    st.markdown("### 📐 Price per m²")
    fg3 = go.Figure()
    fg3.add_bar(x=selected, y=[v(sqm_c,c) for c in selected], name="$/m² Centre",  marker_color="#ab47bc")
    fg3.add_bar(x=selected, y=[v(sqm_o,c) for c in selected], name="$/m² Suburbs", marker_color="#e1bee7")
    fg3.update_layout(barmode="group", yaxis_tickformat="$,.0f", title="Price per m²")
    chart(fg3, 380)

    # mortgage calculator
    st.markdown("### 🏦 Mortgage Calculator (80 m², 20 % down, 20-year fixed)")
    mort = []
    sal_ser = row(SAL_METRIC, selected)
    for city in selected:
        price = v(buy_c, city)
        r_ann = v(rrate, city)
        if any(np.isnan(x) for x in [price, r_ann]):
            continue
        loan = price * .80
        r_mo = r_ann / 100 / 12
        n    = 240
        pmt  = loan * r_mo * (1+r_mo)**n / ((1+r_mo)**n - 1) if r_mo > 0 else loan/n
        sl   = v(sal_ser, city)
        mort.append({
            "City"             : city,
            "Price 80m²"       : price,
            "Mortgage Rate %"  : r_ann,
            "Monthly Payment"  : pmt,
            "% of Dev Salary"  : pmt/sl*100 if sl and sl > 0 else np.nan,
            "Years (full price)": price/(sl*12) if sl and sl > 0 else np.nan,
        })
    if mort:
        mdf = pd.DataFrame(mort).set_index("City")
        st.dataframe(
            mdf.style
               .format({"Price 80m²":"${:,.0f}","Monthly Payment":"${:,.0f}",
                        "Mortgage Rate %":"{:.2f}%","% of Dev Salary":"{:.1f}%",
                        "Years (full price)":"{:.1f}"})
               .background_gradient(cmap="RdYlGn_r", subset=["% of Dev Salary","Years (full price)"]),
            use_container_width=True,
        )

        fig_mort = px.bar(
            mdf.reset_index(), x="City", y="Monthly Payment",
            title="Mortgage Monthly Payment (USD)",
            color="% of Dev Salary", color_continuous_scale="RdYlGn_r",
            labels={"Monthly Payment":"$/mo"},
        )
        chart(fig_mort, 360)

    # property-to-income ratio
    st.markdown("### 📊 Property-to-Income Ratio (years of salary)")
    fig_pti = px.bar(
        x=selected, y=[v(pti,c) for c in selected],
        title="Years of Salary to Buy 80m² (lower = more affordable)",
        color=[v(pti,c) for c in selected], color_continuous_scale="RdYlGn_r",
        labels={"x":"City","y":"Years"},
    )
    fig_pti.update_layout(showlegend=False)
    chart(fig_pti, 380)


# ══════════════════════════════════════════════════════════════════════════════
# 3 · COST OF LIVING
# ══════════════════════════════════════════════════════════════════════════════
with TABS[3]:
    st.subheader("🛒 Cost of Living")

    # stacked spend breakdown
    SPEND_CATS: Dict[str, str] = {
        "Rent 1BR"       : "1 bed. ap. in city centre",
        "Groceries"      : "Groceries, USD/wk * 4,3",
        "Transport"      : "Transport, USD/mo",
        "Restaurants"    : "Restaurant, USD/mo",
        "Utilities"      : "Utilites, USD/mo",
        "Sports/Leisure" : "Sports and Leisure",
        "Clothing"       : "Clothing and Shoes, 2 pairs/year / 12",
    }
    spend_df = pd.DataFrame({cat: row(m, selected) for cat, m in SPEND_CATS.items()}, index=selected)
    fig_st = px.bar(spend_df, barmode="stack",
                    title="Monthly Spending Breakdown by Category",
                    color_discrete_sequence=px.colors.qualitative.Set3)
    fig_st.update_layout(yaxis_tickformat="$,.0f")
    chart(fig_st, 520)

    # ── Salary coverage: how many months of each category does salary buy?
    st.markdown("### 🔬 Spend as % of Dev Salary")
    sal_ser2 = row(SAL_METRIC, selected)
    pct_data = {}
    for cat, m in SPEND_CATS.items():
        cat_row = row(m, selected)
        pct_data[cat] = pd.Series({
            c: v(cat_row,c)/v(sal_ser2,c)*100
            if pd.notna(v(cat_row,c)) and v(sal_ser2,c) and v(sal_ser2,c)>0
            else np.nan
            for c in selected
        })
    pct_df = pd.DataFrame(pct_data, index=selected)
    fig_pct = px.bar(pct_df, barmode="stack",
                     title="Each spending category as % of Dev salary",
                     color_discrete_sequence=px.colors.qualitative.Set3)
    fig_pct.add_hline(y=100, line_dash="dash", line_color="red", opacity=.5,
                      annotation_text="100% of salary")
    fig_pct.update_layout(yaxis_title="% of salary")
    chart(fig_pct, 460)

    # ── Grocery heatmap
    st.markdown("### 🧺 Grocery Basket Prices (USD)")
    GROCERIES: Dict[str, str] = {
        "Milk 1L"     : "milk, l",
        "Bread 0.5kg" : "white bread,  0,5kg",
        "Eggs 12"     : "eggs, 12",
        "Cheese kg"   : "cheese, kg",
        "Chicken kg"  : "chicken filltets, kg",
        "Beef kg"     : "beef, kg",
        "Apples kg"   : "apples, kg",
        "Potatoes kg" : "potatoes, kg",
        "Tomatoes kg" : "tomatoes, kg",
        "Bananas kg"  : "bananas, kg",
        "Rice kg"     : "rice, kg",
        "Water 1.5L"  : "bottled of watter, 1,5l",
        "Cappuccino"  : "cappuccino",
        "McCombo"     : "combo at mcdonald's",
        "Inexp. meal" : "meal at an inexp.rest.",
        "Dinner for 2": "meal for two at a mid rest.",
    }
    groc_df = pd.DataFrame({k: row(m, selected) for k, m in GROCERIES.items()}, index=selected).T
    fig_gr = px.imshow(groc_df, color_continuous_scale="RdYlGn_r",
                       title="Grocery & Restaurant Prices — USD", aspect="auto")
    chart(fig_gr, 520)

    # ── Transport + cars
    st.markdown("### 🚌 Transport & Cars")
    col1, col2 = st.columns(2)
    with col1:
        pass_r   = row("public transport pass, mo", selected)
        oneway_r = row("one way ticket", selected)
        fig_tr = go.Figure()
        fig_tr.add_bar(x=selected, y=[v(pass_r,c)   for c in selected], name="Monthly Pass", marker_color="#42a5f5")
        fig_tr.add_bar(x=selected, y=[v(oneway_r,c) for c in selected], name="One-way",      marker_color="#90caf9")
        fig_tr.update_layout(barmode="group", yaxis_tickformat="$,.2f",
                             title="Public Transport (USD/mo)")
        chart(fig_tr, 360)
    with col2:
        golf_r  = row("VF Golf 1,5", selected)
        toyo_r  = row("Toyoyta Corolla Sedan 1,6", selected)
        gas_r   = row("gasoline, l", selected)
        fig_ca = go.Figure()
        fig_ca.add_bar(x=selected, y=[v(golf_r, c) for c in selected], name="VW Golf 1.5",    marker_color="#26a69a")
        fig_ca.add_bar(x=selected, y=[v(toyo_r, c) for c in selected], name="Toyota Corolla", marker_color="#7e57c2")
        fig_ca.update_layout(barmode="group", yaxis_tickformat="$,.0f", title="New Car Prices (USD)")
        chart(fig_ca, 360)

    # ── Numbeo indices
    st.markdown("### 📊 Numbeo Indices (New York = 100)")
    idx_data = {
        "Cost of Living" : row("Cost of living",   selected),
        "Rent Index"     : row("Rent Index",        selected),
        "Groceries Index": row("Groceries Index",   selected),
    }
    idx_df = pd.DataFrame(idx_data, index=selected)
    fig_idx = px.bar(idx_df, barmode="group",
                     color_discrete_sequence=["#26a69a","#42a5f5","#7e57c2"])
    fig_idx.update_layout(title="Numbeo Indices vs New York (100)")
    chart(fig_idx, 380)

    # ── Full spending table
    st.markdown("### 📋 Complete Spending Table")
    ALL_SPEND = {
        **SPEND_CATS,
        "Monthly Pass"   : "public transport pass, mo",
        "Gasoline/L"     : "gasoline, l",
        "Taxi/km"        : "taxi, 1km",
        "Fitness Club/mo": "fitness club, mo",
        "Cinema Ticket"  : "cinema ticket",
        "Cappuccino"     : "cappuccino",
        "McCombo"        : "combo at mcdonald's",
        "Dinner for 2"   : "meal for two at a mid rest.",
        "Internet/mo"    : "Internet (unlim,  >= 60Mpbs )",
        "Mobile/mo"      : "mobile phone (calls + 10GB+ Data)",
    }
    full_df = pd.DataFrame({k: row(m, selected) for k, m in ALL_SPEND.items()}, index=selected)
    st.dataframe(full_df.round(2).style.background_gradient(cmap="RdYlGn_r", axis=0),
                 use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# 4 · QUALITY OF LIFE
# ══════════════════════════════════════════════════════════════════════════════
with TABS[4]:
    st.subheader("🌿 Quality of Life & Environment")

    QOL_ITEMS: Dict[str, str] = {
        "Quality of Life" : "Quality of life",
        "Safety"          : "Safety",
        "Crime Index"     : "Crime index",
        "Healthcare"      : "Health care",
        "Purch. Power"    : "Purch, power",
        "Pollution"       : "Pollution",
        "Climate"         : "Climate",
        "Traffic"         : "Traffic",
        "Rent Index"      : "Rent Index",
    }
    qol_df = pd.DataFrame({k: row(m, selected) for k, m in QOL_ITEMS.items()}, index=selected)
    fig_qhm = px.imshow(qol_df.T, color_continuous_scale="RdYlGn",
                        title="Quality of Life Heatmap", aspect="auto")
    chart(fig_qhm, 440)

    col1, col2 = st.columns(2)
    with col1:
        cr = row("Crime index", selected)
        fig_cr = px.bar(x=selected, y=[v(cr,c) for c in selected],
                        title="Crime Index (lower = safer)",
                        color=[v(cr,c) for c in selected], color_continuous_scale="RdYlGn_r")
        fig_cr.update_layout(showlegend=False)
        chart(fig_cr, 360)
    with col2:
        hc = row("Health care", selected)
        fig_hc = px.bar(x=selected, y=[v(hc,c) for c in selected],
                        title="Healthcare Index (higher = better)",
                        color=[v(hc,c) for c in selected], color_continuous_scale="RdYlGn")
        fig_hc.update_layout(showlegend=False)
        chart(fig_hc, 360)

    # weather
    st.markdown("### ☀️ Climate & Weather")
    col1, col2 = st.columns(2)
    with col1:
        sun_r  = row("Sunny days, %",  selected)
        rain_r = row("Rainy days, %",  selected)
        summ_r = row("Avg Summer temp", selected)
        wint_r = row("Avg Winter temp", selected)
        fg_wr = go.Figure()
        fg_wr.add_bar(x=selected, y=[v(sun_r,c)  for c in selected], name="Sunny %", marker_color="#ffd54f")
        fg_wr.add_bar(x=selected, y=[v(rain_r,c) for c in selected], name="Rainy %", marker_color="#64b5f6")
        fg_wr.update_layout(barmode="group", title="Sunny vs Rainy Days (%)")
        chart(fg_wr, 380)
    with col2:
        fg_tp = go.Figure()
        fg_tp.add_bar(x=selected, y=[v(summ_r,c) for c in selected], name="Summer °C", marker_color="#ff7043")
        fg_tp.add_bar(x=selected, y=[v(wint_r,c) for c in selected], name="Winter °C", marker_color="#42a5f5")
        fg_tp.update_layout(barmode="group", title="Avg Temperatures (°C)")
        chart(fg_tp, 380)

    col1, col2 = st.columns(2)
    with col1:
        pol = row("Pollution", selected)
        fig_pol = px.bar(x=selected, y=[v(pol,c) for c in selected],
                         title="Pollution Index (lower = cleaner)",
                         color=[v(pol,c) for c in selected], color_continuous_scale="RdYlGn_r")
        fig_pol.update_layout(showlegend=False)
        chart(fig_pol, 360)
    with col2:
        inet = row("Internet Speed (Mbps)", selected)
        fig_inet = px.bar(x=selected, y=[v(inet,c) for c in selected],
                          title="Internet Speed (Mbps)",
                          color=[v(inet,c) for c in selected], color_continuous_scale="Blues")
        fig_inet.update_layout(showlegend=False)
        chart(fig_inet, 360)

    # nature + expats
    st.markdown("### 🌲 Nature Access & Community")
    col1, col2 = st.columns(2)
    with col1:
        nat = pd.DataFrame({
            "Ocean/Sea h": row("Ocean/Sea, h",  selected),
            "Forest h"   : row("Forest, h",     selected),
            "Mountains h": row("Mountains, h",  selected),
        }, index=selected)
        fig_nat = px.bar(nat, barmode="group", title="Hours to Nature (by car)",
                         color_discrete_sequence=["#26a69a","#66bb6a","#8d6e63"])
        fig_nat.update_layout(yaxis_title="Hours")
        chart(fig_nat, 380)
    with col2:
        imm_p = row("Immigrants, %", selected)
        uv_r  = row("Avg Peak UV index", selected)
        fg_iv = go.Figure()
        fg_iv.add_bar(x=selected, y=[v(imm_p,c) for c in selected], name="Immigrants %", marker_color="#7986cb")
        fg_iv.add_bar(x=selected, y=[v(uv_r, c) for c in selected], name="Peak UV",       marker_color="#ffb74d")
        fg_iv.update_layout(barmode="group", title="Expat Share & UV Index")
        chart(fg_iv, 380)


# ══════════════════════════════════════════════════════════════════════════════
# 5 · IMMIGRATION
# ══════════════════════════════════════════════════════════════════════════════
with TABS[5]:
    st.subheader("🛂 Immigration & Residency Pathways")

    IMM_ITEMS: Dict[str, str] = {
        "PR Fast Track (yrs)"    : "PR (Years) Fast Track",
        "PR General (yrs)"       : "PR (Years) General Path",
        "Citizenship Fast (yrs)" : "Citizenship (Years) Fast",
        "Citizenship Real (yrs)" : "Citizenship (Years) Real",
        "Passport Difficulty"    : "BY Passport Difficulty",
        "IELTS Required"         : "IELTS / Language",
        "1-yr Buffer ($k)"       : "1 Year Buffer (USD), $k",
        "Salary 4-5y ($k)"       : "Salary (4-5y exp), $k",
    }
    imm_df = pd.DataFrame({k: row(m, selected) for k, m in IMM_ITEMS.items()}, index=selected)

    fig_ihm = px.imshow(imm_df.T, title="Immigration Metrics Heatmap",
                        color_continuous_scale="RdYlGn_r", aspect="auto")
    chart(fig_ihm, 420)

    st.markdown("### ⏱️ Residency Timeline (years)")
    fig_tl = go.Figure()
    for lbl, m, clr in [
        ("PR Fast Track",    "PR (Years) Fast Track",    "#66bb6a"),
        ("PR General",       "PR (Years) General Path",  "#ffa726"),
        ("Citizenship Fast", "Citizenship (Years) Fast", "#42a5f5"),
        ("Citizenship Real", "Citizenship (Years) Real", "#ef5350"),
    ]:
        rr = row(m, selected)
        fig_tl.add_bar(x=selected, y=[v(rr,c) for c in selected], name=lbl, marker_color=clr)
    fig_tl.update_layout(barmode="group", title="Years to PR / Citizenship")
    chart(fig_tl, 440)

    # buffer vs savings
    st.markdown("### 🛡️ Financial Buffer: Required vs Your Savings")
    buf_r = row("1 Year Buffer (USD), $k", selected) * 1000
    fg_bf = go.Figure()
    fg_bf.add_bar(x=selected, y=[v(buf_r,c)     for c in selected],
                  name="Required Buffer",            marker_color="#ef5350")
    fg_bf.add_bar(x=selected, y=[FINAL_SAVED]*len(selected),
                  name=f"Your Savings by {target_year}", marker_color="#66bb6a")
    fg_bf.update_layout(barmode="group", yaxis_tickformat="$,.0f",
                        title="1-Year Buffer Required vs Projected Savings")
    chart(fg_bf, 420)

    # coverage gauge per city
    st.markdown("### 📊 Buffer Coverage %")
    cov_vals = [
        FINAL_SAVED / v(buf_r, c) * 100
        if pd.notna(v(buf_r, c)) and v(buf_r, c) > 0 else 0
        for c in selected
    ]
    fig_cg = px.bar(x=selected, y=cov_vals,
                    title="Your savings as % of required 1-year buffer",
                    color=cov_vals, color_continuous_scale="RdYlGn")
    fig_cg.add_hline(y=100, line_dash="dash", line_color="white",
                     annotation_text="100% goal")
    fig_cg.update_layout(showlegend=False, yaxis_title="%")
    chart(fig_cg, 360)

    st.markdown("### 📋 Immigration Summary")
    st.dataframe(
        imm_df.style.background_gradient(
            cmap="RdYlGn_r",
            subset=["PR Fast Track (yrs)","PR General (yrs)",
                    "Citizenship Fast (yrs)","Citizenship Real (yrs)","Passport Difficulty"]),
        use_container_width=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
# 6 · SAVINGS PLAN
# ══════════════════════════════════════════════════════════════════════════════
with TABS[6]:
    st.subheader("📈 Personal Savings & Relocation Plan")
    st.info(
        f"**Saving:** ${monthly_save:,.0f}/mo  ·  growing {growth}%/yr  ·  "
        f"{prep_years} years ({start_year} → {target_year})  ·  **Target: ${FINAL_SAVED:,.0f}**"
    )

    fig_acc = px.area(x=_yr_range, y=_accum,
                      title="Accumulated Savings Over Time",
                      labels={"x":"Year","y":"Total Saved ($)"},
                      color_discrete_sequence=["#42a5f5"])
    fig_acc.update_layout(yaxis_tickformat="$,.0f")
    chart(fig_acc, 420)

    # ── Coverage & break-even table
    st.markdown(f"### 🎯 Savings by {target_year}: **${FINAL_SAVED:,.0f}**")
    buf_r  = row("1 Year Buffer (USD), $k", selected) * 1000
    exp_r  = row(EXP_METRIC, selected)
    sav_r  = row(SAV_METRIC, selected)
    sal_r2 = row(SAL_METRIC, selected)

    comp = []
    for city in selected:
        bv   = v(buf_r, city)
        ev   = v(exp_r, city)
        sv   = v(sal_r2, city)
        nmo  = v(sav_r, city)
        if np.isnan(nmo) and pd.notna(sv) and pd.notna(ev):
            nmo = sv - ev
        cov   = FINAL_SAVED / bv * 100 if bv and bv > 0 else np.nan
        months= FINAL_SAVED / ev       if ev and ev > 0 else np.nan
        beven = (bv - FINAL_SAVED) / nmo if (
            pd.notna(nmo) and nmo > 0 and pd.notna(bv) and bv > FINAL_SAVED
        ) else 0.0
        # 10-year wealth = savings + 10y * 12 * net_mo
        wealth10 = FINAL_SAVED + nmo * 120 if pd.notna(nmo) else np.nan
        comp.append({
            "City"           : city,
            "Buffer Required": bv,
            "Your Savings"   : FINAL_SAVED,
            "Coverage %"     : cov,
            "Months of Cover": months,
            "Net Save/mo"    : nmo,
            "Break-even (mo)": max(float(beven), 0),
            "Wealth +10yr"   : wealth10,
        })
    comp_df = pd.DataFrame(comp).set_index("City")
    st.dataframe(
        comp_df.style
               .format({"Buffer Required":"${:,.0f}","Your Savings":"${:,.0f}",
                        "Coverage %":"{:.1f}%","Months of Cover":"{:.1f}",
                        "Net Save/mo":"${:,.0f}","Break-even (mo)":"{:.1f}",
                        "Wealth +10yr":"${:,.0f}"})
               .background_gradient(cmap="RdYlGn",   subset=["Coverage %","Months of Cover"])
               .background_gradient(cmap="RdYlGn_r", subset=["Break-even (mo)"]),
        use_container_width=True,
    )

    fig_cov2 = px.bar(comp_df.reset_index(), x="City", y="Coverage %",
                      title=f"Buffer Coverage — ${FINAL_SAVED:,.0f} vs Required",
                      color="Coverage %", color_continuous_scale="RdYlGn")
    fig_cov2.add_hline(y=100, line_dash="dash", line_color="white", annotation_text="100% target")
    fig_cov2.update_layout(showlegend=False)
    chart(fig_cov2, 400)

    # ── Wealth projection
    st.markdown("### 📊 Wealth Projection After Relocation (10 years)")
    proj_yrs = list(range(target_year, target_year + 11))
    fig_proj = go.Figure()
    for i, city in enumerate(selected):
        nmo = v(sav_r, city)
        sv  = v(sal_r2, city)
        ev  = v(exp_r, city)
        if np.isnan(nmo) and pd.notna(sv) and pd.notna(ev):
            nmo = sv - ev
        if np.isnan(nmo):
            continue
        wealth = [FINAL_SAVED + nmo * 12 * yr for yr in range(len(proj_yrs))]
        fig_proj.add_scatter(x=proj_yrs, y=wealth, mode="lines+markers",
                             name=city, line_color=COLORS[i % len(COLORS)])
    fig_proj.update_layout(yaxis_tickformat="$,.0f",
                           title="Cumulative Wealth: savings + net income on-site (10 yr)")
    chart(fig_proj, 460)


# ══════════════════════════════════════════════════════════════════════════════
# 7 · FINAL SCORE
# ══════════════════════════════════════════════════════════════════════════════
with TABS[7]:
    st.subheader("🏆 Composite Relocation Score")
    st.markdown("""
    **Weighted composite (0–100)** built entirely from your spreadsheet data:

    | Dimension | Weight | Key inputs |
    |---|---|---|
    | Financial | **30 %** | Net savings rate, purchasing power |
    | Quality of Life | **25 %** | QoL index, safety, healthcare, climate |
    | Housing Affordability | **20 %** | Rent/salary ratio, property-to-income |
    | Immigration Ease | **15 %** | PR fast-track years, passport difficulty, buffer coverage |
    | Environment | **10 %** | Pollution, internet speed, expat community % |
    """)

    score_rows = []
    for city in selected:
        def g(m: str) -> float:
            val = v(row(m, selected), city)
            return float(val) if pd.notna(val) else 0.0

        # ── financial
        sv = g(SAL_METRIC); ev = g(EXP_METRIC)
        sav_v = g(SAV_METRIC)
        if sav_v and sv > 0:
            save_r = sav_v / sv * 100
        elif sv > 0 and ev > 0:
            save_r = (sv - ev) / sv * 100
        else:
            save_r = 0.0
        fin = np.clip(save_r * 1.2, 0, 60) + np.clip((g("Purch, power") - 50) / 150 * 40, 0, 40)

        # ── quality of life
        qol = (np.clip(g("Quality of life") / 220 * 100, 0, 100) * .40 +
               np.clip(g("Safety"),       0, 100)                 * .25 +
               np.clip(g("Health care"),  0, 100)                 * .20 +
               np.clip(g("Climate"),      0, 100)                 * .15)

        # ── housing
        r1c_v = g("1 bed. ap. in city centre")
        rr = max((1 - r1c_v / sv) * 100, 0) if sv > 0 else 50
        hsg = (np.clip(rr,                                0, 100) * .50 +
               np.clip((20 - g("Prop, to income"))/17*100, 0, 100) * .50)

        # ── immigration
        buf_v  = g("1 Year Buffer (USD), $k") * 1000
        buf_c  = np.clip(FINAL_SAVED / buf_v * 100, 0, 100) if buf_v > 0 else 0
        imm = (np.clip((10 - g("PR (Years) Fast Track")) / 10 * 100, 0, 100) * .50 +
               np.clip(g("BY Passport Difficulty") / 5 * 100,          0, 100) * .30 +
               buf_c                                                             * .20)

        # ── environment
        env = (np.clip(100 - g("Pollution"),             0, 100) * .40 +
               np.clip(g("Internet Speed (Mbps)") / 350 * 100, 0, 100) * .30 +
               np.clip(g("Immigrants, %"),              0, 100) * .30)

        composite = fin*.30 + qol*.25 + hsg*.20 + imm*.15 + env*.10
        score_rows.append({
            "City"           : city,
            "Financial"      : round(fin, 1),
            "Quality of Life": round(qol, 1),
            "Housing"        : round(hsg, 1),
            "Immigration"    : round(imm, 1),
            "Environment"    : round(env, 1),
            "⭐ Score"       : round(composite, 1),
        })

    score_df = (pd.DataFrame(score_rows)
                  .set_index("City")
                  .sort_values("⭐ Score", ascending=False))

    if len(score_df) > 0:
        top = score_df.index[0]
        st.success(
            f"### 🥇 Best match for your profile: **{top}**  "
            f"— Score: **{score_df.loc[top,'⭐ Score']:.1f} / 100**"
        )

    # horizontal bar
    fig_sc = px.bar(
        score_df.reset_index().sort_values("⭐ Score"),
        x="⭐ Score", y="City", orientation="h",
        color="⭐ Score", color_continuous_scale="RdYlGn",
        title="Composite Relocation Score (0–100)",
    )
    fig_sc.update_layout(showlegend=False, height=max(300, len(selected)*55))
    st.plotly_chart(fig_sc, use_container_width=True, config=PLOTLY_CFG)

    # stacked dimension contribution
    DIM_COLS = ["Financial","Quality of Life","Housing","Immigration","Environment"]
    DIM_W    = [.30, .25, .20, .15, .10]
    wt = score_df[DIM_COLS].copy()
    for c, w in zip(DIM_COLS, DIM_W):
        wt[c] = wt[c] * w
    fig_dim = px.bar(wt.reset_index(), x="City", y=DIM_COLS, barmode="stack",
                     title="Score Contribution by Dimension (weighted points)",
                     color_discrete_sequence=px.colors.qualitative.Set2)
    fig_dim.update_layout(yaxis_title="Weighted Points")
    chart(fig_dim, 460)

    # full table
    st.markdown("### 📋 Score Details")
    st.dataframe(
        score_df.style.background_gradient(cmap="RdYlGn", vmin=0, vmax=100),
        use_container_width=True,
    )

    # city profile cards
    st.markdown("---")
    st.markdown("### 💡 City Profiles")
    sal_ser3  = row(SAL_METRIC, selected)
    exp_ser3  = row(EXP_METRIC, selected)
    sav_ser3  = row(SAV_METRIC, selected)
    pr_ser3   = row("PR (Years) Fast Track", selected)
    city_list: List[str] = [str(c) for c in score_df.index]
    for city_n, rs in score_df.iterrows():
        city_s: str = str(city_n)          # cast Hashable → str
        rank   = city_list.index(city_s) + 1
        medal  = "🥇" if rank == 1 else "🥈" if rank == 2 else "🥉" if rank == 3 else "🏅"
        sv_    = v(sal_ser3, city_s)
        ev_    = v(exp_ser3, city_s)
        nmo_   = v(sav_ser3, city_s)
        if np.isnan(nmo_) and pd.notna(sv_) and pd.notna(ev_):
            nmo_ = sv_ - ev_
        pr_f   = v(pr_ser3, city_s)
        base_cap = (
            f"Salary: {fmt_usd(sv_)}/mo  ·  Expenses: {fmt_usd(ev_)}/mo  ·  "
            f"Net save: {fmt_usd(nmo_)}/mo"
        )
        caption = base_cap + (f"  ·  PR fast-track: {pr_f:.0f} yrs" if pd.notna(pr_f) else "")
        with st.expander(f"{medal} #{rank}  **{city_s}**  —  {rs['⭐ Score']:.1f} / 100"):
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Financial",       f"{rs['Financial']:.1f}")
            c2.metric("Quality of Life", f"{rs['Quality of Life']:.1f}")
            c3.metric("Housing",         f"{rs['Housing']:.1f}")
            c4.metric("Immigration",     f"{rs['Immigration']:.1f}")
            c5.metric("Environment",     f"{rs['Environment']:.1f}")
            st.caption(caption)