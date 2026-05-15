"""
Relocation Intelligence 2026
─────────────────────────────
Streamlit dashboard — 31 cities, 80+ metrics.

requirements.txt (place next to app.py):
    streamlit>=1.35.0
    pandas>=2.0.0
    numpy>=1.24.0
    plotly>=5.20.0
    openpyxl>=3.1.0
    matplotlib>=3.7.0
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

CITY_TO_COUNTRY: Dict[str, str] = {
    "Auckland": "New Zealand 🇳🇿",    "Wellington": "New Zealand 🇳🇿",
    "Christchurch": "New Zealand 🇳🇿",
    "Melbourne": "Australia 🇦🇺",     "Sydney": "Australia 🇦🇺",
    "Brisbane": "Australia 🇦🇺",      "Perth": "Australia 🇦🇺",
    "Adelaide": "Australia 🇦🇺",
    "Toronto": "Canada 🇨🇦",          "Montreal": "Canada 🇨🇦",
    "Ottawa": "Canada 🇨🇦",
    "Chicago": "USA 🇺🇸",             "Philadelphia": "USA 🇺🇸",
    "San Diego": "USA 🇺🇸",           "Dallas": "USA 🇺🇸",
    "Austin": "USA 🇺🇸",              "Seattle": "USA 🇺🇸",
    "Birmingham": "United Kingdom 🇬🇧", "Liverpool": "United Kingdom 🇬🇧",
    "Southhampton": "United Kingdom 🇬🇧", "Manchester": "United Kingdom 🇬🇧",
    "Amsterdam": "Netherlands 🇳🇱",   "Rotterdam": "Netherlands 🇳🇱",
    "Oslo": "Norway 🇳🇴",
    "Belgrade": "Serbia 🇷🇸",
    "Warsaw": "Poland 🇵🇱",           "Krakow": "Poland 🇵🇱",
    "Lodz": "Poland 🇵🇱",             "Gdansk": "Poland 🇵🇱",
    "Moscow": "Russia 🇷🇺",
    "Minsk": "Belarus",
}

PLOTLY_CFG: Dict[str, Any] = {
    "displayModeBar": True, "displaylogo": False,
    "responsive": True, "scrollZoom": False,
}
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
[data-testid="stMetricValue"]  { font-size:1.2rem; font-weight:700; }
[data-testid="stMetricLabel"]  { font-size:.75rem; color:#aaa; }
[data-testid="stMetricDelta"]  { font-size:.78rem; }
.block-container {
    padding-top:.5rem !important;
    padding-left:max(.5rem,2vw) !important;
    padding-right:max(.5rem,2vw) !important;
    max-width:1440px;
}
[data-testid="stExpander"] { border:1px solid #2a3a4a !important; border-radius:8px !important; }
[data-testid="stTabs"] [role="tablist"] { flex-wrap:wrap !important; gap:2px !important; }
[data-testid="stTabs"] button { min-width:0 !important; white-space:normal !important; text-align:center !important; }
@media(max-width:640px){
    [data-testid="stMetricValue"] { font-size:.88rem !important; }
    [data-testid="stMetricLabel"] { font-size:.68rem !important; }
    h1 { font-size:1.05rem !important; line-height:1.25 !important; }
    h2,h3 { font-size:.88rem !important; }
    [data-testid="stTabs"] button { font-size:.65rem !important; padding:.18rem .28rem !important; }
    div[data-testid="column"] { min-width:100% !important; flex:1 1 100% !important; }
    [data-testid="stMultiSelect"] { width:100% !important; }
    .js-plotly-plot .plotly { min-width:0 !important; }
}
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING ─────────────────────────────────────────────────────────────
def _num(val: Any) -> float:
    """Parse cell → float. Handles EU decimals, thousands, k-suffix, $, %."""
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, np.integer, float, np.floating)) and not isinstance(val, bool):
        return float(val)
    s = str(val).strip().replace("\u00a0", "").replace("$", "").replace("%", "")
    if not s or s.lower() in ("nan", "na", "n/a", "-", ""):
        return np.nan
    mult = 1.0
    if s.lower().endswith("k"):
        mult, s = 1_000.0, s[:-1].strip()
    sc = s.replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(?:,\d{3})+", sc):          # 1,234 or 1,234,567
        return float(sc.replace(",", "")) * mult
    if "," in sc and "." in sc:
        sc = sc.replace(".", "").replace(",", ".")           # 1.234,56 → 1234.56
    elif "," in sc:
        sc = sc.replace(",", ".")                            # 1,5 → 1.5
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

    # find city-header row (col C == "Auckland")
    city_hdr = next(
        (i for i in range(len(raw)) if str(raw.iloc[i, 2]).strip() == "Auckland"), -1
    )
    if city_hdr == -1:
        return pd.DataFrame(), [], "City header row not found."

    # find country-header row (col B == "Country")
    country_hdr = next(
        (i for i in range(city_hdr) if str(raw.iloc[i, 1]).strip().lower() == "country"), -1
    )

    skip_vals = {"nan", "city", "country", "category", "spendings", ""}
    hrow = raw.iloc[city_hdr]
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

    def parse_block(r0: int, r1: int) -> List[Dict[str, Any]]:
        out = []
        for ri in range(r0, r1):
            name = str(raw.iloc[ri, 1]).strip()
            if not name or name.lower() in skip_metrics:
                continue
            rec: Dict[str, Any] = {"Metric": name}
            for j, city in city_map.items():
                rec[city] = _num(raw.iloc[ri, j])
            out.append(rec)
        return out

    parsed = parse_block(city_hdr + 1, len(raw))
    city_metric_set = {r["Metric"] for r in parsed}

    # broadcast country-level rows to all cities of that country
    if country_hdr >= 0:
        ch = raw.iloc[country_hdr]
        ctry_cols: List[Tuple[int, str]] = [
            (j, str(ch.iloc[j]).strip())
            for j in range(2, len(ch))
            if str(ch.iloc[j]).strip().lower() not in skip_vals
        ]
        for ri in range(country_hdr + 1, city_hdr):
            name = str(raw.iloc[ri, 1]).strip()
            if not name or name.lower() in skip_metrics or name in city_metric_set:
                continue
            cvals = {lbl: _num(raw.iloc[ri, j]) for j, lbl in ctry_cols}
            rec = {"Metric": name}
            for city in cities:
                cc = CITY_TO_COUNTRY.get(city)
                rec[city] = cvals.get(cc, np.nan) if cc else np.nan
            parsed.append(rec)

    df = pd.DataFrame(parsed)
    df = df[~df["Metric"].duplicated(keep="first")]
    return df, cities, None


df, CITIES, _err = load_data()
if _err:
    st.error(_err)
if df.empty:
    st.error("Could not load data — place the CSV next to app.py.")
    st.stop()

MINDEX = df.set_index("Metric")


# ─── HELPERS ──────────────────────────────────────────────────────────────────
def row(metric: str, cities: List[str]) -> pd.Series:
    """Return a float Series for *metric* indexed by *cities*."""
    if metric not in MINDEX.index:
        return pd.Series({c: np.nan for c in cities}, dtype=float)
    # Use [[metric]] (list) → always DataFrame → .iloc[0] → always Series
    return pd.to_numeric(MINDEX.loc[[metric]].iloc[0].reindex(cities), errors="coerce")


def v(s: pd.Series, city: str) -> float:
    """Safe single-city value extraction."""
    val = s.get(city, np.nan)
    return float(val) if pd.notna(val) else np.nan


def fusd(val: float) -> str:
    return f"${val:,.0f}" if pd.notna(val) and not np.isnan(val) else "—"


def fpct(val: float) -> str:
    return f"{val:.1f}%" if pd.notna(val) and not np.isnan(val) else "—"


def chart(fig: go.Figure, h: int = 420) -> None:
    fig.update_layout(height=h, margin=dict(l=6, r=6, t=46, b=10))
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CFG)


def styled_df(df_in: pd.DataFrame, money_cols: List[str],
              pct_cols: Optional[List[str]] = None,
              plain_cols: Optional[List[str]] = None) -> Any:
    """
    Return a Styler with color-coding but WITHOUT background_gradient
    (which requires matplotlib and crashes on Streamlit Cloud).
    Uses bar() instead — works with only pandas+plotly.
    """
    s = df_in.style
    for col in money_cols:
        if col in df_in.columns:
            s = s.format(f"${{:,.0f}}", subset=[col], na_rep="—")
    for col in (pct_cols or []):
        if col in df_in.columns:
            s = s.format("{:.1f}%", subset=[col], na_rep="—")
    for col in (plain_cols or []):
        if col in df_in.columns:
            s = s.format("{:.1f}", subset=[col], na_rep="—")
    # colour bars — no matplotlib needed
    num_cols = df_in.select_dtypes(include="number").columns.tolist()
    for col in num_cols:
        s = s.bar(subset=[col], color=["#e05c5c", "#5cb85c"], align="mid")
    return s


def saving_metric(people: int, all_exp: bool, inc: str) -> str:
    pool = "all" if all_exp else "base"
    if people == 1:
        sal = "Avg Dev salary" if inc == "dev" else "Avg salary"
        return f"Saving {pool} 1 people on {sal}, USD/mo "
    sal = "Avg Dev + Des salary" if inc == "dev_des" else "Avg Dev salary"
    return f"Saving {pool} 2 people on {sal}, USD/mo "


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Timeline & Savings")
    start_year   = st.number_input("Start year", 2024, 2035, 2026)
    prep_years   = st.slider("Years to prepare", 1, 15, 5)
    target_year  = int(start_year) + int(prep_years)
    st.divider()
    monthly_save = st.number_input("Monthly savings ($)", 0, 100_000, 2_000, step=100)
    growth       = st.slider("Savings growth (% / yr)", 0, 50, 10)
    st.divider()
    st.caption(f"Data: 09.05.2026 · {len(CITIES)} cities · {len(MINDEX)} metrics")

# ─── TITLE + CITY SELECTOR ────────────────────────────────────────────────────
st.title("🌍 Global Relocation Intelligence 2026")
st.caption(f"{len(CITIES)} cities · {len(MINDEX)} metrics · Updated May 2026")

DEFAULTS = [c for c in CITIES
            if any(n in c for n in ["Auckland", "Warsaw", "Toronto", "Amsterdam", "Belgrade"])][:5]
selected: List[str] = st.multiselect("Select cities to compare:", options=CITIES, default=DEFAULTS)
if not selected:
    st.info("Choose at least one city above to begin.")
    st.stop()

# ─── PROFILE ──────────────────────────────────────────────────────────────────
st.subheader("📐 Calculation Profile")
pc1, pc2, pc3 = st.columns(3)

with pc1:
    family = int(st.segmented_control(
        "Family size", [1, 2], default=1,
        format_func=lambda x: "Solo" if x == 1 else "Couple", key="fam",
    ) or 1)

with pc2:
    exp_mode = str(st.segmented_control("Expenses", ["Base", "All"], default="Base", key="exp") or "Base")
    all_exp  = exp_mode == "All"

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

EXP_METRIC = f"Total {'All' if all_exp else 'Base'} / {family} people, USD/mo"
SAV_METRIC = saving_metric(family, all_exp, inc)
SAL_METRIC = "Avg Salary Dev, USD/mo" if (family == 2 or inc == "dev") else "Avg Salary, USD/mo"

# ── Custom salary inputs ───────────────────────────────────────────────────────
st.markdown("#### 💼 Your actual salary (override CSV averages)")
sc1, sc2, sc3 = st.columns(3)
with sc1:
    use_custom_sal = st.toggle("Use my own salary", value=False, key="use_csal")
with sc2:
    custom_sal_1 = st.number_input(
        "Person 1 net salary ($/mo)" if family == 2 else "Your net salary ($/mo)",
        0, 500_000, 5_000, step=100, key="csal1",
        disabled=not use_custom_sal,
        help="Your monthly take-home pay in USD after taxes",
    )
with sc3:
    custom_sal_2 = st.number_input(
        "Person 2 net salary ($/mo)",
        0, 500_000, 4_000, step=100, key="csal2",
        disabled=(not use_custom_sal or family == 1),
        help="Partner monthly take-home pay in USD. Only used when Family = Couple.",
    )

CUSTOM_SAL_TOTAL: float = float(custom_sal_1 + (custom_sal_2 if family == 2 else 0))
CUSTOM_SAL_1: float = float(custom_sal_1)
CUSTOM_SAL_2: float = float(custom_sal_2 if family == 2 else 0)

if use_custom_sal:
    if family == 2:
        st.caption(
            f"👤 Person 1: **${CUSTOM_SAL_1:,.0f}/mo**  +  "
            f"👤 Person 2: **${CUSTOM_SAL_2:,.0f}/mo**  =  "
            f"🏠 Combined: **${CUSTOM_SAL_TOTAL:,.0f}/mo**"
        )
    else:
        st.caption(f"💰 Your salary: **${CUSTOM_SAL_TOTAL:,.0f}/mo**")

with st.expander("Active CSV rows", expanded=False):
    st.caption(f"**Salary:** `{SAL_METRIC}`")
    st.caption(f"**Expenses:** `{EXP_METRIC}`")
    st.caption(f"**Saving:** `{SAV_METRIC.strip()}`")
    if use_custom_sal:
        st.caption(f"**Custom salary override:** ${CUSTOM_SAL_TOTAL:,.0f}/mo total")

# ── Unforeseen / buffer expenses ──────────────────────────────────────────────
st.markdown("#### 🆘 Unforeseen / Emergency Buffer")
_unf_c1, _unf_c2 = st.columns([3, 1])
with _unf_c1:
    UNFORESEEN_PCT: float = float(st.slider(
        "Reserve for unexpected expenses (% of net income)",
        min_value=0, max_value=50, value=15, step=5,
        help="Applied as a deduction from net saving in all calculations. 10-20% recommended.",
        key="unforeseen_pct",
    ))
with _unf_c2:
    st.metric("Reserved", f"{UNFORESEEN_PCT:.0f}%",
              help="This % is deducted from your net monthly saving everywhere.")
UNFORESEEN_FACTOR: float = 1.0 - UNFORESEEN_PCT / 100.0

st.divider()

# ─── SAVINGS ACCUMULATION (precompute) ────────────────────────────────────────
_yr_range: List[int] = list(range(int(start_year), target_year + 1))
_balance, _annual, _accum = 0.0, float(monthly_save) * 12.0, []
for _ in _yr_range:
    _balance += _annual
    _accum.append(_balance)
    _annual *= 1.0 + float(growth) / 100.0
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
    "🆚 Compare vs Minsk",
])


# ══════════════════════════════════════════════════════════════════════════════
# 0 · OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with TABS[0]:
    sal_s  = row(SAL_METRIC,  selected)
    exp_s  = row(EXP_METRIC,  selected)
    sav_s  = row(SAV_METRIC,  selected)
    rent_s = row("1 bed. ap. in city centre", selected)
    qol_s  = row("Quality of life",   selected)
    safe_s = row("Safety",             selected)

    # ── per-city snapshot cards
    for city in selected:
        sv   = CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_s, city)
        ev   = v(exp_s,  city)
        savv = v(sav_s,  city)
        if use_custom_sal:
            net = sv - ev if pd.notna(ev) else np.nan
        else:
            net = savv if pd.notna(savv) else (sv - ev if pd.notna(sv) and pd.notna(ev) else np.nan)
        rat = net / sv * 100 if pd.notna(net) and sv and sv > 0 else np.nan
        with st.expander(f"**{city}**", expanded=len(selected) <= 4):
            c1, c2, c3 = st.columns(3)
            c4, c5, c6 = st.columns(3)
            c1.metric("💰 Salary",              fusd(sv))
            c2.metric(f"💳 {exp_mode} Expenses", fusd(ev))
            c3.metric("📦 Net Save / mo",        fusd(net),
                      delta=fpct(rat) + " of income" if pd.notna(rat) else "")
            c4.metric("🏠 Rent 1BR",             fusd(v(rent_s, city)))
            c5.metric("🌟 Quality of Life",
                      f"{v(qol_s, city):.0f} / 250" if pd.notna(v(qol_s, city)) else "—")
            c6.metric("🛡️ Safety",
                      f"{v(safe_s, city):.0f} / 100" if pd.notna(v(safe_s, city)) else "—")

    # ── radar
    st.markdown("### 🕸️ Multi-Dimension Radar")
    R_M = [
        "Quality of life","Safety","Purch, power","Health care",
        "Climate","Internet Speed (Mbps)","Pollution",
    ]
    R_L = [
        "Quality of Life","Safety","Purch. Power","Healthcare",
        "Climate","Internet Mbps","Clean Air ↑",
    ]
    R_X   = [250, 100, 200, 100, 100, 350, 100]
    # For "Pollution" lower = better → invert
    R_INV = [False, False, False, False, False, False, True]

    fig_r = go.Figure()
    for i, city in enumerate(selected):
        norm = []
        for m_, mx_, inv_ in zip(R_M, R_X, R_INV):
            raw = v(row(m_, selected), city) or 0
            pct_ = min(raw / mx_ * 100, 100)
            norm.append(100 - pct_ if inv_ else pct_)

        fig_r.add_trace(go.Scatterpolar(
            r=norm + [norm[0]],
            theta=R_L + [R_L[0]],
            fill="toself",
            name=city,
            line=dict(color=COLORS[i % len(COLORS)], width=2.5),
            fillcolor=COLORS[i % len(COLORS)].replace(")", ",0.15)").replace("rgb","rgba")
                        if "rgb" in COLORS[i % len(COLORS)] else COLORS[i % len(COLORS)],
            opacity=0.9,
        ))

    fig_r.update_layout(
        polar=dict(
            bgcolor="rgba(10,20,35,0.7)",
            radialaxis=dict(
                visible=True, range=[0, 100],
                tickvals=[25, 50, 75, 100],
                ticktext=["25","50","75","100"],
                tickfont=dict(size=9, color="#aaa"),
                gridcolor="rgba(255,255,255,.12)",
                linecolor="rgba(255,255,255,.1)",
            ),
            angularaxis=dict(
                tickfont=dict(size=12, color="#ddd"),
                gridcolor="rgba(255,255,255,.08)",
                linecolor="rgba(255,255,255,.12)",
            ),
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.14, xanchor="center", x=0.5,
                    font=dict(size=11)),
        margin=dict(l=30, r=30, t=30, b=60),
    )
    chart(fig_r, 560)

    # ── quick comparison table  (no background_gradient — uses bar instead)
    st.markdown("### 📋 Quick Comparison")
    q_meta: Dict[str, str] = {
        "Salary ($/mo)"       : SAL_METRIC,
        "Base Exp 1p ($/mo)"  : "Total Base / 1 people, USD/mo",
        "Net Save/mo"         : SAV_METRIC,
        "Rent 1BR centre"     : "1 bed. ap. in city centre",
        "Quality of Life"     : "Quality of life",
        "Safety"              : "Safety",
        "Crime Index"         : "Crime index",
        "Healthcare"          : "Health care",
        "Climate"             : "Climate",
        "Purch. Power"        : "Purch, power",
        "Internet (Mbps)"     : "Internet Speed (Mbps)",
        "Pollution"           : "Pollution",
    }
    q_df = pd.DataFrame(
        {lbl: row(m, selected) for lbl, m in q_meta.items()},
        index=selected,
    ).T
    # Format: $ for money rows, 0 decimals for indices
    money_rows = ["Salary ($/mo)", "Base Exp 1p ($/mo)", "Net Save/mo", "Rent 1BR centre"]
    fmt_map = {}
    for r_name in q_df.index:
        if r_name in money_rows:
            fmt_map[r_name] = "${:,.0f}"
        else:
            fmt_map[r_name] = "{:.0f}"

    # Build plotly heatmap — no matplotlib, clean formatting
    import math
    q_norm = q_df.copy().astype(float)
    invert_rows = {"Crime Index", "Pollution"}  # lower = better
    for idx_r in q_norm.index:
        d = q_norm.loc[idx_r]; mn, mx = d.min(), d.max()
        if mx > mn:
            if idx_r in invert_rows:
                q_norm.loc[idx_r] = 1 - (d - mn) / (mx - mn)
            else:
                q_norm.loc[idx_r] = (d - mn) / (mx - mn)
        else:
            q_norm.loc[idx_r] = 0.5

    # text labels with proper formatting
    text_q = []
    for r_name in q_df.index:
        row_txt = []
        fmt_str = fmt_map.get(r_name, "{:.0f}")
        for c_name in q_df.columns:
            val_ = q_df.loc[r_name, c_name]
            if pd.notna(val_) and not math.isnan(float(val_)):
                try:
                    row_txt.append(fmt_str.format(float(val_)))
                except Exception:
                    row_txt.append(f"{val_:.0f}")
            else:
                row_txt.append("—")
        text_q.append(row_txt)

    fig_qhm = go.Figure(go.Heatmap(
        z=q_norm.values,
        x=list(q_df.columns),
        y=list(q_df.index),
        text=text_q,
        texttemplate="%{text}",
        colorscale="RdYlGn",
        showscale=False,
    ))
    fig_qhm.update_layout(
        height=max(300, len(q_df) * 36 + 60),
        margin=dict(l=130, r=4, t=30, b=8),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        xaxis=dict(side="top", tickangle=-30),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_qhm, use_container_width=True, config=PLOTLY_CFG)


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

    # show custom salary line if enabled
    if use_custom_sal:
        st.info(
            f"🔧 **Custom salary override active:** "
            f"${custom_sal_1:,.0f}/mo (person 1)"
            + (f" + ${custom_sal_2:,.0f}/mo (person 2) = **${CUSTOM_SAL_TOTAL:,.0f}/mo total**"
               if family == 2 else "")
        )

    # salary vs expenses grouped bar
    fig_fe = go.Figure()
    if use_custom_sal:
        fig_fe.add_scatter(
            x=selected, y=[CUSTOM_SAL_TOTAL] * len(selected),
            mode="lines+markers", name="Your Salary",
            line=dict(color="#f9ca24", dash="dash", width=2),
            marker=dict(size=8),
        )
    fig_fe.add_bar(x=selected, y=[v(sal_avg, c) for c in selected],
                   name="Avg Salary", marker_color="#90caf9")
    fig_fe.add_bar(x=selected, y=[v(sal_dev, c) for c in selected],
                   name="Dev Salary", marker_color="#1565c0")
    fig_fe.add_bar(x=selected, y=[v(exp_b1,  c) for c in selected],
                   name="Base Exp (solo)",   marker_color="#ef9a9a")
    fig_fe.add_bar(x=selected, y=[v(exp_b2,  c) for c in selected],
                   name="Base Exp (couple)", marker_color="#b71c1c")
    fig_fe.update_layout(barmode="group", yaxis_tickformat="$,.0f",
                         title="Salary vs Monthly Expenses")
    chart(fig_fe, 480)

    # ── all 8 savings scenarios
    st.markdown("### 💧 Net Monthly Savings — all CSV scenarios")
    SAV_SC: Dict[str, str] = {
        "1·Base·Avg"    : "Saving base 1 people on Avg salary, USD/mo ",
        "1·Base·Dev"    : "Saving base 1 people on Avg Dev salary, USD/mo ",
        "1·All·Avg"     : "Saving all 1 people on Avg salary, USD/mo ",
        "1·All·Dev"     : "Saving all 1 people on Avg Dev salary, USD/mo ",
        "2·Base·Dev"    : "Saving base 2 people on Avg Dev salary, USD/mo ",
        "2·Base·Dev+Des": "Saving base 2 people on Avg Dev + Des salary, USD/mo ",
        "2·All·Dev"     : "Saving all 2 people on Avg Dev salary, USD/mo ",
        "2·All·Dev+Des" : "Saving all 2 people on Avg Dev + Des salary, USD/mo ",
    }
    sav_long = (
        pd.DataFrame({lbl: row(m, selected) for lbl, m in SAV_SC.items()}, index=selected)
        .reset_index(names="City")
        .melt(id_vars="City", var_name="Scenario", value_name="USD/mo")
    )
    # ── custom salary net saving overlay
    if use_custom_sal:
        exp_row = row(EXP_METRIC, selected)
        csal_rows = []
        for city in selected:
            ev = v(exp_row, city)
            csal_rows.append({
                "City": city, "Scenario": "★ Your Salary",
                "USD/mo": CUSTOM_SAL_TOTAL - ev if pd.notna(ev) else np.nan,
            })
        sav_long = pd.concat([sav_long, pd.DataFrame(csal_rows)], ignore_index=True)

    # apply UNFORESEEN_FACTOR to all saving values
    sav_long["USD/mo"] = sav_long["USD/mo"] * UNFORESEEN_FACTOR

    n_scenarios = sav_long["Scenario"].nunique()
    n_cities    = len(selected)
    bar_h = max(420, 60 + n_cities * n_scenarios * 22)

    fig_sav = px.bar(
        sav_long, x="City", y="USD/mo", color="Scenario", barmode="group",
        title=f"Net monthly savings — all scenarios × {100-UNFORESEEN_PCT:.0f}% kept (USD/mo)",
        color_discrete_sequence=COLORS[:n_scenarios],
        category_orders={"Scenario": list(SAV_SC.keys()) + (["★ Your Salary"] if use_custom_sal else [])},
    )
    fig_sav.add_hline(y=0, line_dash="dash", line_color="#e74c3c", opacity=.7)
    fig_sav.update_layout(
        yaxis_tickformat="$,.0f",
        height=bar_h,
        legend=dict(orientation="h", yanchor="top", y=-0.22,
                    xanchor="left", x=0, font=dict(size=10)),
        margin=dict(l=6, r=6, t=46, b=max(130, n_scenarios * 18)),
        bargap=0.08, bargroupgap=0.04,
        xaxis=dict(tickangle=-35, automargin=True),
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_sav, use_container_width=True, config=PLOTLY_CFG)

    active_lbl = next((k for k, m in SAV_SC.items() if m == SAV_METRIC), SAV_METRIC.strip())
    st.caption(f"▶ Active profile: **{active_lbl}** · Unforeseen deduction: **{UNFORESEEN_PCT:.0f}%**")

    # ── purchasing power + tax
    st.markdown("### 💪 Purchasing Power & Tax")
    col1, col2 = st.columns(2)
    with col1:
        pp = row("Purch, power", selected)
        fig_pp = px.bar(
            x=selected, y=[v(pp, c) for c in selected],
            title="Purchasing Power Index (NYC = 100)",
            color=[v(pp, c) for c in selected],
            color_continuous_scale=[[0,"#e74c3c"],[0.5,"#f39c12"],[1,"#27ae60"]],
            labels={"x":"City","y":"Index"},
        )
        fig_pp.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fig_pp, 380)
    with col2:
        tax_sub = df[df["Metric"].isin(["salary tax, %","VAT, %"])][["Metric"] + selected]
        if not tax_sub.empty:
            tax_long = tax_sub.melt(id_vars="Metric", var_name="City", value_name="Rate")
            tax_long["Rate"] = pd.to_numeric(tax_long["Rate"], errors="coerce")
            fig_tax = px.bar(tax_long, x="City", y="Rate", color="Metric", barmode="group",
                             title="Salary Tax & VAT (%)",
                             color_discrete_map={"salary tax, %":"#e74c3c","VAT, %":"#f39c12"})
            chart(fig_tax, 380)

    # ── financial summary table
    st.markdown("### 📋 Financial Summary")
    fin = pd.DataFrame({
        "Avg Salary"        : sal_avg,
        "Dev Salary"        : sal_dev,
        "Base Exp (1p)"     : exp_b1,
        "All Exp (1p)"      : exp_a1,
        "Base Exp (2p)"     : exp_b2,
        "All Exp (2p)"      : exp_a2,
        "Net Save (profile)": row(SAV_METRIC, selected),
        "Purch. Power"      : row("Purch, power", selected),
    }, index=selected)
    if use_custom_sal:
        fin["★ Net (your sal)"] = pd.Series({
            c: CUSTOM_SAL_TOTAL - v(row(EXP_METRIC, selected), c) for c in selected
        })
    st.dataframe(fin.round(0), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# 2 · HOUSING
# ══════════════════════════════════════════════════════════════════════════════
with TABS[2]:
    st.subheader("🏠 Housing Market")

    r1c   = row("1 bed. ap. in city centre",                 selected)
    r1o   = row("1 bed. ap. out. city centre",               selected)
    r3c   = row("3 bed. ap. in city centre",                 selected)
    r3o   = row("3 bed. ap. out. city centre",               selected)
    buy_c = row("Buy Apartment Price in centre, 80m*m / USD",   selected)
    buy_o = row("Buy Apartment Price out. centre, 80m*m / USD", selected)
    sqm_c = row("USD/m*m in city centre",                    selected)
    sqm_o = row("USD/m*m out. city centre",                  selected)
    rrate = row("Interest Rate (20-Year Fixed, in %)",       selected)
    pti   = row("Prop, to income",                           selected)

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
        fg2.add_bar(x=selected, y=[v(buy_c,c) for c in selected],
                    name="80m² Centre",  marker_color="#e74c3c")
        fg2.add_bar(x=selected, y=[v(buy_o,c) for c in selected],
                    name="80m² Suburbs", marker_color="#f1948a")
        fg2.update_layout(barmode="group", yaxis_tickformat="$,.0f", title="Purchase Price — 80 m²")
        chart(fg2, 420)

    st.markdown("### 📐 Price per m²")
    fg3 = go.Figure()
    fg3.add_bar(x=selected, y=[v(sqm_c,c) for c in selected], name="$/m² Centre",  marker_color="#ab47bc")
    fg3.add_bar(x=selected, y=[v(sqm_o,c) for c in selected], name="$/m² Suburbs", marker_color="#ce93d8")
    fg3.update_layout(barmode="group", yaxis_tickformat="$,.0f", title="Price per m²")
    chart(fg3, 360)

    # ── mortgage calculator
    st.markdown("### 🏦 Mortgage Calculator (80 m², 20% down, 20-year fixed)")
    sal_ser  = row(SAL_METRIC, selected)
    eff_sal_m = {c: (CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_ser, c)) for c in selected}
    sav_ser_m = row(SAV_METRIC, selected)
    exp_ser_m = row(EXP_METRIC, selected)

    # effective monthly NET saving (after unforeseen %)
    def eff_net_m(city_: str) -> float:
        sl_ = eff_sal_m[city_]
        ev_ = v(exp_ser_m, city_)
        ns_ = v(sav_ser_m, city_)
        if use_custom_sal and pd.notna(ev_):
            ns_ = sl_ - ev_
        elif np.isnan(ns_) and pd.notna(sl_) and pd.notna(ev_):
            ns_ = sl_ - ev_
        return float(ns_) * UNFORESEEN_FACTOR if pd.notna(ns_) else np.nan

    mort_rows_list = []
    for city in selected:
        price = v(buy_c, city); r_ann = v(rrate, city)
        if any(np.isnan(x) for x in [price, r_ann]):
            continue
        loan = price * .80; r_mo = r_ann / 100 / 12; n = 240
        pmt  = loan * r_mo * (1 + r_mo)**n / ((1 + r_mo)**n - 1) if r_mo > 0 else loan / n
        sl   = eff_sal_m[city]
        net_ = eff_net_m(city)
        mort_rows_list.append({
            "City"                 : city,
            "Price 80m²"           : price,
            "Mortgage Rate %"      : r_ann,
            "Monthly Payment"      : pmt,
            "% of Salary"          : pmt / sl * 100 if sl and sl > 0 else np.nan,
            "% of Net Saving"      : pmt / net_ * 100 if pd.notna(net_) and net_ > 0 else np.nan,
            "Full price (yrs net)" : price / (net_ * 12) if pd.notna(net_) and net_ > 0 else np.nan,
        })
    if mort_rows_list:
        mdf = pd.DataFrame(mort_rows_list).set_index("City")
        st.dataframe(
            mdf.style
               .format({"Price 80m²":"${:,.0f}", "Monthly Payment":"${:,.0f}",
                        "Mortgage Rate %":"{:.2f}%", "% of Salary":"{:.1f}%",
                        "% of Net Saving":"{:.1f}%", "Full price (yrs net)":"{:.1f}"}),
            use_container_width=True,
        )
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            fig_mort = px.bar(
                mdf.reset_index(), x="City", y="Monthly Payment",
                title="Mortgage Monthly Payment ($/mo)",
                color="% of Salary", color_continuous_scale="RdYlGn_r",
            )
            fig_mort.update_layout(coloraxis_showscale=False)
            chart(fig_mort, 360)
        with col_m2:
            fig_mort2 = px.bar(
                mdf.reset_index(), x="City", y="Full price (yrs net)",
                title=f"Years of net saving to buy (after {UNFORESEEN_PCT:.0f}% reserved)",
                color="Full price (yrs net)", color_continuous_scale="RdYlGn_r",
                text=[f"{v_:.1f}y" if pd.notna(v_) else "—"
                      for v_ in mdf["Full price (yrs net)"]],
            )
            fig_mort2.update_traces(textposition="outside")
            fig_mort2.update_layout(showlegend=False, coloraxis_showscale=False)
            chart(fig_mort2, 360)

    st.markdown("### 📊 Property-to-Savings Ratio (years to buy from net savings)")
    sav_ser_h = row(SAV_METRIC, selected)
    sal_ser_h = row(SAL_METRIC, selected)
    exp_ser_h = row(EXP_METRIC, selected)

    pts_vals, pts_labels = [], []
    for c in selected:
        price_c = v(buy_c, c)
        sv_h = CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_ser_h, c)
        ev_h = v(exp_ser_h, c)
        net_h = v(sav_ser_h, c)
        if use_custom_sal and pd.notna(ev_h):
            net_h = sv_h - ev_h
        elif np.isnan(net_h) and pd.notna(sv_h) and pd.notna(ev_h):
            net_h = sv_h - ev_h
        # apply unforeseen deduction
        net_h_uf = float(net_h) * UNFORESEEN_FACTOR if pd.notna(net_h) and not np.isnan(net_h) else np.nan
        annual_saving = net_h_uf * 12 if pd.notna(net_h_uf) and not np.isnan(net_h_uf) else np.nan
        yrs = price_c / annual_saving if (
            pd.notna(price_c) and pd.notna(annual_saving) and annual_saving > 0
        ) else np.nan
        pts_vals.append(yrs)
        pts_labels.append(f"{yrs:.1f}y" if pd.notna(yrs) and not np.isnan(yrs) else "—")

    col_pti1, col_pti2 = st.columns(2)
    with col_pti1:
        fig_pti = px.bar(
            x=selected, y=pts_vals,
            title=f"Years of net savings to buy 80m² ({100-UNFORESEEN_PCT:.0f}% kept)",
            color=pts_vals, color_continuous_scale="RdYlGn_r",
            text=pts_labels,
        )
        fig_pti.update_traces(textposition="outside")
        fig_pti.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fig_pti, 380)
    with col_pti2:
        fig_pti2 = px.bar(
            x=selected, y=[v(pti,c) for c in selected],
            title="CSV Property-to-Income (years of avg salary, reference)",
            color=[v(pti,c) for c in selected], color_continuous_scale="RdYlGn_r",
            text=[f"{v(pti,c):.1f}y" if pd.notna(v(pti,c)) else "—" for c in selected],
        )
        fig_pti2.update_traces(textposition="outside")
        fig_pti2.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fig_pti2, 380)
    st.caption(
        f"Left: years of **net saving after {UNFORESEEN_PCT:.0f}% unforeseen reserve** to buy 80m² (centre).  "
        "Right: CSV benchmark (years of average local salary)."
    )


# ══════════════════════════════════════════════════════════════════════════════
# 3 · COST OF LIVING
# ══════════════════════════════════════════════════════════════════════════════
with TABS[3]:
    st.subheader("🛒 Cost of Living")

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
                    title="Monthly Spending Breakdown",
                    color_discrete_sequence=px.colors.qualitative.Set3)
    # overlay salary line
    sal_line = row(SAL_METRIC, selected)
    y_sal = [CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_line, c) for c in selected]
    fig_st.add_scatter(x=selected, y=y_sal, mode="lines+markers",
                       name="Salary" + (" (yours)" if use_custom_sal else ""),
                       line=dict(color="#f9ca24", width=2, dash="dash"),
                       marker=dict(size=7))
    fig_st.update_layout(yaxis_tickformat="$,.0f")
    chart(fig_st, 520)

    # spend as % of salary
    st.markdown("### 🔬 Spend as % of salary")
    sal_ser2 = row(SAL_METRIC, selected)
    pct_data: Dict[str, pd.Series] = {}
    for cat, m in SPEND_CATS.items():
        cat_row = row(m, selected)
        pct_data[cat] = pd.Series({
            c: v(cat_row,c) / (CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_ser2,c)) * 100
            if pd.notna(v(cat_row,c)) and (CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_ser2,c)) > 0
            else np.nan
            for c in selected
        })
    pct_df = pd.DataFrame(pct_data, index=selected)
    fig_pct = px.bar(pct_df, barmode="stack",
                     title="Spending categories as % of salary",
                     color_discrete_sequence=px.colors.qualitative.Set3)
    fig_pct.add_hline(y=100, line_dash="dash", line_color="#e74c3c", opacity=.6,
                      annotation_text="100% — full salary consumed")
    fig_pct.update_layout(yaxis_title="% of salary")
    chart(fig_pct, 460)

    # grocery heatmap (px.imshow — no matplotlib)
    st.markdown("### 🧺 Grocery & Restaurant Prices (USD)")
    GROC: Dict[str, str] = {
        "Milk 1L": "milk, l", "Bread 0.5kg": "white bread,  0,5kg",
        "Eggs 12": "eggs, 12", "Cheese kg": "cheese, kg",
        "Chicken kg": "chicken filltets, kg", "Beef kg": "beef, kg",
        "Apples kg": "apples, kg", "Potatoes kg": "potatoes, kg",
        "Tomatoes kg": "tomatoes, kg", "Bananas kg": "bananas, kg",
        "Rice kg": "rice, kg", "Water 1.5L": "bottled of watter, 1,5l",
        "Cappuccino": "cappuccino", "McCombo": "combo at mcdonald's",
        "Inexp.Meal": "meal at an inexp.rest.", "Dinner×2": "meal for two at a mid rest.",
    }
    groc_df = pd.DataFrame({k: row(m, selected) for k, m in GROC.items()}, index=selected).T
    fig_gr = px.imshow(groc_df, color_continuous_scale="RdYlGn_r",
                       title="Price Heatmap (USD)", aspect="auto", text_auto=".2f")
    fig_gr.update_coloraxes(showscale=False)
    chart(fig_gr, 520)

    # transport + cars
    st.markdown("### 🚌 Transport & Cars")
    col1, col2 = st.columns(2)
    with col1:
        fg_tr = go.Figure()
        fg_tr.add_bar(x=selected,
                      y=[v(row("public transport pass, mo", selected), c) for c in selected],
                      name="Monthly Pass", marker_color="#42a5f5")
        fg_tr.add_bar(x=selected,
                      y=[v(row("one way ticket", selected), c) for c in selected],
                      name="One-way ticket", marker_color="#90caf9")
        fg_tr.update_layout(barmode="group", yaxis_tickformat="$,.2f", title="Public Transport")
        chart(fg_tr, 360)
    with col2:
        fg_ca = go.Figure()
        fg_ca.add_bar(x=selected,
                      y=[v(row("VF Golf 1,5", selected), c) for c in selected],
                      name="VW Golf 1.5", marker_color="#26a69a")
        fg_ca.add_bar(x=selected,
                      y=[v(row("Toyoyta Corolla Sedan 1,6", selected), c) for c in selected],
                      name="Toyota Corolla", marker_color="#7e57c2")
        fg_ca.update_layout(barmode="group", yaxis_tickformat="$,.0f", title="New Car Prices")
        chart(fg_ca, 360)

    # Numbeo indices
    st.markdown("### 📊 Numbeo Indices (NYC = 100)")
    idx_df = pd.DataFrame({
        "Cost of Living": row("Cost of living",  selected),
        "Rent Index"    : row("Rent Index",       selected),
        "Groceries"     : row("Groceries Index",  selected),
    }, index=selected)
    fig_idx = px.bar(idx_df, barmode="group", title="Numbeo Indices vs New York (100)",
                     color_discrete_sequence=["#26a69a","#42a5f5","#7e57c2"])
    chart(fig_idx, 360)

    # full table
    st.markdown("### 📋 Complete Spending Table")
    ALL_SPEND = {
        **SPEND_CATS,
        "Monthly Pass"   : "public transport pass, mo",
        "Gasoline/L"     : "gasoline, l",
        "Taxi/km"        : "taxi, 1km",
        "Fitness/mo"     : "fitness club, mo",
        "Cinema"         : "cinema ticket",
        "Cappuccino"     : "cappuccino",
        "McCombo"        : "combo at mcdonald's",
        "Dinner×2"       : "meal for two at a mid rest.",
        "Internet/mo"    : "Internet (unlim,  >= 60Mpbs )",
        "Mobile/mo"      : "mobile phone (calls + 10GB+ Data)",
        "Electricity etc": "electr, heating, cooling, water, garbage, etc",
    }
    full_df = pd.DataFrame({k: row(m, selected) for k, m in ALL_SPEND.items()}, index=selected)
    st.dataframe(full_df.round(2), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# 4 · QUALITY OF LIFE
# ══════════════════════════════════════════════════════════════════════════════
with TABS[4]:
    st.subheader("🌿 Quality of Life & Environment")

    QOL_M: Dict[str, str] = {
        "Quality of Life": "Quality of life", "Safety": "Safety",
        "Crime Index": "Crime index", "Healthcare": "Health care",
        "Purch.Power": "Purch, power", "Pollution": "Pollution",
        "Climate": "Climate", "Traffic": "Traffic",
    }
    qol_df = pd.DataFrame({k: row(m, selected) for k, m in QOL_M.items()}, index=selected)
    fig_qhm = px.imshow(qol_df.T, color_continuous_scale="RdYlGn",
                        title="Quality of Life Heatmap", aspect="auto", text_auto=".1f")
    fig_qhm.update_coloraxes(showscale=False)
    chart(fig_qhm, 420)

    col1, col2 = st.columns(2)
    with col1:
        cr = row("Crime index", selected)
        fc = px.bar(x=selected, y=[v(cr,c) for c in selected], title="Crime Index (↓ safer)",
                    color=[v(cr,c) for c in selected],
                    color_continuous_scale="RdYlGn_r")
        fc.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fc, 340)
    with col2:
        hc = row("Health care", selected)
        fh = px.bar(x=selected, y=[v(hc,c) for c in selected], title="Healthcare Index (↑ better)",
                    color=[v(hc,c) for c in selected],
                    color_continuous_scale="RdYlGn")
        fh.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fh, 340)

    st.markdown("### ☀️ Climate & Weather")
    col1, col2 = st.columns(2)
    with col1:
        fg_wr = go.Figure()
        fg_wr.add_bar(x=selected,
                      y=[v(row("Sunny days, %", selected), c) for c in selected],
                      name="Sunny %", marker_color="#f9ca24")
        fg_wr.add_bar(x=selected,
                      y=[v(row("Rainy days, %", selected), c) for c in selected],
                      name="Rainy %", marker_color="#74b9ff")
        fg_wr.update_layout(barmode="group", title="Sunny vs Rainy Days (%)")
        chart(fg_wr, 360)
    with col2:
        summ_r = row("Avg Summer temp", selected)
        wint_r = row("Avg Winter temp", selected)
        fg_tp = go.Figure()
        fg_tp.add_bar(x=selected, y=[v(summ_r,c) for c in selected],
                      name="Summer °C", marker_color="#e17055")
        fg_tp.add_bar(x=selected, y=[v(wint_r,c) for c in selected],
                      name="Winter °C", marker_color="#74b9ff")
        fg_tp.update_layout(barmode="group", title="Avg Temperatures (°C)")
        chart(fg_tp, 360)

    col1, col2 = st.columns(2)
    with col1:
        pol = row("Pollution", selected)
        fp = px.bar(x=selected, y=[v(pol,c) for c in selected], title="Pollution (↓ cleaner)",
                    color=[v(pol,c) for c in selected], color_continuous_scale="RdYlGn_r")
        fp.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fp, 340)
    with col2:
        inet = row("Internet Speed (Mbps)", selected)
        fi = px.bar(x=selected, y=[v(inet,c) for c in selected], title="Internet Speed (Mbps)",
                    color=[v(inet,c) for c in selected], color_continuous_scale="Blues")
        fi.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fi, 340)

    st.markdown("### 🌲 Nature Access (hours by car) & Community")
    col1, col2 = st.columns(2)
    with col1:
        nat = pd.DataFrame({
            "Ocean/Sea h": row("Ocean/Sea, h", selected),
            "Forest h"   : row("Forest, h",    selected),
            "Mountains h": row("Mountains, h", selected),
        }, index=selected)
        fn = px.bar(nat, barmode="group", title="Hours to Nature",
                    color_discrete_sequence=["#00cec9","#55efc4","#8d6e63"])
        chart(fn, 360)
    with col2:
        fg_iv = go.Figure()
        fg_iv.add_bar(x=selected,
                      y=[v(row("Immigrants, %", selected), c) for c in selected],
                      name="Expats %", marker_color="#a29bfe")
        fg_iv.add_bar(x=selected,
                      y=[v(row("Avg Peak UV index", selected), c) for c in selected],
                      name="Peak UV", marker_color="#fdcb6e")
        fg_iv.update_layout(barmode="group", title="Expat Community & UV Index")
        chart(fg_iv, 360)


# ══════════════════════════════════════════════════════════════════════════════
# 5 · IMMIGRATION
# ══════════════════════════════════════════════════════════════════════════════
with TABS[5]:
    st.subheader("🛂 Immigration, Residency & Passport Pathways")

    # ── raw data rows
    pr_fast  = row("PR (Years) Fast Track",    selected)
    pr_gen   = row("PR (Years) General Path",  selected)
    cit_fast = row("Citizenship (Years) Fast", selected)
    cit_real = row("Citizenship (Years) Real", selected)
    pass_d   = row("BY Passport Difficulty",   selected)
    ielts    = row("IELTS / Language",         selected)
    buf_rr   = row("1 Year Buffer (USD), $k",  selected) * 1000
    sal_exp  = row("Salary (4-5y exp), $k",    selected)

    # ── 4 total journey combos:
    # PR Fast  + Citizenship Fast  → "Best case" (shortest total)
    # PR Fast  + Citizenship Real  → "Realistic fast PR"
    # PR Gen   + Citizenship Fast  → "General PR, fast citizen"
    # PR Gen   + Citizenship Real  → "Worst case" (longest total)

    st.markdown("### 🗺️ Full Pathway: Residency → Passport")
    st.caption(
        "Each bar shows the **total years from arrival to passport**.  "
        "The coloured segments show the two stages: "
        "🟦 Residency (PR) and 🟧 Citizenship on top of PR."
    )

    path_data: List[Dict] = []
    for city in selected:
        prf  = v(pr_fast,  city)
        prg  = v(pr_gen,   city)
        cf   = v(cit_fast, city)
        cr   = v(cit_real, city)
        # Citizenship years in CSV = years AFTER getting PR (incremental)
        for combo, pr_yrs, cit_yrs, label, colour_pr, colour_cit in [
            ("🟢 Best case",          prf, cf, "PR Fast + Citizen Fast",   "#27ae60", "#00b894"),
            ("🟡 Fast PR · slow cit", prf, cr, "PR Fast + Citizen Real",   "#f39c12", "#fdcb6e"),
            ("🟠 Slow PR · fast cit", prg, cf, "PR General + Citizen Fast","#e17055", "#fab1a0"),
            ("🔴 Worst case",         prg, cr, "PR General + Citizen Real","#d63031", "#ff7675"),
        ]:
            if pd.notna(pr_yrs) and pd.notna(cit_yrs):
                path_data.append({
                    "City":      city,
                    "Scenario":  combo,
                    "Label":     label,
                    "PR Stage":  pr_yrs,
                    "Citizenship Stage": cit_yrs,
                    "Total":     pr_yrs + cit_yrs,
                    "colour_pr": colour_pr,
                    "colour_cit":colour_cit,
                })

    if path_data:
        path_df = pd.DataFrame(path_data)

        # ── Stacked horizontal bar — one group per city, 4 scenarios
        fig_path = go.Figure()

        scenarios = ["🟢 Best case","🟡 Fast PR · slow cit",
                     "🟠 Slow PR · fast cit","🔴 Worst case"]
        pr_colors  = ["#27ae60","#f39c12","#e17055","#d63031"]
        cit_colors = ["#00b894","#fdcb6e","#fab1a0","#ff7675"]

        for sc, pc, cc in zip(scenarios, pr_colors, cit_colors):
            sub = path_df[path_df["Scenario"] == sc]
            if sub.empty:
                continue
            fig_path.add_bar(
                name=f"PR — {sc}",
                x=sub["PR Stage"],
                y=[f"{r['City']} | {r['Scenario']}" for _, r in sub.iterrows()],
                orientation="h",
                marker_color=pc,
                text=[f"PR: {r['PR Stage']:.0f}y" for _, r in sub.iterrows()],
                textposition="inside",
                legendgroup=sc,
                hovertemplate="<b>%{y}</b><br>PR stage: %{x:.1f} yrs<extra></extra>",
            )
            fig_path.add_bar(
                name=f"Passport — {sc}",
                x=sub["Citizenship Stage"],
                y=[f"{r['City']} | {r['Scenario']}" for _, r in sub.iterrows()],
                orientation="h",
                marker_color=cc,
                text=[f"+{r['Citizenship Stage']:.0f}y → 🛂{r['Total']:.0f}y total"
                      for _, r in sub.iterrows()],
                textposition="inside",
                legendgroup=sc,
                hovertemplate="<b>%{y}</b><br>Citizenship: +%{x:.1f} yrs<extra></extra>",
            )

        fig_path.update_layout(
            barmode="stack",
            title="Years from Arrival → Residency (PR) → Passport",
            xaxis_title="Years",
            height=max(420, len(selected) * 4 * 32 + 80),
            legend=dict(orientation="h", y=-0.18, xanchor="left", x=0, font_size=10),
            margin=dict(l=4, r=4, t=48, b=120),
            xaxis=dict(gridcolor="rgba(255,255,255,.1)"),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_path, use_container_width=True, config=PLOTLY_CFG)

    # ── Grouped bar: compare best vs worst per city
    st.markdown("### ⚡ Best vs Worst Case Total (years to passport)")
    best_tot  = [v(pr_fast, c) + v(cit_fast, c)
                 if pd.notna(v(pr_fast, c)) and pd.notna(v(cit_fast, c)) else np.nan
                 for c in selected]
    worst_tot = [v(pr_gen, c) + v(cit_real, c)
                 if pd.notna(v(pr_gen, c)) and pd.notna(v(cit_real, c)) else np.nan
                 for c in selected]

    fig_bw = go.Figure()
    fig_bw.add_bar(x=selected, y=best_tot,  name="🟢 Best (PR fast + Citizen fast)",
                   marker_color="#27ae60",
                   text=[f"{t:.0f}y" if pd.notna(t) else "—" for t in best_tot],
                   textposition="outside")
    fig_bw.add_bar(x=selected, y=worst_tot, name="🔴 Worst (PR general + Citizen real)",
                   marker_color="#d63031",
                   text=[f"{t:.0f}y" if pd.notna(t) else "—" for t in worst_tot],
                   textposition="outside")
    fig_bw.update_layout(
        barmode="group",
        title="Total Years to Passport: Best vs Worst Scenario",
        yaxis_title="Years from arrival",
    )
    chart(fig_bw, 420)

    # ── PR only chart
    st.markdown("### 🏠 Residency (PR) Timeline")
    col1, col2 = st.columns(2)
    with col1:
        fig_pr = go.Figure()
        fig_pr.add_bar(x=selected, y=[v(pr_fast,c) for c in selected],
                       name="PR Fast Track",   marker_color="#27ae60",
                       text=[f"{v(pr_fast,c):.0f}y" if pd.notna(v(pr_fast,c)) else "—"
                             for c in selected], textposition="outside")
        fig_pr.add_bar(x=selected, y=[v(pr_gen,c) for c in selected],
                       name="PR General Path", marker_color="#f39c12",
                       text=[f"{v(pr_gen,c):.0f}y" if pd.notna(v(pr_gen,c)) else "—"
                             for c in selected], textposition="outside")
        fig_pr.update_layout(barmode="group", title="Years to Permanent Residency")
        chart(fig_pr, 380)
    with col2:
        # Difficulty index for BY passport holders
        fig_pd = px.bar(x=selected, y=[v(pass_d,c) for c in selected],
                        title="BY Passport Difficulty (higher = harder to immigrate)",
                        color=[v(pass_d,c) for c in selected],
                        color_continuous_scale="RdYlGn_r",
                        text=[f"{v(pass_d,c):.1f}" if pd.notna(v(pass_d,c)) else "—"
                              for c in selected])
        fig_pd.update_traces(textposition="outside")
        fig_pd.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fig_pd, 380)

    # ── Citizenship chart
    st.markdown("### 🛂 Citizenship Timeline (years AFTER getting PR)")
    fig_cit = go.Figure()
    fig_cit.add_bar(x=selected, y=[v(cit_fast,c) for c in selected],
                    name="Citizenship Fast",   marker_color="#3498db",
                    text=[f"+{v(cit_fast,c):.0f}y" if pd.notna(v(cit_fast,c)) else "—"
                          for c in selected], textposition="outside")
    fig_cit.add_bar(x=selected, y=[v(cit_real,c) for c in selected],
                    name="Citizenship Real",   marker_color="#8e44ad",
                    text=[f"+{v(cit_real,c):.0f}y" if pd.notna(v(cit_real,c)) else "—"
                          for c in selected], textposition="outside")
    fig_cit.update_layout(
        barmode="group",
        title="Additional Years for Citizenship (on top of PR)",
        yaxis_title="Years after PR",
    )
    chart(fig_cit, 400)

    # ── Summary bubble: total range per city
    st.markdown("### 🎯 Passport Timeline Range per City")
    rng_data = []
    for city in selected:
        bt = (v(pr_fast,city) + v(cit_fast,city)
              if pd.notna(v(pr_fast,city)) and pd.notna(v(cit_fast,city)) else np.nan)
        wt_ = (v(pr_gen,city) + v(cit_real,city)
               if pd.notna(v(pr_gen,city)) and pd.notna(v(cit_real,city)) else np.nan)
        if pd.notna(bt) and pd.notna(wt_):
            rng_data.append({"City":city,"Best":bt,"Worst":wt_,"Range":wt_-bt})

    if rng_data:
        rng_df = pd.DataFrame(rng_data).sort_values("Best")
        fig_range = go.Figure()
        for _, r_ in rng_df.iterrows():
            fig_range.add_shape(
                type="line",
                x0=r_["Best"], x1=r_["Worst"],
                y0=r_["City"],  y1=r_["City"],
                line=dict(color="rgba(255,255,255,.3)", width=8),
            )
        fig_range.add_scatter(
            x=rng_df["Best"],  y=rng_df["City"],
            mode="markers+text",
            marker=dict(size=16, color="#27ae60", symbol="circle"),
            text=[f"{v:.0f}y" for v in rng_df["Best"]],
            textposition="middle right",
            name="🟢 Best case",
        )
        fig_range.add_scatter(
            x=rng_df["Worst"], y=rng_df["City"],
            mode="markers+text",
            marker=dict(size=16, color="#d63031", symbol="circle"),
            text=[f"{v:.0f}y" for v in rng_df["Worst"]],
            textposition="middle left",
            name="🔴 Worst case",
        )
        fig_range.update_layout(
            title="Passport Timeline Range: Best → Worst (years from arrival)",
            xaxis_title="Years",
            height=max(320, len(rng_df) * 55 + 80),
            xaxis=dict(gridcolor="rgba(255,255,255,.12)"),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_range, use_container_width=True, config=PLOTLY_CFG)

    # ── Buffer vs savings
    st.markdown("### 🛡️ Financial Buffer Required vs Your Savings")
    fg_bf = go.Figure()
    fg_bf.add_bar(x=selected, y=[v(buf_rr,c) for c in selected],
                  name="Required Buffer", marker_color="#e74c3c",
                  text=[f"${v(buf_rr,c):,.0f}" if pd.notna(v(buf_rr,c)) else "—"
                        for c in selected], textposition="outside")
    fg_bf.add_bar(x=selected, y=[FINAL_SAVED]*len(selected),
                  name=f"Your Savings ({target_year})", marker_color="#27ae60",
                  text=[f"${FINAL_SAVED:,.0f}"]*len(selected), textposition="outside")
    fg_bf.update_layout(barmode="group", yaxis_tickformat="$,.0f",
                        title="1-Year Buffer Required vs Your Projected Savings")
    chart(fg_bf, 400)

    cov_vals = [
        FINAL_SAVED / v(buf_rr, c) * 100
        if pd.notna(v(buf_rr, c)) and v(buf_rr, c) > 0 else 0.0
        for c in selected
    ]
    fig_cg = px.bar(x=selected, y=cov_vals,
                    title="Buffer Coverage % (your savings / required)",
                    color=cov_vals,
                    color_continuous_scale=[[0,"#e74c3c"],[0.5,"#f39c12"],[1,"#27ae60"]],
                    text=[f"{vv:.0f}%" for vv in cov_vals])
    fig_cg.update_traces(textposition="outside")
    fig_cg.add_hline(y=100, line_dash="dash", line_color="white", annotation_text="100% goal")
    fig_cg.update_layout(showlegend=False, yaxis_title="%", coloraxis_showscale=False)
    chart(fig_cg, 340)

    # ── immigration summary table
    st.markdown("### 📋 Immigration Data Table")
    IMM_M: Dict[str, str] = {
        "PR Fast (yrs)"     : "PR (Years) Fast Track",
        "PR General (yrs)"  : "PR (Years) General Path",
        "Citizen Fast (yrs)": "Citizenship (Years) Fast",
        "Citizen Real (yrs)": "Citizenship (Years) Real",
        "Best total (yrs)"  : "",  # computed below
        "Worst total (yrs)" : "",
        "Passport Diff."    : "BY Passport Difficulty",
        "IELTS"             : "IELTS / Language",
        "1yr Buffer ($k)"   : "1 Year Buffer (USD), $k",
        "Salary 4-5y ($k)"  : "Salary (4-5y exp), $k",
    }
    imm_df = pd.DataFrame({k: row(m, selected) if m else np.nan
                            for k, m in IMM_M.items()}, index=selected)
    for city in selected:
        bt_ = v(pr_fast,city)+v(cit_fast,city) if pd.notna(v(pr_fast,city)) and pd.notna(v(cit_fast,city)) else np.nan
        wt_ = v(pr_gen,city)+v(cit_real,city)  if pd.notna(v(pr_gen,city))  and pd.notna(v(cit_real,city))  else np.nan
        imm_df.loc[city, "Best total (yrs)"]  = bt_
        imm_df.loc[city, "Worst total (yrs)"] = wt_
    st.dataframe(imm_df.round(1), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# 6 · SAVINGS PLAN  (most detailed tab)
# ══════════════════════════════════════════════════════════════════════════════
with TABS[6]:
    st.subheader("📈 Savings Plan & Relocation Readiness")

    # ── parameters recap
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Monthly Saving",   fusd(float(monthly_save)))
    col2.metric("📈 Growth / yr",       fpct(float(growth)))
    col3.metric("🗓️ Target Year",       str(target_year))
    col4.metric("🎯 Projected Savings", fusd(FINAL_SAVED))

    # ── accumulation area chart
    st.markdown("### 📊 Savings Accumulation Curve")
    accum_df = pd.DataFrame({"Year": _yr_range, "Saved ($)": _accum})
    fig_acc = px.area(accum_df, x="Year", y="Saved ($)",
                      title="Total savings accumulated over time",
                      color_discrete_sequence=["#3498db"])
    fig_acc.update_traces(fill="tozeroy", fillcolor="rgba(52,152,219,0.2)",
                          line_color="#3498db")
    fig_acc.update_layout(yaxis_tickformat="$,.0f")
    # overlay annual contribution bars
    annual_contribs = [_accum[0]] + [_accum[i] - _accum[i-1] for i in range(1, len(_accum))]
    fig_acc.add_bar(x=_yr_range, y=annual_contribs, name="Annual contribution",
                    marker_color="rgba(39,174,96,0.6)", yaxis="y")
    chart(fig_acc, 440)

    # ── year-by-year table
    st.markdown("### 🗓️ Year-by-Year Breakdown")
    yy_rows = []
    bal = 0.0; ann = float(monthly_save) * 12.0
    for yr in _yr_range:
        contrib = ann
        bal    += contrib
        yy_rows.append({"Year": yr, "Annual Contribution": contrib,
                        "Total Saved": bal, "Monthly Equiv": contrib / 12})
        ann *= 1.0 + float(growth) / 100.0
    yy_df = pd.DataFrame(yy_rows).set_index("Year")
    st.dataframe(
        yy_df.style
             .format({"Annual Contribution":"${:,.0f}", "Total Saved":"${:,.0f}",
                      "Monthly Equiv":"${:,.0f}"})
             .bar(subset=["Total Saved"], color=["#3498db","#3498db"], align="left"),
        use_container_width=True,
    )

    st.divider()
    st.markdown(f"### 🎯 Readiness by {target_year}: **{fusd(FINAL_SAVED)}**")

    buf_r2 = row("1 Year Buffer (USD), $k", selected) * 1000
    exp_r2 = row(EXP_METRIC, selected)
    sav_r2 = row(SAV_METRIC, selected)
    sal_r2 = row(SAL_METRIC, selected)

    comp_rows_list = []
    for city in selected:
        bv   = v(buf_r2, city)
        ev   = v(exp_r2, city)
        sv   = CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_r2, city)
        nmo  = v(sav_r2, city)
        if np.isnan(nmo):
            nmo = sv - ev if pd.notna(sv) and pd.notna(ev) else np.nan
        elif use_custom_sal:
            nmo = CUSTOM_SAL_TOTAL - ev if pd.notna(ev) else np.nan
        cov      = FINAL_SAVED / bv * 100       if bv  and bv  > 0 else np.nan
        months   = FINAL_SAVED / ev             if ev  and ev  > 0 else np.nan
        beven    = (bv - FINAL_SAVED) / nmo     if (pd.notna(nmo) and nmo > 0
                                                    and pd.notna(bv) and bv > FINAL_SAVED) else 0.0
        wealth10 = FINAL_SAVED + (nmo or 0) * 120
        # months until buffer is achieved via pre-relocation saving
        mo_to_buf = (bv - FINAL_SAVED) / float(monthly_save) if (
            pd.notna(bv) and bv > FINAL_SAVED and monthly_save > 0) else 0.0
        comp_rows_list.append({
            "City"             : city,
            "Buffer Required"  : bv,
            "Your Savings"     : FINAL_SAVED,
            "Coverage %"       : cov,
            "Months of Cover"  : months,
            "Net Save/mo"      : nmo,
            "Break-even (mo)"  : max(float(beven), 0),
            "Extra mos needed" : max(float(mo_to_buf), 0),
            "Wealth +10yr"     : wealth10,
        })
    comp_df = pd.DataFrame(comp_rows_list).set_index("City")

    st.dataframe(
        comp_df.style
               .format({
                   "Buffer Required" : "${:,.0f}",
                   "Your Savings"    : "${:,.0f}",
                   "Coverage %"      : "{:.1f}%",
                   "Months of Cover" : "{:.1f}",
                   "Net Save/mo"     : "${:,.0f}",
                   "Break-even (mo)" : "{:.1f}",
                   "Extra mos needed": "{:.1f}",
                   "Wealth +10yr"    : "${:,.0f}",
               })
               .bar(subset=["Coverage %"],       color=["#e74c3c","#27ae60"], align="left")
               .bar(subset=["Break-even (mo)"],  color=["#27ae60","#e74c3c"], align="left"),
        use_container_width=True,
    )

    # ── buffer coverage bar
    fig_cov = px.bar(comp_df.reset_index(), x="City", y="Coverage %",
                     title=f"Buffer coverage — {fusd(FINAL_SAVED)} vs required",
                     color="Coverage %",
                     color_continuous_scale=[[0,"#e74c3c"],[0.5,"#f39c12"],[1,"#27ae60"]])
    fig_cov.add_hline(y=100, line_dash="dash", line_color="white",
                      annotation_text="100% — ready to move")
    fig_cov.update_layout(showlegend=False, coloraxis_showscale=False)
    chart(fig_cov, 380)

    # ── per-city savings target timeline
    st.markdown("### 📅 When will you hit the buffer for each city?")
    timeline_rows = []
    cumbal = 0.0; ann_ = float(monthly_save) * 12.0
    cumul_by_year: Dict[int, float] = {}
    for yr in range(int(start_year), int(start_year) + 30):
        cumbal += ann_; cumul_by_year[yr] = cumbal
        ann_ *= 1.0 + float(growth) / 100.0
    for city in selected:
        bv = v(buf_r2, city)
        if np.isnan(bv) or bv <= 0:
            continue
        hit_yr = next((yr for yr, bal in cumul_by_year.items() if bal >= bv), None)
        timeline_rows.append({
            "City"            : city,
            "Buffer Required" : bv,
            "Year to Hit"     : hit_yr or ">2055",
            "Years Away"      : (hit_yr - int(start_year)) if hit_yr else 30,
        })
    if timeline_rows:
        tl_df = pd.DataFrame(timeline_rows).set_index("City")
        fig_tl2 = px.bar(tl_df.reset_index(), x="City", y="Years Away",
                         title="Years from now until buffer is saved",
                         color="Years Away",
                         color_continuous_scale=[[0,"#27ae60"],[0.5,"#f39c12"],[1,"#e74c3c"]],
                         text="Year to Hit")
        fig_tl2.update_traces(textposition="outside")
        fig_tl2.update_layout(showlegend=False, coloraxis_showscale=False)
        chart(fig_tl2, 380)

    # ── 10-year wealth projection
    st.markdown("### 📈 10-Year Wealth Projection After Relocation")
    proj_yrs = list(range(target_year, target_year + 11))
    fig_proj = go.Figure()
    for i, city in enumerate(selected):
        sv   = CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_r2, city)
        ev   = v(exp_r2, city)
        nmo  = v(sav_r2, city)
        if np.isnan(nmo):
            nmo = sv - ev if pd.notna(sv) and pd.notna(ev) else np.nan
        elif use_custom_sal:
            nmo = sv - ev if pd.notna(ev) else np.nan
        if np.isnan(nmo):
            continue
        wealth = [FINAL_SAVED + nmo * 12 * yr for yr in range(len(proj_yrs))]
        fig_proj.add_scatter(
            x=proj_yrs, y=wealth, mode="lines+markers", name=city,
            line_color=COLORS[i % len(COLORS)],
            line=dict(width=2.5),
        )
    fig_proj.update_layout(yaxis_tickformat="$,.0f",
                           title="Cumulative wealth: savings + net income on-site")
    chart(fig_proj, 480)


# ══════════════════════════════════════════════════════════════════════════════
# 7 · FINAL SCORE
# ══════════════════════════════════════════════════════════════════════════════
with TABS[7]:
    st.subheader("🏆 Composite Relocation Score")
    st.markdown("""
    **Weighted composite (0–100)** — all inputs from your CSV:

    | Dimension | Weight | Inputs |
    |---|---|---|
    | Financial | **30%** | Net savings rate, purchasing power |
    | Quality of Life | **25%** | QoL, safety, healthcare, climate |
    | Housing | **20%** | Rent/salary, property-to-income |
    | Immigration | **15%** | PR speed, passport difficulty, buffer coverage |
    | Environment | **10%** | Pollution, internet, expat % |
    """)

    score_rows_list = []
    for city in selected:
        sal_row_s = row(SAL_METRIC, selected)
        exp_row_s = row(EXP_METRIC, selected)
        sav_row_s = row(SAV_METRIC, selected)

        sv   = CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_row_s, city)
        ev   = v(exp_row_s, city)
        savv = v(sav_row_s, city)
        if use_custom_sal:
            save_r = (CUSTOM_SAL_TOTAL - ev) / CUSTOM_SAL_TOTAL * 100 if CUSTOM_SAL_TOTAL > 0 else 0.0
        elif pd.notna(savv) and sv > 0:
            save_r = savv / sv * 100
        elif sv > 0 and pd.notna(ev):
            save_r = (sv - ev) / sv * 100
        else:
            save_r = 0.0

        def g(m: str) -> float:
            val = v(row(m, selected), city)
            return float(val) if pd.notna(val) else 0.0

        pp_v = g("Purch, power")
        fin  = np.clip(save_r * 1.2, 0, 60) + np.clip((pp_v - 50) / 150 * 40, 0, 40)

        qol  = (np.clip(g("Quality of life") / 220 * 100, 0, 100) * .40 +
                np.clip(g("Safety"),       0, 100)                 * .25 +
                np.clip(g("Health care"),  0, 100)                 * .20 +
                np.clip(g("Climate"),      0, 100)                 * .15)

        r1cv = g("1 bed. ap. in city centre")
        rr   = max((1 - r1cv / sv) * 100, 0) if sv > 0 else 50.0
        hsg  = (np.clip(rr,                                  0, 100) * .50 +
                np.clip((20 - g("Prop, to income")) / 17*100, 0, 100) * .50)

        buf_v = g("1 Year Buffer (USD), $k") * 1000
        buf_c = np.clip(FINAL_SAVED / buf_v * 100, 0, 100) if buf_v > 0 else 0.0
        imm   = (np.clip((10 - g("PR (Years) Fast Track")) / 10 * 100, 0, 100) * .50 +
                 np.clip(g("BY Passport Difficulty") / 5 * 100,          0, 100) * .30 +
                 buf_c                                                             * .20)

        env  = (np.clip(100 - g("Pollution"),                  0, 100) * .40 +
                np.clip(g("Internet Speed (Mbps)") / 350 * 100, 0, 100) * .30 +
                np.clip(g("Immigrants, %"),                     0, 100) * .30)

        composite = fin*.30 + qol*.25 + hsg*.20 + imm*.15 + env*.10
        score_rows_list.append({
            "City"           : city,
            "Financial"      : round(fin, 1),
            "Quality of Life": round(qol, 1),
            "Housing"        : round(hsg, 1),
            "Immigration"    : round(imm, 1),
            "Environment"    : round(env, 1),
            "⭐ Score"       : round(composite, 1),
        })

    score_df = (
        pd.DataFrame(score_rows_list)
          .set_index("City")
          .sort_values("⭐ Score", ascending=False)
    )

    if len(score_df) > 0:
        top = str(score_df.index[0])
        st.success(
            f"### 🥇 Best match: **{top}**  "
            f"— Score: **{score_df.loc[top, '⭐ Score']:.1f} / 100**"
        )

    # horizontal score bar
    fig_sc = px.bar(
        score_df.reset_index().sort_values("⭐ Score"),
        x="⭐ Score", y="City", orientation="h",
        color="⭐ Score",
        color_continuous_scale=[[0,"#e74c3c"],[0.5,"#f39c12"],[1,"#27ae60"]],
        title="Composite Score (0–100)",
        text="⭐ Score",
    )
    fig_sc.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    fig_sc.update_layout(showlegend=False, coloraxis_showscale=False,
                         height=max(320, len(selected) * 58))
    st.plotly_chart(fig_sc, use_container_width=True, config=PLOTLY_CFG)

    # stacked dimension contributions
    DIM_COLS = ["Financial","Quality of Life","Housing","Immigration","Environment"]
    DIM_W    = [.30, .25, .20, .15, .10]
    wt = score_df[DIM_COLS].copy()
    for col, w in zip(DIM_COLS, DIM_W):
        wt[col] = wt[col] * w
    fig_dim = px.bar(wt.reset_index(), x="City", y=DIM_COLS, barmode="stack",
                     title="Score Contribution by Dimension (weighted)",
                     color_discrete_sequence=px.colors.qualitative.Set2)
    fig_dim.update_layout(yaxis_title="Weighted Points")
    chart(fig_dim, 440)

    # score table
    st.markdown("### 📋 Score Details")
    st.dataframe(
        score_df.style
                .format("{:.1f}")
                .bar(subset=["⭐ Score"], color=["#e74c3c","#27ae60"], align="left"),
        use_container_width=True,
    )

    # city profile cards
    st.markdown("---")
    st.markdown("### 💡 City Profiles")
    sal_ser3 = row(SAL_METRIC, selected)
    exp_ser3 = row(EXP_METRIC, selected)
    sav_ser3 = row(SAV_METRIC, selected)
    pr_ser3  = row("PR (Years) Fast Track", selected)
    city_list: List[str] = [str(c) for c in score_df.index]

    for city_n, rs in score_df.iterrows():
        city_s: str = str(city_n)
        rank   = city_list.index(city_s) + 1
        medal  = "🥇" if rank == 1 else "🥈" if rank == 2 else "🥉" if rank == 3 else "🏅"
        sv_    = CUSTOM_SAL_TOTAL if use_custom_sal else v(sal_ser3, city_s)
        ev_    = v(exp_ser3, city_s)
        nmo_   = v(sav_ser3, city_s)
        if np.isnan(nmo_):
            nmo_ = sv_ - ev_ if pd.notna(sv_) and pd.notna(ev_) else np.nan
        elif use_custom_sal:
            nmo_ = sv_ - ev_ if pd.notna(ev_) else np.nan
        pr_f  = v(pr_ser3, city_s)
        buf_v = v(row("1 Year Buffer (USD), $k", selected), city_s) * 1000
        cov   = FINAL_SAVED / buf_v * 100 if buf_v > 0 else np.nan

        base_cap = (
            f"Salary: {fusd(sv_)}/mo  ·  Expenses: {fusd(ev_)}/mo  ·  "
            f"Net save: {fusd(nmo_)}/mo  ·  "
            f"Buffer coverage: {fpct(cov)}"
        )
        caption = base_cap + (f"  ·  PR: {pr_f:.0f} yrs" if pd.notna(pr_f) else "")

        with st.expander(f"{medal} #{rank}  **{city_s}**  —  {rs['⭐ Score']:.1f} / 100"):
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Financial",       f"{rs['Financial']:.1f}")
            c2.metric("Quality of Life", f"{rs['Quality of Life']:.1f}")
            c3.metric("Housing",         f"{rs['Housing']:.1f}")
            c4.metric("Immigration",     f"{rs['Immigration']:.1f}")
            c5.metric("Environment",     f"{rs['Environment']:.1f}")
            st.caption(caption)


# ══════════════════════════════════════════════════════════════════════════════
# 8 · COMPARE vs MINSK
# ══════════════════════════════════════════════════════════════════════════════
with TABS[8]:
    st.subheader("🆚 City vs City Comparison (always vs Minsk)")

    # ensure Minsk is loaded
    MINSK = "Minsk"
    if MINSK not in CITIES:
        st.warning("Minsk not found in data.")
    else:
        # pick up to 2 non-Minsk cities to compare
        other_choices = [c for c in CITIES if c != MINSK]
        default_cmp = [c for c in selected if c != MINSK][:2] or other_choices[:2]
        cmp_cities = st.multiselect(
            "Pick 1–2 cities to compare vs Minsk:",
            options=other_choices,
            default=default_cmp,
            max_selections=2,
            key="cmp_cities",
        )
        if not cmp_cities:
            st.info("Select at least one city above.")
        else:
            CMP_ALL: List[str] = cmp_cities + [MINSK]
            # ── ensure Minsk is in MINDEX columns
            cmp_avail = [c for c in CMP_ALL if c in MINDEX.columns]

            def crow(metric: str) -> pd.Series:
                return row(metric, cmp_avail)

            def ceff(city_: str, sal_ser_: "pd.Series") -> float:
                if use_custom_sal and city_ != MINSK:
                    return CUSTOM_SAL_TOTAL
                return v(sal_ser_, city_)

            # ──────────────────────────────────────────────────────────────────
            # DIMENSION TABLE with delta vs Minsk
            # ──────────────────────────────────────────────────────────────────
            st.markdown("### 📊 Head-to-Head Metrics vs Minsk")

            CMP_METRICS: Dict[str, str] = {
                "Avg Salary $/mo"        : "Avg Salary, USD/mo",
                "Dev Salary $/mo"        : "Avg Salary Dev, USD/mo",
                "Base Expenses (1p)"     : "Total Base / 1 people, USD/mo",
                "Base Expenses (2p)"     : "Total Base / 2 people, USD/mo",
                "Rent 1BR centre"        : "1 bed. ap. in city centre",
                "Buy 80m² centre"        : "Buy Apartment Price in centre, 80m*m / USD",
                "Quality of Life"        : "Quality of life",
                "Safety"                 : "Safety",
                "Crime Index ↓"          : "Crime index",
                "Healthcare"             : "Health care",
                "Climate"                : "Climate",
                "Pollution ↓"            : "Pollution",
                "Internet Mbps"          : "Internet Speed (Mbps)",
                "Purch. Power"           : "Purch, power",
                "PR Fast Track (yrs)"    : "PR (Years) Fast Track",
                "Citizenship Fast (yrs)" : "Citizenship (Years) Fast",
                "1yr Buffer $k"          : "1 Year Buffer (USD), $k",
                "Sunny days %"           : "Sunny days, %",
                "Immigrants %"           : "Immigrants, %",
            }

            cmp_df = pd.DataFrame(
                {lbl: crow(m) for lbl, m in CMP_METRICS.items()},
                index=cmp_avail,
            ).T.round(1)

            # compute delta vs Minsk for each metric
            minsk_col = cmp_df.get(MINSK) if MINSK in cmp_df.columns else None
            lower_better = {"Crime Index ↓", "Pollution ↓", "PR Fast Track (yrs)",
                            "Citizenship Fast (yrs)", "1yr Buffer $k",
                            "Base Expenses (1p)", "Base Expenses (2p)", "Rent 1BR centre"}

            # render as plotly table with delta vs Minsk
            tbl_rows = []
            for metric_lbl in cmp_df.index:
                row_d = {"Metric": metric_lbl}
                for city_ in cmp_avail:
                    val_ = cmp_df.loc[metric_lbl, city_] if city_ in cmp_df.columns else np.nan
                    row_d[city_] = val_
                # delta for non-Minsk cities
                if minsk_col is not None:
                    m_val = float(minsk_col.get(metric_lbl, np.nan)) if hasattr(minsk_col, "get") else float(minsk_col.loc[metric_lbl])
                    for cmp_c in cmp_cities:
                        if cmp_c in cmp_df.columns:
                            c_val = float(cmp_df.loc[metric_lbl, cmp_c]) if not pd.isna(cmp_df.loc[metric_lbl, cmp_c]) else np.nan
                            if pd.notna(c_val) and pd.notna(m_val) and m_val != 0:
                                delta_pct = (c_val - m_val) / abs(m_val) * 100
                                better = (delta_pct > 0) if metric_lbl not in lower_better else (delta_pct < 0)
                                sign = "▲" if delta_pct > 0 else "▼"
                                color = "#27ae60" if better else "#e74c3c"
                                row_d[f"Δ vs Minsk ({cmp_c})"] = f"<span style='color:{color}'>{sign}{abs(delta_pct):.0f}%</span>"
                tbl_rows.append(row_d)

            # show as dataframe with numeric values (delta in separate expander)
            plain_df = cmp_df.copy()
            st.dataframe(plain_df.round(1), use_container_width=True)

            # ──────────────────────────────────────────────────────────────────
            # SPIDER / RADAR comparison
            # ──────────────────────────────────────────────────────────────────
            st.markdown("### 🕸️ Radar: City vs Minsk")
            R_CMP_M = ["Quality of life","Safety","Purch, power","Health care",
                       "Climate","Internet Speed (Mbps)","Pollution"]
            R_CMP_L = ["Quality","Safety","Purch.Power","Healthcare",
                       "Climate","Internet","Clean Air"]
            R_CMP_X = [250,100,200,100,100,350,100]
            R_CMP_I = [False,False,False,False,False,False,True]

            fig_cmp_r = go.Figure()
            cmp_colors = ["#f06292","#42a5f5","#66bb6a","#ffa726"]  # pink for first
            for i_, city_ in enumerate(cmp_avail):
                nrm = []
                for m_,mx_,inv_ in zip(R_CMP_M, R_CMP_X, R_CMP_I):
                    raw_ = v(crow(m_), city_) or 0
                    p_   = min(raw_/mx_*100, 100)
                    nrm.append(100 - p_ if inv_ else p_)
                clr_ = cmp_colors[i_ % len(cmp_colors)]
                lw_  = 3.5 if city_ == MINSK else 2.0
                fig_cmp_r.add_trace(go.Scatterpolar(
                    r=nrm+[nrm[0]], theta=R_CMP_L+[R_CMP_L[0]],
                    fill="toself", name=city_,
                    line=dict(color=clr_, width=lw_,
                              dash="dot" if city_==MINSK else "solid"),
                    opacity=0.85,
                ))
            fig_cmp_r.update_layout(
                polar=dict(
                    bgcolor="rgba(10,20,35,0.7)",
                    radialaxis=dict(visible=True, range=[0,100],
                                   tickvals=[25,50,75,100],
                                   tickfont=dict(size=9,color="#aaa"),
                                   gridcolor="rgba(255,255,255,.12)"),
                    angularaxis=dict(tickfont=dict(size=12,color="#ddd"),
                                    gridcolor="rgba(255,255,255,.08)"),
                ),
                paper_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", y=-0.15, xanchor="center", x=0.5),
                margin=dict(l=30,r=30,t=30,b=60),
            )
            chart(fig_cmp_r, 560)

            # ──────────────────────────────────────────────────────────────────
            # % DELTA BARS per city vs Minsk
            # ──────────────────────────────────────────────────────────────────
            st.markdown("### 📊 % Difference vs Minsk (per metric)")
            delta_metric_keys = list(CMP_METRICS.keys())
            for cmp_city in cmp_cities:
                st.markdown(f"#### {cmp_city} vs Minsk")
                d_vals, d_colors, d_labels = [], [], []
                for lbl_d in delta_metric_keys:
                    m_v = v(crow(CMP_METRICS[lbl_d]), MINSK)
                    c_v = v(crow(CMP_METRICS[lbl_d]), cmp_city)
                    if pd.notna(m_v) and pd.notna(c_v) and m_v != 0:
                        delta_d = (c_v - m_v) / abs(m_v) * 100
                        better  = delta_d > 0 if lbl_d not in lower_better else delta_d < 0
                        d_vals.append(delta_d)
                        d_colors.append("#27ae60" if better else "#e74c3c")
                        d_labels.append(lbl_d)
                    else:
                        d_vals.append(0); d_colors.append("#888"); d_labels.append(lbl_d)

                fig_delta = go.Figure(go.Bar(
                    y=d_labels, x=d_vals,
                    orientation="h",
                    marker_color=d_colors,
                    text=[f"{x:+.0f}%" for x in d_vals],
                    textposition="outside",
                ))
                fig_delta.add_vline(x=0, line_dash="dash", line_color="white", opacity=.5)
                fig_delta.update_layout(
                    title=f"{cmp_city} relative to Minsk (green = better, red = worse)",
                    height=max(400, len(d_labels)*28 + 80),
                    xaxis_title="% difference vs Minsk",
                    margin=dict(l=4, r=60, t=46, b=8),
                    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                )
                st.plotly_chart(fig_delta, use_container_width=True, config=PLOTLY_CFG)

            # ──────────────────────────────────────────────────────────────────
            # NARRATIVE: Why better / worse than Minsk
            # ──────────────────────────────────────────────────────────────────
            st.markdown("### 💬 Why Better or Worse than Minsk?")
            for cmp_city in cmp_cities:
                better_pts, worse_pts = [], []
                NARR_METRICS = {
                    "Salary"         : ("Avg Salary, USD/mo",            False, "$"),
                    "Dev Salary"     : ("Avg Salary Dev, USD/mo",         False, "$"),
                    "Expenses (1p)"  : ("Total Base / 1 people, USD/mo",  True,  "$"),
                    "Rent 1BR"       : ("1 bed. ap. in city centre",      True,  "$"),
                    "Quality/Life"   : ("Quality of life",                False, ""),
                    "Safety"         : ("Safety",                         False, ""),
                    "Crime"          : ("Crime index",                    True,  ""),
                    "Healthcare"     : ("Health care",                    False, ""),
                    "Climate"        : ("Climate",                        False, ""),
                    "Pollution"      : ("Pollution",                      True,  ""),
                    "Internet"       : ("Internet Speed (Mbps)",          False, "Mbps"),
                    "Purch.Power"    : ("Purch, power",                   False, ""),
                    "PR Fast (yrs)"  : ("PR (Years) Fast Track",          True,  "yrs"),
                    "Citizenship"    : ("Citizenship (Years) Fast",       True,  "yrs"),
                }
                for lbl_n, (met_n, lower_b, unit) in NARR_METRICS.items():
                    m_v = v(crow(met_n), MINSK)
                    c_v = v(crow(met_n), cmp_city)
                    if pd.isna(m_v) or pd.isna(c_v) or m_v == 0:
                        continue
                    delta_n = (c_v - m_v) / abs(m_v) * 100
                    better  = delta_n > 0 if not lower_b else delta_n < 0
                    unit_s  = f" {unit}" if unit == "$" else (f" {unit}" if unit else "")
                    if unit == "$":
                        val_s = f"${c_v:,.0f}/mo vs ${m_v:,.0f}/mo"
                    elif unit == "yrs":
                        val_s = f"{c_v:.0f} vs {m_v:.0f} yrs"
                    elif unit == "Mbps":
                        val_s = f"{c_v:.0f} vs {m_v:.0f} Mbps"
                    else:
                        val_s = f"{c_v:.0f} vs {m_v:.0f}"
                    txt = f"**{lbl_n}**: {val_s} ({delta_n:+.0f}%)"
                    if better:
                        better_pts.append(txt)
                    else:
                        worse_pts.append(txt)

                with st.expander(f"📋 {cmp_city} vs Minsk — full breakdown", expanded=True):
                    col_b, col_w = st.columns(2)
                    with col_b:
                        st.markdown(f"#### ✅ {cmp_city} is **better**")
                        for pt in better_pts:
                            st.markdown(f"- {pt}")
                        if not better_pts:
                            st.caption("No advantages found in these metrics.")
                    with col_w:
                        st.markdown(f"#### ❌ {cmp_city} is **worse**")
                        for pt in worse_pts:
                            st.markdown(f"- {pt}")
                        if not worse_pts:
                            st.caption("No disadvantages found in these metrics.")

            # ──────────────────────────────────────────────────────────────────
            # FINANCIAL side-by-side bar
            # ──────────────────────────────────────────────────────────────────
            st.markdown("### 💰 Financial Direct Comparison")
            fin_cmp_metrics = {
                "Avg Salary"    : "Avg Salary, USD/mo",
                "Dev Salary"    : "Avg Salary Dev, USD/mo",
                "Base Exp (1p)" : "Total Base / 1 people, USD/mo",
                "Rent 1BR"      : "1 bed. ap. in city centre",
            }
            fig_fin_cmp = go.Figure()
            fin_colors_cmp = ["#f06292","#42a5f5","#66bb6a","#ffa726"]
            for i_c, city_ in enumerate(cmp_avail):
                fig_fin_cmp.add_bar(
                    x=list(fin_cmp_metrics.keys()),
                    y=[v(crow(m_), city_) for m_ in fin_cmp_metrics.values()],
                    name=city_,
                    marker_color=fin_colors_cmp[i_c % len(fin_colors_cmp)],
                    text=[f"${v(crow(m_),city_):,.0f}" for m_ in fin_cmp_metrics.values()],
                    textposition="outside",
                )
            fig_fin_cmp.update_layout(
                barmode="group",
                yaxis_tickformat="$,.0f",
                title="Financial Metrics: Selected Cities vs Minsk",
                legend=dict(orientation="h", y=1.12),
            )
            chart(fig_fin_cmp, 440)