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
from typing import Any, Dict, List, Optional, Tuple, cast

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


def to_safe_float(val: Any) -> float:
    try:
        return float(val)
    except (ValueError, TypeError):
        return float('nan')

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

    city_hdr = next(
        (i for i in range(len(raw)) if str(raw.iloc[i, 2]).strip() == "Auckland"), -1
    )
    if city_hdr == -1:
        return pd.DataFrame(), [], "City header row not found."

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
            
            rec: Dict[str, Any] = {"Metric": name}
            
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
    "🌐 Visa & Residency Guide",
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
            f_val = to_safe_float(val_)
            
            if pd.notna(val_) and not math.isnan(f_val):
                try:
                    row_txt.append(fmt_str.format(f_val))
                except Exception:
                    row_txt.append(f"{f_val:.0f}")
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
                       title="Price Heatmap (USD)", aspect="auto")
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
                        title="Quality of Life Heatmap", aspect="auto")
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
                    # Приводим к Any, чтобы Pylance разрешил конвертацию во float
                    if hasattr(minsk_col, "get"):
                        m_val = float(cast(Any, minsk_col.get(metric_lbl, np.nan)))
                    else:
                        m_val = float(cast(Any, minsk_col.loc[metric_lbl]))
                        
                    for cmp_c in cmp_cities:
                        if cmp_c in cmp_df.columns:
                            # Выносим сырое значение для читаемости и проверяем на натривиальные NaN
                            raw_val = cmp_df.loc[metric_lbl, cmp_c]
                            c_val = float(cast(Any, raw_val)) if not pd.isna(raw_val) else np.nan
                            
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


# ══════════════════════════════════════════════════════════════════════════════
# 9 · VISA & RESIDENCY GUIDE  (AI-powered, live web search via Claude API)
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
# 9 · VISA & RESIDENCY GUIDE  (static, structured, official sources)
# ══════════════════════════════════════════════════════════════════════════════

# ── DATA ──────────────────────────────────────────────────────────────────────
VISA_DB: Dict[str, Any] = {

  "New Zealand 🇳🇿": {
    "flag": "🇳🇿",
    "official": "https://www.immigration.govt.nz",
    "points_system": True,
    "dual_citizenship": True,
    "eu_member": False,
    "notes": "Points-based Skilled Migrant Category. Strong demand for IT, healthcare, trades.",
    "visas": [
      {"name": "Visitor Visa", "type": "tourist", "duration": "Up to 9 months",
       "fee": "NZD 246", "processing": "20–30 days",
       "requirements": ["Passport valid 3+ months", "Return ticket", "Sufficient funds (~NZD 1000/mo)", "Travel insurance"],
       "link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/about-visa/visitor-visa"},
      {"name": "Working Holiday Visa", "type": "work", "duration": "12 months (extendable to 23)",
       "fee": "NZD 310", "processing": "Online, instant–7 days",
       "requirements": ["Age 18–30 (35 for some countries)", "BY/RU not eligible — check list", "NZD 4,200 funds", "Return ticket or funds"],
       "link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/about-visa/working-holiday-visa"},
      {"name": "Skilled Migrant Category (SMC)", "type": "skilled_work", "duration": "Permanent",
       "fee": "NZD 4,310", "processing": "6–18 months",
       "requirements": ["160+ points EOI", "Job offer in NZ preferred", "IELTS 6.5+", "Age under 56", "Skilled occupation"],
       "link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/about-visa/skilled-migrant-category-resident-visa"},
      {"name": "Accredited Employer Work Visa", "type": "work", "duration": "1–3 years",
       "fee": "NZD 750", "processing": "4–8 weeks",
       "requirements": ["Job offer from accredited NZ employer", "Relevant skills/qualifications", "Median wage NZD 29.66/hr+"],
       "link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/about-visa/accredited-employer-work-visa"},
      {"name": "Student Visa", "type": "study", "duration": "Duration of course",
       "fee": "NZD 375", "processing": "4–6 weeks",
       "requirements": ["Enrollment at approved NZ institution", "NZD 15,000/yr funds", "Health insurance", "English proficiency"],
       "link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/about-visa/student-visa"},
      {"name": "Investor Visa (Investor 2)", "type": "investment", "duration": "Permanent",
       "fee": "NZD 5,050", "processing": "18+ months",
       "requirements": ["NZD 3 million investment for 4 years", "Business experience", "English or commitment to learn", "Age under 66"],
       "link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/about-visa/investor-2-category-resident-visa"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "Accredited Employer Work Visa / Student Visa",
       "min_years": "1–3 years", "key_req": "Valid job offer OR study enrollment",
       "cost": "NZD 750–1,500", "link": "https://www.immigration.govt.nz"},
      {"stage": "Permanent (PR)", "name": "Skilled Migrant Category",
       "min_years": "2 years on work visa", "key_req": "160 SMC points, job offer strongly preferred, IELTS 6.5+",
       "cost": "NZD 4,310", "link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/about-visa/skilled-migrant-category-resident-visa"},
      {"stage": "Citizenship", "name": "Naturalization",
       "min_years": "5 years PR (1,350 days physically present)", "key_req": "Good character, intent to stay, basic English",
       "cost": "NZD 470", "link": "https://www.immigration.govt.nz/new-zealand-visas/preparing-a-visa-application/citizenship"},
    ],
    "points": {
      "system_name": "Skilled Migrant Category Points",
      "min_score": 160,
      "current_cutoff": "160 (as of 2024)",
      "calculator_link": "https://www.immigration.govt.nz/new-zealand-visas/apply-for-a-visa/tools-and-information/tools/points-indicator",
      "criteria": [
        {"factor": "Age 20–39", "points": 30},
        {"factor": "Age 40–44", "points": 20},
        {"factor": "Age 45–49", "points": 10},
        {"factor": "Skilled employment in NZ (current)", "points": 50},
        {"factor": "Skilled employment offer in NZ", "points": 50},
        {"factor": "Skilled work experience 2–5 years", "points": 10},
        {"factor": "Skilled work experience 6–9 years", "points": 20},
        {"factor": "Skilled work experience 10+ years", "points": 30},
        {"factor": "NZ work experience (each year)", "points": 10},
        {"factor": "PhD from NZ institution", "points": 70},
        {"factor": "Bachelor/Master's degree", "points": 50},
        {"factor": "Diploma/Trade cert", "points": 40},
        {"factor": "Partner with skilled employment/offer", "points": 20},
        {"factor": "Job/study outside Auckland", "points": 30},
      ],
    },
  },

  "Australia 🇦🇺": {
    "flag": "🇦🇺",
    "official": "https://immi.homeaffairs.gov.au",
    "points_system": True,
    "dual_citizenship": True,
    "eu_member": False,
    "notes": "SkillSelect EOI system. 189/190/491 visas for skilled migrants. Strong tech demand.",
    "visas": [
      {"name": "Visitor Visa (subclass 600)", "type": "tourist", "duration": "3–12 months",
       "fee": "AUD 190", "processing": "1 day – 8 weeks",
       "requirements": ["Genuine visitor intent", "Sufficient funds", "Strong home ties"],
       "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/visitor-600"},
      {"name": "Skilled Independent (189)", "type": "skilled_work", "duration": "Permanent",
       "fee": "AUD 4,770", "processing": "12–24 months",
       "requirements": ["65+ EOI points", "Age under 45", "Skills assessment", "IELTS 6.0+", "Occupation on MLTSSL"],
       "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/skilled-independent-189"},
      {"name": "Skilled Nominated (190)", "type": "skilled_work", "duration": "Permanent",
       "fee": "AUD 4,770", "processing": "12–24 months",
       "requirements": ["65+ EOI points (+5 for nomination)", "State nomination", "Occupation on state list"],
       "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/skilled-nominated-190"},
      {"name": "Skilled Regional (491)", "type": "skilled_work", "duration": "5 years (pathway to PR)",
       "fee": "AUD 4,770", "processing": "12–24 months",
       "requirements": ["65+ points (+15 for regional)", "State/family nomination", "Live in regional area"],
       "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/skilled-work-regional-491"},
      {"name": "Temporary Skill Shortage (482)", "type": "work", "duration": "2–4 years",
       "fee": "AUD 1,455–2,770", "processing": "3–10 months",
       "requirements": ["Employer sponsorship", "Relevant skills + 2 yrs experience", "Occupation on STSOL/MLTSSL"],
       "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/temporary-skill-shortage-482"},
      {"name": "Global Talent (858)", "type": "exceptional", "duration": "Permanent",
       "fee": "AUD 4,770", "processing": "2–6 months",
       "requirements": ["Internationally recognized in target sector", "AUD 162,000+ salary or equivalent", "Endorsement"],
       "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/global-talent-858"},
      {"name": "Business Innovation & Investment (888)", "type": "investment", "duration": "Permanent",
       "fee": "AUD 4,770", "processing": "24+ months",
       "requirements": ["AUD 800k–1.5M investment depending on stream", "Business ownership", "State nomination"],
       "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/business-innovation-and-investment-888"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "TSS 482 / 491 / Student 500",
       "min_years": "2–5 years", "key_req": "Employer sponsorship or skills assessment",
       "cost": "AUD 1,455–4,770", "link": "https://immi.homeaffairs.gov.au"},
      {"stage": "Permanent (PR)", "name": "Subclass 189 / 190",
       "min_years": "N/A (direct PR if invited)", "key_req": "65+ points, skills assessment, IELTS 6.0+",
       "cost": "AUD 4,770", "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/skilled-independent-189"},
      {"stage": "Permanent (PR) via 491", "name": "Skilled Regional 191",
       "min_years": "3 years on 491", "key_req": "3 yrs in regional area, income threshold",
       "cost": "AUD 395", "link": "https://immi.homeaffairs.gov.au/visas/getting-a-visa/visa-listing/permanent-residence-skilled-regional-191"},
      {"stage": "Citizenship", "name": "Australian Citizenship",
       "min_years": "4 years (1 yr as PR)", "key_req": "Physical presence 4 of last 5 yrs, good character, English",
       "cost": "AUD 490", "link": "https://immi.homeaffairs.gov.au/citizenship/become-a-citizen"},
    ],
    "points": {
      "system_name": "SkillSelect Points Test",
      "min_score": 65,
      "current_cutoff": "80–95 for 189 (varies by round, check SkillSelect)",
      "calculator_link": "https://immi.homeaffairs.gov.au/points-test-advice-online/points-test",
      "criteria": [
        {"factor": "Age 25–32", "points": 30},
        {"factor": "Age 33–39", "points": 25},
        {"factor": "Age 18–24 / 40–44", "points": 15},
        {"factor": "Age 45–49", "points": 0},
        {"factor": "Nominated / regional (190/491 bonus)", "points": "5 / 15"},
        {"factor": "English: Competent (IELTS 6)", "points": 0},
        {"factor": "English: Proficient (IELTS 7)", "points": 10},
        {"factor": "English: Superior (IELTS 8)", "points": 20},
        {"factor": "Overseas work exp 3–4 yrs", "points": 5},
        {"factor": "Overseas work exp 5–7 yrs", "points": 10},
        {"factor": "Overseas work exp 8–10 yrs", "points": 15},
        {"factor": "Australian work exp 1 yr", "points": 5},
        {"factor": "Australian work exp 3 yrs", "points": 10},
        {"factor": "Australian work exp 5+ yrs", "points": 20},
        {"factor": "Bachelor degree", "points": 15},
        {"factor": "PhD", "points": 20},
        {"factor": "Australian study (2 yrs)", "points": 5},
        {"factor": "Regional study in Australia", "points": 5},
        {"factor": "Accredited community language", "points": 5},
        {"factor": "Partner skills on MLTSSL", "points": 10},
        {"factor": "Partner IELTS 6+", "points": 5},
        {"factor": "Single / partner AUS citizen", "points": 10},
        {"factor": "Professional year in AUS", "points": 5},
      ],
    },
  },

  "Canada 🇨🇦": {
    "flag": "🇨🇦",
    "official": "https://www.canada.ca/en/immigration-refugees-citizenship.html",
    "points_system": True,
    "dual_citizenship": True,
    "eu_member": False,
    "notes": "Express Entry CRS system. Provincial Nominee Programs add 600 points guaranteed.",
    "visas": [
      {"name": "Visitor Visa (TRV)", "type": "tourist", "duration": "Up to 6 months",
       "fee": "CAD 100", "processing": "2–8 weeks",
       "requirements": ["Valid passport", "Financial proof", "No criminal record", "Ties to home country"],
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/visit-canada.html"},
      {"name": "Express Entry — Federal Skilled Worker", "type": "skilled_work", "duration": "Permanent",
       "fee": "CAD 1,365 principal + CAD 1,365 spouse", "processing": "6 months after ITA",
       "requirements": ["67+ FSW points (education, experience, language, age, job offer, adaptability)", "CRS score for ITA", "IELTS CLB 7+", "1 yr skilled work exp"],
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/immigrate-canada/express-entry/eligibility/federal-skilled-workers.html"},
      {"name": "Express Entry — Canadian Experience Class", "type": "skilled_work", "duration": "Permanent",
       "fee": "CAD 1,365", "processing": "6 months",
       "requirements": ["1 yr Canadian work exp in NOC 0/A/B", "CLB 7+ (NOC 0/A) or CLB 5+ (NOC B)"],
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/immigrate-canada/express-entry/eligibility/canadian-experience-class.html"},
      {"name": "Provincial Nominee Program (PNP)", "type": "skilled_work", "duration": "Permanent",
       "fee": "CAD 1,365 + provincial fee", "processing": "12–24 months",
       "requirements": ["Province-specific criteria", "Enhanced PNP: +600 CRS points (almost certain ITA)"],
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/immigrate-canada/provincial-nominees.html"},
      {"name": "Startup Visa", "type": "entrepreneur", "duration": "Permanent",
       "fee": "CAD 1,575", "processing": "12–16 months",
       "requirements": ["Support from designated org (VC/angel/accelerator)", "CLB 5+ English", "Sufficient funds"],
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/immigrate-canada/start-up-visa.html"},
      {"name": "Study Permit", "type": "study", "duration": "Duration of study",
       "fee": "CAD 150", "processing": "4–16 weeks",
       "requirements": ["Acceptance from DLI", "Financial proof CAD 10,000+/yr", "Language proficiency"],
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/study-canada.html"},
      {"name": "Post-Graduation Work Permit", "type": "work", "duration": "8 months – 3 years",
       "fee": "CAD 255", "processing": "3–5 months",
       "requirements": ["Grad from eligible Canadian school", "Full-time study 8+ months"],
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/study-canada/work/after-graduation.html"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "Work Permit / Study Permit / IEC",
       "min_years": "1–3 years", "key_req": "Job offer, study enrollment, or IEC eligibility",
       "cost": "CAD 150–255", "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/work-canada.html"},
      {"stage": "Permanent (PR)", "name": "Express Entry / PNP",
       "min_years": "No minimum in Canada required for FSW; 1 yr CEC",
       "key_req": "CRS score above cutoff (varies ~450–550), CLB 7+, skills assessment",
       "cost": "CAD 1,365 (main app) + CAD 500 right of permanent residence",
       "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/immigrate-canada/express-entry.html"},
      {"stage": "Citizenship", "name": "Canadian Citizenship",
       "min_years": "3 of last 5 years as PR (1,095 days)", "key_req": "CLB 4+ English/French, knowledge test, tax filing, age 18–54",
       "cost": "CAD 630", "link": "https://www.canada.ca/en/immigration-refugees-citizenship/services/canadian-citizenship/become-canadian-citizen.html"},
    ],
    "points": {
      "system_name": "Comprehensive Ranking System (CRS)",
      "min_score": 67,
      "current_cutoff": "~490–520 for general rounds; lower for targeted draws (check IRCC)",
      "calculator_link": "https://www.cic.gc.ca/english/immigrate/skilled/crs-tool.asp",
      "criteria": [
        {"factor": "Age 20–29 (core / with spouse)", "points": "110 / 100"},
        {"factor": "Age 30–34", "points": "105 / 95"},
        {"factor": "Age 35–39", "points": "99 / 90"},
        {"factor": "Age 40–44", "points": "79 / 72"},
        {"factor": "PhD", "points": "150 / 140"},
        {"factor": "Master's or professional", "points": "135 / 126"},
        {"factor": "Bachelor's 3+ yrs", "points": "120 / 112"},
        {"factor": "2-year diploma", "points": "98 / 91"},
        {"factor": "1-year diploma", "points": "90 / 84"},
        {"factor": "IELTS CLB 9 (all bands)", "points": "136 / 128"},
        {"factor": "IELTS CLB 8 (all bands)", "points": "124 / 116"},
        {"factor": "IELTS CLB 7 (all bands)", "points": "116 / 108"},
        {"factor": "Canadian work exp 1 yr (NOC 0/A)", "points": "40"},
        {"factor": "Canadian work exp 2–3 yrs", "points": "53–64"},
        {"factor": "Canadian work exp 4–5 yrs", "points": "72–80"},
        {"factor": "Foreign work exp 1–2 yrs", "points": "0–13"},
        {"factor": "Foreign work exp 3+ yrs", "points": "25"},
        {"factor": "PNP nomination", "points": "+600 (guaranteed ITA)"},
        {"factor": "Valid job offer NOC 0/A/B", "points": "50–200"},
        {"factor": "Canadian study 3+ yrs post-secondary", "points": "30"},
        {"factor": "Sibling in Canada (PR/citizen)", "points": "15"},
        {"factor": "French CLB 7+ (French-strong draws)", "points": "25–50"},
      ],
    },
  },

  "Poland 🇵🇱": {
    "flag": "🇵🇱",
    "official": "https://www.gov.pl/web/mswia-en/foreigners",
    "points_system": False,
    "dual_citizenship": True,
    "eu_member": True,
    "notes": "EU member. Getting PR = EU Long-Term Residence = right to move within EU. Popular route for BY/RU/UA.",
    "visas": [
      {"name": "National Visa (D) — Work", "type": "work", "duration": "Up to 1 year, renewable",
       "fee": "EUR 80", "processing": "2–8 weeks",
       "requirements": ["Work permit or employer declaration", "Employment contract", "Accommodation proof", "Health insurance"],
       "link": "https://www.gov.pl/web/mswia-en/visas"},
      {"name": "National Visa (D) — IT Specialist", "type": "work", "duration": "Up to 3 years",
       "fee": "EUR 80", "processing": "2–4 weeks",
       "requirements": ["IT job offer", "Relevant degree or experience", "Salary above threshold"],
       "link": "https://www.gov.pl/web/poland-your-business/blue-card"},
      {"name": "EU Blue Card (Poland)", "type": "skilled_work", "duration": "3 years",
       "fee": "~PLN 440", "processing": "30 days",
       "requirements": ["Higher education (min Bachelor's)", "Job offer min PLN 10,000 gross/mo", "Relevant qualification"],
       "link": "https://udsc.gov.pl/en/cudzoziemcy/rodzaje-wiz-i-zezwolen/zezwolenie-na-pobyt-czasowy/niebieska-karta-ue/"},
      {"name": "Poland Business Harbor", "type": "startup", "duration": "1 year, renewable",
       "fee": "EUR 80", "processing": "7–14 days",
       "requirements": ["IT specialist from BY/RU/UA (specific countries)", "Employer registered in Poland"],
       "link": "https://www.gov.pl/web/polska-cyfrowa-en/poland-business-harbour"},
      {"name": "Student Visa (D)", "type": "study", "duration": "Duration of study",
       "fee": "EUR 80", "processing": "2–4 weeks",
       "requirements": ["University acceptance letter", "Accommodation", "PLN 701+/mo funds", "Health insurance"],
       "link": "https://nawa.gov.pl/en/students/study-in-poland"},
      {"name": "Karta Polaka (Polish Heritage)", "type": "special", "duration": "Permanent card",
       "fee": "Free", "processing": "1–3 months",
       "requirements": ["Polish ancestry or Polish language + Polish identity", "Interview at Polish consulate"],
       "link": "https://www.gov.pl/web/mswia-en/karta-polaka"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "Temporary Residence Permit (Karta Pobytu Czasowy)",
       "min_years": "1–3 years (renewable)", "key_req": "Valid job/study/family reason, address, health insurance, income PLN 776+/mo per person",
       "cost": "PLN 440 (work) / PLN 340 (study)", "link": "https://udsc.gov.pl/en/"},
      {"stage": "Permanent (PR)", "name": "Permanent Residence Permit / EU Long-Term Resident",
       "min_years": "5 years continuous legal stay", "key_req": "Stable income (min PLN 776+/mo), no criminal record, B1 Polish language or integration course",
       "cost": "PLN 640", "link": "https://udsc.gov.pl/en/cudzoziemcy/rodzaje-wiz-i-zezwolen/zezwolenie-na-pobyt-staly/"},
      {"stage": "Citizenship", "name": "Polish Naturalization",
       "min_years": "3 yrs PR (5 yrs if no PR)", "key_req": "3 yrs continuous PR, B1+ Polish language, stable income, no criminal record",
       "cost": "PLN 219", "link": "https://www.gov.pl/web/mswia-en/naturalisation"},
    ],
    "points": None,
  },

  "Netherlands 🇳🇱": {
    "flag": "🇳🇱",
    "official": "https://ind.nl/en",
    "points_system": False,
    "dual_citizenship": False,
    "eu_member": True,
    "notes": "No dual citizenship allowed (with exceptions). Highly developed knowledge economy. 30% tax ruling for expats.",
    "visas": [
      {"name": "Highly Skilled Migrant (HSM)", "type": "skilled_work", "duration": "1–3 years",
       "fee": "EUR 290", "processing": "2–4 weeks",
       "requirements": ["Recognised sponsor (registered company)", "Salary ≥ EUR 5,688/mo (2024, age 30+) or EUR 4,171 (<30) or EUR 3,170 (after study in NL)", "Degree"],
       "link": "https://ind.nl/en/work/working_in_the_Netherlands/Pages/Highly-skilled-migrant.aspx"},
      {"name": "EU Blue Card (Netherlands)", "type": "skilled_work", "duration": "4 years",
       "fee": "EUR 290", "processing": "4 weeks",
       "requirements": ["Higher education (min 3 yrs)", "Salary ≥ EUR 6,245/mo (2024)", "Recognised sponsor"],
       "link": "https://ind.nl/en/work/working_in_the_Netherlands/Pages/EU-Blue-Card.aspx"},
      {"name": "Orientation Year (Zoekjaar)", "type": "job_search", "duration": "1 year",
       "fee": "EUR 174", "processing": "2–4 weeks",
       "requirements": ["Graduated from top-200 university in last 3 yrs", "Or age 30+ with HSM 5+ yrs experience"],
       "link": "https://ind.nl/en/work/working_in_the_Netherlands/Pages/Orientation-year-for-highly-educated-persons.aspx"},
      {"name": "Startup Visa", "type": "entrepreneur", "duration": "1 year",
       "fee": "EUR 174", "processing": "3–4 weeks",
       "requirements": ["Innovative business plan", "Facilitator (accredited Dutch org)", "Sufficient funds"],
       "link": "https://ind.nl/en/work/working_in_the_Netherlands/Pages/Start-up.aspx"},
      {"name": "Self-Employed / Freelance (Zelfstandige)", "type": "entrepreneur", "duration": "2 years",
       "fee": "EUR 174", "processing": "3 months",
       "requirements": ["Points-based assessment (essential services, personal experience, business plan)", "Min 44/100 points for assessment", "Sufficient funds EUR 16,100"],
       "link": "https://ind.nl/en/work/working_in_the_Netherlands/Pages/Self-employed-person-and-free-lancer.aspx"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "MVV + Residence Permit (Verblijfsvergunning)",
       "min_years": "1–4 years", "key_req": "HSM / EU Blue Card / recognised sponsor; valid job",
       "cost": "EUR 290", "link": "https://ind.nl/en"},
      {"stage": "Permanent (PR)", "name": "Permanent Residence / EU Long-Term Resident",
       "min_years": "5 continuous years", "key_req": "5 yrs legal stay, A2+ Dutch OR civic integration exam, stable income (min 100% bijstandsnorm), no benefit dependency",
       "cost": "EUR 174", "link": "https://ind.nl/en/residence-permits/permanent-residence"},
      {"stage": "Citizenship", "name": "Dutch Naturalization (Inburgeringseisen)",
       "min_years": "5 yrs PR", "key_req": "5 yrs lawful stay, B1 Dutch (inburgeringsexamen), renounce other nationality (exceptions for BY/RU apply in some cases), good behavior",
       "cost": "EUR 965", "link": "https://ind.nl/en/dutch-citizenship"},
    ],
    "points": None,
  },

  "Norway 🇳🇴": {
    "flag": "🇳🇴",
    "official": "https://www.udi.no/en",
    "points_system": False,
    "dual_citizenship": True,
    "eu_member": False,
    "notes": "EEA but not EU. High salaries. Accepts dual citizenship since 2020. Skilled worker visa is main route.",
    "visas": [
      {"name": "Skilled Worker Permit", "type": "skilled_work", "duration": "1–3 years",
       "fee": "NOK 6,300", "processing": "2–8 weeks",
       "requirements": ["Job offer in Norway", "Higher education OR trade certificate", "Salary min NOK 650,000/yr (2024 threshold)", "Housing proof"],
       "link": "https://www.udi.no/en/want-to-apply/work-immigration/skilled-workers-from-countries-outside-the-eu-eea/"},
      {"name": "Job Seeker Visa (Skilled Worker)", "type": "job_search", "duration": "6 months",
       "fee": "NOK 6,300", "processing": "2–4 weeks",
       "requirements": ["Higher education or vocational", "Sufficient funds NOK 142,000+", "No prior refusals"],
       "link": "https://www.udi.no/en/want-to-apply/work-immigration/job-seeker-visa-for-skilled-workers/"},
      {"name": "Student Residence Permit", "type": "study", "duration": "Duration of study",
       "fee": "NOK 5,900", "processing": "2–4 weeks",
       "requirements": ["Accepted at Norwegian university", "NOK 142,048+/yr funds (2024)", "Housing"],
       "link": "https://www.udi.no/en/want-to-apply/studies/"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "Skilled Worker Residence Permit",
       "min_years": "1–3 years", "key_req": "Job offer, salary above threshold, qualifications",
       "cost": "NOK 6,300", "link": "https://www.udi.no/en/want-to-apply/work-immigration/"},
      {"stage": "Permanent (PR)", "name": "Permanent Residence Permit",
       "min_years": "3 years on skilled worker + total 3 yrs", "key_req": "3 yrs work permit, lived in Norway 3 of last 10 yrs, A2 Norwegian or documented language, no convictions, self-sufficient",
       "cost": "NOK 5,900", "link": "https://www.udi.no/en/want-to-apply/permanent-residence/"},
      {"stage": "Citizenship", "name": "Norwegian Citizenship",
       "min_years": "7 years (3 for EEA, special rules)", "key_req": "7 yrs legal stay (of which 2 yrs recent), A2 Norwegian language, no sentences, age 18+. Dual citizenship allowed since 2020.",
       "cost": "Free", "link": "https://www.udi.no/en/want-to-apply/citizenship/"},
    ],
    "points": None,
  },

  "Serbia 🇷🇸": {
    "flag": "🇷🇸",
    "official": "https://www.mup.gov.rs/wps/portal/en",
    "points_system": False,
    "dual_citizenship": True,
    "eu_member": False,
    "notes": "Very popular for BY/RU/UA. Visa-free entry 30 days. Easy temporary residence via work/freelance. EU candidate country.",
    "visas": [
      {"name": "Visa-Free Stay", "type": "tourist", "duration": "30 days (up to 90 in 180 for some passports)",
       "fee": "Free", "processing": "Instant",
       "requirements": ["Check your passport — BY citizens 30 days visa-free", "No registration required first 24h, then must register address"],
       "link": "https://www.mfa.gov.rs/en/consular-affairs/entry-serbia/visa-free-regimes"},
      {"name": "Temporary Residence — Employment", "type": "work", "duration": "1 year, renewable",
       "fee": "~RSD 6,720", "processing": "1–4 weeks",
       "requirements": ["Employment contract with Serbian company", "Registered address", "Health insurance", "Criminal record clearance"],
       "link": "https://www.mup.gov.rs/wps/portal/en/residence-and-movement-of-foreigners/"},
      {"name": "Temporary Residence — Self-Employment / Freelance", "type": "entrepreneur", "duration": "1 year, renewable",
       "fee": "~RSD 6,720", "processing": "2–4 weeks",
       "requirements": ["Registered company (d.o.o. ~EUR 100) or entrepreneur (preduzetnik)", "Address in Serbia", "Proof of income/work"],
       "link": "https://www.apr.gov.rs/eng/home.1026.html"},
      {"name": "Temporary Residence — Property Owner", "type": "property", "duration": "1 year, renewable",
       "fee": "~RSD 6,720", "processing": "2–4 weeks",
       "requirements": ["Own real estate in Serbia", "Proof of income from abroad or savings"],
       "link": "https://www.mup.gov.rs/wps/portal/en/"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "Privremeni Boravak (Temporary Stay)",
       "min_years": "1 year (renewable)", "key_req": "Employment contract / company / property / family / study. No income minimum specified but must be self-sufficient.",
       "cost": "~EUR 60", "link": "https://www.mup.gov.rs/wps/portal/en/residence-and-movement-of-foreigners/"},
      {"stage": "Permanent (PR)", "name": "Stalni Boravak (Permanent Stay)",
       "min_years": "5 years continuous temporary residence", "key_req": "5 yrs legal stay, registered address, no criminal record, basic Serbian language (A2)",
       "cost": "~EUR 60", "link": "https://www.mup.gov.rs/wps/portal/en/"},
      {"stage": "Citizenship", "name": "Serbian Citizenship by Naturalization",
       "min_years": "3 years PR (or 8 yrs legal stay)", "key_req": "3 yrs PR, renounce previous citizenship (dual citizenship not automatically recognized), basic Serbian, oath",
       "cost": "~EUR 20", "link": "https://www.mup.gov.rs/wps/portal/en/"},
    ],
    "points": None,
  },

  "Germany 🇩🇪": {
    "flag": "🇩🇪",
    "official": "https://www.bamf.de/EN",
    "points_system": True,
    "dual_citizenship": True,
    "eu_member": True,
    "notes": "Introduced Chancenkarte (Opportunity Card) in 2024. Dual citizenship now allowed. Strong IT demand.",
    "visas": [
      {"name": "EU Blue Card Germany", "type": "skilled_work", "duration": "4 years",
       "fee": "EUR 100", "processing": "4–8 weeks",
       "requirements": ["University degree (recognised in Germany)", "Job offer min EUR 45,300 gross/yr (2024) OR EUR 41,042 for shortage occupations (IT, engineering, medicine)", "A1 German or willingness to learn"],
       "link": "https://www.bamf.de/EN/Themen/MigrationAufenthalt/ZuwandererDrittstaaten/Blaue-Karte-EU/blaue-karte-eu-node.html"},
      {"name": "Skilled Worker Visa (Fachkräftevisum)", "type": "skilled_work", "duration": "4 years",
       "fee": "EUR 75", "processing": "4–12 weeks",
       "requirements": ["Recognised vocational/academic qualification", "Job offer", "Salary per tariff agreement", "German recognition of qualification (or exemption)"],
       "link": "https://www.bamf.de/EN/Themen/MigrationAufenthalt/ZuwandererDrittstaaten/Arbeit/arbeit-node.html"},
      {"name": "Chancenkarte (Opportunity Card)", "type": "job_search", "duration": "1 year",
       "fee": "EUR 100", "processing": "4–8 weeks",
       "requirements": ["6 out of potential 12 points: degree (4pts), German B2 (3pts), other language (1pt), Germany exp (1pt), age <35 (1pt), verified employment (1pt) + vocational qualification", "Funds EUR 1,027/mo", "A1 German or English"],
       "link": "https://www.bamf.de/EN/Themen/MigrationAufenthalt/ZuwandererDrittstaaten/Chancenkarte/chancenkarte-node.html"},
      {"name": "Student Visa", "type": "study", "duration": "2 yrs after graduation (job search)",
       "fee": "EUR 75", "processing": "4–8 weeks",
       "requirements": ["University acceptance letter", "EUR 11,208/yr in blocked account (Sperrkonto)", "Health insurance", "Language proof (B2 German or English-taught course)"],
       "link": "https://www.study-in-germany.de/en/plan-your-studies/start-planning/student-visa.html"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "Aufenthaltserlaubnis (Residence Permit)",
       "min_years": "1–4 years", "key_req": "EU Blue Card / Skilled Worker / Study",
       "cost": "EUR 100", "link": "https://www.bamf.de/EN"},
      {"stage": "Permanent (PR)", "name": "Niederlassungserlaubnis / EU Long-Term Resident",
       "min_years": "4 yrs EU Blue Card OR 5 yrs general", "key_req": "4 yrs Blue Card: B1 German, pension contributions, secure livelihood. 5 yrs general: A1+ German, pension, clean record",
       "cost": "EUR 113", "link": "https://www.bamf.de/EN/Themen/MigrationAufenthalt/ZuwandererDrittstaaten/DaueraufenthaltsberechtigungenEU/daueraufenthaltsberechtigungen-eu-node.html"},
      {"stage": "Citizenship", "name": "German Citizenship (Einbürgerung)",
       "min_years": "5 years PR (reduced to 3 yrs for special integration)", "key_req": "5 yrs legal stay, B1 German, self-sufficient, no criminal record, civic knowledge test. Dual nationality now allowed (since 2024).",
       "cost": "EUR 255", "link": "https://www.bamf.de/EN/Themen/Integration/Integrationskurse/integrationskurse-node.html"},
    ],
    "points": {
      "system_name": "Chancenkarte (Opportunity Card) Points",
      "min_score": 6,
      "current_cutoff": "6 / 12 points minimum",
      "calculator_link": "https://www.bamf.de/EN/Themen/MigrationAufenthalt/ZuwandererDrittstaaten/Chancenkarte/chancenkarte-node.html",
      "criteria": [
        {"factor": "University degree (German-recognised)", "points": 4},
        {"factor": "Vocational qualification (German-recognised)", "points": 3},
        {"factor": "German language B2 or higher", "points": 3},
        {"factor": "Other language at B2 level", "points": 1},
        {"factor": "Previous work/study experience in Germany", "points": 1},
        {"factor": "Age under 35", "points": 1},
        {"factor": "Verified employment secured", "points": 1},
        {"factor": "Degree from Top 200 university (QS/THE/Shanghai)", "points": 1},
      ],
    },
  },

  "Portugal 🇵🇹": {
    "flag": "🇵🇹",
    "official": "https://aima.gov.pt",
    "points_system": False,
    "dual_citizenship": True,
    "eu_member": True,
    "notes": "Popular for remote workers. D8 Digital Nomad visa. NHR tax regime (flat 20% for 10 yrs). EU PR after 5 yrs.",
    "visas": [
      {"name": "D8 Digital Nomad Visa", "type": "remote_work", "duration": "1 yr (→ 2yr residence)",
       "fee": "EUR 90", "processing": "2–4 months",
       "requirements": ["Remote work for non-PT company/clients", "Income ≥ EUR 3,480/mo (4x Portuguese min wage 2024)", "Health insurance", "Criminal record", "Accommodation proof"],
       "link": "https://aima.gov.pt"},
      {"name": "D7 Passive Income Visa", "type": "passive_income", "duration": "1 yr (→ 2yr residence)",
       "fee": "EUR 90", "processing": "2–4 months",
       "requirements": ["Passive income ≥ EUR 870/mo (min wage)", "Pension, rental, investments, dividends", "Health insurance", "Criminal record"],
       "link": "https://aima.gov.pt"},
      {"name": "D2 Entrepreneur / Freelancer Visa", "type": "entrepreneur", "duration": "1 year",
       "fee": "EUR 90", "processing": "2–4 months",
       "requirements": ["Business plan or freelance activity in Portugal", "Investment / economic contribution", "Tax registration"],
       "link": "https://aima.gov.pt"},
      {"name": "Tech Visa (Vistos Tech)", "type": "skilled_work", "duration": "4 years",
       "fee": "EUR 90", "processing": "1–2 months",
       "requirements": ["Job offer from certified Tech company in Portugal", "Bachelor's+ degree or 5 yrs experience"],
       "link": "https://startupportugal.com/tech-visa/"},
      {"name": "Golden Visa (ARI) — Investment", "type": "investment", "duration": "2 yrs (renewable)",
       "fee": "EUR 533 (application)", "processing": "6–12 months",
       "requirements": ["EUR 500,000+ fund investment (real estate route closed since 2023)", "No minimum stay required", "Clean criminal record"],
       "link": "https://aima.gov.pt/en/golden-visa"},
    ],
    "residency": [
      {"stage": "Temporary", "name": "Autorização de Residência Temporária",
       "min_years": "1+2 years", "key_req": "Valid visa (D7/D8/D2/work), address, NIF, NISS, health insurance",
       "cost": "EUR 533", "link": "https://aima.gov.pt"},
      {"stage": "Permanent (PR)", "name": "Autorização de Residência Permanente / EU Long-Term",
       "min_years": "5 years legal residence", "key_req": "5 yrs stay, A2 Portuguese language, no convictions, fiscal compliance",
       "cost": "EUR 533", "link": "https://aima.gov.pt"},
      {"stage": "Citizenship", "name": "Portuguese Naturalization",
       "min_years": "5 years legal residence", "key_req": "5 yrs, A2 Portuguese, clean record, ties to Portugal. Dual citizenship allowed.",
       "cost": "EUR 250", "link": "https://justica.gov.pt/Servicos/Pedido-de-aquisicao-da-nacionalidade-portuguesa"},
    ],
    "points": None,
  },

}

with TABS[9]:
    st.subheader("🌐 Visa & Residency Guide")
    st.caption("Static reference data from official sources. Always verify current rules at official government websites before applying.")

    # ── Country selector
    country_names = list(VISA_DB.keys())
    sel_country = st.selectbox(
        "🌍 Select country:",
        country_names,
        key="vg_country_sel",
    )
    cdata = VISA_DB[sel_country]

    # ── Country header bar
    st.markdown(f"""
<div style="background:linear-gradient(135deg,#0d2137,#1a3a5c);border-radius:12px;padding:18px 24px;margin:8px 0 18px;display:flex;align-items:center;gap:16px;">
  <span style="font-size:2.8rem">{cdata['flag']}</span>
  <div>
    <div style="font-size:1.5rem;font-weight:700;color:#fff">{sel_country}</div>
    <div style="font-size:.85rem;color:#90caf9;margin-top:3px">{cdata['notes']}</div>
  </div>
  <div style="margin-left:auto;display:flex;gap:12px;flex-wrap:wrap;">
    <span style="background:{'#1b5e20' if cdata['eu_member'] else '#37474f'};color:#fff;padding:3px 10px;border-radius:20px;font-size:.8rem">{'🇪🇺 EU Member' if cdata['eu_member'] else '🌐 Non-EU'}</span>
    <span style="background:{'#1b5e20' if cdata['dual_citizenship'] else '#b71c1c'};color:#fff;padding:3px 10px;border-radius:20px;font-size:.8rem">{'✅ Dual Citizenship' if cdata['dual_citizenship'] else '❌ No Dual Citizenship'}</span>
    <span style="background:{'#1565c0' if cdata['points_system'] else '#37474f'};color:#fff;padding:3px 10px;border-radius:20px;font-size:.8rem">{'🧮 Points System' if cdata['points_system'] else '📋 Non-points'}</span>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── navigation within country
    subtabs = st.tabs(["🛫 All Visas", "🏠 Residency Pathway", "🛂 Citizenship",
                       "🧮 Points Calculator" if cdata["points_system"] else "📋 Key Requirements",
                       "🔗 Official Links"])

    # ══ SUBTAB 0: ALL VISAS ══════════════════════════════════════════════════
    with subtabs[0]:
        st.markdown("### Available Visa Types")

        # filter
        VISA_TYPE_LABELS = {
            "tourist": "✈️ Tourist/Visitor", "work": "💼 Work",
            "skilled_work": "🎓 Skilled Worker", "study": "📚 Study",
            "entrepreneur": "🚀 Entrepreneur/Startup", "investment": "💰 Investment",
            "job_search": "🔍 Job Search", "remote_work": "💻 Remote/Digital Nomad",
            "passive_income": "💴 Passive Income", "special": "⭐ Special",
            "exceptional": "🏆 Exceptional Talent", "property": "🏡 Property Owner",
        }
        all_types = list({v["type"] for v in cdata["visas"]})
        type_labels = [VISA_TYPE_LABELS.get(t, t) for t in all_types]
        selected_types = st.multiselect(
            "Filter by type:", type_labels,
            default=type_labels, key="vg_type_filter"
        )
        selected_raw = [all_types[i] for i, lbl in enumerate(type_labels) if lbl in selected_types]

        for visa in cdata["visas"]:
            if visa["type"] not in selected_raw:
                continue
            type_lbl = VISA_TYPE_LABELS.get(visa["type"], visa["type"])
            badge_color = {
                "tourist": "#37474f", "work": "#1565c0", "skilled_work": "#1b5e20",
                "study": "#4a148c", "entrepreneur": "#e65100", "investment": "#827717",
                "job_search": "#006064", "remote_work": "#0d47a1", "passive_income": "#1a237e",
                "special": "#880e4f", "exceptional": "#bf360c", "property": "#33691e",
            }.get(visa["type"], "#333")

            with st.expander(f"**{visa['name']}** — {type_lbl}", expanded=False):
                c1, c2, c3 = st.columns(3)
                c1.metric("⏱ Duration", visa["duration"])
                c2.metric("💳 Fee", visa["fee"])
                c3.metric("🕐 Processing", visa["processing"])

                st.markdown("**📋 Requirements:**")
                for req in visa["requirements"]:
                    st.markdown(f"- {req}")

                st.markdown(
                    f'<a href="{visa["link"]}" target="_blank" style="'
                    f'background:{badge_color};color:#fff;padding:6px 14px;'
                    f'border-radius:6px;text-decoration:none;font-size:.85rem;'
                    f'display:inline-block;margin-top:6px">🔗 Official Information</a>',
                    unsafe_allow_html=True
                )

    # ══ SUBTAB 1: RESIDENCY PATHWAY ══════════════════════════════════════════
    with subtabs[1]:
        st.markdown("### 🗺️ Residency → Permanent Residence → Citizenship Pathway")

        stages = cdata["residency"]
        stage_colors = {"Temporary": "#1565c0", "Permanent (PR)": "#1b5e20",
                        "Permanent (PR) via 491": "#2e7d32", "Citizenship": "#6a1b9a"}

        # Timeline visualization
        timeline_html = '<div style="display:flex;align-items:flex-start;gap:0;overflow-x:auto;padding:16px 0 24px;">'
        for i, stage in enumerate(stages):
            clr = stage_colors.get(stage["stage"], "#37474f")
            is_last = i == len(stages) - 1
            timeline_html += f'''
<div style="display:flex;flex-direction:column;align-items:center;min-width:160px;flex:1;">
  <div style="background:{clr};color:#fff;border-radius:50%;width:44px;height:44px;
              display:flex;align-items:center;justify-content:center;font-size:1.3rem;
              font-weight:700;flex-shrink:0;z-index:1;">
    {i+1}
  </div>
  <div style="background:{clr}22;border:1px solid {clr};border-radius:8px;padding:10px 12px;
              margin-top:8px;width:90%;text-align:center;">
    <div style="color:{clr};font-weight:700;font-size:.85rem">{stage["stage"]}</div>
    <div style="color:#fff;font-size:.8rem;margin-top:4px">{stage["name"]}</div>
    <div style="color:#aaa;font-size:.75rem;margin-top:2px">{stage["min_years"]}</div>
  </div>
</div>
'''
            if not is_last:
                timeline_html += '<div style="margin-top:20px;color:#555;font-size:1.3rem;flex-shrink:0;">→</div>'
        timeline_html += '</div>'
        st.markdown(timeline_html, unsafe_allow_html=True)

        # Detailed cards
        for stage in stages:
            clr = stage_colors.get(stage["stage"], "#37474f")
            st.markdown(f"""
<div style="border-left:4px solid {clr};background:{clr}15;border-radius:0 8px 8px 0;
            padding:14px 18px;margin:10px 0;">
  <div style="font-weight:700;color:{clr};font-size:1rem">{stage['stage']}: {stage['name']}</div>
  <div style="margin-top:8px;display:grid;grid-template-columns:1fr 1fr;gap:8px;">
    <div><span style="color:#aaa;font-size:.78rem">⏱ TIMELINE</span>
         <div style="color:#eee;font-size:.88rem">{stage['min_years']}</div></div>
    <div><span style="color:#aaa;font-size:.78rem">💳 FEE</span>
         <div style="color:#eee;font-size:.88rem">{stage['cost']}</div></div>
  </div>
  <div style="margin-top:8px;"><span style="color:#aaa;font-size:.78rem">📋 KEY REQUIREMENTS</span>
       <div style="color:#ddd;font-size:.85rem;margin-top:2px">{stage['key_req']}</div></div>
  <a href="{stage['link']}" target="_blank" style="display:inline-block;margin-top:10px;
     background:{clr};color:#fff;padding:4px 12px;border-radius:4px;
     text-decoration:none;font-size:.8rem;">🔗 Official source</a>
</div>
""", unsafe_allow_html=True)

    # ══ SUBTAB 2: CITIZENSHIP ════════════════════════════════════════════════
    with subtabs[2]:
        st.markdown("### 🛂 Citizenship & Naturalization Details")
        citizenship = [s for s in cdata["residency"] if "Citizenship" in s["stage"]]
        if citizenship:
            cs = citizenship[0]
            col1, col2 = st.columns(2)
            col1.metric("📅 Min. Legal Stay", cs["min_years"])
            col2.metric("💳 Application Fee", cs["cost"])

            st.markdown(f"""
<div style="background:#4a148c22;border:1px solid #9c27b0;border-radius:10px;padding:16px 20px;margin:12px 0">
  <div style="color:#ce93d8;font-weight:700;font-size:.95rem;margin-bottom:10px">📋 Key Requirements</div>
  <div style="color:#e1bee7;font-size:.88rem;line-height:1.7">{cs['key_req']}</div>
</div>
""", unsafe_allow_html=True)

        # Dual citizenship warning
        if cdata["dual_citizenship"]:
            st.success("✅ **Dual citizenship is allowed** — you do not need to renounce your current passport.")
        else:
            st.error("❌ **Dual citizenship is generally NOT allowed** — you may need to renounce your current passport. Check exceptions with a lawyer.")

        if citizenship:
            st.markdown(f"[🔗 Official Citizenship Information]({citizenship[0]['link']})")

    # ══ SUBTAB 3: POINTS / KEY REQUIREMENTS ═════════════════════════════════
    with subtabs[3]:
        if cdata["points_system"] and cdata.get("points"):
            pts = cdata["points"]
            st.markdown(f"### 🧮 {pts['system_name']}")

            col1, col2, col3 = st.columns(3)
            col1.metric("Minimum Score", str(pts["min_score"]))
            col2.metric("Current Cutoff", pts["current_cutoff"].split("(")[0].strip())
            col3.markdown(f"[📊 Official Calculator]({pts['calculator_link']})")

            st.markdown("#### 📋 Scoring Criteria")
            pts_data_rows = [
                {"Factor": c["factor"], "Points": str(c["points"])}
                for c in pts["criteria"]
            ]
            pts_display = pd.DataFrame(pts_data_rows)

            # Interactive calculator
            st.markdown("#### 🧮 Interactive Points Estimator")
            st.caption("Check boxes that apply to you to estimate your score:")
            user_score = 0
            for i, c in enumerate(pts["criteria"], start=1): 
                pts_val = c["points"]
                # handle ranges like "5 / 15"
                if isinstance(pts_val, str) and "/" in str(pts_val):
                    pts_num = int(str(pts_val).split("/")[0].strip())
                    label = f"✅ {c['factor']} — **{pts_val} pts**"
                elif isinstance(pts_val, str) and "+" in str(pts_val):
                    pts_num = int(str(pts_val).replace("+","").split()[0])
                    label = f"✅ {c['factor']} — **{pts_val} pts**"
                else:
                    try:
                        pts_num = int(str(pts_val))
                    except Exception:
                        pts_num = 0
                    label = f"✅ {c['factor']} — **{pts_val} pts**"
                
                # 2. МЕНЯЕМ KEY (добавляем {i} и убираем срез [:20])
                if st.checkbox(label, key=f"pts_{sel_country}_{i}"): 
                    user_score += pts_num

            # Score result
            min_s = pts["min_score"]
            pct_of_min = user_score / min_s * 100 if min_s > 0 else 0
            bar_clr = "#27ae60" if user_score >= min_s else ("#f39c12" if pct_of_min >= 70 else "#e74c3c")
            st.markdown(f"""
<div style="background:#0d2137;border-radius:10px;padding:16px 20px;margin:12px 0;text-align:center;">
  <div style="font-size:2.2rem;font-weight:800;color:{bar_clr}">{user_score} pts</div>
  <div style="color:#aaa;font-size:.9rem">Your estimated score</div>
  <div style="background:#1a3a5c;border-radius:8px;height:12px;margin:12px 0;overflow:hidden;">
    <div style="background:{bar_clr};height:100%;width:{min(pct_of_min,100):.0f}%;
                border-radius:8px;transition:width .3s;"></div>
  </div>
  <div style="color:#ddd;font-size:.85rem">
    Minimum required: <b style="color:{bar_clr}">{min_s}</b> pts
    {'— ✅ <b style="color:#27ae60">You qualify!</b>' if user_score >= min_s
      else f'— ❌ Need <b style="color:#e74c3c">{min_s - user_score} more points</b>'}
  </div>
</div>
""", unsafe_allow_html=True)
            st.dataframe(pts_display, use_container_width=True, hide_index=True)

        else:
            st.markdown("### 📋 Key Requirements Summary")
            st.info(f"{sel_country} does not use a points-based immigration system. Eligibility is determined by individual circumstances, job offers, and documentation.")
            # Show residency summary table
            rdata = []
            for stage in cdata["residency"]:
                rdata.append({
                    "Stage": stage["stage"],
                    "Name": stage["name"],
                    "Min. Years": stage["min_years"],
                    "Cost": stage["cost"],
                })
            st.dataframe(pd.DataFrame(rdata), use_container_width=True, hide_index=True)

    # ══ SUBTAB 4: OFFICIAL LINKS ═════════════════════════════════════════════
    with subtabs[4]:
        st.markdown("### 🔗 Official Sources & Useful Links")

        st.markdown(f"""
<a href="{cdata['official']}" target="_blank"
   style="display:flex;align-items:center;gap:12px;background:#1a3a5c;border:1px solid #2a5a8c;
          border-radius:10px;padding:16px 20px;text-decoration:none;margin-bottom:12px;">
  <span style="font-size:2rem">{cdata['flag']}</span>
  <div>
    <div style="color:#4fc3f7;font-weight:700;font-size:1rem">Official Immigration Portal</div>
    <div style="color:#aaa;font-size:.82rem">{cdata['official']}</div>
  </div>
  <span style="margin-left:auto;color:#4fc3f7">→</span>
</a>
""", unsafe_allow_html=True)

        # visa-specific links
        st.markdown("**All visa types — direct links:**")
        link_cols = st.columns(2)
        for i, visa in enumerate(cdata["visas"]):
            type_lbl = VISA_TYPE_LABELS.get(visa["type"], visa["type"])
            link_cols[i % 2].markdown(f"[{type_lbl}: {visa['name']}]({visa['link']})")

        # citizenship link
        cit_links = [s for s in cdata["residency"] if "Citizenship" in s["stage"]]
        if cit_links:
            st.markdown(f"\n**Citizenship:** [{cit_links[0]['name']}]({cit_links[0]['link']})")

        if cdata.get("points") and cdata["points"].get("calculator_link"):
            st.markdown(f"\n**Points Calculator:** [Official Tool]({cdata['points']['calculator_link']})")

        # general tips
        st.markdown("---")
        st.markdown("### 💡 General Tips")
        tips = [
            "Always check the **official government portal** for the most up-to-date fee and requirement information — rules change frequently.",
            "**Language tests** (IELTS, DELF, Goethe) should be booked 2–3 months in advance as seats fill quickly.",
            "For **qualification recognition** in EU countries, use ENIC-NARIC network: [enic-naric.net](https://www.enic-naric.net)",
            "Keep **apostilled copies** of all documents — most countries require official translation + apostille.",
            "**Tax residency** changes when you move — consult a tax advisor about potential double taxation agreements.",
            "For EU countries, PR gives you the right to **EU Long-Term Resident** status — mobility within the EU.",
        ]
        for tip in tips:
            st.markdown(f"- {tip}")