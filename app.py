import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

DATA_PATH = Path(__file__).resolve().parent / "Country.xlsx - Лист1.csv"

# City → country label as in the CSV country header row (with flags where present).
CITY_COUNTRY: Dict[str, str] = {
    "Auckland": "New Zealand 🇳🇿",
    "Wellington": "New Zealand 🇳🇿",
    "Christchurch": "New Zealand 🇳🇿",
    "Melbourne": "Australia 🇦🇺",
    "Sydney": "Australia 🇦🇺",
    "Brisbane": "Australia 🇦🇺",
    "Perth": "Australia 🇦🇺",
    "Adelaide": "Australia 🇦🇺",
    "Toronto": "Canada 🇨🇦",
    "Montreal": "Canada 🇨🇦",
    "Ottawa": "Canada 🇨🇦",
    "Chicago": "USA 🇺🇸",
    "Philadelphia": "USA 🇺🇸",
    "San Diego": "USA 🇺🇸",
    "Dallas": "USA 🇺🇸",
    "Austin": "USA 🇺🇸",
    "Seattle": "USA 🇺🇸",
    "Birmingham": "United Kingdom 🇬🇧",
    "Liverpool": "United Kingdom 🇬🇧",
    "Southhampton": "United Kingdom 🇬🇧",
    "Manchester": "United Kingdom 🇬🇧",
    "Amsterdam": "Netherlands 🇳🇱",
    "Rotterdam": "Netherlands 🇳🇱",
    "Oslo": "Norway 🇳🇴",
    "Belgrade": "Serbia 🇷🇸",
    "Warsaw": "Poland 🇵🇱",
    "Krakow": "Poland 🇵🇱",
    "Lodz": "Poland 🇵🇱",
    "Gdansk": "Poland 🇵🇱",
    "Moscow": "Russia 🇷🇺",
    "Minsk": "Belarus",
}

PLOTLY_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "responsive": True,
    "scrollZoom": False,
}

st.set_page_config(
    page_title="Relocation Intelligence 2026",
    layout="wide",
    initial_sidebar_state="collapsed",
    page_icon="🌍",
)

st.markdown(
    """
<style>
[data-testid="stMetricValue"] { font-size: 1.35rem; }
.block-container {
    padding-top: 0.75rem !important;
    padding-left: min(1.25rem, 4vw) !important;
    padding-right: min(1.25rem, 4vw) !important;
    max-width: 1400px;
}
@media (max-width: 768px) {
    [data-testid="stMetricValue"] { font-size: 1.05rem !important; }
    h1 { font-size: 1.35rem !important; line-height: 1.25 !important; }
    h2, h3 { font-size: 1.05rem !important; }
    [data-testid="stSlider"] label p { font-size: 0.9rem !important; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        min-width: min(100%, 320px) !important;
        flex: 1 1 auto !important;
    }
    [data-testid="stTabs"] button { padding-left: 0.35rem !important; padding-right: 0.35rem !important; font-size: 0.78rem !important; }
}
</style>
""",
    unsafe_allow_html=True,
)


# ─── DATA LOADING ─────────────────────────────────────────────────────────────
def _parse_cell_value(v: Any) -> float:
    """Parse numbers from CSV/Excel export (EU decimals, thousands, $, k suffix)."""
    if pd.isna(v):
        return np.nan
    if isinstance(v, (int, np.integer, float, np.floating)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip().replace("\u00a0", "").replace("$", "").replace("%", "")
    if not s or s.lower() in ("nan", "na", "n/a", "-"):
        return np.nan
    mult = 1.0
    if s.lower().endswith("k"):
        mult = 1000.0
        s = s[:-1].strip()
    s_clean = s.replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(?:,\d{3})+", s_clean):
        return float(s_clean.replace(",", "")) * mult
    if "," in s_clean and "." in s_clean:
        s2 = s_clean.replace(".", "").replace(",", ".")
    elif "," in s_clean:
        s2 = s_clean.replace(",", ".")
    else:
        s2 = s_clean
    try:
        return float(s2) * mult
    except ValueError:
        return np.nan


@st.cache_data
def load_data() -> Tuple[pd.DataFrame, List[str], Optional[str]]:
    if not DATA_PATH.is_file():
        return pd.DataFrame(), [], f"Data file not found: {DATA_PATH}"

    try:
        df_raw = pd.read_csv(DATA_PATH, header=None, encoding="utf-8-sig", engine="python")
    except Exception as e:
        return pd.DataFrame(), [], f"Load error: {e}"

    header_idx = -1
    for i in range(len(df_raw)):
        if str(df_raw.iloc[i, 2]).strip() == "Auckland":
            header_idx = i
            break
    if header_idx == -1:
        return pd.DataFrame(), [], "City header row not found (expected 'Auckland' in column 3)."

    country_header_idx = -1
    for i in range(header_idx):
        if str(df_raw.iloc[i, 1]).strip().lower() == "country":
            country_header_idx = i
            break

    header_row = df_raw.iloc[header_idx]
    city_map: Dict[int, str] = {}
    for col_idx in range(2, len(header_row)):
        val = str(header_row.iloc[col_idx]).strip()
        if val and val.lower() not in ("nan", "city", "country", "category", "spendings"):
            city_map[col_idx] = val
    cities_list = list(city_map.values())

    skip = {
        "city", "country", "category", "spendings", "", "nan",
        "population, kk", "land area, km^2", "last update",
    }

    parsed_rows: List[Dict[str, Any]] = []
    for row_idx in range(header_idx + 1, len(df_raw)):
        row = df_raw.iloc[row_idx]
        metric_name = str(row.iloc[1]).strip()
        if not metric_name or metric_name.lower() in skip:
            continue
        new_row: Dict[str, Any] = {"Metric": metric_name}
        for col_idx, city in city_map.items():
            new_row[city] = _parse_cell_value(row.iloc[col_idx])
        parsed_rows.append(new_row)

    df_city = pd.DataFrame(parsed_rows)
    city_metric_names = set(df_city["Metric"])

    if country_header_idx >= 0:
        ch = df_raw.iloc[country_header_idx]
        country_cols = [
            (j, str(ch.iloc[j]).strip())
            for j in range(2, len(ch))
            if str(ch.iloc[j]).strip()
            and str(ch.iloc[j]).strip().lower() not in ("nan", "country")
        ]
        for row_idx in range(country_header_idx + 1, header_idx):
            row = df_raw.iloc[row_idx]
            metric_name = str(row.iloc[1]).strip()
            if not metric_name or metric_name.lower() in skip:
                continue
            if metric_name in city_metric_names:
                continue
            country_vals = {lbl: _parse_cell_value(row.iloc[j]) for j, lbl in country_cols}
            new_row = {"Metric": metric_name}
            for city in city_map.values():
                cc = CITY_COUNTRY.get(city)
                new_row[city] = country_vals.get(cc, np.nan) if cc else np.nan
            parsed_rows.append(new_row)

    return pd.DataFrame(parsed_rows), cities_list, None


df, cities_list, _load_err = load_data()
if _load_err:
    st.error(_load_err)
if df.empty:
    st.error("Could not load data.")
    st.stop()

metric_df = df.set_index("Metric", verify_integrity=True)


# ─── HELPERS ──────────────────────────────────────────────────────────────────
def get_row(metric: str, cities: List[str]) -> pd.Series:
    if metric not in metric_df.index:
        return pd.Series({c: np.nan for c in cities}, dtype=float)
    return pd.to_numeric(metric_df.loc[metric].reindex(cities), errors="coerce")

def fmt_usd(v):
    return f"${v:,.1f}" if pd.notna(v) and not np.isnan(v) else "—"

def fmt_pct(v):
    return f"{v:.1f}%" if pd.notna(v) and not np.isnan(v) else "—"


def build_saving_metric_name(
    people: int, expenses_all: bool, income_key: str
) -> str:
    """CSV metric for precomputed monthly net saving (exact strings from export)."""
    pool = "all" if expenses_all else "base"
    if people == 1:
        sal_part = "Avg Dev salary" if income_key == "dev" else "Avg salary"
        return f"Saving {pool} 1 people on {sal_part}, USD/mo "
    sal_part = "Avg Dev + Des salary" if income_key == "dev_des" else "Avg Dev salary"
    return f"Saving {pool} 2 people on {sal_part}, USD/mo "


# ─── SIDEBAR (таймлайн и план накоплений) ────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Parameters")
    st.subheader("Timeline")
    start_year = st.number_input("Start year", 2024, 2035, 2026)
    prep_years = st.slider("Years to prepare", 1, 15, 5)
    target_year = start_year + prep_years

    st.subheader("Savings Plan")
    monthly_save = st.number_input("Monthly savings ($)", 0, 50_000, 2_000, step=100)
    growth = st.slider("Savings growth rate (% / yr)", 0, 50, 10)
    st.divider()
    st.caption("Города и профиль (переключатели) — в основной колонке.")
    st.caption(f"Data: 09.05.2026 · {len(cities_list)} cities · {len(metric_df)} metrics")

# ─── CITY SELECTOR ────────────────────────────────────────────────────────────
st.title("🌍 Global Relocation Intelligence 2026")
st.caption(f"{len(cities_list)} cities · {len(metric_df)} metrics · Updated May 2026")

defaults = [c for c in cities_list if any(n in c for n in
            ["Auckland", "Warsaw", "Toronto", "Amsterdam", "Belgrade"])][:5]
selected = st.multiselect("Select cities to compare:", options=cities_list, default=defaults)
if not selected:
    st.info("Choose at least one city.")
    st.stop()

# ─── ПРОФИЛЬ (segmented control в основной колонке) ───────────────────────────
st.subheader("Профиль расчёта")
st.caption("Переключатели задают строки Total / Saving из CSV.")
pc1, pc2, pc3 = st.columns((1, 1, 1))
with pc1:
    family_size = st.segmented_control(
        "Человек в семье",
        options=[1, 2],
        default=1,
        format_func=lambda x: "Один" if x == 1 else "Пара",
        key="prof_family",
        required=True,
    )
with pc2:
    exp_mode = st.segmented_control(
        "Расходы",
        options=["Base", "All"],
        default="Base",
        key="prof_exp",
        required=True,
    )
    expenses_all = exp_mode == "All"
with pc3:
    if family_size == 1:
        inc_mode = st.segmented_control(
            "Зарплата",
            options=["avg", "dev"],
            default="dev",
            format_func=lambda x: "Средняя" if x == "avg" else "IT (Dev)",
            key="prof_inc_1",
            required=True,
        )
        income_key: str = str(inc_mode)
    else:
        inc_mode = st.segmented_control(
            "Доход пары",
            options=["dev", "dev_des"],
            default="dev",
            format_func=lambda x: "Оба IT" if x == "dev" else "IT + Designer",
            key="prof_inc_2",
            required=True,
        )
        income_key = str(inc_mode)
        st.caption("В CSV для пары только Dev / Dev+Des.")

total_label = "All" if expenses_all else "Base"
exp_metric = f"Total {total_label} / {family_size} people, USD/mo"
saving_metric = build_saving_metric_name(family_size, expenses_all, income_key)
sal_metric = (
    "Avg Salary Dev, USD/mo"
    if (family_size == 2 or income_key == "dev")
    else "Avg Salary, USD/mo"
)
st.caption(f"Активная строка **Saving** в CSV: `{saving_metric.strip()}`")

st.divider()

# precompute savings accumulation
years_range = list(range(start_year, target_year + 1))
accum: List[float] = []
balance = 0.0
annual_contribution = monthly_save * 12
growth_factor = 1.0 + growth / 100.0
for _ in years_range:
    balance += annual_contribution
    accum.append(balance)
    annual_contribution *= growth_factor
final_saved = accum[-1] if accum else 0.0

tabs = st.tabs([
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
# TAB 0 – OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.subheader("At-a-Glance Snapshot")

    sal    = get_row(sal_metric, selected)
    exp    = get_row(exp_metric, selected)
    sav_ds = get_row(saving_metric, selected)
    rent   = get_row("1 bed. ap. in city centre", selected)
    qol    = get_row("Quality of life", selected)
    safe   = get_row("Safety", selected)
    exp_lbl = "Все расходы" if expenses_all else "Базовые расходы"

    for city in selected:
        s_v = sal.get(city, np.nan)
        e_v = exp.get(city, np.nan)
        r_v = rent.get(city, np.nan)
        saving = sav_ds.get(city, np.nan)
        if pd.isna(saving) and pd.notna(s_v) and pd.notna(e_v):
            saving = s_v - e_v
        ratio = saving / s_v * 100 if pd.notna(saving) and s_v > 0 else np.nan
        with st.expander(f"**{city}**", expanded=(len(selected) <= 4)):
            r1, r2, r3 = st.columns(3)
            r4, r5, r6 = st.columns(3)
            r1.metric("💰 Salary", fmt_usd(s_v))
            r2.metric(f"💳 {exp_lbl}", fmt_usd(e_v))
            r3.metric(
                "📦 Net Saving/mo",
                fmt_usd(saving),
                delta=fmt_pct(ratio) + " of income" if pd.notna(ratio) else "",
            )
            r4.metric("🏠 Rent 1BR", fmt_usd(r_v))
            r5.metric("🌟 Quality/Life", f"{qol.get(city, 0):.1f}/250")
            r6.metric("🛡️ Safety", f"{safe.get(city, 0):.1f}/100")

    # Radar
    st.markdown("### 🕸️ Multi-Dimension Radar")
    radar_metrics = ["Quality of life", "Safety", "Purch, power",
                     "Health care", "Climate", "Internet Speed (Mbps)"]
    radar_labels  = ["Quality of Life", "Safety", "Purchasing Power",
                     "Healthcare", "Climate", "Internet Speed"]
    radar_maxes   = [250, 100, 200, 100, 100, 350]

    fig_radar = go.Figure()
    colors = px.colors.qualitative.Set2
    for i, city in enumerate(selected):
        vals = [get_row(m, selected).get(city, 0) or 0 for m in radar_metrics]
        norm = [min(v / m * 100, 100) for v, m in zip(vals, radar_maxes)]
        fig_radar.add_trace(go.Scatterpolar(
            r=norm + [norm[0]], theta=radar_labels + [radar_labels[0]],
            fill='toself', name=city, line_color=colors[i % len(colors)],
        ))
    fig_radar.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])), height=520)
    st.plotly_chart(fig_radar, use_container_width=True, config=PLOTLY_CONFIG)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – FINANCES
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("💰 Income vs Expenses")

    sal_avg = get_row("Avg Salary, USD/mo", selected)
    sal_dev = get_row("Avg Salary Dev, USD/mo", selected)
    exp_b1  = get_row("Total Base / 1 people, USD/mo", selected)
    exp_b2  = get_row("Total Base / 2 people, USD/mo", selected)
    exp_a1  = get_row("Total All / 1 people, USD/mo", selected)
    exp_a2  = get_row("Total All / 2 people, USD/mo", selected)

    fig_fe = go.Figure()
    fig_fe.add_bar(x=selected, y=[sal_avg[c] for c in selected], name="Avg Salary",       marker_color="#90caf9")
    fig_fe.add_bar(x=selected, y=[sal_dev[c] for c in selected], name="Dev Salary",       marker_color="#1565c0")
    fig_fe.add_bar(x=selected, y=[exp_b1[c]  for c in selected], name="Base Exp (solo)",  marker_color="#ef9a9a")
    fig_fe.add_bar(x=selected, y=[exp_b2[c]  for c in selected], name="Base Exp (couple)",marker_color="#b71c1c")
    fig_fe.update_layout(barmode='group', yaxis_tickformat="$,.1f",
                         title="Salary vs Base Monthly Expenses", height=460)
    st.plotly_chart(fig_fe, use_container_width=True, config=PLOTLY_CONFIG)

    # Savings scenarios — long format (px.bar wide DF без x/y часто не рисуется)
    st.markdown("### 💧 Monthly Net Savings — все сценарии из таблицы")
    sav_scenarios = {
        "1·Base·Avg": "Saving base 1 people on Avg salary, USD/mo ",
        "1·Base·Dev": "Saving base 1 people on Avg Dev salary, USD/mo ",
        "1·All·Avg": "Saving all 1 people on Avg salary, USD/mo ",
        "1·All·Dev": "Saving all 1 people on Avg Dev salary, USD/mo ",
        "2·Base·Dev": "Saving base 2 people on Avg Dev salary, USD/mo ",
        "2·Base·Dev+Des": "Saving base 2 people on Avg Dev + Des salary, USD/mo ",
        "2·All·Dev": "Saving all 2 people on Avg Dev salary, USD/mo ",
        "2·All·Dev+Des": "Saving all 2 people on Avg Dev + Des salary, USD/mo ",
    }
    sav_df = pd.DataFrame(
        {lbl: get_row(m, selected) for lbl, m in sav_scenarios.items()},
        index=selected,
    )
    sav_long = sav_df.reset_index(names="City").melt(
        id_vars="City", var_name="Scenario", value_name="USD_mo"
    )
    scenario_order = list(sav_scenarios.keys())
    fig_sav = px.bar(
        sav_long,
        x="City",
        y="USD_mo",
        color="Scenario",
        barmode="group",
        category_orders={"Scenario": scenario_order},
        title="Monthly net savings — all 8 CSV scenarios (USD/mo)",
        color_discrete_sequence=px.colors.qualitative.Set2,
    )
    fig_sav.add_hline(y=0, line_dash="dash", line_color="red", opacity=0.6)
    _chart_h = min(720, max(380, 140 + len(selected) * 36))
    fig_sav.update_layout(
        yaxis_tickformat="$,.1f",
        height=_chart_h,
        xaxis_title="",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.28,
            xanchor="left",
            x=0,
            font=dict(size=11),
        ),
        margin=dict(l=10, r=10, t=48, b=140),
        xaxis=dict(tickangle=-42, automargin=True),
        yaxis=dict(automargin=True, title=""),
        bargap=0.12,
        bargroupgap=0.06,
    )
    st.plotly_chart(fig_sav, use_container_width=True, config=PLOTLY_CONFIG)

    sel_lbl = next((k for k, v in sav_scenarios.items() if v == saving_metric), saving_metric.strip())
    st.caption(f"**Выбранный профиль:** `{sel_lbl}` → `{saving_metric.strip()}`")

    # Purchasing power
    st.markdown("### 💪 Purchasing Power & Tax")
    col1, col2 = st.columns(2)
    with col1:
        pp = get_row("Purch, power", selected)
        fig_pp = px.bar(x=selected, y=[pp[c] for c in selected],
                        title="Purchasing Power Index",
                        color=[pp[c] for c in selected],
                        color_continuous_scale="Blues")
        fig_pp.update_layout(showlegend=False, height=380)
        st.plotly_chart(fig_pp, use_container_width=True, config=PLOTLY_CONFIG)
    with col2:
        tax_df = df[df["Metric"].isin(["salary tax, %", "VAT, %"])][["Metric"] + selected]
        if not tax_df.empty:
            tax_melt = tax_df.melt(id_vars="Metric", var_name="City", value_name="Rate")
            tax_melt["Rate"] = pd.to_numeric(tax_melt["Rate"], errors='coerce')
            fig_tax = px.bar(tax_melt, x="City", y="Rate", color="Metric", barmode="group",
                             title="Salary Tax & VAT (%)",
                             color_discrete_map={"salary tax, %": "#ef5350", "VAT, %": "#ffa726"})
            fig_tax.update_layout(height=380)
            st.plotly_chart(fig_tax, use_container_width=True, config=PLOTLY_CONFIG)

    # Full financial table
    st.markdown("### 📋 Financial Summary Table")
    fin_table = pd.DataFrame(
        {
            "Avg Salary": sal_avg,
            "Dev Salary": sal_dev,
            "Base Exp (1)": exp_b1,
            "All Exp (1)": exp_a1,
            "Base Exp (2)": exp_b2,
            "All Exp (2)": exp_a2,
            "Net save (профиль)": get_row(saving_metric, selected),
            "Purch. Power": get_row("Purch, power", selected),
        },
        index=selected,
    )
    money_cols = [c for c in fin_table.columns if c != "Purch. Power"]
    st.dataframe(
        fin_table.round(1)
        .style
        .format("${:,.1f}", subset=list(money_cols))
        .format("{:.1f}", subset=["Purch. Power"])
        .background_gradient(cmap="RdYlGn", axis=0),
        use_container_width=True
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – HOUSING
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("🏠 Housing Market")

    r1c   = get_row("1 bed. ap. in city centre", selected)
    r1o   = get_row("1 bed. ap. out. city centre", selected)
    r3c   = get_row("3 bed. ap. in city centre", selected)
    r3o   = get_row("3 bed. ap. out. city centre", selected)
    buy_c = get_row("Buy Apartment Price in centre, 80m*m / USD", selected)
    buy_o = get_row("Buy Apartment Price out. centre, 80m*m / USD", selected)
    sqm_c = get_row("USD/m*m in city centre", selected)
    sqm_o = get_row("USD/m*m out. city centre", selected)
    rate  = get_row("Interest Rate (20-Year Fixed, in %)", selected)
    pti   = get_row("Prop, to income", selected)

    col1, col2 = st.columns(2)
    with col1:
        fig_rent = go.Figure()
        fig_rent.add_bar(x=selected, y=[r1c[c] for c in selected], name="1BR Centre",  marker_color="#42a5f5")
        fig_rent.add_bar(x=selected, y=[r1o[c] for c in selected], name="1BR Suburbs", marker_color="#90caf9")
        fig_rent.add_bar(x=selected, y=[r3c[c] for c in selected], name="3BR Centre",  marker_color="#1565c0")
        fig_rent.add_bar(x=selected, y=[r3o[c] for c in selected], name="3BR Suburbs", marker_color="#4fc3f7")
        fig_rent.update_layout(barmode='group', yaxis_tickformat="$,.1f",
                               title="Monthly Rent by Type", height=420)
        st.plotly_chart(fig_rent, use_container_width=True, config=PLOTLY_CONFIG)
    with col2:
        fig_buy = go.Figure()
        fig_buy.add_bar(x=selected, y=[buy_c[c] for c in selected], name="80m² Centre",  marker_color="#ef5350")
        fig_buy.add_bar(x=selected, y=[buy_o[c] for c in selected], name="80m² Suburbs", marker_color="#ffcdd2")
        fig_buy.update_layout(barmode='group', yaxis_tickformat="$,.1f",
                              title="Purchase Price (80 m²)", height=420)
        st.plotly_chart(fig_buy, use_container_width=True, config=PLOTLY_CONFIG)

    # Price per m²
    st.markdown("### 📐 Price per m²")
    fig_sqm = go.Figure()
    fig_sqm.add_bar(x=selected, y=[sqm_c[c] for c in selected], name="$/m² Centre",  marker_color="#ab47bc")
    fig_sqm.add_bar(x=selected, y=[sqm_o[c] for c in selected], name="$/m² Suburbs", marker_color="#e1bee7")
    fig_sqm.update_layout(barmode='group', yaxis_tickformat="$,.1f", title="Price per m²", height=380)
    st.plotly_chart(fig_sqm, use_container_width=True, config=PLOTLY_CONFIG)

    # Mortgage calculator
    st.markdown("### 🏦 Mortgage Calculator (80 m², 20% down, 20-yr fixed)")
    mort_rows = []
    for city in selected:
        price = buy_c.get(city, np.nan)
        r_ann = rate.get(city, np.nan)
        if pd.isna(price) or pd.isna(r_ann):
            continue
        loan = price * 0.80
        r_mo = r_ann / 100 / 12
        n = 20 * 12
        pmt = loan * r_mo * (1 + r_mo)**n / ((1 + r_mo)**n - 1) if r_mo > 0 else loan / n
        sal_v = get_row(sal_metric, selected).get(city, np.nan)
        pct   = pmt / sal_v * 100 if pd.notna(sal_v) and sal_v > 0 else np.nan
        yrs   = price / (sal_v * 12) if pd.notna(sal_v) and sal_v > 0 else np.nan
        mort_rows.append({
            "City"            : city,
            "Price (80m²)"    : price,
            "Rate (%)"        : r_ann,
            "Monthly Payment" : pmt,
            "% of Salary"     : pct,
            "Years to Pay Off": yrs,
        })
    if mort_rows:
        mdf = pd.DataFrame(mort_rows).set_index("City")
        st.dataframe(
            mdf.style
            .format({"Price (80m²)": "${:,.1f}", "Monthly Payment": "${:,.1f}",
                     "Rate (%)": "{:.2f}%", "% of Salary": "{:.1f}%", "Years to Pay Off": "{:.1f}"})
            .background_gradient(cmap="RdYlGn_r", subset=["% of Salary", "Years to Pay Off"]),
            use_container_width=True
        )

    # Property-to-income ratio
    st.markdown("### 📊 Property-to-Income Ratio")
    fig_pti = px.bar(x=selected, y=[pti.get(c, 0) for c in selected],
                     title="Years of Salary to Buy 80m² Apt (centre) — lower = more affordable",
                     color=[pti.get(c, 0) for c in selected],
                     color_continuous_scale="RdYlGn_r",
                     labels={"x": "City", "y": "Years"})
    fig_pti.update_layout(showlegend=False, height=380)
    st.plotly_chart(fig_pti, use_container_width=True, config=PLOTLY_CONFIG)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – COST OF LIVING
# ══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("🛒 Cost of Living Deep Dive")

    spend_cats = {
        "Rent 1BR"        : "1 bed. ap. in city centre",
        "Groceries"       : "Groceries, USD/wk * 4,3",
        "Transport"       : "Transport, USD/mo",
        "Restaurants"     : "Restaurant, USD/mo",
        "Utilities"       : "Utilites, USD/mo",
        "Sports/Leisure"  : "Sports and Leisure",
        "Clothing"        : "Clothing and Shoes, 2 pairs/year / 12",
    }
    spend_df = pd.DataFrame({cat: get_row(m, selected) for cat, m in spend_cats.items()}, index=selected)
    fig_stack = px.bar(spend_df, barmode='stack',
                       title="Monthly Spending Breakdown by Category ($)",
                       color_discrete_sequence=px.colors.qualitative.Set3)
    fig_stack.update_layout(yaxis_tickformat="$,.1f", height=520)
    st.plotly_chart(fig_stack, use_container_width=True, config=PLOTLY_CONFIG)

    # Grocery basket heatmap
    st.markdown("### 🧺 Grocery Basket Prices (USD)")
    grocery_items = {
        "Milk 1L"          : "milk, l",
        "Bread 0.5kg"      : "white bread,  0,5kg",
        "Eggs 12"          : "eggs, 12",
        "Cheese kg"        : "cheese, kg",
        "Chicken kg"       : "chicken filltets, kg",
        "Beef kg"          : "beef, kg",
        "Apples kg"        : "apples, kg",
        "Potatoes kg"      : "potatoes, kg",
        "Water 1.5L"       : "bottled of watter, 1,5l",
        "Tomatoes kg"      : "tomatoes, kg",
        "Bananas kg"       : "bananas, kg",
        "Rice kg"          : "rice, kg",
        "Cappuccino"       : "cappuccino",
        "McCombo"          : "combo at mcdonald's",
    }
    groc_df = pd.DataFrame({item: get_row(m, selected) for item, m in grocery_items.items()}, index=selected).T
    fig_groc = px.imshow(groc_df, color_continuous_scale="RdYlGn_r",
                         title="Grocery Price Heatmap (USD)", aspect="auto")
    fig_groc.update_layout(height=480)
    st.plotly_chart(fig_groc, use_container_width=True, config=PLOTLY_CONFIG)

    # Transport + car prices
    st.markdown("### 🚌 Transport & Cars")
    col1, col2 = st.columns(2)
    with col1:
        trans_data = {
            "Monthly Pass"  : get_row("public transport pass, mo", selected),
            "One-way Ticket": get_row("one way ticket", selected),
            "Gasoline/L"    : get_row("gasoline, l", selected),
            "Taxi/km"       : get_row("taxi, 1km", selected),
        }
        trans_df = pd.DataFrame(trans_data, index=selected)
        fig_trans = px.bar(trans_df[["Monthly Pass"]], title="Monthly Transport Pass ($)",
                           color_discrete_sequence=["#42a5f5"])
        fig_trans.update_layout(height=360, yaxis_tickformat="$,.1f")
        st.plotly_chart(fig_trans, use_container_width=True, config=PLOTLY_CONFIG)
    with col2:
        car_data = {
            "VW Golf 1.5"    : get_row("VF Golf 1,5", selected),
            "Toyota Corolla" : get_row("Toyoyta Corolla Sedan 1,6", selected),
        }
        car_df = pd.DataFrame(car_data, index=selected)
        fig_car = px.bar(car_df, barmode='group', title="New Car Prices (USD)",
                         color_discrete_sequence=["#26a69a", "#7e57c2"])
        fig_car.update_layout(height=360, yaxis_tickformat="$,.1f")
        st.plotly_chart(fig_car, use_container_width=True, config=PLOTLY_CONFIG)

    # Indices
    st.markdown("### 📊 Numbeo Indices (New York = 100)")
    idx_data = {
        "Cost of Living" : get_row("Cost of living", selected),
        "Rent Index"     : get_row("Rent Index", selected),
        "Groceries Index": get_row("Groceries Index", selected),
    }
    idx_df = pd.DataFrame(idx_data, index=selected)
    fig_idx = px.bar(idx_df, barmode='group',
                     color_discrete_sequence=["#26a69a", "#42a5f5", "#7e57c2"])
    fig_idx.update_layout(height=380, title="Numbeo Indices vs New York (100)")
    st.plotly_chart(fig_idx, use_container_width=True, config=PLOTLY_CONFIG)

    # Full spending table
    st.markdown("### 📋 Complete Spending Table")
    all_spend = {**spend_cats,
                 "Monthly Pass"   : "public transport pass, mo",
                 "Gasoline/L"     : "gasoline, l",
                 "Fitness Club/mo": "fitness club, mo",
                 "Cinema Ticket"  : "cinema ticket",
                 "Cappuccino"     : "cappuccino",
                 "McCombo"        : "combo at mcdonald's",
                 "Dinner for 2"   : "meal for two at a mid rest.",
                 }
    full_df = pd.DataFrame({cat: get_row(m, selected) for cat, m in all_spend.items()}, index=selected)
    st.dataframe(full_df.round(1).style.background_gradient(cmap="RdYlGn_r", axis=0),
                 use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – QUALITY OF LIFE
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("🌿 Quality of Life & Environment")

    qol_items = {
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
    qol_df = pd.DataFrame({k: get_row(m, selected) for k, m in qol_items.items()}, index=selected)
    fig_hm = px.imshow(qol_df.T, color_continuous_scale="RdYlGn",
                       title="Quality of Life Heatmap", aspect="auto")
    fig_hm.update_layout(height=440)
    st.plotly_chart(fig_hm, use_container_width=True, config=PLOTLY_CONFIG)

    col1, col2 = st.columns(2)
    with col1:
        fig_cr = px.bar(x=selected, y=[get_row("Crime index", selected).get(c, 0) for c in selected],
                        title="Crime Index (lower = safer)",
                        color=[get_row("Crime index", selected).get(c, 0) for c in selected],
                        color_continuous_scale="RdYlGn_r")
        fig_cr.update_layout(showlegend=False, height=360)
        st.plotly_chart(fig_cr, use_container_width=True, config=PLOTLY_CONFIG)
    with col2:
        fig_hc = px.bar(x=selected, y=[get_row("Health care", selected).get(c, 0) for c in selected],
                        title="Healthcare Index (higher = better)",
                        color=[get_row("Health care", selected).get(c, 0) for c in selected],
                        color_continuous_scale="RdYlGn")
        fig_hc.update_layout(showlegend=False, height=360)
        st.plotly_chart(fig_hc, use_container_width=True, config=PLOTLY_CONFIG)

    # Weather
    st.markdown("### ☀️ Climate & Weather")
    col1, col2 = st.columns(2)
    with col1:
        sun  = get_row("Sunny days, %", selected)
        rain = get_row("Rainy days, %", selected)
        fig_wr = go.Figure()
        fig_wr.add_bar(x=selected, y=[sun[c]  for c in selected], name="Sunny %",  marker_color="#ffd54f")
        fig_wr.add_bar(x=selected, y=[rain[c] for c in selected], name="Rainy %",  marker_color="#64b5f6")
        fig_wr.update_layout(barmode='group', title="Sunny vs Rainy Days (%)", height=380)
        st.plotly_chart(fig_wr, use_container_width=True, config=PLOTLY_CONFIG)
    with col2:
        summ = get_row("Avg Summer temp", selected)
        wint = get_row("Avg Winter temp", selected)
        fig_temp = go.Figure()
        fig_temp.add_bar(x=selected, y=[summ[c] for c in selected], name="Summer °C", marker_color="#ff7043")
        fig_temp.add_bar(x=selected, y=[wint[c] for c in selected], name="Winter °C", marker_color="#42a5f5")
        fig_temp.update_layout(barmode='group', title="Avg Temperatures (°C)", height=380)
        st.plotly_chart(fig_temp, use_container_width=True, config=PLOTLY_CONFIG)

    col1, col2 = st.columns(2)
    with col1:
        fig_pol = px.bar(x=selected, y=[get_row("Pollution", selected).get(c, 0) for c in selected],
                         title="Pollution Index (lower = cleaner)",
                         color=[get_row("Pollution", selected).get(c, 0) for c in selected],
                         color_continuous_scale="RdYlGn_r")
        fig_pol.update_layout(showlegend=False, height=360)
        st.plotly_chart(fig_pol, use_container_width=True, config=PLOTLY_CONFIG)
    with col2:
        fig_inet = px.bar(x=selected, y=[get_row("Internet Speed (Mbps)", selected).get(c, 0) for c in selected],
                          title="Internet Speed (Mbps)",
                          color=[get_row("Internet Speed (Mbps)", selected).get(c, 0) for c in selected],
                          color_continuous_scale="Blues")
        fig_inet.update_layout(showlegend=False, height=360)
        st.plotly_chart(fig_inet, use_container_width=True, config=PLOTLY_CONFIG)

    # Nature access
    st.markdown("### 🌲 Nature Access (hours by car)")
    nat_items = {"Ocean/Sea": "Ocean/Sea, h", "Forest": "Forest, h", "Mountains": "Mountains, h"}
    nat_df = pd.DataFrame({k: get_row(m, selected) for k, m in nat_items.items()}, index=selected)
    fig_nat = px.bar(nat_df, barmode='group', title="Hours to Nearest Nature Feature",
                     color_discrete_sequence=["#26a69a", "#66bb6a", "#8d6e63"])
    fig_nat.update_layout(height=380, yaxis_title="Hours")
    st.plotly_chart(fig_nat, use_container_width=True, config=PLOTLY_CONFIG)

    # Expat share
    st.markdown("### 🌐 Expat Community & UV")
    col1, col2 = st.columns(2)
    with col1:
        fig_imm = px.bar(x=selected, y=[get_row("Immigrants, %", selected).get(c, 0) for c in selected],
                         title="Immigrant Population (%)", color_discrete_sequence=["#7986cb"])
        fig_imm.update_layout(height=360)
        st.plotly_chart(fig_imm, use_container_width=True, config=PLOTLY_CONFIG)
    with col2:
        fig_uv = px.bar(x=selected, y=[get_row("Avg Peak UV index", selected).get(c, 0) for c in selected],
                        title="Peak UV Index",
                        color=[get_row("Avg Peak UV index", selected).get(c, 0) for c in selected],
                        color_continuous_scale="YlOrRd")
        fig_uv.update_layout(showlegend=False, height=360)
        st.plotly_chart(fig_uv, use_container_width=True, config=PLOTLY_CONFIG)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – IMMIGRATION
# ══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("🛂 Immigration & Residency Pathways")

    imm_items = {
        "PR Fast Track (yrs)"    : "PR (Years) Fast Track",
        "PR General (yrs)"       : "PR (Years) General Path",
        "Citizenship Fast (yrs)" : "Citizenship (Years) Fast",
        "Citizenship Real (yrs)" : "Citizenship (Years) Real",
        "Passport Difficulty"    : "BY Passport Difficulty",
        "IELTS Required"         : "IELTS / Language",
        "1yr Buffer ($k)"        : "1 Year Buffer (USD), $k",
        "Salary 4-5y ($k)"       : "Salary (4-5y exp), $k",
    }
    imm_df = pd.DataFrame({k: get_row(m, selected) for k, m in imm_items.items()}, index=selected)

    fig_imm_hm = px.imshow(imm_df.T, title="Immigration Metrics Heatmap",
                           color_continuous_scale="RdYlGn_r", aspect="auto")
    fig_imm_hm.update_layout(height=420)
    st.plotly_chart(fig_imm_hm, use_container_width=True, config=PLOTLY_CONFIG)

    st.markdown("### ⏱️ Residency Timeline (years)")
    fig_tl = go.Figure()
    for label, metric, color in [
        ("PR Fast Track", "PR (Years) Fast Track", "#66bb6a"),
        ("PR General",    "PR (Years) General Path", "#ffa726"),
        ("Citizenship Fast", "Citizenship (Years) Fast", "#42a5f5"),
        ("Citizenship Real", "Citizenship (Years) Real", "#ef5350"),
    ]:
        vals = [get_row(metric, selected).get(c, 0) for c in selected]
        fig_tl.add_bar(x=selected, y=vals, name=label, marker_color=color)
    fig_tl.update_layout(barmode='group', title="Years to PR / Citizenship", height=440)
    st.plotly_chart(fig_tl, use_container_width=True, config=PLOTLY_CONFIG)

    # Buffer analysis
    st.markdown("### 🛡️ Financial Buffer: Required vs Your Savings")
    buf = get_row("1 Year Buffer (USD), $k", selected) * 1000
    fig_buf_cmp = go.Figure()
    fig_buf_cmp.add_bar(x=selected, y=[buf.get(c, 0) for c in selected],
                        name="Required Buffer", marker_color="#ef5350")
    fig_buf_cmp.add_bar(x=selected, y=[final_saved] * len(selected),
                        name=f"Your Savings by {target_year}", marker_color="#66bb6a")
    fig_buf_cmp.update_layout(barmode='group', yaxis_tickformat="$,.1f",
                              title="1-Year Buffer Required vs Your Projected Savings", height=420)
    st.plotly_chart(fig_buf_cmp, use_container_width=True, config=PLOTLY_CONFIG)

    st.markdown("### 📋 Immigration Summary")
    st.dataframe(
        imm_df.style.background_gradient(
            cmap="RdYlGn_r",
            subset=["PR Fast Track (yrs)", "PR General (yrs)", "Citizenship Fast (yrs)",
                    "Citizenship Real (yrs)", "Passport Difficulty"]),
        use_container_width=True
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 – SAVINGS PLAN
# ══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("📈 Personal Savings & Relocation Plan")

    st.info(f"**Plan:** ${monthly_save:,.1f}/mo · growing {growth}%/yr · "
            f"{prep_years} years ({start_year}→{target_year}) · **Target: ${final_saved:,.1f}**")

    fig_acc = px.area(x=years_range, y=accum,
                      title="Accumulated Savings Over Time",
                      labels={"x": "Year", "y": "Total Saved ($)"},
                      color_discrete_sequence=["#42a5f5"])
    fig_acc.update_layout(yaxis_tickformat="$,.1f", height=420)
    st.plotly_chart(fig_acc, use_container_width=True, config=PLOTLY_CONFIG)

    # Coverage analysis
    st.markdown(f"### 🎯 Your savings by {target_year}: **${final_saved:,.1f}**")
    buf = get_row("1 Year Buffer (USD), $k", selected) * 1000
    exp_1 = get_row(exp_metric, selected)
    sav_row = get_row(saving_metric, selected)

    comp_rows = []
    for city in selected:
        buf_v = buf.get(city, np.nan)
        exp_v = exp_1.get(city, np.nan)
        sal_v = get_row(sal_metric, selected).get(city, np.nan)
        net_mo = sav_row.get(city, np.nan)
        if pd.isna(net_mo) and pd.notna(sal_v) and pd.notna(exp_v):
            net_mo = sal_v - exp_v
        months = final_saved / exp_v if pd.notna(exp_v) and exp_v > 0 else np.nan
        cov_pct= final_saved / buf_v * 100 if pd.notna(buf_v) and buf_v > 0 else np.nan
        # break-even: months until savings restored after move
        break_even = (buf_v - final_saved) / net_mo if (pd.notna(net_mo) and net_mo > 0
                     and pd.notna(buf_v) and buf_v > final_saved) else 0
        comp_rows.append({
            "City"               : city,
            "Buffer Required"    : buf_v,
            "Your Savings"       : final_saved,
            "Coverage %"         : cov_pct,
            "Months of Cover"    : months,
            "Net Save/mo"        : net_mo,
            "Break-even (mo)"    : max(break_even, 0),
        })
    comp_df = pd.DataFrame(comp_rows).set_index("City")
    st.dataframe(
        comp_df.style
        .format({"Buffer Required": "${:,.1f}", "Your Savings": "${:,.1f}",
                 "Coverage %": "{:.1f}%", "Months of Cover": "{:.1f}",
                 "Net Save/mo": "${:,.1f}", "Break-even (mo)": "{:.1f}"})
        .background_gradient(cmap="RdYlGn", subset=["Coverage %", "Months of Cover"])
        .background_gradient(cmap="RdYlGn_r", subset=["Break-even (mo)"]),
        use_container_width=True
    )

    # Coverage bar chart
    fig_cov = px.bar(comp_df.reset_index(), x="City", y="Coverage %",
                     title=f"Buffer Coverage — ${final_saved:,.1f} vs Required",
                     color="Coverage %", color_continuous_scale="RdYlGn")
    fig_cov.add_hline(y=100, line_dash="dash", line_color="white",
                      annotation_text="100% target")
    fig_cov.update_layout(showlegend=False, height=400)
    st.plotly_chart(fig_cov, use_container_width=True, config=PLOTLY_CONFIG)

    # Long-term wealth projection in each city
    st.markdown("### 📊 5-Year Wealth Projection After Relocation")
    proj_years = list(range(target_year, target_year + 6))
    fig_proj = go.Figure()
    colors = px.colors.qualitative.Set2
    for i, city in enumerate(selected):
        sal_v = get_row(sal_metric, selected).get(city, np.nan)
        exp_v = exp_1.get(city, np.nan)
        net_mo = sav_row.get(city, np.nan)
        if pd.isna(net_mo) and pd.notna(sal_v) and pd.notna(exp_v):
            net_mo = sal_v - exp_v
        if pd.isna(sal_v) or pd.isna(exp_v) or pd.isna(net_mo):
            continue
        wealth = [final_saved + net_mo * 12 * yr for yr in range(len(proj_years))]
        fig_proj.add_scatter(x=proj_years, y=wealth, mode='lines+markers',
                             name=city, line_color=colors[i % len(colors)])
    fig_proj.update_layout(yaxis_tickformat="$,.1f",
                           title="Projected Cumulative Wealth (savings + on-site net income)",
                           height=450)
    st.plotly_chart(fig_proj, use_container_width=True, config=PLOTLY_CONFIG)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 – FINAL SCORE
# ══════════════════════════════════════════════════════════════════════════════
with tabs[7]:
    st.subheader("🏆 Composite Relocation Score")
    st.markdown("""
    **Weighted score (0–100):**
    Financial (savings rate + purchasing power) **30%** ·
    Quality of Life (QoL + safety + health + climate) **25%** ·
    Housing affordability (rent/salary + prop-to-income) **20%** ·
    Immigration ease (PR speed + passport + buffer) **15%** ·
    Environment (pollution + internet + expat %) **10%**
    """)

    score_rows = []
    for city in selected:
        def g(m): return get_row(m, selected).get(city, np.nan) or 0

        sal_v = get_row(sal_metric, selected).get(city, np.nan)
        exp_v = get_row(exp_metric, selected).get(city, np.nan)
        sav_v = get_row(saving_metric, selected).get(city, np.nan)
        if pd.notna(sav_v) and sal_v > 0:
            save_r = sav_v / sal_v * 100
        elif pd.notna(sal_v) and pd.notna(exp_v) and sal_v > 0:
            save_r = (sal_v - exp_v) / sal_v * 100
        else:
            save_r = 0
        pp_v = get_row("Purch, power", selected).get(city, np.nan) or 0
        fin_score = np.clip(save_r * 1.2, 0, 60) + np.clip((pp_v - 50) / 150 * 40, 0, 40)

        qol_v  = g("Quality of life"); safe_v = g("Safety")
        hc_v   = g("Health care");     clim_v = g("Climate")
        qol_score = (np.clip(qol_v / 220 * 100, 0, 100) * 0.40 +
                     np.clip(safe_v, 0, 100)             * 0.25 +
                     np.clip(hc_v, 0, 100)               * 0.20 +
                     np.clip(clim_v, 0, 100)              * 0.15)

        r1c_v = g("1 bed. ap. in city centre"); pti_v = g("Prop, to income")
        rent_ratio = max((1 - r1c_v / sal_v) * 100, 0) if sal_v > 0 else 50
        hs_score = (np.clip(rent_ratio, 0, 100) * 0.5 +
                    np.clip((20 - pti_v) / 17 * 100, 0, 100) * 0.5)

        pr_f   = g("PR (Years) Fast Track"); pass_d = g("BY Passport Difficulty")
        buf_v  = g("1 Year Buffer (USD), $k") * 1000
        buf_cov = np.clip(final_saved / buf_v * 100, 0, 100) if buf_v > 0 else 0
        imm_score = (np.clip((10 - pr_f) / 10 * 100, 0, 100) * 0.50 +
                     np.clip(pass_d / 5  * 100, 0, 100)       * 0.30 +
                     buf_cov                                   * 0.20)

        pol_v  = g("Pollution"); inet_v = g("Internet Speed (Mbps)"); imm_p = g("Immigrants, %")
        env_score = (np.clip((100 - pol_v), 0, 100)    * 0.40 +
                     np.clip(inet_v / 350 * 100, 0, 100)* 0.30 +
                     np.clip(imm_p, 0, 100)              * 0.30)

        composite = (fin_score  * 0.30 + qol_score * 0.25 +
                     hs_score   * 0.20 + imm_score * 0.15 + env_score * 0.10)

        score_rows.append({
            "City"           : city,
            "Financial"      : round(fin_score, 1),
            "Quality of Life": round(qol_score, 1),
            "Housing"        : round(hs_score, 1),
            "Immigration"    : round(imm_score, 1),
            "Environment"    : round(env_score, 1),
            "⭐ Score"       : round(composite, 1),
        })

    score_df = (pd.DataFrame(score_rows)
                .set_index("City")
                .sort_values("⭐ Score", ascending=False))

    if len(score_df) > 0:
        top = score_df.index[0]
        st.success(f"### 🥇 Best match: **{top}**  —  Score: **{score_df.loc[top, '⭐ Score']:.1f} / 100**")

    # Horizontal bar
    fig_sc = px.bar(score_df.reset_index().sort_values("⭐ Score"),
                    x="⭐ Score", y="City", orientation='h',
                    color="⭐ Score", color_continuous_scale="RdYlGn",
                    title="Composite Score (higher = better match)")
    fig_sc.update_layout(showlegend=False, height=max(300, len(selected) * 55))
    st.plotly_chart(fig_sc, use_container_width=True, config=PLOTLY_CONFIG)

    # Stacked dimension contributions
    dim_cols = ["Financial", "Quality of Life", "Housing", "Immigration", "Environment"]
    weights  = [0.30, 0.25, 0.20, 0.15, 0.10]
    wt_df = score_df[dim_cols].copy()
    for col, w in zip(dim_cols, weights):
        wt_df[col] = wt_df[col] * w
    fig_dim = px.bar(wt_df.reset_index(), x="City", y=dim_cols, barmode='stack',
                     title="Score Contribution by Dimension (weighted)",
                     color_discrete_sequence=px.colors.qualitative.Set2)
    fig_dim.update_layout(yaxis_title="Weighted Points", height=460)
    st.plotly_chart(fig_dim, use_container_width=True, config=PLOTLY_CONFIG)

    # Full table
    st.markdown("### 📋 Score Details")
    st.dataframe(
        score_df.style.background_gradient(cmap="RdYlGn", vmin=0, vmax=100),
        use_container_width=True
    )

    # Per-city summary cards
    st.markdown("---")
    st.markdown("### 💡 City Profiles")
    for city_n, row_s in score_df.iterrows():
        rank = list(score_df.index).index(city_n) + 1
        medal = ["🥇", "🥈", "🥉"][rank - 1] if rank <= 3 else "🏅"
        sal_v = get_row(sal_metric, selected).get(city_n, 0)
        exp_v = get_row(exp_metric, selected).get(city_n, 0)
        net_ds = get_row(saving_metric, selected).get(city_n, np.nan)
        net = float(net_ds) if pd.notna(net_ds) else (sal_v - exp_v)
        with st.expander(f"{medal} #{rank} **{city_n}** — {row_s['⭐ Score']:.1f}/100"):
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Financial",       f"{row_s['Financial']:.1f}")
            c2.metric("Quality of Life", f"{row_s['Quality of Life']:.1f}")
            c3.metric("Housing",         f"{row_s['Housing']:.1f}")
            c4.metric("Immigration",     f"{row_s['Immigration']:.1f}")
            c5.metric("Environment",     f"{row_s['Environment']:.1f}")
            st.caption(
                f"Salary: {fmt_usd(sal_v)}/mo · Expenses: {fmt_usd(exp_v)}/mo · "
                f"Net saving: {fmt_usd(net)}/mo · "
                f"PR Fast Track: {get_row('PR (Years) Fast Track', selected).get(city_n, '?'):.1f} yrs"
            )
