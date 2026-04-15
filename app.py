"""
Alphavest — CSE Risk Analytics Dashboard  v0.2
===============================================
Streamlit app reading from pre-computed Excel outputs.
Data refreshed daily via automated pipeline.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Alphavest",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="stSidebar"] {
    background: #F8FAFC;
    border-right: 1px solid #E2E8F0;
  }
  [data-testid="stMetric"] {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 16px 20px;
  }
  [data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 700;
    color: #0F172A;
  }
  [data-testid="stMetricLabel"] {
    font-size: 0.78rem !important;
    color: #64748B;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  h1 { color: #0F172A; font-weight: 800; letter-spacing: -0.5px; }
  h2, h3 { color: #1E293B; font-weight: 700; }
  [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
  .block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ── Paths ─────────────────────────────────────────────────────────────────────
DATA_DIR       = Path(__file__).parent / "data"
STOCKS_FILE    = DATA_DIR / "risk_metrics_by_window_latest.xlsx"
PORTFOLIO_FILE = DATA_DIR / "portfolio_metrics_latest.xlsx"

# ── Plotly theme ───────────────────────────────────────────────────────────────
PLOT_BG  = "#FFFFFF"
PAPER_BG = "#FFFFFF"
GRID_COL = "#F1F5F9"
TEXT_COL = "#0F172A"
BLUE     = "#2563EB"
GREEN    = "#16A34A"
AMBER    = "#D97706"
RED      = "#DC2626"
MUTED    = "#94A3B8"

def base_layout(**kwargs):
    return dict(
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=PLOT_BG,
        font=dict(color=TEXT_COL, family="Inter, sans-serif", size=12),
        margin=dict(t=20, b=40, l=10, r=10),
        **kwargs
    )

AXIS_STYLE = dict(gridcolor=GRID_COL, linecolor="#E2E8F0", zeroline=False)

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def load_stocks():
    xl = pd.ExcelFile(STOCKS_FILE)
    sheets = {}
    for sheet in xl.sheet_names:
        if sheet == "Daily_Returns":
            continue
        df = xl.parse(sheet, index_col=0)
        sheets[sheet] = df
    return sheets

@st.cache_data(ttl=3600)
def load_returns():
    df = pd.read_excel(STOCKS_FILE, sheet_name="Daily_Returns")
    date_col = df.columns[0]
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
    df = df.dropna(subset=[date_col]).set_index(date_col)
    df.index.name = "Date"
    return df

@st.cache_data(ttl=3600)
def load_portfolio_returns():
    df = pd.read_excel(PORTFOLIO_FILE, sheet_name="Portfolio_Returns")
    df.columns = df.columns.str.strip()
    date_col = df.columns[0]
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
    df = df.dropna(subset=[date_col]).set_index(date_col)
    df.index.name = "Date"
    return df

@st.cache_data(ttl=3600)
def load_portfolio_metrics():
    return pd.read_excel(PORTFOLIO_FILE, sheet_name="All_Metrics", index_col=None)

metrics_data  = load_stocks()
returns_data  = load_returns()
portfolio_ret = load_portfolio_returns()
portfolio_all = load_portfolio_metrics()

# ── Global variables ──────────────────────────────────────────────────────────
WINDOWS       = list(list(metrics_data.values())[0].columns)
SHORT_WINDOWS = [w.split(" ")[0] for w in WINDOWS]
TICKERS       = [t for t in list(list(metrics_data.values())[0].index) if t != "MASI"]
TICKERS_CLEAN = [t.replace(".CS", "") for t in TICKERS]
METRICS       = list(metrics_data.keys())

METRIC_LABELS = {
    "Annualized_Volatility": "Volatility (ann.)",
    "Annualized_Return":     "Return (ann.)",
    "Sharpe_Ratio":          "Sharpe Ratio",
    "Beta":                  "Beta",
    "Tracking_Error":        "Tracking Error",
    "Alpha":                 "Jensen's Alpha",
    "Max_Drawdown":          "Max Drawdown",
    "VaR_95":                "VaR 95%",
    "CVaR_95":               "CVaR 95%",
    "Sortino_Ratio":         "Sortino Ratio",
}

PCT_METRICS = {"Annualized_Volatility", "Annualized_Return", "Tracking_Error",
               "Max_Drawdown", "VaR_95", "CVaR_95"}

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding: 8px 0 20px 0;'>
      <span style='font-size:1.5rem; font-weight:900; color:#2563EB; letter-spacing:-1px;'>
        ALPHAVEST
      </span><br>
      <span style='font-size:0.72rem; color:#64748B; letter-spacing:0.1em;'>
        CSE RISK ANALYTICS
      </span>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio(
        "Navigation",
        ["Market Overview", "Stock Analysis", "Portfolio",
         "Portfolio Optimizer", "Screener", "AI Assistant", "Workflows"],
        label_visibility="collapsed"
    )

    st.markdown(
        "<div style='margin:20px 0 8px 0;font-size:0.72rem;color:#94A3B8;"
        "text-transform:uppercase;letter-spacing:0.08em;'>Time Window</div>",
        unsafe_allow_html=True
    )
    window = st.selectbox("Time Window", WINDOWS, index=0, label_visibility="collapsed")
    win_idx = WINDOWS.index(window)

    end_label = WINDOWS[0].split("to")[-1].strip() if "to" in WINDOWS[0] else "latest"
    st.markdown(
        f"<div style='font-size:0.72rem;color:#94A3B8;margin-top:6px;'>"
        f"Data as of <b style='color:#475569;'>{end_label}</b></div>",
        unsafe_allow_html=True
    )
    st.markdown(
        "<div style='position:absolute;bottom:20px;font-size:0.68rem;color:#CBD5E1;'>"
        "v0.2 · Casablanca SE</div>",
        unsafe_allow_html=True
    )

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(val, metric):
    if pd.isna(val):
        return "n/a"
    if metric in PCT_METRICS:
        return f"{val:.2f}%"
    return f"{val:.3f}"

def metric_bar_colors(metric, values):
    if metric in {"Annualized_Return", "Sharpe_Ratio", "Alpha", "Sortino_Ratio"}:
        return [GREEN if v > 0 else RED for v in values]
    if metric in {"Max_Drawdown", "VaR_95", "CVaR_95"}:
        return [AMBER if v > -20 else RED for v in values]
    return [BLUE] * len(values)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — MARKET OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
if page == "Market Overview":
    st.title("Market Overview")
    st.caption("Casablanca Stock Exchange — 80 listed companies")

    # ── MASI KPI cards ────────────────────────────────────────────────────────
    def masi_val(metric):
        if "MASI" in metrics_data.get(metric, pd.DataFrame()).index:
            return metrics_data[metric].loc["MASI", window]
        return float("nan")

    masi_ret = masi_val("Annualized_Return")
    masi_vol = masi_val("Annualized_Volatility")
    masi_shr = masi_val("Sharpe_Ratio")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Stocks Tracked",   f"{len(TICKERS)}")
    c2.metric("MASI Ann. Return", f"{masi_ret:.1f}%" if not pd.isna(masi_ret) else "n/a")
    c3.metric("MASI Volatility",  f"{masi_vol:.1f}%" if not pd.isna(masi_vol) else "n/a")
    c4.metric("MASI Sharpe",      f"{masi_shr:.2f}"  if not pd.isna(masi_shr) else "n/a")

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

    # ── Metric Explorer bar chart ─────────────────────────────────────────────
    st.subheader("Metric Explorer")
    metric_choice = st.selectbox(
        "Select metric",
        METRICS,
        format_func=lambda m: METRIC_LABELS.get(m, m),
        key="metric_explorer"
    )

    stocks_only = metrics_data[metric_choice].dropna()
    stocks_only = stocks_only[stocks_only.index != "MASI"]
    ascending   = metric_choice in {"Max_Drawdown", "VaR_95", "CVaR_95",
                                    "Annualized_Volatility", "Tracking_Error"}
    ranked = stocks_only[window].sort_values(ascending=ascending).head(20)
    colors = metric_bar_colors(metric_choice, ranked.values)

    fig_bar = go.Figure(go.Bar(
        x=[t.replace(".CS", "") for t in ranked.index],
        y=ranked.values,
        marker_color=colors,
        marker_line_width=0,
        text=[fmt(v, metric_choice) for v in ranked.values],
        textposition="outside",
        textfont=dict(size=10),
    ))
    if metric_choice == "Sharpe_Ratio":
        fig_bar.add_hline(y=1, line_dash="dot", line_color=BLUE, opacity=0.5,
                          annotation_text="Sharpe = 1", annotation_position="top right",
                          annotation_font_color=BLUE)
    fig_bar.update_layout(
        **base_layout(height=360),
        xaxis={**AXIS_STYLE, "tickangle": -40},
        yaxis={**AXIS_STYLE, "title": METRIC_LABELS.get(metric_choice, metric_choice)},
        showlegend=False,
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── Return vs Risk + Heatmap ───────────────────────────────────────────────
    col_l, col_r = st.columns([1, 1])

    with col_l:
        st.subheader("Return vs Risk")
        scatter_df = pd.DataFrame({
            "Return":     metrics_data["Annualized_Return"][window],
            "Volatility": metrics_data["Annualized_Volatility"][window],
            "Sharpe":     metrics_data["Sharpe_Ratio"][window],
        }).dropna()
        scatter_df = scatter_df[scatter_df.index != "MASI"].copy()
        scatter_df["Ticker"] = [t.replace(".CS", "") for t in scatter_df.index]

        fig_sc = px.scatter(
            scatter_df, x="Volatility", y="Return",
            color="Sharpe", hover_name="Ticker",
            color_continuous_scale=[[0, RED], [0.4, AMBER], [1, GREEN]],
            labels={"Volatility": "Volatility (%)", "Return": "Return (%)"},
        )
        fig_sc.update_traces(marker=dict(size=7, line=dict(width=0.5, color="white")))
        fig_sc.update_layout(
            **base_layout(height=400),
            xaxis={**AXIS_STYLE, "title": "Volatility (%)"},
            yaxis={**AXIS_STYLE, "title": "Return (%)"},
            coloraxis_colorbar=dict(title="Sharpe", thickness=12, len=0.8),
        )
        st.plotly_chart(fig_sc, use_container_width=True)

    with col_r:
        st.subheader("Market Heatmap")
        hm_metrics = ["Annualized_Return", "Annualized_Volatility", "Sharpe_Ratio",
                      "Beta", "Max_Drawdown", "VaR_95"]
        hm_table = pd.DataFrame(index=TICKERS)
        for m in hm_metrics:
            if m in metrics_data:
                hm_table[METRIC_LABELS[m]] = metrics_data[m][window].reindex(TICKERS).round(2)
        hm_table.index = TICKERS_CLEAN

        styled = hm_table.style
        for col_name in hm_table.columns:
            cmap = "RdYlGn_r" if col_name in {"Volatility (ann.)", "Max Drawdown", "VaR 95%"} else "RdYlGn"
            styled = styled.background_gradient(cmap=cmap, subset=[col_name], axis=0)
        styled = styled.format(precision=2)
        st.dataframe(styled, use_container_width=True, height=400)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — STOCK ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Stock Analysis":
    st.title("Stock Analysis")

    ticker_label = st.selectbox("Select stock", TICKERS_CLEAN)
    ticker = ticker_label + ".CS"
    if ticker not in TICKERS:
        ticker = TICKERS[0]
        ticker_label = ticker.replace(".CS", "")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    col_left, col_right = st.columns([1, 2])

    # ── Left: metric cards ────────────────────────────────────────────────────
    with col_left:
        st.subheader("Risk Metrics")
        for metric in METRICS:
            if ticker in metrics_data[metric].index:
                val = metrics_data[metric].loc[ticker, window]
                label = METRIC_LABELS.get(metric, metric)
                if not pd.isna(val):
                    display = f"{val:.2f}%" if metric in PCT_METRICS else f"{val:.3f}"
                    st.metric(label, display)

    # ── Right: charts ─────────────────────────────────────────────────────────
    with col_right:
        st.subheader("Metrics Across Windows")
        rows = []
        for w, sw in zip(WINDOWS, SHORT_WINDOWS):
            row = {"Window": sw}
            for metric in ["Annualized_Return", "Sharpe_Ratio", "Annualized_Volatility",
                           "Max_Drawdown", "Beta", "VaR_95"]:
                if metric in metrics_data and ticker in metrics_data[metric].index:
                    val = metrics_data[metric].loc[ticker, w]
                    row[METRIC_LABELS[metric]] = round(val, 2) if metric in PCT_METRICS else round(val, 3)
            rows.append(row)
        window_df = pd.DataFrame(rows).set_index("Window")
        st.dataframe(window_df, use_container_width=True)

        if ticker in returns_data.columns:
            ret_s = returns_data[ticker].dropna() / 100

            # Rolling Volatility
            st.subheader("Rolling Volatility")
            fig_vol = go.Figure()
            for w_days, label, color in [(21, "1M", "#93C5FD"), (63, "3M", BLUE), (126, "6M", "#1E3A8A")]:
                rv = ret_s.rolling(w_days).std() * np.sqrt(252) * 100
                fig_vol.add_trace(go.Scatter(
                    x=rv.index, y=rv.values,
                    name=label, line=dict(color=color, width=1.5),
                ))
            fig_vol.update_layout(
                **base_layout(height=220),
                xaxis=AXIS_STYLE,
                yaxis={**AXIS_STYLE, "title": "Ann. Volatility (%)"},
                legend=dict(orientation="h", y=1.12, x=0),
            )
            st.plotly_chart(fig_vol, use_container_width=True)

            # Drawdown Chart
            st.subheader("Drawdown")
            cum     = (1 + ret_s).cumprod()
            run_max = cum.cummax()
            dd      = (cum - run_max) / run_max * 100

            fig_dd = go.Figure()
            fig_dd.add_trace(go.Scatter(
                x=dd.index, y=dd.values,
                line=dict(color=RED, width=1.5),
                fill="tozeroy",
                fillcolor="rgba(220,38,38,0.08)",
            ))
            max_dd_idx = dd.idxmin()
            fig_dd.add_annotation(
                x=max_dd_idx, y=float(dd.min()),
                text=f"Max DD: {dd.min():.1f}%",
                showarrow=True, arrowhead=2,
                font=dict(color=RED, size=10),
                bgcolor="white",
            )
            fig_dd.update_layout(
                **base_layout(height=200),
                xaxis=AXIS_STYLE,
                yaxis={**AXIS_STYLE, "title": "Drawdown (%)"},
                showlegend=False,
            )
            st.plotly_chart(fig_dd, use_container_width=True)

            # Return Distribution
            st.subheader("Return Distribution")
            daily_pct = returns_data[ticker].dropna()
            mean_r = float(daily_pct.mean())
            var95  = float(np.percentile(daily_pct, 5))

            fig_hist = go.Figure(go.Histogram(
                x=daily_pct.values,
                nbinsx=60,
                marker_color=BLUE,
                opacity=0.75,
            ))
            fig_hist.add_vline(x=mean_r, line_dash="dash", line_color=GREEN,
                               annotation_text=f"Mean {mean_r:.2f}%",
                               annotation_font_color=GREEN)
            fig_hist.add_vline(x=var95, line_dash="dot", line_color=RED,
                               annotation_text=f"VaR 95% {var95:.2f}%",
                               annotation_position="top left",
                               annotation_font_color=RED)
            fig_hist.update_layout(
                **base_layout(height=200),
                xaxis={**AXIS_STYLE, "title": "Daily Return (%)"},
                yaxis={**AXIS_STYLE, "title": "Frequency"},
                showlegend=False,
            )
            st.plotly_chart(fig_hist, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — PORTFOLIO
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Portfolio":
    st.title("Portfolio Analytics")
    st.caption("MASI-weighted market portfolio")

    # ── Parse All_Metrics ─────────────────────────────────────────────────────
    pm        = portfolio_all.copy()
    metric_col = pm.columns[0]
    entity_col = pm.columns[1]
    val_cols   = list(pm.columns[2:])
    short_val_labels = [c.split(" ")[0] if " " in str(c) else str(c) for c in val_cols]

    port_df = pm[pm[entity_col].str.contains("Portfolio", na=False)].copy()
    masi_df = pm[pm[entity_col].str.contains("MASI",      na=False)].copy()

    # Match selected window to a val_col
    sel_short   = SHORT_WINDOWS[win_idx]
    matched_col = val_cols[0]
    for i, sl in enumerate(short_val_labels):
        if sl == sel_short:
            matched_col = val_cols[i]
            break

    # ── Section A: Metric cards ───────────────────────────────────────────────
    st.subheader("Portfolio vs MASI — Selected Window")
    port_vals = port_df.set_index(metric_col)[matched_col]
    masi_vals = masi_df.set_index(metric_col)[matched_col]

    card_metrics = list(port_vals.index)
    for row_start in range(0, len(card_metrics), 5):
        row_m = card_metrics[row_start:row_start + 5]
        cols  = st.columns(len(row_m))
        for col, m_name in zip(cols, row_m):
            try:
                p_val = float(port_vals.get(m_name, float("nan")))
            except (TypeError, ValueError):
                p_val = float("nan")
            try:
                b_val = float(masi_vals.get(m_name, float("nan")))
            except (TypeError, ValueError):
                b_val = float("nan")

            if not pd.isna(p_val):
                if not pd.isna(b_val):
                    col.metric(m_name, f"{p_val:.3g}", delta=f"{p_val - b_val:+.3g} vs MASI")
                else:
                    col.metric(m_name, f"{p_val:.3g}")

    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

    # ── Section B: Metric Evolution Charts ────────────────────────────────────
    st.subheader("Metric Evolution Across Time Horizons")
    st.caption("How each metric changes as the lookback extends from 1Y → 5Y — Portfolio (blue) vs MASI (grey)")

    all_metric_names = list(port_df[metric_col].dropna().unique())

    # Group into rows of 3 charts
    chunk_size = 3
    for row_start in range(0, len(all_metric_names), chunk_size):
        group = all_metric_names[row_start:row_start + chunk_size]
        n = len(group)
        fig_evo = make_subplots(
            rows=1, cols=n,
            subplot_titles=group,
        )
        legend_added = False
        for ci, m_name in enumerate(group, start=1):
            p_row = port_df[port_df[metric_col] == m_name]
            m_row = masi_df[masi_df[metric_col] == m_name]

            if not p_row.empty:
                p_vals = []
                for c in val_cols:
                    try:
                        p_vals.append(float(p_row.iloc[0][c]))
                    except (TypeError, ValueError):
                        p_vals.append(None)
                fig_evo.add_trace(go.Scatter(
                    x=short_val_labels, y=p_vals,
                    name="Portfolio",
                    line=dict(color=BLUE, width=2),
                    marker=dict(size=6),
                    showlegend=not legend_added,
                ), row=1, col=ci)
                legend_added = True

            if not m_row.empty:
                m_vals = []
                for c in val_cols:
                    try:
                        m_vals.append(float(m_row.iloc[0][c]))
                    except (TypeError, ValueError):
                        m_vals.append(None)
                if any(v is not None for v in m_vals):
                    fig_evo.add_trace(go.Scatter(
                        x=short_val_labels, y=m_vals,
                        name="MASI",
                        line=dict(color=MUTED, width=1.5, dash="dot"),
                        marker=dict(size=5),
                        showlegend=(ci == 1 and not any(
                            t.name == "MASI" for t in fig_evo.data
                        )),
                    ), row=1, col=ci)

            fig_evo.update_xaxes(gridcolor=GRID_COL, linecolor="#E2E8F0",
                                 zeroline=False, row=1, col=ci)
            fig_evo.update_yaxes(gridcolor=GRID_COL, linecolor="#E2E8F0",
                                 zeroline=False, row=1, col=ci)

        fig_evo.update_layout(
            paper_bgcolor=PAPER_BG,
            plot_bgcolor=PLOT_BG,
            font=dict(color=TEXT_COL, family="Inter, sans-serif", size=11),
            margin=dict(t=45, b=30, l=10, r=10),
            height=270,
            legend=dict(orientation="h", y=1.18, x=0),
        )
        st.plotly_chart(fig_evo, use_container_width=True)

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

    # ── Section C: Weights ────────────────────────────────────────────────────
    st.subheader("Portfolio Weights")
    weights_file = DATA_DIR / "portfolio_weights.xlsx"
    if weights_file.exists():
        w_df  = pd.read_excel(weights_file, index_col=0)
        w_df.index = [str(t).replace(".CS", "") for t in w_df.index]
        top15 = w_df.iloc[:, 0].sort_values(ascending=False).head(15)
        fig_w = go.Figure(go.Bar(
            y=top15.index, x=top15.values,
            orientation="h",
            marker_color=BLUE,
            marker_line_width=0,
            text=[f"{v:.2%}" for v in top15.values],
            textposition="outside",
        ))
        fig_w.update_layout(
            **base_layout(height=380),
            xaxis={**AXIS_STYLE, "title": "Weight"},
            yaxis=AXIS_STYLE,
            showlegend=False,
        )
        st.plotly_chart(fig_w, use_container_width=True)
    else:
        st.info("Weights = MASI index composition. Detailed weight chart coming in v0.2.")

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

    # ── Section D: MASI Benchmark Table ───────────────────────────────────────
    st.subheader("MASI — Benchmark Reference")
    masi_table = masi_df[[metric_col] + val_cols].copy()
    masi_table.columns = [metric_col] + short_val_labels
    masi_table = masi_table.set_index(metric_col)
    st.dataframe(masi_table, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — PORTFOLIO OPTIMIZER
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Portfolio Optimizer":
    st.title("Portfolio Optimizer")
    st.info("No data available yet.")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — SCREENER
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Screener":
    st.title("Stock Screener")
    st.caption("Filter stocks by risk/return criteria — all 10 metrics")

    slider_config = {
        "Sharpe_Ratio":          ("Sharpe Ratio",         -3.0,  5.0,   (-3.0, 5.0),   0.1),
        "Sortino_Ratio":         ("Sortino Ratio",         -3.0,  8.0,   (-3.0, 8.0),   0.1),
        "Annualized_Return":     ("Return (ann.) %",      -50.0, 100.0,  (-50.0, 100.0), 1.0),
        "Annualized_Volatility": ("Volatility (ann.) %",   0.0,  100.0,  (0.0,  100.0),  1.0),
        "Beta":                  ("Beta",                  -1.0,   3.0,  (-1.0,  3.0),   0.05),
        "Alpha":                 ("Jensen's Alpha",        -20.0,  20.0, (-20.0, 20.0),  0.5),
        "Tracking_Error":        ("Tracking Error %",       0.0,   50.0, (0.0,  50.0),   0.5),
        "Max_Drawdown":          ("Max Drawdown %",       -100.0,   0.0, (-100.0, 0.0),  1.0),
        "VaR_95":                ("VaR 95% %",             -50.0,   0.0, (-50.0, 0.0),   0.5),
        "CVaR_95":               ("CVaR 95% %",            -50.0,   0.0, (-50.0, 0.0),   0.5),
    }

    active_metrics = [m for m in slider_config if m in metrics_data]
    filter_ranges  = {}

    for row_metrics in [active_metrics[i:i+3] for i in range(0, len(active_metrics), 3)]:
        cols = st.columns(3)
        for col, m in zip(cols, row_metrics):
            label, lo, hi, default, step = slider_config[m]
            filter_ranges[m] = col.slider(label, lo, hi, default, step)

    base_idx = pd.Index(TICKERS)
    mask     = pd.Series(True, index=base_idx)
    for m, (lo_val, hi_val) in filter_ranges.items():
        if m in metrics_data:
            series = metrics_data[m][window].reindex(base_idx)
            mask   = mask & (series >= lo_val) & (series <= hi_val)

    result_data = {"Ticker": TICKERS_CLEAN}
    for m in active_metrics:
        result_data[METRIC_LABELS.get(m, m)] = metrics_data[m][window].reindex(base_idx).round(2).values

    result_df = pd.DataFrame(result_data, index=base_idx)[mask]
    result_df = result_df.sort_values("Sharpe Ratio", ascending=False)

    count = int(mask.sum())
    color = GREEN if count > 10 else AMBER if count > 0 else RED
    st.markdown(
        f"<div style='font-size:1.1rem;font-weight:700;color:{color};margin-bottom:12px;'>"
        f"{count} stocks match</div>",
        unsafe_allow_html=True
    )

    styled_result = result_df.style
    for col_name in result_df.select_dtypes(include="number").columns:
        cmap = "RdYlGn_r" if any(k in col_name for k in ["Drawdown", "VaR", "CVaR", "Volatility"]) else "RdYlGn"
        styled_result = styled_result.background_gradient(cmap=cmap, subset=[col_name], axis=0)
    styled_result = styled_result.format(precision=2)
    st.dataframe(styled_result, use_container_width=True, height=500)


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 6 — AI ASSISTANT
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "AI Assistant":
    st.title("AI Assistant")
    st.info("No data available yet.")


# ═══════════════════════════════════════════════════════════════════════════════
# PAGE 7 — WORKFLOWS
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "Workflows":
    st.title("Workflows")
    st.info("No data available yet.")
