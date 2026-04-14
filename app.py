"""
Alphavest — CSE Risk Analytics Dashboard  v0.1
===============================================
Streamlit app reading from pre-computed Excel outputs.
Data refreshed daily via automated pipeline.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
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
  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #F8FAFC;
    border-right: 1px solid #E2E8F0;
  }
  /* Metric cards */
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
  /* Page title */
  h1 { color: #0F172A; font-weight: 800; letter-spacing: -0.5px; }
  h2, h3 { color: #1E293B; font-weight: 700; }
  /* Dataframe */
  [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
  /* Remove default top padding */
  .block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ── Paths ─────────────────────────────────────────────────────────────────────
DATA_DIR       = Path(__file__).parent / "data"
STOCKS_FILE    = DATA_DIR / "risk_metrics_by_window_latest.xlsx"
PORTFOLIO_FILE = DATA_DIR / "portfolio_metrics_latest.xlsx"

# ── Plotly theme helpers ───────────────────────────────────────────────────────
PLOT_BG   = "#FFFFFF"
PAPER_BG  = "#FFFFFF"
GRID_COL  = "#F1F5F9"
TEXT_COL  = "#0F172A"
BLUE      = "#2563EB"
GREEN     = "#16A34A"
AMBER     = "#D97706"
RED       = "#DC2626"
MUTED     = "#94A3B8"

def base_layout(**kwargs):
    return dict(
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=PLOT_BG,
        font=dict(color=TEXT_COL, family="Inter, sans-serif", size=12),
        xaxis=dict(gridcolor=GRID_COL, linecolor="#E2E8F0", zeroline=False),
        yaxis=dict(gridcolor=GRID_COL, linecolor="#E2E8F0", zeroline=False),
        margin=dict(t=20, b=40, l=10, r=10),
        **kwargs
    )

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
    return pd.read_excel(STOCKS_FILE, sheet_name="Daily_Returns", index_col=0)

@st.cache_data(ttl=3600)
def load_portfolio():
    return pd.read_excel(PORTFOLIO_FILE, sheet_name="All_Metrics", index_col=0)

metrics_data   = load_stocks()
returns_data   = load_returns()
portfolio_data = load_portfolio()

WINDOWS  = list(list(metrics_data.values())[0].columns)
TICKERS  = [t for t in list(list(metrics_data.values())[0].index) if t != "MASI"]
TICKERS_CLEAN = [t.replace(".CS", "") for t in TICKERS]
METRICS  = list(metrics_data.keys())

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
        ["Market Overview", "Stock Analysis", "Portfolio", "Screener"],
        label_visibility="collapsed"
    )

    st.markdown("<div style='margin: 20px 0 8px 0; font-size:0.72rem; color:#94A3B8; text-transform:uppercase; letter-spacing:0.08em;'>Time Window</div>", unsafe_allow_html=True)
    window = st.selectbox("Time Window", WINDOWS, index=0, label_visibility="collapsed")

    end_label = WINDOWS[0].split("to")[-1].strip() if "to" in WINDOWS[0] else "latest"
    st.markdown(f"<div style='font-size:0.72rem; color:#94A3B8; margin-top:6px;'>Data as of <b style='color:#475569;'>{end_label}</b></div>", unsafe_allow_html=True)

    st.markdown("<div style='position:absolute; bottom:20px; font-size:0.68rem; color:#CBD5E1;'>v0.1 · Casablanca SE</div>", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(val, metric):
    if pd.isna(val):
        return "n/a"
    if metric in PCT_METRICS:
        return f"{val:.2f}%"
    return f"{val:.3f}"

def sharpe_color(val):
    if pd.isna(val): return MUTED
    if val > 1:      return GREEN
    if val > 0:      return AMBER
    return RED

# ── Page: Market Overview ─────────────────────────────────────────────────────
if page == "Market Overview":
    st.title("Market Overview")
    st.caption("Casablanca Stock Exchange — 80 listed companies")

    sharpe_df = metrics_data["Sharpe_Ratio"].dropna()
    ret_df    = metrics_data["Annualized_Return"].dropna()
    vol_df    = metrics_data["Annualized_Volatility"].dropna()

    # Filter out MASI from stock-level stats
    sharpe_stocks = sharpe_df[sharpe_df.index != "MASI"]
    ret_stocks    = ret_df[ret_df.index != "MASI"]
    vol_stocks    = vol_df[vol_df.index != "MASI"]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Stocks Tracked",      f"{len(TICKERS)}")
    col2.metric("Avg Sharpe",          f"{sharpe_stocks[window].mean():.2f}")
    col3.metric("Avg Ann. Return",     f"{ret_stocks[window].mean():.1f}%")
    col4.metric("Avg Ann. Volatility", f"{vol_stocks[window].mean():.1f}%")

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.subheader("Sharpe Ratio Ranking")
        sr = sharpe_stocks[window].sort_values(ascending=False).head(20)
        colors = [GREEN if v > 1 else AMBER if v > 0 else RED for v in sr.values]
        fig = go.Figure(go.Bar(
            x=[t.replace(".CS", "") for t in sr.index],
            y=sr.values,
            marker_color=colors,
            marker_line_width=0,
            text=[f"{v:.2f}" for v in sr.values],
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig.add_hline(y=1, line_dash="dot", line_color=BLUE, opacity=0.5,
                      annotation_text="Sharpe = 1", annotation_position="top right",
                      annotation_font_color=BLUE)
        fig.update_layout(
            **base_layout(height=360),
            xaxis=dict(tickangle=-40, gridcolor=GRID_COL, linecolor="#E2E8F0"),
            yaxis=dict(title="Sharpe Ratio", gridcolor=GRID_COL),
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.subheader("Return vs Risk")
        scatter_df = pd.DataFrame({
            "Return":     metrics_data["Annualized_Return"][window],
            "Volatility": metrics_data["Annualized_Volatility"][window],
            "Sharpe":     metrics_data["Sharpe_Ratio"][window],
        }).dropna()
        scatter_df = scatter_df[scatter_df.index != "MASI"].copy()
        scatter_df["Ticker"] = [t.replace(".CS", "") for t in scatter_df.index]

        fig2 = px.scatter(
            scatter_df, x="Volatility", y="Return",
            color="Sharpe", hover_name="Ticker",
            color_continuous_scale=[[0, RED], [0.4, AMBER], [1, GREEN]],
            labels={"Volatility": "Volatility (%)", "Return": "Return (%)"},
        )
        fig2.update_traces(marker=dict(size=7, line=dict(width=0.5, color="white")))
        fig2.update_layout(
            **base_layout(height=360),
            coloraxis_colorbar=dict(title="Sharpe", thickness=12, len=0.8),
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("All Stocks — Key Metrics")
    table = pd.DataFrame(index=TICKERS)
    for metric in ["Annualized_Return", "Annualized_Volatility", "Sharpe_Ratio",
                   "Beta", "Max_Drawdown", "VaR_95"]:
        if metric in metrics_data:
            col_data = metrics_data[metric][window].reindex(TICKERS)
            if metric in PCT_METRICS:
                table[METRIC_LABELS[metric]] = col_data.round(2).astype(str) + "%"
            else:
                table[METRIC_LABELS[metric]] = col_data.round(3)
    table.index = TICKERS_CLEAN
    st.dataframe(table, use_container_width=True, height=420)

# ── Page: Stock Analysis ──────────────────────────────────────────────────────
elif page == "Stock Analysis":
    st.title("Stock Analysis")

    ticker_label = st.selectbox("Select stock", TICKERS_CLEAN)
    ticker = ticker_label + ".CS"
    if ticker not in TICKERS:
        ticker = TICKERS[0]

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("Risk Metrics")
        for metric in METRICS:
            if ticker in metrics_data[metric].index:
                val = metrics_data[metric].loc[ticker, window]
                label = METRIC_LABELS.get(metric, metric)
                if not pd.isna(val):
                    display = f"{val:.2f}%" if metric in PCT_METRICS else f"{val:.3f}"
                    st.metric(label, display)

    with col2:
        st.subheader("Metrics Across Windows")
        rows = []
        for w in WINDOWS:
            row = {"Window": w.split("(")[0].strip()}
            for metric in ["Annualized_Return", "Sharpe_Ratio", "Annualized_Volatility", "Max_Drawdown"]:
                if metric in metrics_data and ticker in metrics_data[metric].index:
                    val = metrics_data[metric].loc[ticker, w]
                    row[METRIC_LABELS[metric]] = round(val, 2) if metric in PCT_METRICS else round(val, 3)
            rows.append(row)
        window_df = pd.DataFrame(rows).set_index("Window")
        st.dataframe(window_df, use_container_width=True)

        if ticker in returns_data.columns:
            st.subheader("Cumulative Return (5Y)")
            ret_series = returns_data[ticker].dropna() / 100
            cum = (1 + ret_series).cumprod() - 1

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=cum.index, y=cum.values * 100,
                name=ticker_label,
                line=dict(color=BLUE, width=2),
                fill="tozeroy",
                fillcolor="rgba(37,99,235,0.07)",
            ))
            if "MASI" in returns_data.columns:
                masi_cum = (1 + returns_data["MASI"].dropna() / 100).cumprod() - 1
                fig.add_trace(go.Scatter(
                    x=masi_cum.index, y=masi_cum.values * 100,
                    name="MASI", line=dict(color=MUTED, width=1.5, dash="dot"),
                ))
            fig.update_layout(
                **base_layout(height=280),
                yaxis=dict(title="Cumulative Return (%)", gridcolor=GRID_COL),
                legend=dict(orientation="h", y=1.12, x=0),
            )
            st.plotly_chart(fig, use_container_width=True)

# ── Page: Portfolio ───────────────────────────────────────────────────────────
elif page == "Portfolio":
    st.title("Portfolio Analytics")
    st.caption("MASI-weighted market portfolio")

    if not portfolio_data.empty:
        st.dataframe(portfolio_data, use_container_width=True)
    else:
        st.info("Portfolio data not available.")

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    st.subheader("Portfolio Optimization")
    st.info("Coming soon — Equal Weight / Min Variance / Max Sharpe / Risk Parity strategies")

# ── Page: Screener ────────────────────────────────────────────────────────────
elif page == "Screener":
    st.title("Stock Screener")
    st.caption("Filter stocks by risk/return criteria")

    col1, col2, col3 = st.columns(3)
    min_sharpe = col1.slider("Min Sharpe Ratio", -2.0, 3.0, 0.5, 0.1)
    max_vol    = col2.slider("Max Volatility (%)", 5, 80, 40, 1)
    max_dd     = col3.slider("Max Drawdown (%)", 5, 80, 60, 1)

    sharpe = metrics_data["Sharpe_Ratio"][window].reindex(TICKERS)
    vol    = metrics_data["Annualized_Volatility"][window].reindex(TICKERS)
    dd     = metrics_data["Max_Drawdown"][window].reindex(TICKERS)
    ret    = metrics_data["Annualized_Return"][window].reindex(TICKERS)

    mask = (sharpe >= min_sharpe) & (vol <= max_vol) & (dd <= max_dd)
    filtered = pd.DataFrame({
        "Ticker":     TICKERS_CLEAN,
        "Return (%)": ret.values.round(2),
        "Vol (%)":    vol.values.round(2),
        "Sharpe":     sharpe.values.round(3),
        "Max DD (%)": dd.values.round(2),
    }, index=TICKERS)[mask]

    count = int(mask.sum())
    color = GREEN if count > 10 else AMBER if count > 0 else RED
    st.markdown(f"<div style='font-size:1.1rem; font-weight:700; color:{color}; margin-bottom:12px;'>{count} stocks match</div>", unsafe_allow_html=True)
    st.dataframe(
        filtered.sort_values("Sharpe", ascending=False),
        use_container_width=True, height=500
    )
