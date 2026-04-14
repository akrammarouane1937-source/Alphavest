"""
Alphavest — CSE Risk Analytics Dashboard
=========================================
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

# ── Paths ─────────────────────────────────────────────────────────────────────
DATA_DIR        = Path(__file__).parent / "data"
STOCKS_FILE     = DATA_DIR / "risk_metrics_by_window_latest.xlsx"
PORTFOLIO_FILE  = DATA_DIR / "portfolio_metrics_latest.xlsx"

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

metrics_data  = load_stocks()
returns_data  = load_returns()
portfolio_data = load_portfolio()

WINDOWS  = [c for c in list(metrics_data.values())[0].columns]
TICKERS  = [t for t in list(list(metrics_data.values())[0].index) if t != 'MASI']
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
    st.image("https://placehold.co/200x60/1B1F2E/FFFFFF?text=ALPHAVEST", width=200)
    st.markdown("---")

    page = st.radio(
        "Navigation",
        ["Market Overview", "Stock Analysis", "Portfolio", "Screener"],
        label_visibility="collapsed"
    )

    st.markdown("---")
    window = st.selectbox("Time Window", WINDOWS, index=0)
    st.caption(f"Data as of: **{WINDOWS[0].split('to')[-1].strip() if 'to' in WINDOWS[0] else 'latest'}**")

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(val, metric):
    if pd.isna(val):
        return "—"
    if metric in PCT_METRICS:
        return f"{val:.2f}%"
    return f"{val:.3f}"

def color_sharpe(val):
    if pd.isna(val): return "color: gray"
    if val > 1:      return "color: #2ecc71; font-weight: bold"
    if val > 0:      return "color: #f39c12"
    return "color: #e74c3c"

# ── Page: Market Overview ─────────────────────────────────────────────────────
if page == "Market Overview":
    st.title("Market Overview")
    st.caption("Casablanca Stock Exchange — all 80 listed companies")

    # ── KPI row ──
    sharpe_df = metrics_data["Sharpe_Ratio"][[window]].dropna()
    ret_df    = metrics_data["Annualized_Return"][[window]].dropna()
    vol_df    = metrics_data["Annualized_Volatility"][[window]].dropna()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Stocks tracked",      f"{len(TICKERS)}")
    col2.metric("Avg Sharpe",          f"{sharpe_df[window].mean():.2f}")
    col3.metric("Avg Ann. Return",     f"{ret_df[window].mean():.1f}%")
    col4.metric("Avg Ann. Volatility", f"{vol_df[window].mean():.1f}%")

    st.markdown("---")

    # ── Sharpe ranking chart ──
    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.subheader("Sharpe Ratio Ranking")
        sr = sharpe_df[window].sort_values(ascending=False).head(20)
        colors = ["#2ecc71" if v > 1 else "#f39c12" if v > 0 else "#e74c3c" for v in sr.values]
        fig = go.Figure(go.Bar(
            x=[t.replace(".CS","") for t in sr.index],
            y=sr.values,
            marker_color=colors,
            text=[f"{v:.2f}" for v in sr.values],
            textposition="outside"
        ))
        fig.add_hline(y=1, line_dash="dash", line_color="white", opacity=0.4)
        fig.update_layout(
            paper_bgcolor="#0E1117", plot_bgcolor="#0E1117",
            font_color="white", height=380,
            xaxis=dict(tickangle=-45),
            yaxis_title="Sharpe Ratio",
            margin=dict(t=10, b=60)
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.subheader("Return vs Risk")
        scatter_df = pd.DataFrame({
            "Return":     metrics_data["Annualized_Return"][window],
            "Volatility": metrics_data["Annualized_Volatility"][window],
            "Sharpe":     metrics_data["Sharpe_Ratio"][window],
        }).dropna()
        scatter_df = scatter_df[scatter_df.index != 'MASI']
        scatter_df["Ticker"] = [t.replace(".CS","") for t in scatter_df.index]
        fig2 = px.scatter(
            scatter_df, x="Volatility", y="Return",
            color="Sharpe", hover_name="Ticker",
            color_continuous_scale="RdYlGn",
            labels={"Volatility": "Volatility (%)", "Return": "Return (%)"},
        )
        fig2.update_layout(
            paper_bgcolor="#0E1117", plot_bgcolor="#0E1117",
            font_color="white", height=380,
            coloraxis_colorbar=dict(title="Sharpe"),
            margin=dict(t=10)
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── Full metrics table ──
    st.subheader("All Stocks — Key Metrics")
    table = pd.DataFrame(index=TICKERS)
    for metric in ["Annualized_Return", "Annualized_Volatility", "Sharpe_Ratio",
                   "Beta", "Max_Drawdown", "VaR_95"]:
        if metric in metrics_data:
            col_data = metrics_data[metric][window]
            if metric in PCT_METRICS:
                table[METRIC_LABELS[metric]] = col_data.round(2).astype(str) + "%"
            else:
                table[METRIC_LABELS[metric]] = col_data.round(3)
    table.index = [t.replace(".CS","") for t in table.index]
    st.dataframe(table, use_container_width=True, height=400)

# ── Page: Stock Analysis ──────────────────────────────────────────────────────
elif page == "Stock Analysis":
    st.title("Stock Analysis")

    ticker_label = st.selectbox("Select stock", TICKERS_CLEAN)
    ticker = ticker_label + ".CS"

    if ticker not in TICKERS:
        ticker = TICKERS[0]

    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("Risk Metrics")
        for metric in METRICS:
            val = metrics_data[metric].loc[ticker, window] if ticker in metrics_data[metric].index else None
            label = METRIC_LABELS.get(metric, metric)
            if val is not None and not pd.isna(val):
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

        # Cumulative returns chart
        if ticker in returns_data.columns:
            st.subheader("Cumulative Return (5Y)")
            ret_series = returns_data[ticker].dropna() / 100
            cum = (1 + ret_series).cumprod() - 1
            masi_cum = None
            if "MASI" in returns_data.columns:
                masi_cum = (1 + returns_data["MASI"].dropna() / 100).cumprod() - 1

            fig = go.Figure()
            fig.add_trace(go.Scatter(x=cum.index, y=cum.values * 100,
                                     name=ticker_label, line=dict(color="#3498db")))
            if masi_cum is not None:
                fig.add_trace(go.Scatter(x=masi_cum.index, y=masi_cum.values * 100,
                                         name="MASI", line=dict(color="#95a5a6", dash="dash")))
            fig.update_layout(
                paper_bgcolor="#0E1117", plot_bgcolor="#0E1117",
                font_color="white", height=300,
                yaxis_title="Cumulative Return (%)",
                legend=dict(orientation="h", y=1.1),
                margin=dict(t=10)
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

    st.markdown("---")
    st.subheader("Portfolio Optimization")
    st.info("Coming soon — Equal Weight / Min Variance / Max Sharpe / Risk Parity strategies")

# ── Page: Screener ────────────────────────────────────────────────────────────
elif page == "Screener":
    st.title("Stock Screener")
    st.caption("Filter stocks by risk/return criteria")

    col1, col2, col3 = st.columns(3)
    min_sharpe = col1.slider("Min Sharpe Ratio", -2.0, 3.0, 0.5, 0.1)
    max_vol    = col2.slider("Max Volatility (%)", 5, 80, 30, 1)
    max_dd     = col3.slider("Max Drawdown (%)", 5, 80, 40, 1)

    sharpe = metrics_data["Sharpe_Ratio"][window]
    vol    = metrics_data["Annualized_Volatility"][window]
    dd     = metrics_data["Max_Drawdown"][window]
    ret    = metrics_data["Annualized_Return"][window]

    mask = (sharpe >= min_sharpe) & (vol <= max_vol) & (dd <= max_dd)
    filtered = pd.DataFrame({
        "Ticker":     [t.replace(".CS","") for t in TICKERS],
        "Return (%)": ret.values.round(2),
        "Vol (%)":    vol.values.round(2),
        "Sharpe":     sharpe.values.round(3),
        "Max DD (%)": dd.values.round(2),
    }, index=TICKERS)[mask]

    st.metric("Stocks matching criteria", len(filtered))
    st.dataframe(filtered.sort_values("Sharpe", ascending=False),
                 use_container_width=True, height=500)
