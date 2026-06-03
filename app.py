import streamlit as st
import pandas as pd
import plotly.graph_objects as go

st.set_page_config(page_title="Battery Data Analyzer", layout="wide")
st.title("Battery Data Analyzer")

uploaded_file = st.file_uploader(
    "Upload an Excel or CSV file", type=["xlsx", "xls", "csv"]
)

if uploaded_file is None:
    st.info("Upload a file to get started.")
    st.stop()

# ── Load ──────────────────────────────────────────────────────────────────────
with st.spinner("Loading file…"):
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file, header=0, low_memory=False)
    else:
        df = pd.read_excel(uploaded_file, header=0)

st.success(f"Loaded {len(df):,} rows × {len(df.columns)} columns")

# ── Parse time column (column A, index 0) ────────────────────────────────────
try:
    time_col = pd.to_datetime(df.iloc[:, 0], format="%d/%m/%Y  %H:%M:%S")
except Exception:
    time_col = pd.to_datetime(df.iloc[:, 0], dayfirst=True)

# ── Helper ────────────────────────────────────────────────────────────────────
def make_chart(col_indices: list, title: str) -> go.Figure:
    fig = go.Figure()
    for i in col_indices:
        if i < len(df.columns):
            fig.add_trace(
                go.Scatter(
                    x=time_col,
                    y=pd.to_numeric(df.iloc[:, i], errors="coerce"),
                    mode="lines",
                    name=df.columns[i],
                )
            )
    fig.update_layout(
        title=title,
        xaxis_title="Time",
        xaxis=dict(rangeslider=dict(visible=True), type="date"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
        height=500,
    )
    return fig


# ── Chart 1: Column R (index 17) ─────────────────────────────────────────────
st.subheader("Chart 1 — Full Charge Capacity")
st.plotly_chart(
    make_chart([17], "SE Full_Charge_Capacity [Ah] vs Time"),
    use_container_width=True,
)

# ── Chart 2: Columns AF–AT (indices 31–45, 15 cols) ──────────────────────────
st.subheader("Chart 2 — Columns AF–AT")
st.plotly_chart(
    make_chart(list(range(31, 46)), "AF–AT vs Time"),
    use_container_width=True,
)

# ── Chart 3: Columns CY–DM (indices 102–116, 15 cols) ────────────────────────
st.subheader("Chart 3 — Columns CY–DM")
st.plotly_chart(
    make_chart(list(range(102, 117)), "CY–DM vs Time"),
    use_container_width=True,
)

# ── Chart 4: Columns CJ–CX (indices 87–101, 15 cols) ─────────────────────────
st.subheader("Chart 4 — Columns CJ–CX")
st.plotly_chart(
    make_chart(list(range(87, 102)), "CJ–CX vs Time"),
    use_container_width=True,
)
