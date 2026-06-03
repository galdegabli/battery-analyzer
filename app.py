import io
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Battery Data Analyzer", layout="wide")
st.title("Battery Data Analyzer")

# ── 15 distinct colors shared across all charts (index = cell/series number) ──
COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
    "#aec7e8", "#ffbb78", "#98df8a", "#ff9896", "#c5b0d5",
]

def series_color(col_name: str, fallback_idx: int) -> str:
    """Return the color for a column based on its trailing number (e.g. _7 → COLORS[7])."""
    m = re.search(r"_(\d+)", col_name)
    idx = int(m.group(1)) if m else fallback_idx
    return COLORS[idx % len(COLORS)]


# ── File upload ───────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload an Excel or CSV file", type=["xlsx", "xls", "csv"]
)

if uploaded_file is None:
    st.info("Upload a file to get started.")
    st.stop()

with st.spinner("Loading file…"):
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file, header=0, low_memory=False)
    else:
        df = pd.read_excel(uploaded_file, header=0)

st.success(f"Loaded {len(df):,} rows × {len(df.columns)} columns")

# ── Parse time column (column A, index 0) ────────────────────────────────────
raw_time = df.iloc[:, 0].astype(str).str.strip()
# Auto-detect format: ISO (YYYY-MM-DD …) vs DD/MM/YYYY
sample = raw_time.dropna().iloc[0] if not raw_time.dropna().empty else ""
if re.match(r"\d{4}-\d{2}-\d{2}", sample):
    time_col = pd.to_datetime(raw_time, format="%Y-%m-%d %H:%M:%S", errors="coerce")
else:
    # DD/MM/YYYY  HH:MM:SS  (single or double space)
    time_col = pd.to_datetime(raw_time.str.replace(r"\s+", " ", regex=True),
                               format="%d/%m/%Y %H:%M:%S", errors="coerce")

# ── Chart definitions ─────────────────────────────────────────────────────────
CHARTS = [
    {"title": "Full Charge Capacity",  "cols": [17],              "subtitle": "SE Full_Charge_Capacity [Ah] vs Time"},
    {"title": "Cell Voltage",          "cols": list(range(31, 46)), "subtitle": "CellVoltage_0 – CellVoltage_14 vs Time"},
    {"title": "Cell Distance",         "cols": list(range(102,117)),"subtitle": "CellDistanceAh_0 – CellDistanceAh_14 vs Time"},
    {"title": "Voltage Derivative",    "cols": list(range(87, 102)),"subtitle": "CellFdFVdQ_0 – CellFdFVdQ_14 vs Time"},
]

# ── Plotly helper ─────────────────────────────────────────────────────────────
def make_plotly_chart(col_indices: list, title: str) -> go.Figure:
    fig = go.Figure()
    for fb_idx, i in enumerate(col_indices):
        if i >= len(df.columns):
            continue
        col_name = df.columns[i]
        fig.add_trace(go.Scatter(
            x=time_col,
            y=pd.to_numeric(df.iloc[:, i], errors="coerce"),
            mode="lines",
            name=col_name,
            line=dict(color=series_color(col_name, fb_idx)),
        ))
    fig.update_layout(
        title=title,
        xaxis_title="Time",
        xaxis=dict(rangeslider=dict(visible=True), type="date"),
        yaxis=dict(fixedrange=False),          # allow Y-axis zoom & pan
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
        hoverlabel=dict(namelength=-1),        # show full series name in tooltip
        height=500,
    )
    return fig


# ── Render charts ─────────────────────────────────────────────────────────────
for chart_def in CHARTS:
    st.subheader(f"{'Chart ' + str(CHARTS.index(chart_def)+1)} — {chart_def['title']}")
    st.plotly_chart(
        make_plotly_chart(chart_def["cols"], chart_def["subtitle"]),
        use_container_width=True,
    )


# ── Excel export ──────────────────────────────────────────────────────────────
st.divider()
st.subheader("Download as Excel")

if st.button("Generate Excel file"):
    with st.spinner("Building Excel workbook…"):

        # Collect all relevant column indices
        all_col_indices = []
        for c in CHARTS:
            all_col_indices.extend(c["cols"])
        all_col_indices = sorted(set(all_col_indices))

        # Build a trimmed dataframe: time + all relevant columns
        export_cols = [df.columns[0]] + [df.columns[i] for i in all_col_indices if i < len(df.columns)]
        df_export = df[export_cols].copy()
        df_export[df_export.columns[0]] = time_col.dt.strftime("%Y-%m-%d %H:%M:%S")

        wb = Workbook()

        # ── Data sheet ────────────────────────────────────────────────────────
        ws_data = wb.active
        ws_data.title = "Data"

        for r in dataframe_to_rows(df_export, index=False, header=True):
            ws_data.append(r)

        # Bold header row
        for cell in ws_data[1]:
            cell.font = Font(bold=True)

        # Map original col index → Excel column letter in the Data sheet
        # Data sheet columns: col 1 = time, col 2 onwards = all_col_indices in order
        excel_col_map = {}  # original df col index → 1-based col number in ws_data
        excel_col_map[0] = 1  # time col
        for sheet_col, orig_idx in enumerate(
            [i for i in all_col_indices if i < len(df.columns)], start=2
        ):
            excel_col_map[orig_idx] = sheet_col

        nrows = len(df_export)

        # ── Chart sheets ──────────────────────────────────────────────────────
        for chart_def in CHARTS:
            ws_chart = wb.create_sheet(title=chart_def["title"][:31])

            chart = LineChart()
            chart.title = chart_def["subtitle"]
            chart.style = 10
            chart.height = 15
            chart.width = 30
            chart.x_axis.title = "Row (time)"
            chart.y_axis.title = "Value"

            for fb_idx, orig_idx in enumerate(chart_def["cols"]):
                if orig_idx not in excel_col_map:
                    continue
                sheet_col = excel_col_map[orig_idx]
                col_name = df.columns[orig_idx]

                data_ref = Reference(ws_data, min_col=sheet_col, min_row=1, max_row=nrows + 1)
                series = chart.series.append(data_ref)  # returns None; build differently
                chart.series[-1].title = SeriesLabel(v=col_name)
                chart.series[-1].graphicalProperties.line.solidFill = (
                    series_color(col_name, fb_idx).lstrip("#")
                )

            ws_chart.add_chart(chart, "A1")

        # ── Save to bytes ─────────────────────────────────────────────────────
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

    st.download_button(
        label="⬇️ Download Excel",
        data=buf,
        file_name="battery_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
