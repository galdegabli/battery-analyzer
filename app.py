import io
import json
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
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

# ── Chart HTML builder (custom sorted hover tooltip via JS) ───────────────────
def make_chart_html(col_indices: list, title: str, chart_id: str) -> str:
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
            hoverinfo="none",   # suppress Plotly's own tooltip; we draw ours
        ))
    fig.update_layout(
        title=title,
        xaxis_title="Time",
        xaxis=dict(rangeslider=dict(visible=True), type="date"),
        yaxis=dict(fixedrange=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x",          # fires hover event across all traces at cursor x
        height=500,
        margin=dict(t=60, b=80, l=60, r=20),
    )
    fig_json = fig.to_json()

    return f"""
<!DOCTYPE html><html><head>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  body {{ margin:0; padding:0; background:transparent; }}
  #tip_{chart_id} {{
    position: fixed;
    display: none;
    background: rgba(20,20,30,0.96);
    color: #e0e0f0;
    padding: 10px 14px;
    border-radius: 8px;
    font-size: 12px;
    font-family: "Segoe UI", monospace;
    pointer-events: none;
    z-index: 9999;
    max-height: 420px;
    overflow-y: auto;
    border: 1px solid #444466;
    box-shadow: 0 4px 16px rgba(0,0,0,0.5);
    line-height: 1.9;
    min-width: 220px;
  }}
  hr.tip-hr {{ border: none; border-top: 1px solid #444466; margin: 4px 0 6px 0; }}
</style>
</head><body>
<div id="plt_{chart_id}"></div>
<div id="tip_{chart_id}"></div>
<script>
(function() {{
  var fig = {fig_json};
  Plotly.newPlot('plt_{chart_id}', fig.data, fig.layout, {{responsive:true}}).then(function() {{
    var div = document.getElementById('plt_{chart_id}');
    var tip = document.getElementById('tip_{chart_id}');

    div.on('plotly_hover', function(ev) {{
      var pts = ev.points.slice().filter(function(p) {{
        return p.y !== null && p.y !== undefined && !isNaN(p.y);
      }});
      // sort descending by value
      pts.sort(function(a,b) {{ return b.y - a.y; }});

      if (!pts.length) {{ tip.style.display='none'; return; }}

      var timeStr = pts[0].x;
      var html = '<b style="color:#a0a8ff">' + timeStr + '</b>'
               + '<hr class="tip-hr">';
      pts.forEach(function(p) {{
        var col = (p.data.line && p.data.line.color) ? p.data.line.color : '#fff';
        var val = (typeof p.y === 'number')
          ? p.y.toLocaleString(undefined, {{maximumFractionDigits:4}})
          : p.y;
        html += '<span style="color:' + col + '">━</span>&nbsp;'
              + p.data.name + ':&nbsp;<b>' + val + '</b><br>';
      }});
      tip.innerHTML = html;
      tip.style.display = 'block';
    }});

    div.on('plotly_unhover', function() {{
      tip.style.display = 'none';
    }});

    document.addEventListener('mousemove', function(e) {{
      var x = e.clientX + 16, y = e.clientY - 12;
      var tw = tip.offsetWidth || 240, th = tip.offsetHeight || 300;
      if (x + tw > window.innerWidth)  x = e.clientX - tw - 10;
      if (y + th > window.innerHeight) y = e.clientY - th - 10;
      tip.style.left = x + 'px';
      tip.style.top  = y + 'px';
    }});
  }});
}})();
</script>
</body></html>
"""


# ── Render charts ─────────────────────────────────────────────────────────────
for idx, chart_def in enumerate(CHARTS):
    st.subheader(f"Chart {idx+1} — {chart_def['title']}")
    components.html(
        make_chart_html(chart_def["cols"], chart_def["subtitle"], f"c{idx}"),
        height=570,
        scrolling=False,
    )


# ── Excel export ──────────────────────────────────────────────────────────────
st.divider()
st.subheader("Download as Excel")

col1, col2 = st.columns([3, 1])
with col1:
    excel_filename = st.text_input("File name", value="battery_analysis", label_visibility="collapsed",
                                   placeholder="File name (without .xlsx)")
with col2:
    generate = st.button("Generate Excel file", use_container_width=True)

if generate:
    fname = (excel_filename.strip() or "battery_analysis").removesuffix(".xlsx") + ".xlsx"

    with st.spinner("Building Excel workbook…"):

        # ── Collect relevant columns ──────────────────────────────────────────
        all_col_indices = sorted(set(i for c in CHARTS for i in c["cols"] if i < len(df.columns)))

        # ── Full data export dataframe (time as real datetime) ────────────────
        export_cols = [df.columns[0]] + [df.columns[i] for i in all_col_indices]
        df_full = df[export_cols].copy()
        df_full[df_full.columns[0]] = time_col  # keep as datetime objects

        wb = Workbook()

        # ── Helper: write a dataframe to a sheet with real timestamps ─────────
        def write_sheet(ws, data: pd.DataFrame):
            ws.append(list(data.columns))
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for row_vals in data.itertuples(index=False, name=None):
                ws.append(list(row_vals))
            dt_fmt = "YYYY-MM-DD HH:MM:SS"
            for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
                for cell in row:
                    cell.number_format = dt_fmt

        # ── Data sheet (full resolution) ──────────────────────────────────────
        ws_data = wb.active
        ws_data.title = "Data"
        write_sheet(ws_data, df_full)

        # col index map for Data sheet
        chart_col_map = {0: 1}
        for sc, oi in enumerate(all_col_indices, start=2):
            chart_col_map[oi] = sc
        n_data_rows = len(df_full)

        # ── Single Charts sheet — all 4 charts stacked vertically ─────────────
        ws_charts = wb.create_sheet("Charts")
        CHART_H = 15   # height in cm
        ROW_OFFSET = 30  # rows between chart anchors (~15 cm each)

        for chart_idx, chart_def in enumerate(CHARTS):
            chart = LineChart()
            chart.title = chart_def["subtitle"]
            chart.style = 10
            chart.height = CHART_H
            chart.width = 30
            chart.x_axis.title = "Time"
            chart.y_axis.title = "Value"

            valid_cols = [i for i in chart_def["cols"] if i in chart_col_map]
            if valid_cols:
                min_sc = chart_col_map[valid_cols[0]]
                max_sc = chart_col_map[valid_cols[-1]]
                data_ref = Reference(ws_data, min_col=min_sc, max_col=max_sc,
                                     min_row=1, max_row=n_data_rows + 1)
                chart.add_data(data_ref, titles_from_data=True)

                for fb_idx, (orig_idx, ser) in enumerate(zip(valid_cols, chart.series)):
                    col_name = df.columns[orig_idx]
                    ser.graphicalProperties.line.solidFill = series_color(col_name, fb_idx).lstrip("#")
                    ser.graphicalProperties.line.width = 12000  # 1.2pt

            anchor = f"A{1 + chart_idx * ROW_OFFSET}"
            ws_charts.add_chart(chart, anchor)

        # ── Save ──────────────────────────────────────────────────────────────
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

    st.download_button(
        label=f"⬇️ Download {fname}",
        data=buf,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
