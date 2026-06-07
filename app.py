import io
import json
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.styles import Font

st.set_page_config(page_title="Battery Data Analyzer", layout="wide")
st.title("Battery Data Analyzer")

# ── 15 distinct colors shared across all charts (index = cell/series number) ──
COLORS = [
    "#1f77b4",  # 0  steel blue
    "#e6331a",  # 1  strong red
    "#2ca02c",  # 2  forest green
    "#ff7f0e",  # 3  vivid orange
    "#9467bd",  # 4  purple
    "#00aacc",  # 5  cyan
    "#8c564b",  # 6  brown
    "#ff1493",  # 7  deep pink
    "#17becf",  # 8  teal
    "#b8b800",  # 9  dark yellow
    "#8c1aff",  # 10 violet
    "#d62728",  # 11 crimson
    "#00b300",  # 12 bright green
    "#cc6600",  # 13 dark orange
    "#0040ff",  # 14 cobalt blue
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
        is_re_export = False
    else:
        xl = pd.ExcelFile(uploaded_file)
        is_re_export = "_meta" in xl.sheet_names
        sheet = "Data" if is_re_export else xl.sheet_names[0]
        df = xl.parse(sheet, header=0)

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

# ── Column name → index resolver ─────────────────────────────────────────────
def find_cols(pattern: str) -> list[int]:
    """Return column indices whose names start with `pattern`, sorted numerically."""
    def num_key(i):
        m = re.search(r"(\d+)", df.columns[i][len(pattern):])
        return int(m.group(1)) if m else 0
    return sorted(
        [i for i, c in enumerate(df.columns) if c.startswith(pattern)],
        key=num_key
    )

def find_col(name: str) -> list[int]:
    """Return a single-element list for an exact column name match."""
    matches = [i for i, c in enumerate(df.columns) if c == name]
    return matches[:1]

# ── Chart definitions (resolved by column name, not position) ────────────────
CHARTS = [
    {
        "title":    "Full Charge Capacity",
        "cols":     find_col("SE Full_Charge_Capacity [Ah]"),
        "subtitle": "SE Full_Charge_Capacity [Ah] vs Time",
    },
    {
        "title":    "Cell Voltage",
        "cols":     find_cols("CellVoltage_"),
        "subtitle": "CellVoltage_0 – CellVoltage_14 vs Time",
        "auto_multiply": True,   # factor computed from data so result has 1 digit before decimal
    },
    {
        "title":    "Cell Distance",
        "cols":     find_cols("CellDistanceAh_"),
        "subtitle": "CellDistanceAh_0 – CellDistanceAh_14 vs Time",
    },
    {
        "title":    "Voltage Derivative",
        "cols":     find_cols("CellFdFVdQ_"),
        "subtitle": "CellFdFVdQ_0 – CellFdFVdQ_14 vs Time",
        "decimals": 0,
    },
]

# ── Auto-multiply: bring Cell Voltage (or any flagged chart) to 1 digit before decimal ──
import math as _math

def _auto_multiply(col_indices: list) -> float:
    """Return a power-of-10 factor so the typical value has exactly 1 digit before the decimal."""
    if not col_indices:
        return 1.0
    vals = pd.concat([pd.to_numeric(df.iloc[:, i], errors="coerce") for i in col_indices])
    median = vals.abs().median()
    if pd.isna(median) or median == 0:
        return 1.0
    digits_before_decimal = _math.floor(_math.log10(float(median))) + 1
    return 10 ** (-(digits_before_decimal - 1))

for _chart in CHARTS:
    if _chart.get("auto_multiply") and _chart["cols"]:
        _chart["multiply"] = _auto_multiply(_chart["cols"])

# ── Chart HTML builder (custom sorted hover tooltip via JS) ───────────────────
def make_chart_html(col_indices: list, title: str, chart_id: str,
                    x_range: tuple | None = None,
                    multiply: float | None = None,
                    decimals: int | None = None) -> str:
    fig = go.Figure()
    for fb_idx, i in enumerate(col_indices):
        if i >= len(df.columns):
            continue
        col_name = df.columns[i]
        y = pd.to_numeric(df.iloc[:, i], errors="coerce")
        if multiply is not None:
            y = y * multiply
        if decimals is not None:
            y = y.round(decimals)
        fig.add_trace(go.Scatter(
            x=time_col,
            y=y,
            mode="lines",
            name=col_name,
            line=dict(color=series_color(col_name, fb_idx)),
            hoverinfo="none",   # suppress Plotly's own tooltip; we draw ours
        ))
    xaxis_cfg = dict(type="date", rangeslider=dict(visible=False))
    if x_range:
        xaxis_cfg["range"] = [x_range[0].isoformat(), x_range[1].isoformat()]
    fig.update_layout(
        title=dict(text=title, y=0.97, x=0.5, xanchor="center", yanchor="top"),
        xaxis_title="Time",
        xaxis=xaxis_cfg,
        yaxis=dict(fixedrange=False),
        # legend placed below the chart, outside the plot area
        legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
        hovermode="x",
        height=480,
        margin=dict(t=40, b=100, l=60, r=20),
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
      if (x < 0) x = 0;
      if (y + th > window.innerHeight) y = e.clientY - th - 10;
      if (y < 0) y = 0;
      tip.style.left = x + 'px';
      tip.style.top  = y + 'px';
    }});
  }});
}})();
</script>
</body></html>
"""


# ── Global time range slider ──────────────────────────────────────────────────
from datetime import timedelta

t_min = time_col.min().to_pydatetime()
t_max = time_col.max().to_pydatetime()


st.markdown("#### Time range")
t_start, t_end = st.slider(
    "time_range",
    min_value=t_min,
    max_value=t_max,
    value=(t_min, t_max),
    step=timedelta(minutes=1),
    format="DD/MM/YY HH:mm",
    label_visibility="collapsed",
)

# ── Render charts ─────────────────────────────────────────────────────────────
for idx, chart_def in enumerate(CHARTS):
    st.subheader(f"Chart {idx+1} — {chart_def['title']}")
    components.html(
        make_chart_html(chart_def["cols"], chart_def["subtitle"], f"c{idx}",
                        x_range=(t_start, t_end),
                        multiply=chart_def.get("multiply"),
                        decimals=chart_def.get("decimals")),
        height=520,
        scrolling=False,
    )


# ── Excel export ──────────────────────────────────────────────────────────────
import math
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

st.divider()
st.subheader("Download as Excel")

col1, col2 = st.columns([3, 1])
with col1:
    default_name = uploaded_file.name.rsplit(".", 1)[0]
    excel_filename = st.text_input("File name", value=default_name, label_visibility="collapsed",
                                   placeholder="File name (without .xlsx)")
with col2:
    generate = st.button("Generate Excel file", use_container_width=True)

if generate:
    fname = (excel_filename.strip() or "battery_analysis").removesuffix(".xlsx") + ".xlsx"

    with st.spinner("Building Excel workbook…"):

        # ── Collect relevant columns ──────────────────────────────────────────
        all_col_indices = sorted(set(i for c in CHARTS for i in c["cols"] if i < len(df.columns)))
        export_cols = [df.columns[0]] + [df.columns[i] for i in all_col_indices]

        # Sheet column number for each original df column index (0-based, col 0 = time)
        sheet_col = {0: 0}
        for pos, oi in enumerate(all_col_indices, start=1):
            sheet_col[oi] = pos
        n = len(df)

        # Per-column transform: orig df col index → (multiply, decimals)
        col_transforms = {}
        for chart_def in CHARTS:
            m = chart_def.get("multiply")
            d = chart_def.get("decimals")
            if m is not None or d is not None:
                for oi in chart_def["cols"]:
                    col_transforms[oi] = (m, d)

        buf = io.BytesIO()
        wb = xlsxwriter.Workbook(buf, {'in_memory': True})

        # ── Formats ───────────────────────────────────────────────────────────
        fmt_bold   = wb.add_format({'bold': True})
        fmt_date   = wb.add_format({'num_format': 'dd/mm/yyyy hh:mm:ss', 'bold': False})

        # ── Data sheet ────────────────────────────────────────────────────────
        ws_data = wb.add_worksheet('Data')

        # Header row
        ws_data.write(0, 0, df.columns[0], fmt_bold)
        for pos, oi in enumerate(all_col_indices, start=1):
            ws_data.write(0, pos, df.columns[oi], fmt_bold)

        # Data rows
        for row_i, (t, *vals) in enumerate(
            zip(time_col, *[df.iloc[:, oi] for oi in all_col_indices]), start=1
        ):
            dt = t.to_pydatetime() if pd.notna(t) else None
            if dt:
                ws_data.write_datetime(row_i, 0, dt, fmt_date)
            for pos, (oi, v) in enumerate(zip(all_col_indices, vals), start=1):
                fv = float(v) if (pd.notna(v) and not (isinstance(v, float) and math.isnan(v))) else None
                if fv is not None:
                    tm, td = col_transforms.get(oi, (None, None))
                    if tm is not None:
                        fv *= tm
                    if td is not None:
                        fv = round(fv, td)
                    ws_data.write_number(row_i, pos, fv)

        # ── Charts sheet ──────────────────────────────────────────────────────
        ws_charts = wb.add_worksheet('Charts')

        # Excel date serial = days since 1899-12-30 (accounts for Excel's 1900 leap-year bug)
        from datetime import datetime as _dt
        _xl_epoch = _dt(1899, 12, 30)
        def _to_xl_serial(ts):
            d = ts.to_pydatetime().replace(tzinfo=None) - _xl_epoch
            return d.days + d.seconds / 86400

        _xl_min = _to_xl_serial(time_col.dropna().min())
        _xl_max = _to_xl_serial(time_col.dropna().max())

        CHART_H   = 500   # pixels
        CHART_W   = 960
        ROW_PX    = 20    # approximate row height in pixels
        CHART_GAP = CHART_H + 20

        for chart_idx, chart_def in enumerate(CHARTS):
            chart = wb.add_chart({'type': 'scatter', 'subtype': 'smooth'})

            valid_cols = [i for i in chart_def["cols"] if i in sheet_col]
            if not valid_cols:
                continue
            for fb_idx, orig_idx in enumerate(valid_cols):
                col_name  = df.columns[orig_idx]
                cli       = sheet_col[orig_idx]
                col_ltr   = xl_col_to_name(cli)   # e.g. 0→A, 25→Z, 26→AA
                color     = series_color(col_name, fb_idx)

                chart.add_series({
                    'name':       f"=Data!${col_ltr}$1",
                    'categories': f"=Data!$A$2:$A${n+1}",
                    'values':     f"=Data!${col_ltr}$2:${col_ltr}${n+1}",
                    'line':       {'color': color, 'width': 1.25},
                    'marker':     {'type': 'none'},
                })

            chart.set_title({'name': chart_def["title"]})
            chart.set_x_axis({
                'name':            'Time',
                'date_axis':       True,
                'num_format':      'dd/mm/yy hh:mm',
                'min':             _xl_min,
                'max':             _xl_max,
                'major_gridlines': {'visible': True,  'line': {'color': '#D0D0D0'}},
                'minor_gridlines': {'visible': False},
            })
            chart.set_y_axis({
                'name':            'Value',
                'major_gridlines': {'visible': True,  'line': {'color': '#D0D0D0'}},
                'minor_gridlines': {'visible': True,  'line': {'color': '#EBEBEB'}},
            })
            chart.set_legend({'position': 'bottom'})
            chart.set_size({'width': CHART_W, 'height': CHART_H})

            # Stack charts vertically
            top_px = chart_idx * CHART_GAP
            ws_charts.insert_chart(0, 0, chart, {'x_offset': 0, 'y_offset': top_px})

        # Hidden marker so the app knows this file was exported by us
        ws_meta = wb.add_worksheet('_meta')
        ws_meta.write(0, 0, 'battery-analyzer-export-v1')
        ws_meta.hide()

        wb.close()
        buf.seek(0)

    st.download_button(
        label=f"⬇️ Download {fname}",
        data=buf,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
