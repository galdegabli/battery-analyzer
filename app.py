import io
import math as _math
import re
from datetime import timedelta, datetime as _dt

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(page_title="Battery Data Analyzer", layout="wide")
st.title("Battery Data Analyzer")

# ── 15 distinct colors (index = cell number) ────────────────────────────────────
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
    m = re.search(r"_(\d+)", col_name)
    idx = int(m.group(1)) if m else fallback_idx
    return COLORS[idx % len(COLORS)]


# ── File upload (multiple) ───────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Upload Excel or CSV files", type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload a file to get started.")
    st.stop()


# ── Per-file processing helpers ──────────────────────────────────────────────────
def load_file(f):
    """Load a file into a DataFrame; detect re-exports via hidden _meta sheet."""
    if f.name.lower().endswith(".csv"):
        df = pd.read_csv(f, header=0, low_memory=False)
        is_re_export = False
    else:
        xl = pd.ExcelFile(f)
        is_re_export = "_meta" in xl.sheet_names
        if is_re_export:
            data_sheets = [s for s in xl.sheet_names if s.startswith("Data")]
            sheet = data_sheets[0] if data_sheets else xl.sheet_names[0]
        else:
            sheet = xl.sheet_names[0]
        df = xl.parse(sheet, header=0)
    return df, is_re_export


def parse_time(df):
    """Parse the first column as datetime, auto-detecting ISO vs DD/MM/YYYY."""
    raw = df.iloc[:, 0].astype(str).str.strip()
    sample = raw.dropna().iloc[0] if not raw.dropna().empty else ""
    if re.match(r"\d{4}-\d{2}-\d{2}", sample):
        return pd.to_datetime(raw, format="%Y-%m-%d %H:%M:%S", errors="coerce")
    return pd.to_datetime(
        raw.str.replace(r"\s+", " ", regex=True),
        format="%d/%m/%Y %H:%M:%S", errors="coerce",
    )


def find_cols_in(df, pattern):
    def num_key(i):
        m = re.search(r"(\d+)", df.columns[i][len(pattern):])
        return int(m.group(1)) if m else 0
    return sorted(
        [i for i, c in enumerate(df.columns) if c.startswith(pattern)],
        key=num_key,
    )


def find_col_in(df, name):
    return [i for i, c in enumerate(df.columns) if c == name][:1]


def _auto_multiply(df, col_indices):
    """Power-of-10 factor so the typical value has exactly 1 digit before decimal."""
    if not col_indices:
        return 1.0
    vals = pd.concat([pd.to_numeric(df.iloc[:, i], errors="coerce") for i in col_indices])
    median = vals.abs().median()
    if pd.isna(median) or median == 0:
        return 1.0
    digits = _math.floor(_math.log10(float(median))) + 1
    return 10 ** (-(digits - 1))


def build_charts(df):
    charts = [
        {
            "title":    "Full Charge Capacity",
            "cols":     find_col_in(df, "SE Full_Charge_Capacity [Ah]"),
            "subtitle": "SE Full_Charge_Capacity [Ah] vs Time",
        },
        {
            "title":        "Cell Voltage",
            "cols":         find_cols_in(df, "CellVoltage_"),
            "subtitle":     "CellVoltage_0 – CellVoltage_14 vs Time",
            "auto_multiply": True,
        },
        {
            "title":    "Cell Distance",
            "cols":     find_cols_in(df, "CellDistanceAh_"),
            "subtitle": "CellDistanceAh_0 – CellDistanceAh_14 vs Time",
        },
        {
            "title":    "Voltage Derivative",
            "cols":     find_cols_in(df, "CellFdFVdQ_"),
            "subtitle": "CellFdFVdQ_0 – CellFdFVdQ_14 vs Time",
            "decimals": 0,
        },
    ]
    for c in charts:
        if c.get("auto_multiply") and c["cols"]:
            c["multiply"] = _auto_multiply(df, c["cols"])
    return charts


# ── Chart HTML builder (custom sorted hover via JS) ──────────────────────────────
def make_chart_html(df, time_col, col_indices, title, chart_id,
                    x_range=None, multiply=None, decimals=None):
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
            x=time_col, y=y, mode="lines", name=col_name,
            line=dict(color=series_color(col_name, fb_idx)),
            hoverinfo="none",
        ))
    xaxis_cfg = dict(type="date", rangeslider=dict(visible=False))
    if x_range:
        xaxis_cfg["range"] = [x_range[0].isoformat(), x_range[1].isoformat()]
    fig.update_layout(
        title=dict(text=title, y=0.97, x=0.5, xanchor="center", yanchor="top"),
        xaxis_title="Time", xaxis=xaxis_cfg,
        yaxis=dict(fixedrange=False),
        legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
        hovermode="x", height=480,
        margin=dict(t=40, b=100, l=60, r=20),
    )
    fig_json = fig.to_json()
    return f"""
<!DOCTYPE html><html><head>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  body {{ margin:0; padding:0; background:transparent; }}
  #tip_{chart_id} {{
    position: fixed; display: none;
    background: rgba(20,20,30,0.96); color: #e0e0f0;
    padding: 10px 14px; border-radius: 8px;
    font-size: 12px; font-family: "Segoe UI", monospace;
    pointer-events: none; z-index: 9999;
    max-height: 420px; overflow-y: auto;
    border: 1px solid #444466;
    box-shadow: 0 4px 16px rgba(0,0,0,0.5);
    line-height: 1.9; min-width: 220px;
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
      pts.sort(function(a,b) {{ return b.y - a.y; }});
      if (!pts.length) {{ tip.style.display='none'; return; }}
      var timeStr = pts[0].x;
      var html = '<b style="color:#a0a8ff">' + timeStr + '</b><hr class="tip-hr">';
      pts.forEach(function(p) {{
        var col = (p.data.line && p.data.line.color) ? p.data.line.color : '#fff';
        var val = (typeof p.y === 'number')
          ? p.y.toLocaleString(undefined, {{maximumFractionDigits:4}}) : p.y;
        html += '<span style="color:' + col + '">━</span>&nbsp;'
              + p.data.name + ':&nbsp;<b>' + val + '</b><br>';
      }});
      tip.innerHTML = html;
      tip.style.display = 'block';
    }});
    div.on('plotly_unhover', function() {{ tip.style.display = 'none'; }});
    document.addEventListener('mousemove', function(e) {{
      var x = e.clientX + 16, y = e.clientY - 12;
      var tw = tip.offsetWidth || 240, th = tip.offsetHeight || 300;
      if (x + tw > window.innerWidth)  x = e.clientX - tw - 10;
      if (x < 0) x = 0;
      if (y + th > window.innerHeight) y = e.clientY - th - 10;
      if (y < 0) y = 0;
      tip.style.left = x + 'px'; tip.style.top = y + 'px';
    }});
  }});
}})();
</script>
</body></html>
"""


def render_charts(df, time_col, charts, file_prefix=""):
    """Render the time slider + 4 charts for one file."""
    t_min = time_col.min().to_pydatetime()
    t_max = time_col.max().to_pydatetime()
    st.markdown("#### Time range")
    t_start, t_end = st.slider(
        f"time_range_{file_prefix}",
        min_value=t_min, max_value=t_max,
        value=(t_min, t_max), step=timedelta(minutes=1),
        format="%d/%m/%y %H:%M", label_visibility="collapsed",
    )
    for idx, chart_def in enumerate(charts):
        st.subheader(f"Chart {idx+1} — {chart_def['title']}")
        components.html(
            make_chart_html(
                df, time_col, chart_def["cols"], chart_def["subtitle"],
                f"{file_prefix}c{idx}",
                x_range=(t_start, t_end),
                multiply=chart_def.get("multiply"),
                decimals=chart_def.get("decimals"),
            ),
            height=520, scrolling=False,
        )


# ── Load all uploaded files ──────────────────────────────────────────────────────
file_data = []
for f in uploaded_files:
    with st.spinner(f"Loading {f.name}…"):
        df, is_re_export = load_file(f)
        time_col = parse_time(df)
        charts = build_charts(df)
        file_data.append({
            "name":         f.name,
            "stem":         f.name.rsplit(".", 1)[0],
            "df":           df,
            "time_col":     time_col,
            "charts":       charts,
            "is_re_export": is_re_export,
        })
    st.success(f"**{f.name}**: {len(df):,} rows × {len(df.columns)} columns")


# ── Display charts ───────────────────────────────────────────────────────────────
if len(file_data) == 1:
    render_charts(
        file_data[0]["df"], file_data[0]["time_col"],
        file_data[0]["charts"], file_prefix="f0",
    )
else:
    tabs = st.tabs([fd["name"] for fd in file_data])
    for ti, (tab, fd) in enumerate(zip(tabs, file_data)):
        with tab:
            render_charts(fd["df"], fd["time_col"], fd["charts"], file_prefix=f"f{ti}")


# ── Excel export ─────────────────────────────────────────────────────────────────
def _to_xl_serial(ts):
    """Convert a pandas Timestamp to an Excel date serial number."""
    d = ts.to_pydatetime().replace(tzinfo=None) - _dt(1899, 12, 30)
    return d.days + d.seconds / 86400


def _build_excel(fd: dict) -> bytes:
    """Build a complete Excel workbook for one file and return its bytes."""
    df       = fd["df"]
    time_col = fd["time_col"]
    charts   = fd["charts"]

    all_col_indices = sorted(set(
        i for c in charts for i in c["cols"] if i < len(df.columns)
    ))
    sheet_col = {0: 0}
    for pos, oi in enumerate(all_col_indices, start=1):
        sheet_col[oi] = pos
    n = len(df)

    col_transforms: dict = {}
    for chart_def in charts:
        m_val = chart_def.get("multiply")
        d_val = chart_def.get("decimals")
        if m_val is not None or d_val is not None:
            for oi in chart_def["cols"]:
                col_transforms[oi] = (m_val, d_val)

    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})
    fmt_bold = wb.add_format({"bold": True})
    fmt_date = wb.add_format({"num_format": "dd/mm/yyyy hh:mm:ss", "bold": False})

    # ── Data sheet ────────────────────────────────────────────────────────────
    ws_data = wb.add_worksheet("Data")
    ws_data.write(0, 0, df.columns[0], fmt_bold)
    for pos, oi in enumerate(all_col_indices, start=1):
        ws_data.write(0, pos, df.columns[oi], fmt_bold)

    for row_i, (t, *vals) in enumerate(
        zip(time_col, *[df.iloc[:, oi] for oi in all_col_indices]), start=1
    ):
        dt_val = t.to_pydatetime() if pd.notna(t) else None
        if dt_val:
            ws_data.write_datetime(row_i, 0, dt_val, fmt_date)
        for pos, (oi, v) in enumerate(zip(all_col_indices, vals), start=1):
            fv = float(v) if (
                pd.notna(v) and not (isinstance(v, float) and _math.isnan(v))
            ) else None
            if fv is not None:
                tm, td = col_transforms.get(oi, (None, None))
                if tm is not None:
                    fv *= tm
                if td is not None:
                    fv = round(fv, td)
                ws_data.write_number(row_i, pos, fv)

    # ── Charts sheet ──────────────────────────────────────────────────────────
    ws_charts = wb.add_worksheet("Charts")
    CHART_H, CHART_W, CHART_GAP = 500, 960, 520
    _xl_min = _to_xl_serial(time_col.dropna().min())
    _xl_max = _to_xl_serial(time_col.dropna().max())

    for chart_idx, chart_def in enumerate(charts):
        chart = wb.add_chart({"type": "scatter", "subtype": "smooth"})
        valid_cols = [i for i in chart_def["cols"] if i in sheet_col]
        if not valid_cols:
            continue
        for fb_idx, orig_idx in enumerate(valid_cols):
            col_name = df.columns[orig_idx]
            col_ltr  = xl_col_to_name(sheet_col[orig_idx])
            chart.add_series({
                "name":       f"=Data!${col_ltr}$1",
                "categories": f"=Data!$A$2:$A${n+1}",
                "values":     f"=Data!${col_ltr}$2:${col_ltr}${n+1}",
                "line":       {"color": series_color(col_name, fb_idx), "width": 1.25},
                "marker":     {"type": "none"},
            })
        chart.set_title({"name": chart_def["title"]})
        chart.set_x_axis({
            "name": "Time", "date_axis": True, "num_format": "dd/mm/yy hh:mm",
            "min": _xl_min, "max": _xl_max,
            "major_gridlines": {"visible": True, "line": {"color": "#D0D0D0"}},
            "minor_gridlines": {"visible": False},
        })
        chart.set_y_axis({
            "name": "Value",
            "major_gridlines": {"visible": True, "line": {"color": "#D0D0D0"}},
            "minor_gridlines": {"visible": True, "line": {"color": "#EBEBEB"}},
        })
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": CHART_W, "height": CHART_H})
        ws_charts.insert_chart(0, 0, chart, {"x_offset": 0, "y_offset": chart_idx * CHART_GAP})

    ws_meta = wb.add_worksheet("_meta")
    ws_meta.write(0, 0, "battery-analyzer-export-v1")
    ws_meta.hide()

    wb.close()
    return buf.getvalue()


st.divider()
st.subheader("Download as Excel")

# Invalidate cached results when the uploaded file set changes
_file_key = tuple(fd["name"] for fd in file_data)
if st.session_state.get("_excel_key") != _file_key:
    st.session_state.pop("_excel_results", None)
    st.session_state["_excel_key"] = _file_key

if len(file_data) == 1:
    col1, col2 = st.columns([3, 1])
    with col1:
        excel_filename = st.text_input(
            "File name", value=file_data[0]["stem"],
            label_visibility="collapsed",
            placeholder="File name (without .xlsx)",
        )
    with col2:
        generate = st.button("Generate Excel file", use_container_width=True)
    if generate:
        with st.spinner("Building Excel workbook…"):
            data = _build_excel(file_data[0])
        fname = (excel_filename.strip() or file_data[0]["stem"]).removesuffix(".xlsx") + ".xlsx"
        st.session_state["_excel_results"] = [(fname, data)]
else:
    generate = st.button("Generate Excel files")
    if generate:
        results = []
        for fd in file_data:
            with st.spinner(f"Building {fd['stem']}.xlsx…"):
                results.append((fd["stem"] + ".xlsx", _build_excel(fd)))
        st.session_state["_excel_results"] = results

if "_excel_results" in st.session_state:
    for fname, data in st.session_state["_excel_results"]:
        st.download_button(
            label=f"⬇️ Download {fname}",
            data=data,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{fname}",
        )
