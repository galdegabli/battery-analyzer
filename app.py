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

st.set_page_config(page_title="LOD Analyzer", layout="wide")
st.title("LOD Analyzer")

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


def detect_charge_cycles(soc, min_len=50):
    """Label charging rows with a cycle ID (1-based). NaN = discharge/idle."""
    s = pd.to_numeric(soc, errors="coerce").reset_index(drop=True)
    # 5-step diff smooths small noise; positive = SOC rising = charging
    charging = s.diff(periods=5).fillna(0) > 0
    boundary = (charging != charging.shift(fill_value=False)).cumsum()
    result = pd.Series(float("nan"), index=s.index)
    cyc = 0
    for _bid, grp in charging.groupby(boundary):
        if not grp.iloc[0]:
            continue
        if len(grp) >= min_len:
            cyc += 1
            result.iloc[grp.index] = float(cyc)
    return result


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


def find_time_col_idx(df) -> int:
    """Return the index of the date/time column by name, defaulting to 0."""
    for i, col in enumerate(df.columns):
        if re.search(r"date|time", str(col), re.IGNORECASE):
            return i
    return 0


def parse_time(df):
    """Parse the date/time column (found by name) as datetime."""
    idx = find_time_col_idx(df)
    raw = df.iloc[:, idx].astype(str).str.strip()
    # Collapse any run of whitespace to a single space
    normalised = raw.str.replace(r"\s+", " ", regex=True)
    sample = normalised.dropna().iloc[0] if not normalised.dropna().empty else ""

    # ISO format: YYYY-MM-DD HH:MM:SS
    if re.match(r"\d{4}-\d{2}-\d{2}", sample):
        result = pd.to_datetime(normalised, format="%Y-%m-%d %H:%M:%S", errors="coerce")
        if result.notna().any():
            return result

    # Try explicit DD/MM/YYYY %H:%M:%S (zero-padded hours, single space)
    result = pd.to_datetime(normalised, format="%d/%m/%Y %H:%M:%S", errors="coerce")
    if result.notna().any():
        return result

    # Zero-pad single-digit hour: " 8:" → " 08:", then retry
    zero_padded = normalised.str.replace(r"(?<= )(\d)(?=:)", r"0\1", regex=True)
    result = pd.to_datetime(zero_padded, format="%d/%m/%Y %H:%M:%S", errors="coerce")
    if result.notna().any():
        return result

    # Fallback: pandas auto-detect with dayfirst=True
    try:
        return pd.to_datetime(normalised, format="mixed", dayfirst=True, errors="coerce")
    except TypeError:
        return pd.to_datetime(normalised, dayfirst=True, errors="coerce")


def find_cols_in(df, pattern):
    def num_key(i):
        m = re.search(r"(\d+)", df.columns[i][len(pattern):])
        return int(m.group(1)) if m else 0
    return sorted(
        [i for i, c in enumerate(df.columns) if c.startswith(pattern)],
        key=num_key,
    )


def find_col_in(df, name):
    target = name.strip().lower()
    return [i for i, c in enumerate(df.columns) if str(c).strip().lower() == target][:1]


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
    # Pre-compute column groups for charts that need format detection
    _v_old   = find_cols_in(df, "CellVoltage_")           # old format, stored in V
    _v_new   = find_cols_in(df, "Cell Voltage [mV]_")     # new format, stored in mV
    _v_mm    = (find_col_in(df, "Max Cell Voltage [mV]")
                + find_col_in(df, "Min Cell Voltage [mV]"))
    _afe_old = find_col_in(df, "BMS_AFE_current [10mA]")  # old format, in 10 mA
    _afe_new = find_col_in(df, "AFE Current [A]")          # new format, already in A
    _t_old   = find_cols_in(df, "CellTempSensor_")
    _t_new   = find_cols_in(df, "Cell Temp Sensor [C]_")

    charts = [
        {
            "title":    "Full Charge Capacity",
            "cols":     (find_col_in(df, "SE Full_Charge_Capacity [Ah]")
                         + find_col_in(df, "SE_Full_Charge_Capacity [10mAh]")
                         + find_col_in(df, "Full_Charge_Capacity [Ah]")),
            "subtitle": "Full Charge Capacity vs Time",
            "col_multiplies": {i: 0.01 for i in find_col_in(df, "SE_Full_Charge_Capacity [10mAh]")},
            "y_label":  "Capacity [Ah]",
        },
        {
            "title":    "Cell Voltage",
            "cols":     _v_old + _v_new + _v_mm,
            "subtitle": "Cell Voltages vs Time",
            "multiply":  _auto_multiply(df, _v_old) if _v_old else None,
            "col_multiplies": {i: 0.001 for i in _v_new + _v_mm},
            "hidden_cols":    set(_v_mm),
            "y_label":        "Voltage [V]",
        },
        {
            "title":    "Cell Distance",
            "cols":     find_cols_in(df, "CellDistanceAh_"),
            "subtitle": "CellDistanceAh_0 – CellDistanceAh_14 vs Time",
            "y_label":  "Distance [Ah]",
        },
        {
            "title":    "Voltage Derivative",
            "cols":     find_cols_in(df, "CellFdFVdQ_"),
            "subtitle": "CellFdFVdQ_0 – CellFdFVdQ_14 vs Time",
            "decimals": 0,
        },
        {
            "title":    "Current",
            "cols":     _afe_old + _afe_new,
            "subtitle": "Current vs Time",
            "col_multiplies": {i: 0.01 for i in _afe_old},
            "y_label":  "Current [A]",
        },
        {
            "title":    "Cell Temperature",
            "cols":     _t_old + _t_new,
            "subtitle": "Cell Temperature vs Time",
            "y_label":  "Temperature [°C]",
            "add_avg":  True,
        },
        {
            "title":         "Derivative vs SOC",
            "cols":          find_cols_in(df, "CellFdFVdQ_"),
            "subtitle":      "CellFdFVdQ_0 – CellFdFVdQ_14 vs SOC",
            "decimals":      0,
            "y_label":       "dV/dQ",
            "x_col_indices": find_col_in(df, "SE_SOC [0.01%]"),
            "x_multiply":    0.01,
            "x_label":       "SOC [%]",
            "detect_cycles": True,
        },
    ]
    return charts


# ── Chart HTML builder (custom sorted hover via JS) ──────────────────────────────
def make_chart_html(df, time_col, col_indices, title, chart_id,
                    x_range=None, multiply=None, decimals=None, y_label="Value",
                    col_multiplies=None, hidden_cols=None, add_avg=False,
                    x_label="Time", x_is_date=True):
    fig = go.Figure()
    for fb_idx, i in enumerate(col_indices):
        if i >= len(df.columns):
            continue
        col_name = df.columns[i]
        y = pd.to_numeric(df.iloc[:, i], errors="coerce")
        m = (col_multiplies or {}).get(i, multiply)
        if m is not None:
            y = y * m
        if decimals is not None:
            y = y.round(decimals)
        is_hidden = hidden_cols and i in hidden_cols
        fig.add_trace(go.Scatter(
            x=time_col, y=y, mode="lines", name=col_name,
            line=dict(color=series_color(col_name, fb_idx)),
            hoverinfo="none",
            visible="legendonly" if is_hidden else True,
        ))
    if add_avg and col_indices:
        valid = [i for i in col_indices if i < len(df.columns)]
        if valid:
            all_y = pd.concat(
                [pd.to_numeric(df.iloc[:, i], errors="coerce") * ((col_multiplies or {}).get(i, multiply) or 1)
                 for i in valid], axis=1
            )
            y_avg = all_y.mean(axis=1)
            if decimals is not None:
                y_avg = y_avg.round(decimals)
            fig.add_trace(go.Scatter(
                x=time_col, y=y_avg, mode="lines", name="Average",
                line=dict(color="#000000", width=2.5, dash="dash"),
                hoverinfo="none",
            ))
    if x_is_date:
        xaxis_cfg = dict(type="date", rangeslider=dict(visible=False), title=x_label)
        if x_range:
            xaxis_cfg["range"] = [x_range[0].isoformat(), x_range[1].isoformat()]
    else:
        xaxis_cfg = dict(title=x_label)
    fig.update_layout(
        title=dict(text=title, y=0.97, x=0.5, xanchor="center", yanchor="top"),
        xaxis=xaxis_cfg,
        yaxis=dict(fixedrange=False, title=y_label),
        legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5,
                    entrywidth=160, entrywidthmode="pixels"),
        hovermode="x", height=560,
        margin=dict(t=40, b=130, l=60, r=20),
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
    valid_times = time_col.dropna().sort_values()
    if valid_times.empty:
        st.error("No valid timestamps found in the time column.")
        return

    # Use median ± 10 years to clip extreme outlier timestamps
    median_time = valid_times.iloc[len(valid_times) // 2]
    clip_delta = pd.Timedelta(days=365 * 10)
    clipped = valid_times[(valid_times >= median_time - clip_delta) &
                          (valid_times <= median_time + clip_delta)]
    if clipped.empty:
        clipped = valid_times  # fallback: use all

    t_min = clipped.min().to_pydatetime()
    t_max = clipped.max().to_pydatetime()

    st.markdown("#### Time range")
    t_start, t_end = st.slider(
        f"time_range_{file_prefix}",
        min_value=t_min, max_value=t_max,
        value=(t_min, t_max), step=timedelta(minutes=1),
        label_visibility="collapsed",
    )
    for idx, chart_def in enumerate(charts):
        if not chart_def["cols"]:
            st.markdown(
                f'<p style="opacity:0.35;font-size:1.15em;font-weight:600;margin:0.6em 0">'
                f'Chart {idx+1} — {chart_def["title"]}'
                f' &nbsp;<em style="font-weight:normal;font-size:0.85em">'
                f'(no matching data in this file)</em></p>',
                unsafe_allow_html=True,
            )
            continue
        st.subheader(f"Chart {idx+1} — {chart_def['title']}")
        x_col_indices = chart_def.get("x_col_indices")
        if x_col_indices:
            xi = x_col_indices[0]
            x_data = pd.to_numeric(df.iloc[:, xi], errors="coerce") * chart_def.get("x_multiply", 1)
            if chart_def.get("detect_cycles") and xi < len(df.columns):
                cycle_ids = detect_charge_cycles(x_data)
                n_cycles = int(cycle_ids.max()) if cycle_ids.notna().any() else 0
                if n_cycles > 0:
                    all_labels = [f"Cycle {i}" for i in range(1, n_cycles + 1)]
                    selected = st.multiselect(
                        f"cycles_{file_prefix}_{idx}",
                        options=all_labels, default=all_labels,
                        label_visibility="collapsed",
                    )
                    sel_nums = {int(s.split()[1]) for s in selected}
                    mask = cycle_ids.isin(sel_nums)
                    df_plot = df[mask].reset_index(drop=True)
                    x_arg = x_data[mask].reset_index(drop=True)
                else:
                    df_plot, x_arg = df, x_data
            else:
                df_plot, x_arg = df, x_data
            x_range_arg, x_is_date = None, False
        else:
            df_plot = df
            x_arg, x_range_arg, x_is_date = time_col, (t_start, t_end), True
        components.html(
            make_chart_html(
                df_plot, x_arg, chart_def["cols"], chart_def["subtitle"],
                f"{file_prefix}c{idx}",
                x_range=x_range_arg,
                multiply=chart_def.get("multiply"),
                decimals=chart_def.get("decimals"),
                y_label=chart_def.get("y_label", "Value"),
                col_multiplies=chart_def.get("col_multiplies"),
                hidden_cols=chart_def.get("hidden_cols"),
                add_avg=chart_def.get("add_avg", False),
                x_label=chart_def.get("x_label", "Time"),
                x_is_date=x_is_date,
            ),
            height=630, scrolling=False,
        )


# ── Load all uploaded files ──────────────────────────────────────────────────────
file_data = []
for f in uploaded_files:
    with st.spinner(f"Loading {f.name}…"):
        df, is_re_export = load_file(f)
        time_col_idx = find_time_col_idx(df)
        time_col = parse_time(df)
        charts = build_charts(df)
        file_data.append({
            "name":         f.name,
            "stem":         f.name.rsplit(".", 1)[0],
            "df":           df,
            "time_col":     time_col,
            "time_col_idx": time_col_idx,
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
    df           = fd["df"]
    time_col     = fd["time_col"]
    charts       = fd["charts"]
    time_col_idx = fd.get("time_col_idx", 0)
    time_col_ltr = xl_col_to_name(time_col_idx)
    n            = len(df)

    # Transforms only apply to specific chart columns
    col_transforms: dict = {}
    for chart_def in charts:
        m_val = chart_def.get("multiply")
        d_val = chart_def.get("decimals")
        col_mult_overrides = chart_def.get("col_multiplies", {})
        for oi in chart_def["cols"]:
            m = col_mult_overrides.get(oi, m_val)
            if m is not None or d_val is not None:
                col_transforms[oi] = (m, d_val)

    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {"in_memory": True})
    fmt_bold = wb.add_format({"bold": True})
    fmt_date = wb.add_format({"num_format": "dd/mm/yyyy hh:mm:ss", "bold": False})

    # ── Data sheet: ALL original columns ─────────────────────────────────────
    ws_data = wb.add_worksheet("Data")

    # Header row
    for ci, col_name in enumerate(df.columns):
        ws_data.write(0, ci, str(col_name), fmt_bold)

    # Time column — write as datetime at its actual position
    for row_i, t in enumerate(time_col, start=1):
        dt_val = t.to_pydatetime() if pd.notna(t) else None
        if dt_val:
            ws_data.write_datetime(row_i, time_col_idx, dt_val, fmt_date)

    # All other columns — numeric where possible, fall back to original text
    for col_i in range(len(df.columns)):
        if col_i == time_col_idx:
            continue  # already written as datetime
        col_series = df.iloc[:, col_i]
        numeric = pd.to_numeric(col_series, errors="coerce")
        tm, td = col_transforms.get(col_i, (None, None))
        if tm is not None:
            numeric = numeric * tm
        if td is not None:
            numeric = numeric.round(td)
        values = []
        for num_v, orig_v in zip(numeric, col_series):
            if pd.notna(num_v):
                values.append(float(num_v))
            elif pd.notna(orig_v):
                values.append(str(orig_v))
            else:
                values.append(None)
        ws_data.write_column(1, col_i, values)

    # ── Charts sheet ──────────────────────────────────────────────────────────
    ws_charts = wb.add_worksheet("Charts")
    CHART_H, CHART_W, CHART_GAP = 500, 960, 520
    _xl_min = _to_xl_serial(time_col.dropna().min())
    _xl_max = _to_xl_serial(time_col.dropna().max())

    for chart_idx, chart_def in enumerate(charts):
        chart = wb.add_chart({"type": "scatter", "subtype": "smooth"})
        valid_cols = [i for i in chart_def["cols"] if i < len(df.columns)]
        if not valid_cols:
            continue
        for fb_idx, orig_idx in enumerate(valid_cols):
            col_name = df.columns[orig_idx]
            col_ltr  = xl_col_to_name(orig_idx)   # sheet col = df col index
            chart.add_series({
                "name":       f"=Data!${col_ltr}$1",
                "categories": f"=Data!${time_col_ltr}$2:${time_col_ltr}${n+1}",
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
