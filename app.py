# app.py — Walmart Valuation Explorer
# Clean UI • Header auto-detect • Valuation Summary • Valuation Charts • Sheet-grounded chat

import os, re, json, sys, pkgutil
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ───────── CONFIG ─────────
FILE_NAME = "FIN42030 WMT Valuation (2).xlsx"

# Hard-coded Azure OpenAI (optional; safe fallback if unset)
AZURE_OPENAI_ENDPOINT    = "https://testaisentiment.openai.azure.com/"
AZURE_OPENAI_API_KEY     = "cb1c33772b3c4edab77db69ae18c9a43"
AZURE_OPENAI_API_VERSION = "2024-02-15-preview"
AZURE_OPENAI_DEPLOYMENT  = "aipocexploration"

st.set_page_config(page_title="Walmart Valuation Explorer", page_icon="📊", layout="wide")

# ───────── UI THEME / CSS ─────────
if "ui_theme" not in st.session_state:
    st.session_state.ui_theme = "Light"
with st.sidebar:
    st.markdown("### Appearance")
    st.session_state.ui_theme = st.radio("Theme", ["Light","Dark"], index=0, horizontal=True)

LIGHT = {
    "BG":"#ffffff", "PANEL":"#f7f9fc", "BORDER":"#dde5f0", "TEXT":"#0b1220", "MUTED":"#44536a", "PRIMARY":"#0f6fff"
}
DARK = {
    "BG":"#0b1220", "PANEL":"#0f172a", "BORDER":"#233043", "TEXT":"#e8f1ff", "MUTED":"#9ab0cf", "PRIMARY":"#22d3ee"
}
C = LIGHT if st.session_state.ui_theme=="Light" else DARK

st.markdown(f"""
<style>
  :root {{
    --bg:{C["BG"]}; --panel:{C["PANEL"]}; --border:{C["BORDER"]};
    --text:{C["TEXT"]}; --muted:{C["MUTED"]}; --primary:{C["PRIMARY"]};
  }}
  html, body, .block-container {{
    background:var(--bg); color:var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    font-size: 18px; line-height: 1.5;
  }}
  .headline {{ font-size:2.1rem; font-weight:900; letter-spacing:-.01em; margin:0 0 .25rem 0; }}
  .soft {{ color:var(--muted); }}
  .card {{ background:var(--panel); border:1px solid var(--border); border-radius:18px; padding:16px; }}
  .kpi {{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:16px; }}
  .kpi h4 {{ margin:.1rem 0 .4rem 0; font-size:.95rem; color:var(--muted); }}
  .kpi .v {{ font-size:1.6rem; font-weight:900; }}
  .kpi .d {{ font-size:.92rem; color:var(--muted); }}
  div[data-testid="stChatMessage"] p {{ font-size:1rem; }}
  div[data-baseweb="select"] span {{ white-space:normal !important; }}
</style>
""", unsafe_allow_html=True)

# ───────── OPTIONAL LLM ─────────
OPENAI_OK = False
client = None
if AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY and AZURE_OPENAI_DEPLOYMENT:
    try:
        from openai import AzureOpenAI
        client = AzureOpenAI(
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            api_key=AZURE_OPENAI_API_KEY,
            api_version=AZURE_OPENAI_API_VERSION,
        )
        OPENAI_OK = True
    except Exception:
        OPENAI_OK = False

def ask_gpt(messages, temperature=0.2, max_tokens=900):
    if not OPENAI_OK:
        return "(Chat disabled — missing Azure OpenAI env vars.)"
    try:
        r = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT, messages=messages,
            temperature=temperature, max_tokens=max_tokens
        )
        return r.choices[0].message.content
    except Exception as e:
        return f"(LLM error: {e})"

# ───────── HELPERS: detection & reshaping ─────────
def need_openpyxl() -> bool:
    return pkgutil.find_loader("openpyxl") is None

YEAR_PAT = re.compile(r"^(19|20)\d{2}$")

def _row_has_years(row_vals) -> int:
    cnt = 0
    for v in row_vals:
        s = str(v).strip()
        if YEAR_PAT.match(s): cnt += 1
    return cnt

def detect_header_row(df: pd.DataFrame, scan_rows: int = 12) -> int | None:
    scan = min(scan_rows, len(df))
    for i in range(scan):
        if _row_has_years(df.iloc[i].tolist()) >= 3:
            return i
    return None

def normalize_sheet(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    # detect header row (years)
    header_row = detect_header_row(work)
    if header_row is not None:
        new_cols = work.iloc[header_row].astype(str).str.strip().tolist()
        work = work.iloc[header_row + 1:].reset_index(drop=True)
        work.columns = new_cols
    # remove empty columns
    empties = [c for c in work.columns if work[c].isna().all()]
    if empties: work = work.drop(columns=empties)
    # first column as 'Line'
    if len(work.columns):
        first = str(work.columns[0]).strip().lower()
        if first.startswith("unnamed") or first in ("", "nan"):
            work.columns = ["Line"] + [str(c).strip() for c in work.columns[1:]]
        else:
            if work.iloc[:10, 0].astype(str).str.len().mean() > 2:
                work.rename(columns={work.columns[0]: "Line"}, inplace=True)
    # clean column names + kill lingering 'Unnamed'
    clean_cols = []
    for c in work.columns:
        s = str(c).strip()
        clean_cols.append("Line" if s.lower().startswith("unnamed") or s in ("", "nan") else s)
    work.columns = clean_cols
    if "Line" in work.columns:
        work["Line"] = work["Line"].astype(str).str.strip()
    return work

@st.cache_data(show_spinner=False)
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    if need_openpyxl():
        raise ImportError("openpyxl not installed. Add to requirements.txt and redeploy.")
    xl = pd.ExcelFile(path, engine="openpyxl")
    out = {}
    for name in xl.sheet_names:
        raw = xl.parse(name, header=None)  # read without trusting row 1
        out[name] = normalize_sheet(raw)
    return out

def is_date_col(s: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(s): return True
    if s.dtype == object:
        try:
            parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
            return parsed.notna().mean() >= 0.6
        except Exception: return False
    return False

def num_cols(df):  return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
def date_cols(df): return [c for c in df.columns if is_date_col(df[c])]

def find_year_header_cols(df: pd.DataFrame) -> List[str]:
    pat = re.compile(r"(?:19|20)\d{2}")
    yearish = [c for c in df.columns if pat.search(str(c))]
    if not yearish: return []
    tmp = df.copy()
    for c in yearish: tmp[c] = pd.to_numeric(tmp[c], errors="coerce")
    return [c for c in yearish if tmp[c].notna().any()]

def wide_years_to_long(df: pd.DataFrame, year_cols: List[str], label_col=None) -> pd.DataFrame | None:
    if not year_cols: return None
    non_year = [c for c in df.columns if c not in year_cols]
    if label_col is None: label_col = non_year[0] if non_year else None
    w = df.copy()
    for yc in year_cols: w[yc] = pd.to_numeric(w[yc], errors="coerce")
    long = w.melt(id_vars=[label_col] if label_col else None, value_vars=year_cols,
                  var_name="Year", value_name="Value")
    long["Year"] = long["Year"].astype(str).str.extract(r"((?:19|20)\d{{2}})").astype(float)
    long = long.dropna(subset=["Year","Value"])
    if long.empty: return None
    long["Year"] = long["Year"].astype(int)
    # limit to most varying series to avoid spaghetti
    if label_col and label_col in long.columns:
        var_rank = long.groupby(label_col)["Value"].var().sort_values(ascending=False)
        long = long[long[label_col].isin(list(var_rank.head(6).index))]
    return long

def safe_preview(df: pd.DataFrame, n=8) -> str:
    try:
        import tabulate as _  # noqa
        return df.head(n).to_markdown(index=False)
    except Exception:
        return df.head(n).to_string(index=False)

# ───────── Valuation extractor & smart summary ─────────
def _first_number_to_right(row: pd.Series) -> float | None:
    for v in row[1:]:
        try:
            x = float(str(v).replace(",", "").replace("%",""))
            return x
        except Exception:
            continue
    return None

VAL_PATTERNS = {
    "wacc": re.compile(r"\bwacc\b", re.I),
    "g": re.compile(r"(^g$|terminal growth|lt growth)", re.I),
    "coeq": re.compile(r"cost of equity", re.I),
    "fcff": re.compile(r"^fcff$", re.I),
    "fcfe": re.compile(r"^fcfe$", re.I),
    "pv_fcff": re.compile(r"pv of fcff", re.I),
    "pv_fcfe": re.compile(r"pv of fcfe", re.I),
    "enterprise_value": re.compile(r"(enterprise.*value)|(firm.*value)", re.I),
    "equity_value": re.compile(r"^equity value$|^total equity$", re.I),
    "debt": re.compile(r"^total debt", re.I),
    "shares": re.compile(r"number of outstanding shares", re.I),
    "pps": re.compile(r"^price per share$", re.I),
    "pps_current": re.compile(r"current price per share", re.I),
}

def extract_valuation(df: pd.DataFrame) -> dict | None:
    if "Line" not in df.columns: return None
    vals = {k: None for k in VAL_PATTERNS.keys()}
    scan = min(len(df), 200)
    for i in range(scan):
        label = str(df.iloc[i, 0]).strip()
        if not label or label.lower().startswith("nan"): 
            continue
        for key, pat in VAL_PATTERNS.items():
            if pat.search(label):
                vals[key] = _first_number_to_right(df.iloc[i])
    if all(vals[k] is None for k in ("pps","equity_value","enterprise_value","wacc","fcff","fcfe")):
        return None
    for k in ("wacc","g","coeq"):
        if vals[k] is not None and vals[k] > 1.0:  # interpret 7.5 as 7.5%
            vals[k] = vals[k] / 100.0
    return vals

def sheet_summary_smart(df: pd.DataFrame, name: str) -> str:
    facts = extract_valuation(df)
    if facts:
        parts = [f"**{name} — Valuation Summary**"]
        if facts.get("wacc") is not None:
            parts.append(f"- WACC: **{facts['wacc']*100:.2f}%**")
        if facts.get("coeq") is not None:
            parts.append(f"- Cost of Equity: **{facts['coeq']*100:.2f}%**")
        if facts.get("g") is not None:
            parts.append(f"- Terminal growth (g): **{facts['g']*100:.2f}%**")

        pps = facts.get("pps"); cur = facts.get("pps_current")
        if pps is not None:
            msg = f"- **FCFF price per share:** **${pps:,.2f}**"
            if cur is not None:
                up = (pps/cur - 1.0)*100
                msg += f" vs current **${cur:,.2f}** → **{up:+.1f}%**"
            parts.append(msg)

        if facts.get("equity_value") is not None:
            parts.append(f"- Equity value: **${facts['equity_value']:,.0f}**")
        if facts.get("enterprise_value") is not None:
            parts.append(f"- Enterprise (firm) value: **${facts['enterprise_value']:,.0f}**")
        if facts.get("shares") is not None:
            parts.append(f"- Shares outstanding: **{facts['shares']:,.0f}**")
        return "\n".join(parts)

    # fallback generic
    cols = [str(c) for c in df.columns]
    dcols = date_cols(df); ycols = find_year_header_cols(df)
    axis_hint = "date column" if dcols else ("years in headers" if ycols else "no time axis")
    return f"**{name} — Sheet Summary**\n- Columns: {', '.join(cols[:15])}{'…' if len(cols)>15 else ''}\n- Chart basis: {axis_hint}"

# ───────── Valuation chart block (forced) ─────────
def build_valuation_long(df: pd.DataFrame) -> pd.DataFrame | None:
    """
    Return long-form dataframe for FCFF/FCFE across years:
      columns: Line, Year, Value
    """
    if "Line" not in df.columns: return None
    ycols = find_year_header_cols(df)
    if not ycols: return None
    long = wide_years_to_long(df, ycols, label_col="Line")
    if long is None or long.empty: return None
    # Keep FCFF / FCFE rows
    mask = long["Line"].str.contains(r"^fcff$|^fcfe$", case=False, regex=True)
    out = long[mask].copy()
    return out if not out.empty else None

def show_valuation_block(df: pd.DataFrame):
    st.markdown("### Valuation (FCFF / FCFE)")
    long = build_valuation_long(df)
    facts = extract_valuation(df)

    if long is None:
        st.caption("Could not find FCFF/FCFE with year columns in this sheet.")
        return

    template = "plotly_white" if st.session_state.ui_theme=="Light" else "plotly_dark"
    fig = px.line(long, x="Year", y="Value", color="Line", markers=True)
    fig.update_traces(line=dict(width=3))
    fig.update_layout(height=420, template=template, margin=dict(l=10,r=10,t=10,b=10),
                      legend_title_text="")
    # annotate price per share from FCFF / current price if available
    if facts:
        pps = facts.get("pps"); cur = facts.get("pps_current")
        annot_y = long["Value"].max() * 1.02 if len(long) else None
        ann = []
        if pps is not None:
            ann.append(f"FCFF P/S: ${pps:,.2f}")
        if cur is not None:
            ann.append(f"Current: ${cur:,.2f}")
        if ann and annot_y is not None:
            fig.add_annotation(
                xref="paper", yref="y",
                x=1.0, y=annot_y,
                text=" • ".join(ann),
                showarrow=False, font=dict(size=14)
            )
    st.plotly_chart(fig, use_container_width=True)
# ─────────────────────────────────────────────────────

# ───────── LOAD WORKBOOK ─────────
st.markdown(f'<div class="headline">Walmart Valuation Explorer</div><div class="soft">Python: {sys.executable}</div>', unsafe_allow_html=True)

if not os.path.exists(FILE_NAME):
    st.error(f"File not found: {FILE_NAME} (place it next to app.py).")
    st.stop()

try:
    dfs = load_workbook(FILE_NAME)
except Exception as e:
    st.error(str(e)); st.stop()

# ───────── SIDEBAR SHEETS ─────────
st.sidebar.header("Sheets")
flt = st.sidebar.text_input("Filter sheets", "")
sheet_names = [n for n in dfs.keys() if flt.lower() in n.lower()] or list(dfs.keys())
selected = st.sidebar.selectbox("Select a sheet", sheet_names, index=0)

# ───────── OVERVIEW (placeholder KPIs) ─────────
st.markdown("## Overview")
k1,k2,k3,k4 = st.columns(4)
for k in (k1,k2,k3,k4):
    with k: st.markdown('<div class="kpi"><h4>—</h4><div class="v">—</div></div>', unsafe_allow_html=True)

# ───────── SHEET VIEW ─────────
st.markdown(f"## {selected}")
df = dfs[selected].copy()

# remove fully-empty unnamed columns (rare after normalization)
unnamed = [c for c in df.columns if str(c).lower().startswith("unnamed") and df[c].isna().all()]
if unnamed: df = df.drop(columns=unnamed)

with st.expander("Preview", expanded=False):
    st.dataframe(df, use_container_width=True, height=420)

with st.expander("What’s in this sheet?", expanded=True):
    st.markdown(sheet_summary_smart(df, selected))

# ---- Standard charts
st.markdown("### Charts")
dcols = date_cols(df); ncols = num_cols(df); ycols = find_year_header_cols(df)

if dcols:
    x = dcols[0]
    dfx = df.copy(); dfx[x] = pd.to_datetime(dfx[x], errors="coerce"); dfx = dfx.dropna(subset=[x]).sort_values(x)
    dfx, ncols = (lambda x, cols: (x.assign(**{c:pd.to_numeric(x[c],errors='coerce') for c in cols}), [c for c in cols if pd.to_numeric(x[c],errors='coerce').notna().any()]))(dfx, ncols)
    targets = ncols[:3]
    if targets:
        choose = st.multiselect("Series", options=ncols, default=targets)
        fig = px.line(dfx, x=x, y=choose)
        fig.update_layout(height=380, template="plotly_white" if st.session_state.ui_theme=="Light" else "plotly_dark",
                          margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No numeric columns to plot against the date column.")
elif ycols:
    long = wide_years_to_long(df, ycols, label_col="Line" if "Line" in df.columns else None)
    if long is not None and not long.empty:
        label_cols=[c for c in long.columns if c not in ("Year","Value")]
        if label_cols:
            label = label_cols[0]
            options = sorted(long[label].unique().tolist()); default = options[:min(5,len(options))]
            sel = st.multiselect("Series", options, default)
            sub = long[long[label].isin(sel)] if sel else long
            fig = px.line(sub, x="Year", y="Value", color=label, markers=True)
        else:
            fig = px.line(long, x="Year", y="Value", markers=True)
        fig.update_layout(height=380, template="plotly_white" if st.session_state.ui_theme=="Light" else "plotly_dark",
                          margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Couldn’t reshape year columns into a time series.")
else:
    st.info("No date or year structure detected for a time series.")

# ---- FORCED Valuation block
show_valuation_block(df)
st.divider()

# ───────── ANALYST CHAT ─────────
st.markdown("## Analyst Chat")
st.caption("Ask things like: “summarise this sheet”, “key takeaways”, “what’s WACC and implied upside?”")

if "chat" not in st.session_state:
    st.session_state.chat=[{"role":"assistant","content":"Hi! I’ll summarise the active sheet with real numbers when you ask."}]

for m in st.session_state.chat:
    with st.chat_message("assistant" if m["role"]=="assistant" else "user"):
        st.write(m["content"])

SYSTEM = """You are a valuation analyst. Prefer concrete numbers from the provided SHEET FACTS.
If you use workbook numbers, cite the sheet name once in the answer. Keep answers tight."""

def answer(user_q: str, df: pd.DataFrame, sheet_name: str) -> str:
    qlow = (user_q or "").lower()
    if any(k in qlow for k in ["summarise","summarize","summary","key takeaway","what is the analysis"]):
        return sheet_summary_smart(df, sheet_name)

    # fallback to LLM with computed context
    prev = safe_preview(df, n=8)
    facts = sheet_summary_smart(df, sheet_name)
    context = f"{facts}\n\nPreview (first 8 rows):\n{prev}"
    messages=[
        {"role":"system","content":SYSTEM},
        {"role":"user","content": f"SHEET FACTS:\n{context}\n\nQuestion: {user_q}"}
    ]
    out = ask_gpt(messages)
    if out.startswith("("):
        return "Chat is disabled (no Azure OpenAI env vars). Use charts and summary above."
    if f"**{sheet_name}**" not in out:
        out += f"\n\n(Source: **{sheet_name}**)"
    return out

prompt = st.chat_input("Ask about the selected sheet…")
if prompt:
    st.session_state.chat.append({"role":"user","content":prompt})
    st.session_state.chat.append({"role":"assistant","content":answer(prompt, df, selected)})
    st.rerun()
