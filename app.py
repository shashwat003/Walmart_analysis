# app.py — Walmart Workbook Analyzer (Analyst-first, no downloads, no LLM)
import io, os, json, re
from typing import Dict, List
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Walmart Workbook Analyzer", page_icon="📊", layout="wide")

# ---- CONFIG ----
DEFAULT_XLSX_PATH = "FIN42030 WMT Valuation (2).xlsx"   # put your file next to app.py

# ---- THEME ----
BG, PANEL, BORDER, TEXT, MUTED = "#0b1220", "#111827", "#22314a", "#eef2f7", "#9fb0c7"
PRIMARY = "#22d3ee"
st.markdown(f"""
<style>
  html, body, .block-container {{ background:{BG}; color:{TEXT}; }}
  .card {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:18px; }}
  .headline {{ font-size:1.8rem; font-weight:800; }}
  .soft {{ color:{MUTED}; }}
</style>
""", unsafe_allow_html=True)

# ---- HELPERS ----
def need_openpyxl():
    try:
        import openpyxl  # noqa
        return False
    except Exception:
        return True

@st.cache_data(show_spinner=False)
def load_workbook(file_bytes_or_path) -> Dict[str, pd.DataFrame]:
    if need_openpyxl():
        raise ImportError("openpyxl is required to read .xlsx files. Install it and restart.")
    if isinstance(file_bytes_or_path, (str, os.PathLike)):
        xl = pd.ExcelFile(file_bytes_or_path, engine="openpyxl")
    else:
        xl = pd.ExcelFile(io.BytesIO(file_bytes_or_path), engine="openpyxl")
    return {name: xl.parse(name) for name in xl.sheet_names}

def is_date_col(s: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(s): return True
    if s.dtype == object:
        try: return pd.to_datetime(s, errors="coerce").notna().mean() >= 0.6
        except Exception: return False
    return False

def num_cols(df): return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
def date_cols(df): return [c for c in df.columns if is_date_col(df[c])]
def small_cat_cols(df, max_u=20):
    out=[]
    for c in df.columns:
        if c in num_cols(df): continue
        try:
            u = df[c].nunique(dropna=True)
            if 1 < u <= max_u: out.append(c)
        except Exception: pass
    return out

def sheet_profile(df: pd.DataFrame, name: str) -> dict:
    stats={}
    for c in df.columns:
        s=df[c]
        if pd.api.types.is_numeric_dtype(s):
            stats[str(c)] = {
                "count": int(s.count()),
                "mean": float(np.nanmean(s)) if s.count() else np.nan,
                "min":  float(np.nanmin(s))  if s.count() else np.nan,
                "max":  float(np.nanmax(s))  if s.count() else np.nan,
            }
    with pd.option_context("display.max_columns", 10, "display.width", 1000):
        preview = df.head(8).to_markdown(index=False)
    return {"name":name,"rows":int(df.shape[0]),"cols":int(df.shape[1]),
            "columns":[str(x) for x in df.columns],"num_stats":stats,"preview":preview}

def guess_overview_metrics(dfs: Dict[str, pd.DataFrame]) -> dict:
    """Heuristics to find common metrics from typical sheet names/columns."""
    m = {"revenue":None, "ebitda":None, "fcf":None}
    # Look through likely statement/summary sheets
    for sname, df in dfs.items():
        low = sname.lower()
        if any(k in low for k in ["summary","income","p&l","statement","fcf","cash"]):
            cols = [str(c).lower() for c in df.columns]
            # try to find revenue
            cand = [c for c in cols if "revenue" in c or "sales" in c or "turnover" in c]
            if cand and m["revenue"] is None: m["revenue"] = (sname, cand[0])
            # ebitda
            cand = [c for c in cols if "ebitda" in c]
            if cand and m["ebitda"] is None: m["ebitda"] = (sname, cand[0])
            # fcf
            cand = [c for c in cols if "free cash" in c or "fcf" in c]
            if cand and m["fcf"] is None: m["fcf"] = (sname, cand[0])
    return m

# ---- STATE & LOAD ----
st.markdown(f"""
<div style="display:flex;gap:14px;align-items:center;">
  <div class="headline">Walmart Workbook Analyzer</div>
</div>
<div class="soft">Overview KPIs + per-sheet auto charts. Put your Excel next to this script or upload it below.</div>
""", unsafe_allow_html=True)

dfs = {}
load_error = None
# try autoload from DEFAULT_XLSX_PATH first
if os.path.exists(DEFAULT_XLSX_PATH):
    try:
        dfs = load_workbook(DEFAULT_XLSX_PATH)
        st.success(f"Loaded: {DEFAULT_XLSX_PATH}")
    except Exception as e:
        load_error = str(e)

# uploader (overrides autoload if used)
up = st.file_uploader("Upload Excel model (.xlsx)", type=["xlsx"])
if up:
    try:
        dfs = load_workbook(up.getvalue())
        st.success("Workbook loaded from upload.")
        load_error = None
    except Exception as e:
        load_error = str(e)

if load_error:
    st.error(load_error)

if not dfs:
    st.warning("No workbook loaded yet. Install `openpyxl`, then either keep your file next to app.py or upload it here.")
    st.stop()

# ---- OVERVIEW ----
st.markdown("## Overview")
metrics = guess_overview_metrics(dfs)

# KPI cards (best-effort; may be None if columns not detected)
c1,c2,c3 = st.columns(3)
def kpi_from(sheet_col):
    if not sheet_col: return ("—","—")
    sname, col = sheet_col
    df = dfs[sname].copy()
    # find a time axis
    dcols = date_cols(df)
    if dcols:
        tcol = dcols[0]; df[tcol] = pd.to_datetime(df[tcol], errors="coerce"); df = df.dropna(subset=[tcol]).sort_values(tcol)
        series = pd.to_numeric(df[col], errors="coerce").dropna()
    else:
        # fallback: just take order
        series = pd.to_numeric(df[col], errors="coerce").dropna()
    if series.empty: return ("—","—")
    try:
        first, last, n = float(series.iloc[0]), float(series.iloc[-1]), max(1, len(series)-1)
        cagr = (last/first)**(1/n)-1 if first>0 and last>0 else np.nan
        return (f"{last:,.0f}", f"{cagr:.1%}" if np.isfinite(cagr) else "—")
    except Exception:
        return (f"{series.iloc[-1]:,.0f}", "—")

rev_val, rev_cagr = kpi_from(metrics["revenue"])
ebitda_val, _ = kpi_from(metrics["ebitda"])
fcf_val, _ = kpi_from(metrics["fcf"])

with c1: st.metric("Revenue (latest)", rev_val, delta=rev_cagr if rev_cagr!="—" else None)
with c2: st.metric("EBITDA (latest)", ebitda_val)
with c3: st.metric("Free Cash Flow (latest)", fcf_val)

# ---- SHEETS (one tab per sheet with auto charts) ----
st.markdown("## Sheets")
tabs = st.tabs(list(dfs.keys()))
for (sheet_name, df), tab in zip(dfs.items(), tabs):
    with tab:
        st.subheader(sheet_name)
        with st.expander("Preview", expanded=False):
            st.dataframe(df, use_container_width=True, height=360)

        ncols = num_cols(df)
        dcols = date_cols(df)
        cats  = small_cat_cols(df)

        # Time series
        st.markdown("#### Time-series")
        plotted=False
        if dcols and ncols:
            x = dcols[0]
            dff = df.copy()
            dff[x] = pd.to_datetime(dff[x], errors="coerce")
            dff = dff.dropna(subset=[x]).sort_values(x)
            ys = ncols[:5] if len(ncols)>5 else ncols
            if ys:
                st.line_chart(dff.set_index(x)[ys], use_container_width=True)
                plotted=True
        if not plotted:
            st.info("No reliable date column found for a line chart.")

        # Category bars
        st.markdown("#### Category bars / histograms")
        if cats and ncols:
            cat = cats[0]
            ys = ncols[:3]
            agg = df.groupby(cat)[ys].sum(numeric_only=True).sort_values(ys[0], ascending=False).head(20)
            st.bar_chart(agg, use_container_width=True)
        elif ncols:
            # histograms fallback
            for c in ncols[:3]:
                st.markdown(f"**Histogram — {c}**")
                st.bar_chart(pd.Series(pd.to_numeric(df[c], errors='coerce').dropna()), use_container_width=True)
        else:
            st.info("No numeric columns to chart.")

        # Correlations
        st.markdown("#### Correlation (numeric)")
        if len(ncols) >= 2:
            try:
                corr = df[ncols].corr(numeric_only=True)
                st.dataframe(corr.style.background_gradient(cmap="Blues"), use_container_width=True)
            except Exception:
                st.write("Could not compute correlations.")

        # Sheet summary
        st.markdown("#### Notes")
        prof = sheet_profile(df, sheet_name)
        bullets = [
            f"- Rows: {prof['rows']}, Columns: {prof['cols']}",
            f"- Columns: {', '.join(prof['columns'][:12])}" + ("…" if len(prof['columns'])>12 else ""),
            f"- Numeric columns detected: {len(prof['num_stats'])}",
        ]
        low = sheet_name.lower()
        if any(k in low for k in ["income","p&l","profit"]):
            bullets.append("- Likely **Income Statement** → track revenue growth & margin trends.")
        if "balance" in low:
            bullets.append("- Likely **Balance Sheet** → examine working capital & leverage.")
        if any(k in low for k in ["cash","cf","free cash"]):
            bullets.append("- Likely **Cash Flow** → CFO vs CAPEX → FCF.")
        if "assumption" in low or "driver" in low:
            bullets.append("- **Assumptions** → confirm WACC, growth, terminal value.")
        st.markdown("\n".join(bullets))
        st.markdown("\n**Preview**:\n\n" + prof["preview"])
