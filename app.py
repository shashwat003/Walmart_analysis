# app.py — Walmart Sheets Visualizer (auto-tabs & visuals for each Excel sheet)
# - No uploaders / downloads / chat
# - Loads a specific Excel file from the repo root
# - Robust plotting (date columns, "year header" tables, categories, histograms, correlation)

import os
import re
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st

# ────────────────────────────────────────────────────────────────────────────────
# CONFIG
# ────────────────────────────────────────────────────────────────────────────────
FILE_NAME = "FIN42030 WMT Valuation (2).xlsx"   # keep this file in the repo root (same folder as app.py)

st.set_page_config(page_title="Walmart Sheets Visualizer", page_icon="📊", layout="wide")

# Theme
BG, PANEL, BORDER, TEXT, MUTED = "#0b1220", "#111827", "#22314a", "#eef2f7", "#9fb0c7"
st.markdown(f"""
<style>
  html, body, .block-container {{ background:{BG}; color:{TEXT}; }}
  .card {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:18px; }}
  .headline {{ font-size:1.9rem; font-weight:800; }}
  .soft {{ color:{MUTED}; }}
</style>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────────────────────────────────────────
# HELPERS
# ────────────────────────────────────────────────────────────────────────────────
def need_openpyxl() -> bool:
    try:
        import openpyxl  # noqa
        return False
    except Exception:
        return True

@st.cache_data(show_spinner=False)
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    """Read all sheets from an .xlsx into a dict of DataFrames."""
    if need_openpyxl():
        raise ImportError("openpyxl is not installed in this environment. Add it to requirements.txt and redeploy.")
    xl = pd.ExcelFile(path, engine="openpyxl")
    return {name: xl.parse(name) for name in xl.sheet_names}

def is_date_col(s: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(s):
        return True
    if s.dtype == object:
        try:
            parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
            return parsed.notna().mean() >= 0.6
        except Exception:
            return False
    return False

def num_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def date_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if is_date_col(df[c])]

def small_cat_cols(df: pd.DataFrame, max_unique=20) -> List[str]:
    cats=[]
    for c in df.columns:
        if c in num_cols(df):
            continue
        try:
            u = df[c].nunique(dropna=True)
            if 1 < u <= max_unique:
                cats.append(c)
        except Exception:
            pass
    return cats

def sheet_profile(df: pd.DataFrame) -> str:
    nums = num_cols(df)
    return (
        f"- Rows: **{df.shape[0]}**, Columns: **{df.shape[1]}**\n"
        f"- Numeric columns (pre-coerce): **{len(nums)}**\n"
        f"- Sample columns: {', '.join(map(str, df.columns[:12]))}" + ("…" if df.shape[1] > 12 else "")
    )

def find_year_header_cols(df: pd.DataFrame) -> List[str]:
    """Find columns that look like years (2018 / FY2019 / 2020A / 2021E)."""
    yearish = []
    pat = re.compile(r"(19|20)\d{2}")
    for c in df.columns:
        s = str(c)
        if pat.search(s):
            yearish.append(c)
    # keep only columns that have any numeric values after coercion
    if yearish:
        tmp = df.copy()
        for c in yearish:
            tmp[c] = pd.to_numeric(tmp[c], errors="coerce")
        yearish = [c for c in yearish if tmp[c].notna().any()]
    return yearish

def wide_years_to_long(df: pd.DataFrame, year_cols: List[str], label_col: str | None = None) -> pd.DataFrame | None:
    """Turn a wide financial table (metrics x years) into long format (Year, Metric, Value)."""
    if not year_cols:
        return None
    if label_col is None:
        non_year = [c for c in df.columns if c not in year_cols]
        label_col = non_year[0] if non_year else None
    work = df.copy()
    for yc in year_cols:
        work[yc] = pd.to_numeric(work[yc], errors="coerce")
    long = work.melt(id_vars=[label_col] if label_col else None, value_vars=year_cols,
                     var_name="Year", value_name="Value")
    long["Year"] = long["Year"].astype(str).str.extract(r"((?:19|20)\d{2})").astype(float)
    long = long.dropna(subset=["Year", "Value"])
    if long.empty:
        return None
    long["Year"] = long["Year"].astype(int)
    # pick top metrics by variance for readability
    if label_col and label_col in long.columns:
        vard = long.groupby(label_col)["Value"].var().sort_values(ascending=False)
        top_labels = list(vard.head(5).index)
        long = long[long[label_col].isin(top_labels)]
    return long

def coerce_numeric(df: pd.DataFrame, cols: List[str]) -> tuple[pd.DataFrame, List[str]]:
    """Convert columns to numeric where possible; return df and list of non-empty numeric cols."""
    out = df.copy()
    keep=[]
    for c in cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")
        if out[c].notna().any():
            keep.append(c)
    return out, keep

def guess_overview_metrics(dfs: Dict[str, pd.DataFrame]) -> dict:
    """Best-effort heuristic to find Revenue / EBITDA / FCF columns across sheets."""
    m = {"revenue":None, "ebitda":None, "fcf":None}
    for sname, df in dfs.items():
        low = sname.lower()
        if any(k in low for k in ["summary","income","p&l","statement","fcf","cash","forecast","integrated"]):
            cols_lower = [str(c).lower() for c in df.columns]
            if m["revenue"] is None:
                cand = [c for c in cols_lower if "revenue" in c or "sales" in c]
                if cand: m["revenue"] = (sname, df.columns[cols_lower.index(cand[0])])
            if m["ebitda"] is None:
                cand = [c for c in cols_lower if "ebitda" in c]
                if cand: m["ebitda"] = (sname, df.columns[cols_lower.index(cand[0])])
            if m["fcf"] is None:
                cand = [c for c in cols_lower if "free cash" in c or c.startswith("fcf")]
                if cand: m["fcf"] = (sname, df.columns[cols_lower.index(cand[0])])
    return m

def latest_and_cagr(df: pd.DataFrame, value_col: str) -> tuple[str, str]:
    """Return latest value and rough CAGR when possible."""
    # prefer a real date/period column
    dcols = date_cols(df)
    series = None
    if dcols:
        t = dcols[0]
        x = df.copy()
        x[t] = pd.to_datetime(x[t], errors="coerce")
        x = x.dropna(subset=[t]).sort_values(t)
        x[value_col] = pd.to_numeric(x[value_col], errors="coerce")
        series = x[value_col].dropna()
    else:
        # try year headers layout
        year_cols = find_year_header_cols(df)
        if year_cols:
            long = wide_years_to_long(df, year_cols)
            if long is not None and not long.empty:
                series = long.groupby("Year")["Value"].sum().sort_index()
        else:
            # fallback: coerce the column
            x = pd.to_numeric(df[value_col], errors="coerce").dropna()
            if not x.empty:
                series = x
    if series is None or len(series) == 0:
        return ("—", "—")
    try:
        first, last, n = float(series.iloc[0]), float(series.iloc[-1]), max(1, len(series)-1)
        cagr = (last/first)**(1/n) - 1 if first>0 and last>0 else np.nan
        return (f"{last:,.0f}", f"{cagr:.1%}" if np.isfinite(cagr) else "—")
    except Exception:
        return (f"{series.iloc[-1]:,.0f}", "—")

# ────────────────────────────────────────────────────────────────────────────────
# HEADER
# ────────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="headline">Walmart Sheets Visualizer</div>
<div class="soft">Auto-creates tabs and visuals for every sheet in <b>{FILE_NAME}</b>.</div>
""", unsafe_allow_html=True)
st.markdown("")

# ────────────────────────────────────────────────────────────────────────────────
# LOAD EXCEL
# ────────────────────────────────────────────────────────────────────────────────
if not os.path.exists(FILE_NAME):
    st.error(f"File not found: {FILE_NAME}. Place the Excel next to app.py in the repo.")
    st.stop()

try:
    dfs = load_workbook(FILE_NAME)
except ImportError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"Failed to read Excel: {e}")
    st.stop()

# ────────────────────────────────────────────────────────────────────────────────
# OVERVIEW (best-effort KPIs)
# ────────────────────────────────────────────────────────────────────────────────
st.success(f"Loaded {FILE_NAME} • {len(dfs)} sheet(s)")
st.markdown("## Overview")

c1, c2, c3, c4 = st.columns(4)
with c1: st.metric("Sheets", len(dfs))
with c2: st.metric("Total Rows", sum(df.shape[0] for df in dfs.values()))
with c3: st.metric("Total Columns", sum(df.shape[1] for df in dfs.values()))

metrics = guess_overview_metrics(dfs)
rev_val = ebitda_val = fcf_val = "—"
rev_delta = ebitda_delta = fcf_delta = None

if metrics["revenue"]:
    v, d = latest_and_cagr(dfs[metrics["revenue"][0]], metrics["revenue"][1])
    rev_val, rev_delta = v, (d if d != "—" else None)
if metrics["ebitda"]:
    v, _ = latest_and_cagr(dfs[metrics["ebitda"][0]], metrics["ebitda"][1]); ebitda_val = v
if metrics["fcf"]:
    v, _ = latest_and_cagr(dfs[metrics["fcf"][0]], metrics["fcf"][1]); fcf_val = v

with c4:
    st.metric("Revenue (latest)", rev_val, delta=rev_delta)

c5, c6 = st.columns(2)
with c5: st.metric("EBITDA (latest)", ebitda_val)
with c6: st.metric("Free Cash Flow (latest)", fcf_val)

# ────────────────────────────────────────────────────────────────────────────────
# SHEETS — one tab per sheet with robust visuals
# ────────────────────────────────────────────────────────────────────────────────
st.markdown("## Sheets")
tabs = st.tabs(list(dfs.keys()))

for (sheet_name, raw_df), tab in zip(dfs.items(), tabs):
    with tab:
        st.subheader(sheet_name)

        with st.expander("Preview", expanded=False):
            st.dataframe(raw_df, use_container_width=True, height=380)

        st.markdown("#### Summary")
        st.markdown(sheet_profile(raw_df))

        # Identify columns
        dcols = date_cols(raw_df)
        ncols = num_cols(raw_df)
        year_cols = find_year_header_cols(raw_df)

        # ── Time-series A: real date/period column
        st.markdown("#### Time-series")
        plotted_ts = False
        if dcols and ncols:
            x = dcols[0]
            df_ts = raw_df.copy()
            df_ts[x] = pd.to_datetime(df_ts[x], errors="coerce")
            df_ts = df_ts.dropna(subset=[x]).sort_values(x)
            df_ts, safe_ncols = coerce_numeric(df_ts, ncols)
            safe_ncols = safe_ncols[:5]
            if safe_ncols and not df_ts.empty:
                try:
                    st.line_chart(df_ts.set_index(x)[safe_ncols], use_container_width=True)
                    plotted_ts = True
                except Exception:
                    pass

        # ── Time-series B: wide year headers → long format
        if not plotted_ts and year_cols:
            long = wide_years_to_long(raw_df, year_cols)
            if long is not None and not long.empty:
                try:
                    label_cols = [c for c in long.columns if c not in ("Year", "Value")]
                    if label_cols:
                        label_col = label_cols[0]
                        piv = long.pivot_table(index="Year", columns=label_col, values="Value", aggfunc="mean").sort_index()
                        st.line_chart(piv, use_container_width=True)
                        plotted_ts = True
                except Exception:
                    pass

        if not plotted_ts:
            st.info("No reliable time axis found for a line chart.")

        # ── Category bars / groupings
        st.markdown("#### Category bars / groupings")
        cats = small_cat_cols(raw_df, max_unique=20)
        if cats:
            cat = cats[0]
            co_df = raw_df.copy()
            numeric_after=[]
            for c in raw_df.columns:
                if c == cat:
                    continue
                # attempt coercion to numeric; keep if any numeric values exist
                co_df[c] = pd.to_numeric(co_df[c], errors="coerce")
                if co_df[c].notna().any():
                    numeric_after.append(c)
            show = numeric_after[:3]
            if show:
                try:
                    agg = co_df.groupby(cat)[show].sum(numeric_only=True).sort_values(show[0], ascending=False).head(20)
                    st.bar_chart(agg, use_container_width=True)
                except Exception:
                    st.write("Could not render grouped bars for this sheet.")
            else:
                st.caption("No numeric data to aggregate by category.")
        else:
            st.caption("No small-cardinality category column detected.")

        # ── Numeric histograms (fallback)
        st.markdown("#### Numeric histograms (fallback)")
        co_df = raw_df.copy()
        numeric_candidates=[]
        for c in raw_df.columns:
            co_df[c] = pd.to_numeric(co_df[c], errors="coerce")
            if co_df[c].notna().any():
                numeric_candidates.append(c)
        numeric_candidates = [c for c in numeric_candidates if c in raw_df.columns][:3]
        if numeric_candidates:
            for c in numeric_candidates:
                st.markdown(f"**{c}**")
                try:
                    st.bar_chart(co_df[c].dropna().reset_index(drop=True), use_container_width=True)
                except Exception:
                    st.write("Could not render histogram for this column.")
        else:
            st.caption("No numeric columns available for histograms.")

        # ── Correlation matrix
        st.markdown("#### Correlation (numeric)")
        num_for_corr = [c for c in co_df.columns if pd.api.types.is_numeric_dtype(co_df[c]) and co_df[c].notna().any()]
        if len(num_for_corr) >= 2:
            try:
                corr = co_df[num_for_corr].corr(numeric_only=True)
                st.dataframe(corr.style.background_gradient(cmap="Blues"), use_container_width=True)
            except Exception:
                st.write("Could not compute correlations.")
        else:
            st.caption("Not enough numeric columns for correlations.")

# Footer
st.markdown("<hr style='border-color:#22314a;'>", unsafe_allow_html=True)
st.caption("© Walmart Sheets Visualizer — auto-tabs & visuals generated from your Excel model.")
