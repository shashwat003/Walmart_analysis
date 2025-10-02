# app.py — Walmart Sheets Visualizer (no uploaders, no chat, no downloads)
import os, io
from typing import Dict, List
import numpy as np
import pandas as pd
import streamlit as st

# ---------- CONFIG ----------
FILE_NAME = "FIN42030 WMT Valuation (2).xlsx"   # place this next to app.py

# ---------- PAGE / THEME ----------
st.set_page_config(page_title="Walmart Sheets Visualizer", page_icon="📊", layout="wide")
BG, PANEL, BORDER, TEXT, MUTED = "#0b1220", "#111827", "#22314a", "#eef2f7", "#9fb0c7"
st.markdown(f"""
<style>
  html, body, .block-container {{ background:{BG}; color:{TEXT}; }}
  .card {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:18px; }}
  .headline {{ font-size:1.8rem; font-weight:800; }}
  .soft {{ color:{MUTED}; }}
</style>
""", unsafe_allow_html=True)

# ---------- HELPERS ----------
def need_openpyxl() -> bool:
    try:
        import openpyxl  # noqa
        return False
    except Exception:
        return True

@st.cache_data(show_spinner=False)
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    if need_openpyxl():
        raise ImportError("openpyxl is not installed; install it with `pip install openpyxl`.")
    xl = pd.ExcelFile(path, engine="openpyxl")
    return {name: xl.parse(name) for name in xl.sheet_names}

def is_date_col(s: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(s): return True
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
        if c in num_cols(df): continue
        try:
            u = df[c].nunique(dropna=True)
            if 1 < u <= max_unique: cats.append(c)
        except Exception:
            pass
    return cats

def sheet_summary(df: pd.DataFrame) -> str:
    nums = num_cols(df)
    parts = [
        f"- Rows: **{df.shape[0]}**, Columns: **{df.shape[1]}**",
        f"- Numeric columns: **{len(nums)}**",
        f"- Sample columns: {', '.join(map(str, df.columns[:12]))}" + ("…" if df.shape[1] > 12 else "")
    ]
    return "\n".join(parts)

# ---------- UI HEADER ----------
st.markdown(f"""
<div class="headline">Walmart Sheets Visualizer</div>
<div class="soft">Auto-creates tabs and visuals for every sheet in <b>{FILE_NAME}</b>.</div>
""", unsafe_allow_html=True)
st.markdown("")

# ---------- LOAD ----------
if not os.path.exists(FILE_NAME):
    st.error(f"File not found: {FILE_NAME}. Place the Excel file next to app.py and re-run.")
    st.stop()

try:
    dfs = load_workbook(FILE_NAME)
except ImportError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"Failed to read Excel: {e}")
    st.stop()

# ---------- OVERVIEW ----------
st.success(f"Loaded {FILE_NAME} • {len(dfs)} sheet(s)")
st.markdown("### Overview")
colA, colB, colC = st.columns(3)
with colA: st.metric("Sheets", len(dfs))
with colB: st.metric("Total Rows", sum(df.shape[0] for df in dfs.values()))
with colC: st.metric("Total Columns", sum(df.shape[1] for df in dfs.values()))

# ---------- TABS PER SHEET ----------
st.markdown("### Sheets")
tabs = st.tabs(list(dfs.keys()))

for (sheet_name, df), tab in zip(dfs.items(), tabs):
    with tab:
        st.subheader(sheet_name)

        with st.expander("Preview", expanded=False):
            st.dataframe(df, use_container_width=True, height=360)

        st.markdown("#### Summary")
        st.markdown(sheet_summary(df))

        ncols = num_cols(df)
        dcols = date_cols(df)
        cats  = small_cat_cols(df)

        # Time-series (if date/period exists)
        st.markdown("#### Time-series")
        plotted = False
        if dcols and ncols:
            x = dcols[0]
            dff = df.copy()
            dff[x] = pd.to_datetime(dff[x], errors="coerce")
            dff = dff.dropna(subset=[x]).sort_values(x)
            ys = ncols[:5] if len(ncols) > 5 else ncols
            if not dff.empty and ys:
                st.line_chart(dff.set_index(x)[ys], use_container_width=True)
                plotted = True
        if not plotted:
            st.info("No reliable date/period column found for a line chart.")

        # Category bars / group-by (if small-cardinality text exists)
        st.markdown("#### Category bars / groupings")
        if cats and ncols:
            cat = cats[0]
            ys  = ncols[:3]
            try:
                agg = df.groupby(cat)[ys].sum(numeric_only=True).sort_values(ys[0], ascending=False).head(20)
                st.bar_chart(agg, use_container_width=True)
            except Exception:
                st.write("Could not render grouped bars for this sheet.")
        elif ncols:
            st.caption("No small-cardinality text column; showing histograms for first numeric columns.")
            for c in ncols[:3]:
                st.markdown(f"**Histogram — {c}**")
                vals = pd.to_numeric(df[c], errors="coerce").dropna()
                if not vals.empty:
                    st.bar_chart(pd.Series(vals.values), use_container_width=True)
                else:
                    st.write("No numeric data to plot.")

        # Correlation matrix
        st.markdown("#### Correlation (numeric)")
        if len(ncols) >= 2:
            try:
                corr = df[ncols].corr(numeric_only=True)
                st.dataframe(corr.style.background_gradient(cmap="Blues"), use_container_width=True)
            except Exception:
                st.write("Could not compute correlations for this sheet.")
        else:
            st.write("Not enough numeric columns for correlations.")
