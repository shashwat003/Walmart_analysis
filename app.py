# app.py — Walmart Workbook Analyzer
# - Auto-charts EVERY sheet (no “download my model / assignment” buttons).
# - Upload one Excel model (.xlsx). The app builds visuals + quick analysis per tab.
# - Analyst Chat answers questions grounded in the workbook (and general Walmart context when needed).

import io
import os
import re
import json
from typing import Dict, List

import numpy as np
import pandas as pd
import streamlit as st

# ────────────────────────────────────────────────────────────────────────────────
# PAGE / THEME
# ────────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Walmart Workbook Analyzer", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

BG, PANEL, BORDER, TEXT, MUTED = "#0b1220", "#111827", "#22314a", "#eef2f7", "#9fb0c7"
PRIMARY, PRIMARY2, ACCENT = "#22d3ee", "#0fb5cf", "#7c3aed"
st.markdown(f"""
<style>
  html, body, .block-container {{ background:{BG}; color:{TEXT}; }}
  .card {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:18px; box-shadow:0 10px 30px rgba(3,12,24,.45); }}
  .headline {{ font-size:2rem; font-weight:800; letter-spacing:-.02em; }}
  .soft {{ color:{MUTED}; }}
  .pill {{ display:inline-flex; align-items:center; gap:8px; padding:6px 12px; border-radius:999px; font-size:.85rem;
          background:rgba(34,211,238,.12); color:{PRIMARY}; border:1px solid rgba(34,211,238,.28); }}
  .chat-wrap {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:0; }}
  div[data-testid="stChatMessage"] {{ margin-left:0 !important; margin-right:0 !important; }}
  div[data-testid="stChatInput"] textarea {{ background:#0c1423 !important; color:{TEXT} !important; border:2px solid {BORDER} !important; }}
  .stChatMessage .stMarkdown p {{ color:{TEXT}; }}
</style>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────────────────────────────────────────
# OPTIONAL LLM (Azure OpenAI)
# ────────────────────────────────────────────────────────────────────────────────
OPENAI_OK, client = True, None
AZURE_OPENAI_ENDPOINT    = os.getenv("AZURE_OPENAI_ENDPOINT", "https://testaisentiment.openai.azure.com/")
AZURE_OPENAI_API_KEY     = os.getenv("AZURE_OPENAI_API_KEY",     "cb1c33772b3c4edab77db69ae18c9a43")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")
AZURE_OPENAI_DEPLOYMENT  = os.getenv("AZURE_OPENAI_DEPLOYMENT",  "aipocexploration")
try:
    from openai import AzureOpenAI
    client = AzureOpenAI(azure_endpoint=AZURE_OPENAI_ENDPOINT, api_key=AZURE_OPENAI_API_KEY,
                         api_version=AZURE_OPENAI_API_VERSION)
except Exception:
    OPENAI_OK = False

def ask_gpt(messages, temperature=0.2, max_tokens=900):
    if not OPENAI_OK or not AZURE_OPENAI_DEPLOYMENT:
        return "(LLM not configured. The chat will only work after you set Azure OpenAI env vars.)"
    try:
        r=client.chat.completions.create(model=AZURE_OPENAI_DEPLOYMENT,
                                         messages=messages, temperature=temperature, max_tokens=max_tokens)
        return r.choices[0].message.content
    except Exception as e:
        return f"(Error calling Azure OpenAI: {e})"

# ────────────────────────────────────────────────────────────────────────────────
# HELPERS
# ────────────────────────────────────────────────────────────────────────────────
def need_openpyxl():
    try:
        import openpyxl  # noqa
        return False
    except Exception:
        return True

@st.cache_data(show_spinner=False)
def load_workbook(file_bytes_or_path) -> Dict[str, pd.DataFrame]:
    if need_openpyxl():
        raise ImportError("openpyxl is required to read .xlsx files. Install it and restart the app.")
    if isinstance(file_bytes_or_path, (str, os.PathLike)):
        xl = pd.ExcelFile(file_bytes_or_path, engine="openpyxl")
    else:
        xl = pd.ExcelFile(io.BytesIO(file_bytes_or_path), engine="openpyxl")
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

def numeric_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def date_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if is_date_col(df[c])]

def first_categorical(df: pd.DataFrame, max_unique=20) -> List[str]:
    cats=[]
    for c in df.columns:
        if c in numeric_cols(df): continue
        u = df[c].nunique(dropna=True)
        if 1 < u <= max_unique:
            cats.append(c)
    return cats

def sheet_profile(df: pd.DataFrame, name: str) -> dict:
    stats = {}
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_numeric_dtype(s):
            stats[str(c)] = {
                "count": int(s.count()),
                "mean": float(np.nanmean(s)) if s.count() else np.nan,
                "std":  float(np.nanstd(s))  if s.count() else np.nan,
                "min":  float(np.nanmin(s))  if s.count() else np.nan,
                "max":  float(np.nanmax(s))  if s.count() else np.nan,
            }
    with pd.option_context("display.max_columns", 10, "display.width", 1000):
        preview = df.head(8).to_markdown(index=False)
    return {"name":name, "rows":int(df.shape[0]), "cols":int(df.shape[1]),
            "columns":[str(x) for x in df.columns], "num_stats":stats, "preview":preview}

def build_corpus(profiles: List[dict]) -> List[str]:
    chunks=[]
    for p in profiles:
        chunks.append(
f"""Sheet: {p['name']}
Rows: {p['rows']}, Cols: {p['cols']}
Columns: {', '.join(p['columns'])}
Numeric summary: {json.dumps(p['num_stats'])[:1600]}
Preview:
{p['preview']}
""")
    return chunks or ["(no workbook context loaded)"]

def simple_retriever(query: str, chunks: List[str], top_k: int=4) -> List[str]:
    toks = re.findall(r"[a-z0-9\-%\.]+", (query or "").lower())
    scored=[]
    for ch in chunks:
        t=ch.lower(); score=sum(t.count(k) for k in toks if len(k)>2)
        scored.append((score, ch))
    scored.sort(key=lambda x:x[0], reverse=True)
    return [c for s,c in scored[:top_k] if s>0] or [scored[0][1]]

SYSTEM_PROMPT = """You are a Walmart (NYSE: WMT) valuation analyst assistant.
Use ONLY the provided workbook context for figures; name the sheet(s) you’re using.
If a question needs data that isn’t in the workbook context, say so briefly.
Be concise and precise. Note assumptions. Use the currency present in the sheet(s).
"""

def chat_over_sheets(user_msg: str, corpus_chunks: List[str]) -> str:
    try:
        context_parts = simple_retriever(user_msg, corpus_chunks or ["(no workbook)"], top_k=4)
        context = "\n\n---\n\n".join([str(c) for c in context_parts])
    except Exception:
        context = "(no workbook)"
    messages = [
        {"role":"system", "content":SYSTEM_PROMPT},
        {"role":"user", "content": f"Workbook context:\n{context}\n\nQuestion: {user_msg}"}
    ]
    out = ask_gpt(messages)
    if out.startswith("("):
        return "LLM not configured. You can still explore the charts; to enable chat, set Azure OpenAI env vars."
    return out

# ────────────────────────────────────────────────────────────────────────────────
# STATE
# ────────────────────────────────────────────────────────────────────────────────
if "dfs" not in st.session_state: st.session_state.dfs = {}
if "profiles" not in st.session_state: st.session_state.profiles = []
if "corpus" not in st.session_state: st.session_state.corpus = []
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role":"assistant", "content":"Hi! Upload your Walmart Excel model (.xlsx). I’ll visualize every sheet and answer questions about the tabs."}
    ]

# ────────────────────────────────────────────────────────────────────────────────
# HEADER
# ────────────────────────────────────────────────────────────────────────────────
hl, hr = st.columns([0.75, 0.25])
with hl:
    st.markdown(f"""
      <div style="display:flex;gap:14px;align-items:center;">
        <div style="width:46px;height:46px;background:{ACCENT};border-radius:12px;"></div>
        <div class="headline">Walmart Workbook Analyzer</div>
        <span class="pill">● Auto-visualization</span>
      </div>""", unsafe_allow_html=True)
with hr:
    st.markdown('<div style="text-align:right;"><span class="pill">DCF • Statements • Sensitivity</span></div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="card" style="margin-top:8px;margin-bottom:12px;">
  <div class="soft">Upload your Excel file; I’ll create charts for each tab (time-series, bar/histograms, and correlations) and summarize what’s in each sheet.</div>
</div>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ────────────────────────────────────────────────────────────────────────────────
left, right = st.columns([0.58, 0.42])

# LEFT: Workbook + Auto Charts for ALL sheets
with left:
    st.markdown('<div class="card"><div style="font-weight:800;font-size:1.15rem;">📁 Upload Workbook</div><div class="soft" style="margin-top:4px;">Load your Excel model. I’ll analyze every sheet automatically.</div></div>', unsafe_allow_html=True)

    up = st.file_uploader("Upload Excel model (.xlsx)", type=["xlsx"])
    if up:
        try:
            st.session_state.dfs = load_workbook(up.getvalue())
            st.session_state.profiles = [sheet_profile(df, name) for name, df in st.session_state.dfs.items()]
            st.session_state.corpus = build_corpus(st.session_state.profiles)
            st.success("Workbook loaded. Scroll down for auto-generated charts.")
        except ImportError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Failed to load Excel: {e}")

    if not st.session_state.dfs:
        st.warning("No workbook loaded yet.")
    else:
        # make a tab for EACH sheet with auto charts
        tab_objs = st.tabs(list(st.session_state.dfs.keys()))
        for (sheet_name, df), tab in zip(st.session_state.dfs.items(), tab_objs):
            with tab:
                st.subheader(f"{sheet_name}")

                # Preview
                with st.expander("Preview", expanded=False):
                    st.dataframe(df, use_container_width=True, height=360)

                # Quick metrics
                n_num = len(numeric_cols(df))
                n_cat = len([c for c in df.columns if c not in numeric_cols(df)])
                c1,c2,c3 = st.columns(3)
                with c1: st.metric("Rows", df.shape[0])
                with c2: st.metric("Columns", df.shape[1])
                with c3: st.metric("Numeric cols", n_num)

                # Auto charts
                st.markdown("#### 📈 Time-series / Line charts")
                dcols = date_cols(df)
                ncols = numeric_cols(df)
                plotted=False
                if dcols and ncols:
                    # Try to coerce first date column
                    x = dcols[0]
                    try:
                        dff = df.copy()
                        dff[x] = pd.to_datetime(dff[x], errors="coerce")
                        dff = dff.dropna(subset=[x])
                        show_cols = ncols[:5] if len(ncols) > 5 else ncols
                        if show_cols:
                            st.line_chart(dff.set_index(x)[show_cols], use_container_width=True)
                            plotted=True
                    except Exception:
                        pass
                if not plotted:
                    st.info("No reliable date column found for a line chart.")

                st.markdown("#### 📊 Category bars / histograms")
                cats = first_categorical(df, max_unique=20)
                if cats and ncols:
                    cat = cats[0]
                    # choose up to 3 numeric columns
                    show = ncols[:3]
                    if not show:
                        st.info("No numeric columns detected for bar charts.")
                    else:
                        # Build a grouped bar by the first categorical
                        agg = df.groupby(cat)[show].sum(numeric_only=True).sort_values(show[0], ascending=False).head(20)
                        st.bar_chart(agg, use_container_width=True)
                elif ncols:
                    # Fallback: histogram for the first few numerics
                    from math import ceil
                    to_show = ncols[:3]
                    for c in to_show:
                        st.markdown(f"**Histogram — {c}**")
                        try:
                            st.bar_chart(pd.Series(df[c].dropna().values), use_container_width=True)
                        except Exception:
                            st.write("Could not render histogram for this column.")

                st.markdown("#### 🔗 Correlations (numeric)")
                if len(ncols) >= 2:
                    try:
                        corr = df[ncols].corr(numeric_only=True)
                        st.dataframe(corr.style.background_gradient(cmap="Blues"), use_container_width=True)
                    except Exception:
                        st.write("Could not compute correlations for this sheet.")
                else:
                    st.write("Not enough numeric columns to compute correlations.")

                # Quick analysis text (what’s in the tab)
                with st.expander("📝 Sheet summary / interpretation"):
                    prof = sheet_profile(df, sheet_name)
                    # human-ish writeup based on profile
                    summary_parts = [
                        f"- Rows: {prof['rows']}, Columns: {prof['cols']}",
                        f"- Columns detected: {', '.join(prof['columns'][:12])}" + ("…" if len(prof['columns'])>12 else ""),
                        f"- Numeric columns: {len(prof['num_stats'])}"
                    ]
                    # If common financial sheet names, add hints
                    low = sheet_name.lower()
                    if any(k in low for k in ["income", "p&l", "profit"]):
                        summary_parts.append("- Looks like an **Income Statement** style sheet: consider margin trends, revenue CAGR, and operating leverage.")
                    if any(k in low for k in ["balance", "bs"]):
                        summary_parts.append("- Looks like a **Balance Sheet**: check working capital and leverage ratios.")
                    if any(k in low for k in ["cash", "cf"]):
                        summary_parts.append("- Looks like a **Cash Flow**: examine CFO vs CAPEX to derive FCF.")
                    if any(k in low for k in ["assumption", "input", "drivers"]):
                        summary_parts.append("- **Assumptions/Drivers**: confirm WACC, growth, and terminal value inputs drive valuation tabs.")
                    st.markdown("\n".join(summary_parts))
                    st.markdown("\n**Preview**:\n\n" + prof["preview"])

# RIGHT: Chat
with right:
    st.markdown(f"""
    <div class="card" style="margin-bottom:10px;">
      <div style="display:flex;align-items:center;justify-content:space-between;">
        <div style="font-weight:800;font-size:1.15rem;">💬 Analyst Chat</div>
        <span class="pill">Grounded in your sheets</span>
      </div>
      <div class="soft">Ask about assumptions, statements, WACC, growth, margins, FCF, or sensitivities. I’ll reference sheet names where relevant.</div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="chat-wrap">', unsafe_allow_html=True)

    for m in st.session_state.messages:
        with st.chat_message("assistant" if m["role"]=="assistant" else "user"):
            st.write(m["content"])

    user_text = st.chat_input("Ask a question about your Walmart workbook…")
    if user_text:
        st.session_state.messages.append({"role":"user","content":user_text})
        reply = chat_over_sheets(user_text, st.session_state.corpus)
        st.session_state.messages.append({"role":"assistant","content":reply})
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<hr style='border-color:#1f2937;'>", unsafe_allow_html=True)
st.caption("© Walmart Workbook Analyzer — auto-visualizes each Excel tab and answers questions grounded in your model.")
