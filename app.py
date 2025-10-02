# app.py — Walmart Valuation Explorer (Streamlit)
# What changed from your banking app:
# - Rebranded UI for valuation.
# - Left pane: workbook loader, sheet explorer, chart builder.
# - Right pane: Chatbot grounded in the Excel sheets (and optional Walmart context).
# - No banking verification flows; chat logic rewritten to reason over your sheets.

import io
import os
import re
import json
import math
import textwrap
from typing import Dict, List, Tuple

import pandas as pd
import numpy as np
import streamlit as st

# ────────────────────────────────────────────────────────────────────────────────
# PAGE / CONSTANTS
# ────────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Walmart Valuation Explorer", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

# Default file paths (your uploads). You can still upload a new one in the UI.
DEFAULT_XLSX_PATH = "/mnt/data/FIN42030 WMT Valuation (2).xlsx"
DEFAULT_ASSIGNMENT_PDF = "/mnt/data/Assignment1 (1).pdf"

# Optional Azure OpenAI (same fields as before — safe fallback if unset)
AZURE_OPENAI_ENDPOINT    = os.getenv("AZURE_OPENAI_ENDPOINT", "https://testaisentiment.openai.azure.com/")
AZURE_OPENAI_API_KEY     = os.getenv("AZURE_OPENAI_API_KEY",     "cb1c33772b3c4edab77db69ae18c9a43")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-15-preview")
AZURE_OPENAI_DEPLOYMENT  = os.getenv("AZURE_OPENAI_DEPLOYMENT",  "aipocexploration")

# Theme / palette (kept your style)
BG, PANEL, BORDER, TEXT, MUTED = "#0b1220", "#111827", "#22314a", "#eef2f7", "#9fb0c7"
PRIMARY, PRIMARY2, ACCENT, GOOD, WARN, DANGER = "#22d3ee", "#0fb5cf", "#7c3aed", "#22c55e", "#f59e0b", "#ef4444"
st.markdown(f"""
<style>
  html, body, .block-container {{ background:{BG}; color:{TEXT}; }}
  .card {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:18px; box-shadow:0 10px 30px rgba(3,12,24,.45); }}
  .headline {{ font-size:2rem; font-weight:800; letter-spacing:-.02em; }}
  .soft {{ color:{MUTED}; }}
  .pill {{ display:inline-flex; align-items:center; gap:8px; padding:6px 12px; border-radius:999px; font-size:.85rem;
          background:rgba(34,211,238,.12); color:{PRIMARY}; border:1px solid rgba(34,211,238,.28); }}
  .stButton>button {{ background:#1f2937; color:{TEXT}; border:1px solid {BORDER}; border-radius:10px; font-weight:700; padding:.6rem 1rem; }}
  .stButton>button:hover {{ background:linear-gradient(180deg,{PRIMARY} 0%, {PRIMARY2} 100%); color:#001016; border:0;
                            box-shadow:0 6px 18px rgba(34,211,238,.18); }}
  .chat-wrap {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:0; }}
  div[data-testid="stChatMessage"] {{ margin-left:0 !important; margin-right:0 !important; }}
  div[data-testid="stChatInput"] textarea {{ background:#0c1423 !important; color:{TEXT} !important; border:2px solid {BORDER} !important; }}
  .stChatMessage .stMarkdown p {{ color:{TEXT}; }}
</style>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────────────────────────────────────────
# OPTIONAL LLM
# ────────────────────────────────────────────────────────────────────────────────
OPENAI_OK=True
client=None
try:
    from openai import AzureOpenAI
    client=AzureOpenAI(azure_endpoint=AZURE_OPENAI_ENDPOINT, api_key=AZURE_OPENAI_API_KEY,
                       api_version=AZURE_OPENAI_API_VERSION)
except Exception:
    OPENAI_OK=False

def ask_gpt(messages, temperature=0.2, max_tokens=800):
    """Generic chat completion wrapper."""
    if not OPENAI_OK or not AZURE_OPENAI_DEPLOYMENT:
        return "(Model not configured.)"
    try:
        r=client.chat.completions.create(model=AZURE_OPENAI_DEPLOYMENT,
                                         messages=messages, temperature=temperature, max_tokens=max_tokens)
        return r.choices[0].message.content
    except Exception as e:
        return f"(Error calling Azure OpenAI: {e})"

# ────────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ────────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes_or_path) -> Dict[str, pd.DataFrame]:
    """Load all sheets from Excel into dataframes (lowercase sheet names kept exact for display)."""
    if isinstance(file_bytes_or_path, (str, os.PathLike)):
        xl = pd.ExcelFile(file_bytes_or_path, engine="openpyxl")
    else:
        xl = pd.ExcelFile(io.BytesIO(file_bytes_or_path), engine="openpyxl")
    dfs = {name: xl.parse(name) for name in xl.sheet_names}
    return dfs

def is_date_col(s: pd.Series) -> bool:
    # robust-ish date detection
    if np.issubdtype(s.dtype, np.datetime64):
        return True
    if s.dtype == object:
        try:
            sample = s.dropna().head(10)
            parsed = pd.to_datetime(sample, errors="coerce", infer_datetime_format=True)
            return parsed.notna().mean() >= 0.7
        except Exception:
            return False
    return False

def sheet_profile(df: pd.DataFrame, name: str) -> Dict:
    """Quick metadata + stats per sheet to feed the chatbot."""
    prof = {"name": name, "rows": int(df.shape[0]), "cols": int(df.shape[1]), "columns": list(map(str, df.columns))}
    stats = {}
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_numeric_dtype(s):
            stats[c] = {
                "count": int(s.count()),
                "mean": float(np.nanmean(s)) if s.count() else np.nan,
                "std": float(np.nanstd(s)) if s.count() else np.nan,
                "min": float(np.nanmin(s)) if s.count() else np.nan,
                "max": float(np.nanmax(s)) if s.count() else np.nan,
            }
    prof["num_stats"] = stats
    # A tiny preview snippet to help LLM ground simple lookups
    with pd.option_context("display.max_columns", 12, "display.width", 1000):
        preview = df.head(8).to_markdown(index=False)
    prof["preview"] = preview
    return prof

def build_corpus(profiles: List[Dict]) -> List[str]:
    """Turn sheet profiles into small text chunks for keyword retrieval."""
    chunks=[]
    for p in profiles:
        chunk = f"""Sheet: {p['name']}
Rows: {p['rows']}, Cols: {p['cols']}
Columns: {', '.join(p['columns'])}
Numeric summary: {json.dumps(p['num_stats'])[:1800]}
Preview:
{p['preview']}
"""
        chunks.append(chunk)
    return chunks

def simple_retriever(query: str, chunks: List[str], top_k: int=3) -> List[str]:
    """Very light keyword overlap retriever (no external deps)."""
    q = re.findall(r"[a-zA-Z0-9\-%\.]+", query.lower())
    scored=[]
    for ch in chunks:
        text = ch.lower()
        score = sum(text.count(tok) for tok in q if len(tok) > 2)
        scored.append((score, ch))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [c for s,c in scored[:top_k] if s>0] or scored[:1]

SYSTEM_PROMPT = """You are a valuation analyst assistant for Walmart (NYSE: WMT).
Goal: answer questions using ONLY the provided workbook context unless the question is clearly general, non-quantitative Walmart knowledge.
When using workbook numbers, cite the sheet name(s) explicitly in your prose.
If the data needed isn't present in context, say you don't have enough data from the sheets.
Be concise, precise, and note assumptions. Use EUR/€ only if the sheet uses EUR; otherwise keep native currency from the sheet.
"""

def chat_over_sheets(user_msg: str, corpus_chunks: List[str]) -> str:
    context_parts = simple_retriever(user_msg, corpus_chunks, top_k=4)
    context = "\n\n---\n\n".join(context_parts)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": f"Workbook context (from Excel sheets):\n{context}\n\nQuestion: {user_msg}"}
    ]
    out = ask_gpt(messages)
    # Fallback if model not configured
    if out.startswith("("):
        return "Model isn't configured here. Try enabling Azure OpenAI, or ask a sheet-specific question I can compute (e.g., 'Average revenue by year in the Income Statement sheet')."
    return out

def numeric_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def date_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if is_date_col(df[c])]

# ────────────────────────────────────────────────────────────────────────────────
# STATE
# ────────────────────────────────────────────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = [{"role":"assistant","content":"Hi! Upload your model or use the default file, then pick a sheet to explore and ask me questions about the data or Walmart."}]
if "dfs" not in st.session_state:
    # Attempt default load; will be replaced on upload
    try:
        st.session_state.dfs = load_workbook(DEFAULT_XLSX_PATH)
    except Exception:
        st.session_state.dfs = {}
if "profiles" not in st.session_state:
    st.session_state.profiles = [sheet_profile(df, name) for name, df in st.session_state.dfs.items()]
if "corpus" not in st.session_state:
    st.session_state.corpus = build_corpus(st.session_state.profiles)

# ────────────────────────────────────────────────────────────────────────────────
# HEADER
# ────────────────────────────────────────────────────────────────────────────────
hl, hr = st.columns([0.75,0.25])
with hl:
    st.markdown(f"""
      <div style="display:flex;gap:14px;align-items:center;">
        <div style="width:46px;height:46px;background:{ACCENT};border-radius:12px;"></div>
        <div class="headline">Walmart Valuation Explorer</div>
        <span class="pill">● Interactive</span>
      </div>""", unsafe_allow_html=True)
with hr:
    st.markdown('<div style="text-align:right;"><span class="pill">DCF • Multiples • Sensitivity</span></div>', unsafe_allow_html=True)

st.markdown(f"""
<div class="card" style="margin-top:8px;margin-bottom:12px;">
  <div class="soft">Tip: Ask things like “Show revenue CAGR by segment from the <b>Income Statement</b> sheet” or “What WACC did we use in the <b>Assumptions</b> sheet?”</div>
</div>
""", unsafe_allow_html=True)

# ────────────────────────────────────────────────────────────────────────────────
# LAYOUT
# ────────────────────────────────────────────────────────────────────────────────
left, right = st.columns([0.54, 0.46])

# LEFT: Workbook & Charts
with left:
    st.markdown('<div class="card"><div style="font-weight:800;font-size:1.15rem;">📁 Workbook</div><div class="soft" style="margin-top:4px;">Load your Excel model and explore every sheet.</div></div>', unsafe_allow_html=True)

    up = st.file_uploader("Upload Excel model (.xlsx)", type=["xlsx"])
    if up:
        try:
            st.session_state.dfs = load_workbook(up.getvalue())
            st.session_state.profiles = [sheet_profile(df, name) for name, df in st.session_state.dfs.items()]
            st.session_state.corpus = build_corpus(st.session_state.profiles)
            st.success("Workbook loaded.")
        except Exception as e:
            st.error(f"Failed to load Excel: {e}")

    # Download buttons for your submitted files (optional helper)
    dl1, dl2 = st.columns([0.5,0.5])
    with dl1:
        if os.path.exists(DEFAULT_XLSX_PATH):
            with open(DEFAULT_XLSX_PATH, "rb") as f:
                st.download_button("⬇️ Download My Walmart Model", f, file_name="WMT_Valuation.xlsx", use_container_width=True)
    with dl2:
        if os.path.exists(DEFAULT_ASSIGNMENT_PDF):
            with open(DEFAULT_ASSIGNMENT_PDF, "rb") as f:
                st.download_button("⬇️ Download Assignment Brief", f, file_name="Assignment1.pdf", use_container_width=True)

    if not st.session_state.dfs:
        st.warning("No workbook loaded yet.")
    else:
        sheet_names = list(st.session_state.dfs.keys())
        st.markdown("### 📑 Sheets")
        selected = st.selectbox("Choose a sheet", sheet_names, index=0)

        df = st.session_state.dfs[selected]
        with st.expander(f"Preview — {selected}", expanded=True):
            st.dataframe(df, use_container_width=True, height=400)

        # Quick facts
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Rows", df.shape[0])
        with c2: st.metric("Columns", df.shape[1])
        with c3: st.metric("Numeric Columns", len(numeric_cols(df)))

        st.markdown("### 📈 Build a Chart")
        # Chart builder
        dcols = date_cols(df)
        ncols = numeric_cols(df)
        a1, a2, a3 = st.columns(3)
        with a1:
            x_col = st.selectbox("X axis", dcols + list(df.columns), index=0 if dcols else 0)
        with a2:
            y_cols = st.multiselect("Y axis (numeric)", ncols, default=ncols[:1])
        with a3:
            chart_type = st.selectbox("Chart type", ["line", "area", "bar"])

        plot_df = df.copy()
        # try convert date x if needed
        if x_col in plot_df.columns:
            try:
                plot_df[x_col] = pd.to_datetime(plot_df[x_col], errors="ignore", infer_datetime_format=True)
            except Exception:
                pass

        if y_cols:
            if chart_type == "line":
                st.line_chart(plot_df.set_index(x_col)[y_cols], use_container_width=True)
            elif chart_type == "area":
                st.area_chart(plot_df.set_index(x_col)[y_cols], use_container_width=True)
            else:
                st.bar_chart(plot_df.set_index(x_col)[y_cols], use_container_width=True)
        else:
            st.info("Select at least one numeric Y column to plot.")

        # Quick describe (numeric)
        with st.expander("Descriptive statistics (numeric)", expanded=False):
            st.dataframe(df[ncols].describe().T if ncols else pd.DataFrame({"note":["No numeric columns detected."]}), use_container_width=True)

# RIGHT: Chatbot
with right:
    st.markdown(f"""
    <div class="card" style="margin-bottom:10px;">
      <div style="display:flex;align-items:center;justify-content:space-between;">
        <div style="font-weight:800;font-size:1.15rem;">💬 Analyst Chat</div>
        <span class="pill">Grounded in your sheets</span>
      </div>
      <div class="soft">Ask about assumptions, drivers, WACC, segment growth, margins, FCF, or sensitivity. I’ll cite sheet names in answers.</div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="chat-wrap">', unsafe_allow_html=True)

    # render chat (skip the hidden system)
    for m in st.session_state.messages:
        with st.chat_message("assistant" if m["role"]=="assistant" else "user"):
            st.write(m["content"])

    user_text = st.chat_input("Ask a question about the workbook or Walmart…")
    if user_text:
        st.session_state.messages.append({"role":"user","content":user_text})

        # First: try some simple computed patterns (e.g., CAGR, mean of a column)
        replied = False
        df_map = st.session_state.dfs

        # Pattern: "CAGR of <col> in <sheet>"
        m = re.search(r"cagr of ([\w \-/\(\)%\.]+) in ([\w \-]+)", user_text, flags=re.I)
        if m and not replied:
            col, sh = m.group(1).strip(), m.group(2).strip()
            if sh in df_map and col in df_map[sh].columns:
                dfx = df_map[sh].dropna(subset=[col]).copy()
                # try find a time or index order
                idx = None
                for c in date_cols(dfx):
                    idx = c; break
                if idx is None:
                    # fallback to first column if looks like year
                    for c in dfx.columns:
                        if re.fullmatch(r"\d{4}", str(dfx[c].iloc[0])) or "year" in c.lower():
                            idx = c; break
                try:
                    if idx:
                        dfx = dfx.sort_values(by=idx)
                    v0 = float(dfx[col].iloc[0]); vN = float(dfx[col].iloc[-1]); n = max(1, len(dfx)-1)
                    cagr = (vN / v0)**(1/n) - 1 if v0>0 and vN>0 else float("nan")
                    ans = f"CAGR of **{col}** in **{sh}** ≈ **{cagr:.2%}** (from {v0:,.2f} to {vN:,.2f} over {n} periods)."
                except Exception as e:
                    ans = f"I couldn’t compute CAGR from **{sh}** → **{col}** ({e})."
                st.session_state.messages.append({"role":"assistant","content":ans})
                replied = True

        # If not matched, route to LLM grounded in sheet corpus
        if not replied:
            corpus = st.session_state.corpus if st.session_state.corpus else ["(no workbook context loaded)"]
            reply = chat_over_sheets(user_text, corpus)
            st.session_state.messages.append({"role":"assistant","content":reply})

        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# Footer
st.markdown("<hr style='border-color:#1f2937;'>", unsafe_allow_html=True)
st.caption("© Walmart Valuation Explorer — for coursework demonstration. This tool reads your Excel model and answers questions grounded in your sheets.")
