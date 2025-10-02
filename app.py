# app.py — Walmart Valuation Explorer (Analyst UI + Chat)
# - Sidebar sheet navigator (searchable)
# - Finance Overview (Revenue, EBITDA, FCF, margins, WACC if found)
# - Smart charts per sheet (handles year headers & real dates)
# - Per-sheet English summary
# - Chatbot grounded in workbook + general Walmart context (Azure OpenAI optional)

import os, io, re, json, pkgutil, sys
from typing import Dict, List, Tuple
import numpy as np
import pandas as pd
import streamlit as st

# ============ CONFIG ============
FILE_NAME = "FIN42030 WMT Valuation (2).xlsx"   # Excel kept next to app.py
PAGE_TITLE = "Walmart Valuation Explorer"


# Hard-coded Azure OpenAI (optional; safe fallback if unset)
AZURE_OPENAI_ENDPOINT    = "https://testaisentiment.openai.azure.com/"
AZURE_OPENAI_API_KEY     = "cb1c33772b3c4edab77db69ae18c9a43"
AZURE_OPENAI_API_VERSION = "2024-02-15-preview"
AZURE_OPENAI_DEPLOYMENT  = "aipocexploration"

st.set_page_config(page_title=PAGE_TITLE, page_icon="📊", layout="wide")

# ------ Theme (readable, analyst-y) ------
BG, PANEL, BORDER, TEXT, MUTED = "#0b1220", "#0f172a", "#233043", "#e5edf7", "#98a6bd"
PRIMARY = "#22d3ee"
st.markdown(f"""
<style>
  html, body, .block-container {{ background:{BG}; color:{TEXT}; }}
  .card {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:14px; padding:16px; }}
  .headline {{ font-size:1.8rem; font-weight:800; letter-spacing:-.01em; }}
  .soft {{ color:{MUTED}; }}
  .pill {{ display:inline-flex; gap:8px; padding:6px 12px; border-radius:999px; font-size:.85rem; 
           background:rgba(34,211,238,.12); color:{PRIMARY}; border:1px solid rgba(34,211,238,.28); }}
  .kpi {{ background:#0b1324; border:1px solid {BORDER}; border-radius:14px; padding:14px; text-align:center; }}
  .kpi .v {{ font-size:1.3rem; font-weight:800; }}
  .kpi .l {{ font-size:.85rem; color:{MUTED}; }}
  /* make sidebar comfy */
  .css-1d391kg, .css-12oz5g7 {{ padding-top: 1rem; }}
</style>
""", unsafe_allow_html=True)

# ------ Optional LLM (Azure OpenAI) ------
client = None
OPENAI_OK = False
if AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY and AZURE_OPENAI_DEPLOYMENT:
    try:
        from openai import AzureOpenAI
        client = AzureOpenAI(
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            api_key=AZURE_OPENAI_API_KEY,
            api_version=AZURE_OPENAI_API_VERSION
        )
        OPENAI_OK = True
    except Exception:
        OPENAI_OK = False

def ask_gpt(messages, temperature=0.2, max_tokens=900):
    if not OPENAI_OK:
        return "(LLM not configured.)"
    try:
        r = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens
        )
        return r.choices[0].message.content
    except Exception as e:
        return f"(LLM error: {e})"

# ------ Helpers ------
def need_openpyxl() -> bool:
    return pkgutil.find_loader("openpyxl") is None

@st.cache_data(show_spinner=False)
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    if need_openpyxl():
        raise ImportError("openpyxl is not installed in this environment. Add it to requirements.txt and redeploy.")
    xl = pd.ExcelFile(path, engine="openpyxl")
    return {name: xl.parse(name) for name in xl.sheet_names}

def is_date_col(s: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(s): return True
    if s.dtype == object:
        try:
            parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
            return parsed.notna().mean() >= 0.6
        except Exception: return False
    return False

def num_cols(df): return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
def date_cols(df): return [c for c in df.columns if is_date_col(df[c])]

def find_year_header_cols(df: pd.DataFrame) -> List[str]:
    pat = re.compile(r"(?:19|20)\d{2}")
    yearish = [c for c in df.columns if pat.search(str(c))]
    if yearish:
        tmp = df.copy()
        for c in yearish: tmp[c] = pd.to_numeric(tmp[c], errors="coerce")
        yearish = [c for c in yearish if tmp[c].notna().any()]
    return yearish

def wide_years_to_long(df: pd.DataFrame, year_cols: List[str], label_col=None) -> pd.DataFrame | None:
    if not year_cols: return None
    if label_col is None:
        non_year = [c for c in df.columns if c not in year_cols]
        label_col = non_year[0] if non_year else None
    work = df.copy()
    for yc in year_cols: work[yc] = pd.to_numeric(work[yc], errors="coerce")
    long = work.melt(id_vars=[label_col] if label_col else None, value_vars=year_cols,
                     var_name="Year", value_name="Value")
    long["Year"] = long["Year"].astype(str).str.extract(r"((?:19|20)\d{2})").astype(float)
    long = long.dropna(subset=["Year","Value"])
    if long.empty: return None
    long["Year"] = long["Year"].astype(int)
    if label_col and label_col in long.columns:
        vard = long.groupby(label_col)["Value"].var().sort_values(ascending=False)
        long = long[long[label_col].isin(list(vard.head(5).index))]
    return long

def coerce_numeric(df: pd.DataFrame, cols: List[str]) -> tuple[pd.DataFrame, List[str]]:
    out = df.copy(); keep=[]
    for c in cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")
        if out[c].notna().any(): keep.append(c)
    return out, keep

def sheet_profile(df: pd.DataFrame, name: str) -> dict:
    # small profile object that feeds the chatbot
    stats = {}
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_numeric_dtype(s):
            stats[str(c)] = {
                "count": int(s.count()),
                "mean": float(np.nanmean(s)) if s.count() else np.nan,
                "min": float(np.nanmin(s)) if s.count() else np.nan,
                "max": float(np.nanmax(s)) if s.count() else np.nan,
            }
    with pd.option_context("display.max_columns", 12, "display.width", 1000):
        preview = df.head(8).to_markdown(index=False)
    return {"name":name, "rows":int(df.shape[0]), "cols":int(df.shape[1]),
            "columns":[str(x) for x in df.columns], "num_stats":stats, "preview":preview}

def build_corpus(profiles: List[dict]) -> List[str]:
    chunks=[]
    for p in profiles:
        chunks.append(
f"""Sheet: {p['name']}
Columns: {', '.join(p['columns'][:24])}
Numeric summary: {json.dumps(p['num_stats'])[:1600]}
Preview:
{p['preview']}
""")
    return chunks or ["(no workbook context)"]

def simple_retriever(query: str, chunks: List[str], top_k: int=4) -> List[str]:
    toks = re.findall(r"[a-z0-9\-%\.]+", (query or "").lower())
    scored=[]
    for ch in chunks:
        t=ch.lower(); score=sum(t.count(k) for k in toks if len(k)>2)
        scored.append((score, ch))
    scored.sort(key=lambda x:x[0], reverse=True)
    return [c for s,c in scored[:top_k] if s>0] or [scored[0][1]]

def guess_overview_metrics(dfs: Dict[str, pd.DataFrame]) -> dict:
    m = {"revenue":None, "ebitda":None, "fcf":None, "wacc":None}
    for sname, df in dfs.items():
        cols_lower = [str(c).lower() for c in df.columns]
        low = sname.lower()
        if any(k in low for k in ["income","statement","integrated","summary","forecast","valuation","assumption","wacc","market"]):
            # revenue
            if m["revenue"] is None:
                cands = [c for c in cols_lower if "revenue" in c or "sales" in c]
                if cands: m["revenue"]=(sname, df.columns[cols_lower.index(cands[0])])
            # ebitda
            if m["ebitda"] is None:
                cands = [c for c in cols_lower if "ebitda" in c]
                if cands: m["ebitda"]=(sname, df.columns[cols_lower.index(cands[0])])
            # fcf
            if m["fcf"] is None:
                cands = [c for c in cols_lower if "free cash" in c or c.startswith("fcf")]
                if cands: m["fcf"]=(sname, df.columns[cols_lower.index(cands[0])])
            # wacc (as scalar column or cell)
            if m["wacc"] is None:
                wcols = [c for c in cols_lower if "wacc" in c or "discount rate" in c]
                if wcols: m["wacc"]=(sname, df.columns[cols_lower.index(wcols[0])])
    return m

def latest_and_cagr(df: pd.DataFrame, value_col: str) -> tuple[str, str]:
    dcols = date_cols(df)
    series = None
    if dcols:
        t=dcols[0]; x=df.copy(); x[t]=pd.to_datetime(x[t], errors="coerce")
        x=x.dropna(subset=[t]).sort_values(t); x[value_col]=pd.to_numeric(x[value_col], errors="coerce")
        series = x[value_col].dropna()
    else:
        year_cols = find_year_header_cols(df)
        if year_cols:
            long = wide_years_to_long(df, year_cols)
            if long is not None and not long.empty:
                series = long.groupby("Year")["Value"].sum().sort_index()
        else:
            series = pd.to_numeric(df[value_col], errors="coerce").dropna()
    if series is None or len(series)==0: return ("—","—")
    try:
        first,last,n=float(series.iloc[0]),float(series.iloc[-1]),max(1,len(series)-1)
        cagr=(last/first)**(1/n)-1 if first>0 and last>0 else np.nan
        return (f"{last:,.0f}", f"{cagr:.1%}" if np.isfinite(cagr) else "—")
    except Exception:
        return (f"{series.iloc[-1]:,.0f}", "—")

# ------ Load Excel ------
header_left, header_right = st.columns([0.8,0.2])
with header_left:
    st.markdown(f'<div class="headline">{PAGE_TITLE}</div>', unsafe_allow_html=True)
    st.caption(f"Python: {sys.executable} • openpyxl: {'yes' if not need_openpyxl() else 'no'}")

if not os.path.exists(FILE_NAME):
    st.error(f"File not found: {FILE_NAME}. Keep your Excel next to app.py in the repo.")
    st.stop()

try:
    dfs = load_workbook(FILE_NAME)
except Exception as e:
    st.error(str(e)); st.stop()

# Profiles & corpus for chat
profiles = [sheet_profile(df, name) for name, df in dfs.items()]
corpus = build_corpus(profiles)

# ------ Sidebar: sheet navigator + search ------
st.sidebar.markdown("### Sheets")
query = st.sidebar.text_input("Filter sheets", "")
sheet_names = [n for n in dfs.keys() if query.lower() in n.lower()]
if not sheet_names:
    st.sidebar.info("No sheets match your filter.")
    sheet_names = list(dfs.keys())
selected = st.sidebar.radio("Select a sheet", sheet_names, label_visibility="collapsed", index=0)

# ------ Overview KPIs (analyst-friendly) ------
st.markdown("## Overview")
metrics = guess_overview_metrics(dfs)
rev_val = ebitda_val = fcf_val = "—"; rev_delta = None; wacc_val = "—"

if metrics["revenue"]:
    v, d = latest_and_cagr(dfs[metrics["revenue"][0]], metrics["revenue"][1]); rev_val, rev_delta = v, d if d!="—" else None
if metrics["ebitda"]:
    v, _ = latest_and_cagr(dfs[metrics["ebitda"][0]], metrics["ebitda"][1]); ebitda_val = v
if metrics["fcf"]:
    v, _ = latest_and_cagr(dfs[metrics["fcf"][0]], metrics["fcf"][1]); fcf_val = v
if metrics["wacc"]:
    # try to read a scalar-like WACC
    wdf = dfs[metrics["wacc"][0]]
    try:
        s = pd.to_numeric(wdf[metrics["wacc"][1]], errors="coerce").dropna()
        if not s.empty: wacc_val = f"{float(s.iloc[0])*100:.2f}%" if s.iloc[0] < 1 else f"{float(s.iloc[0]):.2f}%"
    except Exception: pass

k1,k2,k3,k4 = st.columns(4)
with k1: st.markdown(f'<div class="kpi"><div class="l">Revenue (latest)</div><div class="v">{rev_val}</div><div class="l">{rev_delta or ""}</div></div>', unsafe_allow_html=True)
with k2: st.markdown(f'<div class="kpi"><div class="l">EBITDA (latest)</div><div class="v">{ebitda_val}</div></div>', unsafe_allow_html=True)
with k3: st.markdown(f'<div class="kpi"><div class="l">Free Cash Flow (latest)</div><div class="v">{fcf_val}</div></div>', unsafe_allow_html=True)
with k4: st.markdown(f'<div class="kpi"><div class="l">WACC</div><div class="v">{wacc_val}</div></div>', unsafe_allow_html=True)

st.divider()

# ------ Main: selected sheet viewer ------
st.markdown(f"## {selected}")

df = dfs[selected]

with st.expander("Preview", expanded=False):
    st.dataframe(df, use_container_width=True, height=420)

# Summarise in English
def summarise_sheet(df: pd.DataFrame, name: str) -> str:
    cols = [str(c) for c in df.columns]
    hints=[]
    low=name.lower()
    if any(k in low for k in ["income","p&l","profit"]):
        hints.append("appears to be an *Income Statement* layout — look at revenue growth, gross/operating margin, EPS.")
    if "balance" in low or low.endswith(" bs"):
        hints.append("looks like a *Balance Sheet* — check working capital, leverage, cash & debt trends.")
    if any(k in low for k in ["cash","fcf","free cash","cf"]):
        hints.append("reads like *Cash Flow* — compare CFO vs CAPEX to infer FCF.")
    if any(k in low for k in ["assumption","wacc","market","drivers"]):
        hints.append("contains *Assumptions/WACC* — verify discount rate, growth, terminal inputs that drive valuation.")
    year_cols = find_year_header_cols(df)
    dcols = date_cols(df)
    axis_hint = "time axis from the first date column" if dcols else ("years in headers" if year_cols else "no obvious time axis")
    return f"- Columns detected: {', '.join(cols[:12])}{'…' if len(cols)>12 else ''}\n- Chart basis: {axis_hint}\n" + (f"- Notes: {', '.join(hints)}" if hints else "")

with st.expander("What’s in this sheet?", expanded=True):
    st.markdown(summarise_sheet(df, selected))

# Charts
dcols = date_cols(df)
ncols = num_cols(df)
year_cols = find_year_header_cols(df)

st.markdown("### Time-series")
plotted_ts=False
if dcols and ncols:
    x=dcols[0]
    dff=df.copy()
    dff[x]=pd.to_datetime(dff[x], errors="coerce")
    dff=dff.dropna(subset=[x]).sort_values(x)
    dff, safe = coerce_numeric(dff, ncols); safe=safe[:6]
    if safe and not dff.empty:
        st.line_chart(dff.set_index(x)[safe], use_container_width=True); plotted_ts=True
if not plotted_ts and year_cols:
    long = wide_years_to_long(df, year_cols)
    if long is not None and not long.empty:
        label_cols=[c for c in long.columns if c not in ("Year","Value")]
        if label_cols:
            piv = long.pivot_table(index="Year", columns=label_cols[0], values="Value", aggfunc="mean").sort_index()
            st.line_chart(piv, use_container_width=True); plotted_ts=True
if not plotted_ts:
    st.info("No reliable time axis found.")

st.markdown("### Category summary")
# small-cardinality text columns
cats=[]
for c in df.columns:
    if pd.api.types.is_numeric_dtype(df[c]): continue
    try:
        u=df[c].nunique(dropna=True)
        if 1 < u <= 20: cats.append(c)
    except Exception: pass
if cats:
    cat=cats[0]
    co=df.copy(); agg_cols=[]
    for c in df.columns:
        if c==cat: continue
        co[c]=pd.to_numeric(co[c], errors="coerce")
        if co[c].notna().any(): agg_cols.append(c)
    show=agg_cols[:3]
    if show:
        try:
            agg=co.groupby(cat)[show].sum(numeric_only=True).sort_values(show[0], ascending=False).head(20)
            st.bar_chart(agg, use_container_width=True)
        except Exception:
            st.caption("Could not render grouped bars.")
    else:
        st.caption("No numeric data to aggregate by category.")
else:
    st.caption("No small-cardinality category column detected.")

st.markdown("### Numeric histograms")
co=df.copy(); num_cands=[]
for c in df.columns:
    co[c]=pd.to_numeric(co[c], errors="coerce")
    if co[c].notna().any(): num_cands.append(c)
for c in num_cands[:3]:
    st.markdown(f"**{c}**")
    try:
        st.bar_chart(co[c].dropna().reset_index(drop=True), use_container_width=True)
    except Exception:
        st.caption("Could not render histogram.")

st.markdown("### Correlation (numeric)")
num_for_corr=[c for c in co.columns if pd.api.types.is_numeric_dtype(co[c]) and co[c].notna().any()]
if len(num_for_corr) >= 2:
    corr = co[num_for_corr].corr(numeric_only=True)
    st.dataframe(corr.style.background_gradient(cmap="Blues"), use_container_width=True)
else:
    st.caption("Not enough numeric columns for correlations.")

st.divider()

# ------ Chatbot (grounded in workbook) ------
st.markdown("## Analyst Chat")
st.caption("Ask anything about Walmart or this workbook. I’ll cite sheet names when I use workbook numbers.")
if "chat" not in st.session_state:
    st.session_state.chat = [{"role":"assistant","content":"Hi! Ask about assumptions, WACC, margins, growth, FCF, or any tab."}]

# render history
for m in st.session_state.chat:
    with st.chat_message("assistant" if m["role"]=="assistant" else "user"):
        st.write(m["content"])

prompt = st.chat_input("Type your question…")
SYSTEM = """You are a Walmart (NYSE: WMT) valuation analyst.
Use the provided workbook context if it contains the answer; name the sheet(s) you used.
If the workbook lacks the data, say so and answer with general finance knowledge about Walmart only if relevant.
Be concise; provide numeric answers when possible; note assumptions briefly.
"""

def chat_over_sheets(user_msg: str) -> str:
    ctx_parts = simple_retriever(user_msg, corpus, top_k=4)
    context = "\n\n---\n\n".join([str(c) for c in ctx_parts])
    messages=[
        {"role":"system","content":SYSTEM},
        {"role":"user","content": f"Workbook context:\n{context}\n\nQuestion: {user_msg}"}
    ]
    out = ask_gpt(messages)
    if out.startswith("("):
        return "LLM not configured. Add Azure OpenAI env vars to enable chat."
    return out

if prompt:
    st.session_state.chat.append({"role":"user","content":prompt})
    reply = chat_over_sheets(prompt)
    st.session_state.chat.append({"role":"assistant","content":reply})
    st.rerun()
