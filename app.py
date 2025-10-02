# app.py — Walmart Valuation Explorer (Readable UI + Analyst KPIs + Chat)
# - Sidebar navigator (search + select) for all sheets
# - Finance KPIs (Revenue, EBITDA, FCF, WACC) with CAGR where possible
# - Clean charts using Plotly (auto-detect date vs. "years-in-columns")
# - Per-sheet English summary
# - Chatbot grounded in workbook (optional via Azure OpenAI env vars)
# - Robust to "Unnamed" columns and non-numeric text

import os, re, json, sys, pkgutil
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ============== CONFIG ==============
FILE_NAME = "FIN42030 WMT Valuation (2).xlsx"  # keep this file next to app.py

# Optional Azure OpenAI (set as repository secrets/env in Streamlit Cloud)
# Hard-coded Azure OpenAI (optional; safe fallback if unset)
AZURE_OPENAI_ENDPOINT    = "https://testaisentiment.openai.azure.com/"
AZURE_OPENAI_API_KEY     = "cb1c33772b3c4edab77db69ae18c9a43"
AZURE_OPENAI_API_VERSION = "2024-02-15-preview"
AZURE_OPENAI_DEPLOYMENT  = "aipocexploration"

# ============== PAGE / THEME ==============
st.set_page_config(page_title="Walmart Valuation Explorer", page_icon="📊", layout="wide")

# High-contrast, readable dark theme
BG, PANEL, BORDER, TEXT, MUTED = "#0b1220", "#0f172a", "#233043", "#e8f1ff", "#9ab0cf"
PRIMARY = "#22d3ee"

st.markdown(f"""
<style>
  :root {{ --text:{TEXT}; --muted:{MUTED}; --panel:{PANEL}; --border:{BORDER}; --primary:{PRIMARY}; }}
  html, body, .block-container {{ background:{BG}; color:{TEXT}; font-size:16px; }}
  .headline {{ font-size:2rem; font-weight:900; letter-spacing:-.01em; margin-bottom:.5rem; }}
  .subtle  {{ color:{MUTED}; }}
  .card    {{ background:{PANEL}; border:1px solid {BORDER}; border-radius:16px; padding:16px; }}
  .kpi     {{ background:#0b1324; border:1px solid {BORDER}; border-radius:16px; padding:18px; }}
  .kpi h4  {{ margin:0 0 .4rem 0; font-size:.95rem; color:{MUTED}; }}
  .kpi .v  {{ font-size:1.6rem; font-weight:900; }}
  .kpi .d  {{ font-size:.9rem; color:{MUTED}; }}
  /* sidebar readability */
  section[data-testid="stSidebar"] .css-1v0mbdj, section[data-testid="stSidebar"] .css-1d391kg {{ padding-top:8px; }}
  /* fix long selectbox entries */
  div[data-baseweb="select"] span {{ white-space:normal !important; }}
</style>
""", unsafe_allow_html=True)

# ============== OPTIONAL LLM CLIENT ==============
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
        return "(LLM not configured. Add Azure env vars to enable chat.)"
    try:
        r = client.chat.completions.create(
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens,
        )
        return r.choices[0].message.content
    except Exception as e:
        return f"(LLM error: {e})"

# ============== HELPERS ==============
def need_openpyxl() -> bool:
    return pkgutil.find_loader("openpyxl") is None

@st.cache_data(show_spinner=False)
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    if need_openpyxl():
        raise ImportError("openpyxl not installed in this environment. Add it to requirements.txt and redeploy.")
    xl = pd.ExcelFile(path, engine="openpyxl")
    dfs = {}
    for name in xl.sheet_names:
        df = xl.parse(name)
        # clean: drop empty columns, trim headers, remove extra "Unnamed" if totally empty
        new_cols = []
        for c in df.columns:
            s = str(c).strip()
            new_cols.append(s)
        df.columns = new_cols
        # drop fully empty columns
        empty_cols = [c for c in df.columns if df[c].isna().all()]
        if empty_cols:
            df = df.drop(columns=empty_cols)
        dfs[name] = df
    return dfs

def is_date_col(s: pd.Series) -> bool:
    if pd.api.types.is_datetime64_any_dtype(s): return True
    if s.dtype == object:
        try:
            parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
            return parsed.notna().mean() >= 0.6
        except Exception: return False
    return False

def num_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

def date_cols(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if is_date_col(df[c])]

def find_year_header_cols(df: pd.DataFrame) -> List[str]:
    pat = re.compile(r"(?:19|20)\d{2}")
    yearish = [c for c in df.columns if pat.search(str(c))]
    if not yearish: return []
    tmp = df.copy()
    for c in yearish:
        tmp[c] = pd.to_numeric(tmp[c], errors="coerce")
    return [c for c in yearish if tmp[c].notna().any()]

def wide_years_to_long(df: pd.DataFrame, year_cols: List[str], label_col=None) -> pd.DataFrame | None:
    if not year_cols: return None
    if label_col is None:
        non_year = [c for c in df.columns if c not in year_cols]
        label_col = non_year[0] if non_year else None
    work = df.copy()
    for yc in year_cols:
        work[yc] = pd.to_numeric(work[yc], errors="coerce")
    long = work.melt(id_vars=[label_col] if label_col else None, value_vars=year_cols,
                     var_name="Year", value_name="Value")
    long["Year"] = long["Year"].astype(str).str.extract(r"((?:19|20)\d{2})").astype(float)
    long = long.dropna(subset=["Year","Value"])
    if long.empty: return None
    long["Year"] = long["Year"].astype(int)
    # choose up to 5 most varying series to avoid clutter
    if label_col and label_col in long.columns:
        var_rank = long.groupby(label_col)["Value"].var().sort_values(ascending=False)
        long = long[long[label_col].isin(list(var_rank.head(5).index))]
    return long

def coerce_numeric(df: pd.DataFrame, cols: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    out = df.copy(); keep=[]
    for c in cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")
        if out[c].notna().any(): keep.append(c)
    return out, keep

def safe_preview(df: pd.DataFrame, n=8) -> str:
    try:
        import tabulate as _  # noqa
        return df.head(n).to_markdown(index=False)
    except Exception:
        return df.head(n).to_string(index=False)

def sheet_summary_text(df: pd.DataFrame, name: str) -> str:
    cols = [str(c) for c in df.columns]
    hints=[]
    low=name.lower()
    if any(k in low for k in ["income","p&l","profit","is"]): hints.append("Appears to be an **Income Statement**; focus on growth & margins.")
    if "balance" in low or low.endswith(" bs"):             hints.append("Looks like a **Balance Sheet**; check working capital & leverage.")
    if any(k in low for k in ["cash","fcf","free cash","cf"]): hints.append("Reads like a **Cash Flow**; CFO vs CAPEX → FCF.")
    if any(k in low for k in ["assumption","wacc","market","drivers","valuation"]): hints.append("Contains **Assumptions / WACC** driving valuation.")
    year_cols = find_year_header_cols(df)
    dcols = date_cols(df)
    axis_hint = "time axis from a date column" if dcols else ("years in headers" if year_cols else "no obvious time axis")
    return f"- Columns: {', '.join(cols[:12])}{'…' if len(cols)>12 else ''}\n- Chart basis: {axis_hint}\n" + (f"- Notes: {' '.join(hints)}" if hints else "")

def guess_overview_metrics(dfs: Dict[str, pd.DataFrame]) -> dict:
    """Heuristic scan to find Revenue/EBITDA/FCF/WACC columns."""
    m = {"revenue":None, "ebitda":None, "fcf":None, "wacc":None}
    for sname, df in dfs.items():
        low = sname.lower()
        if any(k in low for k in ["income","statement","integrated","summary","forecast","valuation","assumption","wacc","market","fin"]):
            cols_lower = [str(c).lower() for c in df.columns]
            # revenue
            if m["revenue"] is None:
                cand=[c for c in cols_lower if "revenue" in c or "sales" in c]
                if cand: m["revenue"]=(sname, df.columns[cols_lower.index(cand[0])])
            # ebitda
            if m["ebitda"] is None:
                cand=[c for c in cols_lower if "ebitda" in c]
                if cand: m["ebitda"]=(sname, df.columns[cols_lower.index(cand[0])])
            # fcf
            if m["fcf"] is None:
                cand=[c for c in cols_lower if "free cash" in c or c.startswith("fcf")]
                if cand: m["fcf"]=(sname, df.columns[cols_lower.index(cand[0])])
            # wacc
            if m["wacc"] is None:
                cand=[c for c in cols_lower if "wacc" in c or "discount rate" in c]
                if cand: m["wacc"]=(sname, df.columns[cols_lower.index(cand[0])])
    return m

def series_latest_and_cagr(df: pd.DataFrame, value_col: str) -> Tuple[str,str]:
    """Return latest value and rough CAGR if time order can be inferred."""
    dcols = date_cols(df); series=None
    if dcols:
        t=dcols[0]
        x=df.copy(); x[t]=pd.to_datetime(x[t], errors="coerce"); x=x.dropna(subset=[t]).sort_values(t)
        x[value_col]=pd.to_numeric(x[value_col], errors="coerce"); series=x[value_col].dropna()
    else:
        ycols = find_year_header_cols(df)
        if ycols:
            long = wide_years_to_long(df, ycols)
            if long is not None and not long.empty:
                series = long.groupby("Year")["Value"].sum().sort_index()
        else:
            series = pd.to_numeric(df[value_col], errors="coerce").dropna()
    if series is None or series.empty: return ("—","—")
    try:
        first,last,n=float(series.iloc[0]),float(series.iloc[-1]),max(1,len(series)-1)
        cagr=(last/first)**(1/n)-1 if first>0 and last>0 else np.nan
        return (f"{last:,.0f}", f"{cagr:.1%}" if np.isfinite(cagr) else "—")
    except Exception:
        return (f"{series.iloc[-1]:,.0f}", "—")

def build_chat_corpus(dfs: Dict[str,pd.DataFrame]) -> List[str]:
    chunks=[]
    for name, df in dfs.items():
        preview = safe_preview(df, n=8)
        # small numeric summary (capped)
        stats={}
        for c in df.columns:
            s=df[c]
            if pd.api.types.is_numeric_dtype(s):
                stats[str(c)]={"mean":float(np.nanmean(s)) if s.count() else np.nan,
                               "min":float(np.nanmin(s)) if s.count() else np.nan,
                               "max":float(np.nanmax(s)) if s.count() else np.nan}
        chunks.append(
            f"Sheet: {name}\nColumns: {', '.join(map(str, df.columns))[:350]}\n"
            f"Numeric summary: {json.dumps(stats)[:1600]}\nPreview:\n{preview}\n"
        )
    return chunks

def retrieve_context(q:str, chunks:List[str], k:int=4)->List[str]:
    toks = re.findall(r"[a-z0-9\-\.%]+", (q or "").lower())
    scored=[]
    for ch in chunks:
        t=ch.lower(); score=sum(t.count(tok) for tok in toks if len(tok)>2)
        scored.append((score, ch))
    scored.sort(key=lambda x:x[0], reverse=True)
    return [c for s,c in scored[:k] if s>0] or [scored[0][1]]

# ============== LOAD WORKBOOK ==============
st.markdown(f'<div class="headline">Walmart Valuation Explorer</div><div class="subtle">Python: {sys.executable} • openpyxl: {"yes" if not need_openpyxl() else "no"}</div>', unsafe_allow_html=True)

if not os.path.exists(FILE_NAME):
    st.error(f"File not found: {FILE_NAME}. Place it next to app.py.")
    st.stop()

try:
    dfs = load_workbook(FILE_NAME)
except Exception as e:
    st.error(str(e)); st.stop()

# ============== SIDEBAR SHEET PICKER ==============
st.sidebar.header("Sheets")
filter_text = st.sidebar.text_input("Filter sheets", "")
sheet_options = [n for n in dfs.keys() if filter_text.lower() in n.lower()] or list(dfs.keys())
selected_sheet = st.sidebar.selectbox("Select a sheet", sheet_options)

# ============== OVERVIEW KPIs ==============
st.markdown("## Overview")
metrics = guess_overview_metrics(dfs)

rev_val = ebitda_val = fcf_val = "—"
rev_delta = None
wacc_val = "—"

if metrics["revenue"]:
    v,d = series_latest_and_cagr(dfs[metrics["revenue"][0]], metrics["revenue"][1])
    rev_val, rev_delta = v, (d if d!="—" else None)
if metrics["ebitda"]:
    v,_ = series_latest_and_cagr(dfs[metrics["ebitda"][0]], metrics["ebitda"][1]); ebitda_val=v
if metrics["fcf"]:
    v,_ = series_latest_and_cagr(dfs[metrics["fcf"][0]], metrics["fcf"][1]); fcf_val=v
if metrics["wacc"]:
    try:
        s = pd.to_numeric(dfs[metrics["wacc"][0]][metrics["wacc"][1]], errors="coerce").dropna()
        if not s.empty:
            wacc_val = f"{float(s.iloc[0])*100:.2f}%" if s.iloc[0] < 1.0 else f"{float(s.iloc[0]):.2f}%"
    except Exception:
        pass

k1,k2,k3,k4 = st.columns(4)
with k1: st.markdown(f'<div class="kpi"><h4>Revenue (latest)</h4><div class="v">{rev_val}</div><div class="d">{rev_delta or ""}</div></div>', unsafe_allow_html=True)
with k2: st.markdown(f'<div class="kpi"><h4>EBITDA (latest)</h4><div class="v">{ebitda_val}</div></div>', unsafe_allow_html=True)
with k3: st.markdown(f'<div class="kpi"><h4>Free Cash Flow (latest)</h4><div class="v">{fcf_val}</div></div>', unsafe_allow_html=True)
with k4: st.markdown(f'<div class="kpi"><h4>WACC</h4><div class="v">{wacc_val}</div></div>', unsafe_allow_html=True)

st.divider()

# ============== SELECTED SHEET VIEW ==============
st.markdown(f"## {selected_sheet}")
df_raw = dfs[selected_sheet].copy()

# Drop "Unnamed" columns that are fully nan
candidate_unnamed = [c for c in df_raw.columns if c.lower().startswith("unnamed")]
for c in candidate_unnamed:
    if df_raw[c].isna().all():
        df_raw = df_raw.drop(columns=[c])

with st.expander("Preview", expanded=False):
    st.dataframe(df_raw, use_container_width=True, height=420)

with st.expander("What’s in this sheet?", expanded=True):
    st.markdown(sheet_summary_text(df_raw, selected_sheet))

# ---- Chart controls & plotting
dcols = date_cols(df_raw)
ncols = num_cols(df_raw)
ycols = find_year_header_cols(df_raw)

st.markdown("### Charts")

# time-series
ts_col = None
if dcols:
    ts_col = dcols[0]
    df_ts = df_raw.copy()
    df_ts[ts_col] = pd.to_datetime(df_ts[ts_col], errors="coerce")
    df_ts = df_ts.dropna(subset=[ts_col]).sort_values(ts_col)
    df_ts, safe = coerce_numeric(df_ts, ncols)
    targets = st.multiselect("Select series to plot", options=safe, default=safe[:3], help="Pick numeric columns to show")
    if targets:
        fig = px.line(df_ts, x=ts_col, y=targets)
        fig.update_layout(height=360, template="plotly_dark", margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No numeric columns found to plot against the time column.")
elif ycols:
    long = wide_years_to_long(df_raw, ycols)
    if long is not None and not long.empty:
        label_cols = [c for c in long.columns if c not in ("Year","Value")]
        label = label_cols[0] if label_cols else None
        if label:
            choices = sorted(long[label].unique().tolist())
            default_choices = choices[:min(4,len(choices))]
            selected = st.multiselect("Select series to plot", choices, default_choices)
            subset = long[long[label].isin(selected)] if selected else long
            fig = px.line(subset, x="Year", y="Value", color=label, markers=True)
            fig.update_layout(height=360, template="plotly_dark", margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(fig, use_container_width=True)
        else:
            fig = px.line(long, x="Year", y="Value")
            fig.update_layout(height=360, template="plotly_dark", margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Couldn’t reshape year columns into a time-series.")
else:
    st.info("No date or year structure detected to build a time-series.")

# category bars
st.markdown("### Category summary")
cat_cols=[]
for c in df_raw.columns:
    if pd.api.types.is_numeric_dtype(df_raw[c]): 
        continue
    try:
        u = df_raw[c].nunique(dropna=True)
        if 1 < u <= 20: cat_cols.append(c)
    except Exception: pass

if cat_cols:
    cat = st.selectbox("Group by", options=cat_cols, index=0)
    co = df_raw.copy()
    agg_candidates=[]
    for c in df_raw.columns:
        if c==cat: continue
        co[c]=pd.to_numeric(co[c], errors="coerce")
        if co[c].notna().any(): agg_candidates.append(c)
    yagg = st.multiselect("Aggregate numeric columns", options=agg_candidates, default=agg_candidates[:1])
    if yagg:
        agg = co.groupby(cat)[yagg].sum(numeric_only=True).sort_values(yagg[0], ascending=False).head(30).reset_index()
        fig = px.bar(agg, x=cat, y=yagg, barmode="group")
        fig.update_layout(height=360, template="plotly_dark", margin=dict(l=10,r=10,t=10,b=10), xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption("No numeric columns to aggregate by category.")
else:
    st.caption("No small-cardinality category column detected.")

# histograms
st.markdown("### Numeric histograms")
co = df_raw.copy()
numeric_candidates=[]
for c in df_raw.columns:
    co[c]=pd.to_numeric(co[c], errors="coerce")
    if co[c].notna().any():
        numeric_candidates.append(c)
if numeric_candidates:
    pick = st.multiselect("Choose columns for histograms", options=numeric_candidates, default=numeric_candidates[:2])
    for c in pick[:3]:
        fig = px.histogram(co, x=c, nbins=40)
        fig.update_layout(height=300, template="plotly_dark", margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
else:
    st.caption("No numeric data available for histograms.")

# correlation
st.markdown("### Correlation (numeric)")
num_for_corr = [c for c in co.columns if pd.api.types.is_numeric_dtype(co[c]) and co[c].notna().any()]
if len(num_for_corr) >= 2:
    corr = co[num_for_corr].corr(numeric_only=True)
    st.dataframe(corr.style.background_gradient(cmap="Blues"), use_container_width=True)
else:
    st.caption("Not enough numeric columns for correlations.")

st.divider()

# ============== ANALYST CHAT ==============
st.markdown("## Analyst Chat")
st.caption("Ask about assumptions, WACC, segment growth, margins, FCF, or sensitivity. I’ll cite sheet names when I use workbook numbers.")

if "chat" not in st.session_state:
    st.session_state.chat=[{"role":"assistant","content":"Hi! Which sheet or metric would you like to discuss?"}]

# build / cache corpus once
if "corpus" not in st.session_state:
    st.session_state.corpus = build_chat_corpus(dfs)

for m in st.session_state.chat:
    with st.chat_message("assistant" if m["role"]=="assistant" else "user"):
        st.write(m["content"])

user_q = st.chat_input("Type your question…")
SYS = """You are a Walmart valuation analyst. Prefer answers grounded in the provided workbook context; cite the sheet names you used. 
Be concise and numeric where possible. If the workbook lacks the answer, say so briefly then answer with general Walmart finance knowledge only if appropriate."""
def answer_with_context(q: str) -> str:
    ctx = retrieve_context(q, st.session_state.corpus, k=4)
    messages=[
        {"role":"system","content":SYS},
        {"role":"user","content": "Workbook context:\n" + "\n\n---\n\n".join(ctx) + f"\n\nQuestion: {q}"}
    ]
    out = ask_gpt(messages)
    if out.startswith("("):   # no LLM
        return "Chat is disabled (no Azure OpenAI env vars). You can still use the charts and summaries above."
    return out

if user_q:
    st.session_state.chat.append({"role":"user","content":user_q})
    st.session_state.chat.append({"role":"assistant","content":answer_with_context(user_q)})
    st.rerun()
