# app.py — Walmart Valuation Explorer (Readable UI + Sheet-grounded Chat)
# - Light/Dark toggle, big readable system font
# - Sidebar sheet picker (search)
# - Plotly charts (date or years-as-headers)
# - Per-sheet numeric summary
# - Chat grounded in the ACTIVE sheet with computed facts; falls back to LLM only if needed

import os, re, json, sys, pkgutil
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ───────────────────────── CONFIG ─────────────────────────
FILE_NAME = "FIN42030 WMT Valuation (2).xlsx"

# Optional Azure OpenAI (add env vars in Streamlit Cloud to enable chat)
# Hard-coded Azure OpenAI (optional; safe fallback if unset)
AZURE_OPENAI_ENDPOINT    = "https://testaisentiment.openai.azure.com/"
AZURE_OPENAI_API_KEY     = "cb1c33772b3c4edab77db69ae18c9a43"
AZURE_OPENAI_API_VERSION = "2024-02-15-preview"
AZURE_OPENAI_DEPLOYMENT  = "aipocexploration"

st.set_page_config(page_title="Walmart Valuation Explorer", page_icon="📊", layout="wide")

# ───────────────────────── THEME / CSS ─────────────────────────
# Sidebar toggle for Light / Dark (Light is the default for readability)
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
  /* cards and headings */
  .card {{ background:var(--panel); border:1px solid var(--border); border-radius:18px; padding:16px; }}
  .headline {{ font-size: 2.1rem; font-weight: 900; letter-spacing:-.01em; margin: 0 0 .25rem 0; }}
  .soft {{ color:var(--muted); }}
  /* KPI blocks */
  .kpi {{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:16px; }}
  .kpi h4 {{ margin:.1rem 0 .4rem 0; font-size:.95rem; color:var(--muted); }}
  .kpi .v {{ font-size: 1.6rem; font-weight: 900; }}
  .kpi .d {{ font-size:.92rem; color:var(--muted); }}
  /* chat message spacing */
  div[data-testid="stChatMessage"] p {{ font-size: 1rem; }}
  /* sidebar spacing */
  section[data-testid="stSidebar"] .css-1d391kg {{ padding-top: 0.5rem; }}
  /* selectboxes: allow wrapping */
  div[data-baseweb="select"] span {{ white-space: normal !important; }}
</style>
""", unsafe_allow_html=True)

# ───────────────────────── LLM (optional) ─────────────────────────
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
            model=AZURE_OPENAI_DEPLOYMENT,
            messages=messages, temperature=temperature, max_tokens=max_tokens
        )
        return r.choices[0].message.content
    except Exception as e:
        return f"(LLM error: {e})"

# ───────────────────────── HELPERS ─────────────────────────
def need_openpyxl() -> bool:
    return pkgutil.find_loader("openpyxl") is None

@st.cache_data(show_spinner=False)
def load_workbook(path: str) -> Dict[str, pd.DataFrame]:
    if need_openpyxl():
        raise ImportError("openpyxl not installed. Add to requirements.txt and redeploy.")
    xl = pd.ExcelFile(path, engine="openpyxl")
    dfs = {}
    for name in xl.sheet_names:
        df = xl.parse(name)
        # tidy headers
        df.columns = [str(c).strip() for c in df.columns]
        # drop totally empty columns
        empties = [c for c in df.columns if df[c].isna().all()]
        if empties: df = df.drop(columns=empties)
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
    if label_col and label_col in long.columns:
        var_rank = long.groupby(label_col)["Value"].var().sort_values(ascending=False)
        long = long[long[label_col].isin(list(var_rank.head(6).index))]
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
    if any(k in low for k in ["income","p&l","profit","is"]): hints.append("**Income Statement**: growth & margins.")
    if "balance" in low or low.endswith(" bs"):             hints.append("**Balance Sheet**: working capital & leverage.")
    if any(k in low for k in ["cash","fcf","free cash","cf"]): hints.append("**Cash Flow**: CFO vs CAPEX → FCF.")
    if any(k in low for k in ["assumption","wacc","market","drivers","valuation"]): hints.append("**Assumptions/WACC** driving valuation.")
    year_cols = find_year_header_cols(df); dcols = date_cols(df)
    axis_hint = "date column" if dcols else ("years in headers" if year_cols else "no time axis")
    return f"- Columns: {', '.join(cols[:12])}{'…' if len(cols)>12 else ''}\n- Chart basis: {axis_hint}\n" + (f"- Notes: {' '.join(hints)}" if hints else "")

def build_sheet_facts(df: pd.DataFrame, name: str) -> dict:
    """Compute quick facts we can cite in chat: year range + latest values + CAGRs."""
    facts = {"sheet": name, "series": []}
    dcols = date_cols(df); ycols = find_year_header_cols(df)
    if dcols:
        t=dcols[0]; x=df.copy(); x[t]=pd.to_datetime(x[t], errors="coerce"); x=x.dropna(subset=[t]).sort_values(t)
        ncols = num_cols(x); x, ncols = coerce_numeric(x, ncols); ncols=ncols[:6]
        for c in ncols:
            s = x[[t,c]].dropna()
            if s.empty: continue
            first, last = s.iloc[0,0], s.iloc[-1,0]
            v0, v1 = float(s.iloc[0,1]), float(s.iloc[-1,1])
            n = max(1, len(s)-1); cagr = (v1/v0)**(1/n)-1 if v0>0 and v1>0 else np.nan
            facts["series"].append({"label":c, "latest":v1, "start":str(first)[:10], "end":str(last)[:10], "cagr":cagr})
    elif ycols:
        long = wide_years_to_long(df, ycols)
        if long is not None and not long.empty:
            label_cols=[c for c in long.columns if c not in ("Year","Value")]
            label = label_cols[0] if label_cols else None
            if label:
                for lbl, seg in long.groupby(label):
                    seg = seg.sort_values("Year")
                    v0, v1 = float(seg["Value"].iloc[0]), float(seg["Value"].iloc[-1])
                    y0, y1 = int(seg["Year"].iloc[0]), int(seg["Year"].iloc[-1])
                    n = max(1, len(seg)-1); cagr = (v1/v0)**(1/n)-1 if v0>0 and v1>0 else np.nan
                    facts["series"].append({"label":str(lbl), "latest":v1, "start":y0, "end":y1, "cagr":cagr})
            else:
                # single series case
                seg = long.sort_values("Year")
                v0, v1 = float(seg["Value"].iloc[0]), float(seg["Value"].iloc[-1])
                y0, y1 = int(seg["Year"].iloc[0]), int(seg["Year"].iloc[-1])
                n = max(1, len(seg)-1); cagr = (v1/v0)**(1/n)-1 if v0>0 and v1>0 else np.nan
                facts["series"].append({"label":"Series", "latest":v1, "start":y0, "end":y1, "cagr":cagr})
    return facts

def facts_to_text(facts: dict) -> str:
    if not facts or not facts.get("series"): return "(no computed series)"
    lines=[f"Sheet: {facts['sheet']} — computed time series:"]
    for s in facts["series"]:
        cagr = f"{s['cagr']*100:.1f}%" if s["cagr"]==s["cagr"] else "n/a"
        lines.append(f"- {s['label']}: latest={s['latest']:,.0f} (from {s['start']} to {s['end']}), CAGR {cagr}")
    return "\n".join(lines)

def build_chat_context(df: pd.DataFrame, name: str) -> str:
    """Context string for the LLM: preview + computed facts."""
    prev = safe_preview(df, n=8)
    facts = facts_to_text(build_sheet_facts(df, name))
    return f"{facts}\n\nPreview (first 8 rows):\n{prev}"

# ───────────────────────── LOAD WORKBOOK ─────────────────────────
st.markdown(f'<div class="headline">Walmart Valuation Explorer</div><div class="soft">Python: {sys.executable}</div>', unsafe_allow_html=True)

if not os.path.exists(FILE_NAME):
    st.error(f"File not found: {FILE_NAME} (place it next to app.py).")
    st.stop()

try: dfs = load_workbook(FILE_NAME)
except Exception as e:
    st.error(str(e)); st.stop()

# ───────────────────────── SIDEBAR: Sheets ─────────────────────────
st.sidebar.header("Sheets")
flt = st.sidebar.text_input("Filter sheets", "")
sheet_names = [n for n in dfs.keys() if flt.lower() in n.lower()] or list(dfs.keys())
selected = st.sidebar.selectbox("Select a sheet", sheet_names, index=0)

# ───────────────────────── OVERVIEW KPIs (simple heuristic) ─────────────────────────
st.markdown("## Overview")
k1,k2,k3,k4 = st.columns(4)
for k in (k1,k2,k3,k4):
    with k: st.markdown('<div class="kpi"><h4>—</h4><div class="v">—</div></div>', unsafe_allow_html=True)

# ───────────────────────── SELECTED SHEET VIEW ─────────────────────────
st.markdown(f"## {selected}")
df = dfs[selected].copy()

# remove fully-empty "Unnamed" columns
unnamed = [c for c in df.columns if c.lower().startswith("unnamed") and df[c].isna().all()]
if unnamed: df = df.drop(columns=unnamed)

with st.expander("Preview", expanded=False):
    st.dataframe(df, use_container_width=True, height=420)

with st.expander("What’s in this sheet?", expanded=True):
    st.markdown(sheet_summary_text(df, selected))

# Charts
st.markdown("### Charts")
dcols = date_cols(df); ncols = num_cols(df); ycols = find_year_header_cols(df)

if dcols:
    x = dcols[0]
    dfx = df.copy(); dfx[x] = pd.to_datetime(dfx[x], errors="coerce"); dfx = dfx.dropna(subset=[x]).sort_values(x)
    dfx, ncols = coerce_numeric(dfx, ncols); targets = ncols[:3]
    if targets:
        choose = st.multiselect("Series", options=ncols, default=targets)
        fig = px.line(dfx, x=x, y=choose)
        fig.update_layout(height=380, template="plotly_white" if st.session_state.ui_theme=="Light" else "plotly_dark",
                          margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No numeric columns to plot against the date column.")
elif ycols:
    long = wide_years_to_long(df, ycols)
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

# Category bars
st.markdown("### Category summary")
cats=[]
for c in df.columns:
    if pd.api.types.is_numeric_dtype(df[c]): continue
    try:
        u = df[c].nunique(dropna=True)
        if 1 < u <= 20: cats.append(c)
    except Exception: pass
if cats:
    cat = st.selectbox("Group by", options=cats)
    co = df.copy(); agg_cols=[]
    for c in df.columns:
        if c==cat: continue
        co[c] = pd.to_numeric(co[c], errors="coerce")
        if co[c].notna().any(): agg_cols.append(c)
    if agg_cols:
        pick = st.multiselect("Aggregate columns", options=agg_cols, default=agg_cols[:1])
        if pick:
            agg = co.groupby(cat)[pick].sum(numeric_only=True).sort_values(pick[0], ascending=False).head(30).reset_index()
            fig = px.bar(agg, x=cat, y=pick, barmode="group")
            fig.update_layout(height=360, template="plotly_white" if st.session_state.ui_theme=="Light" else "plotly_dark",
                              margin=dict(l=10,r=10,t=10,b=10), xaxis_tickangle=-25)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.caption("Pick at least one numeric column.")
    else:
        st.caption("No numeric columns to aggregate by category.")
else:
    st.caption("No small-cardinality category column detected.")

# Histograms
st.markdown("### Numeric histograms")
co = df.copy(); num_cands=[]
for c in df.columns:
    co[c] = pd.to_numeric(co[c], errors="coerce")
    if co[c].notna().any(): num_cands.append(c)
if num_cands:
    sel = st.multiselect("Columns", options=num_cands, default=num_cands[:2])
    for c in sel[:3]:
        fig = px.histogram(co, x=c, nbins=40)
        fig.update_layout(height=300, template="plotly_white" if st.session_state.ui_theme=="Light" else "plotly_dark",
                          margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig, use_container_width=True)
else:
    st.caption("No numeric columns for histograms.")

st.divider()

# ───────────────────────── ANALYST CHAT (sheet-grounded) ─────────────────────────
st.markdown("## Analyst Chat")
st.caption("I’ll use the **selected sheet** first, compute facts (latest + CAGR), then answer. If the workbook lacks it, I’ll say so.")

if "chat" not in st.session_state:
    st.session_state.chat=[{"role":"assistant","content":"Hi! Ask about revenue growth, margins, FCF, WACC, or ‘summarise this sheet’."}]

# render history
for m in st.session_state.chat:
    with st.chat_message("assistant" if m["role"]=="assistant" else "user"):
        st.write(m["content"])

SYSTEM = """You are a valuation analyst. First rely on the provided SHEET FACTS; answer quantitatively and concisely.
If you use workbook numbers, **cite the sheet name** exactly once in the answer.
If the facts don’t contain the requested number, say what’s missing briefly. Avoid generic filler.
"""

def answer(user_q: str, df: pd.DataFrame, sheet_name: str) -> str:
    # 1) try deterministic “summarise / key takeaways”
    qlow = (user_q or "").lower()
    facts = build_sheet_facts(df, sheet_name)
    if any(k in qlow for k in ["summarise","summarize","summary","key takeaway","what is the analysis"]):
        if facts and facts.get("series"):
            txt = facts_to_text(facts)
            return txt + f"\n\n(Source: **{sheet_name}**)"
        # fallback to simple column summary
        return sheet_summary_text(df, sheet_name) + f"\n\n(Source: **{sheet_name}**)"

    # 2) build context and ask LLM (if configured)
    context = build_chat_context(df, sheet_name)
    messages=[
        {"role":"system","content":SYSTEM},
        {"role":"user","content": f"SHEET FACTS:\n{context}\n\nQuestion: {user_q}"}
    ]
    out = ask_gpt(messages)
    if out.startswith("("):   # chat disabled
        return "Chat is disabled (no Azure OpenAI env vars). Use the charts above."
    # ensure sheet citation present once
    if f"**{sheet_name}**" not in out:
        out += f"\n\n(Source: **{sheet_name}**)"
    return out

prompt = st.chat_input("Ask about the selected sheet…")
if prompt:
    st.session_state.chat.append({"role":"user","content":prompt})
    reply = answer(prompt, df, selected)
    st.session_state.chat.append({"role":"assistant","content":reply})
    st.rerun()
