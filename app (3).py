
# app.py
# Executive Shipment Analytics Dashboard (Upload-based)
# - Filters to STATUS = 440-BILLED
# - KPIs + SVC counts + DEP distribution + monthly trends + lanes + QC + Cost vs Weight
# - OTP Gross vs Net with controllable/non‑controllable selector
# - All costs displayed as EUR (assumes the Excel is already in EUR)

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re

st.set_page_config(page_title="Executive Shipment Analytics", page_icon="🎯", layout="wide")

# ---------- Styling (lightweight, professional) ----------
st.markdown("""
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    h1, h2, h3 { color: #0f172a; }
    div[data-testid="metric-container"] {
        background: #fff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.9rem;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }
    </style>
""", unsafe_allow_html=True)

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def read_excel(file: BytesIO) -> pd.DataFrame:
    try:
        return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        return pd.DataFrame()

def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def find_col(df: pd.DataFrame, names: list[str]) -> str | None:
    """Find a column by exact normalized match, else by substring match."""
    if df.empty:
        return None
    targets = [norm(n) for n in names]
    cmap = {c: norm(c) for c in df.columns}
    # exact match
    for c, cn in cmap.items():
        if cn in targets:
            return c
    # substring match (handles concatenated headers like 'PICKUP DATE/TIMEAirlineROUTEDEP')
    for c, cn in cmap.items():
        if any(t in cn for t in targets):
            return c
    return None

def to_num(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return s
    return pd.to_numeric(
        s.astype(str).str.replace("€","").str.replace(",","").str.strip(),
        errors="coerce"
    )

def to_dt(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def eur(x) -> str:
    try:
        return f"€{x:,.0f}"
    except Exception:
        return "€0"

def compute_otp(df: pd.DataFrame, actual_col: str | None, promised_col: str | None,
                qc_col: str | None, controllable_keywords: list[str], mode: str = "keywords"):
    """
    Returns:
      gross, net, base_n, late_df
    Definitions:
      Gross OTP = % with actual <= promised
      Net OTP   = treats late due to NON‑controllable as on‑time (i.e., excludes non‑controllable late)
    Here we label NON‑controllable by *not* matching controllable keywords.
    """
    if actual_col is None or promised_col is None:
        return np.nan, np.nan, 0, pd.DataFrame()
    d = df[[actual_col, promised_col] + ([qc_col] if qc_col else [])].copy()
    d["actual"] = to_dt(d[actual_col])
    d["promised"] = to_dt(d[promised_col])
    d = d[d["actual"].notna() & d["promised"].notna()]
    if d.empty:
        return np.nan, np.nan, 0, d
    d["on_time"] = d["actual"] <= d["promised"]
    gross = d["on_time"].mean()

    if qc_col:
        qc_norm = d[qc_col].astype(str).str.lower()
        # mark controllable if any keyword appears
        is_controllable = qc_norm.apply(lambda x: any(k in x for k in controllable_keywords))
        # Non‑controllable if not controllable
        is_non_ctrl = ~is_controllable
        # Net: count late due to NON‑controllable as on‑time
        net = (d["on_time"] | (~d["on_time"] & is_non_ctrl)).mean()
    else:
        net = gross

    late_df = d[~d["on_time"]].copy()
    return float(gross), float(net), int(len(d)), late_df

# ---------- UI ----------
st.title("🎯 Executive Shipment Analytics")
st.caption("Upload your Excel, auto‑filter to **STATUS = 440‑BILLED**, and explore fact‑based KPIs & charts. All amounts shown in EUR.")

uploaded = st.file_uploader("Upload shipment Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Please upload an Excel to begin. Expected headers include: STATUS, TOTAL CHARGES, SVC, DEP, QDT / UPD DEL, POD DATE/TIME, QC NAME, etc.")
    st.stop()

df0 = read_excel(uploaded)
if df0.empty:
    st.stop()

# ---------- Column mapping (robust to concatenations) ----------
STATUS = find_col(df0, ["status","referstatusord","statusord"])
CHARGES = find_col(df0, ["total charges","charges","amount","total_charges"])
SVC = find_col(df0, ["svc"])
SVCDESC = find_col(df0, ["svcdesc","svc desc","service description","service desc"])
DEP = find_col(df0, ["dep","routedep","origin","origin station"])
DEL_CTRY = find_col(df0, ["del ctry","delivery country","dest country","dest ctry"])
WEIGHT = find_col(df0, ["weight(kg)","weight kg","billable weight kg"])
DEPART_DT = find_col(df0, ["depart date / time","depart date","pickup date/time","pu date/time","pupickup date/time"])
POD = find_col(df0, ["pod date/time","pod","arrive date / time","arrive date","arr"])
UPD_DEL = find_col(df0, ["upd del","qdtupd del","qdt upd del","qdt","promised del"])
QC = find_col(df0, ["qc name","qcname","delay reason","quality reason","qccode","reason code"])
REFER = find_col(df0, ["refer","ord","invoice","numord"])

# Clean basics
df = df0.copy()
for c in df.columns:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.strip()

# Filter billed
if STATUS:
    df = df[df[STATUS].str.upper() == "440-BILLED"]

# Numerics & dates
if CHARGES: df[CHARGES] = to_num(df[CHARGES])
if WEIGHT: df[WEIGHT] = to_num(df[WEIGHT])
if DEPART_DT: df[DEPART_DT] = to_dt(df[DEPART_DT])
if POD: df[POD] = to_dt(df[POD])
if UPD_DEL: df[UPD_DEL] = to_dt(df[UPD_DEL])

# Sidebar filters
st.sidebar.header("Filters")
if DEPART_DT and df[DEPART_DT].notna().any():
    min_d, max_d = df[DEPART_DT].min(), df[DEPART_DT].max()
    d1, d2 = st.sidebar.date_input("Depart date range",
                                   (min_d.date(), max_d.date()))
    df = df[(df[DEPART_DT] >= pd.to_datetime(d1)) &
            (df[DEPART_DT] <= pd.to_datetime(d2) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]

if SVC and df[SVC].notna().any():
    svcs = sorted(df[SVC].dropna().unique().tolist())
    sel = st.sidebar.multiselect("SVC (Service Type)", svcs, default=svcs)
    if sel:
        df = df[df[SVC].isin(sel)]

# OTP settings
st.sidebar.header("OTP Settings")
# Per your note: treat 'customs' and 'warehouse' as CONTROLLABLE by default
default_controllable = ["customs","warehouse","w/house","mnx","order entry","late dispatch","data entry","hub","pu agt","del agt"]
controllable_selected = st.sidebar.multiselect(
    "Mark these reasons as CONTROLLABLE (late due to these will be counted as late in Net OTP)",
    options=sorted(set([str(x) for x in (df[QC].dropna().unique() if QC else [])])),
    default=[r for r in (df[QC].dropna().unique() if QC else [])
             if any(k in str(r).lower() for k in default_controllable)]
) if QC else []

ctrl_keywords = [s.lower() for s in controllable_selected] if controllable_selected else default_controllable

# Compute OTP
otp_gross, otp_net, otp_base, late_df = compute_otp(df, POD, UPD_DEL, QC, ctrl_keywords)

# ---------- Executive KPIs ----------
total_ship = len(df)
total_cost = df[CHARGES].sum() if CHARGES else np.nan
avg_cost = df[CHARGES].mean() if CHARGES else np.nan
total_weight = df[WEIGHT].sum() if WEIGHT else np.nan
countries = df[DEL_CTRY].nunique() if DEL_CTRY else 0
qc_rate = df[QC].notna().mean() if QC else np.nan

k1,k2,k3,k4 = st.columns(4)
k1.metric("Total Shipments", f"{total_ship:,}")
k2.metric("Total Cost (EUR)", eur(total_cost) if pd.notna(total_cost) else "—")
k3.metric("Avg Cost / Shipment", eur(avg_cost) if pd.notna(avg_cost) else "—")
k4.metric("Total Weight (KG)", f"{total_weight:,.0f}" if pd.notna(total_weight) else "—")

k5,k6,k7,k8 = st.columns(4)
k5.metric("OTP Gross", f"{otp_gross*100:.1f}%" if pd.notna(otp_gross) else "—", help=f"Base: {otp_base} shipments")
k6.metric("OTP Net", f"{otp_net*100:.1f}%" if pd.notna(otp_net) else "—",
          help="Late due to NON‑controllable reasons are treated as on‑time")
k7.metric("QC Issue Rate", f"{qc_rate*100:.1f}%" if pd.notna(qc_rate) else "—",
          help="Share of billed shipments with any QC reason")
k8.metric("Countries Served", f"{countries}" if countries else "—")

st.markdown("---")

# ---------- Charts (insightful & concise) ----------
# 1) Monthly trends
if DEPART_DT:
    tmp = df.copy()
    tmp["_month"] = tmp[DEPART_DT].dt.to_period("M").dt.to_timestamp()
    monthly = tmp.groupby("_month").agg(
        Shipments=("_month","size"),
        Cost=(CHARGES,"sum") if CHARGES else ("_month","size")
    ).reset_index()

    c1,c2 = st.columns(2)
    with c1:
        st.subheader("Monthly Shipments")
        st.plotly_chart(px.bar(monthly, x="_month", y="Shipments",
                               labels={"_month":"Month"}), use_container_width=True)
    with c2:
        st.subheader("Monthly Cost (EUR)")
        st.plotly_chart(px.bar(monthly, x="_month", y="Cost",
                               labels={"_month":"Month"}), use_container_width=True)

# 2) SVC distribution
if SVC:
    st.subheader("Service Mix (SVC)")
    svc_df = df.groupby(SVC).agg(Shipments=(SVC,"size"),
                                 Cost=(CHARGES,"sum") if CHARGES else (SVC,"size")).reset_index()
    svc_df = svc_df.sort_values("Shipments", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.bar(svc_df, x=SVC, y="Shipments"), use_container_width=True)
    with c2:
        if CHARGES:
            st.plotly_chart(px.bar(svc_df, x=SVC, y="Cost", labels={"Cost":"EUR"}), use_container_width=True)

# 3) DEP distribution
if DEP:
    st.subheader("Distribution by DEP (Origin Station)")
    dep_df = df.groupby(DEP).agg(Shipments=(DEP,"size"),
                                 Cost=(CHARGES,"sum") if CHARGES else (DEP,"size")).reset_index()
    dep_df = dep_df.sort_values("Shipments", ascending=False).head(25)
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.bar(dep_df, x=DEP, y="Shipments"), use_container_width=True)
    with c2:
        if CHARGES:
            st.plotly_chart(px.bar(dep_df, x=DEP, y="Cost", labels={"Cost":"EUR"}), use_container_width=True)

# 4) Top lanes (DEP → Delivery Country)
if DEP and DEL_CTRY:
    st.subheader("Top Lanes (DEP → Delivery Country)")
    lanes = df.copy()
    lanes["LANE"] = lanes[DEP].fillna("").astype(str).str.strip() + " → " + lanes[DEL_CTRY].fillna("").astype(str).str.strip()
    lane_df = lanes.groupby("LANE").agg(Shipments=("LANE","size"),
                                        Cost=(CHARGES,"sum") if CHARGES else ("LANE","size")).reset_index()
    lane_df = lane_df.sort_values(["Shipments","Cost"], ascending=[False, False]).head(25)
    st.plotly_chart(px.bar(lane_df, x="LANE", y="Shipments"), use_container_width=True)

# 5) Cost vs Weight
if CHARGES and WEIGHT:
    st.subheader("Cost vs Weight (KG)")
    st.plotly_chart(px.scatter(df, x=WEIGHT, y=CHARGES,
                               hover_data=[SVC, DEP, DEL_CTRY] if all([SVC,DEP,DEL_CTRY]) else None,
                               labels={WEIGHT:"Weight (KG)", CHARGES:"Cost (EUR)"}),
                    use_container_width=True)

# 6) QC Reasons
if QC:
    st.subheader("QC Reasons (Billed Shipments)")
    qc_counts = df[QC].value_counts(dropna=True).head(25).reset_index()
    qc_counts.columns = ["QC NAME", "Shipments"]
    st.plotly_chart(px.bar(qc_counts, x="QC NAME", y="Shipments"), use_container_width=True)

# 7) Late shipments table
if isinstance(late_df, pd.DataFrame) and not late_df.empty:
    st.subheader("Late Shipments (for actioning)")
    idx = late_df.index
    show_cols = [c for c in [REFER, SVC, DEP, DEL_CTRY, CHARGES, WEIGHT, UPD_DEL, POD, QC] if c]
    st.dataframe(df.loc[idx, show_cols], use_container_width=True)

# ---------- Education / Definitions ----------
with st.expander("ℹ️ OTP Definitions (Gross vs Net)"):
    st.markdown("""
**OTP Gross** = percentage of shipments where **Actual Delivery** ≤ **Promised Delivery** (e.g., `POD DATE/TIME` vs `UPD DEL`).  
**OTP Net** = adjusts OTP by treating **late shipments due to NON‑controllable reasons** as **on‑time** (e.g., weather, airline, strikes).  
In the sidebar, you select which reasons are **CONTROLLABLE** (e.g., customs, warehouse, order entry). Late due to these remain **late** in Net OTP.
""")
