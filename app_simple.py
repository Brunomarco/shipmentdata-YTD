
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
import re

st.set_page_config(page_title="Executive Shipment Analytics", page_icon="🎯", layout="wide")

# ---------- Helpers ----------
@st.cache_data(show_spinner=False)
def read_excel(file):
    try:
        return pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        return pd.DataFrame()

def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def find_col(df, names):
    if df.empty: 
        return None
    targets = [norm(n) for n in names]
    cols = {c: norm(c) for c in df.columns}
    # exact normalized match
    for c, cn in cols.items():
        if cn in targets:
            return c
    # substring match
    for c, cn in cols.items():
        if any(t in cn for t in targets):
            return c
    return None

def to_num(s):
    if pd.api.types.is_numeric_dtype(s): 
        return s
    return pd.to_numeric(
        s.astype(str).str.replace("€","").str.replace(",","").str.strip(),
        errors="coerce"
    )

def to_dt(s):
    return pd.to_datetime(s, errors="coerce")

def compute_otp(df, actual_col, promised_col, qc_col, non_ctrl_keywords):
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
        is_non_ctrl = qc_norm.apply(lambda x: any(k in x for k in non_ctrl_keywords))
        net = (d["on_time"] | (~d["on_time"] & is_non_ctrl)).mean()
    else:
        net = gross
    return float(gross), float(net), int(len(d)), d[~d["on_time"]]

# ---------- UI ----------
st.title("🎯 Executive Shipment Analytics (Simple)")
st.caption("Upload your Excel, auto-filter to STATUS = 440-BILLED, and get core KPIs.")

uploaded = st.file_uploader("Upload shipment Excel (.xlsx)", type=["xlsx"])

if not uploaded:
    st.info("Please upload an Excel to begin.")
    st.stop()

df0 = read_excel(uploaded)
if df0.empty:
    st.stop()

# Normalize basic fields
STATUS = find_col(df0, ["status","referstatusord","statusord"])
CHARGES = find_col(df0, ["total charges","charges","amount","total_charges"])
SVC = find_col(df0, ["svc"])
DEP = find_col(df0, ["dep","routedep","origin"])
DEL_CTRY = find_col(df0, ["del ctry","delivery country","dest country","dest ctry"])
WEIGHT = find_col(df0, ["weight(kg)","weight kg","billable weight kg"])
DEPART_DT = find_col(df0, ["depart date / time","depart date","pickup date/time","pu date/time"])
POD = find_col(df0, ["pod date/time","pod","arrive date / time","arrive date"])
UPD_DEL = find_col(df0, ["upd del","qdtupd del","qdt upd del","qdt","promised del"])
QC = find_col(df0, ["qc name","qcname","delay reason","quality reason","qccode"])

# Clean
df = df0.copy()
for c in df.columns:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.strip()

# Filter billed
if STATUS:
    df = df[df[STATUS].str.upper() == "440-BILLED"]

# Numeric & dates
if CHARGES: df[CHARGES] = to_num(df[CHARGES])
if WEIGHT: df[WEIGHT] = to_num(df[WEIGHT])
if DEPART_DT: df[DEPART_DT] = to_dt(df[DEPART_DT])

# Sidebar filters
st.sidebar.header("Filters")
if DEPART_DT and df[DEPART_DT].notna().any():
    min_d, max_d = df[DEPART_DT].min(), df[DEPART_DT].max()
    d1, d2 = st.sidebar.date_input("Depart date range", (min_d.date(), max_d.date()))
    df = df[(df[DEPART_DT] >= pd.to_datetime(d1)) & (df[DEPART_DT] <= pd.to_datetime(d2) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))]

if SVC:
    svcs = sorted(df[SVC].dropna().unique().tolist())
    sel = st.sidebar.multiselect("SVC", svcs, default=svcs[:10] if len(svcs)>10 else svcs)
    if sel:
        df = df[df[SVC].isin(sel)]

# OTP
st.sidebar.header("OTP Settings")
default_non_ctrl = ["customs","warehouse","w/house","airline","weather","security","strike","customer","consignee","agent"]
non_ctrl_selected = st.sidebar.multiselect(
    "Mark these reasons as NON‑controllable (count late as on‑time for Net OTP)",
    options=sorted(set([str(x) for x in (df[QC].dropna().unique() if QC else [])])),
    default=[r for r in (df[QC].dropna().unique() if QC else []) if any(k in str(r).lower() for k in default_non_ctrl)]
) if QC else []

keywords = [s.lower() for s in non_ctrl_selected] if non_ctrl_selected else default_non_ctrl
otp_gross, otp_net, otp_base, late_df = compute_otp(df, POD, UPD_DEL, QC, keywords)

# KPIs
total_ship = len(df)
total_cost = df[CHARGES].sum() if CHARGES else np.nan
avg_cost = df[CHARGES].mean() if CHARGES else np.nan
total_weight = df[WEIGHT].sum() if WEIGHT else np.nan
countries = df[DEL_CTRY].nunique() if DEL_CTRY else 0
qc_rate = df[QC].notna().mean() if QC else np.nan

k1,k2,k3,k4 = st.columns(4)
k1.metric("Total Shipments", f"{total_ship:,}")
k2.metric("Total Cost (EUR)", f"€{total_cost:,.0f}" if pd.notna(total_cost) else "—")
k3.metric("Avg Cost / Shipment", f"€{avg_cost:,.0f}" if pd.notna(avg_cost) else "—")
k4.metric("Total Weight (KG)", f"{total_weight:,.0f}" if pd.notna(total_weight) else "—")

k5,k6,k7,k8 = st.columns(4)
k5.metric("OTP Gross", f"{otp_gross*100:.1f}%" if pd.notna(otp_gross) else "—", help=f"Base: {otp_base} shipments")
k6.metric("OTP Net", f"{otp_net*100:.1f}%" if pd.notna(otp_net) else "—", help="Non‑controllable late counted on‑time")
k7.metric("QC Issue Rate", f"{qc_rate*100:.1f}%" if pd.notna(qc_rate) else "—")
k8.metric("Countries Served", f"{countries}" if countries else "—")

st.divider()

# Charts (lean set)
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
        st.plotly_chart(px.bar(monthly, x="_month", y="Shipments"), use_container_width=True)
    with c2:
        st.subheader("Monthly Cost (EUR)")
        st.plotly_chart(px.bar(monthly, x="_month", y="Cost"), use_container_width=True)

if SVC:
    st.subheader("Service Mix (SVC)")
    svc_df = df.groupby(SVC).size().reset_index(name="Shipments").sort_values("Shipments", ascending=False)
    st.plotly_chart(px.bar(svc_df, x=SVC, y="Shipments"), use_container_width=True)

if DEP:
    st.subheader("Distribution by DEP")
    dep_df = df.groupby(DEP).size().reset_index(name="Shipments").sort_values("Shipments", ascending=False).head(25)
    st.plotly_chart(px.bar(dep_df, x=DEP, y="Shipments"), use_container_width=True)

if CHARGES and WEIGHT:
    st.subheader("Cost vs Weight (KG)")
    st.plotly_chart(px.scatter(df, x=WEIGHT, y=CHARGES, hover_data=[SVC,DEP,DEL_CTRY] if all([SVC,DEP,DEL_CTRY]) else None), use_container_width=True)

if QC:
    st.subheader("Top QC Reasons")
    qc_counts = df[QC].value_counts().head(20).reset_index()
    qc_counts.columns = ["QC NAME","Shipments"]
    st.plotly_chart(px.bar(qc_counts, x="QC NAME", y="Shipments"), use_container_width=True)

# Late shipments table
if not isinstance(late_df, float) and not late_df.empty:
    st.subheader("Late Shipments")
    show_cols = [c for c in [SVC, DEP, DEL_CTRY, CHARGES, WEIGHT, POD, UPD_DEL, QC] if c]
    st.dataframe(df.loc[late_df.index, show_cols], use_container_width=True)
