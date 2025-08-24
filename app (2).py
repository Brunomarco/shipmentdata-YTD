
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
from io import BytesIO
import re
from pathlib import Path

# -------------------------------
# Page configuration
# -------------------------------
st.set_page_config(
    page_title="Shipment Cost & OTP Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------
# Helpers
# -------------------------------
@st.cache_data(show_spinner=False)
def load_excel_from_path(path: str | Path) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        st.error(f"Excel file not found at: {p}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(p, engine="openpyxl")
        return df
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        return pd.DataFrame()

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    if df.empty:
        return None
    norm_map = {c: re.sub(r'[^a-z0-9]', '', str(c).lower()) for c in df.columns}
    for cand in candidates:
        cand_norm = re.sub(r'[^a-z0-9]', '', cand.lower())
        # exact normalized match
        for col, col_norm in norm_map.items():
            if cand_norm == col_norm:
                return col
    # fallback: substring match (lets us find STATUS inside REFERSTATUSORD, etc.)
    for cand in candidates:
        cand_norm = re.sub(r'[^a-z0-9]', '', cand.lower())
        for col, col_norm in norm_map.items():
            if cand_norm in col_norm:
                return col
    return None

def parse_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def eur(x: float) -> str:
    try:
        return f"€{x:,.2f}"
    except Exception:
        return "€0.00"

def build_lane(row, dep_col, del_ctry_col):
    dep = str(row.get(dep_col, "") or "").strip()
    dest = str(row.get(del_ctry_col, "") or "").strip()
    if dep or dest:
        return f"{dep}->{dest}"
    return None

# -------------------------------
# Data source (fixed path)
# -------------------------------
st.sidebar.header("📥 Data source")
default_path = Path("shipment data YTD 25.xlsx")
st.sidebar.write(f"Using Excel: **{default_path}**")
df_raw = load_excel_from_path(default_path)

if df_raw.empty:
    st.stop()

# -------------------------------
# Column detection
# -------------------------------
STATUS_COL = find_col(df_raw, ["STATUS"])
CHARGES_COL = find_col(df_raw, ["TOTAL CHARGES", "AMOUNT", "TOTAL_CHARGES"])
SVC_COL = find_col(df_raw, ["SVC"])
SVCDESC_COL = find_col(df_raw, ["SVCDESC", "SVC DESC", "SERVICE DESC"])
DEP_COL = find_col(df_raw, ["DEP", "ROUTEDEP", "ORIGIN", "ORIGIN STATION"])
DEL_CTRY_COL = find_col(df_raw, ["DEL CTRY", "DELIVERY COUNTRY", "DEST COUNTRY", "DEST CTRY"])
WEIGHT_KG_COL = find_col(df_raw, ["WEIGHT(KG)", "WEIGHT KG", "WT KG", "Billable Weight KG"])
ARRIVE_COL = find_col(df_raw, ["POD DATE/TIME", "Arrive Date / Time", "ARRIVE DATE/TIME", "DELIVERED AT"])
PROMISED_COL = find_col(df_raw, ["UPD DEL", "QDTUPD DEL", "QDT UPD DEL", "QDT", "PROMISED DEL"])
DEPART_COL = find_col(df_raw, ["Depart Date / Time", "DEPART DATE/TIME", "DEPART DATE", "PUPICKUP DATE/TIME"])
QCNAME_COL = find_col(df_raw, ["QC NAME", "QUALITY REASON", "DELAY REASON", "QCCODE"])

missing = [n for n,v in {
    "STATUS": STATUS_COL,
    "TOTAL CHARGES": CHARGES_COL,
    "SVC": SVC_COL,
    "DEP": DEP_COL,
    "DEL CTRY": DEL_CTRY_COL,
    "WEIGHT(KG)": WEIGHT_KG_COL,
    "ACTUAL DELIVERY": ARRIVE_COL,
    "PROMISED DELIVERY": PROMISED_COL,
}.items() if v is None]

if missing:
    st.warning("Some expected columns were not found. The app will still run, but certain KPIs may be incomplete:\n\n- " + "\n- ".join(missing))

# -------------------------------
# Clean + Filter billed
# -------------------------------
df = df_raw.copy()

# Trim whitespace everywhere
df.columns = [c.strip() for c in df.columns]
for c in df.columns:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.strip()

# Only billed
if STATUS_COL:
    df = df[df[STATUS_COL].str.upper() == "440-BILLED"]

# Coerce numerics
if CHARGES_COL and df[CHARGES_COL].dtype == object:
    df[CHARGES_COL] = pd.to_numeric(df[CHARGES_COL].str.replace(",", "").str.replace("€",""), errors="coerce")

if WEIGHT_KG_COL and df[WEIGHT_KG_COL].dtype == object:
    df[WEIGHT_KG_COL] = pd.to_numeric(df[WEIGHT_KG_COL].str.replace(",", ""), errors="coerce")

# Dates
for c in [ARRIVE_COL, PROMISED_COL, DEPART_COL]:
    if c:
        df[c] = parse_datetime(df[c])

# Build helper fields
if DEP_COL and DEL_CTRY_COL:
    df["LANE"] = df.apply(lambda r: build_lane(r, DEP_COL, DEL_CTRY_COL), axis=1)
else:
    df["LANE"] = None

# Date filters
min_date = pd.to_datetime(df[DEPART_COL].min()) if DEPART_COL else None
max_date = pd.to_datetime(df[DEPART_COL].max()) if DEPART_COL else None

st.sidebar.header("🧭 Filters")
if DEPART_COL and pd.notna(min_date) and pd.notna(max_date):
    date_range = st.sidebar.date_input(
        "Depart Date range",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date()
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        df = df[(df[DEPART_COL] >= start) & (df[DEPART_COL] <= end)]

# -------------------------------
# OTP Logic
# -------------------------------
st.sidebar.header("⏱ OTP Settings")

default_non_ctrl = [
    "Customer", "Consignee", "Customs", "Airline", "Weather", "Security", "Strike", "Agent", "Warehouse"
]

if QCNAME_COL:
    unique_qc = sorted({str(x) for x in df[QCNAME_COL].dropna().unique()})
    with st.sidebar.expander("Select non-controllable reasons for Net OTP", expanded=False):
        preselect = [q for q in unique_qc if any(q.lower().startswith(k.lower()) for k in default_non_ctrl)]
        non_ctrl_selected = st.multiselect(
            "Reasons considered NON-controllable (late due to these will count as on-time in Net OTP)",
            options=unique_qc,
            default=preselect
        )
else:
    non_ctrl_selected = []

def compute_otp(df: pd.DataFrame) -> tuple[float,float,int]:
    if ARRIVE_COL is None or PROMISED_COL is None:
        return (np.nan, np.nan, 0)
    df2 = df.copy()
    df2["actual_delivery"] = df2[ARRIVE_COL]
    df2["promised_delivery"] = df2[PROMISED_COL]
    mask = df2["actual_delivery"].notna() & df2["promised_delivery"].notna()
    df2 = df2[mask]
    if df2.empty:
        return (np.nan, np.nan, 0)
    df2["on_time"] = df2["actual_delivery"] <= df2["promised_delivery"]
    gross = float(df2["on_time"].mean())

    # Net OTP: if late but reason is marked NON-controllable, treat as on-time
    if QCNAME_COL:
        late_nonctrl = (~df2["on_time"]) & (df2[QCNAME_COL].isin(non_ctrl_selected))
        net = float((df2["on_time"] | late_nonctrl).mean())
    else:
        net = gross
    return (gross, net, len(df2))

otp_gross, otp_net, otp_base_n = compute_otp(df)

# -------------------------------
# KPIs (Executive Summary)
# -------------------------------
total_shipments = int(len(df))
total_cost = float(df[CHARGES_COL].sum()) if CHARGES_COL else np.nan
avg_cost = float(df[CHARGES_COL].mean()) if CHARGES_COL else np.nan
total_weight = float(df[WEIGHT_KG_COL].sum()) if WEIGHT_KG_COL else np.nan
countries_served = int(df[DEL_CTRY_COL].nunique()) if DEL_CTRY_COL else 0
qc_issue_rate = float(df[QCNAME_COL].notna().mean()) if QCNAME_COL else np.nan

st.title("📦 Shipment Cost & OTP Dashboard")

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Total Shipments", f"{total_shipments:,}")
kpi2.metric("Total Cost (EUR)", eur(total_cost) if not np.isnan(total_cost) else "—")
kpi3.metric("Average Cost / Shipment (EUR)", eur(avg_cost) if not np.isnan(avg_cost) else "—")
kpi4.metric("Total Weight (KG)", f"{total_weight:,.0f}" if not np.isnan(total_weight) else "—")

kpi5, kpi6, kpi7, kpi8 = st.columns(4)
kpi5.metric("OTP Gross", f"{(otp_gross*100):.1f}%" if not np.isnan(otp_gross) else "—", help=f"Base: {otp_base_n} shipments")
kpi6.metric("OTP Net", f"{(otp_net*100):.1f}%" if not np.isnan(otp_net) else "—", help=f"Non-controllable reasons treated as on-time")
kpi7.metric("QC Issue Rate", f"{(qc_issue_rate*100):.1f}%" if not np.isnan(qc_issue_rate) else "—", help="Share of billed shipments with a QC reason")
kpi8.metric("Countries Served", f"{countries_served}" if countries_served else "—")

st.markdown("---")

# -------------------------------
# Charts
# -------------------------------
# 1) Monthly shipment & cost trends
if DEPART_COL:
    df["_month"] = df[DEPART_COL].dt.to_period("M").dt.to_timestamp()
    monthly = df.groupby("_month").agg(
        shipments=("index","count") if "index" in df.columns else ("_month","size"),
        cost=(CHARGES_COL, "sum") if CHARGES_COL else (CHARGES_COL, "size")
    ).reset_index()
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("📊 Monthly Shipments")
        fig = px.bar(monthly, x="_month", y="shipments", labels={"_month":"Month","shipments":"Shipments"})
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.subheader("💶 Monthly Cost (EUR)")
        fig = px.bar(monthly, x="_month", y="cost", labels={"_month":"Month","cost":"EUR"})
        st.plotly_chart(fig, use_container_width=True)

# 2) SVC distribution
if SVC_COL:
    st.subheader("🧩 Service Mix (SVC)")
    svc_df = df.groupby(SVC_COL).agg(
        shipments=("index","count") if "index" in df.columns else (SVC_COL,"size"),
        cost=(CHARGES_COL,"sum") if CHARGES_COL else (SVC_COL,"size"),
        avg_cost=(CHARGES_COL,"mean") if CHARGES_COL else (SVC_COL,"size")
    ).reset_index().sort_values("shipments", ascending=False)
    t1, t2 = st.columns(2)
    with t1:
        fig = px.bar(svc_df, x=SVC_COL, y="shipments", labels={SVC_COL:"SVC","shipments":"Shipments"})
        st.plotly_chart(fig, use_container_width=True)
    with t2:
        fig = px.bar(svc_df, x=SVC_COL, y="cost", labels={SVC_COL:"SVC","cost":"EUR"})
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(svc_df.rename(columns={"shipments":"Shipments","cost":"Cost (EUR)","avg_cost":"Avg Cost (EUR)"}))

# 3) Origin distribution (DEP)
if DEP_COL:
    st.subheader("🗺️ Distribution by DEP (Origin Station)")
    dep_df = df.groupby(DEP_COL).agg(
        shipments=("index","count") if "index" in df.columns else (DEP_COL,"size"),
        cost=(CHARGES_COL,"sum") if CHARGES_COL else (DEP_COL,"size")
    ).reset_index().sort_values("shipments", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(dep_df.head(25), x=DEP_COL, y="shipments", labels={DEP_COL:"DEP","shipments":"Shipments"})
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(dep_df.head(25), x=DEP_COL, y="cost", labels={DEP_COL:"DEP","cost":"EUR"})
        st.plotly_chart(fig, use_container_width=True)

# 4) Top Lanes (DEP -> Delivery Country)
if "LANE" in df.columns and df["LANE"].notna().any():
    st.subheader("🚚 Top Lanes (DEP → Delivery Country)")
    lane_df = df.groupby("LANE").agg(
        shipments=("index","count") if "index" in df.columns else ("LANE","size"),
        cost=(CHARGES_COL,"sum") if CHARGES_COL else ("LANE","size")
    ).reset_index().sort_values(["shipments","cost"], ascending=[False, False])
    fig = px.bar(lane_df.head(25), x="LANE", y="shipments", labels={"LANE":"Lane","shipments":"Shipments"})
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(lane_df.rename(columns={"shipments":"Shipments","cost":"Cost (EUR)"}))

# 5) Cost vs Weight
if CHARGES_COL and WEIGHT_KG_COL:
    st.subheader("⚖️ Cost vs Weight (KG)")
    fig = px.scatter(df, x=WEIGHT_KG_COL, y=CHARGES_COL, hover_data=[SVC_COL, DEP_COL, DEL_CTRY_COL] if SVC_COL and DEP_COL and DEL_CTRY_COL else None,
                     labels={WEIGHT_KG_COL:"Weight (KG)", CHARGES_COL:"Cost (EUR)"})
    st.plotly_chart(fig, use_container_width=True)

# 6) QC Reasons
if QCNAME_COL:
    st.subheader("🧪 QC Reasons (Billed Shipments)")
    qc_df = df[QCNAME_COL].value_counts(dropna=True).reset_index()
    qc_df.columns = ["QC NAME", "count"]
    fig = px.bar(qc_df.head(25), x="QC NAME", y="count", labels={"count":"Shipments"})
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(qc_df.rename(columns={"count":"Shipments"}))

# 7) Late shipments list (for actioning)
if ARRIVE_COL and PROMISED_COL:
    st.subheader("📋 Late Shipments (for actioning)")
    tmp = df.copy()
    tmp["actual_delivery"] = tmp[ARRIVE_COL]
    tmp["promised_delivery"] = tmp[PROMISED_COL]
    tmp["on_time"] = tmp["actual_delivery"] <= tmp["promised_delivery"]
    late = tmp[tmp["on_time"] == False].copy()
    if not late.empty:
        show_cols = [c for c in ["REFER", SVC_COL, DEP_COL, DEL_CTRY_COL, "LANE", CHARGES_COL, WEIGHT_KG_COL, PROMISED_COL, ARRIVE_COL, QCNAME_COL] if c in late.columns]
        st.dataframe(late[show_cols].sort_values(PROMISED_COL).reset_index(drop=True))
    else:
        st.success("No late shipments in the selected period.")

st.markdown("---")
with st.expander("ℹ️ Notes & Assumptions"):
    st.write(
        """
- **Dataset scope:** This dashboard automatically filters to shipments with **STATUS = 440-BILLED**.
- **OTP Gross** uses *actual delivery time* (prefers **POD DATE/TIME**, else **Arrive Date / Time**) vs **UPD DEL** (updated promised delivery).
- **OTP Net** treats late shipments with *selected* **QC NAME** reasons as **non‑controllable**, counting them as on-time. Adjust in the sidebar to match your policy.
- All monetary values are assumed to be **EUR** as provided by your extract.
- Columns are auto-detected by name; if a column is missing or named differently, KPIs depending on it may show as “—”.
        """
    )
