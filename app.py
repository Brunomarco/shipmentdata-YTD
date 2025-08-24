"""
Shipment Cost Analytics Dashboard
A comprehensive dashboard for analyzing shipment costs and performance metrics
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# -----------------------------
# Page configuration
# -----------------------------
st.set_page_config(
    page_title="Shipment Analytics Dashboard",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------
# Styles
# -----------------------------
st.markdown(
    """
    <style>
    .main { padding: 0rem 1rem; }
    .metric-card { background-color: #f0f2f6; padding: 1rem; border-radius: 0.5rem; margin: 0.5rem 0; }
    .stMetric { background-color: #ffffff; padding: 1rem; border-radius: 0.5rem; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    div[data-testid="metric-container"] { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 1rem; border-radius: 0.5rem; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📦 Shipment Cost Analytics Dashboard")
st.markdown("**Real-time insights for strategic decision-making**")
st.markdown("---")

# -----------------------------
# Constants & helpers
# -----------------------------
CONTROLLABLE_QC_CODES = [
    'MNX-Incorrect QDT', 'MNX-Order Entry error', 'MNX-Late dispatch-Delivery',
    'W/House-Data entry errors', 'Customs delay', 'Customs delay-FDA Hold',
    'Customs-Late PWK-Customer', 'Del Agt-Late del', 'Del Agt-Late del-Out of hours',
    'Del Agt-Missing documents', 'PU Agt -Late pick up', 'Airline-Slow offload',
    'Airline-RTA-DG PWK issue', 'Shipment not ready'
]

USD_TO_EUR = 0.92

@st.cache_data(show_spinner=False)
def load_data(file: bytes) -> pd.DataFrame | None:
    """Load and preprocess the Excel data from an uploaded file-like object."""
    try:
        df = pd.read_excel(file, engine='openpyxl')

        # Ensure columns exist
        if 'STATUS' not in df.columns:
            st.warning("Column 'STATUS' not found. The app expects a status column to filter billed shipments.")

        # Convert date columns if present
        date_columns = [
            'ORD CREATE', 'READY', 'QT PU', 'PICKUP DATE/TIME', 'Depart Date / Time',
            'Arrive Date / Time', 'QDT', 'UPD DEL', 'POD DATE/TIME'
        ]
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Numeric columns
        numeric_cols = ['TOTAL CHARGES', 'PIECES', 'WEIGHT(KG)', 'Billable Weight KG']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        # A helper column to enable counts when grouping
        if 'REFER' not in df.columns:
            df['REFER'] = 1

        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None


def calculate_otp_metrics(df: pd.DataFrame) -> tuple[float, float, int, int]:
    """Return (gross_otp%, net_otp%, on_time_count, total_considered). Robust to missing cols."""
    # Guard: required columns
    if not {'QDT', 'POD DATE/TIME'}.issubset(df.columns):
        return 0.0, 0.0, 0, 0

    otp_df = df[df['QDT'].notna() & df['POD DATE/TIME'].notna()].copy()
    if otp_df.empty:
        return 0.0, 0.0, 0, 0

    otp_df['on_time'] = otp_df['POD DATE/TIME'] <= otp_df['QDT']
    gross_otp = (otp_df['on_time'].mean()) * 100

    # Net OTP excludes controllable QC codes if column exists
    if 'QC NAME' in otp_df.columns:
        net_df = otp_df[~otp_df['QC NAME'].isin(CONTROLLABLE_QC_CODES)]
        net_otp = (net_df['on_time'].mean()) * 100 if len(net_df) else gross_otp
    else:
        net_otp = gross_otp

    return float(gross_otp), float(net_otp), int(otp_df['on_time'].sum()), int(len(otp_df))


# -----------------------------
# App
# -----------------------------

def main():
    # -------- Sidebar: Upload & Filters --------
    with st.sidebar:
        st.header("📊 Data Configuration")
        uploaded_file = st.file_uploader(
            "Upload Excel File",
            type=["xls", "xlsx"],
            help="Upload your shipment data Excel file",
        )

        df = None
        df_filtered = None

        if uploaded_file is not None:
            df = load_data(uploaded_file)
            if df is not None:
                # Filter billed rows if the column exists
                if 'STATUS' in df.columns:
                    df_billed = df[df['STATUS'] == '440-BILLED'].copy()
                else:
                    df_billed = df.copy()

                st.success(f"✅ Data loaded: {len(df_billed):,} records")

                # Date Filter on ORD CREATE if present
                if 'ORD CREATE' in df_billed.columns and not df_billed['ORD CREATE'].dropna().empty:
                    st.subheader("📅 Date Range Filter")
                    min_date = pd.to_datetime(df_billed['ORD CREATE']).min().date()
                    max_date = pd.to_datetime(df_billed['ORD CREATE']).max().date()
                    date_range = st.date_input(
                        "Select date range",
                        value=(min_date, max_date),
                        min_value=min_date,
                        max_value=max_date,
                    )
                    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
                        mask = (
                            df_billed['ORD CREATE'].dt.date >= date_range[0]
                        ) & (
                            df_billed['ORD CREATE'].dt.date <= date_range[1]
                        )
                        df_filtered = df_billed[mask].copy()
                    else:
                        df_filtered = df_billed.copy()
                else:
                    df_filtered = df_billed.copy()

                # Service filter
                if 'SVC' in df_filtered.columns:
                    st.subheader("🚚 Service Type Filter")
                    svc_vals = sorted(df_filtered['SVC'].dropna().unique().tolist())
                    selected_services = st.multiselect(
                        "Select services",
                        options=["All"] + svc_vals,
                        default=["All"],
                    )
                    if selected_services and "All" not in selected_services:
                        df_filtered = df_filtered[df_filtered['SVC'].isin(selected_services)].copy()

                # Departure filter
                if 'DEP' in df_filtered.columns:
                    st.subheader("✈️ Departure Filter")
                    dep_vals = sorted(df_filtered['DEP'].dropna().unique().tolist())
                    selected_deps = st.multiselect(
                        "Select departure points",
                        options=["All"] + dep_vals,
                        default=["All"],
                    )
                    if selected_deps and "All" not in selected_deps:
                        df_filtered = df_filtered[df_filtered['DEP'].isin(selected_deps)].copy()

                st.info(f"📊 Filtered records: {len(df_filtered):,}")

    # -------- Main: Visuals --------
    if uploaded_file is None or df is None or df_filtered is None:
        st.info("👈 Please upload your Excel file in the sidebar to begin analysis")
        st.markdown(
            """
            ### 📁 Required Excel Columns (recommended)
            - `STATUS` (filtered for `440-BILLED`)
            - Dates: `ORD CREATE`, `QDT`, `POD DATE/TIME`, etc.
            - Shipment meta: `SVC`, `SVCDESC`, `DEP`, `ARR`, `DEL CTRY`
            - Costs & weights: `TOTAL CHARGES`, `Billable Weight KG`
            - QC data: `QC NAME`
            """
        )
        return

    # Precompute commonly used fields once
    df_filtered = df_filtered.copy()
    if 'TOTAL CHARGES' in df_filtered.columns:
        df_filtered['TOTAL CHARGES EUR'] = df_filtered['TOTAL CHARGES'] * USD_TO_EUR
    else:
        df_filtered['TOTAL CHARGES EUR'] = 0.0

    gross_otp, net_otp, on_time_count, total_otp_records = calculate_otp_metrics(df_filtered)

    # Tabs
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📈 Overview",
        "⏰ OTP Analysis",
        "💰 Cost Analysis",
        "🌍 Geographic Analysis",
        "📊 Service Analysis",
        "🔍 Quality Control",
    ])

    # -------- Tab 1: Overview --------
    with tab1:
        st.header("Executive Summary")

        total_shipments = int(len(df_filtered))
        total_cost_eur = float(df_filtered['TOTAL CHARGES EUR'].sum())
        avg_cost_eur = float(df_filtered['TOTAL CHARGES EUR'].mean()) if total_shipments else 0.0
        total_weight = float(df_filtered.get('Billable Weight KG', pd.Series(dtype=float)).sum())

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Shipments", f"{total_shipments:,}")
        c2.metric("Total Cost (EUR)", f"€{total_cost_eur:,.2f}")
        c3.metric("Average Cost (EUR)", f"€{avg_cost_eur:,.2f}")
        c4.metric("Total Weight (KG)", f"{total_weight:,.0f}")

        c5, c6, c7, c8 = st.columns(4)
        c5.metric(
            "OTP Gross",
            f"{gross_otp:.1f}%",
            delta=f"{gross_otp - 85:.1f}% vs target",
            delta_color="normal" if gross_otp >= 85 else "inverse",
        )
        c6.metric(
            "OTP Net",
            f"{net_otp:.1f}%",
            delta=f"{net_otp - 90:.1f}% vs target",
            delta_color="normal" if net_otp >= 90 else "inverse",
        )
        qc_issues_n = int(df_filtered['QC NAME'].notna().sum()) if 'QC NAME' in df_filtered.columns else 0
        qc_rate = (qc_issues_n / total_shipments * 100) if total_shipments else 0
        c7.metric(
            "QC Issue Rate",
            f"{qc_rate:.1f}%",
            delta=f"{qc_rate - 10:.1f}% vs target",
            delta_color="inverse" if qc_rate > 10 else "normal",
        )
        unique_countries = int(df_filtered.get('DEL CTRY', pd.Series(dtype=object)).nunique())
        c8.metric("Countries Served", f"{unique_countries}")

        # Monthly trend (use Scatter, not go.Line)
        st.subheader("📊 Monthly Shipment Trends")
        if 'ORD CREATE' in df_filtered.columns and not df_filtered['ORD CREATE'].dropna().empty:
            df_monthly = df_filtered.copy()
            df_monthly['Month'] = df_monthly['ORD CREATE'].dt.to_period('M').astype(str)
            monthly_stats = df_monthly.groupby('Month').agg(
                Shipments=('REFER', 'count'),
                TotalCostEUR=('TOTAL CHARGES EUR', 'sum')
            ).reset_index()

            fig = make_subplots(rows=1, cols=2, subplot_titles=(
                'Monthly Shipment Volume', 'Monthly Cost (EUR)'
            ))
            fig.add_trace(
                go.Bar(x=monthly_stats['Month'], y=monthly_stats['Shipments'], name='Shipments'),
                row=1, col=1
            )
            fig.add_trace(
                go.Scatter(x=monthly_stats['Month'], y=monthly_stats['TotalCostEUR'], mode='lines+markers', name='Cost (EUR)'),
                row=1, col=2
            )
            fig.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    # -------- Tab 2: OTP --------
    with tab2:
        st.header("On-Time Performance Analysis")

        col1, col2 = st.columns(2)
        with col1:
            fig_gross = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=gross_otp,
                delta={'reference': 85},
                title={'text': 'OTP Gross %'},
                gauge={
                    'axis': {'range': [0, 100]},
                    'bar': {'color': 'darkblue'},
                    'steps': [
                        {'range': [0, 50], 'color': 'lightgray'},
                        {'range': [50, 85], 'color': 'gray'}
                    ],
                    'threshold': {'line': {'color': 'red', 'width': 4}, 'thickness': 0.75, 'value': 85}
                }
            ))
            fig_gross.update_layout(height=300)
            st.plotly_chart(fig_gross, use_container_width=True)

        with col2:
            fig_net = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=net_otp,
                delta={'reference': 90},
                title={'text': 'OTP Net % (Excl. Controllable)'},
                gauge={
                    'axis': {'range': [0, 100]},
                    'bar': {'color': 'darkgreen'},
                    'steps': [
                        {'range': [0, 60], 'color': 'lightgray'},
                        {'range': [60, 90], 'color': 'gray'}
                    ],
                    'threshold': {'line': {'color': 'red', 'width': 4}, 'thickness': 0.75, 'value': 90}
                }
            ))
            fig_net.update_layout(height=300)
            st.plotly_chart(fig_net, use_container_width=True)

        # OTP by Service Type (bar)
        st.subheader("OTP by Service Type")
        if 'SVC' in df_filtered.columns:
            rows = []
            for svc, svc_df in df_filtered.groupby('SVC'):
                g, n, _, _ = calculate_otp_metrics(svc_df)
                rows.append({'Service': svc, 'OTP Gross': g, 'OTP Net': n, 'Shipments': len(svc_df)})
            if rows:
                otp_service_df = pd.DataFrame(rows).sort_values('Shipments', ascending=False)
                fig_otp_service = go.Figure()
                fig_otp_service.add_trace(go.Bar(x=otp_service_df['Service'], y=otp_service_df['OTP Gross'], name='OTP Gross'))
                fig_otp_service.add_trace(go.Bar(x=otp_service_df['Service'], y=otp_service_df['OTP Net'], name='OTP Net'))
                fig_otp_service.update_layout(barmode='group', height=420, xaxis_title='Service', yaxis_title='OTP %')
                st.plotly_chart(fig_otp_service, use_container_width=True)

        # Cost trend vs day
        st.subheader("Cost Trends Over Time")
        if 'ORD CREATE' in df_filtered.columns and not df_filtered['ORD CREATE'].dropna().empty:
            daily = df_filtered.copy()
            daily['Date'] = daily['ORD CREATE'].dt.date
            daily_cost = daily.groupby('Date').agg(TotalEUR=('TOTAL CHARGES EUR', 'sum'), Shipments=('REFER', 'count')).reset_index()
            daily_cost['AvgEUR'] = daily_cost['TotalEUR'] / daily_cost['Shipments']

            fig_cost_trend = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                           subplot_titles=('Daily Total Cost (EUR)', 'Daily Average Cost per Shipment (EUR)'))
            fig_cost_trend.add_trace(go.Scatter(x=daily_cost['Date'], y=daily_cost['TotalEUR'], mode='lines+markers', name='Total Cost'), row=1, col=1)
            fig_cost_trend.add_trace(go.Scatter(x=daily_cost['Date'], y=daily_cost['AvgEUR'], mode='lines+markers', name='Avg Cost'), row=2, col=1)
            fig_cost_trend.update_layout(height=600, showlegend=False)
            st.plotly_chart(fig_cost_trend, use_container_width=True)

    # -------- Tab 3: Cost --------
    with tab3:
        st.header("Cost Analysis")
        col1, col2 = st.columns(2)

        # Cost by Service Type (Top 10)
        if 'SVC' in df_filtered.columns:
            cost_by_service = df_filtered.groupby('SVC').agg(
                **{"Total Cost (EUR)": ('TOTAL CHARGES EUR', 'sum'), 'Shipments': ('REFER', 'count')}
            ).reset_index().sort_values('Total Cost (EUR)', ascending=False)
        else:
            cost_by_service = pd.DataFrame(columns=['SVC', 'Total Cost (EUR)', 'Shipments']).rename(columns={'SVC': 'Service'})

        with col1:
            if not cost_by_service.empty:
                fig_cost_service = px.pie(
                    cost_by_service.head(10).rename(columns={'SVC': 'Service'}),
                    values='Total Cost (EUR)',
                    names='Service',
                    title='Cost Distribution by Service Type (Top 10)'
                )
                fig_cost_service.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_cost_service, use_container_width=True)

        with col2:
            if not cost_by_service.empty:
                cbs = cost_by_service.copy()
                cbs['Avg Cost (EUR)'] = cbs['Total Cost (EUR)'] / cbs['Shipments']
                cbs = cbs.rename(columns={'SVC': 'Service'})
                fig_avg_cost = px.bar(
                    cbs.head(10).sort_values('Avg Cost (EUR)', ascending=True),
                    x='Avg Cost (EUR)', y='Service', orientation='h', title='Average Cost per Shipment by Service'
                )
                st.plotly_chart(fig_avg_cost, use_container_width=True)

        # Cost summary table
        st.subheader("Detailed Cost Breakdown")
        if {'SVC', 'SVCDESC'}.issubset(df_filtered.columns):
            cost_summary = (
                df_filtered.groupby(['SVC', 'SVCDESC']).agg(
                    Shipments=('REFER', 'count'),
                    **{"Total Cost (EUR)": ('TOTAL CHARGES EUR', 'sum')},
                    **{"Avg Cost (EUR)": ('TOTAL CHARGES EUR', 'mean')},
                    **{"Min Cost (EUR)": ('TOTAL CHARGES EUR', 'min')},
                    **{"Max Cost (EUR)": ('TOTAL CHARGES EUR', 'max')},
                    **{"Avg Weight (KG)": ('Billable Weight KG', 'mean')}
                ).round(2)
            )
            st.dataframe(
                cost_summary.style.format({
                    'Total Cost (EUR)': '€{:,.2f}', 'Avg Cost (EUR)': '€{:,.2f}',
                    'Min Cost (EUR)': '€{:,.2f}', 'Max Cost (EUR)': '€{:,.2f}',
                    'Avg Weight (KG)': '{:,.1f}'
                }),
                use_container_width=True
            )

    # -------- Tab 4: Geography --------
    with tab4:
        st.header("Geographic Analysis")

        # Departures
        if 'DEP' in df_filtered.columns:
            dep_stats = df_filtered.groupby('DEP').agg(Shipments=('REFER', 'count'), **{"Total Cost (EUR)": ('TOTAL CHARGES EUR', 'sum')}).reset_index()
            dep_stats = dep_stats.sort_values('Shipments', ascending=False).head(15)
            fig_dep = px.bar(dep_stats, x='DEP', y='Shipments', title='Top 15 Departure Points by Volume', color='Total Cost (EUR)')
            fig_dep.update_xaxes(title='Departure')
            st.plotly_chart(fig_dep, use_container_width=True)

        # Destinations
        if 'DEL CTRY' in df_filtered.columns:
            dest_stats = df_filtered.groupby('DEL CTRY').agg(Shipments=('REFER', 'count'), **{"Total Cost (EUR)": ('TOTAL CHARGES EUR', 'sum')}).reset_index()
            dest_stats = dest_stats.sort_values('Shipments', ascending=False).head(15)
            fig_dest = px.treemap(dest_stats, path=['DEL CTRY'], values='Shipments', color='Total Cost (EUR)', title='Destination Countries - Volume and Cost')
            st.plotly_chart(fig_dest, use_container_width=True)

        # Routes
        if {'DEP', 'ARR'}.issubset(df_filtered.columns):
            route_stats = df_filtered.groupby(['DEP', 'ARR']).agg(Shipments=('REFER', 'count'), **{"Avg Cost (EUR)": ('TOTAL CHARGES EUR', 'mean')}).reset_index()
            route_stats = route_stats.sort_values('Shipments', ascending=False).head(20)
            route_stats['Route'] = route_stats['DEP'] + ' → ' + route_stats['ARR']
            fig_routes = px.scatter(route_stats, x='Shipments', y='Avg Cost (EUR)', size='Shipments', hover_data=['Route'], title='Top 20 Routes: Volume vs Average Cost')
            st.plotly_chart(fig_routes, use_container_width=True)

        # OTP by departure
        st.subheader("OTP Performance by Departure Point")
        if 'DEP' in df_filtered.columns:
            rows = []
            for dep, dep_df in df_filtered.groupby('DEP'):
                if len(dep_df) >= 10:
                    g, n, _, _ = calculate_otp_metrics(dep_df)
                    rows.append({'Departure': dep, 'OTP Gross': g, 'OTP Net': n, 'Shipments': len(dep_df)})
            if rows:
                geo_otp_df = pd.DataFrame(rows).sort_values('Shipments', ascending=False)
                fig_geo = px.scatter(geo_otp_df, x='OTP Gross', y='OTP Net', size='Shipments', hover_data=['Departure', 'Shipments'], title='OTP by Departure (size = volume)')
                fig_geo.add_hline(y=90, line_dash='dash', line_color='red', annotation_text='Net OTP Target (90%)')
                fig_geo.add_vline(x=85, line_dash='dash', line_color='blue', annotation_text='Gross OTP Target (85%)')
                st.plotly_chart(fig_geo, use_container_width=True)

    # -------- Tab 5: Service --------
    with tab5:
        st.header("Service Analysis")

        # Distribution
        if 'SVC' in df_filtered.columns:
            svc_dist = df_filtered['SVC'].value_counts().reset_index().rename(columns={'index': 'Service', 'SVC': 'Count'})
            c1, c2 = st.columns(2)
            with c1:
                fig_svc = px.pie(svc_dist, values='Count', names='Service', title='Service Type Distribution')
                fig_svc.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_svc, use_container_width=True)

            with c2:
                # Metrics per service
                rows = []
                for svc, svc_df in df_filtered.groupby('SVC'):
                    g, n, _, _ = calculate_otp_metrics(svc_df)
                    rows.append({
                        'Service': svc,
                        'Volume': len(svc_df),
                        'Total Cost (EUR)': svc_df['TOTAL CHARGES EUR'].sum(),
                        'Avg Cost (EUR)': svc_df['TOTAL CHARGES EUR'].mean(),
                        'OTP Gross': g,
                        'OTP Net': n,
                    })
                svc_metrics_df = pd.DataFrame(rows).sort_values('Volume', ascending=False)

                # Radar for top 5 services
                top5 = svc_metrics_df.head(5)
                if not top5.empty:
                    total_vol = svc_metrics_df['Volume'].sum()
                    total_cost = svc_metrics_df['Total Cost (EUR)'].sum()
                    categories = ['Volume %', 'Cost Share %', 'OTP Gross', 'OTP Net']
                    fig_radar = go.Figure()
                    for _, r in top5.iterrows():
                        values = [
                            (r['Volume'] / total_vol) * 100 if total_vol else 0,
                            (r['Total Cost (EUR)'] / total_cost) * 100 if total_cost else 0,
                            r['OTP Gross'],
                            r['OTP Net'],
                        ]
                        fig_radar.add_trace(go.Scatterpolar(r=values, theta=categories, fill='toself', name=r['Service']))
                    fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), showlegend=True, title='Top 5 Services Performance Comparison')
                    st.plotly_chart(fig_radar, use_container_width=True)

                st.subheader("Service Performance Summary")
                st.dataframe(
                    svc_metrics_df.style.format({
                        'Total Cost (EUR)': '€{:,.2f}', 'Avg Cost (EUR)': '€{:,.2f}', 'OTP Gross': '{:.1f}%', 'OTP Net': '{:.1f}%'
                    }),
                    use_container_width=True,
                )

        # Trends
        if 'ORD CREATE' in df_filtered.columns and 'SVC' in df_filtered.columns and not df_filtered['ORD CREATE'].dropna().empty:
            st.subheader("Service Volume Trends (Top 5)")
            df_svc_trend = df_filtered.copy()
            df_svc_trend['Month'] = df_svc_trend['ORD CREATE'].dt.to_period('M').astype(str)
            svc_monthly = df_svc_trend.groupby(['Month', 'SVC']).size().reset_index(name='Count')
            top_svcs = df_filtered['SVC'].value_counts().head(5).index
            svc_monthly_top = svc_monthly[svc_monthly['SVC'].isin(top_svcs)]
            fig_svc_trend = px.line(svc_monthly_top, x='Month', y='Count', color='SVC', markers=True, title='Monthly Trends - Top 5 Services')
            st.plotly_chart(fig_svc_trend, use_container_width=True)

    # -------- Tab 6: QC --------
    with tab6:
        st.header("Quality Control Analysis")

        has_qc = 'QC NAME' in df_filtered.columns
        qc_issues_df = df_filtered[df_filtered['QC NAME'].notna()] if has_qc else pd.DataFrame(columns=['QC NAME'])

        c1, c2, c3 = st.columns(3)
        c1.metric("Total QC Issues", f"{len(qc_issues_df):,}", f"{(len(qc_issues_df)/len(df_filtered)*100 if len(df_filtered) else 0):.1f}% of shipments")
        controllable_n = int(qc_issues_df['QC NAME'].isin(CONTROLLABLE_QC_CODES).sum()) if has_qc else 0
        uncontrollable_n = max(0, len(qc_issues_df) - controllable_n)
        c2.metric("Controllable Issues", f"{controllable_n:,}", f"{(controllable_n/len(qc_issues_df)*100 if len(qc_issues_df) else 0):.1f}% of QC issues")
        c3.metric("Uncontrollable Issues", f"{uncontrollable_n:,}", f"{(uncontrollable_n/len(qc_issues_df)*100 if len(qc_issues_df) else 0):.1f}% of QC issues")

        # Distribution
        st.subheader("Quality Control Issues Distribution")
        if has_qc and not qc_issues_df.empty:
            qc_dist = qc_issues_df['QC NAME'].value_counts().reset_index().rename(columns={'index': 'QC Issue', 'QC NAME': 'Count'})
            qc_dist['Type'] = qc_dist['QC Issue'].apply(lambda x: 'Controllable' if x in CONTROLLABLE_QC_CODES else 'Uncontrollable')
            fig_qc = px.bar(qc_dist.head(20), x='Count', y='QC Issue', color='Type', orientation='h', title='Top 20 Quality Control Issues')
            st.plotly_chart(fig_qc, use_container_width=True)

        # QC by service
        if has_qc and 'SVC' in df_filtered.columns:
            st.subheader("QC Issues by Service Type (Top 10)")
            qc_by_svc = qc_issues_df.groupby('SVC')['QC NAME'].count().reset_index().rename(columns={'QC NAME': 'QC Issues'})
            qc_by_svc = qc_by_svc.sort_values('QC Issues', ascending=False).head(10)
            fig_qc_svc = px.bar(qc_by_svc, x='SVC', y='QC Issues', title='QC Issues by Service Type (Top 10)')
            fig_qc_svc.update_xaxes(title='Service')
            st.plotly_chart(fig_qc_svc, use_container_width=True)

        # QC trends
        if has_qc and 'ORD CREATE' in df_filtered.columns and not df_filtered['ORD CREATE'].dropna().empty:
            st.subheader("QC Issues Trend Analysis")
            df_qc_trend = df_filtered.copy()
            df_qc_trend['Month'] = df_qc_trend['ORD CREATE'].dt.to_period('M').astype(str)
            monthly_qc = df_qc_trend.groupby('Month').agg(**{"Total Shipments": ('REFER', 'count')}, **{"QC Issues": ('QC NAME', lambda x: x.notna().sum())}).reset_index()
            monthly_qc['QC Rate (%)'] = (monthly_qc['QC Issues'] / monthly_qc['Total Shipments']) * 100
            fig_qc_trend = make_subplots(rows=2, cols=1, shared_xaxes=True, subplot_titles=('QC Issues Over Time', 'QC Issue Rate (%)'))
            fig_qc_trend.add_trace(go.Bar(x=monthly_qc['Month'], y=monthly_qc['QC Issues'], name='QC Issues'), row=1, col=1)
            fig_qc_trend.add_trace(go.Scatter(x=monthly_qc['Month'], y=monthly_qc['QC Rate (%)'], mode='lines+markers', name='QC Rate'), row=2, col=1)
            fig_qc_trend.update_layout(height=600, showlegend=False)
            st.plotly_chart(fig_qc_trend, use_container_width=True)

        st.subheader("📋 Actionable Insights")
        st.info(
            """
            **Key Findings:**
            1. **Controllable Issues:** Focus on MNX-related errors, warehouse data entry quality, and customs documentation.
            2. **Customer Issues:** High volume of customer-requested delays and changed parameters – consider customer education.
            3. **Process Improvements:** Implement automated QDT validation to reduce "Incorrect QDT" issues.
            4. **Training Needs:** Address delivery agent delays via performance management and training.
            """
        )


if __name__ == "__main__":
    main()
