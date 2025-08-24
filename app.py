"""
Shipment Analytics Dashboard - Executive Edition
Streamlined for Strategic Decision Making
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Executive Shipment Analytics",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional styling
st.markdown("""
    <style>
    .main {font-family: 'Arial', sans-serif;}
    div[data-testid="metric-container"] {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border: 1px solid #e2e8f0;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        padding: 0 20px;
        background-color: #ffffff;
        border-radius: 8px 8px 0 0;
        font-weight: 500;
    }
    </style>
    """, unsafe_allow_html=True)

# Title
st.markdown("<h1 style='text-align: center; color: #1e293b;'>🎯 Executive Shipment Analytics</h1>", unsafe_allow_html=True)
st.markdown("---")

# Data loading function
@st.cache_data
def load_data(file):
    """Load and preprocess Excel data"""
    try:
        df = pd.read_excel(file, engine='openpyxl')
        
        # Convert date columns
        date_cols = ['ORD CREATE', 'QDT', 'POD DATE/TIME', 'Depart Date / Time', 'Arrive Date / Time']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Convert numeric columns
        numeric_cols = ['TOTAL CHARGES', 'PIECES', 'WEIGHT(KG)', 'Billable Weight KG']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

# Controllable QC codes
CONTROLLABLE_QC = [
    'MNX-Incorrect QDT', 'MNX-Order Entry error', 'MNX-Late dispatch-Delivery',
    'W/House-Data entry errors', 'Customs delay', 'Customs delay-FDA Hold',
    'Customs-Late PWK-Customer', 'Del Agt-Late del', 'Del Agt-Late del-Out of hours',
    'Del Agt-Missing documents', 'PU Agt -Late pick up', 'Airline-Slow offload',
    'Airline-RTA-DG PWK issue', 'Shipment not ready'
]

USD_TO_EUR = 0.92

def calculate_otp(df):
    """Calculate OTP metrics"""
    otp_df = df[(df['QDT'].notna()) & (df['POD DATE/TIME'].notna())].copy()
    if len(otp_df) == 0:
        return 0, 0, 0
    
    otp_df['on_time'] = otp_df['POD DATE/TIME'] <= otp_df['QDT']
    gross_otp = (otp_df['on_time'].sum() / len(otp_df)) * 100
    
    # Net OTP excluding controllable
    net_df = otp_df[~otp_df['QC NAME'].isin(CONTROLLABLE_QC)]
    net_otp = (net_df['on_time'].sum() / len(net_df)) * 100 if len(net_df) > 0 else gross_otp
    
    controllable_delays = otp_df[~otp_df['on_time'] & otp_df['QC NAME'].isin(CONTROLLABLE_QC)].shape[0]
    
    return gross_otp, net_otp, controllable_delays

def main():
    # Sidebar
    with st.sidebar:
        st.header("📊 Configuration")
        
        uploaded_file = st.file_uploader("Upload Excel File", type=['xls', 'xlsx'])
        
        if uploaded_file:
            df = load_data(uploaded_file)
            
            if df is not None:
                # Filter billed records
                df_billed = df[df['STATUS'] == '440-BILLED'].copy()
                df_billed['TOTAL CHARGES EUR'] = df_billed['TOTAL CHARGES'] * USD_TO_EUR
                
                st.success(f"✅ {len(df_billed):,} records loaded")
                
                # Date filter
                st.subheader("📅 Date Range")
                if 'ORD CREATE' in df_billed.columns:
                    min_date = df_billed['ORD CREATE'].min()
                    max_date = df_billed['ORD CREATE'].max()
                    
                    date_range = st.date_input(
                        "Select dates",
                        value=(max_date - timedelta(days=30), max_date),
                        min_value=min_date,
                        max_value=max_date
                    )
                    
                    if len(date_range) == 2:
                        mask = (df_billed['ORD CREATE'].dt.date >= date_range[0]) & \
                               (df_billed['ORD CREATE'].dt.date <= date_range[1])
                        df_filtered = df_billed[mask]
                    else:
                        df_filtered = df_billed
                else:
                    df_filtered = df_billed
                
                # Service filter
                st.subheader("🚚 Services")
                services = df_filtered['SVC'].dropna().unique()
                selected_svc = st.multiselect("Filter services", services, default=list(services))
                if selected_svc:
                    df_filtered = df_filtered[df_filtered['SVC'].isin(selected_svc)]
                
                st.info(f"📊 Analyzing {len(df_filtered):,} shipments")
    
    # Main dashboard
    if 'df_filtered' in locals():
        # Calculate metrics
        total_shipments = len(df_filtered)
        total_cost = df_filtered['TOTAL CHARGES EUR'].sum()
        avg_cost = df_filtered['TOTAL CHARGES EUR'].mean()
        gross_otp, net_otp, controllable_delays = calculate_otp(df_filtered)
        qc_rate = (df_filtered['QC NAME'].notna().sum() / total_shipments * 100) if total_shipments > 0 else 0
        
        # KPIs Row
        st.markdown("## 📊 Key Performance Indicators")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("Shipments", f"{total_shipments:,}")
        with col2:
            st.metric("Total Cost", f"€{total_cost/1000:.1f}K")
        with col3:
            st.metric("Avg Cost", f"€{avg_cost:.0f}")
        with col4:
            st.metric("OTP Gross", f"{gross_otp:.1f}%", f"{gross_otp-85:.1f}%",
                     delta_color="normal" if gross_otp >= 85 else "inverse")
        with col5:
            st.metric("QC Issues", f"{qc_rate:.1f}%", f"{qc_rate-10:.1f}%",
                     delta_color="inverse" if qc_rate > 10 else "normal")
        
        # OTP Explanation
        with st.expander("🎯 **Understanding OTP Metrics**", expanded=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.info(f"""
                **OTP Gross: {gross_otp:.1f}%**  
                Overall on-time performance including all delays.  
                Target: ≥85%
                """)
            with col2:
                st.success(f"""
                **OTP Net: {net_otp:.1f}%**  
                Performance excluding controllable issues.  
                Target: ≥90%
                """)
            with col3:
                st.warning(f"""
                **Improvement Potential: +{net_otp-gross_otp:.1f}%**  
                {controllable_delays} controllable delays could be fixed.
                """)
        
        # Analysis Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📈 Trends", "💰 Costs", "🌍 Routes", "⚠️ Quality"])
        
        with tab1:
            st.subheader("Performance Trends")
            
            if 'ORD CREATE' in df_filtered.columns:
                # Monthly analysis
                df_monthly = df_filtered.copy()
                df_monthly['Month'] = df_monthly['ORD CREATE'].dt.to_period('M')
                
                monthly = df_monthly.groupby('Month').agg({
                    'REFER': 'count',
                    'TOTAL CHARGES EUR': 'sum',
                    'QC NAME': lambda x: x.notna().sum()
                }).reset_index()
                monthly['Month'] = monthly['Month'].astype(str)
                monthly['QC Rate'] = (monthly['QC NAME'] / monthly['REFER']) * 100
                
                # Create subplots
                fig = make_subplots(
                    rows=2, cols=2,
                    subplot_titles=('Monthly Volume', 'Monthly Cost (EUR)', 'QC Rate Trend', 'OTP Trend')
                )
                
                # Volume
                fig.add_trace(
                    go.Bar(x=monthly['Month'], y=monthly['REFER'], name='Volume',
                          marker_color='#3b82f6'),
                    row=1, col=1
                )
                
                # Cost
                fig.add_trace(
                    go.Line(x=monthly['Month'], y=monthly['TOTAL CHARGES EUR'],
                           name='Cost', marker_color='#10b981'),
                    row=1, col=2
                )
                
                # QC Rate
                fig.add_trace(
                    go.Scatter(x=monthly['Month'], y=monthly['QC Rate'],
                              name='QC Rate', mode='lines+markers',
                              marker_color='#f59e0b'),
                    row=2, col=1
                )
                fig.add_hline(y=10, line_dash="dash", line_color="red", row=2, col=1)
                
                # OTP by month
                monthly_otp = []
                for month in df_monthly['Month'].unique():
                    month_df = df_monthly[df_monthly['Month'] == month]
                    otp, _, _ = calculate_otp(month_df)
                    monthly_otp.append(otp)
                
                fig.add_trace(
                    go.Scatter(x=monthly['Month'], y=monthly_otp,
                              name='OTP', mode='lines+markers',
                              marker_color='#06b6d4'),
                    row=2, col=2
                )
                fig.add_hline(y=85, line_dash="dash", line_color="red", row=2, col=2)
                
                fig.update_layout(height=600, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Cost Analysis")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Service cost breakdown
                svc_cost = df_filtered.groupby('SVC').agg({
                    'TOTAL CHARGES EUR': 'sum',
                    'REFER': 'count'
                }).reset_index()
                svc_cost['Avg Cost'] = svc_cost['TOTAL CHARGES EUR'] / svc_cost['REFER']
                
                fig_pie = px.pie(
                    svc_cost.head(10),
                    values='TOTAL CHARGES EUR',
                    names='SVC',
                    title='Cost Distribution by Service'
                )
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col2:
                # Cost efficiency matrix
                efficiency = []
                for svc in df_filtered['SVC'].dropna().unique():
                    svc_df = df_filtered[df_filtered['SVC'] == svc]
                    if len(svc_df) > 5:
                        otp, _, _ = calculate_otp(svc_df)
                        efficiency.append({
                            'Service': svc,
                            'Avg Cost': svc_df['TOTAL CHARGES EUR'].mean(),
                            'OTP': otp,
                            'Volume': len(svc_df)
                        })
                
                if efficiency:
                    eff_df = pd.DataFrame(efficiency)
                    fig_eff = px.scatter(
                        eff_df,
                        x='Avg Cost',
                        y='OTP',
                        size='Volume',
                        color='Volume',
                        hover_data=['Service'],
                        title='Service Efficiency (Cost vs Performance)'
                    )
                    fig_eff.add_hline(y=85, line_dash="dash", line_color="gray", opacity=0.5)
                    fig_eff.add_vline(x=eff_df['Avg Cost'].median(), line_dash="dash", 
                                     line_color="gray", opacity=0.5)
                    st.plotly_chart(fig_eff, use_container_width=True)
            
            # Cost drivers
            st.subheader("Cost Drivers")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Top expensive destinations
                dest_cost = df_filtered.groupby('DEL CTRY')['TOTAL CHARGES EUR'].mean()
                dest_cost = dest_cost.nlargest(10)
                
                fig_dest = go.Figure(go.Bar(
                    x=dest_cost.values,
                    y=dest_cost.index,
                    orientation='h',
                    marker_color='#ef4444'
                ))
                fig_dest.update_layout(title='Most Expensive Destinations', height=300)
                st.plotly_chart(fig_dest, use_container_width=True)
            
            with col2:
                # Weight vs cost
                weight_cost = df_filtered[['Billable Weight KG', 'TOTAL CHARGES EUR']].dropna()
                if len(weight_cost) > 0:
                    fig_weight = px.scatter(
                        weight_cost.sample(min(500, len(weight_cost))),
                        x='Billable Weight KG',
                        y='TOTAL CHARGES EUR',
                        trendline='ols',
                        title='Weight Impact on Cost'
                    )
                    fig_weight.update_layout(height=300)
                    st.plotly_chart(fig_weight, use_container_width=True)
            
            with col3:
                # Service comparison table
                svc_summary = df_filtered.groupby('SVC').agg({
                    'REFER': 'count',
                    'TOTAL CHARGES EUR': ['sum', 'mean']
                }).round(2)
                svc_summary.columns = ['Count', 'Total €', 'Avg €']
                st.dataframe(svc_summary.sort_values('Total €', ascending=False).head(10))
        
        with tab3:
            st.subheader("Geographic Analysis")
            
            # Hub performance
            col1, col2 = st.columns(2)
            
            with col1:
                # Departure hub analysis
                dep_stats = []
                for dep in df_filtered['DEP'].value_counts().head(15).index:
                    dep_df = df_filtered[df_filtered['DEP'] == dep]
                    otp, _, _ = calculate_otp(dep_df)
                    dep_stats.append({
                        'Hub': dep,
                        'Volume': len(dep_df),
                        'OTP': otp,
                        'Avg Cost': dep_df['TOTAL CHARGES EUR'].mean()
                    })
                
                dep_df = pd.DataFrame(dep_stats)
                
                fig_hub = px.scatter(
                    dep_df,
                    x='Volume',
                    y='OTP',
                    size='Avg Cost',
                    color='Avg Cost',
                    hover_data=['Hub'],
                    title='Hub Performance Analysis',
                    color_continuous_scale='RdYlGn_r'
                )
                fig_hub.add_hline(y=85, line_dash="dash", line_color="red")
                st.plotly_chart(fig_hub, use_container_width=True)
            
            with col2:
                # Top routes
                routes = df_filtered.groupby(['DEP', 'ARR']).size().reset_index(name='Count')
                routes['Route'] = routes['DEP'] + ' → ' + routes['ARR']
                routes = routes.nlargest(15, 'Count')
                
                fig_routes = px.bar(
                    routes,
                    x='Count',
                    y='Route',
                    orientation='h',
                    title='Top 15 Routes by Volume'
                )
                st.plotly_chart(fig_routes, use_container_width=True)
            
            # Country distribution
            country_stats = df_filtered.groupby('DEL CTRY').agg({
                'REFER': 'count',
                'TOTAL CHARGES EUR': 'sum'
            }).reset_index()
            country_stats.columns = ['Country', 'Shipments', 'Revenue']
            
            fig_tree = px.treemap(
                country_stats.head(30),
                path=['Country'],
                values='Shipments',
                color='Revenue',
                title='Destination Countries by Volume and Revenue',
                color_continuous_scale='Viridis'
            )
            fig_tree.update_layout(height=500)
            st.plotly_chart(fig_tree, use_container_width=True)
        
        with tab4:
            st.subheader("Quality Control Analysis")
            
            # QC Overview
            qc_data = df_filtered[df_filtered['QC NAME'].notna()]
            controllable = qc_data[qc_data['QC NAME'].isin(CONTROLLABLE_QC)]
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total QC Issues", f"{len(qc_data):,}")
            with col2:
                st.metric("Controllable", f"{len(controllable):,}")
            with col3:
                st.metric("Uncontrollable", f"{len(qc_data) - len(controllable):,}")
            with col4:
                improvement = (len(controllable) / len(df_filtered) * 100) if len(df_filtered) > 0 else 0
                st.metric("Potential OTP Gain", f"+{improvement:.1f}%")
            
            # QC Distribution
            col1, col2 = st.columns(2)
            
            with col1:
                # Top QC issues
                if not qc_data.empty:
                    qc_counts = qc_data['QC NAME'].value_counts().head(10)
                    
                    fig_qc = go.Figure(go.Bar(
                        x=qc_counts.values,
                        y=qc_counts.index,
                        orientation='h',
                        marker_color=['#ef4444' if issue in CONTROLLABLE_QC else '#3b82f6' 
                                     for issue in qc_counts.index]
                    ))
                    fig_qc.update_layout(title='Top 10 Quality Issues', height=400)
                    st.plotly_chart(fig_qc, use_container_width=True)
            
            with col2:
                # Controllable vs Uncontrollable
                qc_summary = pd.DataFrame({
                    'Type': ['Controllable', 'Uncontrollable'],
                    'Count': [len(controllable), len(qc_data) - len(controllable)]
                })
                
                fig_pie_qc = px.pie(
                    qc_summary,
                    values='Count',
                    names='Type',
                    title='QC Issues Breakdown',
                    color_discrete_map={'Controllable': '#f59e0b', 'Uncontrollable': '#3b82f6'}
                )
                st.plotly_chart(fig_pie_qc, use_container_width=True)
            
            # Recommendations
            st.markdown("### 🎯 Action Items")
            
            col1, col2 = st.columns(2)
            with col1:
                st.warning("""
                **Internal Improvements:**
                - Automate QDT validation
                - Enhance order entry training
                - Upgrade warehouse systems
                - Implement real-time monitoring
                """)
            
            with col2:
                st.info("""
                **Partner Actions:**
                - Review delivery SLAs
                - Implement scorecards
                - Improve customs process
                - Enhance communication
                """)
            
            # ROI Calculation
            if len(controllable) > 0:
                avg_delay_cost = controllable['TOTAL CHARGES EUR'].mean() if 'TOTAL CHARGES EUR' in controllable.columns else 100
                annual_impact = len(controllable) * avg_delay_cost * 0.1 * 12
                
                st.success(f"""
                **💰 ROI Estimation**  
                Annual Savings Potential: €{annual_impact:,.0f}  
                OTP Improvement: +{improvement:.1f}%  
                Customer Satisfaction: High Impact
                """)
    
    else:
        # Welcome screen
        st.info("👈 Upload your Excel file to begin analysis")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("### 📈 Performance Metrics")
            st.write("Track OTP, costs, and quality metrics")
        with col2:
            st.markdown("### 🎯 Strategic Insights")
            st.write("Identify optimization opportunities")
        with col3:
            st.markdown("### 💰 Cost Analysis")
            st.write("Understand cost drivers and savings")

if __name__ == "__main__":
    main()
