"""
Shipment Cost Analytics Dashboard
Executive Dashboard for Strategic Decision Making
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

# Professional styling
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    .main {
        font-family: 'Inter', sans-serif;
        background-color: #f8f9fa;
    }
    
    /* Headers styling */
    h1, h2, h3 {
        font-weight: 600;
        color: #1e293b;
    }
    
    /* Metric cards enhancement */
    div[data-testid="metric-container"] {
        background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
        border: 1px solid #e2e8f0;
        padding: 1.2rem;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: transform 0.2s;
    }
    
    div[data-testid="metric-container"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* Sidebar styling */
    .css-1d391kg {
        background-color: #1e293b;
    }
    
    /* Info boxes */
    .stAlert {
        border-radius: 8px;
        border-left: 4px solid #3b82f6;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-left: 20px;
        padding-right: 20px;
        background-color: #ffffff;
        border-radius: 8px 8px 0 0;
        border: 1px solid #e2e8f0;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #3b82f6;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

# Header with company branding
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.markdown("""
    <div style='text-align: center; padding: 1rem 0;'>
        <h1 style='color: #1e293b; font-size: 2.5rem; margin-bottom: 0.5rem;'>
            🎯 Executive Analytics Dashboard
        </h1>
        <p style='color: #64748b; font-size: 1.1rem;'>
            Real-Time Shipment Performance & Cost Intelligence
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# Load data function with caching
@st.cache_data
def load_data(file):
    """Load and preprocess Excel data with optimization"""
    try:
        df = pd.read_excel(file, engine='openpyxl')
        
        # Convert date columns
        date_columns = ['ORD CREATE', 'READY', 'QT PU', 'PICKUP DATE/TIME', 
                       'Depart Date / Time', 'Arrive Date / Time', 'QDT', 
                       'UPD DEL', 'POD DATE/TIME']
        
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Clean numeric columns
        numeric_cols = ['TOTAL CHARGES', 'PIECES', 'WEIGHT(KG)', 'Billable Weight KG']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df
    except Exception as e:
        st.error(f"⚠️ Error loading data: {str(e)}")
        return None

# Define controllable QC codes with categorization
CONTROLLABLE_QC_CATEGORIES = {
    'Internal Process': [
        'MNX-Incorrect QDT',
        'MNX-Order Entry error',
        'MNX-Late dispatch-Delivery',
        'W/House-Data entry errors'
    ],
    'Customs & Regulatory': [
        'Customs delay',
        'Customs delay-FDA Hold',
        'Customs-Late PWK-Customer'
    ],
    'Partner Operations': [
        'Del Agt-Late del',
        'Del Agt-Late del-Out of hours',
        'Del Agt-Missing documents',
        'PU Agt -Late pick up',
        'Airline-Slow offload',
        'Airline-RTA-DG PWK issue'
    ],
    'Operational': [
        'Shipment not ready'
    ]
}

# Flatten controllable codes for easy checking
CONTROLLABLE_QC_CODES = [code for codes in CONTROLLABLE_QC_CATEGORIES.values() for code in codes]

# Currency conversion
USD_TO_EUR = 0.92

def calculate_otp_metrics(df):
    """Calculate comprehensive OTP metrics with explanations"""
    otp_df = df[(df['QDT'].notna()) & (df['POD DATE/TIME'].notna())].copy()
    
    if len(otp_df) == 0:
        return {
            'gross_otp': 0,
            'net_otp': 0,
            'on_time_count': 0,
            'total_count': 0,
            'late_controllable': 0,
            'late_uncontrollable': 0
        }
    
    # Calculate on-time deliveries
    otp_df['on_time'] = otp_df['POD DATE/TIME'] <= otp_df['QDT']
    otp_df['is_controllable'] = otp_df['QC NAME'].isin(CONTROLLABLE_QC_CODES)
    
    # Gross OTP - ALL shipments
    gross_otp = (otp_df['on_time'].sum() / len(otp_df)) * 100
    
    # Late shipments analysis
    late_df = otp_df[~otp_df['on_time']]
    late_controllable = late_df['is_controllable'].sum()
    late_uncontrollable = len(late_df) - late_controllable
    
    # Net OTP - Excluding controllable issues
    net_otp_df = otp_df[~otp_df['is_controllable']]
    if len(net_otp_df) > 0:
        net_otp = (net_otp_df['on_time'].sum() / len(net_otp_df)) * 100
    else:
        net_otp = gross_otp
    
    return {
        'gross_otp': gross_otp,
        'net_otp': net_otp,
        'on_time_count': otp_df['on_time'].sum(),
        'total_count': len(otp_df),
        'late_controllable': late_controllable,
        'late_uncontrollable': late_uncontrollable
    }

def create_kpi_card(title, value, delta=None, delta_color="normal", prefix="", suffix=""):
    """Create a styled KPI card"""
    delta_html = ""
    if delta is not None:
        color = "#10b981" if delta_color == "normal" else "#ef4444"
        arrow = "↑" if delta > 0 else "↓"
        delta_html = f'<p style="color: {color}; font-size: 0.9rem; margin: 0;">{arrow} {abs(delta):.1f}%</p>'
    
    return f"""
    <div style='background: white; padding: 1.5rem; border-radius: 12px; border-left: 4px solid #3b82f6;'>
        <p style='color: #64748b; font-size: 0.9rem; margin: 0;'>{title}</p>
        <h2 style='color: #1e293b; font-size: 2rem; margin: 0.5rem 0;'>{prefix}{value}{suffix}</h2>
        {delta_html}
    </div>
    """

def main():
    # Sidebar configuration
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 1rem 0;'>
            <h3 style='color: white;'>📊 Data Configuration</h3>
        </div>
        """, unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "Upload Shipment Data",
            type=['xls', 'xlsx'],
            help="Upload your shipment Excel file"
        )
        
        if uploaded_file is not None:
            df = load_data(uploaded_file)
            
            if df is not None:
                # Filter for billed status
                df_billed = df[df['STATUS'] == '440-BILLED'].copy()
                
                st.success(f"✅ {len(df_billed):,} billed records loaded")
                
                # Date range filter
                st.markdown("### 📅 Time Period")
                if 'ORD CREATE' in df_billed.columns:
                    min_date = df_billed['ORD CREATE'].min()
                    max_date = df_billed['ORD CREATE'].max()
                    
                    date_range = st.date_input(
                        "Select range",
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
                st.markdown("### 🚚 Service Types")
                services = df_filtered['SVC'].dropna().unique()
                selected_services = st.multiselect(
                    "Filter by service",
                    services,
                    default=list(services)
                )
                
                if selected_services:
                    df_filtered = df_filtered[df_filtered['SVC'].isin(selected_services)]
                
                # Departure filter
                st.markdown("### ✈️ Departure Points")
                deps = df_filtered['DEP'].dropna().unique()
                top_deps = df_filtered['DEP'].value_counts().head(10).index.tolist()
                selected_deps = st.multiselect(
                    "Filter by departure",
                    deps,
                    default=top_deps
                )
                
                if selected_deps:
                    df_filtered = df_filtered[df_filtered['DEP'].isin(selected_deps)]
                
                st.info(f"📊 Analyzing {len(df_filtered):,} shipments")
    
    # Main dashboard
    if uploaded_file is not None and df is not None:
        # Convert charges to EUR
        df_filtered['TOTAL CHARGES EUR'] = df_filtered['TOTAL CHARGES'] * USD_TO_EUR
        
        # Calculate key metrics
        total_shipments = len(df_filtered)
        total_cost_eur = df_filtered['TOTAL CHARGES EUR'].sum()
        avg_cost_eur = df_filtered['TOTAL CHARGES EUR'].mean()
        total_weight = df_filtered['Billable Weight KG'].sum()
        
        otp_metrics = calculate_otp_metrics(df_filtered)
        gross_otp = otp_metrics['gross_otp']
        net_otp = otp_metrics['net_otp']
        
        qc_issues = df_filtered['QC NAME'].notna().sum()
        qc_rate = (qc_issues / total_shipments * 100) if total_shipments > 0 else 0
        
        # Executive Summary Section
        st.markdown("## 📊 Executive Summary")
        
        # Top KPIs
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with tab3:
            st.markdown("### 🌍 Geographic Intelligence & Route Analysis")
            
            # Geographic performance overview
            col1, col2 = st.columns(2)
            
            with col1:
                # Departure hub performance
                dep_perf = []
                for dep in df_filtered['DEP'].value_counts().head(15).index:
                    dep_df = df_filtered[df_filtered['DEP'] == dep]
                    otp = calculate_otp_metrics(dep_df)['gross_otp']
                    dep_perf.append({
                        'Hub': dep,
                        'Volume': len(dep_df),
                        'OTP': otp,
                        'Avg Cost': dep_df['TOTAL CHARGES EUR'].mean(),
                        'Total Revenue': dep_df['TOTAL CHARGES EUR'].sum()
                    })
                
                dep_perf_df = pd.DataFrame(dep_perf)
                
                # Create bubble chart for hub performance
                fig_hub = px.scatter(
                    dep_perf_df,
                    x='Volume',
                    y='OTP',
                    size='Total Revenue',
                    color='Avg Cost',
                    hover_data=['Hub'],
                    title='Hub Performance Analysis',
                    color_continuous_scale='RdYlGn_r',
                    labels={'Volume': 'Shipment Volume', 'OTP': 'OTP %'}
                )
                
                fig_hub.add_hline(y=85, line_dash="dash", line_color="red", 
                                 annotation_text="Target OTP")
                fig_hub.add_vline(x=dep_perf_df['Volume'].median(), line_dash="dash", 
                                 line_color="gray", opacity=0.5)
                
                for _, row in dep_perf_df.iterrows():
                    if row['OTP'] < 80 or row['OTP'] > 95:
                        fig_hub.add_annotation(
                            x=row['Volume'], y=row['OTP'],
                            text=row['Hub'], showarrow=True,
                            arrowhead=2, arrowsize=1, arrowwidth=1
                        )
                
                fig_hub.update_layout(height=400)
                st.plotly_chart(fig_hub, use_container_width=True)
            
            with col2:
                # Route network visualization
                route_data = df_filtered.groupby(['DEP', 'ARR']).agg({
                    'REFER': 'count',
                    'TOTAL CHARGES EUR': 'mean'
                }).reset_index()
                route_data = route_data.sort_values('REFER', ascending=False).head(20)
                route_data['Route'] = route_data['DEP'] + ' → ' + route_data['ARR']
                
                fig_route = go.Figure()
                
                # Create Sankey diagram for top routes
                fig_sankey = go.Figure(data=[go.Sankey(
                    node=dict(
                        pad=15,
                        thickness=20,
                        line=dict(color="black", width=0.5),
                        label=list(set(route_data['DEP'].tolist() + route_data['ARR'].tolist())),
                        color="#3b82f6"
                    ),
                    link=dict(
                        source=[list(set(route_data['DEP'].tolist() + route_data['ARR'].tolist())).index(x) 
                               for x in route_data['DEP']],
                        target=[list(set(route_data['DEP'].tolist() + route_data['ARR'].tolist())).index(x) 
                               for x in route_data['ARR']],
                        value=route_data['REFER'],
                        color=['rgba(59, 130, 246, 0.4)'] * len(route_data)
                    )
                )])
                
                fig_sankey.update_layout(
                    title="Top 20 Shipping Routes Flow",
                    height=400
                )
                st.plotly_chart(fig_sankey, use_container_width=True)
            
            # Geographic distribution map-style visualization
            st.markdown("#### 🗺️ Global Reach & Performance")
            
            # Country performance metrics
            country_metrics = df_filtered.groupby('DEL CTRY').agg({
                'REFER': 'count',
                'TOTAL CHARGES EUR': 'sum'
            }).reset_index()
            country_metrics.columns = ['Country', 'Shipments', 'Revenue']
            country_metrics = country_metrics.sort_values('Shipments', ascending=False)
            
            # Create treemap for country distribution
            fig_treemap = px.treemap(
                country_metrics.head(30),
                path=['Country'],
                values='Shipments',
                color='Revenue',
                title='Destination Countries by Volume and Revenue',
                color_continuous_scale='Viridis',
                hover_data={'Revenue': ':,.0f', 'Shipments': ':,'}
            )
            fig_treemap.update_layout(height=500)
            st.plotly_chart(fig_treemap, use_container_width=True)
            
            # Route profitability analysis
            col1, col2, col3 = st.columns(3)
            
            with col1:
                top_routes = route_data.head(5)
                st.markdown("**🏆 Top 5 Routes by Volume**")
                for _, route in top_routes.iterrows():
                    st.markdown(f"""
                    <div style='background: #f8f9fa; padding: 0.5rem; margin: 0.2rem 0; border-radius: 4px;'>
                        <strong>{route['Route']}</strong><br>
                        <span style='color: #64748b;'>{route['REFER']:,} shipments | €{route['TOTAL CHARGES EUR']:.0f} avg</span>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col2:
                expensive_routes = route_data.nlargest(5, 'TOTAL CHARGES EUR')
                st.markdown("**💰 Most Expensive Routes**")
                for _, route in expensive_routes.iterrows():
                    st.markdown(f"""
                    <div style='background: #fef3c7; padding: 0.5rem; margin: 0.2rem 0; border-radius: 4px;'>
                        <strong>{route['Route']}</strong><br>
                        <span style='color: #92400e;'>€{route['TOTAL CHARGES EUR']:.0f} average cost</span>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col3:
                # Calculate route OTP
                route_otp = []
                for _, route in route_data.head(10).iterrows():
                    route_df = df_filtered[(df_filtered['DEP'] == route['DEP']) & 
                                          (df_filtered['ARR'] == route['ARR'])]
                    if len(route_df) > 5:
                        otp = calculate_otp_metrics(route_df)['gross_otp']
                        route_otp.append({'Route': route['Route'], 'OTP': otp})
                
                if route_otp:
                    route_otp_df = pd.DataFrame(route_otp).sort_values('OTP')
                    st.markdown("**⚠️ Routes Needing Attention**")
                    for _, route in route_otp_df.head(5).iterrows():
                        color = '#ef4444' if route['OTP'] < 85 else '#f59e0b'
                        st.markdown(f"""
                        <div style='background: #fef2f2; padding: 0.5rem; margin: 0.2rem 0; border-radius: 4px;'>
                            <strong>{route['Route']}</strong><br>
                            <span style='color: {color};'>OTP: {route['OTP']:.1f}%</span>
                        </div>
                        """, unsafe_allow_html=True)
        
        with tab4:
            st.markdown("### 🚚 Service Analysis & Optimization")
            
            # Service performance dashboard
            svc_analysis = []
            for svc in df_filtered['SVC'].dropna().unique():
                svc_df = df_filtered[df_filtered['SVC'] == svc]
                if len(svc_df) > 5:
                    otp_data = calculate_otp_metrics(svc_df)
                    svc_analysis.append({
                        'Service': svc,
                        'Description': svc_df['SVCDESC'].mode()[0] if not svc_df['SVCDESC'].empty else svc,
                        'Volume': len(svc_df),
                        'Revenue': svc_df['TOTAL CHARGES EUR'].sum(),
                        'Avg Cost': svc_df['TOTAL CHARGES EUR'].mean(),
                        'OTP Gross': otp_data['gross_otp'],
                        'OTP Net': otp_data['net_otp'],
                        'QC Rate': (svc_df['QC NAME'].notna().sum() / len(svc_df)) * 100
                    })
            
            svc_df_analysis = pd.DataFrame(svc_analysis)
            
            # Service portfolio matrix
            col1, col2 = st.columns(2)
            
            with col1:
                # BCG Matrix style visualization
                fig_bcg = px.scatter(
                    svc_df_analysis,
                    x='Volume',
                    y='Revenue',
                    size='Avg Cost',
                    color='OTP Gross',
                    hover_data=['Service', 'Description'],
                    title='Service Portfolio Analysis',
                    color_continuous_scale='RdYlGn',
                    labels={'Volume': 'Shipment Volume', 'Revenue': 'Total Revenue (EUR)'}
                )
                
                # Add quadrant dividers
                fig_bcg.add_hline(y=svc_df_analysis['Revenue'].median(), 
                                 line_dash="dash", line_color="gray", opacity=0.3)
                fig_bcg.add_vline(x=svc_df_analysis['Volume'].median(), 
                                 line_dash="dash", line_color="gray", opacity=0.3)
                
                # Add quadrant labels
                fig_bcg.add_annotation(x=svc_df_analysis['Volume'].max()*0.8, 
                                      y=svc_df_analysis['Revenue'].max()*0.9,
                                      text="⭐ Stars<br>High Volume, High Revenue",
                                      showarrow=False, font=dict(size=10))
                fig_bcg.add_annotation(x=svc_df_analysis['Volume'].min()*1.5, 
                                      y=svc_df_analysis['Revenue'].max()*0.9,
                                      text="❓ Question Marks<br>Low Volume, High Revenue",
                                      showarrow=False, font=dict(size=10))
                
                fig_bcg.update_layout(height=450)
                st.plotly_chart(fig_bcg, use_container_width=True)
            
            with col2:
                # Service performance radar
                top_5_services = svc_df_analysis.nlargest(5, 'Volume')
                
                categories = ['Volume Score', 'Revenue Score', 'OTP Score', 'Quality Score', 'Efficiency Score']
                
                fig_radar = go.Figure()
                
                for _, service in top_5_services.iterrows():
                    scores = [
                        (service['Volume'] / svc_df_analysis['Volume'].max()) * 100,
                        (service['Revenue'] / svc_df_analysis['Revenue'].max()) * 100,
                        service['OTP Gross'],
                        100 - service['QC Rate'],
                        (1 - (service['Avg Cost'] / svc_df_analysis['Avg Cost'].max())) * 100
                    ]
                    
                    fig_radar.add_trace(go.Scatterpolar(
                        r=scores,
                        theta=categories,
                        fill='toself',
                        name=service['Service'],
                        hovertemplate='%{theta}: %{r:.1f}<extra></extra>'
                    ))
                
                fig_radar.update_layout(
                    polar=dict(
                        radialaxis=dict(
                            visible=True,
                            range=[0, 100]
                        )
                    ),
                    title="Top 5 Services Performance Scorecard",
                    height=450
                )
                st.plotly_chart(fig_radar, use_container_width=True)
            
            # Service comparison table
            st.markdown("#### 📊 Service Performance Metrics")
            
            # Create styled dataframe
            styled_df = svc_df_analysis.sort_values('Volume', ascending=False)
            styled_df['Volume Rank'] = range(1, len(styled_df) + 1)
            styled_df['Performance'] = styled_df.apply(
                lambda x: '🟢' if x['OTP Gross'] >= 90 else ('🟡' if x['OTP Gross'] >= 85 else '🔴'), 
                axis=1
            )
            
            display_df = styled_df[['Volume Rank', 'Performance', 'Service', 'Description', 
                                   'Volume', 'Revenue', 'Avg Cost', 'OTP Gross', 'OTP Net', 'QC Rate']]
            
            st.dataframe(
                display_df.style.format({
                    'Revenue': '€{:,.0f}',
                    'Avg Cost': '€{:,.0f}',
                    'OTP Gross': '{:.1f}%',
                    'OTP Net': '{:.1f}%',
                    'QC Rate': '{:.1f}%'
                }).background_gradient(subset=['OTP Gross', 'OTP Net'], cmap='RdYlGn', vmin=70, vmax=100)
                .background_gradient(subset=['QC Rate'], cmap='RdYlGn_r', vmin=0, vmax=20),
                use_container_width=True,
                height=400
            )
            
            # Service insights
            col1, col2, col3 = st.columns(3)
            
            with col1:
                best_otp = styled_df.nlargest(1, 'OTP Gross').iloc[0]
                st.success(f"""
                **🏆 Best Performing Service**  
                {best_otp['Service']}: {best_otp['OTP Gross']:.1f}% OTP
                """)
            
            with col2:
                highest_revenue = styled_df.nlargest(1, 'Revenue').iloc[0]
                st.info(f"""
                **💰 Highest Revenue Service**  
                {highest_revenue['Service']}: €{highest_revenue['Revenue']:,.0f}
                """)
            
            with col3:
                needs_attention = styled_df[styled_df['OTP Gross'] < 85]
                if not needs_attention.empty:
                    st.warning(f"""
                    **⚠️ Services Below Target**  
                    {len(needs_attention)} services below 85% OTP
                    """)
        
        with tab5:
            st.markdown("### ⚠️ Quality Control & Improvement Opportunities")
            
            # QC Overview cards
            qc_data = df_filtered[df_filtered['QC NAME'].notna()]
            controllable_qc = qc_data[qc_data['QC NAME'].isin(CONTROLLABLE_QC_CODES)]
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(f"""
                <div style='background: #fee2e2; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #991b1b; margin: 0;'>Total QC Issues</h4>
                    <h2 style='color: #dc2626; margin: 0.5rem 0;'>{len(qc_data):,}</h2>
                    <p style='color: #7f1d1d; font-size: 0.9rem; margin: 0;'>
                        {(len(qc_data)/len(df_filtered)*100):.1f}% of shipments
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div style='background: #fef3c7; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #78350f; margin: 0;'>Controllable</h4>
                    <h2 style='color: #f59e0b; margin: 0.5rem 0;'>{len(controllable_qc):,}</h2>
                    <p style='color: #92400e; font-size: 0.9rem; margin: 0;'>
                        {(len(controllable_qc)/len(qc_data)*100):.1f}% can be fixed
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                uncontrollable = len(qc_data) - len(controllable_qc)
                st.markdown(f"""
                <div style='background: #dbeafe; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #1e3a8a; margin: 0;'>Uncontrollable</h4>
                    <h2 style='color: #3b82f6; margin: 0.5rem 0;'>{uncontrollable:,}</h2>
                    <p style='color: #1e40af; font-size: 0.9rem; margin: 0;'>
                        Customer-driven issues
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                potential_improvement = (len(controllable_qc) / len(df_filtered)) * 100
                st.markdown(f"""
                <div style='background: #d1fae5; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #064e3b; margin: 0;'>Potential OTP Gain</h4>
                    <h2 style='color: #10b981; margin: 0.5rem 0;'>+{potential_improvement:.1f}%</h2>
                    <p style='color: #047857; font-size: 0.9rem; margin: 0;'>
                        If controllable fixed
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            # QC Categories breakdown
            st.markdown("#### 🔍 Quality Issue Categories")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Controllable categories breakdown
                controllable_breakdown = []
                for category, codes in CONTROLLABLE_QC_CATEGORIES.items():
                    count = qc_data[qc_data['QC NAME'].isin(codes)].shape[0]
                    if count > 0:
                        controllable_breakdown.append({
                            'Category': category,
                            'Issues': count,
                            'Percentage': (count / len(qc_data)) * 100
                        })
                
                if controllable_breakdown:
                    ctrl_df = pd.DataFrame(controllable_breakdown)
                    
                    fig_ctrl = px.pie(
                        ctrl_df,
                        values='Issues',
                        names='Category',
                        title='Controllable Issues by Category',
                        color_discrete_sequence=px.colors.sequential.OrRd
                    )
                    fig_ctrl.update_traces(textposition='inside', textinfo='percent+label')
                    fig_ctrl.update_layout(height=350)
                    st.plotly_chart(fig_ctrl, use_container_width=True)
            
            with col2:
                # Top QC issues
                qc_counts = qc_data['QC NAME'].value_counts().head(10)
                
                fig_top_qc = go.Figure(go.Bar(
                    x=qc_counts.values,
                    y=qc_counts.index,
                    orientation='h',
                    marker_color=['#ef4444' if issue in CONTROLLABLE_QC_CODES else '#3b82f6' 
                                 for issue in qc_counts.index],
                    text=qc_counts.values,
                    textposition='outside'
                ))
                
                fig_top_qc.update_layout(
                    title='Top 10 Quality Issues',
                    xaxis_title='Number of Occurrences',
                    yaxis_title='',
                    height=350,
                    showlegend=False
                )
                st.plotly_chart(fig_top_qc, use_container_width=True)
            
            # QC Trend Analysis
            st.markdown("#### 📈 Quality Trends & Patterns")
            
            if 'ORD CREATE' in df_filtered.columns:
                # Weekly QC trend
                df_qc_trend = df_filtered.copy()
                df_qc_trend['Week'] = df_qc_trend['ORD CREATE'].dt.to_period('W')
                
                weekly_qc = df_qc_trend.groupby('Week').apply(
                    lambda x: pd.Series({
                        'Total': len(x),
                        'QC Issues': x['QC NAME'].notna().sum(),
                        'Controllable': x[x['QC NAME'].isin(CONTROLLABLE_QC_CODES)]['QC NAME'].notna().sum(),
                        'QC Rate': (x['QC NAME'].notna().sum() / len(x)) * 100
                    })
                ).reset_index()
                weekly_qc['Week'] = weekly_qc['Week'].astype(str)
                
                fig_qc_trend = make_subplots(
                    rows=2, cols=1,
                    subplot_titles=('QC Issues Over Time', 'QC Rate Trend (%)'),
                    row_heights=[0.6, 0.4]
                )
                
                # Stacked bar chart for QC issues
                fig_qc_trend.add_trace(
                    go.Bar(name='Controllable', x=weekly_qc['Week'], y=weekly_qc['Controllable'],
                          marker_color='#f59e0b'),
                    row=1, col=1
                )
                fig_qc_trend.add_trace(
                    go.Bar(name='Uncontrollable', x=weekly_qc['Week'], 
                          y=weekly_qc['QC Issues'] - weekly_qc['Controllable'],
                          marker_color='#3b82f6'),
                    row=1, col=1
                )
                
                # QC rate trend line
                fig_qc_trend.add_trace(
                    go.Scatter(x=weekly_qc['Week'], y=weekly_qc['QC Rate'],
                              mode='lines+markers', name='QC Rate',
                              line=dict(color='#ef4444', width=3)),
                    row=2, col=1
                )
                fig_qc_trend.add_hline(y=10, line_dash="dash", line_color="green",
                                      annotation_text="Target <10%", row=2, col=1)
                
                fig_qc_trend.update_layout(
                    height=600,
                    barmode='stack',
                    showlegend=True,
                    hovermode='x unified'
                )
                st.plotly_chart(fig_qc_trend, use_container_width=True)
            
            # Action items
            st.markdown("#### 🎯 Recommended Actions")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                <div style='background: #fef3c7; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #78350f;'>🔧 Internal Process Improvements</h4>
                    <ul style='color: #92400e; margin: 0.5rem 0;'>
                        <li>Implement automated QDT validation system</li>
                        <li>Enhance order entry training program</li>
                        <li>Deploy real-time dispatch monitoring</li>
                        <li>Upgrade warehouse data systems</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown("""
                <div style='background: #dbeafe; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #1e3a8a;'>🤝 Partner & Customer Actions</h4>
                    <ul style='color: #1e40af; margin: 0.5rem 0;'>
                        <li>Review SLAs with delivery agents</li>
                        <li>Implement partner scorecards</li>
                        <li>Enhance customer communication</li>
                        <li>Develop customs pre-clearance process</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            
            # ROI calculation
            st.markdown("#### 💰 Improvement ROI Estimation")
            
            avg_delay_cost = df_filtered[df_filtered['QC NAME'].isin(CONTROLLABLE_QC_CODES)]['TOTAL CHARGES EUR'].mean()
            total_controllable_cost = len(controllable_qc) * avg_delay_cost * 0.1  # Assume 10% cost impact
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(
                    "Estimated Annual Impact",
                    f"€{total_controllable_cost*12:,.0f}",
                    help="Potential cost savings from fixing controllable issues"
                )
            
            with col2:
                st.metric(
                    "OTP Improvement Potential",
                    f"+{(len(controllable_qc)/len(df_filtered)*100):.1f}%",
                    help="Potential OTP increase"
                )
            
            with col3:
                st.metric(
                    "Customer Satisfaction Impact",
                    "High",
                    help="Expected improvement in customer satisfaction"
                )
    
    else:
        # Landing page when no data is loaded
        st.markdown("""
        <div style='text-align: center; padding: 3rem;'>
            <h2 style='color: #1e293b;'>Welcome to the Executive Analytics Dashboard</h2>
            <p style='color: #64748b; font-size: 1.1rem; margin: 2rem 0;'>
                Upload your shipment data to unlock powerful insights and drive strategic decisions.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div style='background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>
                <h3 style='color: #3b82f6;'>📈 Performance Metrics</h3>
                <p style='color: #64748b;'>Track OTP, costs, and quality metrics with real-time visualizations</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div style='background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>
                <h3 style='color: #10b981;'>🎯 Strategic Insights</h3>
                <p style='color: #64748b;'>Identify optimization opportunities and cost-saving potential</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div style='background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>
                <h3 style='color: #f59e0b;'>🔍 Root Cause Analysis</h3>
                <p style='color: #64748b;'>Understand quality issues and their business impact</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("""
        <div style='margin-top: 3rem; padding: 2rem; background: #f8f9fa; border-radius: 12px;'>
            <h3 style='color: #1e293b;'>📁 Required Data Format</h3>
            <p style='color: #64748b;'>Your Excel file should contain the following columns:</p>
            <ul style='color: #64748b;'>
                <li><strong>STATUS</strong> - Must include '440-BILLED' records</li>
                <li><strong>Order & Shipment Details</strong> - Reference numbers, dates, routes</li>
                <li><strong>Cost Information</strong> - Total charges in original currency</li>
                <li><strong>Performance Data</strong> - QDT, POD dates for OTP calculation</li>
                <li><strong>Quality Metrics</strong> - QC codes and descriptions</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()

    st.metric(
        label="Total Shipments",
        value=f"{total_shipments:,}",
        help="Total billed shipments in period"
    )
        
        with col2:
            st.metric(
                "Total Cost",
                f"€{total_cost_eur/1000:.1f}K",
                f"{((total_cost_eur/df_billed['TOTAL CHARGES EUR'].sum())-1)*100:.1f}%",
                help="Total shipment costs in EUR"
            )
        
        with col3:
            st.metric(
                "Avg Cost/Shipment",
                f"€{avg_cost_eur:.0f}",
                help="Average cost per shipment"
            )
        
        with col4:
            st.metric(
                "OTP Performance",
                f"{gross_otp:.1f}%",
                f"{gross_otp - 85:.1f}%",
                delta_color="normal" if gross_otp >= 85 else "inverse",
                help="On-Time Performance (target: 85%)"
            )
        
        with col5:
            st.metric(
                "QC Issue Rate",
                f"{qc_rate:.1f}%",
                f"{qc_rate - 10:.1f}%",
                delta_color="inverse" if qc_rate > 10 else "normal",
                help="Quality control issues (target: <10%)"
            )
        
        # OTP Explanation Box
        with st.expander("🎯 **Understanding OTP Metrics** - Click to expand", expanded=True):
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col1:
                st.markdown("""
                <div style='background: #eff6ff; padding: 1rem; border-radius: 8px; border-left: 4px solid #3b82f6;'>
                    <h4 style='color: #1e3a8a; margin: 0;'>📊 OTP Gross</h4>
                    <p style='color: #1e293b; margin: 0.5rem 0; font-size: 0.9rem;'>
                        Measures ALL shipments delivered on-time, including those with issues.
                        This is your overall performance metric.
                    </p>
                    <h3 style='color: #3b82f6; margin: 0;'>{:.1f}%</h3>
                    <p style='color: #64748b; font-size: 0.8rem;'>Target: ≥85%</p>
                </div>
                """.format(gross_otp), unsafe_allow_html=True)
            
            with col2:
                st.markdown("""
                <div style='background: #f0fdf4; padding: 1rem; border-radius: 8px; border-left: 4px solid #10b981;'>
                    <h4 style='color: #14532d; margin: 0;'>🎯 OTP Net</h4>
                    <p style='color: #1e293b; margin: 0.5rem 0; font-size: 0.9rem;'>
                        Excludes delays we can control (customs, warehouse, partners).
                        Shows true customer-impacting performance.
                    </p>
                    <h3 style='color: #10b981; margin: 0;'>{:.1f}%</h3>
                    <p style='color: #64748b; font-size: 0.8rem;'>Target: ≥90%</p>
                </div>
                """.format(net_otp), unsafe_allow_html=True)
            
            with col3:
                improvement_potential = net_otp - gross_otp
                st.markdown(f"""
                <div style='background: #fef3c7; padding: 1rem; border-radius: 8px; border-left: 4px solid #f59e0b;'>
                    <h4 style='color: #78350f; margin: 0;'>💡 Improvement Potential</h4>
                    <p style='color: #1e293b; margin: 0.5rem 0; font-size: 0.9rem;'>
                        By fixing controllable issues, OTP could improve by:
                    </p>
                    <h3 style='color: #f59e0b; margin: 0;'>+{improvement_potential:.1f}%</h3>
                    <p style='color: #64748b; font-size: 0.8rem;'>
                        {otp_metrics['late_controllable']} controllable delays
                    </p>
                </div>
                """, unsafe_allow_html=True)
        
        # Tabs for detailed analysis
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📈 Performance Trends",
            "💰 Cost Analytics", 
            "🌍 Geographic Intelligence",
            "🚚 Service Analysis",
            "⚠️ Quality Control"
        ])
        
        with tab1:
            st.markdown("### 📈 Performance Trends & Insights")
            
            # Monthly trend analysis
            if 'ORD CREATE' in df_filtered.columns:
                df_monthly = df_filtered.copy()
                df_monthly['Month'] = df_monthly['ORD CREATE'].dt.to_period('M')
                
                # Calculate monthly metrics
                monthly_metrics = df_monthly.groupby('Month').apply(
                    lambda x: pd.Series({
                        'Shipments': len(x),
                        'Cost (EUR)': x['TOTAL CHARGES EUR'].sum(),
                        'Avg Cost': x['TOTAL CHARGES EUR'].mean(),
                        'OTP': calculate_otp_metrics(x)['gross_otp'],
                        'QC Rate': (x['QC NAME'].notna().sum() / len(x)) * 100
                    })
                ).reset_index()
                monthly_metrics['Month'] = monthly_metrics['Month'].astype(str)
                
                # Create comprehensive trend chart
                fig = make_subplots(
                    rows=2, cols=2,
                    subplot_titles=(
                        'Shipment Volume Trend',
                        'Cost Trend (EUR)',
                        'OTP Performance Trend',
                        'Quality Issues Trend'
                    ),
                    specs=[[{'secondary_y': False}, {'secondary_y': True}],
                           [{'secondary_y': False}, {'secondary_y': False}]]
                )
                
                # Shipment volume
                fig.add_trace(
                    go.Bar(x=monthly_metrics['Month'], y=monthly_metrics['Shipments'],
                          name='Shipments', marker_color='#3b82f6',
                          text=monthly_metrics['Shipments'],
                          textposition='outside'),
                    row=1, col=1
                )
                
                # Cost trend with average line
                fig.add_trace(
                    go.Bar(x=monthly_metrics['Month'], y=monthly_metrics['Cost (EUR)'],
                          name='Total Cost', marker_color='#10b981'),
                    row=1, col=2, secondary_y=False
                )
                fig.add_trace(
                    go.Scatter(x=monthly_metrics['Month'], y=monthly_metrics['Avg Cost'],
                              name='Avg Cost', line=dict(color='#ef4444', width=3),
                              mode='lines+markers'),
                    row=1, col=2, secondary_y=True
                )
                
                # OTP trend with target line
                fig.add_trace(
                    go.Scatter(x=monthly_metrics['Month'], y=monthly_metrics['OTP'],
                              name='OTP %', line=dict(color='#3b82f6', width=3),
                              mode='lines+markers+text',
                              text=[f"{x:.1f}%" for x in monthly_metrics['OTP']],
                              textposition='top center'),
                    row=2, col=1
                )
                fig.add_hline(y=85, line_dash="dash", line_color="red", 
                            annotation_text="Target: 85%", row=2, col=1)
                
                # QC trend
                fig.add_trace(
                    go.Scatter(x=monthly_metrics['Month'], y=monthly_metrics['QC Rate'],
                              name='QC Rate', line=dict(color='#f59e0b', width=3),
                              mode='lines+markers', fill='tozeroy',
                              fillcolor='rgba(245, 158, 11, 0.2)'),
                    row=2, col=2
                )
                fig.add_hline(y=10, line_dash="dash", line_color="red",
                            annotation_text="Target: <10%", row=2, col=2)
                
                fig.update_layout(height=700, showlegend=False, 
                                title_text="Monthly Performance Dashboard")
                fig.update_yaxes(title_text="Avg Cost (EUR)", secondary_y=True, row=1, col=2)
                st.plotly_chart(fig, use_container_width=True)
            
            # Performance by day of week
            col1, col2 = st.columns(2)
            
            with col1:
                if 'ORD CREATE' in df_filtered.columns:
                    df_filtered['Weekday'] = df_filtered['ORD CREATE'].dt.day_name()
                    weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 
                                   'Friday', 'Saturday', 'Sunday']
                    weekday_stats = df_filtered.groupby('Weekday').agg({
                        'REFER': 'count',
                        'TOTAL CHARGES EUR': 'mean'
                    }).reindex(weekday_order)
                    
                    fig_weekday = go.Figure()
                    fig_weekday.add_trace(go.Bar(
                        x=weekday_stats.index,
                        y=weekday_stats['REFER'],
                        name='Volume',
                        marker_color=['#3b82f6' if d not in ['Saturday', 'Sunday'] 
                                    else '#94a3b8' for d in weekday_stats.index],
                        text=weekday_stats['REFER'],
                        textposition='outside'
                    ))
                    fig_weekday.update_layout(
                        title='Shipment Volume by Day of Week',
                        xaxis_title='Day',
                        yaxis_title='Number of Shipments',
                        height=350
                    )
                    st.plotly_chart(fig_weekday, use_container_width=True)
            
            with col2:
                # OTP by hour of day
                if 'POD DATE/TIME' in df_filtered.columns:
                    df_filtered['Delivery Hour'] = df_filtered['POD DATE/TIME'].dt.hour
                    hourly_otp = df_filtered.groupby('Delivery Hour').apply(
                        lambda x: calculate_otp_metrics(x)['gross_otp']
                    )
                    
                    fig_hourly = go.Figure()
                    fig_hourly.add_trace(go.Scatter(
                        x=hourly_otp.index,
                        y=hourly_otp.values,
                        mode='lines+markers',
                        line=dict(color='#10b981', width=3),
                        fill='tozeroy',
                        fillcolor='rgba(16, 185, 129, 0.2)',
                        name='OTP %'
                    ))
                    fig_hourly.add_hline(y=85, line_dash="dash", line_color="red")
                    fig_hourly.update_layout(
                        title='OTP Performance by Hour of Day',
                        xaxis_title='Hour',
                        yaxis_title='OTP %',
                        height=350
                    )
                    st.plotly_chart(fig_hourly, use_container_width=True)
        
        with tab2:
            st.markdown("### 💰 Cost Analytics & Optimization")
            
            # Cost breakdown analysis
            col1, col2 = st.columns(2)
            
            with col1:
                # Cost by service type - Sunburst chart
                svc_cost = df_filtered.groupby(['SVC', 'SVCDESC']).agg({
                    'TOTAL CHARGES EUR': 'sum',
                    'REFER': 'count'
                }).reset_index()
                svc_cost['Cost per Shipment'] = svc_cost['TOTAL CHARGES EUR'] / svc_cost['REFER']
                
                fig_sunburst = px.sunburst(
                    svc_cost,
                    path=['SVC', 'SVCDESC'],
                    values='TOTAL CHARGES EUR',
                    color='Cost per Shipment',
                    color_continuous_scale='RdYlGn_r',
                    title='Cost Distribution by Service Type'
                )
                fig_sunburst.update_layout(height=400)
                st.plotly_chart(fig_sunburst, use_container_width=True)
            
            with col2:
                # Cost efficiency matrix
                efficiency_data = []
                for svc in df_filtered['SVC'].dropna().unique():
                    svc_df = df_filtered[df_filtered['SVC'] == svc]
                    if len(svc_df) > 5:  # Only include services with sufficient data
                        otp = calculate_otp_metrics(svc_df)['gross_otp']
                        avg_cost = svc_df['TOTAL CHARGES EUR'].mean()
                        efficiency_data.append({
                            'Service': svc,
                            'Avg Cost': avg_cost,
                            'OTP': otp,
                            'Volume': len(svc_df),
                            'Total Cost': svc_df['TOTAL CHARGES EUR'].sum()
                        })
                
                if efficiency_data:
                    eff_df = pd.DataFrame(efficiency_data)
                    
                    fig_efficiency = px.scatter(
                        eff_df,
                        x='Avg Cost',
                        y='OTP',
                        size='Volume',
                        color='Total Cost',
                        hover_data=['Service', 'Volume'],
                        title='Service Efficiency Matrix (Cost vs Performance)',
                        color_continuous_scale='Viridis'
                    )
                    
                    # Add quadrant lines
                    fig_efficiency.add_hline(y=85, line_dash="dash", line_color="gray", opacity=0.5)
                    fig_efficiency.add_vline(x=eff_df['Avg Cost'].median(), line_dash="dash", 
                                            line_color="gray", opacity=0.5)
                    
                    # Add quadrant labels
                    fig_efficiency.add_annotation(x=eff_df['Avg Cost'].min(), y=95,
                                                 text="High Performance<br>Low Cost", showarrow=False)
                    fig_efficiency.add_annotation(x=eff_df['Avg Cost'].max(), y=95,
                                                 text="High Performance<br>High Cost", showarrow=False)
                    
                    fig_efficiency.update_layout(height=400)
                    st.plotly_chart(fig_efficiency, use_container_width=True)
            
            # Cost drivers analysis
            st.markdown("#### 📊 Cost Drivers Analysis")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Weight vs Cost correlation
                weight_cost = df_filtered[['Billable Weight KG', 'TOTAL CHARGES EUR']].dropna()
                if len(weight_cost) > 0:
                    correlation = weight_cost.corr().iloc[0, 1]
                    
                    fig_weight = px.scatter(
                        weight_cost.sample(min(1000, len(weight_cost))),
                        x='Billable Weight KG',
                        y='TOTAL CHARGES EUR',
                        title=f'Weight Impact on Cost (r={correlation:.2f})',
                        trendline='ols',
                        color_discrete_sequence=['#3b82f6']
                    )
                    fig_weight.update_layout(height=300)
                    st.plotly_chart(fig_weight, use_container_width=True)
            
            with col2:
                # Distance vs Cost
                if 'TOT DST' in df_filtered.columns:
                    dist_cost = df_filtered[['TOT DST', 'TOTAL CHARGES EUR']].dropna()
                    dist_cost = dist_cost[dist_cost['TOT DST'] > 0]
                    
                    if len(dist_cost) > 0:
                        fig_dist = px.box(
                            dist_cost,
                            y='TOTAL CHARGES EUR',
                            title='Cost Distribution Analysis',
                            color_discrete_sequence=['#10b981']
                        )
                        fig_dist.update_layout(height=300)
                        st.plotly_chart(fig_dist, use_container_width=True)
            
            with col3:
                # Cost by destination country
                country_cost = df_filtered.groupby('DEL CTRY')['TOTAL CHARGES EUR'].agg(['mean', 'count'])
                country_cost = country_cost[country_cost['count'] > 5].sort_values('mean', ascending=False).head(10)
                
                fig_country = go.Figure(go.Bar(
                    x=country_cost['mean'],
                    y=country_cost.index,
                    orientation='h',
                    marker_color='#f59e0b',
                    text=[f"€{x:.0f}" for x in country_cost['mean']],
                    textposition='outside'
                ))
                fig_country.update_layout(
                    title='Top 10 Most Expensive Destinations',
                    xaxis_title='Average Cost (EUR)',
                    yaxis_title='Country',
                    height=300
                )
                st.plotly_chart(fig_country, use_container_width=True)
        
        with tab3:
            st.markdown("### 🌍 Geographic Intelligence & Route Analysis")
            
            # Geographic performance overview
            col1, col2 = st.columns(2)
            
            with col1:
                # Departure hub performance
                dep_perf = []
                for dep in df_filtered['DEP'].value_counts().head(15).index:
                    dep_df = df_filtered[df_filtered['DEP'] == dep]
                    otp = calculate_otp_metrics(dep_df)['gross_otp']
                    dep_perf.append({
                        'Hub': dep,
                        'Volume': len(dep_df),
                        'OTP': otp,
                        'Avg Cost': dep_df['TOTAL CHARGES EUR'].mean(),
                        'Total Revenue': dep_df['TOTAL CHARGES EUR'].sum()
                    })
                
                dep_perf_df = pd.DataFrame(dep_perf)
                
                # Create bubble chart for hub performance
                fig_hub = px.scatter(
                    dep_perf_df,
                    x='Volume',
                    y='OTP',
                    size='Total Revenue',
                    color='Avg Cost',
                    hover_data=['Hub'],
                    title='Hub Performance Analysis',
                    color_continuous_scale='RdYlGn_r',
                    labels={'Volume': 'Shipment Volume', 'OTP': 'OTP %'}
                )
                
                fig_hub.add_hline(y=85, line_dash="dash", line_color="red", 
                                 annotation_text="Target OTP")
                fig_hub.add_vline(x=dep_perf_df['Volume'].median(), line_dash="dash", 
                                 line_color="gray", opacity=0.5)
                
                for _, row in dep_perf_df.iterrows():
                    if row['OTP'] < 80 or row['OTP'] > 95:
                        fig_hub.add_annotation(
                            x=row['Volume'], y=row['OTP'],
                            text=row['Hub'], showarrow=True,
                            arrowhead=2, arrowsize=1, arrowwidth=1
                        )
                
                fig_hub.update_layout(height=400)
                st.plotly_chart(fig_hub, use_container_width=True)
            
            with col2:
                # Route network visualization
                route_data = df_filtered.groupby(['DEP', 'ARR']).agg({
                    'REFER': 'count',
                    'TOTAL CHARGES EUR': 'mean'
                }).reset_index()
                route_data = route_data.sort_values('REFER', ascending=False).head(20)
                route_data['Route'] = route_data['DEP'] + ' → ' + route_data['ARR']
                
                # Create Sankey diagram for top routes
                all_nodes = list(set(route_data['DEP'].tolist() + route_data['ARR'].tolist()))
                
                fig_sankey = go.Figure(data=[go.Sankey(
                    node=dict(
                        pad=15,
                        thickness=20,
                        line=dict(color="black", width=0.5),
                        label=all_nodes,
                        color="#3b82f6"
                    ),
                    link=dict(
                        source=[all_nodes.index(x) for x in route_data['DEP']],
                        target=[all_nodes.index(x) for x in route_data['ARR']],
                        value=route_data['REFER'],
                        color=['rgba(59, 130, 246, 0.4)'] * len(route_data)
                    )
                )])
                
                fig_sankey.update_layout(
                    title="Top 20 Shipping Routes Flow",
                    height=400
                )
                st.plotly_chart(fig_sankey, use_container_width=True)
            
            # Geographic distribution map-style visualization
            st.markdown("#### 🗺️ Global Reach & Performance")
            
            # Country performance metrics
            country_metrics = df_filtered.groupby('DEL CTRY').agg({
                'REFER': 'count',
                'TOTAL CHARGES EUR': 'sum'
            }).reset_index()
            country_metrics.columns = ['Country', 'Shipments', 'Revenue']
            country_metrics = country_metrics.sort_values('Shipments', ascending=False)
            
            # Create treemap for country distribution
            fig_treemap = px.treemap(
                country_metrics.head(30),
                path=['Country'],
                values='Shipments',
                color='Revenue',
                title='Destination Countries by Volume and Revenue',
                color_continuous_scale='Viridis',
                hover_data={'Revenue': ':,.0f', 'Shipments': ':,'}
            )
            fig_treemap.update_layout(height=500)
            st.plotly_chart(fig_treemap, use_container_width=True)
            
            # Route profitability analysis
            col1, col2, col3 = st.columns(3)
            
            with col1:
                top_routes = route_data.head(5)
                st.markdown("**🏆 Top 5 Routes by Volume**")
                for _, route in top_routes.iterrows():
                    st.markdown(f"""
                    <div style='background: #f8f9fa; padding: 0.5rem; margin: 0.2rem 0; border-radius: 4px;'>
                        <strong>{route['Route']}</strong><br>
                        <span style='color: #64748b;'>{route['REFER']:,} shipments | €{route['TOTAL CHARGES EUR']:.0f} avg</span>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col2:
                expensive_routes = route_data.nlargest(5, 'TOTAL CHARGES EUR')
                st.markdown("**💰 Most Expensive Routes**")
                for _, route in expensive_routes.iterrows():
                    st.markdown(f"""
                    <div style='background: #fef3c7; padding: 0.5rem; margin: 0.2rem 0; border-radius: 4px;'>
                        <strong>{route['Route']}</strong><br>
                        <span style='color: #92400e;'>€{route['TOTAL CHARGES EUR']:.0f} average cost</span>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col3:
                # Calculate route OTP
                route_otp = []
                for _, route in route_data.head(10).iterrows():
                    route_df = df_filtered[(df_filtered['DEP'] == route['DEP']) & 
                                          (df_filtered['ARR'] == route['ARR'])]
                    if len(route_df) > 5:
                        otp = calculate_otp_metrics(route_df)['gross_otp']
                        route_otp.append({'Route': route['Route'], 'OTP': otp})
                
                if route_otp:
                    route_otp_df = pd.DataFrame(route_otp).sort_values('OTP')
                    st.markdown("**⚠️ Routes Needing Attention**")
                    for _, route in route_otp_df.head(5).iterrows():
                        color = '#ef4444' if route['OTP'] < 85 else '#f59e0b'
                        st.markdown(f"""
                        <div style='background: #fef2f2; padding: 0.5rem; margin: 0.2rem 0; border-radius: 4px;'>
                            <strong>{route['Route']}</strong><br>
                            <span style='color: {color};'>OTP: {route['OTP']:.1f}%</span>
                        </div>
                        """, unsafe_allow_html=True)
        
        with tab4:
            st.markdown("### 🚚 Service Analysis & Optimization")
            
            # Service performance dashboard
            svc_analysis = []
            for svc in df_filtered['SVC'].dropna().unique():
                svc_df = df_filtered[df_filtered['SVC'] == svc]
                if len(svc_df) > 5:
                    otp_data = calculate_otp_metrics(svc_df)
                    svc_desc = svc_df['SVCDESC'].mode()[0] if 'SVCDESC' in svc_df.columns and not svc_df['SVCDESC'].empty else svc
                    svc_analysis.append({
                        'Service': svc,
                        'Description': svc_desc,
                        'Volume': len(svc_df),
                        'Revenue': svc_df['TOTAL CHARGES EUR'].sum(),
                        'Avg Cost': svc_df['TOTAL CHARGES EUR'].mean(),
                        'OTP Gross': otp_data['gross_otp'],
                        'OTP Net': otp_data['net_otp'],
                        'QC Rate': (svc_df['QC NAME'].notna().sum() / len(svc_df)) * 100
                    })
            
            if svc_analysis:
                svc_df_analysis = pd.DataFrame(svc_analysis)
                
                # Service portfolio matrix
                col1, col2 = st.columns(2)
                
                with col1:
                    # BCG Matrix style visualization
                    fig_bcg = px.scatter(
                        svc_df_analysis,
                        x='Volume',
                        y='Revenue',
                        size='Avg Cost',
                        color='OTP Gross',
                        hover_data=['Service', 'Description'],
                        title='Service Portfolio Analysis',
                        color_continuous_scale='RdYlGn',
                        labels={'Volume': 'Shipment Volume', 'Revenue': 'Total Revenue (EUR)'}
                    )
                    
                    # Add quadrant dividers
                    fig_bcg.add_hline(y=svc_df_analysis['Revenue'].median(), 
                                     line_dash="dash", line_color="gray", opacity=0.3)
                    fig_bcg.add_vline(x=svc_df_analysis['Volume'].median(), 
                                     line_dash="dash", line_color="gray", opacity=0.3)
                    
                    # Add quadrant labels
                    fig_bcg.add_annotation(x=svc_df_analysis['Volume'].max()*0.8, 
                                          y=svc_df_analysis['Revenue'].max()*0.9,
                                          text="⭐ Stars<br>High Volume, High Revenue",
                                          showarrow=False, font=dict(size=10))
                    fig_bcg.add_annotation(x=svc_df_analysis['Volume'].min()*1.5, 
                                          y=svc_df_analysis['Revenue'].max()*0.9,
                                          text="❓ Question Marks<br>Low Volume, High Revenue",
                                          showarrow=False, font=dict(size=10))
                    
                    fig_bcg.update_layout(height=450)
                    st.plotly_chart(fig_bcg, use_container_width=True)
                
                with col2:
                    # Service performance radar
                    top_5_services = svc_df_analysis.nlargest(5, 'Volume')
                    
                    if not top_5_services.empty:
                        categories = ['Volume Score', 'Revenue Score', 'OTP Score', 'Quality Score', 'Efficiency Score']
                        
                        fig_radar = go.Figure()
                        
                        for _, service in top_5_services.iterrows():
                            scores = [
                                (service['Volume'] / svc_df_analysis['Volume'].max()) * 100,
                                (service['Revenue'] / svc_df_analysis['Revenue'].max()) * 100,
                                service['OTP Gross'],
                                100 - service['QC Rate'],
                                (1 - (service['Avg Cost'] / svc_df_analysis['Avg Cost'].max())) * 100
                            ]
                            
                            fig_radar.add_trace(go.Scatterpolar(
                                r=scores,
                                theta=categories,
                                fill='toself',
                                name=service['Service'],
                                hovertemplate='%{theta}: %{r:.1f}<extra></extra>'
                            ))
                        
                        fig_radar.update_layout(
                            polar=dict(
                                radialaxis=dict(
                                    visible=True,
                                    range=[0, 100]
                                )
                            ),
                            title="Top 5 Services Performance Scorecard",
                            height=450
                        )
                        st.plotly_chart(fig_radar, use_container_width=True)
                
                # Service comparison table
                st.markdown("#### 📊 Service Performance Metrics")
                
                # Create styled dataframe
                styled_df = svc_df_analysis.sort_values('Volume', ascending=False)
                styled_df['Volume Rank'] = range(1, len(styled_df) + 1)
                styled_df['Performance'] = styled_df.apply(
                    lambda x: '🟢' if x['OTP Gross'] >= 90 else ('🟡' if x['OTP Gross'] >= 85 else '🔴'), 
                    axis=1
                )
                
                display_df = styled_df[['Volume Rank', 'Performance', 'Service', 'Description', 
                                       'Volume', 'Revenue', 'Avg Cost', 'OTP Gross', 'OTP Net', 'QC Rate']]
                
                st.dataframe(
                    display_df.style.format({
                        'Revenue': '€{:,.0f}',
                        'Avg Cost': '€{:,.0f}',
                        'OTP Gross': '{:.1f}%',
                        'OTP Net': '{:.1f}%',
                        'QC Rate': '{:.1f}%'
                    }).background_gradient(subset=['OTP Gross', 'OTP Net'], cmap='RdYlGn', vmin=70, vmax=100)
                    .background_gradient(subset=['QC Rate'], cmap='RdYlGn_r', vmin=0, vmax=20),
                    use_container_width=True,
                    height=400
                )
                
                # Service insights
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if not styled_df.empty:
                        best_otp = styled_df.nlargest(1, 'OTP Gross').iloc[0]
                        st.success(f"""
                        **🏆 Best Performing Service**  
                        {best_otp['Service']}: {best_otp['OTP Gross']:.1f}% OTP
                        """)
                
                with col2:
                    if not styled_df.empty:
                        highest_revenue = styled_df.nlargest(1, 'Revenue').iloc[0]
                        st.info(f"""
                        **💰 Highest Revenue Service**  
                        {highest_revenue['Service']}: €{highest_revenue['Revenue']:,.0f}
                        """)
                
                with col3:
                    if not styled_df.empty:
                        needs_attention = styled_df[styled_df['OTP Gross'] < 85]
                        if not needs_attention.empty:
                            st.warning(f"""
                            **⚠️ Services Below Target**  
                            {len(needs_attention)} services below 85% OTP
                            """)
        
        with tab5:
            st.markdown("### ⚠️ Quality Control & Improvement Opportunities")
            
            # QC Overview cards
            qc_data = df_filtered[df_filtered['QC NAME'].notna()]
            controllable_qc = qc_data[qc_data['QC NAME'].isin(CONTROLLABLE_QC_CODES)]
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown(f"""
                <div style='background: #fee2e2; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #991b1b; margin: 0;'>Total QC Issues</h4>
                    <h2 style='color: #dc2626; margin: 0.5rem 0;'>{len(qc_data):,}</h2>
                    <p style='color: #7f1d1d; font-size: 0.9rem; margin: 0;'>
                        {(len(qc_data)/len(df_filtered)*100):.1f}% of shipments
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div style='background: #fef3c7; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #78350f; margin: 0;'>Controllable</h4>
                    <h2 style='color: #f59e0b; margin: 0.5rem 0;'>{len(controllable_qc):,}</h2>
                    <p style='color: #92400e; font-size: 0.9rem; margin: 0;'>
                        {(len(controllable_qc)/len(qc_data)*100):.1f}% can be fixed
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                uncontrollable = len(qc_data) - len(controllable_qc)
                st.markdown(f"""
                <div style='background: #dbeafe; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #1e3a8a; margin: 0;'>Uncontrollable</h4>
                    <h2 style='color: #3b82f6; margin: 0.5rem 0;'>{uncontrollable:,}</h2>
                    <p style='color: #1e40af; font-size: 0.9rem; margin: 0;'>
                        Customer-driven issues
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                potential_improvement = (len(controllable_qc) / len(df_filtered)) * 100 if len(df_filtered) > 0 else 0
                st.markdown(f"""
                <div style='background: #d1fae5; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #064e3b; margin: 0;'>Potential OTP Gain</h4>
                    <h2 style='color: #10b981; margin: 0.5rem 0;'>+{potential_improvement:.1f}%</h2>
                    <p style='color: #047857; font-size: 0.9rem; margin: 0;'>
                        If controllable fixed
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            # QC Categories breakdown
            st.markdown("#### 🔍 Quality Issue Categories")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Controllable categories breakdown
                controllable_breakdown = []
                for category, codes in CONTROLLABLE_QC_CATEGORIES.items():
                    count = qc_data[qc_data['QC NAME'].isin(codes)].shape[0]
                    if count > 0:
                        controllable_breakdown.append({
                            'Category': category,
                            'Issues': count,
                            'Percentage': (count / len(qc_data)) * 100
                        })
                
                if controllable_breakdown:
                    ctrl_df = pd.DataFrame(controllable_breakdown)
                    
                    fig_ctrl = px.pie(
                        ctrl_df,
                        values='Issues',
                        names='Category',
                        title='Controllable Issues by Category',
                        color_discrete_sequence=px.colors.sequential.OrRd
                    )
                    fig_ctrl.update_traces(textposition='inside', textinfo='percent+label')
                    fig_ctrl.update_layout(height=350)
                    st.plotly_chart(fig_ctrl, use_container_width=True)
            
            with col2:
                # Top QC issues
                if not qc_data.empty:
                    qc_counts = qc_data['QC NAME'].value_counts().head(10)
                    
                    fig_top_qc = go.Figure(go.Bar(
                        x=qc_counts.values,
                        y=qc_counts.index,
                        orientation='h',
                        marker_color=['#ef4444' if issue in CONTROLLABLE_QC_CODES else '#3b82f6' 
                                     for issue in qc_counts.index],
                        text=qc_counts.values,
                        textposition='outside'
                    ))
                    
                    fig_top_qc.update_layout(
                        title='Top 10 Quality Issues',
                        xaxis_title='Number of Occurrences',
                        yaxis_title='',
                        height=350,
                        showlegend=False
                    )
                    st.plotly_chart(fig_top_qc, use_container_width=True)
            
            # QC Trend Analysis
            st.markdown("#### 📈 Quality Trends & Patterns")
            
            if 'ORD CREATE' in df_filtered.columns:
                # Weekly QC trend
                df_qc_trend = df_filtered.copy()
                df_qc_trend['Week'] = df_qc_trend['ORD CREATE'].dt.to_period('W')
                
                weekly_qc = df_qc_trend.groupby('Week').apply(
                    lambda x: pd.Series({
                        'Total': len(x),
                        'QC Issues': x['QC NAME'].notna().sum(),
                        'Controllable': x[x['QC NAME'].isin(CONTROLLABLE_QC_CODES)]['QC NAME'].notna().sum(),
                        'QC Rate': (x['QC NAME'].notna().sum() / len(x)) * 100 if len(x) > 0 else 0
                    })
                ).reset_index()
                weekly_qc['Week'] = weekly_qc['Week'].astype(str)
                
                fig_qc_trend = make_subplots(
                    rows=2, cols=1,
                    subplot_titles=('QC Issues Over Time', 'QC Rate Trend (%)'),
                    row_heights=[0.6, 0.4]
                )
                
                # Stacked bar chart for QC issues
                fig_qc_trend.add_trace(
                    go.Bar(name='Controllable', x=weekly_qc['Week'], y=weekly_qc['Controllable'],
                          marker_color='#f59e0b'),
                    row=1, col=1
                )
                fig_qc_trend.add_trace(
                    go.Bar(name='Uncontrollable', x=weekly_qc['Week'], 
                          y=weekly_qc['QC Issues'] - weekly_qc['Controllable'],
                          marker_color='#3b82f6'),
                    row=1, col=1
                )
                
                # QC rate trend line
                fig_qc_trend.add_trace(
                    go.Scatter(x=weekly_qc['Week'], y=weekly_qc['QC Rate'],
                              mode='lines+markers', name='QC Rate',
                              line=dict(color='#ef4444', width=3)),
                    row=2, col=1
                )
                fig_qc_trend.add_hline(y=10, line_dash="dash", line_color="green",
                                      annotation_text="Target <10%", row=2, col=1)
                
                fig_qc_trend.update_layout(
                    height=600,
                    barmode='stack',
                    showlegend=True,
                    hovermode='x unified'
                )
                st.plotly_chart(fig_qc_trend, use_container_width=True)
            
            # Action items
            st.markdown("#### 🎯 Recommended Actions")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                <div style='background: #fef3c7; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #78350f;'>🔧 Internal Process Improvements</h4>
                    <ul style='color: #92400e; margin: 0.5rem 0;'>
                        <li>Implement automated QDT validation system</li>
                        <li>Enhance order entry training program</li>
                        <li>Deploy real-time dispatch monitoring</li>
                        <li>Upgrade warehouse data systems</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown("""
                <div style='background: #dbeafe; padding: 1rem; border-radius: 8px;'>
                    <h4 style='color: #1e3a8a;'>🤝 Partner & Customer Actions</h4>
                    <ul style='color: #1e40af; margin: 0.5rem 0;'>
                        <li>Review SLAs with delivery agents</li>
                        <li>Implement partner scorecards</li>
                        <li>Enhance customer communication</li>
                        <li>Develop customs pre-clearance process</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            
            # ROI calculation
            st.markdown("#### 💰 Improvement ROI Estimation")
            
            if not controllable_qc.empty:
                avg_delay_cost = df_filtered[df_filtered['QC NAME'].isin(CONTROLLABLE_QC_CODES)]['TOTAL CHARGES EUR'].mean()
                total_controllable_cost = len(controllable_qc) * avg_delay_cost * 0.1  # Assume 10% cost impact
            else:
                avg_delay_cost = 0
                total_controllable_cost = 0
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(
                    "Estimated Annual Impact",
                    f"€{total_controllable_cost*12:,.0f}",
                    help="Potential cost savings from fixing controllable issues"
                )
            
            with col2:
                improvement_potential = (len(controllable_qc)/len(df_filtered)*100) if len(df_filtered) > 0 else 0
                st.metric(
                    "OTP Improvement Potential",
                    f"+{improvement_potential:.1f}%",
                    help="Potential OTP increase"
                )
            
            with col3:
                st.metric(
                    "Customer Satisfaction Impact",
                    "High",
                    help="Expected improvement in customer satisfaction"
                )
