import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')
import io

# Page configuration - MUST be first Streamlit command
st.set_page_config(
    page_title="Shipment Cost Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for professional styling
st.markdown("""
    <style>
    .main {
        padding: 0rem 1rem;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h1 {
        color: #1f2937;
        font-weight: 700;
        border-bottom: 3px solid #3b82f6;
        padding-bottom: 10px;
    }
    h2 {
        color: #374151;
        font-weight: 600;
        margin-top: 2rem;
    }
    .metric-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 15px;
        color: white;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)

# Title
st.markdown("# 📊 Shipment Cost Analytics Dashboard")
st.markdown("**Executive Overview - Facts & Figures**")

# Define controllable QC codes
CONTROLLABLE_QC_CODES = [
    262, 287, 183, 197, 199, 308, 309, 319, 326, 278, 203
]

# Optimized data loading function with caching
@st.cache_data(show_spinner=False)
def load_and_process_data(file_bytes):
    """Load and process Excel data with all transformations"""
    # Read Excel file
    df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    
    # Drop duplicate column names
    df = df.loc[:, ~df.columns.duplicated()]
    
    # Filter only 440-BILLED status
    df = df[df['STATUS'] == '440-BILLED'].copy()
    
    # Convert date columns
    date_columns = ['ORD CREATE', 'READY', 'QT PU', 'ACT PU', 'QDT', 'UPD DEL', 
                   'POD DATE/TIME', 'Depart Date / Time', 'Arrive Date / Time', 
                   'PICKUP DATE/TIME', 'Depart Date']
    
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Currency conversion (USD to EUR)
    USD_TO_EUR = 0.92
    df['TOTAL_CHARGES_EUR'] = pd.to_numeric(df['TOTAL CHARGES'], errors='coerce') * USD_TO_EUR
    
    # Create route column
    if {'DEP', 'ARR'}.issubset(df.columns):
        df['Route'] = df['DEP'].astype(str) + ' → ' + df['ARR'].astype(str)
    
    return df

@st.cache_data(show_spinner=False)
def calculate_otp_metrics(df, controllable_codes):
    """Calculate OTP metrics with caching"""
    df_otp = df.dropna(subset=['QDT', 'POD DATE/TIME']).copy()
    
    if len(df_otp) == 0:
        return 0, 0, pd.DataFrame()
    
    df_otp['ON_TIME_GROSS'] = df_otp['POD DATE/TIME'] <= df_otp['QDT']
    df_otp['LATE'] = ~df_otp['ON_TIME_GROSS']
    df_otp['CONTROLLABLE_DELAY'] = df_otp['QCCODE'].isin(controllable_codes)
    df_otp['ON_TIME_NET'] = df_otp['ON_TIME_GROSS'] | (df_otp['LATE'] & ~df_otp['CONTROLLABLE_DELAY'])
    
    gross_otp = (df_otp['ON_TIME_GROSS'].sum() / len(df_otp) * 100)
    net_otp = (df_otp['ON_TIME_NET'].sum() / len(df_otp) * 100)
    
    return gross_otp, net_otp, df_otp

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'], label_visibility="collapsed")

if uploaded_file is not None:
    # Load data with spinner
    with st.spinner('Loading and processing data...'):
        file_bytes = uploaded_file.getvalue()
        df = load_and_process_data(file_bytes)
        gross_otp, net_otp, df_otp = calculate_otp_metrics(df, CONTROLLABLE_QC_CODES)
    
    # Key Metrics Row
    st.markdown("---")
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        total_shipments = len(df)
        st.metric("Total Shipments", f"{total_shipments:,}")
    
    with col2:
        total_cost_eur = df['TOTAL_CHARGES_EUR'].sum()
        st.metric("Total Cost", f"€{total_cost_eur:,.0f}")
    
    with col3:
        avg_cost = df['TOTAL_CHARGES_EUR'].mean()
        st.metric("Avg Cost/Shipment", f"€{avg_cost:,.2f}")
    
    with col4:
        st.metric("OTP Gross", f"{gross_otp:.1f}%", 
                 delta=f"{gross_otp-85:.1f}%" if gross_otp >= 85 else f"{gross_otp-85:.1f}%")
    
    with col5:
        st.metric("OTP Net", f"{net_otp:.1f}%", 
                 delta=f"{net_otp-gross_otp:.1f}%" if net_otp > gross_otp else None)
    
    # Create two columns for main visualizations
    st.markdown("---")
    
    # Row 1: Service Distribution and OTP Comparison
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Service Type Distribution")
        svc_counts = df['SVC'].value_counts().reset_index()
        svc_counts.columns = ['Service', 'Count']
        
        if 'SVCDESC' in df.columns:
            svc_desc_map = df.groupby('SVC')['SVCDESC'].first().to_dict()
            svc_counts['Description'] = svc_counts['Service'].map(svc_desc_map)
        else:
            svc_counts['Description'] = svc_counts['Service']
        
        svc_counts['Percentage'] = (svc_counts['Count'] / svc_counts['Count'].sum() * 100).round(1)
        svc_counts_sorted = svc_counts.head(10).sort_values('Count', ascending=True)
        
        fig_svc = px.bar(svc_counts_sorted, 
                         x='Count', 
                         y='Service',
                         orientation='h',
                         text='Count',
                         hover_data=['Description', 'Percentage'],
                         color='Count',
                         color_continuous_scale='Viridis')
        
        fig_svc.update_traces(texttemplate='%{text}', textposition='outside')
        fig_svc.update_layout(
            height=400,
            showlegend=False,
            xaxis_title="Number of Shipments",
            yaxis_title="Service Type",
            margin=dict(l=0, r=0, t=0, b=0),
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_svc, use_container_width=True)
    
    with col2:
        st.markdown("### OTP Performance: Gross vs Net")
        st.markdown("""
        **Gross OTP**: Shipments delivered on or before quoted delivery time  
        **Net OTP**: Excludes delays from controllable factors (customs, warehouse, MNX errors)
        """)
        
        otp_data = pd.DataFrame({
            'Metric': ['Gross OTP', 'Net OTP', 'Gap'],
            'Percentage': [gross_otp, net_otp, net_otp - gross_otp],
            'Color': ['#3b82f6', '#10b981', '#fbbf24']
        })
        
        fig_otp = go.Figure()
        fig_otp.add_trace(go.Bar(
            x=otp_data['Metric'],
            y=otp_data['Percentage'],
            text=[f"{val:.1f}%" for val in otp_data['Percentage']],
            textposition='outside',
            marker_color=otp_data['Color'],
            hovertemplate='%{x}: %{y:.1f}%<extra></extra>'
        ))
        
        fig_otp.add_hline(y=90, line_dash="dash", line_color="red", 
                         annotation_text="Industry Standard (90%)")
        
        fig_otp.update_layout(
            height=400,
            showlegend=False,
            yaxis_title="Percentage (%)",
            xaxis_title="",
            margin=dict(l=0, r=0, t=20, b=0),
            yaxis=dict(range=[0, 105])
        )
        st.plotly_chart(fig_otp, use_container_width=True)
    
    # Row 2: Departure Airport Analysis and Cost Distribution
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Top Departure (DEP) Airports")
        if 'DEP' in df.columns:
            dep_counts = df['DEP'].value_counts().head(15).reset_index()
            dep_counts.columns = ['Airport', 'Shipments']
            
            dep_cost = df.groupby('DEP')['TOTAL_CHARGES_EUR'].agg(['sum', 'mean']).reset_index()
            dep_cost.columns = ['Airport', 'Total_Cost', 'Avg_Cost']
            dep_counts = dep_counts.merge(dep_cost, on='Airport', how='left')
            
            fig_dep = px.treemap(dep_counts, 
                                path=['Airport'], 
                                values='Shipments',
                                color='Avg_Cost',
                                hover_data={'Shipments': True, 'Total_Cost': ':.0f', 'Avg_Cost': ':.2f'},
                                color_continuous_scale='RdYlGn_r',
                                labels={'Avg_Cost': 'Avg Cost (€)', 'DEP': 'Departure'})
            
            fig_dep.update_layout(
                height=400,
                margin=dict(l=0, r=0, t=0, b=0)
            )
            st.plotly_chart(fig_dep, use_container_width=True)
        else:
            st.info("No departure airport data available")
    
    with col2:
        st.markdown("### Cost Distribution Analysis")
        
        df['Cost_Bin'] = pd.cut(df['TOTAL_CHARGES_EUR'], 
                                bins=[0, 500, 1000, 2000, 5000, float('inf')],
                                labels=['<€500', '€500-1K', '€1K-2K', '€2K-5K', '>€5K'])
        
        cost_dist = df['Cost_Bin'].value_counts().sort_index().reset_index()
        cost_dist.columns = ['Cost Range', 'Count']
        
        fig_cost = px.pie(cost_dist, 
                         values='Count', 
                         names='Cost Range',
                         hole=0.4,
                         color_discrete_sequence=px.colors.sequential.Viridis)
        
        fig_cost.update_traces(textposition='inside', textinfo='percent+label')
        fig_cost.update_layout(
            height=400,
            margin=dict(l=0, r=0, t=0, b=0),
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1)
        )
        st.plotly_chart(fig_cost, use_container_width=True)
    
    # Row 3: Time Analysis and QC Analysis
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Monthly Trend Analysis")
        if 'ORD CREATE' in df.columns:
            df['Month'] = pd.to_datetime(df['ORD CREATE']).dt.to_period('M')
            monthly_stats = df.groupby('Month').agg({
                'REFER': 'count',
                'TOTAL_CHARGES_EUR': 'sum'
            }).reset_index()
            monthly_stats.columns = ['Month', 'Shipments', 'Total_Cost']
            monthly_stats['Month'] = monthly_stats['Month'].astype(str)
            
            fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
            
            fig_trend.add_trace(
                go.Bar(x=monthly_stats['Month'], 
                      y=monthly_stats['Shipments'],
                      name='Shipments',
                      marker_color='#3b82f6',
                      yaxis='y'),
                secondary_y=False
            )
            
            fig_trend.add_trace(
                go.Scatter(x=monthly_stats['Month'], 
                          y=monthly_stats['Total_Cost'],
                          name='Total Cost (€)',
                          mode='lines+markers',
                          marker_color='#ef4444',
                          line=dict(width=3),
                          yaxis='y2'),
                secondary_y=True
            )
            
            fig_trend.update_xaxes(title_text="Month")
            fig_trend.update_yaxes(title_text="Number of Shipments", secondary_y=False)
            fig_trend.update_yaxes(title_text="Total Cost (€)", secondary_y=True)
            fig_trend.update_layout(
                height=400,
                margin=dict(l=0, r=0, t=20, b=0),
                hovermode='x unified',
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            st.plotly_chart(fig_trend, use_container_width=True)
        else:
            st.info("No order creation date data available")
    
    with col2:
        st.markdown("### Quality Control Issues Analysis")
        st.markdown("🟢 **Controllable** (Internal) | 🔴 **Non-Controllable** (External)")
        
        if 'QCCODE' in df.columns and 'QC NAME' in df.columns:
            qc_data = df[df['QCCODE'].notna()].copy()
            
            if len(qc_data) > 0:
                qc_counts = qc_data.groupby(['QCCODE', 'QC NAME']).size().reset_index(name='Count')
                qc_counts['Issue Type'] = qc_counts['QCCODE'].apply(
                    lambda x: 'Controllable' if x in CONTROLLABLE_QC_CODES else 'Non-Controllable'
                )
                qc_counts = qc_counts.sort_values('Count', ascending=False).head(10)
                
                fig_qc = px.bar(qc_counts, 
                               x='Count', 
                               y='QC NAME',
                               orientation='h',
                               color='Issue Type',
                               color_discrete_map={'Controllable': '#10b981', 'Non-Controllable': '#ef4444'},
                               text='Count')
                
                fig_qc.update_traces(texttemplate='%{text}', textposition='outside')
                fig_qc.update_layout(
                    height=400,
                    xaxis_title="Number of Occurrences",
                    yaxis_title="",
                    margin=dict(l=0, r=0, t=20, b=0),
                    legend_title="Issue Type",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig_qc, use_container_width=True)
            else:
                st.info("No quality control issues found in the data")
        else:
            st.info("Quality control data not available")
    
    # Row 4: Weight Analysis and Route Performance
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Weight Distribution & Cost Correlation")
        
        if 'Billable Weight KG' in df.columns and 'PIECES' in df.columns:
            weight_data = df[df['Billable Weight KG'].notna() & (df['Billable Weight KG'] > 0)].copy()
            
            if len(weight_data) > 0:
                sample_size = min(500, len(weight_data))
                fig_weight = px.scatter(weight_data.sample(sample_size), 
                                      x='Billable Weight KG', 
                                      y='TOTAL_CHARGES_EUR',
                                      color='SVC',
                                      size='PIECES',
                                      hover_data=['REFER', 'DEP', 'ARR'] if all(col in df.columns for col in ['REFER', 'DEP', 'ARR']) else [],
                                      labels={'TOTAL_CHARGES_EUR': 'Cost (€)', 
                                             'Billable Weight KG': 'Weight (KG)'},
                                      opacity=0.6)
                
                fig_weight.update_layout(
                    height=400,
                    margin=dict(l=0, r=0, t=20, b=0),
                    legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02)
                )
                st.plotly_chart(fig_weight, use_container_width=True)
            else:
                st.info("No weight data available for analysis")
        else:
            st.info("Weight data columns not available")
    
    with col2:
        st.markdown("### Top Routes Performance")
        
        if 'Route' in df.columns:
            route_stats = df.groupby('Route').agg({
                'REFER': 'count',
                'TOTAL_CHARGES_EUR': 'mean'
            }).reset_index()
            route_stats.columns = ['Route', 'Shipments', 'Avg_Cost']
            route_stats = route_stats[route_stats['Shipments'] >= 5]
            route_stats = route_stats.sort_values('Shipments', ascending=False).head(15)
            
            if len(route_stats) > 0:
                fig_route = px.scatter(route_stats, 
                                     x='Shipments', 
                                     y='Avg_Cost',
                                     size='Shipments',
                                     color='Avg_Cost',
                                     text='Route',
                                     color_continuous_scale='RdYlGn_r',
                                     labels={'Avg_Cost': 'Avg Cost (€)', 'Shipments': 'Number of Shipments'})
                
                fig_route.update_traces(textposition='top center', textfont_size=8)
                fig_route.update_layout(
                    height=400,
                    margin=dict(l=0, r=0, t=20, b=0),
                    showlegend=False,
                    coloraxis_colorbar=dict(title="Avg Cost (€)")
                )
                st.plotly_chart(fig_route, use_container_width=True)
            else:
                st.info("Not enough route data for analysis")
        else:
            st.info("Route data not available")
    
    # Executive Summary
    st.markdown("---")
    st.markdown("### 📋 Executive Summary")
    
    col1, col2 = st.columns(2)
    
    with col1:
        dep_top = df['DEP'].value_counts().index[0] if 'DEP' in df.columns and len(df['DEP'].value_counts()) > 0 else 'N/A'
        dep_pct = (df['DEP'].value_counts().iloc[0] / len(df) * 100) if 'DEP' in df.columns and len(df['DEP'].value_counts()) > 0 else 0
        
        st.markdown("""
        **Key Performance Indicators:**
        - Total shipment volume: **{:,} shipments**
        - Total cost in period: **€{:,.0f}**
        - Average cost per shipment: **€{:.2f}**
        - Main departure hub: **{}** ({:.1f}% of volume)
        """.format(
            len(df),
            df['TOTAL_CHARGES_EUR'].sum(),
            df['TOTAL_CHARGES_EUR'].mean(),
            dep_top,
            dep_pct
        ))
    
    with col2:
        st.markdown("""
        **On-Time Performance Analysis:**
        - Gross OTP: **{:.1f}%** - Actual delivery performance
        - Net OTP: **{:.1f}%** - Excluding controllable delays
        - Performance gap: **{:.1f}%** - Improvement opportunity
        - Controllable delays represent potential cost savings through process optimization
        """.format(gross_otp, net_otp, net_otp - gross_otp))
    
    # Data Quality Metrics
    st.markdown("---")
    st.markdown("### 📊 Data Quality Metrics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        completeness = (1 - df['TOTAL CHARGES'].isna().sum() / len(df)) * 100 if 'TOTAL CHARGES' in df.columns else 0
        st.metric("Cost Data Completeness", f"{completeness:.1f}%")
    
    with col2:
        date_completeness = (1 - df['POD DATE/TIME'].isna().sum() / len(df)) * 100 if 'POD DATE/TIME' in df.columns else 0
        st.metric("Delivery Data Completeness", f"{date_completeness:.1f}%")
    
    with col3:
        if 'QCCODE' in df.columns and 'POD DATE/TIME' in df.columns and 'QDT' in df.columns:
            late_shipments = df[df['POD DATE/TIME'] > df['QDT']]
            qc_coverage = (df['QCCODE'].notna().sum() / len(late_shipments)) * 100 if len(late_shipments) > 0 else 100
        else:
            qc_coverage = 0
        st.metric("QC Code Coverage", f"{qc_coverage:.1f}%")
    
    with col4:
        unique_routes = df['Route'].nunique() if 'Route' in df.columns else 0
        st.metric("Unique Routes", f"{unique_routes:,}")

else:
    # Landing page when no file is uploaded
    st.markdown("""
    ### Welcome to the Shipment Cost Analytics Dashboard
    
    This professional dashboard provides comprehensive analysis of shipment costs and performance metrics.
    
    **Key Features:**
    - 📊 Real-time OTP (On-Time Performance) tracking - Gross vs Net analysis
    - 💰 Cost analysis in EUR with automatic currency conversion
    - 🌍 Departure airport distribution and route performance
    - 📈 Service type breakdown and quality control insights
    - 📅 Temporal trends and monthly performance tracking
    
    **Upload your Excel file to begin analysis.**
    
    The dashboard will automatically:
    - Filter for 440-BILLED status shipments
    - Calculate OTP metrics (Gross and Net)
    - Convert costs to EUR
    - Generate executive-ready visualizations
    """)
    
    st.info("Please upload the shipment data Excel file to generate insights.")
