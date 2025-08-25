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
    page_title="Shipment Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for professional styling
st.markdown("""
    <style>
    .main {
        padding: 0rem 1rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        color: white;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    h1 {
        color: #1e3a8a;
        font-weight: 700;
    }
    h2 {
        color: #334155;
        font-weight: 600;
        margin-top: 2rem;
    }
    </style>
    """, unsafe_allow_html=True)

# Load data function
@st.cache_data
def load_data(file_path):
    """Load and preprocess the Excel data"""
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Filter for 440-BILLED status
    df_billed = df[df['STATUS'] == '440-BILLED'].copy()
    
    # Convert date columns - handle both Excel serial dates and datetime strings
    date_columns = ['ORD CREATE', 'Depart Date / Time', 'Arrive Date / Time', 'POD DATE/TIME', 'QDT', 'READY']
    for col in date_columns:
        if col in df_billed.columns:
            # First try to convert as datetime string
            df_billed[col] = pd.to_datetime(df_billed[col], errors='coerce')
            
            # For any remaining numeric values (Excel serial dates), convert them
            numeric_mask = pd.to_numeric(df_billed[col], errors='coerce').notna()
            if numeric_mask.any():
                # Create a temporary series for numeric values
                temp_numeric = pd.to_numeric(df_billed.loc[numeric_mask, col], errors='coerce')
                # Convert Excel serial date to datetime
                df_billed.loc[numeric_mask, col] = pd.to_datetime('1899-12-30') + pd.to_timedelta(temp_numeric, unit='D')
    
    # Clean numeric columns
    df_billed['TOTAL CHARGES'] = pd.to_numeric(df_billed['TOTAL CHARGES'], errors='coerce')
    df_billed['PIECES'] = pd.to_numeric(df_billed['PIECES'], errors='coerce')
    df_billed['Billable Weight KG'] = pd.to_numeric(df_billed['Billable Weight KG'], errors='coerce')
    df_billed['Time In Transit'] = pd.to_numeric(df_billed['Time In Transit'], errors='coerce')
    
    # Convert to EUR (assuming USD to EUR rate of 0.92)
    USD_TO_EUR = 0.92
    df_billed['TOTAL CHARGES EUR'] = df_billed['TOTAL CHARGES'] * USD_TO_EUR
    
    return df_billed

# Define controllable QC codes based on the data
CONTROLLABLE_QC_CODES = {
    '165': 'Customer-Requested delay',
    '164': 'Customer-Changed delivery parameters',
    '175': 'Customer-Unattainable QDT (Online)',
    '173': 'Customer-Shipment not ready',
    '163': 'Customer-Delayed clearance docs',
    '338': 'Shipment not ready',
    '319': 'Customs-Late PWK-Customer',
    '308': 'Customs delay',
    '309': 'Customs delay-FDA Hold',
    '326': 'W/House-Data entry errors'
}

def calculate_otp(df):
    """Calculate On-Time Performance (OTP) metrics"""
    # Ensure we have the necessary columns
    if 'POD DATE/TIME' not in df.columns or 'QDT' not in df.columns:
        return 0, 0
    
    # Create a copy to avoid warnings
    df_calc = df.copy()
    
    # Gross OTP: All shipments delivered on time
    # Only calculate for rows where both dates are available
    valid_dates = df_calc['POD DATE/TIME'].notna() & df_calc['QDT'].notna()
    if valid_dates.sum() > 0:
        df_calc.loc[valid_dates, 'On_Time'] = pd.to_datetime(df_calc.loc[valid_dates, 'POD DATE/TIME']) <= pd.to_datetime(df_calc.loc[valid_dates, 'QDT'])
        gross_otp = (df_calc['On_Time'].sum() / valid_dates.sum()) * 100
    else:
        gross_otp = 0
        df_calc['On_Time'] = False
    
    # Net OTP: Excluding non-controllable delays
    df_calc['Is_Controllable'] = df_calc['QCCODE'].astype(str).isin(CONTROLLABLE_QC_CODES.keys())
    
    # For net OTP, exclude shipments with non-controllable delays
    df_controllable = df_calc[~df_calc['Is_Controllable'] | df_calc['On_Time']]
    if len(df_controllable) > 0 and 'On_Time' in df_controllable.columns:
        net_otp = (df_controllable['On_Time'].sum() / len(df_controllable)) * 100
    else:
        net_otp = gross_otp
    
    # Add the calculated columns back to the original dataframe
    df['On_Time'] = df_calc['On_Time'] if 'On_Time' in df_calc.columns else False
    df['Is_Controllable'] = df_calc['Is_Controllable'] if 'Is_Controllable' in df_calc.columns else False
    
    return gross_otp, net_otp

# Title and header
st.title("📊 Shipment Performance Dashboard")
st.markdown("**Executive Overview - Year to Date 2025**")

# Sidebar for filters
with st.sidebar:
    st.header("🔍 Filters")
    
    # File upload option
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        df = load_data(uploaded_file)
    else:
        # Try to load default file
        try:
            df = load_data('shipment data YTD 25.xlsx')
        except:
            st.error("Please upload the shipment data Excel file")
            st.stop()
    
    # Date range filter
    st.subheader("Date Range")
    if not df.empty and 'ORD CREATE' in df.columns:
        min_date = df['ORD CREATE'].min()
        max_date = df['ORD CREATE'].max()
        if pd.notna(min_date) and pd.notna(max_date):
            date_range = st.date_input(
                "Select Period",
                value=(min_date.date() if isinstance(min_date, pd.Timestamp) else datetime.now().date(),
                       max_date.date() if isinstance(max_date, pd.Timestamp) else datetime.now().date()),
                format="DD/MM/YYYY"
            )
        else:
            date_range = None
    else:
        date_range = None
    
    # Service filter
    st.subheader("Service Type")
    svc_options = ['All'] + sorted(df['SVC'].dropna().unique().tolist())
    selected_svc = st.multiselect("Select Services", svc_options, default=['All'])
    
    # Departure airport filter
    st.subheader("Departure Airport")
    dep_options = ['All'] + sorted(df['DEP'].dropna().unique().tolist())
    selected_dep = st.multiselect("Select Airports", dep_options, default=['All'])

# Apply filters
df_filtered = df.copy()

if 'All' not in selected_svc:
    df_filtered = df_filtered[df_filtered['SVC'].isin(selected_svc)]

if 'All' not in selected_dep:
    df_filtered = df_filtered[df_filtered['DEP'].isin(selected_dep)]

# Calculate OTP metrics
gross_otp, net_otp = calculate_otp(df_filtered)

# Key Performance Indicators
st.markdown("## 📈 Key Performance Indicators")

col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    total_shipments = len(df_filtered)
    st.metric(
        label="Total Shipments",
        value=f"{total_shipments:,}",
        delta=f"YTD 2025"
    )

with col2:
    total_revenue = df_filtered['TOTAL CHARGES EUR'].sum()
    st.metric(
        label="Total Revenue (€)",
        value=f"€{total_revenue:,.0f}",
        delta=f"Avg: €{total_revenue/total_shipments:,.0f}" if total_shipments > 0 else "€0"
    )

with col3:
    st.metric(
        label="Gross OTP",
        value=f"{gross_otp:.1f}%",
        delta="All shipments",
        delta_color="normal" if gross_otp >= 95 else "inverse"
    )

with col4:
    st.metric(
        label="Net OTP",
        value=f"{net_otp:.1f}%",
        delta="Controllable only",
        delta_color="normal" if net_otp >= 95 else "inverse"
    )

with col5:
    avg_transit = df_filtered['Time In Transit'].mean() if 'Time In Transit' in df_filtered.columns else 0
    st.metric(
        label="Avg Transit Time",
        value=f"{avg_transit:.1f} days" if avg_transit else "N/A",
        delta="Days in transit"
    )

with col6:
    total_weight = df_filtered['Billable Weight KG'].sum()
    st.metric(
        label="Total Weight",
        value=f"{total_weight:,.0f} KG",
        delta=f"Avg: {total_weight/total_shipments:,.0f} KG" if total_shipments > 0 else "0 KG"
    )

# OTP Explanation
with st.expander("ℹ️ Understanding OTP Metrics"):
    st.markdown("""
    ### On-Time Performance (OTP) Explained
    
    **🎯 Gross OTP:** 
    - Measures the percentage of all shipments delivered on or before the quoted delivery time (QDT)
    - Includes ALL delays regardless of cause
    - Industry benchmark: >95%
    
    **✅ Net OTP:** 
    - Excludes delays due to non-controllable factors
    - Controllable factors include: Customer requests, customs delays, warehouse errors
    - Non-controllable factors include: Weather, airline delays, force majeure
    - Provides a clearer picture of operational performance
    
    **Why the difference matters:**
    - Gross OTP shows overall customer experience
    - Net OTP shows operational efficiency within your control
    - Gap between them indicates external factor impact
    """)

# Service Analysis
st.markdown("## 🚚 Service Type Analysis")

col1, col2 = st.columns(2)

with col1:
    # Service distribution pie chart
    svc_dist = df_filtered.groupby('SVC').agg({
        'REFER': 'count',
        'TOTAL CHARGES EUR': 'sum'
    }).reset_index()
    svc_dist.columns = ['Service', 'Count', 'Revenue']
    
    fig_svc_pie = px.pie(
        svc_dist, 
        values='Count', 
        names='Service',
        title='Shipment Distribution by Service Type',
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    fig_svc_pie.update_traces(textposition='inside', textinfo='percent+label')
    fig_svc_pie.update_layout(height=400)
    st.plotly_chart(fig_svc_pie, use_container_width=True)

with col2:
    # Revenue by service bar chart
    fig_svc_revenue = px.bar(
        svc_dist.sort_values('Revenue', ascending=True).tail(10),
        x='Revenue',
        y='Service',
        orientation='h',
        title='Top 10 Services by Revenue (€)',
        color='Revenue',
        color_continuous_scale='Viridis',
        text='Revenue'
    )
    fig_svc_revenue.update_traces(texttemplate='€%{text:,.0f}', textposition='outside')
    fig_svc_revenue.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig_svc_revenue, use_container_width=True)

# Departure Airport Analysis
st.markdown("## ✈️ Departure Airport Performance")

dep_analysis = df_filtered.groupby('DEP').agg({
    'REFER': 'count',
    'TOTAL CHARGES EUR': 'sum',
    'Time In Transit': 'mean',
    'On_Time': lambda x: (x.sum() / len(x) * 100) if len(x) > 0 else 0
}).reset_index()
dep_analysis.columns = ['Airport', 'Shipments', 'Revenue', 'Avg Transit Days', 'OTP %']
dep_analysis = dep_analysis.sort_values('Shipments', ascending=False).head(15)

# Create subplots for airport analysis
fig_airport = make_subplots(
    rows=2, cols=2,
    subplot_titles=('Shipment Volume by Airport', 'Revenue by Airport (€)', 
                    'Average Transit Time by Airport', 'OTP % by Airport'),
    specs=[[{'type': 'bar'}, {'type': 'bar'}],
           [{'type': 'bar'}, {'type': 'scatter'}]]
)

# Shipment volume
fig_airport.add_trace(
    go.Bar(x=dep_analysis['Airport'], y=dep_analysis['Shipments'], 
           name='Shipments', marker_color='lightblue'),
    row=1, col=1
)

# Revenue
fig_airport.add_trace(
    go.Bar(x=dep_analysis['Airport'], y=dep_analysis['Revenue'], 
           name='Revenue', marker_color='lightgreen'),
    row=1, col=2
)

# Transit time
fig_airport.add_trace(
    go.Bar(x=dep_analysis['Airport'], y=dep_analysis['Avg Transit Days'], 
           name='Transit Days', marker_color='coral'),
    row=2, col=1
)

# OTP percentage
fig_airport.add_trace(
    go.Scatter(x=dep_analysis['Airport'], y=dep_analysis['OTP %'], 
               mode='lines+markers', name='OTP %', marker_color='purple'),
    row=2, col=2
)

fig_airport.update_layout(height=800, showlegend=False)
fig_airport.update_xaxes(tickangle=45)
st.plotly_chart(fig_airport, use_container_width=True)

# Time Series Analysis
st.markdown("## 📅 Temporal Trends")

# Prepare time series data - only if we have date data
if 'ORD CREATE' in df_filtered.columns and df_filtered['ORD CREATE'].notna().any():
    df_time = df_filtered[df_filtered['ORD CREATE'].notna()].copy()
    df_time['Month'] = pd.to_datetime(df_time['ORD CREATE']).dt.to_period('M')
    
    monthly_data = df_time.groupby('Month').agg({
        'REFER': 'count',
        'TOTAL CHARGES EUR': 'sum'
    }).reset_index()
    
    # Calculate OTP only if On_Time column exists
    if 'On_Time' in df_time.columns:
        monthly_otp = df_time.groupby('Month')['On_Time'].apply(lambda x: (x.sum() / len(x) * 100) if len(x) > 0 else 0).reset_index()
        monthly_data = monthly_data.merge(monthly_otp, on='Month', how='left')
    else:
        monthly_data['On_Time'] = 0
    
    monthly_data['Month'] = monthly_data['Month'].astype(str)
    monthly_data.columns = ['Month', 'Shipments', 'Revenue', 'OTP %']
    
    # Create time series chart
    fig_timeline = make_subplots(
        rows=3, cols=1,
        subplot_titles=('Monthly Shipment Volume', 'Monthly Revenue (€)', 'Monthly OTP %'),
        row_heights=[0.33, 0.33, 0.34]
    )
    
    fig_timeline.add_trace(
        go.Scatter(x=monthly_data['Month'], y=monthly_data['Shipments'],
                   mode='lines+markers', name='Shipments', fill='tozeroy',
                   line=dict(color='blue', width=3)),
        row=1, col=1
    )
    
    fig_timeline.add_trace(
        go.Scatter(x=monthly_data['Month'], y=monthly_data['Revenue'],
                   mode='lines+markers', name='Revenue', fill='tozeroy',
                   line=dict(color='green', width=3)),
        row=2, col=1
    )
    
    fig_timeline.add_trace(
        go.Scatter(x=monthly_data['Month'], y=monthly_data['OTP %'],
                   mode='lines+markers', name='OTP %',
                   line=dict(color='red', width=3)),
        row=3, col=1
    )
    
    # Add 95% OTP target line
    fig_timeline.add_hline(y=95, line_dash="dash", line_color="gray", 
                           annotation_text="Target: 95%", row=3, col=1)
    
    fig_timeline.update_layout(height=900, showlegend=False)
    st.plotly_chart(fig_timeline, use_container_width=True)
else:
    st.info("Time series analysis requires valid date data")

# Quality Control Analysis
st.markdown("## 🔍 Quality Control Analysis")

col1, col2 = st.columns(2)

with col1:
    # QC distribution
    qc_dist = df_filtered[df_filtered['QCCODE'].notna()].groupby('QC NAME').size().reset_index(name='Count')
    qc_dist = qc_dist.sort_values('Count', ascending=False).head(10)
    
    fig_qc = px.bar(
        qc_dist,
        y='QC NAME',
        x='Count',
        orientation='h',
        title='Top 10 Quality Control Issues',
        color='Count',
        color_continuous_scale='Reds'
    )
    fig_qc.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig_qc, use_container_width=True)

with col2:
    # Controllable vs Non-controllable
    controllable_count = df_filtered['Is_Controllable'].sum()
    non_controllable = len(df_filtered[df_filtered['QCCODE'].notna()]) - controllable_count
    
    fig_control = go.Figure(data=[
        go.Bar(name='Controllable', x=['QC Issues'], y=[controllable_count], marker_color='orange'),
        go.Bar(name='Non-Controllable', x=['QC Issues'], y=[non_controllable], marker_color='gray')
    ])
    fig_control.update_layout(
        title='Controllable vs Non-Controllable Issues',
        barmode='stack',
        height=400
    )
    st.plotly_chart(fig_control, use_container_width=True)

# Cost Analysis
st.markdown("## 💰 Financial Performance")

col1, col2 = st.columns(2)

with col1:
    # Cost distribution histogram
    fig_cost_dist = px.histogram(
        df_filtered[df_filtered['TOTAL CHARGES EUR'] < 5000],  # Filter outliers for better visualization
        x='TOTAL CHARGES EUR',
        nbins=50,
        title='Shipment Cost Distribution (€)',
        labels={'TOTAL CHARGES EUR': 'Cost (€)', 'count': 'Number of Shipments'}
    )
    fig_cost_dist.update_layout(height=400)
    st.plotly_chart(fig_cost_dist, use_container_width=True)

with col2:
    # Weight vs Cost scatter
    fig_weight_cost = px.scatter(
        df_filtered[df_filtered['Billable Weight KG'] < 1000],  # Filter for visualization
        x='Billable Weight KG',
        y='TOTAL CHARGES EUR',
        color='SVC',
        title='Weight vs Cost Analysis',
        labels={'Billable Weight KG': 'Weight (KG)', 'TOTAL CHARGES EUR': 'Cost (€)'},
        trendline='ols'
    )
    fig_weight_cost.update_layout(height=400)
    st.plotly_chart(fig_weight_cost, use_container_width=True)

# Executive Summary
st.markdown("## 📋 Executive Summary & Recommendations")

with st.container():
    st.markdown("""
    ### Key Findings:
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        **Performance Metrics:**
        - Total shipments processed: **{total_shipments:,}**
        - Total revenue generated: **€{total_revenue:,.0f}**
        - Gross OTP: **{gross_otp:.1f}%** {'✅' if gross_otp >= 95 else '⚠️'}
        - Net OTP: **{net_otp:.1f}%** {'✅' if net_otp >= 95 else '⚠️'}
        - OTP Gap: **{abs(net_otp - gross_otp):.1f}%** (external factors impact)
        """)
        
    with col2:
        # Top performing routes
        top_routes = df_filtered.groupby('DEP')['On_Time'].mean().sort_values(ascending=False).head(3)
        st.markdown("**Top Performing Routes:**")
        for route, otp in top_routes.items():
            st.markdown(f"- {route}: {otp*100:.1f}% OTP")
        
        # Bottom performing routes
        bottom_routes = df_filtered.groupby('DEP')['On_Time'].mean().sort_values(ascending=True).head(3)
        st.markdown("**Routes Needing Attention:**")
        for route, otp in bottom_routes.items():
            st.markdown(f"- {route}: {otp*100:.1f}% OTP ⚠️")

    st.markdown("""
    ### 🎯 Strategic Recommendations:
    
    1. **Improve OTP Performance**
       - Focus on controllable factors, particularly customer communication and customs documentation
       - Implement proactive alerts for shipments at risk of delay
       
    2. **Route Optimization**
       - Review underperforming departure airports and consider alternative routing
       - Increase capacity on high-performing routes
       
    3. **Cost Management**
       - Analyze high-cost outliers for potential optimization
       - Review pricing strategy for low-volume, high-cost services
       
    4. **Quality Control**
       - Address top QC issues through targeted training and process improvements
       - Implement preventive measures for recurring controllable delays
    """)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Dashboard generated on: {}</p>
    <p>Data source: Shipment Data YTD 2025 | Status: 440-BILLED</p>
    <p>All amounts displayed in EUR (€)</p>
</div>
""".format(datetime.now().strftime('%Y-%m-%d %H:%M')), unsafe_allow_html=True)
