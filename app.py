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
    .info-box {
        background-color: #e0f2fe;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #0284c7;
        margin: 1rem 0;
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
    
    # Convert date columns - handle various date formats
    date_columns = ['ORD CREATE', 'Depart Date / Time', 'Arrive Date / Time', 'POD DATE/TIME', 
                   'QDT', 'READY', 'UPD DEL', 'PICKUP DATE/TIME', 'QT PU']
    
    for col in date_columns:
        if col in df_billed.columns:
            # Create a temporary column to store converted dates
            temp_dates = pd.Series(index=df_billed.index, dtype='datetime64[ns]')
            
            for idx, value in df_billed[col].items():
                if pd.isna(value) or value == '' or value == ' ':
                    continue
                    
                try:
                    # Try different conversion methods
                    if isinstance(value, (int, float)):
                        # Excel serial date
                        if value > 0 and value < 100000:  # Reasonable range for Excel dates
                            temp_dates[idx] = pd.Timestamp('1899-12-30') + pd.Timedelta(days=value)
                    elif isinstance(value, str):
                        # Clean the string
                        value = value.strip()
                        # Try various date formats
                        for fmt in ['%Y-%m-%d %H:%M', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', 
                                   '%m-%d-%Y', '%m/%d/%Y', '%d/%m/%Y', '%d-%m-%Y']:
                            try:
                                temp_dates[idx] = pd.to_datetime(value, format=fmt)
                                break
                            except:
                                continue
                        # If no format worked, try pandas auto-detection
                        if pd.isna(temp_dates[idx]):
                            temp_dates[idx] = pd.to_datetime(value, errors='coerce')
                    else:
                        # Already a datetime
                        temp_dates[idx] = pd.to_datetime(value, errors='coerce')
                except:
                    # If all else fails, leave as NaT
                    continue
            
            df_billed[col] = temp_dates
    
    # Clean numeric columns
    numeric_columns = ['TOTAL CHARGES', 'PIECES', 'Billable Weight KG', 'Time In Transit', 
                      'WEIGHT(KG)', 'WT LB', 'Billable Weight LB', 'TOT DST', 'AMOUNT']
    for col in numeric_columns:
        if col in df_billed.columns:
            df_billed[col] = pd.to_numeric(df_billed[col], errors='coerce')
    
    # Convert TOTAL CHARGES to EUR (assuming USD to EUR rate of 0.92)
    USD_TO_EUR = 0.92
    df_billed['TOTAL CHARGES EUR'] = df_billed['TOTAL CHARGES'] * USD_TO_EUR
    
    # Extract delivery hour if POD DATE/TIME is available
    if 'POD DATE/TIME' in df_billed.columns:
        df_billed['Delivery Hour'] = pd.to_datetime(df_billed['POD DATE/TIME'], errors='coerce').dt.hour
        df_billed['Delivery Day of Week'] = pd.to_datetime(df_billed['POD DATE/TIME'], errors='coerce').dt.day_name()
    
    return df_billed

# Define controllable QC codes - expanded based on your requirements
def is_controllable_qc(qc_code, qc_name):
    """Determine if a QC code is controllable based on code and name"""
    if pd.isna(qc_code) and pd.isna(qc_name):
        return False
    
    # Convert to string for comparison
    qc_code_str = str(qc_code).lower() if not pd.isna(qc_code) else ''
    qc_name_str = str(qc_name).lower() if not pd.isna(qc_name) else ''
    
    # Keywords that indicate controllable issues
    controllable_keywords = ['custom', 'warehouse', 'w/house', 'agt', 'agent', 'mnx', 
                            'customer', 'shipper', 'consignee', 'data entry', 
                            'documentation', 'pwk', 'clearance']
    
    # Check if any keyword is in the QC name
    for keyword in controllable_keywords:
        if keyword in qc_name_str:
            return True
    
    # Specific QC codes that are controllable
    controllable_codes = ['165', '164', '175', '173', '163', '338', '319', '308', '309', 
                         '326', '262', '287', '199', '203', '278']
    
    if qc_code_str in controllable_codes:
        return True
    
    return False

def calculate_otp(df):
    """Calculate On-Time Performance (OTP) metrics"""
    # Ensure we have the necessary columns
    if 'POD DATE/TIME' not in df.columns or 'QDT' not in df.columns:
        return 0, 0, df
    
    # Create a copy to avoid warnings
    df_calc = df.copy()
    
    # Initialize columns
    df_calc['On_Time'] = False
    df_calc['Is_Controllable'] = False
    
    # Gross OTP: All shipments delivered on time
    # Only calculate for rows where both dates are available
    valid_dates = df_calc['POD DATE/TIME'].notna() & df_calc['QDT'].notna()
    if valid_dates.sum() > 0:
        df_calc.loc[valid_dates, 'On_Time'] = (
            pd.to_datetime(df_calc.loc[valid_dates, 'POD DATE/TIME']) <= 
            pd.to_datetime(df_calc.loc[valid_dates, 'QDT'])
        )
        gross_otp = (df_calc.loc[valid_dates, 'On_Time'].sum() / valid_dates.sum()) * 100
    else:
        gross_otp = 0
    
    # Determine controllable QC issues
    if 'QCCODE' in df_calc.columns and 'QC NAME' in df_calc.columns:
        df_calc['Is_Controllable'] = df_calc.apply(
            lambda row: is_controllable_qc(row['QCCODE'], row['QC NAME']), axis=1
        )
    
    # Net OTP: Excluding non-controllable delays
    # For net OTP, we only count delays that are non-controllable against us
    if valid_dates.sum() > 0:
        # Shipments that are either on-time OR have non-controllable delays
        net_eligible = df_calc[valid_dates & (~df_calc['Is_Controllable'] | df_calc['On_Time'])]
        net_otp = (net_eligible['On_Time'].sum() / len(net_eligible)) * 100 if len(net_eligible) > 0 else gross_otp
    else:
        net_otp = gross_otp
    
    # Add the calculated columns back to the original dataframe
    df['On_Time'] = df_calc['On_Time']
    df['Is_Controllable'] = df_calc['Is_Controllable']
    
    return gross_otp, net_otp, df

# Title and header
st.title("📊 Shipment Performance Dashboard")
st.markdown("**Executive Overview - 440-BILLED Shipments Analysis**")

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
        valid_dates = df['ORD CREATE'].notna()
        if valid_dates.any():
            min_date = df.loc[valid_dates, 'ORD CREATE'].min()
            max_date = df.loc[valid_dates, 'ORD CREATE'].max()
            if pd.notna(min_date) and pd.notna(max_date):
                date_range = st.date_input(
                    "Select Period",
                    value=(min_date.date(), max_date.date()),
                    format="DD/MM/YYYY"
                )
            else:
                date_range = None
        else:
            st.warning("No valid dates found in data")
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

if date_range and len(date_range) == 2 and 'ORD CREATE' in df_filtered.columns:
    mask = (pd.to_datetime(df_filtered['ORD CREATE']).dt.date >= date_range[0]) & \
           (pd.to_datetime(df_filtered['ORD CREATE']).dt.date <= date_range[1])
    df_filtered = df_filtered[mask]

# Calculate OTP metrics
gross_otp, net_otp, df_filtered = calculate_otp(df_filtered)

# Key Performance Indicators
st.markdown("## 📈 Key Performance Indicators")
st.markdown("<div class='info-box'><b>ℹ️ Note:</b> All financial metrics are converted from USD to EUR at rate 0.92. Revenue is calculated from TOTAL CHARGES column.</div>", unsafe_allow_html=True)

col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    total_shipments = len(df_filtered)
    st.metric(
        label="Total Shipments",
        value=f"{total_shipments:,}",
        delta=f"440-BILLED only"
    )

with col2:
    total_revenue = df_filtered['TOTAL CHARGES EUR'].sum()
    avg_revenue = total_revenue/total_shipments if total_shipments > 0 else 0
    st.metric(
        label="Total Revenue (€)",
        value=f"€{total_revenue:,.0f}",
        delta=f"Avg: €{avg_revenue:,.0f}"
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
        delta="Excl. non-controllable",
        delta_color="normal" if net_otp >= 95 else "inverse"
    )

with col5:
    avg_transit = df_filtered['Time In Transit'].mean() if 'Time In Transit' in df_filtered.columns else 0
    st.metric(
        label="Avg Transit Time",
        value=f"{avg_transit:.1f} days" if avg_transit and not pd.isna(avg_transit) else "N/A",
        delta="Days in transit"
    )

with col6:
    total_weight = df_filtered['Billable Weight KG'].sum()
    avg_weight = total_weight/total_shipments if total_shipments > 0 else 0
    st.metric(
        label="Total Weight",
        value=f"{total_weight:,.0f} KG",
        delta=f"Avg: {avg_weight:,.0f} KG"
    )

# OTP Explanation
with st.expander("📖 Understanding OTP Metrics - Click to Expand"):
    st.markdown("""
    ### On-Time Performance (OTP) Explained
    
    **🎯 Gross OTP:** 
    - Percentage of shipments delivered on or before the Quoted Delivery Time (QDT)
    - Includes ALL delays regardless of cause
    - Calculated as: (On-time deliveries / Total deliveries with valid dates) × 100
    - Industry benchmark: >95%
    
    **✅ Net OTP:** 
    - Excludes delays due to non-controllable factors from the calculation
    - **Controllable factors** (impact our performance):
        - Customer-related delays (changed parameters, not ready)
        - Customs delays (documentation, clearance)
        - Warehouse/Agent issues (data entry, late pickup/delivery)
        - MNX operational errors
    - **Non-controllable factors** (external):
        - Airline delays (RTA, slow offload)
        - Weather conditions
        - Force majeure events
    - Shows true operational efficiency within your control
    
    **📊 Why the difference matters:**
    - **Gap = Net OTP - Gross OTP** indicates external factor impact
    - Large gap suggests many delays are outside your control
    - Small gap means most delays are operational (controllable)
    """)

# Service Analysis
st.markdown("## 🚚 Service Type Analysis")
st.markdown("<div class='info-box'><b>What this shows:</b> Distribution of shipments and revenue across different service types (SVC codes). Helps identify which services drive volume vs. revenue.</div>", unsafe_allow_html=True)

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
        title='Shipment Volume Distribution by Service',
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
st.markdown("<div class='info-box'><b>What this shows:</b> Performance metrics by departure airport (DEP). Identifies high-performing routes and problem areas for operational focus.</div>", unsafe_allow_html=True)

if 'DEP' in df_filtered.columns:
    dep_analysis = df_filtered.groupby('DEP').agg({
        'REFER': 'count',
        'TOTAL CHARGES EUR': 'sum'
    }).reset_index()
    dep_analysis.columns = ['Airport', 'Shipments', 'Revenue']
    
    # Add transit time if available
    if 'Time In Transit' in df_filtered.columns:
        transit_by_dep = df_filtered.groupby('DEP')['Time In Transit'].mean().reset_index()
        transit_by_dep.columns = ['Airport', 'Avg Transit Days']
        dep_analysis = dep_analysis.merge(transit_by_dep, on='Airport', how='left')
    else:
        dep_analysis['Avg Transit Days'] = 0
    
    # Add OTP if available
    if 'On_Time' in df_filtered.columns:
        otp_by_dep = df_filtered.groupby('DEP')['On_Time'].apply(
            lambda x: (x.sum() / len(x) * 100) if len(x) > 0 else 0
        ).reset_index()
        otp_by_dep.columns = ['Airport', 'OTP %']
        dep_analysis = dep_analysis.merge(otp_by_dep, on='Airport', how='left')
    else:
        dep_analysis['OTP %'] = 0
    
    dep_analysis = dep_analysis.sort_values('Shipments', ascending=False).head(15)
    
    # Create subplots for airport analysis
    fig_airport = make_subplots(
        rows=2, cols=2,
        subplot_titles=('Shipment Volume by Airport', 'Revenue by Airport (€)', 
                        'Average Transit Time (Days)', 'On-Time Performance (%)'),
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
    
    # OTP percentage with target line
    fig_airport.add_trace(
        go.Scatter(x=dep_analysis['Airport'], y=dep_analysis['OTP %'], 
                   mode='lines+markers', name='OTP %', 
                   marker=dict(size=10, color='purple')),
        row=2, col=2
    )
    
    # Add 95% target line for OTP
    fig_airport.add_hline(y=95, line_dash="dash", line_color="red", 
                         annotation_text="Target", row=2, col=2)
    
    fig_airport.update_layout(height=800, showlegend=False)
    fig_airport.update_xaxes(tickangle=45)
    st.plotly_chart(fig_airport, use_container_width=True)
else:
    st.info("Departure airport analysis requires DEP column in the data")

# Delivery Time Analysis
st.markdown("## 🕐 Delivery Time Analysis")
st.markdown("<div class='info-box'><b>What this shows:</b> When shipments are actually delivered (hour and day). Helps optimize delivery scheduling and resource allocation.</div>", unsafe_allow_html=True)

if 'Delivery Hour' in df_filtered.columns and df_filtered['Delivery Hour'].notna().any():
    col1, col2 = st.columns(2)
    
    with col1:
        # Hourly delivery distribution
        hourly_deliveries = df_filtered['Delivery Hour'].value_counts().sort_index()
        fig_hourly = px.bar(
            x=hourly_deliveries.index,
            y=hourly_deliveries.values,
            title='Deliveries by Hour of Day',
            labels={'x': 'Hour (24h format)', 'y': 'Number of Deliveries'},
            color=hourly_deliveries.values,
            color_continuous_scale='Blues'
        )
        fig_hourly.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig_hourly, use_container_width=True)
    
    with col2:
        # Day of week distribution
        if 'Delivery Day of Week' in df_filtered.columns:
            day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            day_deliveries = df_filtered['Delivery Day of Week'].value_counts()
            day_deliveries = day_deliveries.reindex(day_order, fill_value=0)
            
            fig_daily = px.bar(
                x=day_deliveries.index,
                y=day_deliveries.values,
                title='Deliveries by Day of Week',
                labels={'x': 'Day', 'y': 'Number of Deliveries'},
                color=day_deliveries.values,
                color_continuous_scale='Greens'
            )
            fig_daily.update_layout(height=400, showlegend=False)
            st.plotly_chart(fig_daily, use_container_width=True)
else:
    st.info("Delivery time analysis requires POD DATE/TIME data")

# Time Series Analysis
st.markdown("## 📅 Temporal Trends")
st.markdown("<div class='info-box'><b>What this shows:</b> Trends over time for volume, revenue, and OTP. Identifies seasonality and growth patterns.</div>", unsafe_allow_html=True)

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
        monthly_otp = df_time.groupby('Month')['On_Time'].apply(
            lambda x: (x.sum() / len(x) * 100) if len(x) > 0 else 0
        ).reset_index()
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
    st.info("Time series analysis requires valid order creation dates")

# Quality Control Analysis
st.markdown("## 🔍 Quality Control Analysis")
st.markdown("<div class='info-box'><b>What this shows:</b> Root causes of delays (QC codes). Controllable issues include customs, warehouse, and agent-related delays.</div>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    # QC distribution
    if 'QC NAME' in df_filtered.columns:
        qc_dist = df_filtered[df_filtered['QC NAME'].notna()].groupby('QC NAME').size().reset_index(name='Count')
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
    else:
        st.info("QC analysis requires QC NAME column")

with col2:
    # Controllable vs Non-controllable
    if 'QCCODE' in df_filtered.columns or 'QC NAME' in df_filtered.columns:
        qc_with_codes = df_filtered[(df_filtered['QCCODE'].notna()) | (df_filtered['QC NAME'].notna())]
        
        if 'Is_Controllable' in qc_with_codes.columns:
            controllable_count = qc_with_codes['Is_Controllable'].sum()
            non_controllable = len(qc_with_codes) - controllable_count
        else:
            # Recalculate if not present
            qc_with_codes['Is_Controllable'] = qc_with_codes.apply(
                lambda row: is_controllable_qc(row.get('QCCODE'), row.get('QC NAME')), axis=1
            )
            controllable_count = qc_with_codes['Is_Controllable'].sum()
            non_controllable = len(qc_with_codes) - controllable_count
        
        fig_control = go.Figure(data=[
            go.Bar(name='Controllable', x=['QC Issues'], y=[controllable_count], 
                   marker_color='orange', text=controllable_count, textposition='outside'),
            go.Bar(name='Non-Controllable', x=['QC Issues'], y=[non_controllable], 
                   marker_color='gray', text=non_controllable, textposition='outside')
        ])
        fig_control.update_layout(
            title='Controllable vs Non-Controllable Issues',
            barmode='stack',
            height=400,
            showlegend=True
        )
        st.plotly_chart(fig_control, use_container_width=True)
    else:
        st.info("Quality control analysis requires QC data")

# Cost Analysis
st.markdown("## 💰 Financial Performance")
st.markdown("<div class='info-box'><b>What this shows:</b> Cost distribution and weight-to-cost relationship. Identifies pricing efficiency and outliers.</div>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    # Cost distribution histogram
    if 'TOTAL CHARGES EUR' in df_filtered.columns:
        # Filter outliers for better visualization
        charges_for_viz = df_filtered[df_filtered['TOTAL CHARGES EUR'] < df_filtered['TOTAL CHARGES EUR'].quantile(0.95)]
        
        fig_cost_dist = px.histogram(
            charges_for_viz,
            x='TOTAL CHARGES EUR',
            nbins=50,
            title='Shipment Cost Distribution (€) - 95th Percentile',
            labels={'TOTAL CHARGES EUR': 'Cost (€)', 'count': 'Number of Shipments'}
        )
        fig_cost_dist.update_layout(height=400)
        st.plotly_chart(fig_cost_dist, use_container_width=True)
    else:
        st.info("Cost analysis requires TOTAL CHARGES data")

with col2:
    # Weight vs Cost scatter (without trendline to avoid statsmodels dependency)
    if 'Billable Weight KG' in df_filtered.columns and 'TOTAL CHARGES EUR' in df_filtered.columns:
        # Filter for visualization
        weight_cost_viz = df_filtered[
            (df_filtered['Billable Weight KG'] < df_filtered['Billable Weight KG'].quantile(0.95)) &
            (df_filtered['TOTAL CHARGES EUR'] < df_filtered['TOTAL CHARGES EUR'].quantile(0.95))
        ]
        
        fig_weight_cost = px.scatter(
            weight_cost_viz,
            x='Billable Weight KG',
            y='TOTAL CHARGES EUR',
            color='SVC' if 'SVC' in weight_cost_viz.columns else None,
            title='Weight vs Cost Analysis (95th Percentile)',
            labels={'Billable Weight KG': 'Weight (KG)', 'TOTAL CHARGES EUR': 'Cost (€)'},
            hover_data=['REFER'] if 'REFER' in weight_cost_viz.columns else None
        )
        fig_weight_cost.update_layout(height=400)
        st.plotly_chart(fig_weight_cost, use_container_width=True)
    else:
        st.info("Weight-cost analysis requires weight and cost data")

# Route Performance Matrix
st.markdown("## 🗺️ Route Performance Matrix")
st.markdown("<div class='info-box'><b>What this shows:</b> Performance comparison between departure (DEP) and arrival (ARR) airports. Darker colors indicate more shipments.</div>", unsafe_allow_html=True)

if 'DEP' in df_filtered.columns and 'ARR' in df_filtered.columns:
    # Create route matrix
    route_matrix = df_filtered.groupby(['DEP', 'ARR']).size().reset_index(name='Shipments')
    
    # Get top routes
    top_routes = route_matrix.nlargest(20, 'Shipments')
    
    # Create pivot table for heatmap
    pivot_routes = top_routes.pivot_table(
        index='DEP',
        columns='ARR',
        values='Shipments',
        fill_value=0
    )
    
    fig_heatmap = px.imshow(
        pivot_routes,
        labels=dict(x="Arrival Airport", y="Departure Airport", color="Shipments"),
        title="Top 20 Routes by Volume",
        color_continuous_scale='YlOrRd'
    )
    fig_heatmap.update_layout(height=600)
    st.plotly_chart(fig_heatmap, use_container_width=True)
else:
    st.info("Route analysis requires DEP and ARR columns")

# Data Quality Report
st.markdown("## 📊 Data Quality Report")
st.markdown("<div class='info-box'><b>What this shows:</b> Completeness of critical data fields. Helps identify data quality issues affecting analysis accuracy.</div>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    # Calculate data completeness
    critical_fields = ['ORD CREATE', 'POD DATE/TIME', 'QDT', 'TOTAL CHARGES', 
                       'DEP', 'ARR', 'SVC', 'Billable Weight KG']
    
    completeness_data = []
    for field in critical_fields:
        if field in df_filtered.columns:
            complete_pct = (df_filtered[field].notna().sum() / len(df_filtered)) * 100
            completeness_data.append({'Field': field, 'Completeness %': complete_pct})
    
    if completeness_data:
        completeness_df = pd.DataFrame(completeness_data)
        
        fig_complete = px.bar(
            completeness_df,
            x='Completeness %',
            y='Field',
            orientation='h',
            title='Data Field Completeness',
            color='Completeness %',
            color_continuous_scale='RdYlGn',
            range_color=[0, 100]
        )
        fig_complete.update_layout(height=400)
        st.plotly_chart(fig_complete, use_container_width=True)

with col2:
    # Summary statistics
    st.markdown("### 📈 Summary Statistics")
    
    total_records = len(df_filtered)
    records_with_dates = df_filtered[df_filtered['POD DATE/TIME'].notna()].shape[0] if 'POD DATE/TIME' in df_filtered.columns else 0
    records_with_qc = df_filtered[df_filtered['QCCODE'].notna()].shape[0] if 'QCCODE' in df_filtered.columns else 0
    
    st.markdown(f"""
    - **Total Records:** {total_records:,}
    - **Records with Delivery Date:** {records_with_dates:,} ({records_with_dates/total_records*100:.1f}%)
    - **Records with QC Codes:** {records_with_qc:,} ({records_with_qc/total_records*100:.1f}%)
    - **Unique Services:** {df_filtered['SVC'].nunique() if 'SVC' in df_filtered.columns else 0}
    - **Unique Departure Airports:** {df_filtered['DEP'].nunique() if 'DEP' in df_filtered.columns else 0}
    - **Unique Arrival Airports:** {df_filtered['ARR'].nunique() if 'ARR' in df_filtered.columns else 0}
    """)

# Executive Summary
st.markdown("## 📋 Executive Summary & Recommendations")

with st.container():
    st.markdown("### 🎯 Key Findings:")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        **Performance Metrics:**
        - Total shipments processed: **{total_shipments:,}**
        - Total revenue generated: **€{total_revenue:,.0f}**
        - Gross OTP: **{gross_otp:.1f}%** {'✅' if gross_otp >= 95 else '⚠️'}
        - Net OTP: **{net_otp:.1f}%** {'✅' if net_otp >= 95 else '⚠️'}
        - OTP Gap: **{abs(net_otp - gross_otp):.1f}%** (impact of external factors)
        
        **Cost & Weight Analysis:**
        - Average shipment cost: **€{avg_revenue:,.2f}**
        - Average shipment weight: **{avg_weight:,.2f} KG**
        - Cost per KG: **€{(total_revenue/total_weight):.2f}** (if weight data available)
        """)
        
    with col2:
        # Top performing routes
        if 'On_Time' in df_filtered.columns and 'DEP' in df_filtered.columns:
            top_routes = df_filtered.groupby('DEP')['On_Time'].mean().sort_values(ascending=False).head(3)
            if not top_routes.empty:
                st.markdown("**🏆 Top Performing Routes:**")
                for route, otp in top_routes.items():
                    st.markdown(f"- {route}: {otp*100:.1f}% OTP")
            
            # Bottom performing routes
            bottom_routes = df_filtered.groupby('DEP')['On_Time'].mean().sort_values(ascending=True).head(3)
            if not bottom_routes.empty:
                st.markdown("**⚠️ Routes Needing Attention:**")
                for route, otp in bottom_routes.items():
                    st.markdown(f"- {route}: {otp*100:.1f}% OTP")
        
        # Top QC issues if available
        if 'QC NAME' in df_filtered.columns:
            top_qc = df_filtered['QC NAME'].value_counts().head(3)
            if not top_qc.empty:
                st.markdown("**🔍 Top Quality Issues:**")
                for issue, count in top_qc.items():
                    if pd.notna(issue):
                        st.markdown(f"- {issue}: {count} occurrences")

    st.markdown("""
    ### 🎯 Strategic Recommendations:
    
    1. **📈 Improve OTP Performance**
       - Focus on controllable factors: customs documentation, warehouse operations, and agent coordination
       - Implement proactive monitoring for shipments approaching QDT
       - Target routes with OTP below 95% for immediate improvement
       
    2. **✈️ Route Optimization**
       - Review underperforming departure airports and consider alternative routing
       - Increase capacity allocation on high-performing, high-revenue routes
       - Investigate root causes for delays at specific airports
       
    3. **💰 Revenue Management**
       - Analyze high-cost outliers for pricing optimization opportunities
       - Review cost-per-KG ratios across different service types
       - Focus on high-margin services while maintaining volume
       
    4. **🔧 Operational Excellence**
       - Address top controllable QC issues through:
         * Enhanced customs documentation processes
         * Improved warehouse data entry accuracy
         * Better coordination with agents (AGT)
       - Implement preventive measures for recurring delays
       - Optimize delivery scheduling based on hourly/daily patterns
       
    5. **📊 Data Quality Improvement**
       - Ensure complete capture of POD dates for accurate OTP calculation
       - Standardize QC code recording for better root cause analysis
       - Improve data completeness for critical fields below 90%
    """)

# Footer
st.markdown("---")
st.markdown(f"""
<div style='text-align: center; color: #666;'>
    <p><b>Dashboard Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
    <p><b>Data Source:</b> Excel File | <b>Status Filter:</b> 440-BILLED Only</p>
    <p><b>Currency:</b> All amounts in EUR (€) | USD→EUR Rate: 0.92</p>
    <p><b>Data Coverage:</b> {df_filtered['ORD CREATE'].min().strftime('%Y-%m-%d') if 'ORD CREATE' in df_filtered.columns and df_filtered['ORD CREATE'].notna().any() else 'N/A'} to {df_filtered['ORD CREATE'].max().strftime('%Y-%m-%d') if 'ORD CREATE' in df_filtered.columns and df_filtered['ORD CREATE'].notna().any() else 'N/A'}</p>
</div>
""", unsafe_allow_html=True)
