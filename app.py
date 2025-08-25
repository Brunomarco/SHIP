import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime

# Page configuration - MUST be first
st.set_page_config(
    page_title="Shipment Cost Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Reduce default height for performance
px.defaults.height = 400

# Simple CSS
st.markdown("""
    <style>
    .main {padding: 0rem 1rem;}
    h1 {color: #1f2937; font-weight: 700; border-bottom: 3px solid #3b82f6; padding-bottom: 10px;}
    h2 {color: #374151; font-weight: 600; margin-top: 2rem;}
    .stMetric {background-color: #f0f2f6; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);}
    </style>
    """, unsafe_allow_html=True)

# Title
st.title("📊 Shipment Cost Analytics Dashboard")
st.markdown("### Executive Overview - Facts & Figures")
st.markdown("---")

# Constants
CONTROLLABLE_QC_CODES = [262, 287, 183, 197, 199, 308, 309, 319, 326, 278, 203]
USD_TO_EUR = 0.92

@st.cache_data
def load_and_process_data(file):
    """Load and process Excel data"""
    try:
        # Read Excel
        df = pd.read_excel(file, sheet_name=0)
        
        # Filter 440-BILLED only
        df = df[df['STATUS'] == '440-BILLED'].copy()
        
        # Convert essential dates only
        date_cols = ['QDT', 'POD DATE/TIME', 'ORD CREATE', 'Depart Date / Time', 'Arrive Date / Time']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Currency conversion
        if 'TOTAL CHARGES' in df.columns:
            df['TOTAL_CHARGES_EUR'] = pd.to_numeric(df['TOTAL CHARGES'], errors='coerce') * USD_TO_EUR
        else:
            df['TOTAL_CHARGES_EUR'] = 0
            
        # Create route if possible
        if 'DEP' in df.columns and 'ARR' in df.columns:
            df['Route'] = df['DEP'].astype(str) + ' → ' + df['ARR'].astype(str)
        
        # Extract month
        if 'ORD CREATE' in df.columns:
            df['Month'] = df['ORD CREATE'].dt.to_period('M').astype(str)
            
        # Calculate transit time if dates available
        if 'Depart Date / Time' in df.columns and 'Arrive Date / Time' in df.columns:
            df['Transit_Hours'] = (df['Arrive Date / Time'] - df['Depart Date / Time']).dt.total_seconds() / 3600
        
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def calculate_otp(df):
    """Calculate OTP metrics"""
    if 'QDT' not in df.columns or 'POD DATE/TIME' not in df.columns:
        return 0, 0
    
    df_valid = df.dropna(subset=['QDT', 'POD DATE/TIME']).copy()
    if len(df_valid) == 0:
        return 0, 0
    
    df_valid['ON_TIME_GROSS'] = df_valid['POD DATE/TIME'] <= df_valid['QDT']
    df_valid['LATE'] = ~df_valid['ON_TIME_GROSS']
    
    if 'QCCODE' in df.columns:
        df_valid['CONTROLLABLE'] = df_valid['QCCODE'].isin(CONTROLLABLE_QC_CODES)
        df_valid['ON_TIME_NET'] = df_valid['ON_TIME_GROSS'] | (df_valid['LATE'] & ~df_valid['CONTROLLABLE'])
    else:
        df_valid['ON_TIME_NET'] = df_valid['ON_TIME_GROSS']
    
    gross = (df_valid['ON_TIME_GROSS'].sum() / len(df_valid) * 100)
    net = (df_valid['ON_TIME_NET'].sum() / len(df_valid) * 100)
    
    return gross, net

# File uploader
uploaded_file = st.file_uploader(
    "Upload your Shipment Excel file",
    type=['xlsx', 'xls'],
    help="Upload the Excel file with shipment data"
)

if uploaded_file is not None:
    # Load data
    df = load_and_process_data(uploaded_file)
    
    if df is not None:
        # Calculate metrics
        gross_otp, net_otp = calculate_otp(df)
        
        # Top KPIs Row
        st.markdown("### 🎯 Key Performance Indicators")
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        with col1:
            st.metric("📦 Total Shipments", f"{len(df):,}")
        
        with col2:
            total_cost = df['TOTAL_CHARGES_EUR'].sum()
            st.metric("💰 Total Cost", f"€{total_cost:,.0f}")
        
        with col3:
            avg_cost = df['TOTAL_CHARGES_EUR'].mean()
            st.metric("📊 Avg Cost", f"€{avg_cost:,.2f}")
        
        with col4:
            # Format OTP values to show full number without truncation
            gross_display = f"{gross_otp:.1f}%"
            delta_display = f"{gross_otp-90:.1f}%"
            st.metric("✅ OTP Gross", gross_display, 
                     delta=delta_display if gross_otp != 0 else None)
        
        with col5:
            # Format Net OTP to show full number
            net_display = f"{net_otp:.1f}%"
            delta_net = f"+{net_otp-gross_otp:.1f}%" if net_otp > gross_otp else f"{net_otp-gross_otp:.1f}%"
            st.metric("🎯 OTP Net", net_display,
                     delta=delta_net if net_otp != gross_otp else None)
        
        with col6:
            improvement_potential = net_otp - gross_otp
            improve_display = f"{improvement_potential:.1f}%"
            st.metric("📈 Improvement", improve_display,
                     help="Potential OTP improvement by addressing controllable issues")
        
        st.markdown("---")
        
        # Row 1: Service Distribution and OTP Analysis
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📊 Service Type Distribution (Descending)")
            if 'SVC' in df.columns:
                # Get service counts and descriptions
                svc_counts = df['SVC'].value_counts().reset_index()
                svc_counts.columns = ['Service', 'Count']
                
                # Add descriptions if available
                if 'SVCDESC' in df.columns:
                    svc_desc_map = df.groupby('SVC')['SVCDESC'].first().to_dict()
                    svc_counts['Description'] = svc_counts['Service'].map(svc_desc_map)
                else:
                    svc_counts['Description'] = svc_counts['Service']
                
                # Sort for horizontal bar (ascending for correct visual display)
                svc_top10 = svc_counts.head(10).sort_values('Count', ascending=True)
                
                fig_svc = px.bar(
                    svc_top10,
                    x='Count',
                    y='Service',
                    orientation='h',
                    color='Count',
                    text='Count',
                    hover_data=['Description'],
                    color_continuous_scale='Viridis'
                )
                fig_svc.update_traces(texttemplate='%{text}', textposition='outside')
                fig_svc.update_layout(
                    height=400,
                    xaxis_title="Number of Shipments",
                    yaxis_title="Service Type",
                    showlegend=False,
                    coloraxis_showscale=False
                )
                st.plotly_chart(fig_svc, use_container_width=True)
            else:
                st.info("No service data available")
        
        with col2:
            st.subheader("🎯 OTP Performance Analysis")
            
            # Create detailed OTP breakdown
            otp_breakdown = pd.DataFrame({
                'Category': ['Gross OTP', 'Net OTP', 'Controllable Impact'],
                'Value': [gross_otp, net_otp, net_otp - gross_otp],
                'Type': ['Actual', 'Adjusted', 'Opportunity']
            })
            
            fig_otp = go.Figure()
            
            colors = ['#3b82f6', '#10b981', '#fbbf24']
            for i, row in otp_breakdown.iterrows():
                fig_otp.add_trace(go.Bar(
                    x=[row['Value']],
                    y=[row['Category']],
                    orientation='h',
                    name=row['Category'],
                    marker_color=colors[i],
                    text=f"{row['Value']:.1f}%",
                    textposition='outside',
                    showlegend=False
                ))
            
            fig_otp.update_layout(
                height=400,
                xaxis_title="Percentage (%)",
                yaxis_title="",
                xaxis=dict(range=[0, 105]),
                barmode='overlay'
            )
            fig_otp.add_vline(x=90, line_dash="dash", line_color="red")
            st.plotly_chart(fig_otp, use_container_width=True)
        
        # Row 2: Departure and Delivery Analysis
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🛫 Top 15 Departure Points (DEP)")
            if 'DEP' in df.columns:
                # Clean and standardize DEP data (convert to uppercase to avoid duplicates)
                df['DEP_CLEAN'] = df['DEP'].str.upper().str.strip()
                dep_counts = df.groupby('DEP_CLEAN').agg({
                    'REFER': 'count',
                    'TOTAL_CHARGES_EUR': 'mean'
                }).reset_index()
                dep_counts.columns = ['Location', 'Shipments', 'Avg_Cost']
                dep_counts = dep_counts.sort_values('Shipments', ascending=False).head(15)
                
                # Sort ascending for correct horizontal display
                dep_counts_sorted = dep_counts.sort_values('Shipments', ascending=True)
                
                fig_dep = px.bar(
                    dep_counts_sorted,
                    x='Shipments',
                    y='Location',
                    orientation='h',
                    color='Avg_Cost',
                    text='Shipments',
                    hover_data={'Avg_Cost': ':.2f'},
                    color_continuous_scale='Blues',
                    labels={'Avg_Cost': 'Avg Cost (€)'}
                )
                fig_dep.update_traces(texttemplate='%{text}', textposition='outside')
                fig_dep.update_layout(
                    height=450,
                    xaxis_title="Number of Shipments",
                    yaxis_title="Departure Location",
                    showlegend=False
                )
                st.plotly_chart(fig_dep, use_container_width=True)
            else:
                st.info("No departure data available")
        
        with col2:
            st.subheader("📍 Top 15 Delivery Points (ARR)")
            if 'ARR' in df.columns:
                # Clean and standardize ARR data (convert to uppercase to avoid duplicates)
                df['ARR_CLEAN'] = df['ARR'].str.upper().str.strip()
                arr_counts = df.groupby('ARR_CLEAN').agg({
                    'REFER': 'count',
                    'TOTAL_CHARGES_EUR': 'mean'
                }).reset_index()
                arr_counts.columns = ['Location', 'Shipments', 'Avg_Cost']
                arr_counts = arr_counts.sort_values('Shipments', ascending=False).head(15)
                
                # Sort ascending for correct horizontal display
                arr_counts_sorted = arr_counts.sort_values('Shipments', ascending=True)
                
                fig_arr = px.bar(
                    arr_counts_sorted,
                    x='Shipments',
                    y='Location',
                    orientation='h',
                    color='Avg_Cost',
                    text='Shipments',
                    hover_data={'Avg_Cost': ':.2f'},
                    color_continuous_scale='Greens',
                    labels={'Avg_Cost': 'Avg Cost (€)'}
                )
                fig_arr.update_traces(texttemplate='%{text}', textposition='outside')
                fig_arr.update_layout(
                    height=450,
                    xaxis_title="Number of Shipments",
                    yaxis_title="Delivery Location",
                    showlegend=False
                )
                st.plotly_chart(fig_arr, use_container_width=True)
            else:
                st.info("No delivery data available")
        
        # Row 3: Quality Control Analysis
        st.markdown("---")
        st.subheader("🔧 Quality Control Issues Analysis")
        
        if 'QCCODE' in df.columns and 'QC NAME' in df.columns:
            qc_data = df[df['QCCODE'].notna()].copy()
            
            if len(qc_data) > 0:
                # Classify QC codes
                qc_data['Issue_Type'] = qc_data['QCCODE'].apply(
                    lambda x: 'Controllable (Internal)' if x in CONTROLLABLE_QC_CODES else 'Non-Controllable (External)'
                )
                
                col1, col2 = st.columns(2)
                
                with col1:
                    # QC Issues by Type
                    st.markdown("#### Distribution by Control Type")
                    
                    type_counts = qc_data['Issue_Type'].value_counts()
                    
                    fig_qc_pie = px.pie(
                        values=type_counts.values,
                        names=type_counts.index,
                        color_discrete_map={
                            'Controllable (Internal)': '#10b981',
                            'Non-Controllable (External)': '#ef4444'
                        },
                        hole=0.4
                    )
                    fig_qc_pie.update_traces(textposition='inside', textinfo='percent+label')
                    fig_qc_pie.update_layout(height=350)
                    st.plotly_chart(fig_qc_pie, use_container_width=True)
                    
                    # Metrics
                    col1a, col1b = st.columns(2)
                    with col1a:
                        controllable_pct = (type_counts.get('Controllable (Internal)', 0) / len(qc_data) * 100)
                        st.metric("🟢 Controllable", f"{controllable_pct:.1f}%",
                                help="Issues that can be addressed internally")
                    with col1b:
                        non_controllable_pct = (type_counts.get('Non-Controllable (External)', 0) / len(qc_data) * 100)
                        st.metric("🔴 Non-Controllable", f"{non_controllable_pct:.1f}%",
                                help="External factors beyond direct control")
                
                with col2:
                    # Top QC Issues
                    st.markdown("#### Top 10 Quality Control Issues")
                    
                    qc_counts = qc_data.groupby(['QC NAME', 'Issue_Type']).size().reset_index(name='Count')
                    qc_counts = qc_counts.sort_values('Count', ascending=False).head(10)
                    # Sort ascending for horizontal display
                    qc_counts_sorted = qc_counts.sort_values('Count', ascending=True)
                    
                    fig_qc_bar = px.bar(
                        qc_counts_sorted,
                        x='Count',
                        y='QC NAME',
                        orientation='h',
                        color='Issue_Type',
                        text='Count',
                        color_discrete_map={
                            'Controllable (Internal)': '#10b981',
                            'Non-Controllable (External)': '#ef4444'
                        }
                    )
                    fig_qc_bar.update_traces(texttemplate='%{text}', textposition='outside')
                    fig_qc_bar.update_layout(
                        height=350,
                        xaxis_title="Number of Occurrences",
                        yaxis_title="",
                        legend_title="Issue Type",
                        legend=dict(orientation="h", yanchor="bottom", y=-0.5, x=0),
                        margin=dict(b=100)  # Add bottom margin for legend
                    )
                    st.plotly_chart(fig_qc_bar, use_container_width=True)
            else:
                st.info("No quality control issues found")
        else:
            st.info("No QC data available")
        
        # Row 4: Cost Distribution and Route Analysis
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("💶 Cost Distribution Analysis")
            
            # Create cost segments
            cost_segments = pd.cut(
                df['TOTAL_CHARGES_EUR'],
                bins=[0, 250, 500, 1000, 2500, 5000, float('inf')],
                labels=['<€250', '€250-500', '€500-1K', '€1K-2.5K', '€2.5K-5K', '>€5K']
            ).value_counts().sort_index()
            
            # Create pie chart for cost distribution
            fig_cost_dist = px.pie(
                values=cost_segments.values,
                names=cost_segments.index,
                color_discrete_sequence=px.colors.sequential.Viridis,
                hole=0.4  # Makes it a donut chart
            )
            
            fig_cost_dist.update_traces(
                textposition='inside',
                textinfo='percent+label',
                hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
            )
            
            fig_cost_dist.update_layout(
                height=400,
                title="Shipments by Cost Range",
                showlegend=True,
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1)
            )
            st.plotly_chart(fig_cost_dist, use_container_width=True)
        
        with col2:
            st.subheader("🛣️ Top 10 Routes by Volume")
            if 'Route' in df.columns:
                route_stats = df.groupby('Route').agg({
                    'REFER': 'count',
                    'TOTAL_CHARGES_EUR': ['mean', 'sum']
                }).reset_index()
                route_stats.columns = ['Route', 'Shipments', 'Avg_Cost', 'Total_Cost']
                route_stats = route_stats.sort_values('Shipments', ascending=False).head(10)
                
                # Create bubble chart
                fig_route = px.scatter(
                    route_stats,
                    x='Shipments',
                    y='Avg_Cost',
                    size='Total_Cost',
                    color='Total_Cost',
                    text='Route',
                    color_continuous_scale='Turbo',
                    labels={'Avg_Cost': 'Average Cost (€)', 
                           'Shipments': 'Number of Shipments',
                           'Total_Cost': 'Total Cost (€)'}
                )
                fig_route.update_traces(textposition='top center', textfont_size=9)
                fig_route.update_layout(
                    height=400,
                    showlegend=False
                )
                st.plotly_chart(fig_route, use_container_width=True)
            else:
                st.info("Route data not available")
        
        # Row 5: Time-based Analysis
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📅 Monthly Performance Trend")
            if 'Month' in df.columns:
                monthly = df.groupby('Month').agg({
                    'REFER': 'count',
                    'TOTAL_CHARGES_EUR': 'sum'
                }).reset_index()
                monthly.columns = ['Month', 'Orders', 'Cost']
                
                # Calculate NET OTP by month if possible
                if 'QDT' in df.columns and 'POD DATE/TIME' in df.columns:
                    monthly_net_otp = []
                    for month in monthly['Month']:
                        month_data = df[df['Month'] == month]
                        _, month_net = calculate_otp(month_data)
                        monthly_net_otp.append(month_net)
                    monthly['Net_OTP'] = monthly_net_otp
                
                fig_trend = make_subplots(
                    rows=2, cols=1,
                    subplot_titles=('Shipments & Cost', 'Net OTP Trend' if 'Net_OTP' in monthly.columns else ''),
                    specs=[[{"secondary_y": True}], [{"secondary_y": False}]],
                    row_heights=[0.6, 0.4] if 'Net_OTP' in monthly.columns else [1, 0]
                )
                
                # Shipments and Cost
                fig_trend.add_trace(
                    go.Bar(x=monthly['Month'], y=monthly['Orders'], 
                          name='Orders', marker_color='lightblue'),
                    secondary_y=False, row=1, col=1
                )
                fig_trend.add_trace(
                    go.Scatter(x=monthly['Month'], y=monthly['Cost'], 
                              name='Cost (€)', mode='lines+markers',
                              marker_color='darkblue', line=dict(width=3)),
                    secondary_y=True, row=1, col=1
                )
                
                # Net OTP Trend if available
                if 'Net_OTP' in monthly.columns:
                    fig_trend.add_trace(
                        go.Scatter(x=monthly['Month'], y=monthly['Net_OTP'],
                                  name='Net OTP %', mode='lines+markers',
                                  marker_color='green', line=dict(width=2)),
                        row=2, col=1
                    )
                    fig_trend.add_hline(y=90, line_dash="dash", line_color="red",
                                      annotation_text="Target", row=2, col=1)
                
                fig_trend.update_layout(height=500, hovermode='x unified', showlegend=True)
                fig_trend.update_yaxes(title_text="Orders", secondary_y=False, row=1, col=1)
                fig_trend.update_yaxes(title_text="Cost (€)", secondary_y=True, row=1, col=1)
                if 'Net_OTP' in monthly.columns:
                    fig_trend.update_yaxes(title_text="Net OTP %", row=2, col=1)
                
                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.info("No monthly data available")
        
        with col2:
            st.subheader("⏱️ Transit Time Analysis")
            if 'Transit_Hours' in df.columns:
                # Filter valid transit times
                transit_data = df[df['Transit_Hours'].notna() & (df['Transit_Hours'] > 0) & (df['Transit_Hours'] < 200)]
                
                if len(transit_data) > 0:
                    # Create histogram
                    fig_transit = px.histogram(
                        transit_data,
                        x='Transit_Hours',
                        nbins=30,
                        color_discrete_sequence=['#6366f1']
                    )
                    fig_transit.update_layout(
                        height=400,
                        xaxis_title="Transit Time (Hours)",
                        yaxis_title="Number of Shipments",
                        showlegend=False
                    )
                    
                    # Add average line
                    avg_transit = transit_data['Transit_Hours'].mean()
                    fig_transit.add_vline(x=avg_transit, line_dash="dash", line_color="red",
                                        annotation_text=f"Avg: {avg_transit:.1f}h")
                    
                    st.plotly_chart(fig_transit, use_container_width=True)
                    
                    # Key metrics
                    col2a, col2b, col2c = st.columns(3)
                    with col2a:
                        st.metric("Avg Transit", f"{avg_transit:.1f}h")
                    with col2b:
                        st.metric("Median", f"{transit_data['Transit_Hours'].median():.1f}h")
                    with col2c:
                        st.metric("95th Percentile", f"{transit_data['Transit_Hours'].quantile(0.95):.1f}h")
                else:
                    st.info("No valid transit time data")
            else:
                # Weight analysis as alternative
                st.markdown("#### 📦 Weight Distribution")
                if 'Billable Weight KG' in df.columns:
                    weight_data = df[df['Billable Weight KG'].notna() & (df['Billable Weight KG'] > 0)]
                    
                    if len(weight_data) > 0:
                        # Create weight categories
                        weight_bins = pd.cut(
                            weight_data['Billable Weight KG'],
                            bins=[0, 10, 50, 100, 500, 1000, float('inf')],
                            labels=['<10kg', '10-50kg', '50-100kg', '100-500kg', '500-1000kg', '>1000kg']
                        ).value_counts().sort_index()
                        
                        fig_weight = px.bar(
                            x=weight_bins.index,
                            y=weight_bins.values,
                            color=weight_bins.values,
                            color_continuous_scale='Oranges',
                            text=weight_bins.values
                        )
                        fig_weight.update_traces(texttemplate='%{text}', textposition='outside')
                        fig_weight.update_layout(
                            height=400,
                            xaxis_title="Weight Category",
                            yaxis_title="Number of Shipments",
                            showlegend=False
                        )
                        st.plotly_chart(fig_weight, use_container_width=True)
                    else:
                        st.info("No weight data available")
                else:
                    st.info("No weight data available")
        
        # Executive Summary with actionable insights
        st.markdown("---")
        st.subheader("📋 Executive Summary & Actionable Insights")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("#### 📊 Performance Metrics")
            top_dep = df['DEP'].value_counts().index[0] if 'DEP' in df.columns and len(df['DEP'].value_counts()) > 0 else 'N/A'
            top_route = df['Route'].value_counts().index[0] if 'Route' in df.columns and len(df['Route'].value_counts()) > 0 else 'N/A'
            
            st.markdown(f"""
            **Volume & Cost:**
            - Total shipments: **{len(df):,}**
            - Total cost: **€{df['TOTAL_CHARGES_EUR'].sum():,.0f}**
            - Average cost: **€{df['TOTAL_CHARGES_EUR'].mean():,.2f}**
            
            **Top Locations:**
            - Main departure: **{top_dep}**
            - Busiest route: **{top_route}**
            """)
        
        with col2:
            st.markdown("#### 🎯 OTP Analysis")
            st.markdown(f"""
            **Current Performance:**
            - Gross OTP: **{gross_otp:.1f}%**
            - Net OTP: **{net_otp:.1f}%**
            - Gap to target: **{90-gross_otp:.1f}%**
            
            **Improvement Potential:**
            - Controllable impact: **{net_otp-gross_otp:.1f}%**
            - Status: {'✅ Above target' if gross_otp >= 90 else '⚠️ Below target'}
            """)
        
        with col3:
            st.markdown("#### 💡 Key Actions")
            
            # Calculate controllable issues percentage
            if 'QCCODE' in df.columns:
                qc_issues = df[df['QCCODE'].notna()]
                controllable_issues = qc_issues[qc_issues['QCCODE'].isin(CONTROLLABLE_QC_CODES)]
                controllable_pct = (len(controllable_issues) / len(qc_issues) * 100) if len(qc_issues) > 0 else 0
            else:
                controllable_pct = 0
            
            st.markdown(f"""
            **Recommendations:**
            1. Focus on controllable delays ({controllable_pct:.0f}% of issues)
            2. Optimize top 3 departure points for cost
            3. Review high-cost routes for efficiency
            4. {'Maintain' if gross_otp >= 90 else 'Improve'} OTP performance
            5. Address internal process gaps
            """)
        
        # Data Quality Footer
        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            data_completeness = (1 - df['TOTAL CHARGES'].isna().sum() / len(df)) * 100 if 'TOTAL CHARGES' in df.columns else 0
            st.metric("Data Completeness", f"{data_completeness:.1f}%")
        
        with col2:
            unique_routes = df['Route'].nunique() if 'Route' in df.columns else 0
            st.metric("Unique Routes", f"{unique_routes:,}")
        
        with col3:
            unique_services = df['SVC'].nunique() if 'SVC' in df.columns else 0
            st.metric("Service Types", f"{unique_services}")
        
        with col4:
            # Download button
            csv = df.to_csv(index=False)
            st.download_button(
                label="📥 Download Data",
                data=csv,
                file_name=f"shipment_data_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        # Footer
        st.markdown("---")
        st.caption(f"Dashboard generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {len(df):,} 440-BILLED shipments analyzed")

else:
    st.info("👆 Please upload your Excel file to begin analysis")
    st.markdown("""
    ### Dashboard Features:
    
    **📊 Key Metrics:**
    - Total shipments and costs in EUR
    - OTP (On-Time Performance) Gross vs Net
    - Controllable vs Non-controllable delays
    
    **📈 Visual Analytics:**
    - Service type distribution
    - Top departure (DEP) and delivery (ARR) locations
    - Quality control issues classification
    - Cost distribution analysis
    - Monthly trends and transit times
    - Route performance metrics
    
    **🎯 Executive Insights:**
    - Actionable recommendations
    - Performance gaps identification
    - Cost optimization opportunities
    
    **Requirements:**
    - Excel file with 440-BILLED status shipments
    - Automatic EUR conversion applied
    - All charts in descending order for clarity
    """)
