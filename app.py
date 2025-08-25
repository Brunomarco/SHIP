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
        date_cols = ['QDT', 'POD DATE/TIME', 'ORD CREATE']
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
        
        # Display metrics
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("📦 Total Shipments", f"{len(df):,}")
        
        with col2:
            total_cost = df['TOTAL_CHARGES_EUR'].sum()
            st.metric("💰 Total Cost", f"€{total_cost:,.0f}")
        
        with col3:
            avg_cost = df['TOTAL_CHARGES_EUR'].mean()
            st.metric("📊 Avg Cost", f"€{avg_cost:,.2f}")
        
        with col4:
            st.metric("✅ OTP Gross", f"{gross_otp:.1f}%")
        
        with col5:
            st.metric("🎯 OTP Net", f"{net_otp:.1f}%")
        
        st.markdown("---")
        
        # Row 1: Service and OTP
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Service Type Distribution")
            if 'SVC' in df.columns:
                svc_data = df['SVC'].value_counts().head(10)
                fig_svc = px.bar(
                    x=svc_data.values,
                    y=svc_data.index,
                    orientation='h',
                    color=svc_data.values,
                    color_continuous_scale='Viridis'
                )
                fig_svc.update_layout(
                    height=400,
                    xaxis_title="Count",
                    yaxis_title="Service",
                    showlegend=False
                )
                st.plotly_chart(fig_svc, use_container_width=True)
            else:
                st.info("No service data available")
        
        with col2:
            st.subheader("OTP Performance")
            otp_data = pd.DataFrame({
                'Metric': ['Gross OTP', 'Net OTP'],
                'Value': [gross_otp, net_otp]
            })
            fig_otp = px.bar(
                otp_data,
                x='Metric',
                y='Value',
                color='Value',
                color_continuous_scale='Blues',
                text='Value'
            )
            fig_otp.update_traces(texttemplate='%{text:.1f}%')
            fig_otp.update_layout(
                height=400,
                yaxis_title="Percentage (%)",
                showlegend=False
            )
            fig_otp.add_hline(y=90, line_dash="dash", line_color="red")
            st.plotly_chart(fig_otp, use_container_width=True)
        
        # Row 2: Departure and Cost Distribution
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Top Departure Airports")
            if 'DEP' in df.columns:
                dep_data = df['DEP'].value_counts().head(10)
                fig_dep = px.bar(
                    x=dep_data.index,
                    y=dep_data.values,
                    color=dep_data.values,
                    color_continuous_scale='Plasma'
                )
                fig_dep.update_layout(
                    height=400,
                    xaxis_title="Airport",
                    yaxis_title="Shipments",
                    showlegend=False
                )
                st.plotly_chart(fig_dep, use_container_width=True)
            else:
                st.info("No departure data available")
        
        with col2:
            st.subheader("Cost Distribution")
            cost_bins = pd.cut(
                df['TOTAL_CHARGES_EUR'],
                bins=[0, 500, 1000, 2000, 5000, float('inf')],
                labels=['<€500', '€500-1K', '€1K-2K', '€2K-5K', '>€5K']
            )
            cost_dist = cost_bins.value_counts()
            fig_cost = px.pie(
                values=cost_dist.values,
                names=cost_dist.index,
                hole=0.4
            )
            fig_cost.update_layout(height=400)
            st.plotly_chart(fig_cost, use_container_width=True)
        
        # Row 3: Monthly Trend and QC Issues
        st.markdown("---")
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Monthly Trend")
            if 'Month' in df.columns:
                monthly = df.groupby('Month').agg({
                    'REFER': 'count',
                    'TOTAL_CHARGES_EUR': 'sum'
                }).reset_index()
                monthly.columns = ['Month', 'Orders', 'Cost']
                
                fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
                fig_trend.add_trace(
                    go.Bar(x=monthly['Month'], y=monthly['Orders'], name='Orders'),
                    secondary_y=False
                )
                fig_trend.add_trace(
                    go.Scatter(x=monthly['Month'], y=monthly['Cost'], name='Cost (€)', mode='lines+markers'),
                    secondary_y=True
                )
                fig_trend.update_layout(height=400, hovermode='x unified')
                fig_trend.update_yaxes(title_text="Orders", secondary_y=False)
                fig_trend.update_yaxes(title_text="Cost (€)", secondary_y=True)
                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.info("No monthly data available")
        
        with col2:
            st.subheader("Quality Control Issues")
            if 'QCCODE' in df.columns and 'QC NAME' in df.columns:
                qc_data = df[df['QCCODE'].notna()]
                if len(qc_data) > 0:
                    qc_counts = qc_data.groupby('QC NAME').size().sort_values(ascending=False).head(10)
                    fig_qc = px.bar(
                        x=qc_counts.values,
                        y=qc_counts.index,
                        orientation='h',
                        color=qc_counts.values,
                        color_continuous_scale='Reds'
                    )
                    fig_qc.update_layout(
                        height=400,
                        xaxis_title="Count",
                        yaxis_title="",
                        showlegend=False
                    )
                    st.plotly_chart(fig_qc, use_container_width=True)
                else:
                    st.info("No QC issues found")
            else:
                st.info("No QC data available")
        
        # Summary
        st.markdown("---")
        st.subheader("📋 Executive Summary")
        
        col1, col2 = st.columns(2)
        
        with col1:
            top_dep = df['DEP'].value_counts().index[0] if 'DEP' in df.columns and len(df['DEP'].value_counts()) > 0 else 'N/A'
            st.markdown(f"""
            **Key Metrics:**
            - Total volume: **{len(df):,} shipments**
            - Total cost: **€{df['TOTAL_CHARGES_EUR'].sum():,.0f}**
            - Average cost: **€{df['TOTAL_CHARGES_EUR'].mean():,.2f}**
            - Main hub: **{top_dep}**
            """)
        
        with col2:
            st.markdown(f"""
            **Performance:**
            - Gross OTP: **{gross_otp:.1f}%**
            - Net OTP: **{net_otp:.1f}%**
            - Gap: **{net_otp - gross_otp:.1f}%**
            - Status: All 440-BILLED
            """)
        
        # Download button
        csv = df.to_csv(index=False)
        st.download_button(
            label="📥 Download Data (CSV)",
            data=csv,
            file_name=f"shipment_data_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )

else:
    st.info("👆 Please upload your Excel file to begin analysis")
    st.markdown("""
    ### Instructions:
    1. Upload your shipment Excel file
    2. Dashboard will filter for 440-BILLED status
    3. All costs converted to EUR automatically
    4. View insights and download processed data
    """)
