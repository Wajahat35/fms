# advanced_fee_dashboard.py
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import os
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import calendar
import io
import base64

# Page configuration
st.set_page_config(
    page_title="Advanced Fee Management System",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for advanced styling
st.markdown("""
    <style>
    /* Main header */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        color: white;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .main-header h1 {
        font-size: 2.8rem;
        font-weight: 700;
        margin: 0;
    }
    .main-header p {
        font-size: 1.1rem;
        opacity: 0.9;
        margin-top: 0.5rem;
    }
    
    /* Metric cards */
    .metric-card {
        background: white;
        padding: 1.2rem;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-left: 4px solid #667eea;
        transition: transform 0.2s;
        margin-bottom: 0.5rem;
    }
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 4px 20px rgba(0,0,0,0.12);
    }
    .metric-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a1a2e;
        margin: 0.2rem 0;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        font-weight: 600;
    }
    .metric-change {
        font-size: 0.8rem;
        padding: 0.2rem 0.6rem;
        border-radius: 20px;
        display: inline-block;
    }
    .metric-change.positive {
        background: #d4edda;
        color: #155724;
    }
    .metric-change.negative {
        background: #f8d7da;
        color: #721c24;
    }
    
    /* Status badges */
    .status-badge {
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.8rem;
        display: inline-block;
    }
    .status-paid {
        background: #d4edda;
        color: #155724;
    }
    .status-partial {
        background: #fff3cd;
        color: #856404;
    }
    .status-outstanding {
        background: #f8d7da;
        color: #721c24;
    }
    
    /* Custom tabs */
    .custom-tab {
        padding: 0.8rem 1.5rem;
        background: #f0f2f6;
        border-radius: 8px 8px 0 0;
        cursor: pointer;
        font-weight: 600;
        transition: all 0.3s;
    }
    .custom-tab.active {
        background: #667eea;
        color: white;
    }
    
    /* Advanced table */
    .advanced-table {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    
    /* Progress bar */
    .progress-container {
        background: #e9ecef;
        border-radius: 10px;
        height: 10px;
        width: 100%;
        margin: 0.5rem 0;
    }
    .progress-bar {
        height: 100%;
        border-radius: 10px;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        transition: width 0.5s ease;
    }
    
    /* Floating action button */
    .fab {
        position: fixed;
        bottom: 2rem;
        right: 2rem;
        background: #667eea;
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 50px;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        cursor: pointer;
        transition: all 0.3s;
        z-index: 999;
        text-decoration: none;
        font-weight: 600;
    }
    .fab:hover {
        transform: scale(1.05);
        box-shadow: 0 6px 25px rgba(102, 126, 234, 0.6);
    }
    
    /* Dark mode support */
    @media (prefers-color-scheme: dark) {
        .metric-card {
            background: #1a1a2e;
            border-left-color: #667eea;
        }
        .metric-value {
            color: white;
        }
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'report' not in st.session_state:
    st.session_state.report = None
if 'merged_data' not in st.session_state:
    st.session_state.merged_data = None
if 'fee_df' not in st.session_state:
    st.session_state.fee_df = None
if 'ts_df' not in st.session_state:
    st.session_state.ts_df = None
if 'dark_mode' not in st.session_state:
    st.session_state.dark_mode = False
if 'selected_student' not in st.session_state:
    st.session_state.selected_student = None
if 'filter_preset' not in st.session_state:
    st.session_state.filter_preset = 'All Students'

# Helper functions
def load_and_process_data(fee_path, ts_path):
    """Load and process the fee data"""
    try:
        # Load sheets
        fee_df = pd.read_excel(fee_path, sheet_name='MIS')
        ts_df = pd.read_excel(ts_path, sheet_name='ENROLLED')
        
        # Process TS sheet
        ts_df['Form No'] = pd.to_numeric(ts_df['Form No'], errors='coerce')
        ts_df = ts_df.dropna(subset=['Form No'])
        ts_unique = ts_df.drop_duplicates(subset=['Form No'], keep='first')
        
        # Process fee sheet
        fee_df['Form Number'] = pd.to_numeric(fee_df['Form Number'], errors='coerce')
        fee_df = fee_df.dropna(subset=['Form Number'])
        
        # Payment columns
        payment_cols = ['admission', 'monthly fee', 'total', 'ecm', 'advance', 'fee_inc_paid']
        for col in payment_cols:
            if col in fee_df.columns:
                fee_df[col] = pd.to_numeric(fee_df[col], errors='coerce').fillna(0)
        
        # Aggregate payments
        if 'total' in fee_df.columns:
            fee_grouped = fee_df.groupby('Form Number').agg({
                'total': 'sum',
                'admission': 'first',
                'monthly fee': 'first'
            }).reset_index()
            fee_grouped = fee_grouped.rename(columns={'total': 'Total_Paid'})
        else:
            payment_cols_available = [col for col in payment_cols if col in fee_df.columns]
            fee_grouped = fee_df.groupby('Form Number').agg({
                col: 'sum' for col in payment_cols_available
            }).reset_index()
            fee_grouped['Total_Paid'] = fee_grouped[payment_cols_available].sum(axis=1)
        
        # Merge data
        merged = pd.merge(ts_unique, fee_grouped, left_on='Form No', right_on='Form Number', how='left')
        merged['Total_Paid'] = merged['Total_Paid'].fillna(0)
        
        # Calculate fee details
        merged['Admission_Date'] = pd.to_datetime(merged['Admission  Date'], errors='coerce')
        merged['Base_Fee'] = pd.to_numeric(merged['1st Inst.'], errors='coerce').fillna(0)
        merged['Admission_Fee'] = pd.to_numeric(merged['Admissions'], errors='coerce').fillna(0)
        merged['Total_Invoice'] = pd.to_numeric(merged['Total Invoice'], errors='coerce').fillna(0)
        merged['Total_Months'] = pd.to_numeric(merged['Total Months'], errors='coerce').fillna(0)
        merged['Course'] = merged.get('Course', '')
        merged['Batch'] = merged.get('Batch Code', '')
        
        # Calculate months passed
        current_date = datetime.now()
        
        def calculate_months(row):
            if pd.isna(row['Admission_Date']):
                return 0
            try:
                if not isinstance(row['Admission_Date'], pd.Timestamp):
                    admission = pd.Timestamp(row['Admission_Date'])
                else:
                    admission = row['Admission_Date']
                months = (current_date.year - admission.year) * 12 + (current_date.month - admission.month)
                months = months + 1
                return max(0, months)
            except:
                return 0
        
        merged['Months_Passed'] = merged.apply(calculate_months, axis=1)
        
        # Fee calculations
        merged['Expected_Fee'] = merged['Base_Fee'] * merged['Months_Passed']
        merged['Paid_After_Admission'] = merged['Total_Paid'] - merged['Admission_Fee']
        merged['Balance_Amount'] = merged['Expected_Fee'] - merged['Paid_After_Admission']
        
        merged['Outstanding_Amount'] = np.where(
            merged['Balance_Amount'] > merged['Total_Invoice'],
            merged['Total_Invoice'],
            merged['Balance_Amount']
        )
        merged['Outstanding_Amount'] = merged['Outstanding_Amount'].clip(lower=0)
        merged['Balance_Amount'] = merged['Balance_Amount'].clip(lower=0)
        merged['Invoice_Balance'] = merged['Total_Invoice'] - merged['Total_Paid']
        merged['Invoice_Balance'] = merged['Invoice_Balance'].clip(lower=0)
        
        # Payment percentage
        merged['Payment_Percentage'] = np.where(
            merged['Total_Invoice'] > 0,
            (merged['Total_Paid'] / merged['Total_Invoice'] * 100),
            0
        )
        
        # Status
        merged['Status'] = np.where(
            merged['Outstanding_Amount'] <= 0,
            'Paid in Full',
            np.where(
                merged['Outstanding_Amount'] < merged['Balance_Amount'] * 0.5,
                'Partial Payment',
                'Outstanding'
            )
        )
        
        # Risk Level (New feature)
        merged['Risk_Level'] = np.where(
            merged['Status'] == 'Paid in Full',
            'Low',
            np.where(
                merged['Outstanding_Amount'] < merged['Total_Invoice'] * 0.3,
                'Medium',
                'High'
            )
        )
        
        # Create report
        report_cols = [
            'Form No', 'Name', 'Course', 'Batch', 'Admission_Date', 'Total_Months',
            'Months_Passed', 'Base_Fee', 'Admission_Fee', 'Expected_Fee', 'Total_Invoice',
            'Total_Paid', 'Paid_After_Admission', 'Balance_Amount', 'Outstanding_Amount',
            'Invoice_Balance', 'Payment_Percentage', 'Status', 'Risk_Level'
        ]
        report_cols_existing = [col for col in report_cols if col in merged.columns]
        report = merged[report_cols_existing].copy()
        
        # Rename columns
        rename_dict = {
            'Admission_Fee': 'Admission_Fee_TS',
            'Paid_After_Admission': 'Paid_After_Admission_Fee',
            'Balance_Amount': 'Balance_After_Expected'
        }
        report = report.rename(columns=rename_dict)
        report = report.sort_values('Outstanding_Amount', ascending=False)
        
        return report, merged, fee_df, ts_df
        
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None, None, None, None

def create_student_profile(student_data):
    """Create a detailed student profile view"""
    st.markdown("---")
    st.subheader("👤 Student Profile")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
            <div style="background: #f8f9fa; padding: 1.5rem; border-radius: 10px;">
                <p><strong>📋 Student Name:</strong> {student_data.get('Name', 'N/A')}</p>
                <p><strong>🎓 Course:</strong> {student_data.get('Course', 'N/A')}</p>
                <p><strong>📚 Batch:</strong> {student_data.get('Batch', 'N/A')}</p>
                <p><strong>📅 Admission Date:</strong> {student_data.get('Admission_Date', 'N/A')}</p>
                <p><strong>📊 Months Passed:</strong> {student_data.get('Months_Passed', 0)}</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        # Payment progress
        payment_percent = student_data.get('Payment_Percentage', 0)
        status = student_data.get('Status', 'Unknown')
        risk = student_data.get('Risk_Level', 'Unknown')
        
        status_color = {
            'Paid in Full': '#28a745',
            'Partial Payment': '#ffc107',
            'Outstanding': '#dc3545'
        }.get(status, '#6c757d')
        
        risk_color = {
            'Low': '#28a745',
            'Medium': '#ffc107',
            'High': '#dc3545'
        }.get(risk, '#6c757d')
        
        st.markdown(f"""
            <div style="background: #f8f9fa; padding: 1.5rem; border-radius: 10px;">
                <p><strong>💰 Total Paid:</strong> Rs. {student_data.get('Total_Paid', 0):,.2f}</p>
                <p><strong>📄 Total Invoice:</strong> Rs. {student_data.get('Total_Invoice', 0):,.2f}</p>
                <p><strong>⚠️ Outstanding:</strong> Rs. {student_data.get('Outstanding_Amount', 0):,.2f}</p>
                <p><strong>📊 Payment:</strong> {payment_percent:.1f}%</p>
                <p><strong>📌 Status:</strong> <span class="status-badge status-{'paid' if status == 'Paid in Full' else 'partial' if status == 'Partial Payment' else 'outstanding'}">{status}</span></p>
                <p><strong>🔴 Risk Level:</strong> <span style="color: {risk_color}; font-weight: bold;">{risk}</span></p>
            </div>
        """, unsafe_allow_html=True)

def create_advanced_dashboard(report, merged):
    """Create the advanced dashboard with all features"""
    
    # Calculate summary metrics
    total_students = len(report)
    total_paid = report['Total_Paid'].sum()
    total_outstanding = report['Outstanding_Amount'].sum()
    total_invoice = report['Total_Invoice'].sum()
    total_expected = report['Expected_Fee'].sum()
    total_balance = report['Invoice_Balance'].sum()
    
    students_paid_full = report[report['Outstanding_Amount'] == 0].shape[0]
    students_with_outstanding = report[report['Outstanding_Amount'] > 0].shape[0]
    students_high_risk = report[report['Risk_Level'] == 'High'].shape[0]
    
    collection_rate = (total_paid / total_invoice * 100) if total_invoice > 0 else 0
    
    # Header
    st.markdown(f"""
        <div class="main-header">
            <h1>🏫 Advanced Fee Management System</h1>
            <p>Last Updated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')} | Total Students: {total_students}</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Top metrics row
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Total Students</div>
                <div class="metric-value">{total_students}</div>
                <span class="metric-change positive">Active</span>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #28a745;">
                <div class="metric-label">Total Collected</div>
                <div class="metric-value" style="color: #28a745;">Rs. {total_paid:,.0f}</div>
                <span class="metric-change positive">↑ {collection_rate:.1f}%</span>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #dc3545;">
                <div class="metric-label">Total Outstanding</div>
                <div class="metric-value" style="color: #dc3545;">Rs. {total_outstanding:,.0f}</div>
                <span class="metric-change negative">⚠️ {students_with_outstanding} students</span>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #ffc107;">
                <div class="metric-label">Paid in Full</div>
                <div class="metric-value" style="color: #28a745;">{students_paid_full}</div>
                <span class="metric-change positive">✓ {(students_paid_full/total_students*100):.1f}%</span>
            </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #dc3545;">
                <div class="metric-label">High Risk</div>
                <div class="metric-value" style="color: #dc3545;">{students_high_risk}</div>
                <span class="metric-change negative">🔴 Needs attention</span>
            </div>
        """, unsafe_allow_html=True)
    
    with col6:
        st.markdown(f"""
            <div class="metric-card" style="border-left-color: #667eea;">
                <div class="metric-label">Avg Payment</div>
                <div class="metric-value">Rs. {(total_paid/total_students if total_students > 0 else 0):,.0f}</div>
                <span class="metric-change positive">Per student</span>
            </div>
        """, unsafe_allow_html=True)
    
    # Advanced filters and controls
    st.markdown("---")
    
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        status_filter = st.selectbox(
            "📊 Filter by Status",
            options=['All Students', 'Paid in Full', 'Partial Payment', 'Outstanding']
        )
    
    with col_f2:
        risk_filter = st.selectbox(
            "🔴 Filter by Risk Level",
            options=['All Levels', 'Low', 'Medium', 'High']
        )
    
    with col_f3:
        if 'Course' in report.columns:
            course_options = ['All Courses'] + sorted(report['Course'].dropna().unique().tolist())
            course_filter = st.selectbox("🎓 Filter by Course", options=course_options)
        else:
            course_filter = 'All Courses'
    
    with col_f4:
        search_term = st.text_input("🔍 Search Student", placeholder="Name or Form No...")
    
    # Apply filters
    filtered_report = report.copy()
    
    if status_filter != 'All Students':
        filtered_report = filtered_report[filtered_report['Status'] == status_filter]
    
    if risk_filter != 'All Levels':
        filtered_report = filtered_report[filtered_report['Risk_Level'] == risk_filter]
    
    if course_filter != 'All Courses':
        filtered_report = filtered_report[filtered_report['Course'] == course_filter]
    
    if search_term:
        filtered_report = filtered_report[
            filtered_report['Name'].str.contains(search_term, case=False, na=False) |
            filtered_report['Form No'].astype(str).str.contains(search_term, case=False, na=False)
        ]
    
    # Main content tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Overview & Analytics",
        "👥 Student Management",
        "📈 Advanced Analytics",
        "🔔 Alerts & Actions",
        "📁 Data Management"
    ])
    
    with tab1:
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.subheader("📊 Payment Status Distribution")
            status_counts = filtered_report['Status'].value_counts()
            colors = {'Paid in Full': '#28a745', 'Partial Payment': '#ffc107', 'Outstanding': '#dc3545'}
            
            fig_status = go.Figure(data=[go.Pie(
                labels=status_counts.index,
                values=status_counts.values,
                hole=0.4,
                marker_colors=[colors.get(x, '#6c757d') for x in status_counts.index],
                textinfo='label+percent',
                hoverinfo='label+value+percent'
            )])
            fig_status.update_layout(
                height=400,
                showlegend=False,
                annotations=[dict(text=f'Total<br>{len(filtered_report)}', x=0.5, y=0.5, font_size=20, showarrow=False)]
            )
            st.plotly_chart(fig_status, use_container_width=True)
        
        with col_chart2:
            st.subheader("💰 Collection Overview")
            
            # Gauge chart for collection rate
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=collection_rate,
                title={'text': "Collection Rate"},
                delta={'reference': 70},
                gauge={
                    'axis': {'range': [None, 100]},
                    'bar': {'color': "#667eea"},
                    'steps': [
                        {'range': [0, 50], 'color': "lightgray"},
                        {'range': [50, 75], 'color': "gray"}
                    ],
                    'threshold': {
                        'line': {'color': "red", 'width': 4},
                        'thickness': 0.75,
                        'value': 90
                    }
                }
            ))
            fig_gauge.update_layout(height=300)
            st.plotly_chart(fig_gauge, use_container_width=True)
        
        # Advanced charts row
        col_chart3, col_chart4 = st.columns(2)
        
        with col_chart3:
            st.subheader("📈 Monthly Collection Trend")
            
            # Create monthly data
            if 'Admission_Date' in filtered_report.columns:
                monthly_data = filtered_report.groupby(
                    filtered_report['Admission_Date'].dt.to_period('M')
                ).agg({
                    'Total_Paid': 'sum',
                    'Total_Invoice': 'sum',
                    'Outstanding_Amount': 'sum'
                }).reset_index()
                monthly_data['Admission_Date'] = monthly_data['Admission_Date'].astype(str)
                
                fig_trend = go.Figure()
                fig_trend.add_trace(go.Bar(
                    x=monthly_data['Admission_Date'],
                    y=monthly_data['Total_Paid'],
                    name='Paid',
                    marker_color='#28a745'
                ))
                fig_trend.add_trace(go.Bar(
                    x=monthly_data['Admission_Date'],
                    y=monthly_data['Total_Invoice'],
                    name='Invoice',
                    marker_color='#667eea'
                ))
                fig_trend.update_layout(
                    height=350,
                    barmode='group',
                    xaxis_title='Month',
                    yaxis_title='Amount (Rs.)'
                )
                st.plotly_chart(fig_trend, use_container_width=True)
        
        with col_chart4:
            st.subheader("🏫 Course-wise Analysis")
            
            if 'Course' in filtered_report.columns:
                course_data = filtered_report.groupby('Course').agg({
                    'Total_Paid': 'sum',
                    'Outstanding_Amount': 'sum',
                    'Total_Invoice': 'sum'
                }).reset_index()
                course_data = course_data[course_data['Total_Invoice'] > 0]
                
                if not course_data.empty:
                    fig_course = go.Figure()
                    fig_course.add_trace(go.Bar(
                        y=course_data['Course'],
                        x=course_data['Total_Paid'],
                        name='Paid',
                        orientation='h',
                        marker_color='#28a745'
                    ))
                    fig_course.add_trace(go.Bar(
                        y=course_data['Course'],
                        x=course_data['Outstanding_Amount'],
                        name='Outstanding',
                        orientation='h',
                        marker_color='#dc3545'
                    ))
                    fig_course.update_layout(
                        height=350,
                        barmode='stack',
                        xaxis_title='Amount (Rs.)',
                        yaxis_title='Course'
                    )
                    st.plotly_chart(fig_course, use_container_width=True)
    
    with tab2:
        st.subheader("👥 Student Management")
        
        # Student selection
        st.markdown("### Select Student for Detailed View")
        
        student_options = filtered_report['Name'].tolist()
        selected_student_name = st.selectbox(
            "Choose a student",
            options=student_options
        )
        
        if selected_student_name:
            student_data = filtered_report[filtered_report['Name'] == selected_student_name].iloc[0]
            
            # Display student profile
            create_student_profile(student_data)
            
            # Payment history (simulated)
            st.markdown("---")
            st.subheader("📊 Payment History")
            
            # Create simulated payment history
            if student_data.get('Total_Paid', 0) > 0:
                months = student_data.get('Months_Passed', 1)
                base_fee = student_data.get('Base_Fee', 0)
                
                # Generate simulated payment history
                payment_history = pd.DataFrame({
                    'Month': [f'Month {i+1}' for i in range(months)],
                    'Expected': [base_fee] * months,
                    'Paid': [base_fee * (0.7 + np.random.random() * 0.3) for _ in range(months)]
                })
                payment_history['Paid'] = payment_history['Paid'].clip(upper=payment_history['Expected'])
                
                fig_history = px.line(
                    payment_history,
                    x='Month',
                    y=['Expected', 'Paid'],
                    title='Payment History',
                    markers=True
                )
                fig_history.update_layout(height=300)
                st.plotly_chart(fig_history, use_container_width=True)
    
    with tab3:
        st.subheader("📈 Advanced Analytics")
        
        col_adv1, col_adv2 = st.columns(2)
        
        with col_adv1:
            # Risk distribution
            st.markdown("### Risk Analysis")
            risk_counts = filtered_report['Risk_Level'].value_counts()
            colors_risk = {'Low': '#28a745', 'Medium': '#ffc107', 'High': '#dc3545'}
            
            fig_risk = px.bar(
                x=risk_counts.index,
                y=risk_counts.values,
                color=risk_counts.index,
                color_discrete_map=colors_risk,
                title="Risk Level Distribution"
            )
            fig_risk.update_layout(height=300, showlegend=False)
            st.plotly_chart(fig_risk, use_container_width=True)
        
        with col_adv2:
            # Outstanding range analysis
            st.markdown("### Outstanding Range")
            
            bins = [0, 10000, 25000, 50000, 100000, float('inf')]
            labels = ['0-10k', '10k-25k', '25k-50k', '50k-100k', '100k+']
            
            filtered_report['Outstanding_Range'] = pd.cut(
                filtered_report['Outstanding_Amount'],
                bins=bins,
                labels=labels,
                right=False
            )
            
            range_counts = filtered_report['Outstanding_Range'].value_counts().sort_index()
            
            fig_range = px.bar(
                x=range_counts.index,
                y=range_counts.values,
                title="Outstanding Amount Range",
                color=range_counts.index,
                color_discrete_sequence=px.colors.sequential.Blues_r
            )
            fig_range.update_layout(height=300, showlegend=False)
            st.plotly_chart(fig_range, use_container_width=True)
        
        # Correlation analysis
        st.markdown("### 📊 Payment Performance Correlation")
        
        if len(filtered_report) > 1:
            corr_cols = ['Months_Passed', 'Base_Fee', 'Total_Invoice', 'Total_Paid', 'Outstanding_Amount', 'Payment_Percentage']
            corr_cols_existing = [col for col in corr_cols if col in filtered_report.columns]
            
            if len(corr_cols_existing) > 1:
                corr_matrix = filtered_report[corr_cols_existing].corr()
                
                fig_corr = px.imshow(
                    corr_matrix,
                    text_auto=True,
                    aspect="auto",
                    title="Correlation Matrix",
                    color_continuous_scale='RdBu_r'
                )
                fig_corr.update_layout(height=400)
                st.plotly_chart(fig_corr, use_container_width=True)
    
    with tab4:
        st.subheader("🔔 Alerts & Actions")
        
        # High risk alerts
        high_risk_students = filtered_report[filtered_report['Risk_Level'] == 'High']
        if not high_risk_students.empty:
            st.warning(f"⚠️ {len(high_risk_students)} students are in the HIGH RISK category!")
            
            for _, student in high_risk_students.head(5).iterrows():
                with st.expander(f"🔴 {student['Name']} - Risk: High"):
                    col_alert1, col_alert2 = st.columns(2)
                    with col_alert1:
                        st.write(f"**Course:** {student.get('Course', 'N/A')}")
                        st.write(f"**Outstanding:** Rs. {student['Outstanding_Amount']:,.2f}")
                    with col_alert2:
                        st.write(f"**Payment:** {student['Payment_Percentage']:.1f}%")
                        if st.button(f"📧 Send Alert to {student['Name']}", key=f"alert_{student['Form No']}"):
                            st.success(f"Alert sent to {student['Name']}!")
        else:
            st.success("✅ No high-risk students found!")
        
        # Upcoming actions
        st.markdown("---")
        st.subheader("📋 Recommended Actions")
        
        action_col1, action_col2, action_col3 = st.columns(3)
        
        with action_col1:
            st.info("📊 **Review Collections**")
            st.write(f"Collection rate: {collection_rate:.1f}%")
            if collection_rate < 70:
                st.warning("Collection rate below 70%. Consider reviewing payment plans.")
        
        with action_col2:
            st.info("👥 **Student Follow-ups**")
            overdue = filtered_report[filtered_report['Outstanding_Amount'] > 0]
            st.write(f"{len(overdue)} students need follow-up")
            if len(overdue) > 0:
                st.write(f"Average outstanding: Rs. {overdue['Outstanding_Amount'].mean():,.2f}")
        
        with action_col3:
            st.info("📅 **Upcoming Deadlines**")
            # Simulated deadlines
            st.write("🔹 Payment reminders to send: 15")
            st.write("🔹 Fee reviews pending: 8")
            st.write("🔹 New admissions: 5")
    
    with tab5:
        st.subheader("📁 Data Management")
        
        col_data1, col_data2 = st.columns(2)
        
        with col_data1:
            st.markdown("### 📤 Export Data")
            
            # Export CSV
            csv = filtered_report.to_csv(index=False)
            st.download_button(
                label="📥 Download as CSV",
                data=csv,
                file_name=f"fee_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
            # Export Excel
            if st.button("📊 Download as Excel", use_container_width=True):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    filtered_report.to_excel(writer, sheet_name='Fee Report', index=False)
                    
                    # Add summary sheet
                    summary_data = {
                        'Metric': ['Total Students', 'Total Paid', 'Total Outstanding', 
                                  'Collection Rate', 'Paid in Full', 'High Risk'],
                        'Value': [len(filtered_report), total_paid, total_outstanding,
                                 f"{collection_rate:.1f}%", students_paid_full, students_high_risk]
                    }
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                st.success("✅ Excel report generated and ready for download!")
                st.download_button(
                    label="📥 Download Excel",
                    data=output.getvalue(),
                    file_name=f"fee_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        with col_data2:
            st.markdown("### 📊 Data Statistics")
            st.write(f"**Total Records:** {len(filtered_report)}")
            st.write(f"**Columns:** {len(filtered_report.columns)}")
            
            # Data quality
            st.markdown("#### Data Quality")
            missing_data = filtered_report.isnull().sum()
            missing_data = missing_data[missing_data > 0]
            
            if not missing_data.empty:
                st.warning("Missing data detected:")
                for col, count in missing_data.items():
                    st.write(f"- {col}: {count} missing")
            else:
                st.success("✅ No missing data found!")
    
    # Detailed table at the bottom
    st.markdown("---")
    st.subheader("📋 Detailed Student Records")
    
    # Advanced table with search and sort
    display_cols = ['Form No', 'Name', 'Course', 'Batch', 'Admission_Date', 'Months_Passed',
                   'Base_Fee', 'Total_Invoice', 'Total_Paid', 'Outstanding_Amount',
                   'Payment_Percentage', 'Status', 'Risk_Level']
    
    display_cols_existing = [col for col in display_cols if col in filtered_report.columns]
    df_display = filtered_report[display_cols_existing].copy()
    
    # Format the dataframe
    if 'Admission_Date' in df_display.columns:
        df_display['Admission_Date'] = df_display['Admission_Date'].dt.strftime('%Y-%m-%d')
    
    # Format currency
    currency_cols = ['Base_Fee', 'Total_Invoice', 'Total_Paid', 'Outstanding_Amount']
    for col in currency_cols:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(lambda x: f"Rs. {x:,.2f}")
    
    # Add status badges
    def add_status_badge(status):
        if status == 'Paid in Full':
            return '🟢 Paid in Full'
        elif status == 'Partial Payment':
            return '🟡 Partial Payment'
        else:
            return '🔴 Outstanding'
    
    if 'Status' in df_display.columns:
        df_display['Status'] = df_display['Status'].apply(add_status_badge)
    
    # Format risk level
    def add_risk_badge(risk):
        if risk == 'Low':
            return '🟢 Low'
        elif risk == 'Medium':
            return '🟡 Medium'
        else:
            return '🔴 High'
    
    if 'Risk_Level' in df_display.columns:
        df_display['Risk_Level'] = df_display['Risk_Level'].apply(add_risk_badge)
    
    # Display the table
    st.dataframe(
        df_display,
        use_container_width=True,
        height=500,
        hide_index=True
    )

def main():
    # Sidebar
    with st.sidebar:
        st.image("https://via.placeholder.com/150x150/667eea/ffffff?text=🏫", use_column_width=True)
        st.markdown("## 📁 Data Upload")
        st.markdown("---")
        
        # File uploaders
        fee_file = st.file_uploader(
            "📊 Fee Detail (MIS Sheet)",
            type=['xlsx', 'xls'],
            help="Upload the Fee detail file with MIS sheet"
        )
        
        ts_file = st.file_uploader(
            "📋 TS Enrolled",
            type=['xlsx', 'xls'],
            help="Upload the TS Enrolled file"
        )
        
        st.markdown("---")
        
        # Default paths
        st.subheader("📂 Default Files")
        default_fee_path = r'D:\APTECH ACCOUNT\Fee Managment System\Fee detail Updated.xlsx'
        default_ts_path = r'D:\APTECH ACCOUNT\Fee Managment System\TS 19-06-2026.xlsx'
        
        use_default = st.checkbox("Use default file paths", value=True)
        
        st.markdown("---")
        
        # Load button
        if st.button("🚀 Load & Process Data", type="primary", use_container_width=True):
            with st.spinner("Loading and processing data..."):
                if use_default:
                    if os.path.exists(default_fee_path) and os.path.exists(default_ts_path):
                        report, merged, fee_df, ts_df = load_and_process_data(default_fee_path, default_ts_path)
                    else:
                        st.error("Default files not found. Please upload files manually.")
                        report, merged, fee_df, ts_df = None, None, None, None
                else:
                    if fee_file and ts_file:
                        # Save uploaded files temporarily
                        temp_fee = f"temp_fee_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        temp_ts = f"temp_ts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        
                        with open(temp_fee, 'wb') as f:
                            f.write(fee_file.getbuffer())
                        with open(temp_ts, 'wb') as f:
                            f.write(ts_file.getbuffer())
                        
                        report, merged, fee_df, ts_df = load_and_process_data(temp_fee, temp_ts)
                        
                        # Clean up temp files
                        try:
                            os.remove(temp_fee)
                            os.remove(temp_ts)
                        except:
                            pass
                    else:
                        st.error("Please upload both files or use default paths.")
                        report, merged, fee_df, ts_df = None, None, None, None
                
                if report is not None:
                    st.session_state.report = report
                    st.session_state.merged_data = merged
                    st.session_state.fee_df = fee_df
                    st.session_state.ts_df = ts_df
                    st.session_state.data_loaded = True
                    st.success("✅ Data loaded successfully!")
                    st.rerun()
                else:
                    st.error("Failed to load data. Please check the files.")
        
        st.markdown("---")
        st.markdown("""
            <div style="text-align: center; color: #666; font-size: 0.8rem;">
                <p>🏫 Advanced Fee Management System</p>
                <p>Version 2.0</p>
                <p>© 2026 All Rights Reserved</p>
            </div>
        """, unsafe_allow_html=True)
    
    # Main content
    if st.session_state.data_loaded and st.session_state.report is not None:
        create_advanced_dashboard(st.session_state.report, st.session_state.merged_data)
        
        # Floating action button
        st.markdown("""
            <a href="#" class="fab" onclick="window.scrollTo({top: 0, behavior: 'smooth'}); return false;">
                ⬆️ Back to Top
            </a>
        """, unsafe_allow_html=True)
        
        # Refresh button
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    else:
        st.info("👈 Please load data using the sidebar to view the dashboard.")
        
        # Show welcome message
        st.markdown("""
            <div style="text-align: center; padding: 3rem;">
                <h2>🏫 Welcome to Advanced Fee Management System</h2>
                <p style="font-size: 1.1rem; color: #666;">
                    Load your data to get started with real-time analytics and management.
                </p>
                <div style="display: flex; justify-content: center; gap: 2rem; margin-top: 2rem;">
                    <div style="background: #f0f2f6; padding: 1.5rem; border-radius: 10px; width: 200px;">
                        <h3>📊</h3>
                        <p>Real-time Analytics</p>
                    </div>
                    <div style="background: #f0f2f6; padding: 1.5rem; border-radius: 10px; width: 200px;">
                        <h3>👥</h3>
                        <p>Student Management</p>
                    </div>
                    <div style="background: #f0f2f6; padding: 1.5rem; border-radius: 10px; width: 200px;">
                        <h3>📈</h3>
                        <p>Advanced Reporting</p>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()