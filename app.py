# --- diagnostics bootstrap: install crash/exception logging before app code ---
import os as _diag_os
import sys as _diag_sys

_diag_dir = _diag_os.path.dirname(_diag_os.path.abspath(__file__))
if _diag_dir not in _diag_sys.path:
    _diag_sys.path.insert(0, _diag_dir)
import app_diagnostics  # noqa: E402

app_diagnostics.install_diagnostics()
# --- end diagnostics bootstrap ---

import streamlit as st
import pypdf
from docx import Document
import io
import json
import os
from databricks.sdk import WorkspaceClient
from databricks.sdk.service.serving import ChatMessage, ChatMessageRole

# Page configuration
st.set_page_config(
    page_title="BizForge - AI Solution Engineering Platform",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium Enterprise SaaS Design - BizForge
st.markdown("""
<style>
    /* Import Premium Font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    /* Global Premium Light Theme */
    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
    }
    
    /* Main Container - Clean White Background */
    .main {
        background: #ffffff;
        padding: 3rem 2rem;
        color: #000000;
    }
    
    /* Text Color Override - Maximum Contrast */
    .stMarkdown, p, span, div {
        color: #000000;
        line-height: 1.6;
    }
    
    /* Force All Headings to Black */
    h1, h2, h3, h4, h5, h6,
    [data-testid="stHeader"],
    [data-testid="stSubheader"],
    [data-testid="stTitle"],
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3,
    [data-testid="stMarkdownContainer"] h4,
    [data-testid="stMarkdownContainer"] h5,
    [data-testid="stMarkdownContainer"] h6,
    [data-testid="stHeading"] {
        color: #000000 !important;
    }
    
    /* Premium Header - Enhanced Readability */
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #000000;
        margin-bottom: 0.5rem;
        letter-spacing: -0.5px;
        text-align: center;
        padding: 1.25rem 0 1rem 0;
        line-height: 1.2;
    }
    
    /* Subtitle */
    .subtitle {
        font-size: 1rem;
        color: #000000;
        margin-bottom: 3rem;
        text-align: center;
        font-weight: 400;
        letter-spacing: 0.5px;
        letter-spacing: 0.5px;
    }
    
    /* Phase Badge */
    .phase-badge {
        background: #ffffff;
        color: #000000;
        padding: 0.5rem 1.2rem;
        border-radius: 6px;
        font-size: 0.9rem;
        font-weight: 500;
        display: inline-block;
        margin: 1rem auto;
        text-align: center;
        border: 1px solid #e2e8f0;
    }
    
    /* Professional Cards - Clean White Background */
    .stExpander {
        background: #ffffff !important;
        border: 2px solid #e2e8f0 !important;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1), 0 1px 2px rgba(0, 0, 0, 0.06);
        margin: 1rem 0;
        transition: all 0.2s ease;
    }
    
    /* Force white background on expander content */
    .stExpander > div {
        background: #ffffff !important;
        color: #000000 !important;
    }
    
    /* Force white background on all expander child elements */
    .stExpander * {
        background-color: transparent !important;
    }
    
    .stExpander div[data-testid="stMarkdownContainer"] {
        background: #ffffff !important;
        color: #000000 !important;
    }
    
    .stExpander:hover {
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15), 0 4px 10px rgba(0, 0, 0, 0.1);
        transform: translateY(-2px);
        border-color: #cbd5e1 !important;
    }
    
    /* Success Box */
    .success-box {
        padding: 1rem;
        background: #ffffff;
        border-left: 3px solid #16a34a;
        border-radius: 6px;
        margin: 1rem 0;
        font-weight: 400;
    }
    
    /* Info Box */
    .info-box {
        padding: 1rem;
        background: #ffffff;
        border-left: 3px solid #2563eb;
        border-radius: 6px;
        margin: 1rem 0;
        font-weight: 400;
    }
    
    /* Warning Box */
    .warning-box {
        padding: 1rem;
        background: #ffffff;
        border-left: 4px solid #f59e0b;
        border-radius: 6px;
        margin: 1rem 0;
        font-weight: 400;
    }
    
    /* Buttons - High Contrast */
    .stButton > button {
        background: #2563eb;
        color: #000000;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 1.75rem;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.2s ease;
        cursor: pointer;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
    }
    
    .stButton > button:hover {
        background: #1d4ed8;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
        transform: translateY(-1px);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* Download Buttons - High Contrast */
    .stDownloadButton > button {
        background: #16a34a;
        color: #000000;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 1.75rem;
        font-weight: 600;
        transition: all 0.2s ease;
        box-shadow: 0 2px 4px rgba(22, 163, 74, 0.2);
    }
    
    .stDownloadButton > button:hover {
        background: #15803d;
        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.3);
        transform: translateY(-1px);
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 1rem;
        background: #ffffff;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border: none;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        color: #000000;
        border-radius: 8px;
        transition: all 0.15s ease;
    }
    
    .stTabs [aria-selected="true"] {
        background: #eff6ff;
        color: #000000 !important;
        border-bottom: 2px solid #2563eb;
        box-shadow: none;
    }
    
    /* Progress Bar */
    .stProgress > div > div > div > div {
        background: #2563eb;
        border-radius: 8px;
    }
    
    /* Metrics - Enhanced Visibility */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 700;
        color: #000000;
    }
    
    /* File Uploader */
    [data-testid="stFileUploader"] {
        background: #ffffff;
        border: 2px dashed #e2e8f0;
        border-radius: 8px;
        padding: 1.25rem;
        transition: all 0.15s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #2563eb;
        background: #eff6ff;
    }
    
    /* Text Area */
    .stTextArea textarea {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
        font-family: 'Monaco', 'Courier New', monospace;
        color: #000000 !important;
        transition: all 0.15s ease;
    }
    
    .stTextArea textarea:focus {
        border-color: #2563eb;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.2);
    }
    
    /* Code Blocks - Clean Design */
    .stCodeBlock {
        background: #f8fafc !important;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
    }
    
    /* Ensure code and pre elements have light backgrounds */
    code, pre {
        background: #f8fafc !important;
        color: #000000 !important;
    }
    
    .stExpander code, .stExpander pre {
        background: #f8fafc !important;
        color: #000000 !important;
    }
    
    /* Sidebar - Clean White Theme */
    [data-testid="stSidebar"] {
        background: #f8fafc;
        padding: 1.5rem 1rem;
        border-right: 2px solid #e2e8f0;
    }
    
    [data-testid="stSidebar"] .stMarkdown {
        color: #000000 !important;
    }
    
    /* Divider */
    hr {
        margin: 2rem 0;
        border: none;
        height: 2px;
        background: #e2e8f0;
    }
    
    /* Animations */
    @keyframes fadeInDown {
        from {
            opacity: 0;
            transform: translateY(-20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    @keyframes slideInLeft {
        from {
            opacity: 0;
            transform: translateX(-20px);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }
    
    @keyframes pulse {
        0%, 100% {
            box-shadow: 0 1px 3px rgba(37, 99, 235, 0.4);
        }
        50% {
            box-shadow: 0 4px 25px rgba(37, 99, 235, 0.7);
        }
    }
    
    /* Tooltips */
    .tooltip {
        position: relative;
        display: inline-block;
    }
    
    /* Custom Scrollbar */
    ::-webkit-scrollbar {
        width: 10px;
        height: 10px;
    }
    
    ::-webkit-scrollbar-track {
        background: #f1f5f9;
        border-radius: 8px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: #2563eb;
        border-radius: 8px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: #2563eb;
    }
    
    /* Select Box */
    .stSelectbox > div > div {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        transition: all 0.15s ease;
    }
    
    .stSelectbox > div > div:hover {
        border-color: #2563eb;
    }
    
    /* Info/Success Messages from Streamlit */
    .stAlert {
        border-radius: 8px;
        border: none;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
    }
    
    /* Professional Footer */
    .footer {
        text-align: center;
        padding: 1.25rem;
        background: #ffffff;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
        margin-top: 3rem;
        font-weight: 500;
        color: #000000;
    }
    
    /* Phase Progress Cards */
    .phase-card {
        background: #ffffff;
        border-radius: 8px;
        padding: 1.25rem;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
        transition: all 0.15s ease;
        text-align: center;
        border-top: 4px solid #2563eb;
    }
    
    .phase-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
    }
    
    /* Status Badges */
    .status-badge {
        display: inline-block;
        padding: 0.4rem 1rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .status-complete {
        background: #ffffff;
        color: #000000;
    }
    
    .status-pending {
        background: #f59e0b;
        color: #000000;
    }
    
    /* Loading Animation */
    .loading-spinner {
        border: 4px solid #e2e8f0;
        border-top: 4px solid #2563eb;
        border-radius: 50%;
        width: 40px;
        height: 40px;
        animation: spin 1s linear infinite;
        margin: 2rem auto;
    }
    
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
</style>
""", unsafe_allow_html=True)

# Initialize ALL session state variables at startup to avoid AttributeError
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
if 'analysis_results' not in st.session_state:
    st.session_state.analysis_results = None
if 'selected_accelerator' not in st.session_state:
    st.session_state.selected_accelerator = None
if 'show_recommendations' not in st.session_state:
    st.session_state.show_recommendations = False
if 'architecture_generated' not in st.session_state:
    st.session_state.architecture_generated = False
if 'architecture_design' not in st.session_state:
    st.session_state.architecture_design = None
if 'artifacts_generated' not in st.session_state:
    st.session_state.artifacts_generated = False
if 'generated_artifacts' not in st.session_state:
    st.session_state.generated_artifacts = None
if 'data_model_generated' not in st.session_state:
    st.session_state.data_model_generated = False
if 'data_model_design' not in st.session_state:
    st.session_state.data_model_design = None
if 'pipeline_design_generated' not in st.session_state:
    st.session_state.pipeline_design_generated = False
if 'pipeline_design' not in st.session_state:
    st.session_state.pipeline_design = None
if 'analytics_design_generated' not in st.session_state:
    st.session_state.analytics_design_generated = False
if 'analytics_design' not in st.session_state:
    st.session_state.analytics_design = None
if 'testing_framework_generated' not in st.session_state:
    st.session_state.testing_framework_generated = False
if 'testing_framework' not in st.session_state:
    st.session_state.testing_framework = None
if 'deployment_plan_generated' not in st.session_state:
    st.session_state.deployment_plan_generated = False
if 'deployment_plan' not in st.session_state:
    st.session_state.deployment_plan = None

# Initialize selected phase for navigation
if 'selected_phase' not in st.session_state:
    st.session_state.selected_phase = None

# Premium Enterprise Sidebar - Clean Vertical Workflow
with st.sidebar:
    # Minimal Logo Header
    st.markdown("""
    <div style="text-align: center; padding: 2rem 0 1.5rem 0; border-bottom: 1px solid #e2e8f0; margin-bottom: 2rem;">
        <div style="font-size: 1.5rem; font-weight: 700; margin-bottom: 0.3rem; letter-spacing: -0.5px;">
            <span style="color: #000000;">BizForge</span>
        </div>
        <p style="color: #000000; font-size: 0.7rem; margin: 0; letter-spacing: 1px; text-transform: uppercase; font-weight: 500;">Solution Engineering</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Progress Tracker
    completed_phases = 0
    if st.session_state.analysis_complete: completed_phases += 1
    if st.session_state.selected_accelerator: completed_phases += 1
    if st.session_state.architecture_generated: completed_phases += 1
    if st.session_state.artifacts_generated: completed_phases += 1
    if st.session_state.data_model_generated: completed_phases += 1
    if st.session_state.pipeline_design_generated: completed_phases += 1
    if st.session_state.analytics_design_generated: completed_phases += 1
    if st.session_state.testing_framework_generated: completed_phases += 1
    if st.session_state.deployment_plan_generated: completed_phases += 1
    
    progress_pct = (completed_phases / 10.0) * 100
    
    st.markdown(f"""
    <div style="background: #ffffff; padding: 1rem; border-radius: 8px; margin-bottom: 1.5rem; border-left: 3px solid #16a34a;">
        <p style="color: #000000; font-weight: 600; margin-bottom: 0.5rem; font-size: 0.85rem;">Overall Progress</p>
        <div style="background: #f8fafc; height: 6px; border-radius: 8px; overflow: hidden;">
            <div style="background: #16a34a; height: 100%; width: {progress_pct}%; border-radius: 8px; transition: width 0.3s ease;"></div>
        </div>
        <p style="color: #000000; margin-top: 0.5rem; font-size: 0.8rem;">{completed_phases}/10 Complete</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Clean Vertical Workflow Navigation
    st.markdown('<p style="color: #000000; font-weight: 600; margin-bottom: 1.5rem; font-size: 0.7rem; letter-spacing: 1px; text-transform: uppercase;">Workflow</p>', unsafe_allow_html=True)
    
    # Define phases with anchors
    nav_phases = [
        ("1. BRD Upload", "phase-upload", st.session_state.get('analysis_results') is not None),
        ("2. AI Analysis", "phase-analysis", st.session_state.analysis_complete),
        ("3. Platform Selection", "phase-platform", st.session_state.selected_accelerator is not None),
        ("4. Architecture", "phase-architecture", st.session_state.architecture_generated),
        ("5. Artifacts", "phase-artifacts", st.session_state.artifacts_generated),
        ("6. Data Model", "phase-datamodel", st.session_state.data_model_generated),
        ("7. Pipeline/ETL", "phase-pipeline", st.session_state.pipeline_design_generated),
        ("8. Analytics", "phase-analytics", st.session_state.analytics_design_generated),
        ("9. Testing", "phase-testing", st.session_state.testing_framework_generated),
        ("10. Deployment", "phase-deployment", st.session_state.deployment_plan_generated)
    ]
    
    for phase_name, anchor, is_complete in nav_phases:
        accent_color = "#16a34a" if is_complete else "#e2e8f0"
        text_color = "#000000" if is_complete else "#000000"
        bg_color = "#ffffff" if is_complete else "#f8fafc"
        hover_bg = "#f1f5f9" if not is_complete else "#f0fdf4"
        
        # Create minimal clickable navigation item
        st.markdown(f"""
        <a href="#{anchor}" style="text-decoration: none;">
            <div style="
                background: {bg_color};
                padding: 0.75rem 1rem;
                margin-bottom: 0.4rem;
                border-radius: 6px;
                border-left: 2px solid {accent_color};
                cursor: pointer;
                transition: all 0.15s ease;
                color: {text_color};
                font-size: 0.8rem;
                font-weight: 500;
                border: 1px solid #e2e8f0;
            " onmouseover="this.style.background='{hover_bg}'" onmouseout="this.style.background='{bg_color}'">
                {phase_name}
            </div>
        </a>
        """, unsafe_allow_html=True)
    
    st.markdown("<div style='height: 2rem;'></div>", unsafe_allow_html=True)
    
    # Platform Info
    if st.session_state.selected_accelerator:
        st.markdown(f"""
        <div style="background: #ffffff; padding: 1rem; border-radius: 8px; border-left: 3px solid #2563eb;">
            <p style="color: #000000; font-weight: 600; font-size: 0.85rem; margin-bottom: 0.5rem;">Selected Platform</p>
            <p style="color: #000000; font-size: 0.85rem; font-weight: 600;">{st.session_state.selected_accelerator}</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<div style='margin-top: auto; padding-top: 2rem;'><div style='text-align: center; padding: 0.75rem; background: #ffffff; border-radius: 6px;'><p style='color: #000000; font-size: 0.7rem; margin: 0;'>Powered by Databricks AI</p></div></div>", unsafe_allow_html=True)

# Premium Minimal Header
st.markdown('''
<div style="text-align: center; padding: 3rem 0 2rem 0;">
    <div style="font-size: 2.5rem; font-weight: 700; color: #000000; margin-bottom: 0.5rem; letter-spacing: -1px;">
        BizForge
    </div>
    <p style="color: #000000; font-size: 0.95rem; margin: 0; font-weight: 400;">AI-Powered Solution Engineering Platform</p>
    <div style="margin-top: 1.5rem;">
        <span style="background: #f8fafc; color: #000000; padding: 0.5rem 1.25rem; border-radius: 20px; font-size: 0.8rem; font-weight: 500; border: 1px solid #e2e8f0;">BRD → AI Analysis → Architecture → Data → Pipeline → Deployment</span>
    </div>
</div>
''', unsafe_allow_html=True)

st.markdown("<div style='border-top: 1px solid #e2e8f0; margin: 2rem 0;'></div>", unsafe_allow_html=True)

# Professional Introduction Section
with st.expander("About BizForge - Enterprise AI Solution Engineering", expanded=False):
    st.markdown("""
    <div style="padding: 1.5rem; background: #ffffff;">
        <h3 style="color: #000000; margin-bottom: 1rem; font-weight: 600;"> What is BizForge?</h3>
        <p style="font-size: 1.05rem; line-height: 1.8; color: #000000;">
            <strong>BizForge</strong> is an enterprise-grade AI platform that transforms business requirements 
            into production-ready technical solutions. Using advanced AI models, it automates the entire 
            solution engineering lifecycle from requirements analysis to deployment.
        </p>
        
        <h3 style="color: #000000; margin-top: 2rem; margin-bottom: 1rem; font-weight: 600;"> How It Works</h3>
        <div style="display: grid; gap: 1rem;">
            <div style="padding: 1.25rem; background: #ffffff; border-radius: 8px; color: #000000; border: 2px solid #e2e8f0; border-left: 4px solid #2563eb; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <strong style="color: #000000;"> Phase 1-2:</strong> Upload BRD & AI Analysis
            </div>
            <div style="padding: 1.25rem; background: #ffffff; border-radius: 8px; color: #000000; border: 2px solid #e2e8f0; border-left: 4px solid #16a34a; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <strong style="color: #000000;"> Phase 3-4:</strong> Platform Selection & Architecture Design
            </div>
            <div style="padding: 1.25rem; background: #ffffff; border-radius: 8px; color: #000000; border: 2px solid #e2e8f0; border-left: 4px solid #0ea5e9; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <strong style="color: #000000;"> Phase 5-7:</strong> Artifacts, Data Model & Pipeline Design
            </div>
            <div style="padding: 1.25rem; background: #ffffff; border-radius: 8px; color: #000000; border: 2px solid #e2e8f0; border-left: 4px solid #16a34a; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <strong style="color: #000000;"> Phase 8-10:</strong> Analytics, Testing & Deployment
            </div>
        </div>
        
        <h3 style="color: #000000; margin-top: 2rem; margin-bottom: 1rem; font-weight: 600;"> Key Features</h3>
        <ul style="font-size: 1.05rem; line-height: 2; color: #000000;">
            <li> <strong>AI-Powered Analysis:</strong> Databricks Llama 3.3 70B for intelligent extraction</li>
            <li> <strong>Multi-Platform Support:</strong> Databricks, Fabric, AWS, Azure, Snowflake, GCP</li>
            <li> <strong>Production-Ready Code:</strong> Complete artifacts with schemas, configs, and scripts</li>
            <li> <strong>End-to-End Pipeline:</strong> From BRD to deployed solution in minutes</li>
            <li> <strong>Export Everything:</strong> JSON and Markdown downloads for all phases</li>
        </ul>
        
        <div style="margin-top: 2rem; padding: 1.25rem; background: #eff6ff; border: 2px solid #2563eb; border-radius: 8px; color: #000000;">
            <strong style="color: #000000;"> Pro Tip:</strong> Have your BRD ready in PDF, DOCX, or TXT format for the best results!
        </div>
    </div>
    """, unsafe_allow_html=True)

# PHASE 1: BRD Upload Section with Anchor
st.markdown('<div id="phase-upload"></div>', unsafe_allow_html=True)
st.markdown("""
<div style="background: #ffffff; padding: 1.25rem; border-radius: 8px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08); margin: 2rem 0; border-left: 3px solid #2563eb;">
    <h2 style="color: #000000; margin-bottom: 0.5rem; font-weight: 700; font-size: 1.5rem;">
        Phase 1: Upload Your Business Requirements Document
    </h2>
    <p style="color: #000000; font-size: 1.05rem; margin-bottom: 0;">
        Start your journey by uploading your BRD. We support multiple formats for your convenience.
    </p>
</div>
""", unsafe_allow_html=True)

# Minimal format cards
col_format1, col_format2, col_format3 = st.columns(3)

with col_format1:
    st.markdown("""
    <div style="background: #ffffff; padding: 1.5rem; border-radius: 8px; text-align: center; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);">
        <div style="width: 48px; height: 48px; background: #eff6ff; border-radius: 8px; display: inline-flex; align-items: center; justify-content: center; margin-bottom: 0.75rem;">
            <span style="font-size: 1.5rem; color: #000000;"></span>
        </div>
        <h4 style="color: #000000; margin: 0.5rem 0; font-weight: 600;">PDF</h4>
        <p style="color: #000000; font-size: 0.85rem; margin: 0;">.pdf files</p>
    </div>
    """, unsafe_allow_html=True)

with col_format2:
    st.markdown("""
    <div style="background: #ffffff; padding: 1.5rem; border-radius: 8px; text-align: center; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);">
        <div style="width: 48px; height: 48px; background: #f0fdf4; border-radius: 8px; display: inline-flex; align-items: center; justify-content: center; margin-bottom: 0.75rem;">
            <span style="font-size: 1.5rem; color: #000000;"></span>
        </div>
        <h4 style="color: #000000; margin: 0.5rem 0; font-weight: 600;">Word</h4>
        <p style="color: #000000; font-size: 0.85rem; margin: 0;">.docx files</p>
    </div>
    """, unsafe_allow_html=True)

with col_format3:
    st.markdown("""
    <div style="background: #ffffff; padding: 1.5rem; border-radius: 8px; text-align: center; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);">
        <div style="width: 48px; height: 48px; background: #f8fafc; border-radius: 8px; display: inline-flex; align-items: center; justify-content: center; margin-bottom: 0.75rem;">
            <span style="font-size: 1.5rem; color: #000000;"></span>
        </div>
        <h4 style="color: #000000; margin: 0.5rem 0; font-weight: 600;">Text</h4>
        <p style="color: #000000; font-size: 0.85rem; margin: 0;">.txt files</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# File uploader
uploaded_file = st.file_uploader(
    "Choose a file",
    type=["pdf", "docx", "txt"],
    help="Upload your Business Requirements Document"
)

if uploaded_file is not None:
    # Display file info
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("File Name", uploaded_file.name)
    with col2:
        st.metric("File Type", uploaded_file.type)
    with col3:
        file_size_mb = uploaded_file.size / (1024 * 1024)
        st.metric("File Size", f"{file_size_mb:.2f} MB")
    
    st.markdown("---")
    
    # Extract text based on file type
    with st.spinner(" Extracting text from document..."):
        try:
            extracted_text = ""
            
            if uploaded_file.type == "application/pdf":
                # Extract text from PDF
                pdf_reader = pypdf.PdfReader(io.BytesIO(uploaded_file.read()))
                extracted_text = ""
                for page_num, page in enumerate(pdf_reader.pages, 1):
                    page_text = page.extract_text()
                    extracted_text += f"\n--- Page {page_num} ---\n{page_text}"
                
                st.markdown('<div class="success-box"> Successfully extracted text from PDF</div>', unsafe_allow_html=True)
                st.info(f" Document Statistics: {len(pdf_reader.pages)} pages, {len(extracted_text)} characters")
            
            elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                # Extract text from DOCX
                doc = Document(io.BytesIO(uploaded_file.read()))
                extracted_text = "\n\n".join([para.text for para in doc.paragraphs if para.text.strip()])
                
                st.markdown('<div class="success-box"> Successfully extracted text from Word document</div>', unsafe_allow_html=True)
                st.info(f" Document Statistics: {len(doc.paragraphs)} paragraphs, {len(extracted_text)} characters")
            
            elif uploaded_file.type == "text/plain":
                # Extract text from TXT
                extracted_text = uploaded_file.read().decode("utf-8")
                
                st.markdown('<div class="success-box"> Successfully extracted text from text file</div>', unsafe_allow_html=True)
                st.info(f" Document Statistics: {len(extracted_text.splitlines())} lines, {len(extracted_text)} characters")
            
            # Display extracted text
            st.header(" Extracted Document Content")
            
            # Create tabs for different views
            tab1, tab2 = st.tabs([" Full Text", " Preview"])
            
            with tab1:
                st.text_area(
                    "Extracted Text",
                    value=extracted_text,
                    height=500,
                    help="Full extracted text from your document"
                )
                
                # Download button
                st.download_button(
                    label=" Download Extracted Text",
                    data=extracted_text,
                    file_name=f"{uploaded_file.name}_extracted.txt",
                    mime="text/plain"
                )
            
            with tab2:
                # Preview first 2000 characters
                preview_text = extracted_text[:2000]
                if len(extracted_text) > 2000:
                    preview_text += "\n\n... (text truncated for preview)"
                
                st.text_area(
                    "Preview (First 2000 characters)",
                    value=preview_text,
                    height=400
                )
            
            # Phase 2: AI-Powered Requirements Analysis
            st.markdown("---")
            st.markdown('<div id="phase-analysis"></div>', unsafe_allow_html=True)
            st.header("AI-Powered Requirements Analysis")
            
            st.markdown('<div class="info-box"> <b>Phase 2: AI Analysis</b><br>Click the button below to analyze your BRD using AI to extract structured requirements, data sources, KPIs, and business rules.</div>', unsafe_allow_html=True)
            
            # AI Analysis button
            col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
            with col_btn2:
                analyze_button = st.button(
                    " Analyze BRD with AI",
                    type="primary",
                    use_container_width=True
                )
            
            if analyze_button:
                with st.spinner(" AI is analyzing your document... This may take 30-60 seconds"):
                    try:
                        # Initialize Databricks SDK client
                        # This uses Databricks Foundation Model APIs - no external API key needed!
                        w = WorkspaceClient()
                        
                        # We'll use Databricks Foundation Model endpoint
                        # Using Llama 3.3 70B - powerful and reliable for document analysis
                        endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                        
                        # Prepare analysis prompt
                        analysis_prompt = f"""
You are an expert business analyst reviewing a Business Requirements Document (BRD). 
Analyze the following document and extract structured information in JSON format.

Document Text:
{extracted_text[:8000]}

Provide your analysis in the following JSON structure:
{{
    "executive_summary": "Brief overview of the project (2-3 sentences)",
    "business_objectives": ["List of key business objectives"],
    "data_sources": ["List of data sources mentioned"],
    "kpis": [{{
        "name": "KPI name",
        "description": "What it measures",
        "calculation": "How it's calculated (if mentioned)"
    }}],
    "business_rules": ["List of business rules and logic"],
    "technical_requirements": ["List of technical requirements"],
    "stakeholders": ["Key stakeholders mentioned"],
    "constraints": ["Timeline, budget, or technical constraints"],
    "functional_requirements": ["Key functional requirements"]
}}

Be thorough and extract as much structured information as possible. If a section is not found, return an empty array or "Not specified".
"""
                        
                        # Call Databricks Foundation Model API
                        response = w.serving_endpoints.query(
                            name=endpoint_name,
                            messages=[
                                ChatMessage(
                                    role=ChatMessageRole.SYSTEM,
                                    content="You are a helpful business analyst that extracts structured information from documents. Always respond with valid JSON only."
                                ),
                                ChatMessage(
                                    role=ChatMessageRole.USER,
                                    content=analysis_prompt
                                )
                            ],
                            temperature=0.3,
                            max_tokens=2000
                        )
                        
                        # Parse response
                        analysis_json = response.choices[0].message.content
                        
                        # Extract JSON from response (handle cases where model includes extra text)
                        if "```json" in analysis_json:
                            analysis_json = analysis_json.split("```json")[1].split("```")[0].strip()
                        elif "```" in analysis_json:
                            analysis_json = analysis_json.split("```")[1].split("```")[0].strip()
                        
                        analysis_results = json.loads(analysis_json)
                        
                        # Store in session state
                        st.session_state.analysis_complete = True
                        st.session_state.analysis_results = analysis_results
                        
                        st.success(" Analysis complete!")
                        st.rerun()
                            
                    except Exception as e:
                        st.error(f"Error during AI analysis: {str(e)}")
                        st.info(" Tip: Make sure you have access to Databricks Foundation Model APIs. The app uses Llama 3.3 70B model.")
                        
                        # Provide fallback instructions
                        with st.expander(" Troubleshooting"):
                            st.markdown("""
                            **Common Issues:**
                            1. **Model endpoint not available**: The endpoint 'databricks-meta-llama-3-3-70b-instruct' must be available in your workspace
                            2. **Permissions**: Make sure the app has permission to query serving endpoints
                            3. **Token limits**: Try with a smaller document if it's very large
                            
                            **Alternative**: Contact your Databricks admin to enable Foundation Model APIs
                            """)
            
            # Display analysis results if available
            if st.session_state.analysis_complete and st.session_state.analysis_results:
                st.markdown("---")
                st.subheader(" Analysis Results")
                
                results = st.session_state.analysis_results
                
                # Create tabs for different sections
                tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                    " Overview",
                    " Data & KPIs", 
                    " Business Rules",
                    " Technical",
                    " Stakeholders",
                    " Download"
                ])
                
                with tab1:
                    st.markdown("### Executive Summary")
                    st.info(results.get("executive_summary", "Not specified"))
                    
                    st.markdown("### Business Objectives")
                    objectives = results.get("business_objectives", [])
                    if objectives:
                        for i, obj in enumerate(objectives, 1):
                            st.markdown(f"{i}. {obj}")
                    else:
                        st.write("No objectives specified")
                    
                    st.markdown("### Functional Requirements")
                    func_reqs = results.get("functional_requirements", [])
                    if func_reqs:
                        for i, req in enumerate(func_reqs, 1):
                            st.markdown(f"- {req}")
                    else:
                        st.write("No functional requirements specified")
                
                with tab2:
                    st.markdown("### Data Sources")
                    data_sources = results.get("data_sources", [])
                    if data_sources:
                        for source in data_sources:
                            st.markdown(f"- {source}")
                    else:
                        st.write("No data sources identified")
                    
                    st.markdown("---")
                    st.markdown("### Key Performance Indicators (KPIs)")
                    kpis = results.get("kpis", [])
                    if kpis:
                        for kpi in kpis:
                            with st.expander(f" {kpi.get('name', 'Unnamed KPI')}"):
                                st.write(f"**Description:** {kpi.get('description', 'N/A')}")
                                st.write(f"**Calculation:** {kpi.get('calculation', 'Not specified')}")
                    else:
                        st.write("No KPIs identified")
                
                with tab3:
                    st.markdown("### Business Rules")
                    rules = results.get("business_rules", [])
                    if rules:
                        for i, rule in enumerate(rules, 1):
                            st.markdown(f"{i}. {rule}")
                    else:
                        st.write("No business rules specified")
                
                with tab4:
                    st.markdown("### Technical Requirements")
                    tech_reqs = results.get("technical_requirements", [])
                    if tech_reqs:
                        for req in tech_reqs:
                            st.markdown(f"- {req}")
                    else:
                        st.write("No technical requirements specified")
                    
                    st.markdown("---")
                    st.markdown("### Constraints")
                    constraints = results.get("constraints", [])
                    if constraints:
                        for constraint in constraints:
                            st.markdown(f"- {constraint}")
                    else:
                        st.write("No constraints specified")
                
                with tab5:
                    st.markdown("### Stakeholders")
                    stakeholders = results.get("stakeholders", [])
                    if stakeholders:
                        for stakeholder in stakeholders:
                            st.markdown(f"- {stakeholder}")
                    else:
                        st.write("No stakeholders identified")
                
                with tab6:
                    st.markdown("### Download Analysis Results")
                    
                    # JSON download
                    json_str = json.dumps(results, indent=2)
                    st.download_button(
                        label=" Download as JSON",
                        data=json_str,
                        file_name=f"{uploaded_file.name}_analysis.json",
                        mime="application/json"
                    )
                    
                    # Markdown summary download
                    md_summary = f"""# BRD Analysis: {uploaded_file.name}

## Executive Summary
{results.get('executive_summary', 'Not specified')}

## Business Objectives
{chr(10).join(f'- {obj}' for obj in results.get('business_objectives', []))}

## Data Sources
{chr(10).join(f'- {ds}' for ds in results.get('data_sources', []))}

## KPIs
{chr(10).join(f"- **{kpi.get('name')}**: {kpi.get('description', 'N/A')}" for kpi in results.get('kpis', []))}

## Business Rules
{chr(10).join(f'- {rule}' for rule in results.get('business_rules', []))}

## Technical Requirements
{chr(10).join(f'- {req}' for req in results.get('technical_requirements', []))}

## Stakeholders
{chr(10).join(f'- {sh}' for sh in results.get('stakeholders', []))}

## Constraints
{chr(10).join(f'- {con}' for con in results.get('constraints', []))}
"""
                    st.download_button(
                        label=" Download as Markdown",
                        data=md_summary,
                        file_name=f"{uploaded_file.name}_analysis.md",
                        mime="text/markdown"
                    )
                    
                    st.markdown("---")
                    st.success(" Phase 2 Complete! Proceed to Phase 3 below.")
                
                # Phase 3: Accelerator Selection
                if st.session_state.analysis_complete and st.session_state.analysis_results:
                    st.markdown("---")
                    st.markdown('<div id="phase-platform"></div>', unsafe_allow_html=True)
                    st.header(" Phase 3: Accelerator Selection")
                    
                    st.markdown('<div class="info-box"> <b>Phase 3: Choose Your Target Platform</b><br>Select the platform/accelerator that best fits your requirements. We\'ll analyze your BRD and provide recommendations.</div>', unsafe_allow_html=True)
                    
                    # Initialize session state for accelerator selection
                    if 'selected_accelerator' not in st.session_state:
                        st.session_state.selected_accelerator = None
                    if 'show_recommendations' not in st.session_state:
                        st.session_state.show_recommendations = False
                    
                    # Accelerator options with details
                    accelerators = {
                        "Databricks": {
                            "icon": "",
                            "tagline": "Data Intelligence Platform",
                            "description": "Unified platform for data engineering, data science, ML, and analytics on lakehouse architecture.",
                            "strengths": [
                                "Real-time data processing & streaming",
                                "Advanced ML/AI capabilities",
                                "Delta Lake for ACID transactions",
                                "Unity Catalog for governance",
                                "Seamless Spark integration"
                            ],
                            "use_cases": "Big data analytics, ML pipelines, real-time streaming, data lakehouse"
                        },
                        "Microsoft Fabric": {
                            "icon": "",
                            "tagline": "Unified Analytics Platform",
                            "description": "All-in-one analytics solution integrating Power BI, Data Factory, Synapse, and more.",
                            "strengths": [
                                "Unified SaaS experience",
                                "Native Power BI integration",
                                "OneLake data foundation",
                                "Microsoft ecosystem integration",
                                "Low-code/no-code options"
                            ],
                            "use_cases": "Enterprise BI, unified analytics, Microsoft-centric environments"
                        },
                        "AWS": {
                            "icon": "",
                            "tagline": "Cloud Data Services",
                            "description": "Comprehensive suite of cloud services including S3, Redshift, Glue, EMR, and Athena.",
                            "strengths": [
                                "Widest range of services",
                                "Serverless options (Athena, Glue)",
                                "S3 for data lake storage",
                                "Mature ecosystem",
                                "Pay-per-use pricing"
                            ],
                            "use_cases": "Cloud-native applications, data lakes, serverless analytics"
                        },
                        "Azure Synapse": {
                            "icon": "",
                            "tagline": "Analytics Service",
                            "description": "Limitless analytics service with data integration, enterprise data warehousing, and big data analytics.",
                            "strengths": [
                                "Unified workspace",
                                "Integrated with Azure services",
                                "Serverless & dedicated options",
                                "Built-in security & governance",
                                "Power BI integration"
                            ],
                            "use_cases": "Enterprise data warehousing, hybrid analytics, Azure ecosystem"
                        },
                        "Snowflake": {
                            "icon": "",
                            "tagline": "Data Cloud Platform",
                            "description": "Cloud-native data platform with separation of storage and compute, built for scale.",
                            "strengths": [
                                "Zero management overhead",
                                "Instant elasticity",
                                "Data sharing capabilities",
                                "Multi-cloud support",
                                "Time travel & cloning"
                            ],
                            "use_cases": "Data warehousing, data sharing, multi-cloud analytics"
                        },
                        "Google Cloud (BigQuery)": {
                            "icon": "",
                            "tagline": "Serverless Data Warehouse",
                            "description": "Fully managed, serverless data warehouse with built-in ML and real-time analytics.",
                            "strengths": [
                                "Serverless architecture",
                                "Built-in ML (BQML)",
                                "Real-time analytics",
                                "Pay-per-query pricing",
                                "Google ecosystem integration"
                            ],
                            "use_cases": "Serverless analytics, ML integration, Google Cloud workloads"
                        }
                    }
                    
                    # AI Recommendations button
                    col_rec1, col_rec2, col_rec3 = st.columns([1, 2, 1])
                    with col_rec2:
                        if st.button(" Get AI Recommendations", type="secondary", use_container_width=True):
                            st.session_state.show_recommendations = True
                            st.rerun()
                    
                    # Show AI recommendations if requested
                    if st.session_state.show_recommendations:
                        st.markdown("### AI-Powered Platform Recommendations")
                        
                        results = st.session_state.analysis_results
                        
                        # Simple recommendation logic based on analysis
                        recommendations = []
                        
                        # Check for real-time/streaming requirements
                        tech_reqs = ' '.join(results.get('technical_requirements', [])).lower()
                        objectives = ' '.join(results.get('business_objectives', [])).lower()
                        
                        if 'real-time' in tech_reqs or 'streaming' in tech_reqs:
                            recommendations.append(("Databricks", "Strong real-time and streaming capabilities", 95))
                        
                        if 'ml' in objectives or 'machine learning' in objectives or 'ai' in objectives:
                            recommendations.append(("Databricks", "Advanced ML/AI capabilities", 90))
                        
                        if 'microsoft' in tech_reqs or 'power bi' in tech_reqs or 'azure' in tech_reqs:
                            recommendations.append(("Microsoft Fabric", "Native Microsoft ecosystem integration", 92))
                            recommendations.append(("Azure Synapse", "Azure-native analytics service", 88))
                        
                        if 'aws' in tech_reqs or 's3' in tech_reqs:
                            recommendations.append(("AWS", "Comprehensive AWS service integration", 90))
                        
                        if 'data sharing' in objectives or 'collaboration' in objectives:
                            recommendations.append(("Snowflake", "Excellent data sharing capabilities", 87))
                        
                        if 'serverless' in tech_reqs or 'low maintenance' in tech_reqs:
                            recommendations.append(("Google Cloud (BigQuery)", "Fully serverless architecture", 89))
                            recommendations.append(("Snowflake", "Zero management overhead", 88))
                        
                        # If no specific recommendations, suggest based on general strengths
                        if not recommendations:
                            recommendations = [
                                ("Databricks", "Best for advanced analytics and ML workloads", 85),
                                ("Microsoft Fabric", "Best for unified enterprise analytics", 83),
                                ("Snowflake", "Best for scalable data warehousing", 82)
                            ]
                        
                        # Remove duplicates and sort by score
                        seen = set()
                        unique_recs = []
                        for rec in recommendations:
                            if rec[0] not in seen:
                                seen.add(rec[0])
                                unique_recs.append(rec)
                        unique_recs.sort(key=lambda x: x[2], reverse=True)
                        
                        # Display top 3 recommendations
                        for i, (platform, reason, score) in enumerate(unique_recs[:3], 1):
                            icon = accelerators[platform]["icon"]
                            if i == 1:
                                st.success(f"**#{i} Recommendation: {icon} {platform}** (Match: {score}%)")
                            else:
                                st.info(f"**#{i} Alternative: {icon} {platform}** (Match: {score}%)")
                            st.markdown(f"*{reason}*")
                    
                    st.markdown("---")
                    st.markdown("### Select Your Target Accelerator")
                    
                    # Display accelerator cards
                    cols = st.columns(2)
                    
                    for idx, (name, details) in enumerate(accelerators.items()):
                        with cols[idx % 2]:
                            with st.container():
                                st.markdown(f"### {details['icon']} {name}")
                                st.markdown(f"**{details['tagline']}**")
                                st.markdown(details['description'])
                                
                                with st.expander(" View Details"):
                                    st.markdown("**Key Strengths:**")
                                    for strength in details['strengths']:
                                        st.markdown(f"- {strength}")
                                    st.markdown(f"\n**Best For:** {details['use_cases']}")
                                
                                # Selection button
                                if st.button(f"Select {name}", key=f"select_{name}", use_container_width=True):
                                    st.session_state.selected_accelerator = name
                                    st.rerun()
                                
                                # Show if selected
                                if st.session_state.selected_accelerator == name:
                                    st.success(f" Selected: {name}")
                                
                                st.markdown("---")
                    
                    # Show summary if accelerator is selected
                    if st.session_state.selected_accelerator:
                        st.markdown("---")
                        st.markdown("###  Phase 3 Summary")
                        
                        selected = st.session_state.selected_accelerator
                        st.markdown(f'<div class="success-box"> <b>Selected Platform:</b> {accelerators[selected]["icon"]} {selected}<br><b>Tagline:</b> {accelerators[selected]["tagline"]}</div>', unsafe_allow_html=True)
                        
                        col_sum1, col_sum2 = st.columns(2)
                        
                        with col_sum1:
                            st.markdown("**Your BRD Analysis:**")
                            st.markdown(f"- Business Objectives: {len(results.get('business_objectives', []))}")
                            st.markdown(f"- Data Sources: {len(results.get('data_sources', []))}")
                            st.markdown(f"- KPIs: {len(results.get('kpis', []))}")
                        
                        with col_sum2:
                            st.markdown("**Next Steps:**")
                            st.markdown("-  Requirements analyzed")
                            st.markdown("-  Platform selected")
                            st.markdown("-  Ready for Phase 4: Solution Architecture")
                        
                        st.success(" Phase 3 Complete! Proceed to Phase 4 below.")
                
                # Phase 4: Solution Architecture Design
                if st.session_state.selected_accelerator:
                    st.markdown("---")
                    st.markdown('<div id="phase-architecture"></div>', unsafe_allow_html=True)
                    st.header(" Phase 4: Solution Architecture Design")
                    
                    st.markdown('<div class="info-box"> <b>Phase 4: Generate Solution Architecture</b><br>AI will design a complete solution architecture based on your BRD and selected platform.</div>', unsafe_allow_html=True)
                    
                    # Initialize session state for architecture
                    if 'architecture_generated' not in st.session_state:
                        st.session_state.architecture_generated = False
                    if 'architecture_design' not in st.session_state:
                        st.session_state.architecture_design = None
                    
                    # Show current selections summary
                    col_summary1, col_summary2 = st.columns(2)
                    with col_summary1:
                        st.markdown("** Your Selections:**")
                        st.info(f"**Platform:** {st.session_state.selected_accelerator}")
                    
                    with col_summary2:
                        results = st.session_state.analysis_results
                        st.markdown("** BRD Summary:**")
                        st.info(f"{len(results.get('business_objectives', []))} objectives, {len(results.get('data_sources', []))} data sources, {len(results.get('kpis', []))} KPIs")
                    
                    # Generate Architecture button
                    col_arch1, col_arch2, col_arch3 = st.columns([1, 2, 1])
                    with col_arch2:
                        if st.button("Generate Solution Architecture", type="primary", use_container_width=True, disabled=st.session_state.architecture_generated):
                            with st.spinner(" AI is designing your solution architecture... This may take 30-60 seconds"):
                                try:
                                    # Initialize Databricks SDK client
                                    w = WorkspaceClient()
                                    endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                                    
                                    # Prepare architecture design prompt
                                    selected_platform = st.session_state.selected_accelerator
                                    brd_summary = json.dumps(results, indent=2)
                                    
                                    architecture_prompt = f"""
You are an expert solution architect. Design a complete technical solution architecture for the following requirements.

**Target Platform:** {selected_platform}

**BRD Analysis:**
{brd_summary[:6000]}

Provide a comprehensive solution architecture in JSON format with the following structure:
{{
    "architecture_overview": "High-level description of the solution (3-4 sentences)",
    "data_layers": {{
        "ingestion": [{{
            "component": "Component name",
            "technology": "Specific technology/service",
            "purpose": "What it does"
        }}],
        "storage": [{{
            "component": "Component name",
            "technology": "Specific technology/service",
            "purpose": "What it does"
        }}],
        "processing": [{{
            "component": "Component name",
            "technology": "Specific technology/service",
            "purpose": "What it does"
        }}],
        "serving": [{{
            "component": "Component name",
            "technology": "Specific technology/service",
            "purpose": "What it does"
        }}]
    }},
    "technology_stack": {{
        "data_platform": ["List of data platform technologies"],
        "processing_engines": ["List of processing engines"],
        "analytics_tools": ["List of analytics tools"],
        "orchestration": ["Workflow orchestration tools"],
        "monitoring": ["Monitoring and observability tools"]
    }},
    "integration_points": [
        {{
            "source": "Source system",
            "target": "Target system",
            "method": "Integration method",
            "frequency": "Data refresh frequency"
        }}
    ],
    "security_governance": {{
        "authentication": "Authentication approach",
        "authorization": "Authorization model",
        "data_governance": "Governance framework",
        "compliance": ["Compliance requirements"]
    }},
    "deployment_architecture": {{
        "environments": ["List of environments (dev, staging, prod)"],
        "cicd": "CI/CD approach",
        "infrastructure": "Infrastructure model (serverless, provisioned, etc.)"
    }},
    "scalability_considerations": ["List of scalability considerations"],
    "estimated_timeline": {{
        "setup": "Time to set up infrastructure",
        "development": "Time for development",
        "testing": "Time for testing",
        "deployment": "Time for deployment"
    }}
}}

Be specific to the {selected_platform} platform. Use actual service names and best practices.
"""
                                    
                                    # Call Databricks Foundation Model API
                                    response = w.serving_endpoints.query(
                                        name=endpoint_name,
                                        messages=[
                                            ChatMessage(
                                                role=ChatMessageRole.SYSTEM,
                                                content="You are an expert solution architect specializing in data platforms. Always respond with valid JSON only."
                                            ),
                                            ChatMessage(
                                                role=ChatMessageRole.USER,
                                                content=architecture_prompt
                                            )
                                        ],
                                        temperature=0.4,
                                        max_tokens=3000
                                    )
                                    
                                    # Parse response
                                    architecture_json = response.choices[0].message.content
                                    
                                    # Extract JSON from response
                                    if "```json" in architecture_json:
                                        architecture_json = architecture_json.split("```json")[1].split("```")[0].strip()
                                    elif "```" in architecture_json:
                                        architecture_json = architecture_json.split("```")[1].split("```")[0].strip()
                                    
                                    # Parse with error handling
                                    try:
                                        architecture_design = json.loads(architecture_json, strict=False)
                                    except json.JSONDecodeError:
                                        # Try to fix common issues
                                        architecture_json_fixed = architecture_json.replace('\\n', '\\\\n').replace('\\t', '\\\\t')
                                        architecture_design = json.loads(architecture_json_fixed, strict=False)
                                    
                                    # Store in session state
                                    st.session_state.architecture_generated = True
                                    st.session_state.architecture_design = architecture_design
                                    
                                    st.success(" Architecture design complete!")
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"Error generating architecture: {str(e)}")
                                    st.info(" Tip: Make sure you have access to Databricks Foundation Model APIs.")
                    
                    # Display architecture if generated
                    if st.session_state.architecture_generated and st.session_state.architecture_design:
                        st.markdown("---")
                        st.subheader(" Solution Architecture Design")
                        
                        arch = st.session_state.architecture_design
                        
                        # Architecture Overview
                        st.markdown("###  Architecture Overview")
                        st.info(arch.get('architecture_overview', 'Not specified'))
                        
                        # Create tabs for different architecture views
                        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                            " Data Layers",
                            " Technology Stack",
                            " Integrations",
                            " Security",
                            "Deployment",
                            " Download"
                        ])
                        
                        with tab1:
                            st.markdown("### Data Architecture Layers")
                            
                            data_layers = arch.get('data_layers', {})
                            
                            # Ingestion Layer
                            with st.expander(" Ingestion Layer", expanded=True):
                                ingestion = data_layers.get('ingestion', [])
                                if ingestion:
                                    for comp in ingestion:
                                        st.markdown(f"**{comp.get('component', 'N/A')}**")
                                        st.markdown(f"- *Technology:* {comp.get('technology', 'N/A')}")
                                        st.markdown(f"- *Purpose:* {comp.get('purpose', 'N/A')}")
                                        st.markdown("")
                                else:
                                    st.write("No ingestion components specified")
                            
                            # Storage Layer
                            with st.expander(" Storage Layer", expanded=True):
                                storage = data_layers.get('storage', [])
                                if storage:
                                    for comp in storage:
                                        st.markdown(f"**{comp.get('component', 'N/A')}**")
                                        st.markdown(f"- *Technology:* {comp.get('technology', 'N/A')}")
                                        st.markdown(f"- *Purpose:* {comp.get('purpose', 'N/A')}")
                                        st.markdown("")
                                else:
                                    st.write("No storage components specified")
                            
                            # Processing Layer
                            with st.expander(" Processing Layer", expanded=True):
                                processing = data_layers.get('processing', [])
                                if processing:
                                    for comp in processing:
                                        st.markdown(f"**{comp.get('component', 'N/A')}**")
                                        st.markdown(f"- *Technology:* {comp.get('technology', 'N/A')}")
                                        st.markdown(f"- *Purpose:* {comp.get('purpose', 'N/A')}")
                                        st.markdown("")
                                else:
                                    st.write("No processing components specified")
                            
                            # Serving Layer
                            with st.expander(" Serving Layer", expanded=True):
                                serving = data_layers.get('serving', [])
                                if serving:
                                    for comp in serving:
                                        st.markdown(f"**{comp.get('component', 'N/A')}**")
                                        st.markdown(f"- *Technology:* {comp.get('technology', 'N/A')}")
                                        st.markdown(f"- *Purpose:* {comp.get('purpose', 'N/A')}")
                                        st.markdown("")
                                else:
                                    st.write("No serving components specified")
                        
                        with tab2:
                            st.markdown("### Technology Stack")
                            
                            tech_stack = arch.get('technology_stack', {})
                            
                            col_tech1, col_tech2 = st.columns(2)
                            
                            with col_tech1:
                                st.markdown("**Data Platform**")
                                for tech in tech_stack.get('data_platform', []):
                                    st.markdown(f"- {tech}")
                                
                                st.markdown("\n**Processing Engines**")
                                for tech in tech_stack.get('processing_engines', []):
                                    st.markdown(f"- {tech}")
                                
                                st.markdown("\n**Analytics Tools**")
                                for tech in tech_stack.get('analytics_tools', []):
                                    st.markdown(f"- {tech}")
                            
                            with col_tech2:
                                st.markdown("**Orchestration**")
                                for tech in tech_stack.get('orchestration', []):
                                    st.markdown(f"- {tech}")
                                
                                st.markdown("\n**Monitoring**")
                                for tech in tech_stack.get('monitoring', []):
                                    st.markdown(f"- {tech}")
                            
                            st.markdown("---")
                            st.markdown("### ⏱ Estimated Timeline")
                            timeline = arch.get('estimated_timeline', {})
                            col_t1, col_t2, col_t3, col_t4 = st.columns(4)
                            with col_t1:
                                st.metric("Setup", timeline.get('setup', 'N/A'))
                            with col_t2:
                                st.metric("Development", timeline.get('development', 'N/A'))
                            with col_t3:
                                st.metric("Testing", timeline.get('testing', 'N/A'))
                            with col_t4:
                                st.metric("Deployment", timeline.get('deployment', 'N/A'))
                        
                        with tab3:
                            st.markdown("### Integration Points")
                            
                            integrations = arch.get('integration_points', [])
                            if integrations:
                                for i, integration in enumerate(integrations, 1):
                                    with st.expander(f" Integration {i}: {integration.get('source', 'N/A')} → {integration.get('target', 'N/A')}"):
                                        st.markdown(f"**Source:** {integration.get('source', 'N/A')}")
                                        st.markdown(f"**Target:** {integration.get('target', 'N/A')}")
                                        st.markdown(f"**Method:** {integration.get('method', 'N/A')}")
                                        st.markdown(f"**Frequency:** {integration.get('frequency', 'N/A')}")
                            else:
                                st.write("No integration points specified")
                            
                            st.markdown("---")
                            st.markdown("###  Scalability Considerations")
                            scalability = arch.get('scalability_considerations', [])
                            if scalability:
                                for consideration in scalability:
                                    st.markdown(f"- {consideration}")
                            else:
                                st.write("No scalability considerations specified")
                        
                        with tab4:
                            st.markdown("### Security & Governance")
                            
                            security = arch.get('security_governance', {})
                            
                            col_sec1, col_sec2 = st.columns(2)
                            
                            with col_sec1:
                                st.markdown("** Authentication**")
                                st.info(security.get('authentication', 'Not specified'))
                                
                                st.markdown("** Authorization**")
                                st.info(security.get('authorization', 'Not specified'))
                            
                            with col_sec2:
                                st.markdown("** Data Governance**")
                                st.info(security.get('data_governance', 'Not specified'))
                                
                                st.markdown("** Compliance**")
                                compliance = security.get('compliance', [])
                                if compliance:
                                    for req in compliance:
                                        st.markdown(f"- {req}")
                                else:
                                    st.write("No compliance requirements specified")
                        
                        with tab5:
                            st.markdown("### Deployment Architecture")
                            
                            deployment = arch.get('deployment_architecture', {})
                            
                            st.markdown("** Environments**")
                            environments = deployment.get('environments', [])
                            if environments:
                                cols = st.columns(len(environments))
                                for i, env in enumerate(environments):
                                    with cols[i]:
                                        st.info(env)
                            else:
                                st.write("No environments specified")
                            
                            st.markdown("\n** CI/CD Pipeline**")
                            st.info(deployment.get('cicd', 'Not specified'))
                            
                            st.markdown("** Infrastructure Model**")
                            st.info(deployment.get('infrastructure', 'Not specified'))
                        
                        with tab6:
                            st.markdown("### Download Architecture Design")
                            
                            # JSON download
                            json_str = json.dumps(arch, indent=2)
                            st.download_button(
                                label=" Download as JSON",
                                data=json_str,
                                file_name=f"{st.session_state.selected_accelerator}_architecture.json",
                                mime="application/json"
                            )
                            
                            # Create markdown document
                            md_doc = f"""# Solution Architecture: {st.session_state.selected_accelerator}

## Architecture Overview
{arch.get('architecture_overview', 'Not specified')}

## Data Layers

### Ingestion Layer
{chr(10).join(f"- **{c.get('component', 'N/A')}**: {c.get('technology', 'N/A')} - {c.get('purpose', 'N/A')}" for c in arch.get('data_layers', {}).get('ingestion', []))}

### Storage Layer
{chr(10).join(f"- **{c.get('component', 'N/A')}**: {c.get('technology', 'N/A')} - {c.get('purpose', 'N/A')}" for c in arch.get('data_layers', {}).get('storage', []))}

### Processing Layer
{chr(10).join(f"- **{c.get('component', 'N/A')}**: {c.get('technology', 'N/A')} - {c.get('purpose', 'N/A')}" for c in arch.get('data_layers', {}).get('processing', []))}

### Serving Layer
{chr(10).join(f"- **{c.get('component', 'N/A')}**: {c.get('technology', 'N/A')} - {c.get('purpose', 'N/A')}" for c in arch.get('data_layers', {}).get('serving', []))}

## Technology Stack

### Data Platform
{chr(10).join(f'- {t}' for t in arch.get('technology_stack', {}).get('data_platform', []))}

### Processing Engines
{chr(10).join(f'- {t}' for t in arch.get('technology_stack', {}).get('processing_engines', []))}

### Analytics Tools
{chr(10).join(f'- {t}' for t in arch.get('technology_stack', {}).get('analytics_tools', []))}

## Security & Governance

- **Authentication**: {arch.get('security_governance', {}).get('authentication', 'Not specified')}
- **Authorization**: {arch.get('security_governance', {}).get('authorization', 'Not specified')}
- **Data Governance**: {arch.get('security_governance', {}).get('data_governance', 'Not specified')}

## Estimated Timeline

- **Setup**: {arch.get('estimated_timeline', {}).get('setup', 'N/A')}
- **Development**: {arch.get('estimated_timeline', {}).get('development', 'N/A')}
- **Testing**: {arch.get('estimated_timeline', {}).get('testing', 'N/A')}
- **Deployment**: {arch.get('estimated_timeline', {}).get('deployment', 'N/A')}
"""
                            st.download_button(
                                label=" Download as Markdown",
                                data=md_doc,
                                file_name=f"{st.session_state.selected_accelerator}_architecture.md",
                                mime="text/markdown"
                            )
                            
                            st.markdown("---")
                            st.success(" Phase 4 Complete! Proceed to Phase 5 below.")
                
                # Phase 5: Technical Artifact Generation
                if st.session_state.architecture_generated and st.session_state.architecture_design:
                    st.markdown("---")
                    st.markdown('<div id="phase-artifacts"></div>', unsafe_allow_html=True)
                    st.header(" Phase 5: Technical Artifact Generation")
                    
                    st.markdown('<div class="info-box"> <b>Phase 5: Generate Build Artifacts</b><br>AI will generate ready-to-use code, configurations, schemas, and deployment scripts for your solution.</div>', unsafe_allow_html=True)
                    
                    # Initialize session state for artifacts
                    if 'artifacts_generated' not in st.session_state:
                        st.session_state.artifacts_generated = False
                    if 'generated_artifacts' not in st.session_state:
                        st.session_state.generated_artifacts = None
                    
                    # Show summary
                    st.markdown("** Artifacts to Generate:**")
                    col_art1, col_art2, col_art3, col_art4 = st.columns(4)
                    with col_art1:
                        st.info(" Pipeline Code")
                    with col_art2:
                        st.info(" Schema Definitions")
                    with col_art3:
                        st.info(" Config Files")
                    with col_art4:
                        st.info("Deployment Scripts")
                    
                    # Generate Artifacts button
                    col_gen1, col_gen2, col_gen3 = st.columns([1, 2, 1])
                    with col_gen2:
                        if st.button(" Generate Technical Artifacts", type="primary", use_container_width=True, disabled=st.session_state.artifacts_generated):
                            with st.spinner(" AI is generating your technical artifacts... This may take 60-90 seconds"):
                                try:
                                    # Initialize Databricks SDK client
                                    w = WorkspaceClient()
                                    endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                                    
                                    # Prepare artifact generation prompt
                                    selected_platform = st.session_state.selected_accelerator
                                    brd_summary = json.dumps(st.session_state.analysis_results, indent=2)[:4000]
                                    arch_summary = json.dumps(st.session_state.architecture_design, indent=2)[:4000]
                                    
                                    artifacts_prompt = f"""
You are an expert developer. Generate production-ready technical artifacts for the following solution.

**Target Platform:** {selected_platform}

**BRD Summary:**
{brd_summary}

**Architecture Summary:**
{arch_summary}

Generate the following artifacts in VALID JSON format. CRITICAL: All code in 'code' and 'content' fields must have newlines escaped as \\n and tabs as \\t:
{{
    "data_pipeline": {{
        "description": "Purpose and overview of the pipeline",
        "code": "Complete pipeline code (Python/SQL) for {selected_platform}",
        "language": "python or sql",
        "filename": "Suggested filename"
    }},
    "schema_definitions": [
        {{
            "table_name": "Table name",
            "schema_type": "bronze/silver/gold or fact/dim",
            "ddl": "CREATE TABLE SQL statement",
            "description": "Table purpose"
        }}
    ],
    "configuration_files": [
        {{
            "filename": "Config filename",
            "format": "yaml/json/properties",
            "content": "Full configuration content",
            "purpose": "What this config does"
        }}
    ],
    "deployment_script": {{
        "script_type": "bash/terraform/arm/cloudformation",
        "content": "Complete deployment script",
        "filename": "Suggested filename",
        "description": "What the script deploys"
    }},
    "data_quality_checks": [
        {{
            "check_name": "Quality check name",
            "code": "Code for the quality check",
            "description": "What it validates"
        }}
    ],
    "documentation": {{
        "readme": "Complete README.md content",
        "setup_guide": "Step-by-step setup instructions"
    }}
}}

Generate REAL, WORKING code specific to {selected_platform}. Use actual service names and best practices.
"""
                                    
                                    # Call Databricks Foundation Model API
                                    response = w.serving_endpoints.query(
                                        name=endpoint_name,
                                        messages=[
                                            ChatMessage(
                                                role=ChatMessageRole.SYSTEM,
                                                content="You are an expert developer who generates production-ready code. Always respond with valid JSON. IMPORTANT: In JSON strings, escape all newlines as \\n, tabs as \\t, and backslashes as \\\\. All code must be in properly escaped JSON strings."
                                            ),
                                            ChatMessage(
                                                role=ChatMessageRole.USER,
                                                content=artifacts_prompt
                                            )
                                        ],
                                        temperature=0.3,
                                        max_tokens=4000
                                    )
                                    
                                    # Parse response
                                    artifacts_json = response.choices[0].message.content
                                    
                                    # Extract JSON from response
                                    if "```json" in artifacts_json:
                                        artifacts_json = artifacts_json.split("```json")[1].split("```")[0].strip()
                                    elif "```" in artifacts_json:
                                        artifacts_json = artifacts_json.split("```")[1].split("```")[0].strip()
                                    
                                    # Clean JSON - handle control characters
                                    # Try parsing with strict=False to allow control characters
                                    try:
                                        generated_artifacts = json.loads(artifacts_json, strict=False)
                                    except json.JSONDecodeError as json_err:
                                        # If that fails, try to fix common issues
                                        # Replace literal newlines in strings with escaped newlines
                                        import re
                                        # Find JSON string values and escape control chars within them
                                        artifacts_json_fixed = artifacts_json.replace('\\n', '\\\\n').replace('\\t', '\\\\t')
                                        try:
                                            generated_artifacts = json.loads(artifacts_json_fixed, strict=False)
                                        except Exception as e:
                                            # Last resort: create a simplified structure
                                            st.warning(f"AI response had formatting issues: {str(e)}. Using simplified output.")
                                            generated_artifacts = {
                                                "data_pipeline": {
                                                    "description": "Data pipeline for " + selected_platform,
                                                    "code": "# Pipeline code\n# See raw AI response for details",
                                                    "language": "python",
                                                    "filename": "pipeline.py"
                                                },
                                                "schema_definitions": [],
                                                "configuration_files": [],
                                                "deployment_script": {
                                                    "script_type": "bash",
                                                    "content": "# Deployment script",
                                                    "filename": "deploy.sh",
                                                    "description": "Deployment automation"
                                                },
                                                "data_quality_checks": [],
                                                "documentation": {
                                                    "readme": "# Project README\n\nSee architecture and BRD analysis for details.",
                                                    "setup_guide": "# Setup Guide\n\n1. Review architecture\n2. Configure environment\n3. Deploy solution"
                                                }
                                            }
                                    
                                    # Store in session state
                                    st.session_state.artifacts_generated = True
                                    st.session_state.generated_artifacts = generated_artifacts
                                    
                                    st.success(" Artifacts generated successfully!")
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"Error generating artifacts: {str(e)}")
                                    st.info(" Tip: This requires Databricks Foundation Model APIs access.")
                    
                    # Display generated artifacts
                    if st.session_state.artifacts_generated and st.session_state.generated_artifacts:
                        st.markdown("---")
                        st.subheader("Generated Technical Artifacts")
                        
                        artifacts = st.session_state.generated_artifacts
                        
                        # Create tabs for different artifact types
                        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                            " Pipeline Code",
                            " Schemas",
                            "Configs",
                            "Deployment",
                            " Quality Checks",
                            " Documentation"
                        ])
                        
                        with tab1:
                            st.markdown("### Data Pipeline Code")
                            
                            pipeline = artifacts.get('data_pipeline', {})
                            st.info(pipeline.get('description', 'No description'))
                            
                            pipeline_code = pipeline.get('code', 'No code generated')
                            pipeline_lang = pipeline.get('language', 'python')
                            pipeline_filename = pipeline.get('filename', 'pipeline.py')
                            
                            st.code(pipeline_code, language=pipeline_lang)
                            
                            # Download button
                            st.download_button(
                                label=f" Download {pipeline_filename}",
                                data=pipeline_code,
                                file_name=pipeline_filename,
                                mime="text/plain"
                            )
                        
                        with tab2:
                            st.markdown("### Schema Definitions")
                            
                            schemas = artifacts.get('schema_definitions', [])
                            if schemas:
                                for schema in schemas:
                                    with st.expander(f" {schema.get('table_name', 'Unknown')} ({schema.get('schema_type', 'N/A')})"):
                                        st.markdown(f"**Description:** {schema.get('description', 'N/A')}")
                                        st.code(schema.get('ddl', 'No DDL provided'), language='sql')
                                
                                # Download all schemas
                                all_schemas = "\n\n".join([
                                    f"-- {s.get('table_name', 'Unknown')}\n-- {s.get('description', 'N/A')}\n{s.get('ddl', '')}" 
                                    for s in schemas
                                ])
                                st.download_button(
                                    label=" Download All Schemas",
                                    data=all_schemas,
                                    file_name="schemas.sql",
                                    mime="text/plain"
                                )
                            else:
                                st.write("No schema definitions generated")
                        
                        with tab3:
                            st.markdown("### Configuration Files")
                            
                            configs = artifacts.get('configuration_files', [])
                            if configs:
                                for config in configs:
                                    with st.expander(f" {config.get('filename', 'Unknown')}"):
                                        st.markdown(f"**Purpose:** {config.get('purpose', 'N/A')}")
                                        st.markdown(f"**Format:** {config.get('format', 'N/A')}")
                                        st.code(config.get('content', 'No content'), language=config.get('format', 'yaml'))
                                        
                                        # Download individual config
                                        st.download_button(
                                            label=f" Download {config.get('filename', 'config')}",
                                            data=config.get('content', ''),
                                            file_name=config.get('filename', 'config.yaml'),
                                            mime="text/plain",
                                            key=f"download_{config.get('filename', 'config')}"
                                        )
                            else:
                                st.write("No configuration files generated")
                        
                        with tab4:
                            st.markdown("### Deployment Script")
                            
                            deployment = artifacts.get('deployment_script', {})
                            st.info(deployment.get('description', 'No description'))
                            
                            deploy_content = deployment.get('content', 'No script generated')
                            deploy_type = deployment.get('script_type', 'bash')
                            deploy_filename = deployment.get('filename', 'deploy.sh')
                            
                            st.code(deploy_content, language=deploy_type)
                            
                            # Download button
                            st.download_button(
                                label=f" Download {deploy_filename}",
                                data=deploy_content,
                                file_name=deploy_filename,
                                mime="text/plain"
                            )
                            
                            st.warning("Review and customize the deployment script before running in production.")
                        
                        with tab5:
                            st.markdown("### Data Quality Checks")
                            
                            quality_checks = artifacts.get('data_quality_checks', [])
                            if quality_checks:
                                for check in quality_checks:
                                    with st.expander(f" {check.get('check_name', 'Unknown')}"):
                                        st.markdown(f"**Description:** {check.get('description', 'N/A')}")
                                        st.code(check.get('code', 'No code'), language='python')
                                
                                # Download all quality checks
                                all_checks = "\n\n".join([
                                    f"# {c.get('check_name', 'Unknown')}\n# {c.get('description', 'N/A')}\n{c.get('code', '')}" 
                                    for c in quality_checks
                                ])
                                st.download_button(
                                    label=" Download All Quality Checks",
                                    data=all_checks,
                                    file_name="quality_checks.py",
                                    mime="text/plain"
                                )
                            else:
                                st.write("No quality checks generated")
                        
                        with tab6:
                            st.markdown("### Documentation")
                            
                            docs = artifacts.get('documentation', {})
                            
                            # README
                            st.markdown("####  README.md")
                            readme = docs.get('readme', 'No README generated')
                            st.markdown(readme)
                            
                            st.download_button(
                                label=" Download README.md",
                                data=readme,
                                file_name="README.md",
                                mime="text/markdown",
                                key="download_readme"
                            )
                            
                            st.markdown("---")
                            
                            # Setup Guide
                            st.markdown("####  Setup Guide")
                            setup_guide = docs.get('setup_guide', 'No setup guide generated')
                            st.markdown(setup_guide)
                            
                            st.download_button(
                                label=" Download Setup Guide",
                                data=setup_guide,
                                file_name="SETUP.md",
                                mime="text/markdown",
                                key="download_setup"
                            )
                        
                        # Download all artifacts as ZIP (conceptual)
                        st.markdown("---")
                        st.markdown("###  Download All Artifacts")
                        
                        # Create a combined download with all artifacts
                        # Pre-compute joined strings to avoid f-string backslash issues
                        schemas_joined = "\n".join([s.get('ddl', '') for s in artifacts.get('schema_definitions', [])])
                        configs_joined = "\n".join([f"### {c.get('filename', 'Unknown')}\n```{c.get('format', 'yaml')}\n{c.get('content', '')}\n```" for c in artifacts.get('configuration_files', [])])
                        
                        all_artifacts_md = f"""# Technical Artifacts - {st.session_state.selected_accelerator}

## Pipeline Code
```{pipeline.get('language', 'python')}
{pipeline.get('code', '')}
```

## Schema Definitions
```sql
{schemas_joined}
```

## Configuration Files
{configs_joined}

## Deployment Script
```{deployment.get('script_type', 'bash')}
{deployment.get('content', '')}
```

## Documentation
{docs.get('readme', '')}

## Setup Guide
{docs.get('setup_guide', '')}
"""
                        
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            st.download_button(
                                label=" Download All Artifacts (Markdown)",
                                data=all_artifacts_md,
                                file_name=f"{st.session_state.selected_accelerator}_artifacts.md",
                                mime="text/markdown",
                                use_container_width=True
                            )
                        
                        with col_dl2:
                            # JSON export
                            json_artifacts = json.dumps(artifacts, indent=2)
                            st.download_button(
                                label=" Download All Artifacts (JSON)",
                                data=json_artifacts,
                                file_name=f"{st.session_state.selected_accelerator}_artifacts.json",
                                mime="application/json",
                                use_container_width=True
                            )
                        
                        st.markdown("---")
                        st.success(" Phase 5 Complete! Ready for Phase 6: Data Model & Schema Design")
                        
                        st.markdown("###  What You've Accomplished So Far:")
                        col_prog1, col_prog2, col_prog3, col_prog4, col_prog5 = st.columns(5)
                        with col_prog1:
                            st.markdown(" **Phase 1**\nBRD Upload")
                        with col_prog2:
                            st.markdown(" **Phase 2**\nAI Analysis")
                        with col_prog3:
                            st.markdown(" **Phase 3**\nPlatform Selection")
                        with col_prog4:
                            st.markdown(" **Phase 4**\nArchitecture Design")
                        with col_prog5:
                            st.markdown(" **Phase 5**\nArtifact Generation")
                
                # Phase 6: Data Model & Schema Design
                if st.session_state.artifacts_generated and st.session_state.generated_artifacts:
                    st.markdown("---")
                    st.markdown('<div id="phase-datamodel"></div>', unsafe_allow_html=True)
                    st.header(" Phase 6: Data Model & Schema Design")
                    
                    st.markdown('<div class="info-box"> <b>Phase 6: Design Your Data Model</b><br>AI will create a comprehensive data model with entity relationships, dimensional design, and data lineage.</div>', unsafe_allow_html=True)
                    
                    # Initialize session state for data model
                    if 'data_model_generated' not in st.session_state:
                        st.session_state.data_model_generated = False
                    if 'data_model_design' not in st.session_state:
                        st.session_state.data_model_design = None
                    
                    # Show summary
                    st.markdown("** Data Model Components:**")
                    col_dm1, col_dm2, col_dm3, col_dm4 = st.columns(4)
                    with col_dm1:
                        st.info(" Entity Model")
                    with col_dm2:
                        st.info("⭐ Dimensional Design")
                    with col_dm3:
                        st.info(" Relationships")
                    with col_dm4:
                        st.info(" Data Lineage")
                    
                    # Generate Data Model button
                    col_dm_gen1, col_dm_gen2, col_dm_gen3 = st.columns([1, 2, 1])
                    with col_dm_gen2:
                        if st.button(" Generate Data Model", type="primary", use_container_width=True, disabled=st.session_state.data_model_generated):
                            with st.spinner(" AI is designing your data model... This may take 60-90 seconds"):
                                try:
                                    # Initialize Databricks SDK client
                                    w = WorkspaceClient()
                                    endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                                    
                                    # Prepare data model prompt
                                    selected_platform = st.session_state.selected_accelerator
                                    brd_summary = json.dumps(st.session_state.analysis_results, indent=2)[:3000]
                                    arch_summary = json.dumps(st.session_state.architecture_design, indent=2)[:3000]
                                    artifacts_summary = json.dumps(st.session_state.generated_artifacts, indent=2)[:3000]
                                    
                                    data_model_prompt = f"""
You are an expert data modeler. Design a comprehensive data model for the following solution.

**Target Platform:** {selected_platform}

**BRD Summary:**
{brd_summary}

**Architecture:**
{arch_summary}

**Artifacts:**
{artifacts_summary}

Generate a complete data model in JSON format:
{{
    "modeling_approach": "Star Schema / Snowflake / Data Vault / Normalized / Lakehouse",
    "model_description": "Overview of the data model approach and rationale",
    "entities": [
        {{
            "entity_name": "Entity/Table name",
            "entity_type": "fact/dimension/bridge/staging/raw",
            "description": "Purpose of this entity",
            "grain": "Data granularity",
            "columns": [
                {{
                    "column_name": "Column name",
                    "data_type": "Data type",
                    "is_primary_key": true/false,
                    "is_foreign_key": true/false,
                    "nullable": true/false,
                    "description": "Column description"
                }}
            ]
        }}
    ],
    "relationships": [
        {{
            "from_entity": "Source entity",
            "to_entity": "Target entity",
            "relationship_type": "1:1 / 1:N / N:M",
            "foreign_key": "FK column name",
            "description": "Relationship description"
        }}
    ],
    "dimensional_model": {{
        "fact_tables": [
            {{
                "name": "Fact table name",
                "measures": ["List of measures"],
                "dimensions": ["List of dimension references"],
                "grain": "Transaction grain"
            }}
        ],
        "dimension_tables": [
            {{
                "name": "Dimension name",
                "type": "Type 1/2/3",
                "attributes": ["Key attributes"],
                "hierarchy": "Dimensional hierarchy if any"
            }}
        ]
    }},
    "data_layers": {{
        "bronze": "Raw data description",
        "silver": "Cleansed data description",
        "gold": "Business-level aggregates description"
    }},
    "data_lineage": [
        {{
            "source": "Source system/table",
            "transformation": "Transformation applied",
            "target": "Target table",
            "frequency": "Update frequency"
        }}
    ],
    "data_quality_rules": [
        {{
            "entity": "Table name",
            "rule_name": "Rule name",
            "rule_type": "completeness/accuracy/consistency/validity",
            "description": "What the rule checks",
            "threshold": "Acceptance threshold"
        }}
    ],
    "scd_strategy": {{
        "dimensions_with_history": ["List of Type 2 dimensions"],
        "approach": "How history is tracked"
    }}
}}

Be specific and comprehensive. Design for {selected_platform} best practices.
"""
                                    
                                    # Call Databricks Foundation Model API
                                    response = w.serving_endpoints.query(
                                        name=endpoint_name,
                                        messages=[
                                            ChatMessage(
                                                role=ChatMessageRole.SYSTEM,
                                                content="You are an expert data modeler specializing in dimensional modeling and data warehouse design. Always respond with valid JSON."
                                            ),
                                            ChatMessage(
                                                role=ChatMessageRole.USER,
                                                content=data_model_prompt
                                            )
                                        ],
                                        temperature=0.3,
                                        max_tokens=4000
                                    )
                                    
                                    # Parse response
                                    data_model_json = response.choices[0].message.content
                                    
                                    # Extract JSON from response
                                    if "```json" in data_model_json:
                                        data_model_json = data_model_json.split("```json")[1].split("```")[0].strip()
                                    elif "```" in data_model_json:
                                        data_model_json = data_model_json.split("```")[1].split("```")[0].strip()
                                    
                                    # Clean JSON - handle control characters
                                    try:
                                        data_model_design = json.loads(data_model_json, strict=False)
                                    except json.JSONDecodeError:
                                        # Try to fix common issues
                                        data_model_json_fixed = data_model_json.replace('\\n', '\\\\n').replace('\\t', '\\\\t')
                                        try:
                                            data_model_design = json.loads(data_model_json_fixed, strict=False)
                                        except:
                                            # Fallback to simplified structure
                                            st.warning("AI response had formatting issues. Using simplified data model.")
                                            data_model_design = {
                                                "modeling_approach": "Star Schema",
                                                "model_description": "Data model for " + selected_platform,
                                                "entities": [],
                                                "relationships": [],
                                                "dimensional_model": {"fact_tables": [], "dimension_tables": []},
                                                "data_layers": {"bronze": "Raw data", "silver": "Cleansed data", "gold": "Business aggregates"},
                                                "data_lineage": [],
                                                "data_quality_rules": [],
                                                "scd_strategy": {"dimensions_with_history": [], "approach": "Type 2 SCD"}
                                            }
                                    
                                    # Store in session state
                                    st.session_state.data_model_generated = True
                                    st.session_state.data_model_design = data_model_design
                                    
                                    st.success(" Data model generated successfully!")
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"Error generating data model: {str(e)}")
                                    st.info(" Tip: This requires Databricks Foundation Model APIs access.")
                    
                    # Display data model
                    if st.session_state.data_model_generated and st.session_state.data_model_design:
                        st.markdown("---")
                        st.subheader(" Data Model Design")
                        
                        dm = st.session_state.data_model_design
                        
                        # Model Overview
                        col_ov1, col_ov2 = st.columns([1, 2])
                        with col_ov1:
                            st.markdown("###  Modeling Approach")
                            st.success(dm.get('modeling_approach', 'Not specified'))
                        with col_ov2:
                            st.markdown("###  Model Description")
                            st.info(dm.get('model_description', 'Not specified'))
                        
                        # Create tabs for different model views
                        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                            " Entities",
                            "⭐ Dimensional Model",
                            " Relationships",
                            " Lineage",
                            " Quality Rules",
                            " Download"
                        ])
                        
                        with tab1:
                            st.markdown("### Entity Model")
                            
                            entities = dm.get('entities', [])
                            if entities:
                                # Filter by entity type
                                entity_types = list(set([e.get('entity_type', 'other') for e in entities]))
                                selected_type = st.selectbox("Filter by type:", ["All"] + entity_types)
                                
                                filtered_entities = entities if selected_type == "All" else [e for e in entities if e.get('entity_type') == selected_type]
                                
                                for entity in filtered_entities:
                                    with st.expander(f" {entity.get('entity_name', 'Unknown')} ({entity.get('entity_type', 'N/A')})"):
                                        st.markdown(f"**Description:** {entity.get('description', 'N/A')}")
                                        st.markdown(f"**Grain:** {entity.get('grain', 'N/A')}")
                                        
                                        st.markdown("\n**Columns:**")
                                        columns = entity.get('columns', [])
                                        if columns:
                                            # Create a table view
                                            import pandas as pd
                                            cols_df = pd.DataFrame(columns)
                                            st.dataframe(cols_df, use_container_width=True)
                                        else:
                                            st.write("No columns defined")
                                
                                st.markdown("---")
                                st.markdown(f"**Total Entities:** {len(filtered_entities)}")
                            else:
                                st.write("No entities defined")
                        
                        with tab2:
                            st.markdown("### Dimensional Model")
                            
                            dim_model = dm.get('dimensional_model', {})
                            
                            # Fact Tables
                            st.markdown("####  Fact Tables")
                            fact_tables = dim_model.get('fact_tables', [])
                            if fact_tables:
                                for fact in fact_tables:
                                    with st.expander(f" {fact.get('name', 'Unknown')}"):
                                        st.markdown(f"**Grain:** {fact.get('grain', 'N/A')}")
                                        
                                        col_f1, col_f2 = st.columns(2)
                                        with col_f1:
                                            st.markdown("**Measures:**")
                                            for measure in fact.get('measures', []):
                                                st.markdown(f"- {measure}")
                                        
                                        with col_f2:
                                            st.markdown("**Dimensions:**")
                                            for dim in fact.get('dimensions', []):
                                                st.markdown(f"- {dim}")
                            else:
                                st.write("No fact tables defined")
                            
                            st.markdown("---")
                            
                            # Dimension Tables
                            st.markdown("####  Dimension Tables")
                            dim_tables = dim_model.get('dimension_tables', [])
                            if dim_tables:
                                for dim in dim_tables:
                                    with st.expander(f" {dim.get('name', 'Unknown')} (SCD {dim.get('type', 'N/A')})"):
                                        st.markdown(f"**Type:** Slowly Changing Dimension {dim.get('type', 'N/A')}")
                                        st.markdown(f"**Hierarchy:** {dim.get('hierarchy', 'None')}")
                                        
                                        st.markdown("**Key Attributes:**")
                                        for attr in dim.get('attributes', []):
                                            st.markdown(f"- {attr}")
                            else:
                                st.write("No dimension tables defined")
                            
                            # Data Layers
                            st.markdown("---")
                            st.markdown("###  Data Layers (Medallion Architecture)")
                            
                            data_layers = dm.get('data_layers', {})
                            col_l1, col_l2, col_l3 = st.columns(3)
                            
                            with col_l1:
                                st.markdown("** Bronze Layer**")
                                st.info(data_layers.get('bronze', 'Not specified'))
                            
                            with col_l2:
                                st.markdown("** Silver Layer**")
                                st.info(data_layers.get('silver', 'Not specified'))
                            
                            with col_l3:
                                st.markdown("** Gold Layer**")
                                st.info(data_layers.get('gold', 'Not specified'))
                        
                        with tab3:
                            st.markdown("### Entity Relationships")
                            
                            relationships = dm.get('relationships', [])
                            if relationships:
                                for rel in relationships:
                                    with st.expander(f" {rel.get('from_entity', '?')} → {rel.get('to_entity', '?')} ({rel.get('relationship_type', 'N/A')})"):
                                        st.markdown(f"**From:** {rel.get('from_entity', 'N/A')}")
                                        st.markdown(f"**To:** {rel.get('to_entity', 'N/A')}")
                                        st.markdown(f"**Type:** {rel.get('relationship_type', 'N/A')}")
                                        st.markdown(f"**Foreign Key:** {rel.get('foreign_key', 'N/A')}")
                                        st.markdown(f"**Description:** {rel.get('description', 'N/A')}")
                                
                                st.markdown("---")
                                st.markdown(f"**Total Relationships:** {len(relationships)}")
                            else:
                                st.write("No relationships defined")
                            
                            # SCD Strategy
                            st.markdown("---")
                            st.markdown("###  Slowly Changing Dimension Strategy")
                            
                            scd = dm.get('scd_strategy', {})
                            st.markdown(f"**Approach:** {scd.get('approach', 'Not specified')}")
                            
                            scd_dims = scd.get('dimensions_with_history', [])
                            if scd_dims:
                                st.markdown("**Dimensions with History Tracking:**")
                                for scd_dim in scd_dims:
                                    st.markdown(f"- {scd_dim}")
                        
                        with tab4:
                            st.markdown("### Data Lineage")
                            
                            lineage = dm.get('data_lineage', [])
                            if lineage:
                                for i, line in enumerate(lineage, 1):
                                    with st.expander(f" Lineage {i}: {line.get('source', '?')} → {line.get('target', '?')}"):
                                        col_lin1, col_lin2 = st.columns(2)
                                        
                                        with col_lin1:
                                            st.markdown(f"**Source:** {line.get('source', 'N/A')}")
                                            st.markdown(f"**Target:** {line.get('target', 'N/A')}")
                                        
                                        with col_lin2:
                                            st.markdown(f"**Transformation:** {line.get('transformation', 'N/A')}")
                                            st.markdown(f"**Frequency:** {line.get('frequency', 'N/A')}")
                                
                                st.markdown("---")
                                st.info(" Data lineage shows how data flows from sources through transformations to final targets.")
                            else:
                                st.write("No lineage information available")
                        
                        with tab5:
                            st.markdown("### Data Quality Rules")
                            
                            quality_rules = dm.get('data_quality_rules', [])
                            if quality_rules:
                                # Group by entity
                                entities_with_rules = list(set([r.get('entity', 'Unknown') for r in quality_rules]))
                                
                                for entity in entities_with_rules:
                                    entity_rules = [r for r in quality_rules if r.get('entity') == entity]
                                    
                                    with st.expander(f" {entity} ({len(entity_rules)} rules)"):
                                        for rule in entity_rules:
                                            st.markdown(f"**{rule.get('rule_name', 'Unknown')}** ({rule.get('rule_type', 'N/A')})")
                                            st.markdown(f"- *Description:* {rule.get('description', 'N/A')}")
                                            st.markdown(f"- *Threshold:* {rule.get('threshold', 'N/A')}")
                                            st.markdown("")
                                
                                st.markdown("---")
                                st.markdown(f"**Total Quality Rules:** {len(quality_rules)}")
                            else:
                                st.write("No quality rules defined")
                        
                        with tab6:
                            st.markdown("### Download Data Model")
                            
                            # JSON download
                            json_dm = json.dumps(dm, indent=2)
                            st.download_button(
                                label=" Download Data Model (JSON)",
                                data=json_dm,
                                file_name=f"{st.session_state.selected_accelerator}_data_model.json",
                                mime="application/json"
                            )
                            
                            # Create comprehensive markdown documentation
                            # Pre-compute joined strings to avoid f-string backslash issues
                            entities_joined = "\n".join([f"### {e.get('entity_name', 'Unknown')} ({e.get('entity_type', 'N/A')})\n{e.get('description', '')}\n**Grain:** {e.get('grain', 'N/A')}\n" for e in dm.get('entities', [])])
                            facts_joined = "\n".join([f"- **{f.get('name', 'Unknown')}**: {f.get('grain', 'N/A')}" for f in dm.get('dimensional_model', {}).get('fact_tables', [])])
                            dims_joined = "\n".join([f"- **{d.get('name', 'Unknown')}** (SCD {d.get('type', 'N/A')})" for d in dm.get('dimensional_model', {}).get('dimension_tables', [])])
                            rels_joined = "\n".join([f"- {r.get('from_entity', '?')} → {r.get('to_entity', '?')} ({r.get('relationship_type', 'N/A')})" for r in dm.get('relationships', [])])
                            quality_joined = "\n".join([f"- **{q.get('entity', 'Unknown')}**: {q.get('rule_name', 'Unknown')} - {q.get('description', 'N/A')}" for q in dm.get('data_quality_rules', [])])
                            
                            md_dm = f"""# Data Model Design: {st.session_state.selected_accelerator}

## Modeling Approach
**{dm.get('modeling_approach', 'Not specified')}**

{dm.get('model_description', '')}

## Entities

{entities_joined}

## Dimensional Model

### Fact Tables
{facts_joined}

### Dimension Tables
{dims_joined}

## Data Layers

- **Bronze (Raw):** {dm.get('data_layers', {}).get('bronze', 'N/A')}
- **Silver (Cleansed):** {dm.get('data_layers', {}).get('silver', 'N/A')}
- **Gold (Business):** {dm.get('data_layers', {}).get('gold', 'N/A')}

## Relationships

{rels_joined}

## Data Quality Rules

{quality_joined}
"""
                            
                            st.download_button(
                                label=" Download Data Model (Markdown)",
                                data=md_dm,
                                file_name=f"{st.session_state.selected_accelerator}_data_model.md",
                                mime="text/markdown"
                            )
                            
                            st.markdown("---")
                            st.success(" Phase 6 Complete! Ready for Phase 7: Pipeline & ETL Design")
                        
                        # Progress tracker
                        st.markdown("---")
                        st.markdown("###  Progress: 60% Complete (6/10 Phases)")
                        st.progress(0.6)
                        
                        col_p1, col_p2, col_p3, col_p4, col_p5, col_p6 = st.columns(6)
                        with col_p1:
                            st.markdown(" **Phase 1**\nUpload")
                        with col_p2:
                            st.markdown(" **Phase 2**\nAnalysis")
                        with col_p3:
                            st.markdown(" **Phase 3**\nPlatform")
                        with col_p4:
                            st.markdown(" **Phase 4**\nArchitecture")
                        with col_p5:
                            st.markdown(" **Phase 5**\nArtifacts")
                        with col_p6:
                            st.markdown(" **Phase 6**\nData Model")
                
                # Phase 7: Pipeline & ETL Design
                if st.session_state.data_model_generated and st.session_state.data_model_design:
                    st.markdown("---")
                    st.markdown('<div id="phase-pipeline"></div>', unsafe_allow_html=True)
                    st.header(" Phase 7: Pipeline & ETL Design")
                    
                    st.markdown('<div class="info-box"> <b>Phase 7: Design Your Data Pipeline & ETL</b><br>AI will create a comprehensive ETL/pipeline design with orchestration, transformations, incremental loading, error handling, and monitoring.</div>', unsafe_allow_html=True)
                    
                    # Initialize session state for Phase 7
                    if 'pipeline_design_generated' not in st.session_state:
                        st.session_state.pipeline_design_generated = False
                    if 'pipeline_design' not in st.session_state:
                        st.session_state.pipeline_design = None
                    
                    # Show Phase 7 overview
                    col_pipe1, col_pipe2, col_pipe3, col_pipe4 = st.columns(4)
                    with col_pipe1:
                        st.info(" Orchestration Design")
                    with col_pipe2:
                        st.info(" Transformation Logic")
                    with col_pipe3:
                        st.info(" Incremental Loading")
                    with col_pipe4:
                        st.info(" Monitoring & Logging")
                    
                    # Generate Pipeline Design button
                    col_gen1, col_gen2, col_gen3 = st.columns([1, 2, 1])
                    with col_gen2:
                        if st.button(" Generate Pipeline & ETL Design", type="primary", use_container_width=True, disabled=st.session_state.pipeline_design_generated):
                            with st.spinner(" AI is designing your data pipeline and ETL processes... This may take 60-90 seconds"):
                                try:
                                    # Initialize Databricks SDK client
                                    w = WorkspaceClient()
                                    endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                                    
                                    # Prepare pipeline design prompt
                                    selected_platform = st.session_state.selected_accelerator
                                    data_model_summary = json.dumps(st.session_state.data_model_design, indent=2)[:3000]
                                    arch_summary = json.dumps(st.session_state.architecture_design, indent=2)[:2000]
                                    
                                    pipeline_prompt = f"""
You are an expert data engineer. Design a comprehensive data pipeline and ETL system for the following solution.

**Target Platform:** {selected_platform}

**Data Model Summary:**
{data_model_summary}

**Architecture Summary:**
{arch_summary}

Generate the following pipeline design in VALID JSON format. CRITICAL: All code in 'code' and 'script' fields must have newlines escaped as \\n and tabs as \\t:
{{
    "pipeline_overview": {{
        "description": "Overall pipeline purpose and approach",
        "architecture_pattern": "ETL/ELT/Batch/Streaming/Lambda/Kappa",
        "orchestration_tool": "Tool used for orchestration",
        "execution_model": "How the pipeline executes"
    }},
    "data_sources": [
        {{
            "source_name": "Source system name",
            "source_type": "Database/API/File/Stream/etc",
            "connection_method": "How to connect",
            "extraction_frequency": "Batch schedule or streaming",
            "extraction_strategy": "Full/Incremental/CDC"
        }}
    ],
    "orchestration_design": {{
        "workflow_structure": "DAG/Sequential/Parallel/Hybrid",
        "task_dependencies": "How tasks depend on each other",
        "scheduling": "Cron/Event/Manual/Continuous",
        "parallelism": "Degree of parallelism"
    }},
    "transformation_jobs": [
        {{
            "job_name": "Transformation job name",
            "layer": "Bronze/Silver/Gold",
            "input_sources": ["List of input tables/files"],
            "transformations": ["List of transformations applied"],
            "output_target": "Output table/location",
            "processing_type": "Batch/Streaming/Micro-batch",
            "estimated_runtime": "Expected duration"
        }}
    ],
    "incremental_loading": {{
        "strategy": "Watermark/Timestamp/CDC/Key-based",
        "tracking_method": "How progress is tracked",
        "state_management": "Where state is stored",
        "backfill_approach": "How to handle historical loads"
    }},
    "data_quality_checks": [
        {{
            "check_name": "Quality check name",
            "check_type": "Schema/Completeness/Accuracy/Freshness",
            "stage": "Pre-load/Post-load/Both",
            "action_on_failure": "Stop/Alert/Log/Quarantine"
        }}
    ],
    "error_handling": {{
        "retry_strategy": "Retry approach (exponential backoff, etc.)",
        "max_retries": "Number of retry attempts",
        "dead_letter_queue": "How failed records are handled",
        "alerting": "Alert mechanisms"
    }},
    "monitoring_logging": {{
        "metrics_tracked": ["List of key metrics"],
        "logging_strategy": "Centralized/Distributed logging approach",
        "dashboard_tools": ["Tools for monitoring dashboards"],
        "sla_requirements": "Service level agreements"
    }},
    "performance_optimization": {{
        "partitioning_strategy": "How data is partitioned",
        "clustering_keys": ["Columns used for clustering"],
        "caching_approach": "Caching strategy",
        "resource_allocation": "Compute resource sizing"
    }},
    "testing_strategy": {{
        "unit_tests": "Unit testing approach",
        "integration_tests": "Integration testing approach",
        "data_validation": "Data validation methodology",
        "regression_testing": "How to prevent regressions"
    }},
    "sample_pipeline_code": {{
        "language": "python/sql/scala",
        "framework": "Spark/dbt/Airflow/etc",
        "code": "Sample orchestration or transformation code"
    }}
}}

Be specific to {selected_platform} best practices and capabilities.
"""
                                    
                                    # Call Databricks Foundation Model API
                                    response = w.serving_endpoints.query(
                                        name=endpoint_name,
                                        messages=[
                                            ChatMessage(
                                                role=ChatMessageRole.SYSTEM,
                                                content="You are an expert data engineer specializing in ETL/ELT pipelines, data orchestration, and data processing frameworks. Always respond with valid JSON. IMPORTANT: In JSON strings, escape all newlines as \\n, tabs as \\t, and backslashes as \\\\."
                                            ),
                                            ChatMessage(
                                                role=ChatMessageRole.USER,
                                                content=pipeline_prompt
                                            )
                                        ],
                                        temperature=0.3,
                                        max_tokens=4000
                                    )
                                    
                                    # Parse response with error handling
                                    pipeline_json = response.choices[0].message.content
                                    
                                    # Extract JSON from response
                                    if "```json" in pipeline_json:
                                        pipeline_json = pipeline_json.split("```json")[1].split("```")[0].strip()
                                    elif "```" in pipeline_json:
                                        pipeline_json = pipeline_json.split("```")[1].split("```")[0].strip()
                                    
                                    # Clean JSON - handle control characters
                                    try:
                                        pipeline_design = json.loads(pipeline_json, strict=False)
                                    except json.JSONDecodeError:
                                        # Try to fix common issues
                                        pipeline_json_fixed = pipeline_json.replace('\\n', '\\\\n').replace('\\t', '\\\\t')
                                        try:
                                            pipeline_design = json.loads(pipeline_json_fixed, strict=False)
                                        except:
                                            # Fallback to simplified structure
                                            st.warning("AI response had formatting issues. Using simplified pipeline design.")
                                            pipeline_design = {
                                                "pipeline_overview": {
                                                    "description": "Data pipeline for " + selected_platform,
                                                    "architecture_pattern": "ELT",
                                                    "orchestration_tool": "Native platform scheduler",
                                                    "execution_model": "Batch processing"
                                                },
                                                "data_sources": [],
                                                "orchestration_design": {},
                                                "transformation_jobs": [],
                                                "incremental_loading": {},
                                                "data_quality_checks": [],
                                                "error_handling": {},
                                                "monitoring_logging": {},
                                                "performance_optimization": {},
                                                "testing_strategy": {},
                                                "sample_pipeline_code": {}
                                            }
                                    
                                    # Store in session state
                                    st.session_state.pipeline_design_generated = True
                                    st.session_state.pipeline_design = pipeline_design
                                    
                                    st.success(" Pipeline design generated successfully!")
                                    st.rerun()
                                    
                                except Exception as e:
                                    st.error(f"Error generating pipeline design: {str(e)}")
                                    st.info(" Tip: This requires Databricks Foundation Model APIs access.")
                    
                    # Display pipeline design
                    if st.session_state.pipeline_design_generated and st.session_state.pipeline_design:
                        st.markdown("---")
                        st.subheader("Pipeline & ETL Design")
                        
                        pd = st.session_state.pipeline_design
                        
                        # Pipeline Overview
                        st.markdown("###  Pipeline Overview")
                        overview = pd.get('pipeline_overview', {})
                        col_ov1, col_ov2 = st.columns(2)
                        with col_ov1:
                            st.markdown(f"**Architecture Pattern:** {overview.get('architecture_pattern', 'N/A')}")
                            st.markdown(f"**Orchestration Tool:** {overview.get('orchestration_tool', 'N/A')}")
                        with col_ov2:
                            st.markdown(f"**Execution Model:** {overview.get('execution_model', 'N/A')}")
                        st.info(overview.get('description', 'No description provided'))
                        
                        # Create tabs for pipeline details
                        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                            " Data Sources",
                            " Orchestration",
                            " Transformations",
                            " Incremental Loading",
                            " Quality & Monitoring",
                            " Code Sample"
                        ])
                        
                        with tab1:
                            st.markdown("### Data Sources & Extraction")
                            data_sources = pd.get('data_sources', [])
                            if data_sources:
                                for i, source in enumerate(data_sources, 1):
                                    with st.expander(f" Source {i}: {source.get('source_name', 'Unknown')}"):
                                        col_s1, col_s2 = st.columns(2)
                                        with col_s1:
                                            st.markdown(f"**Type:** {source.get('source_type', 'N/A')}")
                                            st.markdown(f"**Connection:** {source.get('connection_method', 'N/A')}")
                                        with col_s2:
                                            st.markdown(f"**Frequency:** {source.get('extraction_frequency', 'N/A')}")
                                            st.markdown(f"**Strategy:** {source.get('extraction_strategy', 'N/A')}")
                            else:
                                st.info("No data sources defined")
                        
                        with tab2:
                            st.markdown("### Orchestration Design")
                            orch = pd.get('orchestration_design', {})
                            col_o1, col_o2 = st.columns(2)
                            with col_o1:
                                st.markdown(f"**Workflow Structure:** {orch.get('workflow_structure', 'N/A')}")
                                st.markdown(f"**Scheduling:** {orch.get('scheduling', 'N/A')}")
                            with col_o2:
                                st.markdown(f"**Dependencies:** {orch.get('task_dependencies', 'N/A')}")
                                st.markdown(f"**Parallelism:** {orch.get('parallelism', 'N/A')}")
                        
                        with tab3:
                            st.markdown("### Transformation Jobs")
                            transform_jobs = pd.get('transformation_jobs', [])
                            if transform_jobs:
                                for i, job in enumerate(transform_jobs, 1):
                                    with st.expander(f" Job {i}: {job.get('job_name', 'Unknown')} - {job.get('layer', 'N/A')} Layer"):
                                        st.markdown(f"**Input Sources:** {', '.join(job.get('input_sources', []))}")
                                        st.markdown(f"**Output Target:** {job.get('output_target', 'N/A')}")
                                        st.markdown(f"**Processing Type:** {job.get('processing_type', 'N/A')}")
                                        st.markdown(f"**Estimated Runtime:** {job.get('estimated_runtime', 'N/A')}")
                                        
                                        st.markdown("**Transformations:**")
                                        for trans in job.get('transformations', []):
                                            st.markdown(f"- {trans}")
                            else:
                                st.info("No transformation jobs defined")
                        
                        with tab4:
                            st.markdown("### Incremental Loading Strategy")
                            incr = pd.get('incremental_loading', {})
                            col_i1, col_i2 = st.columns(2)
                            with col_i1:
                                st.markdown(f"**Strategy:** {incr.get('strategy', 'N/A')}")
                                st.markdown(f"**Tracking Method:** {incr.get('tracking_method', 'N/A')}")
                            with col_i2:
                                st.markdown(f"**State Management:** {incr.get('state_management', 'N/A')}")
                                st.markdown(f"**Backfill Approach:** {incr.get('backfill_approach', 'N/A')}")
                        
                        with tab5:
                            st.markdown("### Data Quality Checks")
                            quality_checks = pd.get('data_quality_checks', [])
                            if quality_checks:
                                for check in quality_checks:
                                    with st.expander(f" {check.get('check_name', 'Unknown')}"):
                                        col_q1, col_q2 = st.columns(2)
                                        with col_q1:
                                            st.markdown(f"**Type:** {check.get('check_type', 'N/A')}")
                                            st.markdown(f"**Stage:** {check.get('stage', 'N/A')}")
                                        with col_q2:
                                            st.markdown(f"**On Failure:** {check.get('action_on_failure', 'N/A')}")
                            
                            st.markdown("---")
                            st.markdown("### Error Handling")
                            err = pd.get('error_handling', {})
                            col_e1, col_e2 = st.columns(2)
                            with col_e1:
                                st.markdown(f"**Retry Strategy:** {err.get('retry_strategy', 'N/A')}")
                                st.markdown(f"**Max Retries:** {err.get('max_retries', 'N/A')}")
                            with col_e2:
                                st.markdown(f"**Dead Letter Queue:** {err.get('dead_letter_queue', 'N/A')}")
                                st.markdown(f"**Alerting:** {err.get('alerting', 'N/A')}")
                            
                            st.markdown("---")
                            st.markdown("### Monitoring & Logging")
                            mon = pd.get('monitoring_logging', {})
                            st.markdown(f"**Logging Strategy:** {mon.get('logging_strategy', 'N/A')}")
                            st.markdown(f"**SLA Requirements:** {mon.get('sla_requirements', 'N/A')}")
                            
                            st.markdown("**Metrics Tracked:**")
                            for metric in mon.get('metrics_tracked', []):
                                st.markdown(f"- {metric}")
                            
                            st.markdown("**Dashboard Tools:**")
                            for tool in mon.get('dashboard_tools', []):
                                st.markdown(f"- {tool}")
                        
                        with tab6:
                            st.markdown("### Sample Pipeline Code")
                            sample_code = pd.get('sample_pipeline_code', {})
                            st.markdown(f"**Language:** {sample_code.get('language', 'N/A')}")
                            st.markdown(f"**Framework:** {sample_code.get('framework', 'N/A')}")
                            
                            if sample_code.get('code'):
                                st.code(sample_code.get('code', ''), language=sample_code.get('language', 'python'))
                            else:
                                st.info("No sample code provided")
                            
                            st.markdown("---")
                            st.markdown("### Performance Optimization")
                            perf = pd.get('performance_optimization', {})
                            col_pf1, col_pf2 = st.columns(2)
                            with col_pf1:
                                st.markdown(f"**Partitioning:** {perf.get('partitioning_strategy', 'N/A')}")
                                st.markdown(f"**Caching:** {perf.get('caching_approach', 'N/A')}")
                            with col_pf2:
                                st.markdown(f"**Clustering Keys:** {', '.join(perf.get('clustering_keys', []))}")
                                st.markdown(f"**Resources:** {perf.get('resource_allocation', 'N/A')}")
                            
                            st.markdown("---")
                            st.markdown("### Testing Strategy")
                            test = pd.get('testing_strategy', {})
                            col_t1, col_t2 = st.columns(2)
                            with col_t1:
                                st.markdown(f"**Unit Tests:** {test.get('unit_tests', 'N/A')}")
                                st.markdown(f"**Integration Tests:** {test.get('integration_tests', 'N/A')}")
                            with col_t2:
                                st.markdown(f"**Data Validation:** {test.get('data_validation', 'N/A')}")
                                st.markdown(f"**Regression Testing:** {test.get('regression_testing', 'N/A')}")
                        
                        # Download options
                        st.markdown("---")
                        st.markdown("###  Download Pipeline Design")
                        
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            # JSON download
                            json_pipeline = json.dumps(pd, indent=2)
                            st.download_button(
                                label=" Download Pipeline Design (JSON)",
                                data=json_pipeline,
                                file_name=f"{st.session_state.selected_accelerator}_pipeline_design.json",
                                mime="application/json"
                            )
                        
                        with col_dl2:
                            # Markdown download
                            sources_text = "\n".join([f"- **{s.get('source_name', 'Unknown')}** ({s.get('source_type', 'N/A')}): {s.get('extraction_strategy', 'N/A')}" for s in pd.get('data_sources', [])])
                            jobs_text = "\n".join([f"- **{j.get('job_name', 'Unknown')}** ({j.get('layer', 'N/A')} layer): {j.get('processing_type', 'N/A')}" for j in pd.get('transformation_jobs', [])])
                            quality_text = "\n".join([f"- **{q.get('check_name', 'Unknown')}** ({q.get('check_type', 'N/A')}): {q.get('action_on_failure', 'N/A')} on failure" for q in pd.get('data_quality_checks', [])])
                            metrics_text = "\n".join([f"- {m}" for m in pd.get('monitoring_logging', {}).get('metrics_tracked', [])])
                            
                            md_pipeline = f"""# Pipeline & ETL Design: {st.session_state.selected_accelerator}

## Pipeline Overview

**Architecture Pattern:** {overview.get('architecture_pattern', 'N/A')}

**Orchestration Tool:** {overview.get('orchestration_tool', 'N/A')}

**Execution Model:** {overview.get('execution_model', 'N/A')}

{overview.get('description', '')}

## Data Sources

{sources_text}

## Orchestration

- **Workflow Structure:** {pd.get('orchestration_design', {}).get('workflow_structure', 'N/A')}
- **Scheduling:** {pd.get('orchestration_design', {}).get('scheduling', 'N/A')}
- **Task Dependencies:** {pd.get('orchestration_design', {}).get('task_dependencies', 'N/A')}
- **Parallelism:** {pd.get('orchestration_design', {}).get('parallelism', 'N/A')}

## Transformation Jobs

{jobs_text}

## Incremental Loading

- **Strategy:** {pd.get('incremental_loading', {}).get('strategy', 'N/A')}
- **Tracking Method:** {pd.get('incremental_loading', {}).get('tracking_method', 'N/A')}
- **State Management:** {pd.get('incremental_loading', {}).get('state_management', 'N/A')}
- **Backfill Approach:** {pd.get('incremental_loading', {}).get('backfill_approach', 'N/A')}

## Data Quality Checks

{quality_text}

## Error Handling

- **Retry Strategy:** {pd.get('error_handling', {}).get('retry_strategy', 'N/A')}
- **Max Retries:** {pd.get('error_handling', {}).get('max_retries', 'N/A')}
- **Dead Letter Queue:** {pd.get('error_handling', {}).get('dead_letter_queue', 'N/A')}
- **Alerting:** {pd.get('error_handling', {}).get('alerting', 'N/A')}

## Monitoring & Logging

**Logging Strategy:** {pd.get('monitoring_logging', {}).get('logging_strategy', 'N/A')}

**SLA Requirements:** {pd.get('monitoring_logging', {}).get('sla_requirements', 'N/A')}

**Metrics Tracked:**
{metrics_text}

## Performance Optimization

- **Partitioning Strategy:** {pd.get('performance_optimization', {}).get('partitioning_strategy', 'N/A')}
- **Clustering Keys:** {', '.join(pd.get('performance_optimization', {}).get('clustering_keys', []))}
- **Caching Approach:** {pd.get('performance_optimization', {}).get('caching_approach', 'N/A')}
- **Resource Allocation:** {pd.get('performance_optimization', {}).get('resource_allocation', 'N/A')}

## Testing Strategy

- **Unit Tests:** {pd.get('testing_strategy', {}).get('unit_tests', 'N/A')}
- **Integration Tests:** {pd.get('testing_strategy', {}).get('integration_tests', 'N/A')}
- **Data Validation:** {pd.get('testing_strategy', {}).get('data_validation', 'N/A')}
- **Regression Testing:** {pd.get('testing_strategy', {}).get('regression_testing', 'N/A')}

## Sample Code

**Language:** {pd.get('sample_pipeline_code', {}).get('language', 'N/A')}

**Framework:** {pd.get('sample_pipeline_code', {}).get('framework', 'N/A')}

```{pd.get('sample_pipeline_code', {}).get('language', 'python')}
{pd.get('sample_pipeline_code', {}).get('code', '')}
```
"""
                            
                            st.download_button(
                                label=" Download Pipeline Design (Markdown)",
                                data=md_pipeline,
                                file_name=f"{st.session_state.selected_accelerator}_pipeline_design.md",
                                mime="text/markdown"
                            )
                        
                        st.markdown("---")
                        st.success(" Phase 7 Complete! Ready for Phase 8: Analytics & Reporting Layer")
                        
                        # Progress tracker
                        st.markdown("---")
                        st.markdown("###  Progress: 70% Complete (7/10 Phases)")
                        st.progress(0.7)
                        
                        col_p1, col_p2, col_p3, col_p4, col_p5, col_p6, col_p7 = st.columns(7)
                        with col_p1:
                            st.markdown(" **Phase 1**\nUpload")
                        with col_p2:
                            st.markdown(" **Phase 2**\nAnalysis")
                        with col_p3:
                            st.markdown(" **Phase 3**\nPlatform")
                        with col_p4:
                            st.markdown(" **Phase 4**\nArchitecture")
                        with col_p5:
                            st.markdown(" **Phase 5**\nArtifacts")
                        with col_p6:
                            st.markdown(" **Phase 6**\nData Model")
                        with col_p7:
                            st.markdown(" **Phase 7**\nPipeline/ETL")
                
                # Phase 8: Analytics & Reporting Layer
                if st.session_state.pipeline_design_generated and st.session_state.pipeline_design:
                    st.markdown("---")
                    st.markdown('<div id="phase-analytics"></div>', unsafe_allow_html=True)
                    st.header(" Phase 8: Analytics & Reporting Layer")
                    
                    st.markdown('<div class="info-box"> <b>Phase 8: Design Analytics & Reporting</b><br>AI will create a comprehensive analytics and reporting layer design with dashboards, reports, KPIs, and visualization strategies.</div>', unsafe_allow_html=True)
                    
                    # Show Phase 8 summary
                    col_a1, col_a2, col_a3, col_a4 = st.columns(4)
                    with col_a1:
                        st.info(" Dashboard Design")
                    with col_a2:
                        st.info(" Report Specifications")
                    with col_a3:
                        st.info(" KPI Dashboards")
                    with col_a4:
                        st.info(" Visualization Strategy")
                    
                    # Generate Analytics Design button
                    col_gen1, col_gen2, col_gen3 = st.columns([1, 2, 1])
                    with col_gen2:
                        if st.button(" Generate Analytics Layer", type="primary", use_container_width=True, disabled=st.session_state.analytics_design_generated):
                            with st.spinner(" AI is designing your analytics and reporting layer... This may take 60-90 seconds"):
                                try:
                                    w = WorkspaceClient()
                                    endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                                    
                                    # Prepare analytics design prompt
                                    selected_platform = st.session_state.selected_accelerator
                                    analytics_prompt = f"""You are an expert analytics architect. Design a comprehensive analytics and reporting layer.

Target Platform: {selected_platform}

Generate analytics design in JSON:
{{
    "dashboards": [{{
        "name": "Dashboard name",
        "purpose": "What it shows",
        "audience": "Target users",
        "refresh_frequency": "Update schedule",
        "widgets": ["List of visualizations"]
    }}],
    "reports": [{{
        "name": "Report name",
        "type": "Operational/Analytical/Executive",
        "schedule": "When it runs",
        "outputs": ["PDF/Excel/Email"]
    }}],
    "kpi_metrics": [{{
        "kpi_name": "KPI name",
        "calculation": "How it's calculated",
        "target_value": "Goal",
        "visualization_type": "Chart type"
    }}],
    "visualization_strategy": {{
        "tools": ["BI tools used"],
        "best_practices": ["Design principles"]
    }},
    "access_control": {{
        "row_level_security": "RLS approach",
        "column_level_security": "CLS approach"
    }}
}}

Be specific to {selected_platform}.
"""
                                    
                                    response = w.serving_endpoints.query(
                                        name=endpoint_name,
                                        messages=[
                                            ChatMessage(role=ChatMessageRole.SYSTEM, content="You are an expert analytics architect. Always respond with valid JSON."),
                                            ChatMessage(role=ChatMessageRole.USER, content=analytics_prompt)
                                        ],
                                        temperature=0.3,
                                        max_tokens=3000
                                    )
                                    
                                    analytics_json = response.choices[0].message.content
                                    if "```json" in analytics_json:
                                        analytics_json = analytics_json.split("```json")[1].split("```")[0].strip()
                                    elif "```" in analytics_json:
                                        analytics_json = analytics_json.split("```")[1].split("```")[0].strip()
                                    
                                    try:
                                        analytics_design = json.loads(analytics_json, strict=False)
                                    except:
                                        analytics_design = {"dashboards": [], "reports": [], "kpi_metrics": [], "visualization_strategy": {}, "access_control": {}}
                                    
                                    st.session_state.analytics_design_generated = True
                                    st.session_state.analytics_design = analytics_design
                                    st.success(" Analytics layer design complete!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error generating analytics design: {str(e)}")
                    
                    # Display analytics design
                    if st.session_state.analytics_design_generated and st.session_state.analytics_design:
                        st.markdown("---")
                        st.subheader("Analytics & Reporting Design")
                        ad = st.session_state.analytics_design
                        
                        tab1, tab2, tab3, tab4 = st.tabs([" Dashboards", " Reports", " KPIs", " Download"])
                        
                        with tab1:
                            st.markdown("### Dashboards")
                            for dash in ad.get('dashboards', []):
                                with st.expander(f" {dash.get('name', 'Unknown')}"):
                                    st.markdown(f"**Purpose:** {dash.get('purpose', 'N/A')}")
                                    st.markdown(f"**Audience:** {dash.get('audience', 'N/A')}")
                                    st.markdown(f"**Refresh:** {dash.get('refresh_frequency', 'N/A')}")
                                    st.markdown("**Widgets:**")
                                    for widget in dash.get('widgets', []):
                                        st.markdown(f"- {widget}")
                        
                        with tab2:
                            st.markdown("### Reports")
                            for rep in ad.get('reports', []):
                                with st.expander(f" {rep.get('name', 'Unknown')}"):
                                    st.markdown(f"**Type:** {rep.get('type', 'N/A')}")
                                    st.markdown(f"**Schedule:** {rep.get('schedule', 'N/A')}")
                                    st.markdown(f"**Outputs:** {', '.join(rep.get('outputs', []))}")
                        
                        with tab3:
                            st.markdown("### KPI Metrics")
                            for kpi in ad.get('kpi_metrics', []):
                                with st.expander(f" {kpi.get('kpi_name', 'Unknown')}"):
                                    st.markdown(f"**Calculation:** {kpi.get('calculation', 'N/A')}")
                                    st.markdown(f"**Target:** {kpi.get('target_value', 'N/A')}")
                                    st.markdown(f"**Visualization:** {kpi.get('visualization_type', 'N/A')}")
                        
                        with tab4:
                            json_analytics = json.dumps(ad, indent=2)
                            st.download_button(" Download Analytics Design (JSON)", json_analytics, f"{selected_platform}_analytics.json", "application/json")
                        
                        st.success(" Phase 8 Complete! Ready for Phase 9: Testing Framework")
                        st.progress(0.8)
                
                # Phase 9: Testing & Validation Framework
                if st.session_state.analytics_design_generated and st.session_state.analytics_design:
                    st.markdown("---")
                    st.markdown('<div id="phase-testing"></div>', unsafe_allow_html=True)
                    st.header(" Phase 9: Testing & Validation Framework")
                    
                    st.markdown('<div class="info-box"> <b>Phase 9: Design Testing Framework</b><br>AI will create a comprehensive testing strategy with unit tests, integration tests, data validation, and UAT plans.</div>', unsafe_allow_html=True)
                    
                    # Show Phase 9 summary
                    col_t1, col_t2, col_t3, col_t4 = st.columns(4)
                    with col_t1:
                        st.info(" Unit Tests")
                    with col_t2:
                        st.info(" Integration Tests")
                    with col_t3:
                        st.info(" Data Validation")
                    with col_t4:
                        st.info(" UAT Plans")
                    
                    # Generate Testing Framework button
                    col_gen1, col_gen2, col_gen3 = st.columns([1, 2, 1])
                    with col_gen2:
                        if st.button(" Generate Testing Framework", type="primary", use_container_width=True, disabled=st.session_state.testing_framework_generated):
                            with st.spinner(" AI is designing your testing framework... This may take 60-90 seconds"):
                                try:
                                    w = WorkspaceClient()
                                    endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                                    
                                    selected_platform = st.session_state.selected_accelerator
                                    testing_prompt = f"""You are a QA expert. Design a comprehensive testing framework.

Target Platform: {selected_platform}

Generate testing framework in JSON:
{{
    "unit_tests": [{{
        "component": "Component name",
        "test_cases": ["List of test cases"],
        "framework": "Testing tool"
    }}],
    "integration_tests": [{{
        "integration_point": "What's being tested",
        "test_scenarios": ["Scenarios"],
        "expected_outcomes": ["Expected results"]
    }}],
    "data_validation": [{{
        "validation_name": "Validation name",
        "validation_type": "Schema/Quality/Completeness",
        "criteria": "Pass criteria"
    }}],
    "performance_tests": [{{
        "test_name": "Performance test",
        "metrics": ["Metrics tracked"],
        "threshold": "Acceptable threshold"
    }}],
    "uat_plan": {{
        "test_scenarios": ["User scenarios"],
        "acceptance_criteria": ["Criteria"],
        "stakeholders": ["Who participates"]
    }},
    "test_data_strategy": {{
        "approach": "How test data is managed",
        "data_masking": "PII protection approach"
    }}
}}
"""
                                    
                                    response = w.serving_endpoints.query(
                                        name=endpoint_name,
                                        messages=[
                                            ChatMessage(role=ChatMessageRole.SYSTEM, content="You are a QA expert. Always respond with valid JSON."),
                                            ChatMessage(role=ChatMessageRole.USER, content=testing_prompt)
                                        ],
                                        temperature=0.3,
                                        max_tokens=3000
                                    )
                                    
                                    testing_json = response.choices[0].message.content
                                    if "```json" in testing_json:
                                        testing_json = testing_json.split("```json")[1].split("```")[0].strip()
                                    elif "```" in testing_json:
                                        testing_json = testing_json.split("```")[1].split("```")[0].strip()
                                    
                                    try:
                                        testing_framework = json.loads(testing_json, strict=False)
                                    except:
                                        testing_framework = {"unit_tests": [], "integration_tests": [], "data_validation": [], "performance_tests": [], "uat_plan": {}, "test_data_strategy": {}}
                                    
                                    st.session_state.testing_framework_generated = True
                                    st.session_state.testing_framework = testing_framework
                                    st.success(" Testing framework complete!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error generating testing framework: {str(e)}")
                    
                    # Display testing framework
                    if st.session_state.testing_framework_generated and st.session_state.testing_framework:
                        st.markdown("---")
                        st.subheader("Testing & Validation Framework")
                        tf = st.session_state.testing_framework
                        
                        tab1, tab2, tab3, tab4 = st.tabs([" Unit Tests", " Integration", " Data Validation", " Download"])
                        
                        with tab1:
                            st.markdown("### Unit Tests")
                            for ut in tf.get('unit_tests', []):
                                with st.expander(f" {ut.get('component', 'Unknown')}"):
                                    st.markdown(f"**Framework:** {ut.get('framework', 'N/A')}")
                                    st.markdown("**Test Cases:**")
                                    for tc in ut.get('test_cases', []):
                                        st.markdown(f"- {tc}")
                        
                        with tab2:
                            st.markdown("### Integration Tests")
                            for it in tf.get('integration_tests', []):
                                with st.expander(f" {it.get('integration_point', 'Unknown')}"):
                                    st.markdown("**Test Scenarios:**")
                                    for ts in it.get('test_scenarios', []):
                                        st.markdown(f"- {ts}")
                        
                        with tab3:
                            st.markdown("### Data Validation")
                            for dv in tf.get('data_validation', []):
                                with st.expander(f" {dv.get('validation_name', 'Unknown')}"):
                                    st.markdown(f"**Type:** {dv.get('validation_type', 'N/A')}")
                                    st.markdown(f"**Criteria:** {dv.get('criteria', 'N/A')}")
                        
                        with tab4:
                            json_testing = json.dumps(tf, indent=2)
                            st.download_button(" Download Testing Framework (JSON)", json_testing, f"{selected_platform}_testing.json", "application/json")
                        
                        st.success(" Phase 9 Complete! Ready for Phase 10: Deployment & Documentation")
                        st.progress(0.9)
                
                # Phase 10: Deployment & Documentation
                if st.session_state.testing_framework_generated and st.session_state.testing_framework:
                    st.markdown("---")
                    st.markdown('<div id="phase-deployment"></div>', unsafe_allow_html=True)
                    st.header(" Phase 10: Deployment & Documentation")
                    
                    st.markdown('<div class="info-box"> <b>Phase 10: Deployment Automation & Final Documentation</b><br>AI will create deployment automation, runbooks, and comprehensive project documentation.</div>', unsafe_allow_html=True)
                    
                    # Show Phase 10 summary
                    col_d1, col_d2, col_d3, col_d4 = st.columns(4)
                    with col_d1:
                        st.info(" CI/CD Pipeline")
                    with col_d2:
                        st.info(" Documentation")
                    with col_d3:
                        st.info(" Runbooks")
                    with col_d4:
                        st.info(" Training Materials")
                    
                    # Generate Deployment Plan button
                    col_gen1, col_gen2, col_gen3 = st.columns([1, 2, 1])
                    with col_gen2:
                        if st.button(" Generate Deployment Plan", type="primary", use_container_width=True, disabled=st.session_state.deployment_plan_generated):
                            with st.spinner(" AI is creating your deployment plan and documentation... This may take 60-90 seconds"):
                                try:
                                    w = WorkspaceClient()
                                    endpoint_name = "databricks-meta-llama-3-3-70b-instruct"
                                    
                                    selected_platform = st.session_state.selected_accelerator
                                    deployment_prompt = f"""You are a DevOps expert. Design a complete deployment and documentation strategy.

Target Platform: {selected_platform}

Generate deployment plan in JSON:
{{
    "cicd_pipeline": {{
        "tool": "CI/CD tool used",
        "stages": ["Build/Test/Deploy stages"],
        "automation_level": "Fully automated/Semi-automated"
    }},
    "deployment_strategy": {{
        "approach": "Blue-Green/Canary/Rolling",
        "rollback_plan": "How to rollback",
        "zero_downtime": "Strategy for zero downtime"
    }},
    "infrastructure_as_code": [{{
        "tool": "Terraform/CloudFormation/ARM",
        "components": ["What's automated"],
        "environments": ["Dev/Staging/Prod"]
    }}],
    "runbooks": [{{
        "title": "Runbook name",
        "purpose": "When to use it",
        "steps": ["Step-by-step procedures"]
    }}],
    "documentation": [{{
        "doc_type": "Technical/User/Admin",
        "sections": ["Document sections"],
        "audience": "Target readers"
    }}],
    "training_materials": [{{
        "material_type": "Videos/Guides/Workshops",
        "target_audience": "Who it's for",
        "topics": ["Topics covered"]
    }}],
    "monitoring_post_deployment": {{
        "health_checks": ["What to monitor"],
        "alerting_rules": ["When to alert"],
        "sla_tracking": "How SLAs are tracked"
    }}
}}
"""
                                    
                                    response = w.serving_endpoints.query(
                                        name=endpoint_name,
                                        messages=[
                                            ChatMessage(role=ChatMessageRole.SYSTEM, content="You are a DevOps expert. Always respond with valid JSON."),
                                            ChatMessage(role=ChatMessageRole.USER, content=deployment_prompt)
                                        ],
                                        temperature=0.3,
                                        max_tokens=3000
                                    )
                                    
                                    deployment_json = response.choices[0].message.content
                                    if "```json" in deployment_json:
                                        deployment_json = deployment_json.split("```json")[1].split("```")[0].strip()
                                    elif "```" in deployment_json:
                                        deployment_json = deployment_json.split("```")[1].split("```")[0].strip()
                                    
                                    try:
                                        deployment_plan = json.loads(deployment_json, strict=False)
                                    except:
                                        deployment_plan = {"cicd_pipeline": {}, "deployment_strategy": {}, "infrastructure_as_code": [], "runbooks": [], "documentation": [], "training_materials": [], "monitoring_post_deployment": {}}
                                    
                                    st.session_state.deployment_plan_generated = True
                                    st.session_state.deployment_plan = deployment_plan
                                    st.success(" Deployment plan complete!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error generating deployment plan: {str(e)}")
                    
                    # Display deployment plan
                    if st.session_state.deployment_plan_generated and st.session_state.deployment_plan:
                        st.markdown("---")
                        st.subheader("Deployment & Documentation Plan")
                        dp = st.session_state.deployment_plan
                        
                        tab1, tab2, tab3, tab4 = st.tabs([" CI/CD", " Documentation", " Runbooks", " Download"])
                        
                        with tab1:
                            st.markdown("### CI/CD Pipeline")
                            cicd = dp.get('cicd_pipeline', {})
                            st.markdown(f"**Tool:** {cicd.get('tool', 'N/A')}")
                            st.markdown(f"**Automation Level:** {cicd.get('automation_level', 'N/A')}")
                            st.markdown("**Stages:**")
                            for stage in cicd.get('stages', []):
                                st.markdown(f"- {stage}")
                            
                            st.markdown("---")
                            st.markdown("### Deployment Strategy")
                            ds = dp.get('deployment_strategy', {})
                            st.markdown(f"**Approach:** {ds.get('approach', 'N/A')}")
                            st.markdown(f"**Rollback Plan:** {ds.get('rollback_plan', 'N/A')}")
                            st.markdown(f"**Zero Downtime:** {ds.get('zero_downtime', 'N/A')}")
                        
                        with tab2:
                            st.markdown("### Documentation")
                            for doc in dp.get('documentation', []):
                                with st.expander(f" {doc.get('doc_type', 'Unknown')} Documentation"):
                                    st.markdown(f"**Audience:** {doc.get('audience', 'N/A')}")
                                    st.markdown("**Sections:**")
                                    for sec in doc.get('sections', []):
                                        st.markdown(f"- {sec}")
                        
                        with tab3:
                            st.markdown("### Runbooks")
                            for rb in dp.get('runbooks', []):
                                with st.expander(f" {rb.get('title', 'Unknown')}"):
                                    st.markdown(f"**Purpose:** {rb.get('purpose', 'N/A')}")
                                    st.markdown("**Steps:**")
                                    for step in rb.get('steps', []):
                                        st.markdown(f"- {step}")
                        
                        with tab4:
                            json_deployment = json.dumps(dp, indent=2)
                            st.download_button(" Download Deployment Plan (JSON)", json_deployment, f"{selected_platform}_deployment.json", "application/json")
                        
                        st.markdown("---")
                        st.success(" Phase 10 Complete! ALL PHASES FINISHED!")
                        st.balloons()
                        st.progress(1.0)
                        
                        st.markdown("###    PROJECT COMPLETE!   ")
                        st.markdown("### All 10 Phases Successfully Generated!")
                        
                        col_final1, col_final2, col_final3, col_final4, col_final5 = st.columns(5)
                        with col_final1:
                            st.markdown(" **Phases 1-2**\nBRD & Analysis")
                        with col_final2:
                            st.markdown(" **Phases 3-4**\nPlatform & Architecture")
                        with col_final3:
                            st.markdown(" **Phases 5-6**\nArtifacts & Data Model")
                        with col_final4:
                            st.markdown(" **Phases 7-8**\nPipeline & Analytics")
                        with col_final5:
                            st.markdown(" **Phases 9-10**\nTesting & Deployment")
            
        except Exception as e:
            st.error(f"Error extracting text: {str(e)}")
            st.exception(e)
else:
    # Placeholder when no file is uploaded
    st.info(" Upload a document to get started")
    
    # Sample BRD structure hint
    with st.expander(" What should a BRD include?"):
        st.markdown("""
        A good Business Requirements Document typically includes:
        
        * **Executive Summary**: Overview of the project and objectives
        * **Business Objectives**: Goals and expected outcomes
        * **Scope**: What's included and excluded
        * **Stakeholders**: Key people involved
        * **Functional Requirements**: What the solution must do
        * **Data Requirements**: Data sources, entities, KPIs
        * **Technical Requirements**: Integration, performance, security
        * **Business Rules**: Calculations, validations, workflows
        * **Constraints**: Timeline, budget, technical limitations
        * **Success Criteria**: How success will be measured
        """)

# Professional Footer
st.markdown("""<br><br>""", unsafe_allow_html=True)

st.markdown("""
<div class="footer">
    <div style="text-align: center;">
        <h3 style="color: #000000; margin-bottom: 1rem; font-size: 1.5rem;"> BizForge</h3>
        <p style="font-size: 1.1rem; color: #000000; margin-bottom: 0.5rem;">
            <strong>Enterprise AI Solution Engineering Platform</strong>
        </p>
        <p style="font-size: 0.95rem; color: #000000; margin-bottom: 1.5rem;">
            Accelerating the journey from requirements to reality
        </p>
    </div>
</div>
""", unsafe_allow_html=True)




