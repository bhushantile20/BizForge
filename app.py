import streamlit as st
import pypdf
from docx import Document
import io
import json
import os
from openai import OpenAI

# Page configuration
st.set_page_config(
    page_title="BizForge - AI Solution Engineering Platform",
    page_icon="🏗️",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        background: linear-gradient(90deg, #1E88E5 0%, #00ACC1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    .subtitle {
        font-size: 1.2rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .phase-badge {
        background-color: #E3F2FD;
        color: #1976D2;
        padding: 0.3rem 0.8rem;
        border-radius: 15px;
        font-size: 0.9rem;
        font-weight: 600;
    }
    .success-box {
        padding: 1rem;
        background-color: #E8F5E9;
        border-left: 4px solid #4CAF50;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #E3F2FD;
        border-left: 4px solid #2196F3;
        border-radius: 4px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">🏗️ BizForge</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">AI-Powered Solution Engineering Platform</div>', unsafe_allow_html=True)
st.markdown('<span class="phase-badge">Phase 1 & 2: Document Upload & AI Analysis</span>', unsafe_allow_html=True)

st.markdown("---")

# Introduction
with st.expander("ℹ️ About BizForge", expanded=False):
    st.markdown("""
    **BizForge** converts business requirements into build-ready technical solutions.
    
    **How it works:**
    1. 📄 Upload your Business Requirements Document (BRD)
    2. 🤖 AI analyzes and extracts requirements
    3. 🎯 Select your target accelerator (Databricks, Fabric, AWS, etc.)
    4. 🏗️ Generate solution architecture and build artifacts
    
    **Current Phase:** Document Upload & AI-Powered Requirements Analysis
    
    **Coming Soon:** Accelerator selection and artifact generation
    """)

# Main content
st.header("📄 Upload Your Business Requirements Document")

st.markdown("""
Upload your BRD in any of the following formats:
* **PDF** (.pdf)
* **Word Document** (.docx)
* **Text File** (.txt)
""")

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
    with st.spinner("🔍 Extracting text from document..."):
        try:
            extracted_text = ""
            
            if uploaded_file.type == "application/pdf":
                # Extract text from PDF
                pdf_reader = pypdf.PdfReader(io.BytesIO(uploaded_file.read()))
                extracted_text = ""
                for page_num, page in enumerate(pdf_reader.pages, 1):
                    page_text = page.extract_text()
                    extracted_text += f"\n--- Page {page_num} ---\n{page_text}"
                
                st.markdown('<div class="success-box">✅ Successfully extracted text from PDF</div>', unsafe_allow_html=True)
                st.info(f"📊 Document Statistics: {len(pdf_reader.pages)} pages, {len(extracted_text)} characters")
            
            elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                # Extract text from DOCX
                doc = Document(io.BytesIO(uploaded_file.read()))
                extracted_text = "\n\n".join([para.text for para in doc.paragraphs if para.text.strip()])
                
                st.markdown('<div class="success-box">✅ Successfully extracted text from Word document</div>', unsafe_allow_html=True)
                st.info(f"📊 Document Statistics: {len(doc.paragraphs)} paragraphs, {len(extracted_text)} characters")
            
            elif uploaded_file.type == "text/plain":
                # Extract text from TXT
                extracted_text = uploaded_file.read().decode("utf-8")
                
                st.markdown('<div class="success-box">✅ Successfully extracted text from text file</div>', unsafe_allow_html=True)
                st.info(f"📊 Document Statistics: {len(extracted_text.splitlines())} lines, {len(extracted_text)} characters")
            
            # Display extracted text
            st.header("📝 Extracted Document Content")
            
            # Create tabs for different views
            tab1, tab2 = st.tabs(["📄 Full Text", "📊 Preview"])
            
            with tab1:
                st.text_area(
                    "Extracted Text",
                    value=extracted_text,
                    height=500,
                    help="Full extracted text from your document"
                )
                
                # Download button
                st.download_button(
                    label="💾 Download Extracted Text",
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
            st.header("🤖 AI-Powered Requirements Analysis")
            
            st.markdown('<div class="info-box">📊 <b>Phase 2: AI Analysis</b><br>Click the button below to analyze your BRD using AI to extract structured requirements, data sources, KPIs, and business rules.</div>', unsafe_allow_html=True)
            
            # Initialize session state for analysis
            if 'analysis_complete' not in st.session_state:
                st.session_state.analysis_complete = False
            if 'analysis_results' not in st.session_state:
                st.session_state.analysis_results = None
            
            # AI Analysis button
            col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
            with col_btn2:
                analyze_button = st.button(
                    "🚀 Analyze BRD with AI",
                    type="primary",
                    use_container_width=True
                )
            
            if analyze_button:
                with st.spinner("🧠 AI is analyzing your document... This may take 30-60 seconds"):
                    try:
                        # Initialize OpenAI client
                        # Try to get API key from environment or Databricks secrets
                        api_key = os.environ.get("OPENAI_API_KEY")
                        
                        if not api_key:
                            # Provide instructions if no API key
                            st.error("⚠️ OpenAI API key not configured. Please set OPENAI_API_KEY environment variable.")
                            with st.expander("💡 How to configure API key"):
                                st.markdown("""
                                **Option 1: Use Databricks Secrets (Recommended)**
                                1. Store your OpenAI API key in Databricks secrets
                                2. Update app.yaml to inject the secret as environment variable
                                
                                **Option 2: Temporary Testing**
                                - Enter your API key below (not recommended for production)
                                """)
                                
                                temp_api_key = st.text_input("Enter OpenAI API Key", type="password")
                                if temp_api_key:
                                    api_key = temp_api_key
                        
                        if api_key:
                            client = OpenAI(api_key=api_key)
                            
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
                            
                            # Call OpenAI API
                            response = client.chat.completions.create(
                                model="gpt-4",
                                messages=[
                                    {"role": "system", "content": "You are a helpful business analyst that extracts structured information from documents."},
                                    {"role": "user", "content": analysis_prompt}
                                ],
                                temperature=0.3,
                                response_format={"type": "json_object"}
                            )
                            
                            # Parse response
                            analysis_json = response.choices[0].message.content
                            analysis_results = json.loads(analysis_json)
                            
                            # Store in session state
                            st.session_state.analysis_complete = True
                            st.session_state.analysis_results = analysis_results
                            
                            st.success("✅ Analysis complete!")
                            st.rerun()
                            
                    except Exception as e:
                        st.error(f"❌ Error during AI analysis: {str(e)}")
                        st.info("💡 Tip: Make sure your OpenAI API key is valid and has sufficient credits.")
            
            # Display analysis results if available
            if st.session_state.analysis_complete and st.session_state.analysis_results:
                st.markdown("---")
                st.subheader("📊 Analysis Results")
                
                results = st.session_state.analysis_results
                
                # Create tabs for different sections
                tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                    "📋 Overview",
                    "📊 Data & KPIs", 
                    "⚙️ Business Rules",
                    "💻 Technical",
                    "👥 Stakeholders",
                    "📥 Download"
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
                            with st.expander(f"📈 {kpi.get('name', 'Unnamed KPI')}"):
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
                        label="📥 Download as JSON",
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
                        label="📄 Download as Markdown",
                        data=md_summary,
                        file_name=f"{uploaded_file.name}_analysis.md",
                        mime="text/markdown"
                    )
                    
                    st.markdown("---")
                    st.success("✅ Phase 2 Complete! Ready for Phase 3: Accelerator Selection")
            
        except Exception as e:
            st.error(f"❌ Error extracting text: {str(e)}")
            st.exception(e)
else:
    # Placeholder when no file is uploaded
    st.info("👆 Upload a document to get started")
    
    # Sample BRD structure hint
    with st.expander("💡 What should a BRD include?"):
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

# Footer
st.markdown("---")
st.markdown("**BizForge** - Accelerating the journey from requirements to reality | Phase 2 of 10")

