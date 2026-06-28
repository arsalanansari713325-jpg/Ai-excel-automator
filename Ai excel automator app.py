import streamlit as st
import pandas as pd
import numpy as np
import io

# ==========================================
# 1. PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="AI Excel Automator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Tailwind-like CSS styling for Streamlit elements
st.markdown("""
<style>
    /* Main container styling */
    .main { background-color: #f8fafc; }
    
    /* KPI Card styling */
    .kpi-card {
        background: white;
        padding: 1.25rem;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        text-align: center;
    }
    .kpi-title { font-size: 0.85rem; color: #64748b; font-weight: 600; text-transform: uppercase; margin-bottom: 0.25rem; }
    .kpi-val { font-size: 1.8rem; font-weight: 800; color: #0f172a; }
    
    /* Header Gradient banner */
    .hero-banner {
        background: linear-gradient(135deg, #065f46 0%, #0f766e 50%, #0f172a 100%);
        color: white;
        padding: 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
    }
    .hero-banner h1 { color: white !important; margin: 0; font-size: 2.2rem; font-weight: 900; }
    .hero-banner p { color: #a7f3d0; margin-top: 0.5rem; font-size: 1.05rem; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. APP HEADER
# ==========================================
st.markdown("""
<div class="hero-banner">
    <h1>⚡ AI Excel Automator & Database Transformer</h1>
    <p>Upload messy Excel / CSV sheets to instantly clean data, generate formulas, calculate KPIs, and export reports.</p>
</div>
""", unsafe_allow_html=True)

# Sidebar Navigation & Settings
with st.sidebar:
    st.header("🎛️ Workspace Settings")
    st.caption("Select what you want to do with your spreadsheet today:")
    
    app_mode = st.radio(
        "Navigation Mode",
        ["📁 Upload & Transform Sheet", "🧙‍♂️ Excel Formula Solver"],
        index=0
    )
    
    st.divider()
    st.info("💡 **Pro Tip:** Keep your column names clean (e.g., 'Revenue', 'Cost', 'Date') for auto-detecting financial KPIs!")

# ==========================================
# 3. MODE 1: SHEET UPLOAD & DATABASE TRANSFORMER
# ==========================================
if app_mode == "📁 Upload & Transform Sheet":
    
    # File Uploader Widget
    uploaded_file = st.file_uploader(
        "📤 Upload your Excel (.xlsx, .xls) or CSV file here",
        type=["xlsx", "xls", "csv"],
        help="Maximum recommended file size is 200MB."
    )
    
    if uploaded_file is not None:
        try:
            # Load file based on extension
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                excel_file = pd.ExcelFile(uploaded_file)
                sheet_names = excel_file.sheet_names
                
                # If multiple sheets, let user choose
                selected_sheet = sheet_names[0]
                if len(sheet_names) > 1:
                    selected_sheet = st.selectbox("📑 Select Excel Worksheet", sheet_names)
                df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
                
            st.success(f"Successfully loaded **{uploaded_file.name}** ({df.shape[0]} rows, {df.shape[1]} columns)")
            
            # ------------------------------------------
            # TABS FOR DATA ACTIONS
            # ------------------------------------------
            tab1, tab2, tab3, tab4 = st.tabs([
                "📊 Instant Database View", 
                "📈 Automated KPI Summary", 
                "🧹 Data Cleaner & Filter", 
                "💾 Export & Download"
            ])
            
            # TAB 1: RAW DATA TABLE
            with tab1:
                st.subheader("Interactive Database")
                st.dataframe(df, use_container_width=True, height=450)
                
                with st.expander("🔍 View Column Data Types & Missing Nulls Count"):
                    col_info = pd.DataFrame({
                        "Data Type": df.dtypes.astype(str),
                        "Missing (Null) Values": df.isnull().sum(),
                        "Non-Null Count": df.notnull().sum()
                    })
                    st.table(col_info)

            # TAB 2: AUTOMATED KPI CALCULATOR
            with tab2:
                st.subheader("⚡ Instant Executive KPI Dashboard")
                
                # Auto-detect numeric columns
                numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
                
                if not numeric_cols:
                    st.warning("No numeric columns found in this sheet to calculate KPIs.")
                else:
                    # Allow user to pick primary metrics
                    col_select1, col_select2 = st.columns(2)
                    with col_select1:
                        target_col = st.selectbox("Select Primary Metric Column (Sum/Avg)", numeric_cols, index=0)
                    with col_select2:
                        group_col = st.selectbox("Group By Column (Optional)", ["None"] + df.columns.tolist(), index=0)
                    
                    # Display Top KPI Banner
                    k1, k2, k3, k4 = st.columns(4)
                    total_sum = df[target_col].sum()
                    avg_val = df[target_col].mean()
                    max_val = df[target_col].max()
                    min_val = df[target_col].min()
                    
                    with k1:
                        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total {target_col}</div><div class="kpi-val">{total_sum:,.2f}</div></div>', unsafe_allow_html=True)
                    with k2:
                        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Average {target_col}</div><div class="kpi-val">{avg_val:,.2f}</div></div>', unsafe_allow_html=True)
                    with k3:
                        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Maximum Value</div><div class="kpi-val">{max_val:,.2f}</div></div>', unsafe_allow_html=True)
                    with k4:
                        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Minimum Value</div><div class="kpi-val">{min_val:,.2f}</div></div>', unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # Charting section
                    if group_col != "None" and group_col != target_col:
                        st.markdown(f"#### 📈 {target_col} Breakdown by {group_col}")
                        grouped_df = df.groupby(group_col)[target_col].sum().reset_index()
                        st.bar_chart(grouped_df, x=group_col, y=target_col, use_container_width=True)
                    else:
                        st.markdown(f"#### 📈 Trend Line for {target_col}")
                        st.line_chart(df[target_col], use_container_width=True)

            # TAB 3: DATA CLEANER & TRANSFORMER
            with tab3:
                st.subheader("🧹 1-Click Data Cleaning Tools")
                
                clean_df = df.copy()
                
                c1, c2, c3 = st.columns(3)
                with c1:
                    remove_dups = st.checkbox("Remove Duplicate Rows", value=False)
                with c2:
                    drop_nulls = st.checkbox("Drop Rows with Null/Missing Values", value=False)
                with c3:
                    strip_spaces = st.checkbox("Trim Extra Spaces from Text", value=True)
                
                if remove_dups:
                    before_rows = len(clean_df)
                    clean_df = clean_df.drop_duplicates()
                    st.toast(f"Removed {before_rows - len(clean_df)} duplicate rows.")
                    
                if drop_nulls:
                    clean_df = clean_df.dropna()
                    
                if strip_spaces:
                    str_cols = clean_df.select_dtypes(include=['object']).columns
                    for col in str_cols:
                        clean_df[col] = clean_df[col].astype(str).str.strip()
                
                st.markdown("#### Preview Cleaned Database")
                st.dataframe(clean_df, use_container_width=True, height=300)

            # TAB 4: EXPORT REPORTS
            with tab4:
                st.subheader("📥 Export Processed Reports")
                st.write("Download your cleaned dataset or summary KPIs in your preferred format.")
                
                export_data = clean_df if 'clean_df' in locals() else df
                
                exp_col1, exp_col2 = st.columns(2)
                
                # CSV Export
                csv_buffer = export_data.to_csv(index=False).encode('utf-8')
                with exp_col1:
                    st.download_button(
                        label="📄 Download Cleaned Data (CSV)",
                        data=csv_buffer,
                        file_name="AI_Cleaned_Database.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                
                # Excel Export using BytesIO buffer
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    export_data.to_excel(writer, index=False, sheet_name='Cleaned Data')
                    
                    # Add summary sheet if numeric columns exist
                    if numeric_cols:
                        summary_df = export_data[numeric_cols].describe().reset_index()
                        summary_df.to_excel(writer, index=False, sheet_name='KPI Summary')
                        
                with exp_col2:
                    st.download_button(
                        label="📊 Download Multi-Sheet Excel (.xlsx)",
                        data=excel_buffer.getvalue(),
                        file_name="AI_Automated_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error(f"❌ Error processing spreadsheet: {str(e)}")
            st.info("Make sure the uploaded file is not corrupted or password-protected.")

# ==========================================
# 4. MODE 2: EXCEL FORMULA SOLVER & GENERATOR
# ==========================================
elif app_mode == "🧙‍♂️ Excel Formula Solver":
    
    st.subheader("🧙‍♂️ AI Excel Formula Generator & Explainer")
    st.write("Type what you want to calculate in plain English, and get exact Excel/Google Sheets syntax instantly.")
    
    col_input, col_examples = st.columns([2, 1])
    
    with col_examples:
        st.markdown("#### 💡 Quick Examples")
        if st.button("Calculate VLOOKUP across sheets"):
            st.session_state.prompt_val = "Lookup employee ID in cell A2 from sheet 'Staff' and return salary from column 4"
        if st.button("Extract domain from Email"):
            st.session_state.prompt_val = "Extract domain name after @ symbol from email in cell C2"
        if st.button("Nested IF condition"):
            st.session_state.prompt_val = "If cell B2 is greater than 10000 return Bonus, otherwise return No Bonus"

    with col_input:
        default_prompt = st.session_state.get('prompt_val', '')
        user_query = st.text_area(
            "Describe your calculation problem:",
            value=default_prompt,
            placeholder="e.g., Sum column D only if column B says 'Completed' and column C is after Jan 1st 2026...",
            height=120
        )
        
        solve_btn = st.button("⚡ Generate Excel Formula", type="primary", use_container_width=True)
        
        if solve_btn and user_query:
            query_lower = user_query.lower()
            
            # Rule-based intelligent formula simulation matching
            if "lookup" in query_lower or "find" in query_lower:
                formula = "=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B, \"Not Found\")"
                explanation = "XLOOKUP searches for the value in cell A2 inside Sheet2 column A, and returns the matching value from column B. If not found, it returns 'Not Found'."
            elif "sum" in query_lower and ("if" in query_lower or "only" in query_lower or "where" in query_lower):
                formula = "=SUMIFS(D:D, B:B, \"Completed\", C:C, \">2026-01-01\")"
                explanation = "SUMIFS adds up values in column D where column B equals 'Completed' and column C is greater than Jan 1, 2026."
            elif "extract" in query_lower or "domain" in query_lower or "after @" in query_lower:
                formula = "=RIGHT(C2, LEN(C2) - FIND(\"@\", C2))"
                explanation = "FIND locates the position of '@'. RIGHT extracts all text to the right of that position up to the total length of the cell."
            elif "if" in query_lower and ("greater" in query_lower or "return" in query_lower):
                formula = "=IF(B2 > 10000, \"Bonus\", \"No Bonus\")"
                explanation = "Standard IF logical test. Checks if cell B2 is > 10,000. Returns 'Bonus' if True, 'No Bonus' if False."
            elif "count" in query_lower or "how many" in query_lower:
                formula = "=COUNTIFS(A:A, \"Active\", B:B, \">50\")"
                explanation = "COUNTIFS counts rows where column A matches 'Active' and column B is greater than 50."
            else:
                formula = "=INDEX(C:C, MATCH(1, (A:A=\"Target\")*(B:B=\"Confirmed\"), 0))"
                explanation = "Powerful multi-criteria INDEX/MATCH lookup array formula."
            
            # Display generated output card
            st.markdown("### 🎉 Generated Formula")
            st.code(formula, language="excel")
            
            st.info(f"📘 **How it works:** {explanation}")
            st.caption("💡 Simply click the copy icon on the top right corner of the code box above to paste directly into Excel!")

# ==========================================
# 5. FOOTER
# ==========================================
st.divider()
st.caption("🚀 AI Excel Automator • Built with Python & Streamlit • Deployable on Streamlit Community Cloud")
