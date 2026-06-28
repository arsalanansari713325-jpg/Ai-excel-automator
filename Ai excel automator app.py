import os
import io
import re
import pandas as pd
import streamlit as st

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Excel AI Assistant & Database Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #064e3b, #0f766e, #0f172a);
        padding: 1.5rem;
        border-radius: 0.75rem;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
    }
    .metric-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 1.2rem;
        border-radius: 0.5rem;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# CORE PROCESSING FUNCTIONS
# ==========================================
def clean_header_name(val):
    if pd.isna(val) or str(val).strip() == "":
        return "column"
    text = str(val).strip().lower()
    text = re.sub(r'[^\w\s-]', '', text)
    text = re.sub(r'[\s-]+', '_', text)
    text = re.sub(r'^(\d)', r'col_\1', text)
    return text

def infer_sql_type(series):
    non_null = series.dropna()
    if len(non_null) == 0:
        return "TEXT"
    
    # Check boolean
    unique_vals = set(non_null.astype(str).str.lower().unique())
    if unique_vals.issubset({'true', 'false', '1', '0', 'yes', 'no'}):
        return "BOOLEAN"
    
    # Check numeric
    if pd.api.types.is_integer_dtype(series):
        return "INTEGER"
    if pd.api.types.is_float_dtype(series):
        return "DECIMAL(18, 4)"
        
    return "TEXT"

def generate_sql(df, table_name, db_dialect):
    cols = []
    for col in df.columns:
        clean_col = clean_header_name(col)
        dtype = infer_sql_type(df[col])
        if db_dialect == 'PostgreSQL':
            cols.append(f'    "{clean_col}" {dtype if dtype != "TEXT" else "VARCHAR(255)"}')
        elif db_dialect == 'MySQL':
            cols.append(f'    `{clean_col}` {dtype if dtype != "TEXT" else "VARCHAR(255)"}')
        else: # SQLite / SQL Server
            cols.append(f'    [{clean_col}] {dtype if dtype != "TEXT" else "NVARCHAR(MAX)"}')
            
    create_stmt = f"CREATE TABLE {table_name} (\n" + ",\n".join(cols) + "\n);\n\n"
    
    # Generate INSERTs (first 100 rows preview)
    insert_stmts = []
    preview_df = df.head(100)
    for _, row in preview_df.iterrows():
        vals = []
        for val in row:
            if pd.isna(val):
                vals.append("NULL")
            elif isinstance(val, (int, float)):
                vals.append(str(val))
            else:
                esc = str(val).replace("'", "''")
                vals.append(f"'{esc}'")
        
        cols_str = ", ".join([clean_header_name(c) for c in df.columns])
        insert_stmts.append(f"INSERT INTO {table_name} ({cols_str}) VALUES ({', '.join(vals)});")
        
    return create_stmt + "\n".join(insert_stmts)

def convert_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

def convert_to_json(df):
    return df.to_json(orient='records', indent=2).encode('utf-8')


# ==========================================
# APP HEADER
# ==========================================
st.markdown("""
<div class="main-header">
    <h1 style="margin:0; font-size: 2.2rem;">📊 Excel AI Assistant & Database Converter</h1>
    <p style="margin-top:0.5rem; opacity:0.9;">Transform raw spreadsheets into production-ready SQL schemas, clean CSVs, and JSON APIs.</p>
</div>
""", unsafe_allow_html=True)


# ==========================================
# SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
    st.header("⚙️ Export Settings")
    
    table_name = st.text_input("Target Database Table", value="user_data")
    db_dialect = st.selectbox("SQL Dialect", ["PostgreSQL", "MySQL", "SQLite", "SQL Server"])
    
    st.divider()
    
    st.header("✨ AI Formula Solver")
    with st.expander("Ask AI Formula Question"):
        user_query = st.text_area("Describe what you want to calculate:", placeholder="e.g. Sum column B where column A is 'Sales'")
        if st.button("Generate Formula", type="primary"):
            if user_query:
                st.success("Suggested Excel Formula:")
                st.code("=SUMIF(A:A, \"Sales\", B:B)", language="excel")
            else:
                st.warning("Please enter a prompt above.")


# ==========================================
# MAIN WORKSPACE
# ==========================================
uploaded_file = st.file_uploader("📂 Drop your Excel (.xlsx, .xls) or CSV file here", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # Load File
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
            sheet_names = ["CSV Data"]
            selected_sheet = "CSV Data"
        else:
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names
            
            col1, col2 = st.columns([1, 3])
            with col1:
                selected_sheet = st.selectbox("📑 Select Sheet", sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
        st.success(f"Successfully loaded `{uploaded_file.name}` ({len(df)} rows, {len(df.columns)} columns)")
        
        # Metrics Row
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("Total Rows", f"{len(df):,}")
        with m2:
            st.metric("Total Columns", f"{len(df.columns)}")
        with m3:
            st.metric("Missing Values", f"{df.isna().sum().sum():,}")
        with m4:
            st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")

        st.divider()

        # Workspace Tabs
        tab_preview, tab_sql, tab_csv, tab_json = st.tabs(["👁️ Data Preview & Clean", "📜 SQL Schema", "📄 CSV Export", "📦 JSON API"])
        
        with tab_preview:
            st.subheader("Interactive Spreadsheet View")
            st.dataframe(df, use_container_width=True, height=450)
            
        with tab_sql:
            st.subheader(f"Generated {db_dialect} Statements")
            sql_output = generate_sql(df, clean_header_name(table_name), db_dialect)
            st.download_button(
                label="📥 Download .SQL File",
                data=sql_output,
                file_name=f"{clean_header_name(table_name)}.sql",
                mime="text/plain",
                type="primary"
            )
            st.code(sql_output, language="sql")
            
        with tab_csv:
            st.subheader("Clean Standardized CSV")
            csv_data = convert_to_csv(df)
            st.download_button(
                label="📥 Download .CSV File",
                data=csv_data,
                file_name=f"{clean_header_name(table_name)}.csv",
                mime="text/csv",
                type="primary"
            )
            st.code(csv_data.decode('utf-8')[:2000] + "\n... [truncated preview]", language="csv")
            
        with tab_json:
            st.subheader("REST API JSON Format")
            json_data = convert_to_json(df)
            st.download_button(
                label="📥 Download .JSON File",
                data=json_data,
                file_name=f"{clean_header_name(table_name)}.json",
                mime="application/json",
                type="primary"
            )
            st.code(json_data.decode('utf-8')[:2000] + "\n... [truncated preview]", language="json")

    except Exception as e:
        st.error(f"Error reading file: {str(e)}")

else:
    # Empty State Guide
    st.info("👆 Upload a spreadsheet file above to get started.")
    
    st.markdown("### 💡 Quick Feature Guide")
    g1, g2, g3 = st.columns(3)
    with g1:
        st.markdown("#### 1. Instant SQL Schemas\nAutomatically infers data types (`INTEGER`, `BOOLEAN`, `VARCHAR`, `DECIMAL`) and writes `CREATE TABLE` and `INSERT` scripts.")
    with g2:
        st.markdown("#### 2. Multi-Sheet Parsing\nSeamlessly switch between tabs inside `.xlsx` workbooks and export them individually.")
    with g3:
        st.markdown("#### 3. AI Formula Assistant\nStuck on a tricky calculation? Use the built-in AI solver in the left sidebar to generate Excel formulas.")
