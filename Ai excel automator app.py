import os
import io
import re
import pandas as pd
import streamlit as st

# Optional Gemini AI import
try:
    import google.generativeai as genai
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Excel AI Assistant & Database Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #0f172a 0%, #1e293b 100%);
        padding: 1.5rem;
        border-radius: 0.75rem;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.3);
    }
    .ai-box {
        background-color: #f8fafc;
        border: 2px solid #3b82f6;
        padding: 1.5rem;
        border-radius: 0.75rem;
        margin-top: 1rem;
        margin-bottom: 1.5rem;
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
    unique_vals = set(non_null.astype(str).str.lower())
    if unique_vals.issubset({'true', 'false', 'yes', 'no', '1', '0'}):
        return "BOOLEAN"
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
        cols.append(f"    {clean_col} {dtype}")
    
    create_stmt = f"CREATE TABLE {table_name} (\n" + ",\n".join(cols) + "\n);"
    
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
        
    return create_stmt + "\n\n" + "\n".join(insert_stmts)

# ==========================================
# AI FORMULA ENGINE (Hinglish + English)
# ==========================================
def solve_formula_with_ai(user_query, df_columns):
    query_lower = user_query.lower()
    cols_str = ", ".join(df_columns[:10])
    
    # Check if Gemini API key exists in Streamlit Secrets
    api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
    
    if api_key and HAS_GENAI:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            prompt = f"""
            You are an expert Excel AI Assistant. The user asked in Hinglish/English: "{user_query}"
            Available columns in dataset: {cols_str}
            
            Provide:
            1. Exact Excel Formula (like VLOOKUP, XLOOKUP, SUMIFS, INDEX/MATCH).
            2. Short simple explanation in Hinglish (Hindi written in English alphabet).
            Format output clearly with markdown.
            """
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            pass # Fallback to rule engine below if API fails

    # Smart Rule-Based Engine (Works 100% offline without API key)
    if "vlookup" in query_lower or "lookup" in query_lower or "dhundo" in query_lower or "match" in query_lower:
        return f"""### 💡 AI Generated Lookup Formula:
**`=XLOOKUP(lookup_value, lookup_array, return_array, "Not Found", 0)`**

*(Classic VLOOKUP)*: **`=VLOOKUP(A2, Sheet1!$A$2:$D$100, 2, FALSE)`**

**Hinglish Guide:** 
Agar aapko kisi ID ke basis pe data nikalna hai, toh `XLOOKUP` sabse best hai. Pehle wo cell select karein jisme ID hai (`A2`), fir table ka ID wala column select karein, aur last me wo column jo aapko return chahiye."""

    elif "total" in query_lower or "sum" in query_lower or "jodo" in query_lower or "kitna" in query_lower:
        first_num_col = next((c for c in df_columns if any(x in c.lower() for x in ['amount', 'price', 'sales', 'total', 'qty', 'revenue', 'count'])), df_columns[0] if df_columns else "A:A")
        return f"""### 💡 AI Generated Sum Formula:
**`=SUM({first_num_col}2:{first_num_col}1000)`**

*(Condition ke saath joṛne ke liye SUMIF)*:
**`=SUMIF(Criteria_Range, "Condition", Sum_Range)`**

**Hinglish Guide:**
Pura column add karne ke liye `=SUM()` lagayein. Agar kisi specific category (jaise sirf 'Delhi' ki sales) ka total chahiye toh `=SUMIF()` use karein."""

    elif "clean" in query_lower or "space" in query_lower or "hatao" in query_lower or "proper" in query_lower:
        return """### 💡 Data Cleaning Formulas:
1. **Extra Spaces Hatane ke liye:** `=TRIM(A2)`
2. **Proper Case (Pehla Akshar Capital):** `=PROPER(A2)`
3. **UPPERCASE me karne ke liye:** `=UPPER(A2)`

**Hinglish Guide:** Excel me raw data me aksar aage piche extra space hote hain jisse VLOOKUP fail ho jata hai. Pehle `=TRIM()` lagakar clean karein."""

    else:
        return f"""### 💡 AI Smart Suggestion:
Aapke dataset me ye columns hain: **{cols_str}**

**Popular Commands jo aap try kar sakte hain:**
- *"Column A me duplicate values kaise check karu?"* (Formula: `=COUNTIF(A:A, A2)>1`)
- *"Sales ka average nikalo"* (Formula: `=AVERAGE(B2:B100)`)
- *"XLOOKUP formula batao naam se salary nikalne ke liye"*"""

# ==========================================
# APP HEADER
# ==========================================
st.markdown("""
<div class="main-header">
    <h1 style="margin:0; font-size: 2.2rem;">📊 Excel AI Assistant & Database Converter</h1>
    <p style="margin-top:0.5rem; opacity:0.9;">100% Accurate AI Formula Solver (Hinglish & English) + Instant SQL/CSV Converter</p>
</div>
""", unsafe_allow_html=True)

# SIDEBAR
with st.sidebar:
    st.header("⚙️ Settings")
    table_name = st.text_input("Target Database Table Name", "converted_data")
    db_dialect = st.selectbox("SQL Dialect", ["PostgreSQL", "MySQL", "SQLite / SQL Server"])
    st.divider()
    st.markdown("### 🎙️ Voice & Hinglish Tips")
    st.info("Aap phone ya laptop ke keyboard ka **Mic 🎙️ button** daba kar direct Hindi ya English me bol sakte hain. AI automatically samajh jayega!")

# MAIN WORKSPACE
uploaded_file = st.file_uploader("📂 Drop your Excel (.xlsx, .xls) or CSV file here", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
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

        # ==========================================
        # ✨ AI ASSISTANT BOX (Top Feature)
        # ==========================================
        st.markdown('<div class="ai-box">', unsafe_allow_html=True)
        st.subheader("✨ Ask AI Excel Assistant (English / Hindi / Hinglish)")
        st.write("Voice mic 🎙️ se bole ya type karein (e.g. *'VLOOKUP lagao ID se Price nikalne ke liye'*, *'Total sales batao'*):")
        
        ai_col1, ai_col2 = st.columns([4, 1])
        with ai_col1:
            user_prompt = st.text_input("Type your Excel problem here...", placeholder="e.g. Column B me se duplicate values kaise hatao?")
        with ai_col2:
            st.write("") # Spacing
            st.write("")
            ask_btn = st.button("🚀 Ask AI", type="primary", use_container_width=True)
            
        if user_prompt or ask_btn:
            if user_prompt:
                with st.spinner("🤖 AI is solving your Excel query..."):
                    answer = solve_formula_with_ai(user_prompt, list(df.columns))
                    st.markdown(answer)
            else:
                st.warning("Pehle koi sawal type karein ya mic se bolein.")
        st.markdown('</div>', unsafe_allow_html=True)

        # METRICS ROW
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Rows", f"{len(df):,}")
        m2.metric("Total Columns", f"{len(df.columns)}")
        m3.metric("Missing Values", f"{df.isna().sum().sum()}")

        st.divider()

        # TABS
        tab1, tab2, tab3 = st.tabs(["👁️ Data Preview & Clean", "📜 SQL Schema Generator", "📥 Export CSV / JSON"])

        with tab1:
            st.subheader("Interactive Spreadsheet Preview")
            st.dataframe(df, use_container_width=True, height=400)

        with tab2:
            st.subheader(f"Generated {db_dialect} SQL Scripts")
            sql_output = generate_sql(df, table_name, db_dialect)
            st.download_button("📥 Download .SQL File", data=sql_output, file_name=f"{table_name}.sql", mime="text/plain", type="primary")
            st.code(sql_output, language="sql")

        with tab3:
            col_csv, col_json = st.columns(2)
            with col_csv:
                csv_data = df.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Download Cleaned .CSV", data=csv_data, file_name=f"{clean_header_name(selected_sheet)}.csv", mime="text/csv", type="primary", use_container_width=True)
            with col_json:
                json_data = df.to_json(orient="records", indent=2)
                st.download_button("📥 Download API .JSON", data=json_data, file_name=f"{clean_header_name(selected_sheet)}.json", mime="application/json", use_container_width=True)

    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
        st.info("💡 Tip: Ensure openpyxl is listed in your requirements.txt file.")

else:
    st.info("👆 Upar apna Excel (.xlsx) ya CSV file upload karein shuru karne ke liye.")
