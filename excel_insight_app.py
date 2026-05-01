import streamlit as st
import pandas as pd
from openai import OpenAI
import io

# ── Config (backend only) ─────────────────────────────────────────────────────
MODEL = "gpt-4o-mini"  # 🔒 Model is fixed — users cannot change this

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Insight Report", page_icon="📊", layout="centered")

st.title("📊 Excel Insight Report")
st.caption("Upload your Excel file and get an AI-powered analysis instantly.")

# ── Sidebar: API key only ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
    st.markdown("---")
    st.markdown("**How it works:**\n1. Enter your API key\n2. Upload an Excel file\n3. Click Generate Report")

# ── File Upload ───────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    selected_sheet = st.selectbox("Select Sheet", sheet_names) if len(sheet_names) > 1 else sheet_names[0]
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    st.subheader("📋 Data Preview")
    st.dataframe(df.head(10), use_container_width=True)

    col1, col2, col3 = st.columns(3)
    col1.metric("Rows", df.shape[0])
    col2.metric("Columns", df.shape[1])
    col3.metric("Sheets", len(sheet_names))

    # ── Generate Report ───────────────────────────────────────────────────────
    if st.button("🔍 Generate Insight Report", type="primary", use_container_width=True):
        if not api_key:
            st.error("Please enter your OpenAI API key in the sidebar.")
        else:
            with st.spinner("Analyzing your data..."):
                buffer = io.StringIO()
                df.info(buf=buffer)
                info_str = buffer.getvalue()

                summary = f"""
Dataset: '{selected_sheet}' sheet from '{uploaded_file.name}'
Shape: {df.shape[0]} rows × {df.shape[1]} columns
Columns: {list(df.columns)}

Data Types:
{info_str}

Statistical Summary:
{df.describe(include='all').to_string()}

First 5 rows sample:
{df.head(5).to_string()}

Missing values per column:
{df.isnull().sum().to_string()}
"""

                prompt = f"""You are a senior data analyst. Analyze the following Excel dataset and produce a clear, structured insight report.

{summary}

Write your report in this exact structure:
1. **Top Insights** – 3 meaningful findings or patterns from the data with facts and figures.
2. **Recommendations** – 3 Actionable suggestions based on the data with facts and figures so user can improve their business performance.

Be concise, professional, and use bullet points where helpful. Please do not suggest data cleaning, give actionable recommendations only.
"""

                try:
                    client = OpenAI(api_key=api_key)
                    response = client.chat.completions.create(
                        model=MODEL,
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.4,
                    )
                    report = response.choices[0].message.content

                    st.subheader("📝 Insight Report")
                    st.markdown(report)

                    st.download_button(
                        label="⬇️ Download Report as .txt",
                        data=report,
                        file_name=f"{uploaded_file.name.split('.')[0]}_insight_report.txt",
                        mime="text/plain",
                    )

                except Exception as e:
                    st.error("Something went wrong. Please check your API key and try again.")

else:
    st.info("👆 Upload an Excel file to get started.")
