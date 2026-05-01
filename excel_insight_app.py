import streamlit as st
import pandas as pd
from openai import OpenAI
import io

# ── Config ────────────────────────────────────────────────────────────────────
MODEL = "gpt-4o-mini"  # 🔒 Fixed — users cannot change this

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Data Chatbot", page_icon="📊", layout="centered")

st.title("📊 Excel Data Chatbot")
st.caption("Upload your Excel file and ask questions about your data.")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
    st.markdown("---")

    # File upload inside sidebar
    uploaded_file = st.file_uploader("Upload Excel File (.xlsx / .xls)", type=["xlsx", "xls"])

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("Select Sheet", sheet_names) if len(sheet_names) > 1 else sheet_names[0]
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

        st.success(f"✅ Loaded: **{uploaded_file.name}**")
        st.caption(f"{df.shape[0]} rows × {df.shape[1]} columns")

        if st.button("🗑️ Clear Chat", use_container_width=True):
            st.session_state.messages = []
            st.rerun()

    st.markdown("---")
    st.markdown("**How it works:**\n1. Enter your API key\n2. Upload an Excel file\n3. Ask any question about your data")

# ── Build data context (once per file upload) ─────────────────────────────────
def build_data_context(df, sheet_name, file_name):
    buffer = io.StringIO()
    df.info(buf=buffer)
    info_str = buffer.getvalue()

    return f"""You are a smart Sr. Data analyst assistant . The user has uploaded an Excel file and will ask questions about it Please do not give any generic insight and without any facts and figures. Your each statment should be depended on the data and analysis only like current status and Actionable Insight.

File: '{file_name}', Sheet: '{sheet_name}'
Shape: {df.shape[0]} rows × {df.shape[1]} columns
Columns: {list(df.columns)}

Data Types & Info:
{info_str}

Statistical Summary:
{df.describe(include='all').to_string()}

First 10 rows:
{df.head(10).to_string()}

Missing values:
{df.isnull().sum().to_string()}

Answer the user's questions based on this data. Be concise, use bullet points and facts/figures where relevant. Do not suggest data cleaning. Give actionable insights only."""

# ── Initialize chat history ───────────────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = []

# ── Main chat area ────────────────────────────────────────────────────────────
if not uploaded_file:
    st.info("👈 Please upload an Excel file from the sidebar to get started.")
else:
    # Show data preview in an expander
    with st.expander("📋 Data Preview", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)
        col1, col2, col3 = st.columns(3)
        col1.metric("Rows", df.shape[0])
        col2.metric("Columns", df.shape[1])
        col3.metric("Sheets", len(sheet_names))

    # Suggested starter questions
    if len(st.session_state.messages) == 0:
        st.markdown("#### 💡 Try asking:")
        suggestions = [
            "What are the top insights from this data?",
            "Which category has the highest value?",
            "What trends do you see in this data?",
            "Give me a summary of this dataset.",
        ]
        cols = st.columns(2)
        for i, suggestion in enumerate(suggestions):
            if cols[i % 2].button(suggestion, use_container_width=True):
                st.session_state.messages.append({"role": "user", "content": suggestion})
                st.rerun()

    # Display chat history
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Chat input
    if prompt := st.chat_input("Ask a question about your data..."):
        if not api_key:
            st.error("Please enter your OpenAI API key in the sidebar.")
        else:
            # Add user message
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            # Get AI response
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    try:
                        client = OpenAI(api_key=api_key)

                        # Build messages: system context + full chat history
                        system_context = build_data_context(df, selected_sheet, uploaded_file.name)
                        api_messages = [{"role": "system", "content": system_context}] + [
                            {"role": m["role"], "content": m["content"]}
                            for m in st.session_state.messages
                        ]

                        response = client.chat.completions.create(
                            model=MODEL,
                            messages=api_messages,
                            temperature=0.4,
                        )
                        reply = response.choices[0].message.content
                        st.markdown(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})

                    except Exception as e:
                        st.error("Something went wrong. Please check your API key and try again.")
