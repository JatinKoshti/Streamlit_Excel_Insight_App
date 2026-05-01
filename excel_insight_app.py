import streamlit as st
import pandas as pd
from openai import OpenAI
import io
# ── Build data helpers ──────────────────────────────────────────────────────
def build_data_context(df, sheet_name, file_name, action_summary=""):
    buffer = io.StringIO()
    df.info(buf=buffer)
    info_str = buffer.getvalue()

    action_info = f"Active data action: {action_summary}\n\n" if action_summary else ""

    return f"""You are a smart Sr. Data Analyst assistant. The user has uploaded an Excel file and will ask questions about it. Please do not give any generic insight and without any facts and figures. Give each statement like current status and Actionable Insight.

{action_info}File: '{file_name}', Sheet: '{sheet_name}'
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


def build_filter_summary(selected_action, filter_column=None, operator=None, filter_value=None, second_value=None, group_by=None, agg_column=None, agg_func=None, result_shape=None):
    if selected_action == "Filter" and filter_column:
        summary = f"Filter: {filter_column} {operator} {filter_value}"
        if second_value is not None:
            summary += f" and {second_value}"
        if result_shape is not None:
            summary += f" → {result_shape[0]} rows"
        return summary
    if selected_action == "Aggregate" and agg_column and agg_func:
        summary = f"Aggregation: {agg_func} of {agg_column}"
        if group_by:
            summary += f" grouped by {group_by}"
        return summary
    return ""
# ── Config ────────────────────────────────────────────────────────────────────
MODEL = "gpt-4o"  # 🔒 Fixed — users cannot change this

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel Data Chatbot", page_icon="📊", layout="centered")

st.title("📊 Excel Data Chatbot")
st.caption("Upload your Excel file and ask questions about your data.")

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
    st.markdown("---")

    uploaded_file = st.file_uploader("Upload Excel File (.xlsx / .xls)", type=["xlsx", "xls"])

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # ✅ Always show sheet selector (even for single sheet)
        selected_sheet = st.selectbox("Select Sheet", sheet_names)

        # ✅ Always load df regardless of sheet count
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

        st.success(f"✅ Loaded: **{uploaded_file.name}**")
        st.caption(f"{df.shape[0]} rows × {df.shape[1]} columns")

        numeric_columns = df.select_dtypes(include="number").columns.tolist()
        filtered_df = df
        action_summary = ""
        use_filtered_for_chat = False

        selected_action = st.selectbox("Data action", ["None", "Aggregate", "Filter"], help="Choose a quick data operation.")

        if selected_action == "Aggregate":
            if numeric_columns:
                agg_column = st.selectbox("Numeric column", numeric_columns, key="agg_column")
                agg_func = st.selectbox("Operation", ["sum", "mean", "median", "min", "max", "count"], key="agg_func")
                group_by = st.selectbox("Group by (optional)", [""] + [c for c in df.columns if c != agg_column], key="group_by")

                if group_by:
                    agg_df = df.groupby(group_by, dropna=False)[agg_column].agg(agg_func).reset_index()
                    st.write(f"**{agg_func.title()} of {agg_column} grouped by {group_by}:**")
                    st.dataframe(agg_df, use_container_width=True)
                else:
                    agg_value = getattr(df[agg_column], agg_func)()
                    st.metric(f"{agg_func.title()} of {agg_column}", f"{agg_value}")

                action_summary = build_filter_summary(selected_action, group_by=group_by or None, agg_column=agg_column, agg_func=agg_func)
            else:
                st.warning("No numeric columns available for aggregation.")

        elif selected_action == "Filter":
            filter_column = st.selectbox("Filter column", df.columns, key="filter_column")
            dtype = df[filter_column].dtype

            if pd.api.types.is_numeric_dtype(dtype):
                operator = st.selectbox("Operator", ["=", ">", "<", ">=", "<=", "between"], key="filter_operator")
                if operator == "between":
                    filter_value = st.number_input("Min value", value=float(df[filter_column].min()), key="filter_value_min")
                    second_value = st.number_input("Max value", value=float(df[filter_column].max()), key="filter_value_max")
                    filtered_df = df[df[filter_column].between(filter_value, second_value)]
                else:
                    filter_value = st.number_input("Value", value=float(df[filter_column].median()), key="filter_value")
                    if operator == "=":
                        filtered_df = df[df[filter_column] == filter_value]
                    elif operator == ">":
                        filtered_df = df[df[filter_column] > filter_value]
                    elif operator == "<":
                        filtered_df = df[df[filter_column] < filter_value]
                    elif operator == ">=":
                        filtered_df = df[df[filter_column] >= filter_value]
                    else:
                        filtered_df = df[df[filter_column] <= filter_value]
                    second_value = None
            else:
                operator = st.selectbox("Operator", ["equals", "contains", "starts with", "ends with"], key="filter_operator")
                filter_value = st.text_input("Value", key="filter_value")
                if filter_value:
                    if operator == "equals":
                        filtered_df = df[df[filter_column].astype(str) == filter_value]
                    elif operator == "contains":
                        filtered_df = df[df[filter_column].astype(str).str.contains(filter_value, case=False, na=False)]
                    elif operator == "starts with":
                        filtered_df = df[df[filter_column].astype(str).str.startswith(filter_value, na=False)]
                    else:
                        filtered_df = df[df[filter_column].astype(str).str.endswith(filter_value, na=False)]
                second_value = None

            st.write(f"Filtered rows: **{filtered_df.shape[0]}** of **{df.shape[0]}**")
            st.dataframe(filtered_df.head(10), use_container_width=True)
            use_filtered_for_chat = st.checkbox("Use filtered data for chat", value=True, key="use_filtered_for_chat")
            action_summary = build_filter_summary(selected_action, filter_column=filter_column, operator=operator, filter_value=filter_value, second_value=second_value, result_shape=filtered_df.shape)

        st.markdown("---")

        if st.button("🗑️ Clear Chat", use_container_width=True):
            st.session_state.messages = []
            st.rerun()

    st.markdown("---")
    st.markdown("**How it works:**\n1. Enter your API key\n2. Upload an Excel file\n3. Ask any question about your data")

# ── Initialize chat history ───────────────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = []

# ── Main chat area ────────────────────────────────────────────────────────────
if not uploaded_file:
    st.info("👈 Please upload an Excel file from the sidebar to get started.")
else:
    with st.expander("📋 Data Preview", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)
        col1, col2, col3 = st.columns(3)
        col1.metric("Rows", df.shape[0])
        col2.metric("Columns", df.shape[1])
        col3.metric("Sheets", len(sheet_names))

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

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("Ask a question about your data..."):
        if not api_key:
            st.error("Please enter your OpenAI API key in the sidebar.")
        else:
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    try:
                        client = OpenAI(api_key=api_key)

                        chat_df = filtered_df if use_filtered_for_chat else df
                        system_context = build_data_context(chat_df, selected_sheet, uploaded_file.name, action_summary)
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
