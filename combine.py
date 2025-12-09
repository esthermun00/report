import streamlit as st
import os

# -------------------------------
# LOGIN SYSTEM
# -------------------------------
def login_page():
    st.title("🔐 Login Required")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == "generator1" and password == "report123":
            st.session_state["logged_in"] = True
            st.success("Login successful!")
            st.rerun()
        else:
            st.error("Invalid username or password.")

# If not logged in → show login page
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login_page()
    st.stop()

# -------------------------------
# IMPORTS AFTER LOGIN
# -------------------------------
from sustain import (
    load_excel_as_text,
    detect_references_with_ai,
    answer_references,
    merge_answers_into_text,
    generate_sustainability_report,
    save_report_to_word
)

from mda import (
    load_excel_clean, load_consol_to_json,
    generate_business_overview, ask_ai_for_financials,
    financial_report_text, ask_ai_for_revenue, revenue_table_text,
    generate_revenue_narrative, generate_donut_chart,
    generate_gp_pbt_analysis, save_all_outputs_to_word
)

# -------------------------------
# STREAMLIT PAGE SETTINGS
# -------------------------------
st.set_page_config(page_title="Multi-Report Generator", layout="wide")
st.title("📊 Multi-Report Generator")

# Use three columns: left | divider | right
left_col, divider_col, right_col = st.columns([5, 0.1, 5])

# Divider line
with divider_col:
    st.markdown(
        "<div style='border-left:2px solid #d9d9d9; height:100vh;'></div>", 
        unsafe_allow_html=True
    )

# -------------------------------
# SESSION STATE INITIALIZATION
# -------------------------------
if "ann_files" not in st.session_state:
    st.session_state["ann_files"] = []     # RIGHT column upload state

if "sust_files" not in st.session_state:
    st.session_state["sust_files"] = []    # LEFT column upload state


# ------------------------
# LEFT COLUMN (SUSTAINABILITY)
# ------------------------
# ------------------------
# LEFT COLUMN (SUSTAINABILITY)
# ------------------------
with left_col:
    st.header("🌿 Sustainability Report")
    uploaded_files_sust = st.file_uploader(
        "Upload Excel files for Sustainability Reports", 
        type=["xlsx"], accept_multiple_files=True,
        key="sust_upload"
    )

    if uploaded_files_sust is not None and len(uploaded_files_sust) > 0:
        st.session_state["sust_files"] = uploaded_files_sust

    # Only show Generate button if files exist
    if len(st.session_state.get("sust_files", [])) > 0:
        if st.button("Generate Sustainability Reports", key="gen_sust"):

            with st.spinner("Generating Sustainability Report..."):
                reports = {}
                os.makedirs("reports", exist_ok=True)

                for file in st.session_state["sust_files"]:
                    st.write(f"Processing: {file.name}")
                    sheets_text = load_excel_as_text(file)
                    references = detect_references_with_ai(sheets_text)
                    answers = answer_references(sheets_text, references)
                    merged_text = merge_answers_into_text(sheets_text, answers)
                    report_text = generate_sustainability_report(merged_text)
                    reports[file.name] = report_text

                    word_file_path = os.path.join("reports", f"{file.name}_report.docx")
                    save_report_to_word(report_text, word_file_path)

            st.success("✅ Sustainability reports generated and saved!")

            # Download buttons
            for fname in reports:
                word_path = os.path.join("reports", f"{fname}_report.docx")
                with open(word_path, "rb") as f:
                    st.download_button(
                        label=f"📥 Download Sustainability Report",
                        data=f,
                        file_name=f"{fname}_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )


# ------------------------
# RIGHT COLUMN (MD&A)
# ------------------------
with right_col:
    st.header("🏢 MD&A Report")
    uploaded_files_ann = st.file_uploader(
        "Upload Excel files (MD&A, Consol, etc.)",
        type=["xlsx"], accept_multiple_files=True,
        key="ann_upload"
    )

    # Initialize session_state storage
    if "prev_ann_files" not in st.session_state:
        st.session_state.prev_ann_files = []
    if "ann_ready" not in st.session_state:
        st.session_state.ann_ready = False
    if "ann_messages" not in st.session_state:
        st.session_state.ann_messages = []  # store success/warning messages

    current_filenames = [f.name for f in uploaded_files_ann] if uploaded_files_ann else []
    previous_filenames = st.session_state.prev_ann_files
    new_upload = current_filenames != previous_filenames and len(current_filenames) > 0
    st.session_state.prev_ann_files = current_filenames

    if new_upload:
        with st.spinner("Processing uploaded files..."):
            uploaded_files_dict = {file.name: file for file in uploaded_files_ann}
            raw_text_dict = {name: load_excel_clean(file_obj) for name, file_obj in uploaded_files_dict.items()}

            # Store messages in session_state so they persist
            st.session_state.ann_messages = [
                ("success",f"{len(uploaded_files_ann)} file(s) uploaded and processed successfully!")
            ]

            consol_file_obj = uploaded_files_dict.get("consol.xlsx")
            if consol_file_obj:
                clean_json = load_consol_to_json(consol_file_obj, sheet_name="GRP IS")
                st.session_state.ann_messages.append(("success","Using consol.xlsx for financial extraction."))
            else:
                clean_json = None
                st.session_state.ann_messages.append(("error","No consol.xlsx uploaded. Financial extraction cannot proceed."))

            mdna_files = {k: v for k, v in raw_text_dict.items() if "md&a" in k.lower()}
            other_files = {k: v for k, v in raw_text_dict.items() if k not in mdna_files}

        st.session_state.raw_text_dict = raw_text_dict
        st.session_state.mdna_files = mdna_files
        st.session_state.other_files = other_files
        st.session_state.clean_json = clean_json
        st.session_state.ann_ready = True

    # -----------------------------------------------------------
    # Show persisted messages even on rerun
    # -----------------------------------------------------------
    for item in st.session_state.get("ann_messages", []):
        if isinstance(item, tuple):
            level, msg = item
        else:
            level, msg = "success", item  # backward compatibility

        if level == "error":
            st.error(msg)
        elif level == "warning":
            st.warning(msg)
        else:
            st.success(msg)


    # -----------------------------------------------------------
    # Only show generate button if files processed
    # -----------------------------------------------------------
    if st.session_state.get("ann_ready", False):
        if st.button("Generate MD&A Report", key="gen_ann"):
            if st.session_state.clean_json is None:
                st.error("Cannot generate financial report without consol.xlsx.")
            else:
                with st.spinner("Generating MD&A Financial Report..."):
                    try:
                        raw_text_dict = st.session_state.raw_text_dict
                        mdna_files = st.session_state.mdna_files
                        other_files = st.session_state.other_files
                        clean_json = st.session_state.clean_json

                        business_text = generate_business_overview(mdna_files, other_files)

                        all_sheets_text = {}
                        for v in raw_text_dict.values():
                            all_sheets_text.update(v)

                        financial_data = ask_ai_for_financials(all_sheets_text, clean_json)
                        financial_text = financial_report_text(financial_data)

                        revenue_data = ask_ai_for_revenue(all_sheets_text, clean_json)
                        revenue_text = revenue_table_text(revenue_data)
                        revenue_narrative = generate_revenue_narrative(revenue_data, all_sheets_text)
                        revenue_chart_buf = generate_donut_chart(revenue_data)

                        gp_commentary, pbt_commentary, gp_chart_buf, pbt_chart_buf = generate_gp_pbt_analysis(financial_data)

                        word_buf = save_all_outputs_to_word(
                            business_text, financial_text, revenue_text,
                            revenue_narrative, gp_commentary, pbt_commentary,
                            revenue_chart_buf, gp_chart_buf, pbt_chart_buf
                        )

                        st.success("✅ Annual report generated successfully!")
                        st.download_button(
                            label="📥 Download Annual Report Word",
                            data=word_buf,
                            file_name="Annual_Report.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

                    except Exception as e:
                        st.error(f"❌ Error during report generation: {e}")
