import os
import tempfile
import pandas as pd
import streamlit as st

from checker import run_checks, add_word_comments

st.set_page_config(page_title="Thesis Compliance Checker (APA + Guidelines)", layout="wide")

st.title("Thesis Compliance Checker")
st.caption("DOCX-only. Checks guideline structure, formatting, and APA-only tables/figures and references.")

uploaded = st.file_uploader("Upload thesis (.docx only)", type=["docx"])
degree_type = st.selectbox("Degree type", ["Dissertation", "MPhil", "PhD"], index=0)

if uploaded is None:
    st.stop()

with tempfile.TemporaryDirectory() as tmpdir:
    in_path = os.path.join(tmpdir, "thesis.docx")
    with open(in_path, "wb") as f:
        f.write(uploaded.getbuffer())

    issues, meta = run_checks(in_path, degree_type)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total issues", meta["issues_total"])
    c2.metric("High", meta["issues_high"])
    c3.metric("Medium", meta["issues_medium"])
    c4.metric("Low", meta["issues_low"])

    st.subheader("Checklist results")
    if not issues:
        st.success("No issues detected by the current rule set.")
    else:
        df = pd.DataFrame([{
            "Severity": i.severity,
            "Rule": i.rule_id,
            "Message": i.message,
            "Evidence": i.evidence,
            "Location": i.location_hint,
            "Anchor paragraph": i.anchor_paragraph_index + 1
        } for i in issues])
        st.dataframe(df, use_container_width=True, hide_index=True)

    st.subheader("Downloads")
    annotated_path = os.path.join(tmpdir, "thesis_annotated.docx")
    add_word_comments(in_path, issues, annotated_path)

    with open(annotated_path, "rb") as f:
        st.download_button(
            "Download annotated DOCX (Word comments)",
            data=f,
            file_name="thesis_annotated.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
