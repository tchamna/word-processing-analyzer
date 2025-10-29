import streamlit as st
from docx import Document
from docx.shared import RGBColor
import re
import io
import csv

# Punctuations that should end a sentence.
PUNCTUATION = ".?!;"
# Punctuations that should not have a space before them.
SPACED_PUNCTUATION = ",."

def check_paragraph(i, para):
    """Check one paragraph for capitalization, punctuation, and color."""
    issues = []
    text = para.text.strip()
    if not text:
        return issues

    # Color check (detect any colored text except black)
    colored_text = ""
    for run in para.runs:
        if run.font.color and run.font.color.rgb and run.font.color.rgb != RGBColor(0, 0, 0):
            colored_text += run.text
    if colored_text:
        issues.append({
            "line": i + 1,
            "type": "Contains colored text",
            "sentence": text,
            "suggestion": f"Review colored text: '{colored_text}'"
        })

    # Capitalization check
    if not text[0].isupper():
        issues.append({
            "line": i + 1,
            "type": "Missing capitalization",
            "sentence": text,
            "suggestion": f"Capitalize the first letter: '{text[0].upper() + text[1:]}'"
        })

    # Punctuation check
    if text[-1] not in PUNCTUATION:
        issues.append({
            "line": i + 1,
            "type": "Missing punctuation",
            "sentence": text,
            "suggestion": f"Add punctuation at the end, e.g., '{text + '.'}'"
        })

    # Spacing before punctuation check
    punct_issues = []
    if re.search(r'\s,', text):
        punct_issues.append("comma")
    if re.search(r'\s\.', text):
        punct_issues.append("period")
    if punct_issues:
        punct_str = " and ".join(punct_issues) if len(punct_issues) > 1 else punct_issues[0]
        corrected = re.sub(r'\s([,.])', r'\1', text)
        issues.append({
            "line": i + 1,
            "type": "Incorrect spacing before punctuation",
            "sentence": text,
            "suggestion": f"Remove spaces before {punct_str}: '{corrected}'"
        })

    return issues

st.title("Docx Formatting Analyzer")
st.write("Upload a .docx file to check for formatting issues including colored text, capitalization, punctuation, and spacing.")

uploaded_file = st.file_uploader("Choose a .docx file", type=["docx"])

if uploaded_file is not None:
    try:
        doc = Document(io.BytesIO(uploaded_file.read()))
        all_issues = []
        for i, para in enumerate(doc.paragraphs):
            all_issues.extend(check_paragraph(i, para))

        if all_issues:
            st.success(f"Found {len(all_issues)} formatting issues.")
            st.dataframe(all_issues)

            # CSV download
            csv_output = io.StringIO()
            writer = csv.DictWriter(csv_output, fieldnames=["line", "type", "sentence", "suggestion"])
            writer.writeheader()
            writer.writerows(all_issues)
            csv_data = csv_output.getvalue()

            st.download_button(
                label="Download CSV Report",
                data=csv_data.encode('utf-8-sig'),
                file_name="formatting_issues.csv",
                mime="text/csv",
                key="csv_download"
            )

            # Text summary download
            text_summary = "\n".join([f"Line {issue['line']} | {issue['type']} | {issue['sentence']} | {issue['suggestion']}" for issue in all_issues])
            st.download_button(
                label="Download Text Summary",
                data=text_summary,
                file_name="formatting_issues.txt",
                mime="text/plain",
                key="txt_download"
            )
        else:
            st.info("No formatting issues found in the document!")
    except Exception as e:
        st.error(f"Error processing the file: {e}")
else:
    st.info("Please upload a .docx file to get started.")