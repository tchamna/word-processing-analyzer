import argparse
from docx import Document
from docx.shared import RGBColor
import os
import csv
import re

# ==========================================
# DEFAULT CONFIGURATION (can be overridden from CLI)
# ==========================================
DEFAULT_FILE = r"G:\My Drive\Mbú'ŋwɑ̀'nì\Livres Ewondo\Ewondo_CMR_test_2.docx"
DEFAULT_FILE = r"G:\My Drive\Mbú'ŋwɑ̀'nì\Livres Nufi\Nufi_test_processing.docx"

language_name = os.path.basename(DEFAULT_FILE).split('_')[0]  # Extract language name from filename 

DEFAULT_OUTPUT = f"formatting_issues_{language_name}.txt"
DEFAULT_CSV = f"formatting_issues_{language_name}.csv"
# ==========================================

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


def main():
    """Main entry point with hybrid argument handling."""
    parser = argparse.ArgumentParser(description="Check a .docx file for formatting issues.")
    parser.add_argument("filepath", nargs="?", default=DEFAULT_FILE, help="Path to the .docx file")
    parser.add_argument("--output", help="Output file for text summary.")
    parser.add_argument("--csv", help="Output CSV file name.")
    args = parser.parse_args()

    # Fallback to defaults if arguments not provided
    filepath = args.filepath
    output_txt = args.output or DEFAULT_OUTPUT
    output_csv = args.csv or DEFAULT_CSV

    print(f"\n📘 Using settings:")
    print(f"  File:   {filepath}")
    print(f"  Output: {output_txt}")
    print(f"  CSV:    {output_csv}\n")

    if not os.path.exists(filepath):
        print(f"❌ File not found: {filepath}")
        return

    try:
        doc = Document(filepath)
    except Exception as e:
        print(f"❌ Error opening file: {e}")
        return

    all_issues = []
    for i, para in enumerate(doc.paragraphs):
        all_issues.extend(check_paragraph(i, para))

    # --- Write plain text summary
    with open(output_txt, "w", encoding="utf-8") as f:
        for issue in all_issues:
            f.write(f"Line {issue['line']} | {issue['type']} | {issue['sentence']} | {issue['suggestion']}\n")

    # --- Write CSV report
    with open(output_csv, "w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=["Line", "Issue Type", "Sentence", "Suggested Fix"])
        writer.writeheader()
        for issue in all_issues:
            writer.writerow({
                "Line": issue["line"],
                "Issue Type": issue["type"],
                "Sentence": issue["sentence"],
                "Suggested Fix": issue["suggestion"]
            })

    print(f"✅ Found {len(all_issues)} formatting issues.")
    print(f"📄 Text summary saved to: {output_txt}")
    print(f"📊 CSV report saved to:  {output_csv}")


if __name__ == "__main__":
    main()
