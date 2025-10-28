import argparse
from docx import Document
from docx.shared import RGBColor
import os
import csv

# ==========================================
# DEFAULT CONFIGURATION (can be overridden from CLI)
# ==========================================
DEFAULT_FILE = r"G:\My Drive\Mbú'ŋwɑ̀'nì\Livres Nufi\Nufi_test_processing.docx"
DEFAULT_FILE = r"G:\My Drive\Mbú'ŋwɑ̀'nì\Livres Ewondo\Ewondo_CMR_test_2.docx"
DEFAULT_FILE = r"G:\My Drive\Mbú'ŋwɑ̀'nì\Livres Ewondo\Ewondo_CMR_test_.docx"

language_name = os.path.basename(DEFAULT_FILE).split('_')[0]  # Extract language name from filename 

DEFAULT_COLOR = "red"           # options: "red", "green", or None
DEFAULT_OUTPUT = f"formatting_issues_{language_name}.txt"
DEFAULT_CSV = f"formatting_issues_{language_name}.csv"
# ==========================================

COLOR_MAP = {
    "red": [
        RGBColor.from_string("FF0000"), RGBColor.from_string("8B0000"),
        RGBColor.from_string("DC143C"), RGBColor.from_string("B22222"),
        RGBColor.from_string("CD5C5C"), RGBColor.from_string("F08080"),
        RGBColor.from_string("FA8072"), RGBColor.from_string("E9967A"),
        RGBColor.from_string("FFA07A"), RGBColor.from_string("FF6347"),
        RGBColor.from_string("FF4500"), RGBColor.from_string("800000"),
        RGBColor.from_string("800020"), RGBColor.from_string("7C0A02"),
        RGBColor.from_string("FF0800"), RGBColor.from_string("FF2400"),
        RGBColor.from_string("E23D28"), RGBColor.from_string("AA4A44"),
        RGBColor.from_string("C41E3A"), RGBColor.from_string("D70040"),
        RGBColor.from_string("D2042D"), RGBColor.from_string("F88379"),
        RGBColor.from_string("EE0000")
    ],
    "green": [
        RGBColor.from_string("008000"), RGBColor.from_string("00FF00"),
        RGBColor.from_string("228B22"), RGBColor.from_string("355E3B"),
        RGBColor.from_string("00A36C"), RGBColor.from_string("2AAA8A"),
        RGBColor.from_string("4CBB17"), RGBColor.from_string("50C878"),
        RGBColor.from_string("023020"), RGBColor.from_string("DFFF00"),
        RGBColor.from_string("808000"), RGBColor.from_string("088F8F"),
    ]
}

PUNCTUATION = ".?!;"


def check_paragraph(i, para, color_name, colors_to_check):
    """Check one paragraph for capitalization, punctuation, and color."""
    issues = []
    text = para.text.strip()
    if not text:
        return issues

    # Color check
    if colors_to_check:
        for run in para.runs:
            if run.font.color and run.font.color.rgb in colors_to_check:
                issues.append({
                    "line": i + 1,
                    "type": f"Contains '{color_name}' text",
                    "sentence": text,
                    "suggestion": f"Remove {color_name} formatting or standardize color."
                })
                break

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

    return issues


def main():
    """Main entry point with hybrid argument handling."""
    parser = argparse.ArgumentParser(description="Check a .docx file for formatting issues.")
    parser.add_argument("filepath", nargs="?", default=DEFAULT_FILE, help="Path to the .docx file")
    parser.add_argument("--color", help="Color to check for (e.g., 'red', 'green').")
    parser.add_argument("--output", help="Output file for text summary.")
    parser.add_argument("--csv", help="Output CSV file name.")
    args = parser.parse_args()

    # Fallback to defaults if arguments not provided
    filepath = args.filepath
    color = args.color or DEFAULT_COLOR
    output_txt = args.output or DEFAULT_OUTPUT
    output_csv = args.csv or DEFAULT_CSV

    print(f"\n📘 Using settings:")
    print(f"  File:   {filepath}")
    print(f"  Color:  {color}")
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

    colors_to_check = []
    if color and color.lower() in COLOR_MAP:
        colors_to_check = COLOR_MAP[color.lower()]
    elif color:
        print(f"⚠️ Unsupported color '{color}'. Choose from {list(COLOR_MAP.keys())}")
        return

    all_issues = []
    for i, para in enumerate(doc.paragraphs):
        all_issues.extend(check_paragraph(i, para, color, colors_to_check))

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
