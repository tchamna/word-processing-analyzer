# Word Processing Script

A Python script to analyze .docx files for formatting issues, including extracting lines with colored text (e.g., red), checking capitalization, and punctuation. It generates both a plain text summary and a CSV report.

## Features

- Detects paragraphs with colored text (e.g., red, green, as specified by the user)
- Checks for missing capitalization
- Checks for missing punctuation
- Outputs results in both text and CSV formats
- CSV includes columns: Line, Issue Type, Sentence, Suggested Fix

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/word-processing-script.git
   cd word-processing-script
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the script with optional arguments:

```bash
python word_processing.py [filepath] [--color COLOR] [--output OUTPUT] [--csv CSV]
```

The script uses default values for all parameters if not provided. Defaults are set in the script (e.g., default file path, color, output files).

Examples:

- Use all defaults: `python word_processing.py`
- Specify a different .docx file: `python word_processing.py path/to/your/file.docx`
- Check for green text instead: `python word_processing.py --color green`
- Custom output files: `python word_processing.py --output my_issues.txt --csv my_report.csv`

The script will generate:
- `formatting_issues_{language}.txt` (or custom output file): Plain text summary of issues
- `formatting_issues_{language}.csv` (or custom CSV file): CSV report with columns: Line, Issue Type, Sentence, Suggested Fix

## Requirements

- Python 3.6+
- python-docx library

## License

MIT License