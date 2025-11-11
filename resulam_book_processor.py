#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Resulam Books Database Processor

This script processes the Resulam books database from an Excel file.
It provides functionality to filter books by language or type, format the
details for various outputs, and export the results to text files or Excel sheets.

Key Features:
    - Load book data from an Excel file.
    - Filter books by language or title keywords (book type).
    - Generate formatted text summaries for sharing (e.g., on WhatsApp).
    - Export filtered data into separate text files for all books, by type, or by language.
    - Export filtered data into a single Excel file with a sheet for each language.
    - Clean and validate data, such as sheet names for Excel export.
"""

import os
import re
import shutil
from pathlib import Path
from typing import List, Dict, Optional

import pandas as pd
import unicodedata



# =============================================================================
# Configuration
# =============================================================================

# Set the default path to the books database Excel file.
# Using a raw string (r"...") is recommended for Windows paths.
DEFAULT_BOOKS_PATH = r"G:\My Drive\Mbú'ŋwɑ̀'nì\RoyaltiesResulam\Resulam_books_database_Amazon_base_de_donnee_livres.xlsx"

# Default folder for generated output files.
DEFAULT_OUTPUT_DIR = "Book_Exports"

# Default columns to use when filtering book data.
DEFAULT_COLUMNS = [
    "language_name",
    "title",
    "book_nick_name",
    "paperback",
    "ebook",
    "hard_cover"
]


# =============================================================================
# Helper Functions
# =============================================================================

def clean_sheet_name(name: str) -> str:
    """
    Cleans a string to be a valid Excel sheet name.

    Removes special characters, strips leading/trailing apostrophes,
    and truncates the name to the 31-character limit.

    Args:
        name: The original string to clean.

    Returns:
        A sanitized string suitable for an Excel sheet name.
    """
    # Remove characters that are invalid in sheet names
    no_invalid_chars = re.sub(r'[\\/*?:[\]]', '', name)
    # Replace whitespace sequences with a single space and strip ends
    cleaned = re.sub(r'\s+', ' ', no_invalid_chars).strip()
    # Strip leading/trailing apostrophes, which are invalid at the ends
    cleaned = cleaned.strip("'")
    # Truncate to 31 characters
    return cleaned[:31]


# =============================================================================
# Core Book Processor Class
# =============================================================================

class BookDatabaseProcessor:
    """
    Handles loading, processing, and exporting of the Resulam books database.
    """

    def __init__(self, books_path: str, output_dir: str = DEFAULT_OUTPUT_DIR):
        """
        Initializes the BookDatabaseProcessor.

        Args:
            books_path: The file path to the book database Excel file.
            output_dir: The directory where output files will be saved.
        """
        self.books_path = books_path
        self.output_dir = Path(output_dir)
        self.books_df: Optional[pd.DataFrame] = None
        self.languages: List[str] = []
        self._load_database()

    def _load_database(self) -> None:
        """
        Loads the book database from the specified Excel file.
        Sorts the data by publication date and identifies unique languages.
        """
        print(f"Loading book database from: {self.books_path}")
        try:
            self.books_df = pd.read_excel(self.books_path)
            # Sort by date, most recent first
            self.books_df = self.books_df.sort_values(by="publication_date", ascending=False)
            # Sort languages alphabetically while ignoring diacritics so that
            # names like 'Éwé' appear near 'Ewondo' rather than at the end.
            raw_langs = [str(x) for x in self.books_df["language_name"].dropna().unique()]
            def _strip_accents(s: str) -> str:
                return unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII')

            self.languages = sorted(raw_langs, key=lambda s: _strip_accents(s).lower())
            print(f"Successfully loaded {len(self.books_df)} book records for {len(self.languages)} languages.")
        except FileNotFoundError:
            print(f"Error: The file was not found at '{self.books_path}'.")
            self.books_df = pd.DataFrame() # Ensure books_df is not None
        except Exception as e:
            print(f"An unexpected error occurred while loading the Excel file: {e}")
            self.books_df = pd.DataFrame()

    def get_books(
        self,
        language: Optional[str] = None,
        book_type: Optional[str] = None,
        columns: Optional[List[str]] = None
    ) -> pd.DataFrame:
        """
        Filters and returns books based on language and/or book type.

        Args:
            language: The language to filter by (e.g., "Nufi").
            book_type: A keyword to search for in the book titles (e.g., "conte").
            columns: A list of column names to include in the result.

        Returns:
            A pandas DataFrame containing the filtered book data.
        """
        if self.books_df is None or self.books_df.empty:
            return pd.DataFrame()

        df = self.books_df.copy()

        if language:
            if language not in self.languages:
                print(f"Warning: Language '{language}' not found in the database.")
                return pd.DataFrame()
            df = df[df["language_name"] == language].copy()

        if book_type:
            # Case-insensitive search for the book_type in the title
            mask = df["title"].str.contains(book_type, case=False, na=False)
            df = df[mask]

        if columns:
            # Filter to only include requested columns that actually exist
            existing_columns = [col for col in columns if col in df.columns]
            df = df[existing_columns]

        return df

    def format_for_whatsapp(self, books_df: pd.DataFrame) -> str:
        """
        Formats a DataFrame of book details into a string for WhatsApp.

        Args:
            books_df: DataFrame containing the books to format.

        Returns:
            A formatted string with all book details.
        """
        if books_df.empty:
            return "No books found matching the criteria."

        formatted_text_all = []
        for _, row in books_df.iterrows():
            output_dict = {
                "Book Name": row.get("title"),
                "Ebook": row.get("ebook"),
                "Soft Cover": row.get("paperback"),
                "Hard Cover": row.get("hard_cover")
            }

            # Filter out entries with no link or value
            valid_entries = {k: v for k, v in output_dict.items() if pd.notna(v) and v}

            if valid_entries:
                # Format as *Key*: Value
                formatted_text = "\n".join([f"*{key}*: {value}" for key, value in valid_entries.items()])
                formatted_text_all.append(formatted_text)

        return ("\n" + "*" * 35 + "\n").join(formatted_text_all)

    def export_all_books_to_txt(self) -> None:
        """
        Exports all books, formatted for WhatsApp, to a single text file.
        """
        print("Exporting all books to a single text file...")
        if self.books_df is None or self.books_df.empty:
            print("Cannot export: Book data is not loaded.")
            return

        self.output_dir.mkdir(parents=True, exist_ok=True)
        file_path = self.output_dir / "01_all_resulam_books.txt"

        all_books_df = self.get_books(columns=DEFAULT_COLUMNS)
        formatted_output = self.format_for_whatsapp(all_books_df)
        
        header = "*Resulam's Books Publication from 2015*\n" + "*" * 35
        final_content = header + "\n" + formatted_output

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(final_content)
        print(f"Successfully saved all books to '{file_path}'")

    def export_books_by_type_to_txt(self, book_type: str) -> None:
        """
        Exports books of a specific type to a text file.

        Args:
            book_type: The keyword to filter book titles by.
        """
        print(f"Exporting books of type '{book_type}' to text file...")
        if self.books_df is None or self.books_df.empty:
            print("Cannot export: Book data is not loaded.")
            return

        self.output_dir.mkdir(parents=True, exist_ok=True)
        file_path = self.output_dir / f"02_books_of_type_{book_type.lower()}.txt"

        books_of_type_df = self.get_books(book_type=book_type, columns=DEFAULT_COLUMNS)
        if books_of_type_df.empty:
            print(f"No books found for type '{book_type}'. Nothing to export.")
            return

        formatted_output = self.format_for_whatsapp(books_of_type_df)
        header = f"*Resulam's '{book_type.title()}' Books*\n" + "*" * 35
        final_content = header + "\n" + formatted_output

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(final_content)
        print(f"Successfully saved '{book_type}' books to '{file_path}'")

    def export_books_by_language_to_txts(self) -> None:
        """
        Exports books for each language to its own separate text file.
        """
        print("Exporting books for each language to separate text files...")
        if self.books_df is None or self.books_df.empty:
            print("Cannot export: Book data is not loaded.")
            return

        lang_dir = self.output_dir / "per_language_txt"
        lang_dir.mkdir(parents=True, exist_ok=True)

        for lang in self.languages:
            file_path = lang_dir / f"{clean_sheet_name(lang)}_books.txt"
            lang_df = self.get_books(language=lang, columns=DEFAULT_COLUMNS)

            if not lang_df.empty:
                formatted_output = self.format_for_whatsapp(lang_df)
                header = f"*Resulam's Books in {lang}*\n" + "*" * 35
                final_content = header + "\n" + formatted_output
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(final_content)
        print(f"Successfully saved book lists for {len(self.languages)} languages in '{lang_dir}'")

    def export_books_by_language_to_excel(self) -> None:
        """
        Exports all book data to a single Excel file, with each language
        in its own sheet.
        """
        print("Exporting books for each language to an Excel file...")
        if self.books_df is None or self.books_df.empty:
            print("Cannot export: Book data is not loaded.")
            return

        self.output_dir.mkdir(parents=True, exist_ok=True)
        file_path = self.output_dir / "03_all_books_by_language.xlsx"

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            for lang in self.languages:
                sheetname = clean_sheet_name(lang)
                lang_df = self.get_books(language=lang) # Get all columns for Excel
                lang_df.to_excel(writer, sheet_name=sheetname, index=False)

        print(f"Successfully saved all books to '{file_path}', with each language on a separate sheet.")

    def _format_publication_summary_entry(self, row: pd.Series) -> str:
        """
        Formats a single book's metadata into a detailed summary string.

        Args:
            row: A pandas Series representing a single row (book).

        Returns:
            A formatted string with the book's publication details.
        """
        authors = row.get('authors', 'N/A')
        title = row.get('title', 'N/A')

        # Prefer pages from format-specific scraped columns if available
        pages_candidates = ['pages_paperback', 'pages_hard_cover', 'pages_ebook', 'number_of_pages']
        pages_val = None
        for pc in pages_candidates:
            if pc in row and pd.notna(row.get(pc)) and row.get(pc) != "":
                pages_val = row.get(pc)
                break

        # Normalize pages to an integer string when possible
        pages_str = ""
        if pages_val is not None:
            try:
                # pages_val might be numeric or a string containing digits
                pages_int = int(str(pages_val).strip())
                pages_str = f"({pages_int} pages), "
            except Exception:
                # Fallback: try to extract leading digits
                m = re.search(r"(\d+)", str(pages_val))
                if m:
                    pages_str = f"({int(m.group(1))} pages), "

        # Try common columns that might contain the paperback ISBN (or the scraped isbn_paperback)
        isbn_candidates = [
            'paperback_isbn13',
            'isbn_paperback',
            'paperback_isbn',
            'paperback_isbn10',
            'paperback_isbn-13',
        ]
        isbn_val = None
        for col in isbn_candidates:
            if col in row and pd.notna(row.get(col)) and row.get(col) != 0:
                isbn_val = row.get(col)
                break
        isbn_str = f"ISBN-13: {str(isbn_val).strip()}, " if isbn_val else ""

        pub_date = row.get('publication_date')
        # The date should be a datetime object from _load_database
        date_str = pd.to_datetime(pub_date).strftime('%B %d, %Y') if pd.notna(pub_date) else "N/A"

        link = row.get('paperback')
        link_str = f"Link (Printed): {link}" if pd.notna(link) else ""

        # Combine all parts into the final string
        return f'{authors}, "{title}", {pages_str}{isbn_str}{date_str}. {link_str}'

    def export_publication_summary(self) -> None:
        """
        Generates and exports a formatted summary of all books, categorized by
        language and sorted by publication date.
        """
        print("Exporting formatted publication summary...")
        if self.books_df is None or self.books_df.empty:
            print("Cannot export: Book data is not loaded.")
            return

        self.output_dir.mkdir(parents=True, exist_ok=True)
        file_path = self.output_dir / "books_publication_summary.txt"
        
        full_summary = []

        for lang in self.languages:
            full_summary.append(f"\n{lang}\n" + "-" * len(lang))
            
            # The main DataFrame is already sorted by publication_date
            lang_df = self.get_books(language=lang)
            
            # Format each book entry
            for i, (_, row) in enumerate(lang_df.iterrows(), 1):
                formatted_entry = self._format_publication_summary_entry(row)
                full_summary.append(f"{i}-\t{formatted_entry}")
        
        final_content = "\n".join(full_summary)
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(final_content)
        print(f"Successfully saved publication summary to '{file_path}'")


# =============================================================================
# Main Execution
# =============================================================================

def main():
    """
    Main function to run the book processing script.
    """
    print("Starting Resulam Book Database Processor...")
    print("=" * 50)

    # Initialize the processor with the path to the database.
    processor = BookDatabaseProcessor(books_path=DEFAULT_BOOKS_PATH)

    # Exit if the database could not be loaded.
    if processor.books_df is None or processor.books_df.empty:
        print("\nProcessing stopped due to database loading issues.")
        return

    # --- Perform Export Tasks ---
    # You can comment out or enable the tasks you want to run.

    # 1. Export a single text file containing all books.
    processor.export_all_books_to_txt()

    # 2. Export a text file for a specific type of book (e.g., "conte", "grenier").
    processor.export_books_by_type_to_txt(book_type="conte")
    processor.export_books_by_type_to_txt(book_type="grenier")

    # 3. Export a separate text file for each language.
    processor.export_books_by_language_to_txts()

    # 4. Export a single Excel file with a sheet for each language.
    processor.export_books_by_language_to_excel()

    # 5. Generate and export the formatted publication summary.
    processor.export_publication_summary()

    print("\n" + "=" * 50)
    print("Processing complete. All files saved in the 'Book_Exports' directory.")


if __name__ == "__main__":
    main()