#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Amazon ISBN Scraper

This script reads the Resulam books database from an Excel file, scrapes the
Amazon product pages linked in the 'paperback', 'ebook', and 'hard_cover'
columns to find their ISBN-13s, and saves the updated database to a new
Excel file.
"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
from pathlib import Path
import time
import re

# =============================================================================
# Configuration
# =============================================================================

# Path to the original books database Excel file.
INPUT_DATABASE_PATH = r"G:\My Drive\Mbú'ŋwɑ̀'nì\RoyaltiesResulam\Resulam_books_database_Amazon_base_de_donnee_livres.xlsx"

# Name for the new Excel file with the added ISBNs.
OUTPUT_DATABASE_NAME = "Resulam_books_database_with_ISBNS.xlsx"

# Headers to mimic a browser visit.
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    'Accept-Language': 'en-US,en;q=0.9',
}

# =============================================================================
# Core Functions
# =============================================================================

def get_isbn_from_url(url: str) -> str:
    """
    Scrapes an Amazon product page to find the ISBN-13.

    Args:
        url: The URL of the Amazon product page.

    Returns:
        The ISBN-13 as a string, or an empty string if not found or on error.
    """
    if not isinstance(url, str) or not url.startswith('http'):
        return ""

    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()  # Raise an exception for bad status codes

        soup = BeautifulSoup(response.content, 'html.parser')

        # --- Find ISBN-13 ---
        # Amazon pages can have different structures. We'll try a few common patterns.
        
        # Pattern 1: Look for a span that contains "ISBN-13" and get the next element.
        isbn_element = soup.find('span', string=lambda t: t and "ISBN-13" in t)
        if isbn_element:
            # The actual ISBN might be in the next span or div
            isbn_value = isbn_element.find_next_sibling(['span', 'div'])
            if isbn_value:
                return isbn_value.text.strip()

        # Pattern 2: Look for list items in the product details section.
        detail_bullets = soup.find('div', {'id': 'detailBullets_feature_div'})
        if detail_bullets:
            for li in detail_bullets.find_all('li'):
                if 'ISBN-13' in li.text:
                    # The text is usually like "ISBN-13 : 979-8466614336"
                    parts = li.text.split(':')
                    if len(parts) > 1:
                        return parts[1].strip()
        
        print(f"    - ISBN not found on page: {url}")
        return ""

    except requests.exceptions.RequestException as e:
        print(f"    - Could not fetch URL {url}: {e}")
        return ""
    except Exception as e:
        print(f"    - An unexpected error occurred while scraping {url}: {e}")
        return ""


def get_asin_from_url(url: str) -> str:
    """
    Try to extract the ASIN from an Amazon product URL or page.

    Returns ASIN string or empty string if not found.
    """
    if not isinstance(url, str) or not url.startswith('http'):
        return ""

    # Try to parse ASIN directly from common URL patterns (/dp/<ASIN>, /gp/product/<ASIN>)
    try:
        for token in ['/dp/', '/gp/product/', '/product/']:
            if token in url:
                parts = url.split(token, 1)[1]
                # ASIN is the first path segment
                asin = parts.split('/')[0].split('?')[0].strip()
                # Clean trailing punctuation or spaces
                asin = re.sub(r"[^A-Za-z0-9]", "", asin)
                if asin:
                    return asin
    except Exception:
        pass

    # If not in URL, fetch page and look for ASIN in product details or meta tags
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        # Check detail bullets
        detail_bullets = soup.find('div', {'id': 'detailBullets_feature_div'})
        if detail_bullets:
            for li in detail_bullets.find_all('li'):
                if 'ASIN' in li.text:
                    parts = li.text.split(':')
                    if len(parts) > 1:
                        return clean_identifier(parts[1].strip())

        # Check product details table
        prod_details = soup.find(id='productDetails_detailBullets_sections1')
        if prod_details:
            text = prod_details.get_text(separator=' ')
            m = re.search(r'ASIN\s*[:\u200E\u200F\s]+([A-Za-z0-9]+)', text)
            if m:
                return m.group(1)

        # Check canonical or og:url meta tag
        canonical = soup.find('link', {'rel': 'canonical'})
        if canonical and canonical.get('href'):
            href = canonical['href']
            for token in ['/dp/', '/gp/product/']:
                if token in href:
                    asin = href.split(token, 1)[1].split('/')[0].split('?')[0]
                    asin = re.sub(r"[^A-Za-z0-9]", "", asin)
                    if asin:
                        return asin

        # As a last resort, search for data-asin attribute
        data_asin = soup.find(attrs={'data-asin': True})
        if data_asin and data_asin.get('data-asin'):
            return clean_identifier(data_asin.get('data-asin'))

        return ""
    except requests.exceptions.RequestException as e:
        print(f"    - Could not fetch URL for ASIN {url}: {e}")
        return ""
    except Exception as e:
        print(f"    - Unexpected error extracting ASIN from {url}: {e}")
        return ""


def get_pages_from_url(url: str) -> str:
    """
    Scrape the product page and try to extract the number of pages (if present).

    Returns the page count as a string (digits) or an empty string if not found.
    """
    if not isinstance(url, str) or not url.startswith('http'):
        return ""

    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        # Search detail bullets for patterns like '72 pages' or 'Paperback: 72 pages'
        detail_bullets = soup.find('div', {'id': 'detailBullets_feature_div'})
        text_sources = []
        if detail_bullets:
            text_sources.append(detail_bullets.get_text(separator=' '))

        # product details sections
        prod_details = soup.find(id='productDetails_detailBullets_sections1')
        if prod_details:
            text_sources.append(prod_details.get_text(separator=' '))

        # Also check the general page text as a last resort
        text_sources.append(soup.get_text(separator=' '))

        page_regex = re.compile(r"(\d{1,5})\s+pages", flags=re.IGNORECASE)
        for txt in text_sources:
            if not txt:
                continue
            m = page_regex.search(txt)
            if m:
                # return just the digits (no commas)
                return m.group(1).replace(',', '')

        return ""
    except requests.exceptions.RequestException as e:
        print(f"    - Could not fetch URL for pages {url}: {e}")
        return ""
    except Exception as e:
        print(f"    - Unexpected error extracting pages from {url}: {e}")
        return ""


def clean_identifier(value: str) -> str:
    """
    Clean an identifier string (ISBN or ASIN) by removing
    trailing/leading whitespace, paragraph markers, bidi marks,
    zero-width chars, and normalizing internal whitespace.
    """
    if value is None:
        return ""
    if not isinstance(value, str):
        try:
            value = str(value)
        except Exception:
            return ""

    # Remove common invisible/formatting characters
    # Left-to-right mark, right-to-left mark, paragraph separators, zero-width spaces
    value = re.sub(r"[\u200E\u200F\u2028\u2029\u200B\uFEFF]", "", value)

    # Replace newlines, carriage returns, tabs with a single space
    value = re.sub(r"[\r\n\t]+", " ", value)

    # Replace non-breaking spaces with regular spaces
    value = value.replace('\u00A0', ' ')

    # Collapse multiple spaces into a single space and strip ends
    value = ' '.join(value.split()).strip()

    # Remove leading/trailing punctuation that sometimes appears
    value = value.strip(' :;\u200E\u200F')

    return value


def process_isbns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Iterates through the DataFrame, scrapes ISBNs for each book type,
    and adds them to new columns.

    Args:
        df: The original DataFrame of books.

    Returns:
        The DataFrame updated with new ISBN columns.
    """
    # Initialize new columns
    df['isbn_paperback'] = ''
    df['isbn_ebook'] = ''
    df['isbn_hard_cover'] = ''
    # Initialize pages columns
    df['pages_paperback'] = ''
    df['pages_ebook'] = ''
    df['pages_hard_cover'] = ''

    total_rows = len(df)
    
    # Define the columns to process
    url_columns = {
        'paperback': 'isbn_paperback',
        'ebook': 'isbn_ebook',
        'hard_cover': 'isbn_hard_cover'
    }

    for index, row in df.iterrows():
        print(f"Processing book {index + 1}/{total_rows}: {row['title']}")
        
        for source_col, target_col in url_columns.items():
            url = row.get(source_col)
            if pd.notna(url) and url:
                print(f"  - Scraping {source_col} URL...")
                isbn = get_isbn_from_url(url)
                # Clean the extracted identifier to remove invisible chars and whitespace
                isbn_clean = clean_identifier(isbn)
                if isbn_clean:
                    df.at[index, target_col] = isbn_clean
                    print(f"    - Found ISBN: {isbn_clean}")
                else:
                    # If this is the ebook column and no ISBN found, try ASIN fallback
                    if source_col == 'ebook':
                        asin = get_asin_from_url(url)
                        asin_clean = clean_identifier(asin)
                        if asin_clean:
                            df.at[index, target_col] = asin_clean
                            print(f"    - Found ASIN (used as ISBN_ebook): {asin_clean}")
                # scrape pages for this URL and put in corresponding pages column
                pages_col = f"pages_{source_col}"
                try:
                    pages = get_pages_from_url(url)
                    pages_clean = clean_identifier(pages)
                    if pages_clean:
                        df.at[index, pages_col] = pages_clean
                        print(f"    - Found pages ({pages_col}): {pages_clean}")
                except Exception as e:
                    print(f"    - Error extracting pages for {url}: {e}")
                # Add a small delay to avoid overwhelming the server
                time.sleep(1)
    
    return df


# =============================================================================
# Main Execution
# =============================================================================

def main():
    """
    Main function to execute the ISBN scraping script.
    """
    print("Starting Amazon ISBN Scraper...")
    print("=" * 50)

    # Load the database
    try:
        input_path = Path(INPUT_DATABASE_PATH)
        if not input_path.exists():
            print(f"Error: Input file not found at '{INPUT_DATABASE_PATH}'")
            return
            
        books_df = pd.read_excel(input_path)
        print(f"Successfully loaded {len(books_df)} book records.")
    except Exception as e:
        print(f"An error occurred while loading the Excel file: {e}")
        return

    # Process the books to find ISBNs
    updated_df = process_isbns(books_df)

    # Clean ISBN/ASIN columns again to ensure no trailing spaces or hidden
    # paragraph/formatting characters remain from scraping.
    for col in ['isbn_paperback', 'isbn_ebook', 'isbn_hard_cover']:
        if col in updated_df.columns:
            updated_df[col] = updated_df[col].fillna('').astype(str).apply(clean_identifier)

    # Save the updated database to a new file
    try:
        output_path = Path(OUTPUT_DATABASE_NAME)
        updated_df.to_excel(output_path, index=False, engine='xlsxwriter')
        print(f"\nSuccessfully saved updated database with ISBNs to '{output_path}'")
    except Exception as e:
        print(f"\nAn error occurred while saving the new Excel file: {e}")

    print("\n" + "=" * 50)
    print("Processing complete.")


if __name__ == "__main__":
    main()
