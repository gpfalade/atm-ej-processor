"""
==============================================================
  Postilion ATM/POS Journal Processor
  journal_split.py
==============================================================
  Processes raw Postilion ATM journal .dat files from the
  input_folder and its subfolders.

  What it does:
  1. Detects and converts .dat files from any encoding to UTF-8
  2. Splits each file into individual transaction sections
  3. Validates each section — only processes transactions
     with a non-zero RRN (successful/attempted transactions)
  4. Extracts card BIN, RRN, date and time from PostilionTran tag
  5. Organises output PDFs into subfolders by card BIN
  6. Renames any unresolved PDFs on a second pass
  7. Writes a full audit log to the logs/ folder

  Folder structure expected:
    input_folder/   <- place .dat files here (subfolders supported)
    output_folder/  <- PDFs created here automatically
    logs/           <- log file written here automatically

  Usage:
    python journal_split.py
    OR double-click run_journal_split.bat (Windows)
==============================================================
"""

import os
import re
import logging
from datetime import datetime
from fpdf import FPDF
from PyPDF2 import PdfReader
import chardet

# FOLDER PATHS 

INPUT_FOLDER  = "input_folder"
OUTPUT_FOLDER = "output_folder"
LOGS_FOLDER   = "logs"

# LOGGING SETUP 

os.makedirs(LOGS_FOLDER,   exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

run_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_filename  = os.path.join(LOGS_FOLDER, f"journal_split_{run_timestamp}.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

#  COUNTERS 

stats = {
    "dat_files_found":    0,
    "dat_files_converted":0,
    "txt_files_processed":0,
    "sections_found":     0,
    "sections_skipped":   0,
    "pdfs_created":       0,
    "pdfs_renamed":       0,
    "errors":             0,
}

# UTILITIES 

def sanitize_filename(filename):
    """Remove invalid characters from filenames."""
    invalid_chars_pattern = r'[<>:"/\\|?*\n\r]'
    return re.sub(invalid_chars_pattern, '_', filename)


def clean_text(text):
    """Remove non-ASCII characters for PDF compatibility."""
    return re.sub(r'[^\x00-\x7F]+', '', text)


# PROGRAM 1: CONVERT .dat TO UTF-8 .txt 

def convert_dat_to_utf8(input_folder):
    """
    Walk input_folder recursively, find all .dat files,
    detect encoding and convert to UTF-8 .txt files.
    """
    log.info("=" * 60)
    log.info("STAGE 1 — DAT FILE CONVERSION")
    log.info("=" * 60)

    for root, dirs, files in os.walk(input_folder):
        for file_name in files:
            if file_name.endswith(".dat"):
                stats["dat_files_found"] += 1
                input_path  = os.path.join(root, file_name)
                output_path = os.path.join(
                    root, os.path.splitext(file_name)[0] + "_utf8.txt"
                )
                log.info(f"Found .dat file: {input_path}")

                # Detect encoding
                with open(input_path, 'rb') as f:
                    raw_data = f.read()
                result   = chardet.detect(raw_data)
                encoding = result.get('encoding') or 'utf-8'
                log.info(f"  Detected encoding: {encoding} "
                         f"(confidence: {result.get('confidence', 0):.0%})")

                try:
                    with open(input_path, 'r', encoding=encoding) as f:
                        content = f.read()
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    stats["dat_files_converted"] += 1
                    log.info(f"  Converted → {output_path}")
                except UnicodeDecodeError as e:
                    stats["errors"] += 1
                    log.error(f"  Encoding error on {input_path}: {e}")
                except Exception as e:
                    stats["errors"] += 1
                    log.error(f"  Unexpected error on {input_path}: {e}")

    log.info(f"Conversion complete: {stats['dat_files_converted']} of "
             f"{stats['dat_files_found']} files converted.")


#  PROGRAM 2: PROCESS .txt FILES INTO PDFs 

def process_all_txt_files(input_folder, output_folder):
    """
    Walk input_folder recursively, find all .txt files produced
    from .dat conversion, split into transaction sections and
    save each valid transaction as a PDF.
    """
    log.info("")
    log.info("=" * 60)
    log.info("STAGE 2 — TRANSACTION EXTRACTION AND PDF GENERATION")
    log.info("=" * 60)

    for root, dirs, files in os.walk(input_folder):
        for file_name in files:
            if file_name.endswith(".txt"):
                file_path = os.path.join(root, file_name)
                log.info(f"Processing: {file_path}")
                process_txt_file(file_path, output_folder)

    log.info(f"Extraction complete: {stats['pdfs_created']} PDFs created, "
             f"{stats['sections_skipped']} sections skipped (zero RRN).")


def process_txt_file(file_path, output_folder):
    """Split a single .txt file and process each transaction section."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        stats["errors"] += 1
        log.error(f"  Cannot read {file_path}: {e}")
        return

    sections = content.split("------------------")
    stats["sections_found"] += len(sections)
    log.info(f"  Sections found: {len(sections)}")

    valid = 0
    skipped = 0
    for i, text in enumerate(sections):
        if has_valid_rrn(text):
            unique_name  = extract_unique_name(text)
            unique_name  = sanitize_filename(unique_name)
            cleaned_text = clean_text(text.strip())
            save_as_pdf(cleaned_text, unique_name, output_folder)
            valid   += 1
        else:
            skipped += 1
            stats["sections_skipped"] += 1

    log.info(f"  Valid (saved): {valid} | Skipped (zero RRN): {skipped}")


def has_valid_rrn(split_text):
    """Return True only if PostilionTran tag contains a non-zero RRN."""
    matches = re.findall(
        r'<PostilionTran>(.*?)</PostilionTran>', split_text, re.DOTALL
    )
    for match in matches:
        fields = match.split("|")
        if len(fields) > 11 and fields[11] != "000000000000":
            return True
    return False


def extract_unique_name(split_text):
    """
    Extract a unique filename from the PostilionTran tag.
    Format: BIN_RRN_Date_Time
    """
    matches = re.findall(
        r'<PostilionTran>(.*?)</PostilionTran>', split_text, re.DOTALL
    )
    for match in matches:
        fields = match.split("|")
        if len(fields) > 11:
            rrn       = fields[11]
            date_time = fields[9]
            time_     = fields[10]
            number    = re.search(r'(\d+)\*\*', match)

            if rrn == "000000000000":
                rrn = find_valid_rrn(matches)

            if number:
                return f"{number.group(1)}_{rrn}_{date_time}_{time_}"

    return generate_fallback_name(split_text)


def find_valid_rrn(postilion_sections):
    """Find the first non-zero RRN across all PostilionTran sections."""
    for section in postilion_sections:
        fields = section.split("|")
        if len(fields) > 11 and fields[11] != "000000000000":
            return fields[11]
    return "000000000000"


def generate_fallback_name(split_text):
    """Generate a fallback filename when PostilionTran parsing fails."""
    match = re.search(r'<PostilionTran>(\d+)\*\*', split_text)
    if match:
        return match.group(1)
    return "UnknownTransaction"


def save_as_pdf(split_text, unique_name, output_folder):
    """
    Save a transaction section as a PDF file.
    Organises into subfolders by card BIN (first 6 digits).
    """
    name_parts = unique_name.split('_')
    bin_number = name_parts[0] if name_parts else "UnknownBIN"
    pdf_name   = '_'.join(name_parts[1:]) if len(name_parts) > 1 else "UnnamedTransaction"

    # Create BIN subfolder
    subfolder = os.path.join(output_folder, bin_number)
    os.makedirs(subfolder, exist_ok=True)

    # Handle duplicate filenames
    base_path = os.path.join(subfolder, f"{bin_number}_{pdf_name}.pdf")
    output_path = base_path
    index = 1
    while os.path.exists(output_path):
        output_path = os.path.join(subfolder, f"{bin_number}_{pdf_name}_{index}.pdf")
        index += 1

    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Courier", size=10)
        pdf.multi_cell(0, 6, txt=split_text)
        pdf.output(output_path)
        stats["pdfs_created"] += 1
        log.debug(f"  PDF saved: {output_path}")
    except Exception as e:
        stats["errors"] += 1
        log.error(f"  PDF creation failed for {output_path}: {e}")


# PROGRAM 3: RENAME UNNAMED PDFs

def rename_unresolved_pdfs(output_folder):
    """
    Second pass — rename any PDF files still named
    'UnnamedTransaction' by extracting the PostilionTran
    tag from inside the PDF content.
    """
    log.info("")
    log.info("=" * 60)
    log.info("STAGE 3 — PDF RENAME (UNRESOLVED TRANSACTIONS)")
    log.info("=" * 60)

    renamed = 0
    for root, dirs, files in os.walk(output_folder):
        for file_name in files:
            if "UnnamedTransaction" in file_name and file_name.endswith(".pdf"):
                file_path = os.path.join(root, file_name)
                log.info(f"  Attempting rename: {file_path}")

                try:
                    with open(file_path, 'rb') as f:
                        reader   = PdfReader(f)
                        pdf_text = "".join(
                            page.extract_text() or "" for page in reader.pages
                        )
                    new_name = extract_postilion_name_from_text(pdf_text)

                    if new_name:
                        new_path = os.path.join(root, f"{new_name}.pdf")
                        index = 1
                        while os.path.exists(new_path):
                            new_path = os.path.join(root, f"{new_name}_{index}.pdf")
                            index += 1
                        os.rename(file_path, new_path)
                        stats["pdfs_renamed"] += 1
                        renamed += 1
                        log.info(f"  Renamed → {new_path}")
                    else:
                        log.warning(f"  No PostilionTran tag found — file kept as-is")
                except PermissionError as e:
                    stats["errors"] += 1
                    log.error(f"  Permission error renaming {file_path}: {e}")
                except Exception as e:
                    stats["errors"] += 1
                    log.error(f"  Error processing {file_path}: {e}")

    log.info(f"Rename complete: {renamed} PDFs renamed.")


def extract_postilion_name_from_text(pdf_text):
    """Extract a unique name from PostilionTran tag content inside a PDF."""
    pattern = re.compile(
        r'<PostilionTran>(.*?)</PostilionTran>', re.DOTALL
    )
    matches = pattern.findall(pdf_text)
    for match in matches:
        fields = match.split("|")
        if len(fields) >= 12:
            rrn    = fields[11]
            dt     = fields[9]
            tm     = fields[10]
            number = fields[0][:6]
            if rrn != "000000000000":
                return sanitize_filename(f"{number}_{rrn}_{dt}_{tm}")
    return None


# MAIN

if __name__ == "__main__":
    start_time = datetime.now()

    log.info("=" * 60)
    log.info("  POSTILION JOURNAL PROCESSOR — STARTED")
    log.info(f"  Run ID   : {run_timestamp}")
    log.info(f"  Input    : {os.path.abspath(INPUT_FOLDER)}")
    log.info(f"  Output   : {os.path.abspath(OUTPUT_FOLDER)}")
    log.info(f"  Log file : {os.path.abspath(log_filename)}")
    log.info("=" * 60)

    convert_dat_to_utf8(INPUT_FOLDER)
    process_all_txt_files(INPUT_FOLDER, OUTPUT_FOLDER)
    rename_unresolved_pdfs(OUTPUT_FOLDER)

    elapsed = (datetime.now() - start_time).seconds
    log.info("")
    log.info("=" * 60)
    log.info("  PROCESSING SUMMARY")
    log.info("=" * 60)
    log.info(f"  .dat files found     : {stats['dat_files_found']}")
    log.info(f"  .dat files converted : {stats['dat_files_converted']}")
    log.info(f"  Sections found       : {stats['sections_found']}")
    log.info(f"  Sections skipped     : {stats['sections_skipped']} (zero RRN)")
    log.info(f"  PDFs created         : {stats['pdfs_created']}")
    log.info(f"  PDFs renamed         : {stats['pdfs_renamed']}")
    log.info(f"  Errors               : {stats['errors']}")
    log.info(f"  Runtime              : {elapsed} seconds")
    log.info(f"  Log saved to         : {log_filename}")
    log.info("=" * 60)
    log.info("  DONE — Run excel_check.py to reconcile against Excel ledger")
    log.info("=" * 60)
