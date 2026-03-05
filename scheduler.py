#!/usr/bin/env python3
"""
Scheduler that monitors the input folder for new PDFs, processes them,
appends to a single Excel file (DQ_Anomaly_V2.xlsx), and deletes input files after processing.

Usage:
    python scheduler.py [--input-dir INPUT] [--output-dir OUTPUT] [--poll-interval SECS]

Environment:
    OPENAI_API_KEY  - Required for extraction
    OPENAI_MODEL    - Model to use (default: gpt-5.2)
"""

import argparse
import logging
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from app.config import INPUT_DIR, OUTPUT_DIR
from app.services.excel_service import OUTPUT_EXCEL_NAME, append_rows_to_excel
from app.services.extraction_service import extract_fields_from_images
from app.services.pdf_service import pdf_to_base64_images

logger = logging.getLogger(__name__)

# How often to check for files needing processing
DEFAULT_POLL_INTERVAL = 5


def process_pdf(pdf_path: Path) -> dict | None:
    """
    Process a single PDF and return extracted_fields dict, or None on failure.
    """
    logger.info("Processing: %s", pdf_path.name)

    try:
        images = pdf_to_base64_images(pdf_path)
        if not images:
            logger.warning("No pages rendered from %s", pdf_path.name)
            return None

        extraction_result = extract_fields_from_images(images)
        return extraction_result["extracted_fields"]

    except Exception:
        logger.exception("Failed to process %s", pdf_path.name)
        return None


def process_input_folder(input_dir: Path, output_dir: Path) -> None:
    """
    Process all PDFs in input folder: extract, append to DQ_Anomaly_V2.xlsx, delete each PDF after processing.
    """
    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        return

    logger.info("Found %d PDF(s) in %s", len(pdf_files), input_dir)

    rows: list[dict] = []
    for pdf_path in pdf_files:
        extracted = process_pdf(pdf_path)
        if extracted is not None:
            rows.append(extracted)
            # Delete input file as soon as we've picked and processed it
            try:
                pdf_path.unlink()
                logger.info("Deleted input: %s", pdf_path.name)
            except OSError:
                logger.exception("Could not delete %s", pdf_path)

    if rows:
        output_path = output_dir / OUTPUT_EXCEL_NAME
        append_rows_to_excel(output_path, rows)
        logger.info("Appended %d row(s) to %s", len(rows), output_path)


def run_polling_loop(input_dir: Path, output_dir: Path, interval: int) -> None:
    """Poll input folder periodically and process any PDFs found."""
    logger.info(
        "Scheduler started. Watching %s every %ds. Output: %s",
        input_dir,
        interval,
        output_dir,
    )

    while True:
        try:
            process_input_folder(input_dir, output_dir)
        except Exception:
            logger.exception("Error in scheduler loop")

        time.sleep(interval)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Monitor input folder for PDFs, append to DQ_Anomaly_V2.xlsx, delete input after processing."
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=INPUT_DIR,
        help=f"Input folder to watch (default: {INPUT_DIR})",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT_DIR,
        help=f"Output folder for Excel (default: {OUTPUT_DIR})",
    )
    parser.add_argument(
        "--poll-interval",
        type=int,
        default=DEFAULT_POLL_INTERVAL,
        help=f"Seconds between folder checks (default: {DEFAULT_POLL_INTERVAL})",
    )
    args = parser.parse_args()

    input_dir = args.input_dir.resolve()
    output_dir = args.output_dir.resolve()

    if not input_dir.exists():
        logger.error("Input directory does not exist: %s", input_dir)
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    # Process any existing PDFs on startup
    process_input_folder(input_dir, output_dir)

    run_polling_loop(input_dir, output_dir, args.poll_interval)
    return 0  # unreachable


if __name__ == "__main__":
    sys.exit(main())
