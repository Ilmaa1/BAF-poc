#!/usr/bin/env python3
"""
Batch PDF processing script.

Reads PDFs from the input folder, appends extracted data to DQ_Anomaly_V2.xlsx,
and deletes each input file after successful processing.

Usage:
    python process_pdfs.py [--input-dir INPUT] [--output-dir OUTPUT]

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


def process_pdf(pdf_path: Path) -> dict | None:
    """Process a single PDF and return extracted_fields dict, or None on failure."""
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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Process PDFs from input folder, append to DQ_Anomaly_V2.xlsx, delete input after processing."
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=INPUT_DIR,
        help=f"Input folder containing PDFs (default: {INPUT_DIR})",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT_DIR,
        help=f"Output folder for Excel (default: {OUTPUT_DIR})",
    )
    args = parser.parse_args()

    input_dir = args.input_dir.resolve()
    output_dir = args.output_dir.resolve()

    if not input_dir.exists():
        logger.error("Input directory does not exist: %s", input_dir)
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        logger.warning("No PDF files found in %s", input_dir)
        return 0

    logger.info("Found %d PDF(s) in %s", len(pdf_files), input_dir)

    start = time.perf_counter()
    rows: list[dict] = []
    for pdf_path in pdf_files:
        extracted = process_pdf(pdf_path)
        if extracted is not None:
            rows.append(extracted)
            try:
                pdf_path.unlink()
                logger.info("Deleted input: %s", pdf_path.name)
            except OSError:
                logger.exception("Could not delete %s", pdf_path)

    if rows:
        output_path = output_dir / OUTPUT_EXCEL_NAME
        append_rows_to_excel(output_path, rows)

    elapsed = time.perf_counter() - start
    logger.info(
        "Completed: %d/%d succeeded in %.1fs. Output: %s",
        len(rows),
        len(pdf_files),
        elapsed,
        output_dir / OUTPUT_EXCEL_NAME,
    )

    return 0 if len(rows) == len(pdf_files) else 1


if __name__ == "__main__":
    sys.exit(main())
