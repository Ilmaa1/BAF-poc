import base64
from io import BytesIO
import logging
from pathlib import Path
import time

import pypdfium2 as pdfium

logger = logging.getLogger(__name__)


def pdf_to_base64_images(pdf_path: Path, scale: float = 2.0) -> list[dict]:
    start = time.perf_counter()
    pdf = pdfium.PdfDocument(str(pdf_path))
    images: list[dict] = []
    total_pages = len(pdf)
    logger.info(
        "Starting PDF render. path=%s total_pages=%s scale=%s",
        pdf_path,
        total_pages,
        scale,
    )

    try:
        for index, page in enumerate(pdf):
            page_start = time.perf_counter()
            pil_image = page.render(scale=scale).to_pil()
            buffer = BytesIO()
            pil_image.save(buffer, format="PNG")
            images.append(
                {
                    "page": index + 1,
                    "mime_type": "image/png",
                    "image_base64": base64.b64encode(buffer.getvalue()).decode("utf-8"),
                }
            )
            logger.debug(
                "Rendered page to base64 image. page=%s elapsed_ms=%s",
                index + 1,
                int((time.perf_counter() - page_start) * 1000),
            )
    finally:
        pdf.close()

    logger.info(
        "Completed PDF render. path=%s pages_rendered=%s elapsed_ms=%s",
        pdf_path,
        len(images),
        int((time.perf_counter() - start) * 1000),
    )
    return images
