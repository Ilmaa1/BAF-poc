import base64
from io import BytesIO
import logging
from pathlib import Path
import time

import pypdfium2 as pdfium
from PIL import ImageEnhance, ImageFilter, Image

logger = logging.getLogger(__name__)

# Render scale: higher = more pixels for the model to read handwriting from
_RENDER_SCALE = 3.0

# Preprocessing factors applied to every page image before sending to the model.
# These improve OCR accuracy on handwritten text and phone photos of ID cards.
_CONTRAST_FACTOR = 1.5    # >1 increases contrast
_SHARPNESS_FACTOR = 2.0   # >1 increases sharpness


def _preprocess(pil_image):
    """Enhance contrast, sharpness and apply unsharp mask to improve handwriting OCR accuracy."""
    pil_image = ImageEnhance.Contrast(pil_image).enhance(_CONTRAST_FACTOR)
    pil_image = ImageEnhance.Sharpness(pil_image).enhance(_SHARPNESS_FACTOR)
    # Unsharp mask: radius=2 amplifies fine edges (letter strokes in handwriting)
    pil_image = pil_image.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
    return pil_image


def pdf_to_base64_images(pdf_path: Path, scale: float = _RENDER_SCALE) -> list[dict]:
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
            pil_image = _preprocess(pil_image)
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
