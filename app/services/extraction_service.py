import json
import logging
import time
from typing import Any

import requests

from app.config import OPENAI_API_KEY, OPENAI_MODEL
from app.services.field_config import DOCUMENT_EXTRACTION_CONFIG

logger = logging.getLogger(__name__)


def _build_prompt() -> str:
    schema = DOCUMENT_EXTRACTION_CONFIG["document_extraction_config"]
    return (
        "You are a strict document extraction engine.\n"
        "Follow these rules exactly:\n"
        "1. Extract only values explicitly visible in the provided document images.\n"
        "2. Do not infer, guess, normalize, or transform any value.\n"
        "3. If a field is not explicitly found, return null.\n"
        "4. For field_source, return null unless the source is explicitly visible in the document.\n"
        "5. field_source must match a configured source label for that field when present.\n"
        "6. Return JSON only.\n\n"
        "Return with this exact shape:\n"
        "{\n"
        '  "version": "1.0",\n'
        '  "extracted_fields": {\n'
        '    "FIELD_NAME": "value or null"\n'
        "  },\n"
        '  "field_confidence": {\n'
        '    "FIELD_NAME": "number between 0 and 1, or null when value is null"\n'
        "  },\n"
        '  "field_source": {\n'
        '    "FIELD_NAME": "actual source label where value was found, or null if not found"\n'
        "  }\n"
        "}\n\n"
        f"Field configuration:\n{json.dumps(schema, indent=2)}"
    )


def _empty_result() -> dict[str, Any]:
    fields = DOCUMENT_EXTRACTION_CONFIG["document_extraction_config"]["fields"]
    return {
        "version": DOCUMENT_EXTRACTION_CONFIG["document_extraction_config"]["version"],
        "extracted_fields": {field["field_name"]: None for field in fields},
        "field_confidence": {field["field_name"]: None for field in fields},
        "field_source": {field["field_name"]: None for field in fields},
    }


def _safe_json(text: str) -> dict[str, Any]:
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        logger.warning("Primary JSON parsing failed; attempting to recover JSON object from text.")
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start : end + 1])
        raise


def _extract_output_text(response_json: dict[str, Any]) -> str:
    if isinstance(response_json.get("output_text"), str):
        return response_json["output_text"]

    output_items = response_json.get("output", [])
    texts: list[str] = []

    for item in output_items:
        for content_item in item.get("content", []):
            item_type = content_item.get("type")
            if item_type in {"output_text", "text"}:
                text_value = content_item.get("text")
                if isinstance(text_value, str):
                    texts.append(text_value)
                elif isinstance(text_value, dict) and isinstance(text_value.get("value"), str):
                    texts.append(text_value["value"])

    return "\n".join(texts).strip()


def _normalize_confidence(value: Any) -> float | None:
    if value is None:
        return None
    try:
        score = float(value)
    except (TypeError, ValueError):
        return None
    if score < 0:
        return 0.0
    if score > 1:
        return 1.0
    return score


def _normalize_source(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    cleaned = value.strip()
    return cleaned or None


def extract_fields_from_images(images: list[dict]) -> dict[str, Any]:
    if not OPENAI_API_KEY:
        logger.warning("OPENAI_API_KEY is not configured; returning empty extraction result.")
        return _empty_result()

    logger.info(
        "Starting field extraction request. images=%s model=%s",
        len(images),
        OPENAI_MODEL,
    )
    content: list[dict[str, Any]] = [{"type": "input_text", "text": _build_prompt()}]

    for image in images:
        content.append(
            {
                "type": "input_image",
                "image_url": f"data:{image['mime_type']};base64,{image['image_base64']}",
            }
        )

    payload = {
        "model": OPENAI_MODEL,
        "input": [{"role": "user", "content": content}],
    }
    headers = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json",
    }

    request_start = time.perf_counter()
    try:
        response = requests.post(
            "https://api.openai.com/v1/responses",
            headers=headers,
            json=payload,
            timeout=180,
        )
    except requests.RequestException:
        logger.exception(
            "OpenAI extraction request failed before response. images=%s model=%s",
            len(images),
            OPENAI_MODEL,
        )
        raise

    response.raise_for_status()
    logger.info(
        "OpenAI extraction response received. status_code=%s elapsed_ms=%s",
        response.status_code,
        int((time.perf_counter() - request_start) * 1000),
    )
    response_json = response.json()

    raw_text = _extract_output_text(response_json) or "{}"
    parsed = _safe_json(raw_text)

    result = _empty_result()
    extracted = parsed.get("extracted_fields", {})
    confidence = parsed.get("field_confidence", {})
    source = parsed.get("field_source", {})
    for field_name in result["extracted_fields"]:
        if field_name in extracted:
            result["extracted_fields"][field_name] = extracted[field_name]
        if field_name in confidence:
            result["field_confidence"][field_name] = _normalize_confidence(confidence[field_name])
        if field_name in source:
            result["field_source"][field_name] = _normalize_source(source[field_name])

    populated_fields = sum(1 for value in result["extracted_fields"].values() if value is not None)
    logger.info(
        "Completed extraction response normalization. populated_fields=%s total_fields=%s",
        populated_fields,
        len(result["extracted_fields"]),
    )
    return result
