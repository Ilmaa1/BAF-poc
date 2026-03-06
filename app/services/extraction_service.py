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
        "6. Return JSON only.\n"
        "7. SOURCE PRIORITY IS MANDATORY: For every field, you MUST use the highest-priority source available. See rules below.\n\n"
        "Document identification guide:\n"
        "- 'NIB Card': issued by 'The National Insurance Board of The Commonwealth of The Bahamas'. "
        "The name layout is: first line = NIB number, second line = SURNAME (last name), third line = FIRST NAME, fourth line = MIDDLE NAME — stacked vertically with no field labels. "
        "For example: '10327541 / LIGHTBOURN / LOUISE / ICELYN' means SURNAME=LIGHTBOURN, FIRST_NAME=LOUISE, MIDDLE_NAME=ICELYN. May say 'SENIOR'.\n"
        "- \"Driver's License\": issued by 'Road Traffic Department'. Has fields labelled 'Last Name', 'First Name', 'Mid. Name', 'DOB', 'DL#'.\n"
        "- 'Voter Card': titled 'VOTER'S CARD, COMMONWEALTH OF THE BAHAMAS'. Shows name, place of birth, address. "
        "Addresses on Voter Cards are handwritten — read every character carefully. "
        "The address has two parts: 'Apt./H.#' (house number) and 'Address' (street name) on one line, and the area/boulevard on the next line. "
        "Read each word independently and do not substitute visually similar letters (e.g. do not change F to E, or drop leading words like EL).\n"
        "- 'Passport': issued by 'THE COMMONWEALTH OF THE BAHAMAS'. Contains fields: SURNAME/NOM, GIVEN NAMES/PRÉNOMS, NATIONALITY, DATE OF BIRTH/DATE DE NAISSANCE, NIB/NIR N, SEX/SEXE, PLACE OF BIRTH/LIEU DE NAISSANCE, DATE OF ISSUE, DATE OF EXPIRY, PLACE OF ISSUE. Has MRZ lines at the bottom.\n"
        "- 'Eapp': any BAF insurance form — includes 'BURIAL INSURANCE – APPLICATION FORM' and 'BURIAL INSURANCE – HEALTH QUESTIONNAIRE'. "
        "On the APPLICATION FORM: street address is in 'Street Address:', home phone in 'Home#:', cell/mobile in 'Cell#:'. "
        "On the HEALTH QUESTIONNAIRE: street address is in 'Current Address:', postal code in 'Postal Address:', city in 'City:', island in 'Island:', home phone in 'Home Phone:', mobile in 'Mobile Phone:', email in 'Email:'. "
        "For ADDR4 (country): when no Utility Bill is present, extract from the Eapp. "
        "Scan every label on the Eapp form and extract the value next to ANY label containing the word 'Country' — "
        "this includes 'Country:', 'Country of Birth:', 'Country of Origin:', 'Country of Residence:', or any similar wording. "
        "IMPORTANT: even if the label says 'Country of Birth', its value still maps to ADDR4 (it represents the country for address purposes). "
        "Do NOT skip this field just because the label says 'of Birth' — extract it for ADDR4. "
        "The reference/case number is handwritten at the top of either form (e.g. '903-256/109').\n"
        "- CRITICAL: A field labelled 'Government ID Type' in a form may contain a value such as 'NIB Card', 'Passport', or 'Driver\\'s License'. "
        "This is metadata about which ID was presented — it does NOT mean that document is physically present in the submission. "
        "Do NOT use 'Government ID Type' field values as evidence that a source document exists.\n"
        "- 'Utility Bill': a utility bill showing address details.\n\n"
        "Source priority rules (MANDATORY):\n"
        "- Each field has a `primary_source` and an ordered `fallback_sources` list.\n"
        "- Step 1: Check if the `primary_source` document is present in the images. If yes, extract from it.\n"
        "- Step 2: If primary is absent, scan `fallback_sources` left to right. Use the FIRST document in that list that is present.\n"
        "- NEVER use a document that appears later in `fallback_sources` when an earlier one is present — regardless of legibility or label clarity.\n"
        "- Concrete example: fallback_sources = ['NIB Card', 'Voter Card', \"Driver's License\"]. If NIB Card and Driver's License are both present, you MUST extract FIRST_NAME, MIDDLE_NAME, SURNAME from the NIB Card and set field_source = 'NIB Card'. Using Driver's License for names when NIB Card is present is WRONG.\n"
        "- The NIB Card does not label its name fields — the names are stacked vertically (SURNAME, then FIRST NAME, then MIDDLE NAME). This does NOT make it a lower-priority source. You must still use it over Driver's License.\n"
        "- The field_source value must always be the label of the document you actually extracted the value from.\n\n"
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
        "Starting IDP field extraction request. images=%s",
        len(images),
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
        "temperature": 0,
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
            "IDP extraction request failed before response. images=%s",
            len(images),
        )
        raise

    response.raise_for_status()
    logger.info(
        "IDP extraction response received. status_code=%s elapsed_ms=%s",
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
