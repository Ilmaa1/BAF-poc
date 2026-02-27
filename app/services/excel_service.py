import json
from io import BytesIO
import logging

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from app.services.field_config import DOCUMENT_EXTRACTION_CONFIG, SECTION_ORDER

logger = logging.getLogger(__name__)


def _field_section_map() -> dict[str, str]:
    section_map: dict[str, str] = {}
    for section_name, field_names in SECTION_ORDER:
        for field_name in field_names:
            section_map[field_name] = section_name
    return section_map


def _autosize_columns(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 14), 45)


def _add_headers(worksheet, headers: list[str]) -> None:
    worksheet.append(headers)
    header_fill = PatternFill("solid", fgColor="1D4ED8")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font


def generate_extraction_excel(
    extracted_fields: dict,
    field_confidence: dict,
    field_source: dict,
) -> bytes:
    config = DOCUMENT_EXTRACTION_CONFIG["document_extraction_config"]
    fields = config["fields"]
    section_map = _field_section_map()

    workbook = Workbook()
    logger.info("Generating extraction excel. total_fields=%s", len(fields))

    data_sheet = workbook.active
    data_sheet.title = "Extracted Data"
    _add_headers(
        data_sheet,
        [
            "Section",
            "Field Name",
            "Value",
            "Confidence (%)",
            "Detected Source",
            "Required",
            "Primary Source",
            "Fallback Sources",
            "Notes",
        ],
    )

    for field in fields:
        name = field["field_name"]
        confidence_value = field_confidence.get(name)
        confidence_percent = (
            round(float(confidence_value) * 100, 2) if confidence_value is not None else None
        )
        data_sheet.append(
            [
                section_map.get(name, "Unmapped"),
                name,
                extracted_fields.get(name),
                confidence_percent,
                field_source.get(name),
                "Yes" if field["extraction_required"] else "No",
                field["primary_source"],
                ", ".join(field["fallback_sources"]) if field["fallback_sources"] else None,
                field["notes"],
            ]
        )

    _autosize_columns(data_sheet)
    data_sheet.freeze_panes = "A2"

    raw_json_sheet = workbook.create_sheet(title="Raw JSON")
    _add_headers(raw_json_sheet, ["JSON"])
    raw_payload = {
        "version": config["version"],
        "extracted_fields": extracted_fields,
        "field_confidence": field_confidence,
        "field_source": field_source,
    }
    raw_json_sheet.append([json.dumps(raw_payload, indent=2)])
    raw_json_sheet.column_dimensions["A"].width = 120

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    logger.info("Completed extraction excel generation.")
    return output.getvalue()
