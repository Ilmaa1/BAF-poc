from io import BytesIO
import logging
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Output filename for consolidated Excel
OUTPUT_EXCEL_NAME = "DQ_Anomaly_V2.xlsx"

# Column order: extracted fields first, then workflow columns (empty for extraction)
EXCEL_COLUMNS = [
    "PERNO",
    "FIRST_NAME",
    "MIDDLE_NAME",
    "SURNAME",
    "DOB",
    "NI_NUMBER",
    "CONTACT_MODE",
    "BIRTH_NATION_NO",
    "SMOKER",
    "POL_REF_NO",
    "POLICY_NO",
    "TERM",
    "START_DATE",
    "END_DATE",
    "FREQ",
    "POL_STAT",
    "CURRENT_BRANCH",
    "PAY_ID",
    "ADDRNO",
    "ADDR1",
    "ADDR2",
    "ADDR3",
    "ADDR4",
    "POSTCODE",
    "HOME_PHONES",
    "MOBILE_PHONES",
    "WORK_PHONES",
    "EMAILS",
    "FAX",
    "SMS",
    "Updated by",
    "Updated date",
    "Comments",
    "Reviewed by",
    "Reviewed date",
    "Reviewer Comments",
    "Approved by",
    "Approved date",
    "Approver Comments",
]


def _autosize_columns(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 45)


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
    workbook = Workbook()
    data_sheet = workbook.active
    data_sheet.title = "Extracted Data"

    _add_headers(data_sheet, EXCEL_COLUMNS)

    # Build data row: extracted values for each field, then empty workflow columns
    row = []
    for col in EXCEL_COLUMNS:
        if col in extracted_fields:
            val = extracted_fields.get(col)
            row.append(val if val is not None else "")
        else:
            row.append("")

    data_sheet.append(row)
    _autosize_columns(data_sheet)
    data_sheet.freeze_panes = "A2"

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    logger.info("Completed extraction excel generation.")
    return output.getvalue()


def _extracted_fields_to_row(extracted_fields: dict) -> list:
    """Convert extracted_fields dict to row values in EXCEL_COLUMNS order."""
    row = []
    for col in EXCEL_COLUMNS:
        if col in extracted_fields:
            val = extracted_fields.get(col)
            row.append(val if val is not None else "")
        else:
            row.append("")
    return row


def append_rows_to_excel(
    output_path: Path,
    rows: list[dict],
) -> None:
    """
    Append rows to DQ_Anomaly_V2.xlsx. Creates file with headers if it doesn't exist.
    Each row is an extracted_fields dict.
    """
    if not rows:
        return

    if output_path.exists():
        workbook = load_workbook(output_path)
        data_sheet = workbook.active
    else:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "Extracted Data"
        _add_headers(data_sheet, EXCEL_COLUMNS)

    for extracted_fields in rows:
        row = _extracted_fields_to_row(extracted_fields)
        data_sheet.append(row)

    _autosize_columns(data_sheet)
    data_sheet.freeze_panes = "A2"
    workbook.save(output_path)
    logger.info("Appended %d row(s) to %s", len(rows), output_path.name)
