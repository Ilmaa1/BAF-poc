from app.services.field_config import DOCUMENT_EXTRACTION_CONFIG, SECTION_ORDER


def build_sectioned_results(
    extracted_fields: dict,
    field_confidence: dict,
    field_source: dict,
) -> list[dict]:
    field_map = {
        field["field_name"]: field
        for field in DOCUMENT_EXTRACTION_CONFIG["document_extraction_config"]["fields"]
    }

    sections: list[dict] = []
    for section_title, field_names in SECTION_ORDER:
        section_fields = []
        for field_name in field_names:
            config = field_map[field_name]
            section_fields.append(
                {
                    "name": field_name,
                    "value": extracted_fields.get(field_name),
                    "confidence": field_confidence.get(field_name),
                    "detected_source": field_source.get(field_name),
                    "required": config["extraction_required"],
                    "notes": config["notes"],
                }
            )

        sections.append({"title": section_title, "fields": section_fields})

    return sections
