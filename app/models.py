from dataclasses import dataclass
from datetime import datetime


@dataclass
class DocumentRecord:
    document_id: str
    file_name: str
    file_path: str
    extracted_fields: dict
    field_confidence: dict
    field_source: dict
    created_at: datetime
