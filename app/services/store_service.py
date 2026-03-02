import json
import logging
from datetime import datetime
from pathlib import Path

from app.config import RECORDS_DIR
from app.models import DocumentRecord

logger = logging.getLogger(__name__)


def _record_to_dict(record: DocumentRecord) -> dict:
    return {
        "document_id": record.document_id,
        "file_name": record.file_name,
        "file_path": record.file_path,
        "extracted_fields": record.extracted_fields,
        "field_confidence": record.field_confidence,
        "field_source": record.field_source,
        "created_at": record.created_at.isoformat(),
    }


def _dict_to_record(data: dict) -> DocumentRecord:
    return DocumentRecord(
        document_id=data["document_id"],
        file_name=data["file_name"],
        file_path=data["file_path"],
        extracted_fields=data["extracted_fields"],
        field_confidence=data["field_confidence"],
        field_source=data["field_source"],
        created_at=datetime.fromisoformat(data["created_at"]),
    )


class DocumentStore:
    """File-based persistent store shared across all uvicorn workers."""

    def __init__(self, records_dir: Path | None = None) -> None:
        self._records_dir = records_dir or RECORDS_DIR

    def _record_path(self, document_id: str) -> Path:
        return self._records_dir / f"{document_id}.json"

    def put(self, record: DocumentRecord) -> None:
        path = self._record_path(record.document_id)
        tmp_path = path.with_suffix(".json.tmp")
        try:
            with tmp_path.open("w", encoding="utf-8") as f:
                json.dump(_record_to_dict(record), f, indent=2)
            tmp_path.rename(path)
            logger.info("Stored document record. document_id=%s path=%s", record.document_id, path)
        except Exception:
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)
            raise

    def get(self, document_id: str) -> DocumentRecord | None:
        path = self._record_path(document_id)
        if not path.exists():
            logger.debug("Document record not found. document_id=%s", document_id)
            return None
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            record = _dict_to_record(data)
            logger.debug("Fetched document record. document_id=%s", document_id)
            return record
        except Exception as e:
            logger.warning("Failed to load document record. document_id=%s error=%s", document_id, e)
            return None


document_store = DocumentStore()
