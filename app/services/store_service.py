import logging

from app.models import DocumentRecord

logger = logging.getLogger(__name__)


class DocumentStore:
    def __init__(self) -> None:
        self._records: dict[str, DocumentRecord] = {}

    def put(self, record: DocumentRecord) -> None:
        self._records[record.document_id] = record
        logger.info("Stored document record in memory. document_id=%s", record.document_id)

    def get(self, document_id: str) -> DocumentRecord | None:
        record = self._records.get(document_id)
        if record is None:
            logger.debug("Document record not found in store. document_id=%s", document_id)
        else:
            logger.debug("Fetched document record from store. document_id=%s", document_id)
        return record


document_store = DocumentStore()
