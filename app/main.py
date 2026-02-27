from datetime import datetime
import logging
from pathlib import Path
import shutil
import time
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import BASE_DIR, UPLOAD_DIR
from app.models import DocumentRecord
from app.services.excel_service import generate_extraction_excel
from app.services.extraction_service import extract_fields_from_images
from app.services.pdf_service import pdf_to_base64_images
from app.services.result_service import build_sectioned_results
from app.services.store_service import document_store

app = FastAPI(title="Document Extraction App")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
logger = logging.getLogger(__name__)


@app.get("/")
async def upload_screen(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request, "error": None})


@app.post("/extract")
async def extract_document(request: Request, file: UploadFile = File(...)):
    request_start = time.perf_counter()

    if not file.filename or not file.filename.lower().endswith(".pdf"):
        logger.warning("Rejected upload with invalid file type: %s", file.filename)
        return templates.TemplateResponse(
            "upload.html",
            {"request": request, "error": "Please upload a valid PDF document."},
            status_code=400,
        )

    document_id = str(uuid4())
    file_name = Path(file.filename).name
    saved_path = UPLOAD_DIR / f"{document_id}.pdf"
    logger.info("Started document extraction. document_id=%s file_name=%s", document_id, file_name)

    with saved_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    logger.info(
        "Saved uploaded PDF. document_id=%s path=%s size_bytes=%s",
        document_id,
        saved_path,
        saved_path.stat().st_size,
    )

    try:
        pdf_render_start = time.perf_counter()
        images = pdf_to_base64_images(saved_path)
        pdf_render_elapsed_ms = int((time.perf_counter() - pdf_render_start) * 1000)
        logger.info(
            "Rendered PDF to images. document_id=%s pages=%s elapsed_ms=%s",
            document_id,
            len(images),
            pdf_render_elapsed_ms,
        )

        extraction_start = time.perf_counter()
        extraction_result = extract_fields_from_images(images)
        extraction_elapsed_ms = int((time.perf_counter() - extraction_start) * 1000)
        populated_fields = sum(
            1 for value in extraction_result["extracted_fields"].values() if value is not None
        )
        logger.info(
            "Completed field extraction. document_id=%s populated_fields=%s elapsed_ms=%s",
            document_id,
            populated_fields,
            extraction_elapsed_ms,
        )
    except Exception:
        logger.exception("Document extraction failed. document_id=%s file_name=%s", document_id, file_name)
        if saved_path.exists():
            saved_path.unlink()
            logger.info("Removed failed upload file. document_id=%s path=%s", document_id, saved_path)
        return templates.TemplateResponse(
            "upload.html",
            {
                "request": request,
                "error": "Could not process the document. Check API settings and try again.",
            },
            status_code=500,
        )

    record = DocumentRecord(
        document_id=document_id,
        file_name=file_name,
        file_path=str(saved_path),
        extracted_fields=extraction_result["extracted_fields"],
        field_confidence=extraction_result["field_confidence"],
        field_source=extraction_result["field_source"],
        created_at=datetime.utcnow(),
    )
    document_store.put(record)
    total_elapsed_ms = int((time.perf_counter() - request_start) * 1000)
    logger.info(
        "Stored extraction record. document_id=%s total_elapsed_ms=%s",
        document_id,
        total_elapsed_ms,
    )

    return RedirectResponse(url=f"/results/{document_id}", status_code=303)


@app.get("/results/{document_id}")
async def results_screen(request: Request, document_id: str):
    record = document_store.get(document_id)
    if record is None:
        logger.warning("Results requested for missing document. document_id=%s", document_id)
        raise HTTPException(status_code=404, detail="Document not found.")

    sections = build_sectioned_results(
        record.extracted_fields,
        record.field_confidence,
        record.field_source,
    )
    logger.info("Prepared results page data. document_id=%s", document_id)

    return templates.TemplateResponse(
        "results.html",
        {
            "request": request,
            "document_id": record.document_id,
            "file_name": record.file_name,
            "sections": sections,
        },
    )


@app.get("/documents/{document_id}")
async def view_pdf(document_id: str):
    record = document_store.get(document_id)
    if record is None:
        logger.warning("PDF requested for missing document. document_id=%s", document_id)
        raise HTTPException(status_code=404, detail="Document not found.")

    logger.info("Serving stored PDF. document_id=%s file_name=%s", document_id, record.file_name)

    return FileResponse(
        path=record.file_path,
        media_type="application/pdf",
        filename=record.file_name,
        headers={"Content-Disposition": f'inline; filename="{record.file_name}"'},
    )


@app.get("/results/{document_id}/download-excel")
async def download_excel(document_id: str):
    record = document_store.get(document_id)
    if record is None:
        logger.warning("Excel requested for missing document. document_id=%s", document_id)
        raise HTTPException(status_code=404, detail="Document not found.")

    excel_bytes = generate_extraction_excel(
        extracted_fields=record.extracted_fields,
        field_confidence=record.field_confidence,
        field_source=record.field_source,
    )

    output_name = f"{Path(record.file_name).stem}_extraction.xlsx"
    logger.info("Generated excel download. document_id=%s output_name=%s", document_id, output_name)
    return StreamingResponse(
        content=iter([excel_bytes]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{output_name}"'},
    )
