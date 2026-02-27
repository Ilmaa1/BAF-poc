# BAF Document Extraction App

Two-screen Python application for:
- uploading PDF documents,
- converting PDF pages to base64 images,
- extracting key-value pairs using OpenAI vision model,
- showing browser-native PDF preview and structured extraction results.

## Tech Stack
- FastAPI
- Jinja2 templates
- pypdfium2 for PDF-to-image conversion
- OpenAI Responses API for VLM extraction

## Setup
1. Create and activate a virtual environment.
2. Install dependencies:
   - `pip install -r requirements.txt`
3. Configure environment variables:
   - copy `.env.example` values into your environment
   - `OPENAI_API_KEY`
   - `OPENAI_MODEL` (default: `gpt-5.2`)

## Run
- `uvicorn app.main:app --reload`

Open `http://127.0.0.1:7005`

## Notes
- Extraction follows strict no-inference rules from your schema.
- When a field is not explicitly present, the app returns `null`.
- Results are grouped into dynamic sections (Identity, Policy, Address, Contact).
