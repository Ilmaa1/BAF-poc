# BAF PDF Processor

Monitors the input folder for PDFs, extracts structured key-value pairs using OpenAI vision, and saves Excel output to the output folder.

## How It Works

1. Start the scheduler (Docker or locally)
2. Place PDFs in the `input/` folder
3. New PDFs are processed automatically and appended to a single Excel file
4. Output: `output/DQ_Anomaly_V2.xlsx` (one row per PDF)
5. Input PDFs are deleted after successful processing

## Setup

1. Create and activate a virtual environment
2. Install dependencies: `pip install -r requirements.txt`
3. Configure environment:
   - Copy `.env.example` to `.env`
   - Set `OPENAI_API_KEY` (required)
   - Optionally set `OPENAI_MODEL` (default: `gpt-5.2`)

## Run Locally

```bash
# Start scheduler (monitors input folder continuously)
python scheduler.py

# One-time batch run
python process_pdfs.py
```

With custom folders:

```bash
python scheduler.py --input-dir ./my_pdfs --output-dir ./my_results --poll-interval 60
```

## Run with Docker

```bash
# Start scheduler (runs in background, monitors ./input)
./deploy.sh deploy

# Other commands
./deploy.sh stop      # Stop scheduler
./deploy.sh status    # Container status
./deploy.sh logs      # Follow logs
./deploy.sh run-once  # One-time batch (no scheduler)
```

## Output

Single file `output/DQ_Anomaly_V2.xlsx` with one row per processed PDF. Columns: PERNO, FIRST_NAME, MIDDLE_NAME, SURNAME, DOB, NI_NUMBER, and other configured fields, plus workflow columns (Updated by, Comments, Reviewed by, etc.). New rows are appended; existing data is preserved.

## Extraction Rules

- Extract only values explicitly visible in the document
- No inference, guessing, or normalization
- Returns empty when a field is not found
