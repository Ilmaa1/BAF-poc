from pathlib import Path
import logging
import os
import datetime

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.2")
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

_log_filename = LOGS_DIR / f"extraction_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
_fmt = "%(asctime)s %(levelname)s [%(name)s] %(message)s"

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format=_fmt,
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(_log_filename, encoding="utf-8"),
    ],
)

LOG_FILE = _log_filename
