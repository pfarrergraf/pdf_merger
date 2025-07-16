# utils.py (or config.py)
import os
from dotenv import load_dotenv

# Load .env file into environment variables
load_dotenv()

# Now fetch them (with sensible defaults if you like)
WATCH_FOLDER = os.getenv("WATCH_FOLDER", "")
SMTP_SERVER  = os.getenv("SMTP_SERVER")
SMTP_PORT    = int(os.getenv("SMTP_PORT", 0))
EMAIL_FROM   = os.getenv("EMAIL_FROM")
EMAIL_TO     = os.getenv("EMAIL_TO")
EMAIL_USER   = os.getenv("EMAIL_USER")
EMAIL_PASS   = os.getenv("EMAIL_PASS")
