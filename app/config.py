import os
from dotenv import load_dotenv

load_dotenv()

OUTPUT_DIR = os.getenv("OUTPUT_DIR", "")
EXCEL_PATH = os.getenv("EXCEL_PATH", "")
COVER_LETTER_SOURCE_PATH = os.getenv("COVER_LETTER_SOURCE_PATH", "")
COVER_LETTER_SOURCE_PATH_PDF = os.getenv("COVER_LETTER_SOURCE_PATH_PDF", "")
COVER_LETTER_MASTER_PATH = os.getenv("COVER_LETTER_MASTER_PATH", "")
COVER_LETTER_WORKING_NAME = "Cover_Letter_Mustafa_Ansari.docx"

CV_OPTIONS = {
    "1": {
        "label": "Backend Engineer CV",
        "path": os.getenv("CV_BACKEND_PATH", ""),
    },
    "2": {
        "label": "Python AI Dev CV",
        "path": os.getenv("CV_PYTHON_AI_PATH", ""),
    },
}

EXCEL_FIELDS = [
    ("job_id", "Job ID (optional): ", False),
    ("portal_choice", "Select portal (1-5): ", False),
    ("job_link", "Job Link: ", False),
    ("job_type", "Type (Startup/Mid/Big): ", False),
    ("industry", "Industry: ", False),
    ("location", "Location: ", False),
    ("status", "Status: ", False),
    ("docs_sent", "Docs Sent: ", False),
    ("todo_done", "TO-DO (optional): ", False),
    ("hr_name", "HR Name: ", False),
    ("hr_contact", "HR Contact: ", False),
    ("comments", "Comments (optional): ", False),
    ("id_pw", "ID/PW (optional): ", False),
]
