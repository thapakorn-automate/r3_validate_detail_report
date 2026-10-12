
from pathlib import Path

# Project root
BASE_DIR = Path(__file__).resolve().parent

# Folder paths
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
REPORT_DIR = BASE_DIR / "Reports"

# Excel settings
EXCEL_EXTENSIONS = (".xlsx",)
MATCH_COLUMN = 1  # Column A
INCLUDE_EXTENSION = True

# Output settings
OUTPUT_SUFFIX = "_Filtered"

# Supported Excel files
EXCLUDE_TEMP_FILES = True


def create_directories():
    """Create required folders if they do not exist."""
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)