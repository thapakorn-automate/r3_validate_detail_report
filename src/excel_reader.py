
from pathlib import Path
from typing import List

from config import INPUT_DIR, EXCEL_EXTENSIONS, EXCLUDE_TEMP_FILES


def find_excel_files() -> List[Path]:
    """Return all supported Excel files in the input folder."""
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    files = []

    for path in sorted(INPUT_DIR.iterdir()):
        if not path.is_file():
            continue

        if path.suffix.lower() not in EXCEL_EXTENSIONS:
            continue

        if EXCLUDE_TEMP_FILES and path.name.startswith("~$"):
            continue

        files.append(path)

    return files