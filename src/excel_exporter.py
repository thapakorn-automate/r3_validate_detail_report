
from pathlib import Path

from src.excel_filter import (
    MASTER_FILE_STEM,
    create_combined_workbook,
)


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# Round เริ่มต้น สามารถเปลี่ยนได้จาก Robot Framework
DEFAULT_ROUND = "Round14"

# ชื่อไฟล์ Output จะเป็น <Round>_Combined_Filtered.xlsx
OUTPUT_SUFFIX = "_Combined_Filtered.xlsx"


# ============================================================
# FILE HELPERS
# ============================================================

def get_round_input_dir(round_name=DEFAULT_ROUND):
    """คืนค่าโฟลเดอร์ Input ของ Round ที่เลือก"""

    round_name = str(round_name).strip()

    if not round_name:
        raise ValueError("Round name must not be empty.")

    # ป้องกันการระบุ path ที่ออกนอกโฟลเดอร์ input
    if (
        Path(round_name).name != round_name
        or "/" in round_name
        or "\\" in round_name
        or round_name in {".", ".."}
    ):
        raise ValueError(
            f"Invalid Round name: {round_name!r}. "
            "Example: Round14"
        )

    input_dir = INPUT_DIR / round_name

    if not input_dir.is_dir():
        raise FileNotFoundError(
            f"Input folder for {round_name} was not found: "
            f"{input_dir}"
        )

    return input_dir


def get_round_output_file(round_name=DEFAULT_ROUND):
    """สร้าง path ของไฟล์ Output ตาม Round ที่เลือก"""

    round_name = str(round_name).strip()

    if (
        not round_name
        or Path(round_name).name != round_name
        or "/" in round_name
        or "\\" in round_name
        or round_name in {".", ".."}
    ):
        raise ValueError(f"Invalid Round name: {round_name!r}")

    return OUTPUT_DIR / f"{round_name}{OUTPUT_SUFFIX}"


def get_excel_files(input_dir):
    """
    อ่านไฟล์ Excel เฉพาะในโฟลเดอร์ Round ที่เลือก
    ไม่อ่านไฟล์จากโฟลเดอร์ย่อย
    """

    supported_extensions = {".xlsx", ".xlsm"}

    files = [
        path
        for path in input_dir.iterdir()
        if (
            path.is_file()
            and path.suffix.lower() in supported_extensions
            and not path.name.startswith("~$")
        )
    ]

    return sorted(files, key=lambda path: path.name.casefold())


# ============================================================
# EXPORT
# ============================================================

def export_round(round_name=DEFAULT_ROUND):
    """
    Export ผลลัพธ์ของ Round ที่เลือก

    1. อ่านไฟล์จาก input/<Round>/
    2. สร้าง Combined workbook โดยใช้ Logic เดิม
    3. บันทึกเป็น output/<Round>_Combined_Filtered.xlsx
    """

    round_name = str(round_name).strip()

    input_dir = get_round_input_dir(round_name)
    output_file = get_round_output_file(round_name)

    excel_files = get_excel_files(input_dir)

    if not excel_files:
        raise FileNotFoundError(
            f"No Excel files found in: {input_dir}"
        )

    print(f"Selected Round : {round_name}")
    print(f"Input folder   : {input_dir}")
    print(f"Excel files    : {len(excel_files)}")
    print(f"Output file    : {output_file}")

    master_candidates = [
        path
        for path in excel_files
        if MASTER_FILE_STEM.casefold() in path.stem.casefold()
    ]

    if not master_candidates:
        raise FileNotFoundError(
            f"Master file containing '{MASTER_FILE_STEM}' "
            f"was not found in: {input_dir}"
        )

    if len(master_candidates) > 1:
        paths = "\n".join(str(path) for path in master_candidates)
        raise ValueError(
            f"Multiple Master files found in {input_dir}:\n{paths}"
        )

    master_file = master_candidates[0]
    detail_files = [
        path for path in excel_files if path != master_file
    ]

    if not detail_files:
        raise FileNotFoundError(
            f"No Detail Excel files found in: {input_dir}"
        )

    workbook, metadata = create_combined_workbook(
        master_file,
        detail_files,
    )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    try:
        workbook.save(output_file)
    finally:
        workbook.close()

    print(f"Export completed: {output_file}")

    metadata.update(
        {
            "round_name": round_name,
            "input_dir": str(input_dir),
            "output_file": str(output_file),
            "input_count": len(excel_files),
        }
    )
    return metadata


if __name__ == "__main__":
    # เปลี่ยน Round ตรงนี้ได้ หากต้องการรันผ่าน Python โดยตรง
    export_round(DEFAULT_ROUND)