
from pathlib import Path

from openpyxl import load_workbook

from src.excel_exporter import export_round


class ExcelFilterLibrary:
    ROBOT_LIBRARY_SCOPE = "SUITE"

    def __init__(self):
        self.project_dir = Path(__file__).resolve().parents[1]
        self.last_output = None
        self.last_metadata = None
        self.last_round = None

    def _get_round_input_dir(self, round_name="Round14"):
        round_name = str(round_name).strip()

        if (
            not round_name
            or Path(round_name).name != round_name
            or "/" in round_name
            or "\\" in round_name
            or round_name in {".", ".."}
        ):
            raise ValueError(f"Invalid Round name: {round_name!r}")

        input_dir = self.project_dir / "input" / round_name

        if not input_dir.is_dir():
            raise AssertionError(
                f"Input folder not found: {input_dir}"
            )

        return input_dir

    def process_all_excel_files(self, round_name="Round14"):
        """
        Process Excel files only in input/<Round>.
        Output: output/<Round>_Combined_Filtered.xlsx
        """

        round_name = str(round_name).strip()
        self._get_round_input_dir(round_name)
        metadata = export_round(round_name)
        output_file = Path(metadata["output_file"])

        self.last_output = output_file
        self.last_metadata = metadata
        self.last_round = round_name

        print(f"Selected Round: {round_name}")
        print(f"Input folder: {metadata['input_dir']}")
        print(f"Output file: {output_file}")
        print(f"Master count: {metadata['master_count']}")
        print(f"Detail file count: {len(metadata['detail_files'])}")

        return str(output_file)

    def get_input_file_count(self, round_name="Round14"):
        input_dir = self._get_round_input_dir(round_name)

        return sum(
            1
            for path in input_dir.iterdir()
            if path.is_file()
            and path.suffix.lower() in {".xlsx", ".xlsm"}
            and not path.name.startswith("~$")
        )

    def verify_output_files_exist(self, round_name=None):
        if self.last_output is None:
            raise AssertionError(
                "Output has not been generated yet."
            )

        if round_name is not None:
            if str(round_name).strip() != self.last_round:
                raise AssertionError(
                    f"Selected Round is {round_name}, "
                    f"but last exported Round is {self.last_round}."
                )

        if not self.last_output.is_file():
            raise AssertionError(
                f"Output file does not exist: {self.last_output}"
            )

        if self.last_output.stat().st_size == 0:
            raise AssertionError(
                f"Output file is empty: {self.last_output}"
            )

        return True

    def validate_excel_comparison_results(self):
        if self.last_output is None:
            raise AssertionError(
                "Run Process All Excel Files first."
            )

        if not self.last_output.is_file():
            raise AssertionError(
                f"Output file not found: {self.last_output}"
            )

        workbook = load_workbook(
            self.last_output,
            read_only=False,
            data_only=False,
        )

        try:
            required_sheets = {
                "Summary",
                "Combine result",
            }

            missing_sheets = (
                required_sheets - set(workbook.sheetnames)
            )

            if missing_sheets:
                raise AssertionError(
                    f"Missing sheets: {missing_sheets}"
                )

            ws = workbook["Combine result"]
            status_count = 0

            for row in ws.iter_rows():
                for cell in row:
                    if cell.value not in ("Match", "Diff"):
                        continue

                    status_count += 1
                    color = cell.fill.fgColor.rgb

                    if cell.value == "Match":
                        valid_colors = {
                            "FFD9EAD3",
                            "D9EAD3",
                        }

                        if color not in valid_colors:
                            raise AssertionError(
                                f"Match cell {cell.coordinate} "
                                f"does not have green fill. "
                                f"Found: {color}"
                            )

                    elif cell.value == "Diff":
                        valid_colors = {
                            "FFF4CCCC",
                            "F4CCCC",
                        }

                        if color not in valid_colors:
                            raise AssertionError(
                                f"Diff cell {cell.coordinate} "
                                f"does not have red fill. "
                                f"Found: {color}"
                            )

            if status_count == 0:
                raise AssertionError(
                    "No Match/Diff status cells found."
                )

        finally:
            workbook.close()

        return True