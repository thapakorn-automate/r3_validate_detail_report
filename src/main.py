
from config import create_directories
from src.excel_reader import find_excel_files
from src.excel_exporter import export_filtered_workbook


def main():
    create_directories()
    input_files = find_excel_files()

    if not input_files:
        raise FileNotFoundError(
            "No Excel input files found in the input folder."
        )

    total_files = 0

    for input_path in input_files:
        print(f"\nProcessing: {input_path.name}")

        summary = export_filtered_workbook(input_path)

        print(f"Output: {summary['output_path']}")
        print(f"Sheets: {len(summary['sheets'])}")
        print(f"Rows checked: {summary['total_rows_checked']}")
        print(f"Rows kept: {summary['total_rows_kept']}")
        print(f"Rows removed: {summary['total_rows_removed']}")

        for sheet_name, counts in summary["sheets"].items():
            print(
                f"  {sheet_name}: "
                f"kept={counts['kept']}, "
                f"removed={counts['removed']}"
            )

        total_files += 1

    print(f"\nCompleted: {total_files} file(s)")


if __name__ == "__main__":
    main()