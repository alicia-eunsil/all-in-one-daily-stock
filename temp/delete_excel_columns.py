"""
File: delete_excel_columns.py
Version: v1.0.0
Role: 지정한 엑셀 파일의 특정 시트에서 연속된 열을 삭제한다.
"""

import argparse
from pathlib import Path
import openpyxl
from openpyxl.utils import column_index_from_string


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = BASE_DIR / "KR_Stocks_Individual.xlsx"
DEFAULT_SHEET = "std"
DEFAULT_START_COL = "C"  # 필요에 맞게 수정
DEFAULT_END_COL = "D"    # 필요에 맞게 수정


def parse_col(value: str) -> int:
    value = value.strip()
    try:
        return int(value)
    except ValueError:
        return column_index_from_string(value)


def delete_columns(file_path: Path, sheet_name: str, start_col: str, end_col: str) -> None:
    if not file_path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    wb = openpyxl.load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"시트를 찾을 수 없습니다: {sheet_name}")

    ws = wb[sheet_name]
    start_idx = parse_col(start_col)
    end_idx = parse_col(end_col)
    if start_idx > end_idx:
        start_idx, end_idx = end_idx, start_idx

    amount = end_idx - start_idx + 1
    ws.delete_cols(start_idx, amount)
    wb.save(file_path)
    wb.close()


def main():
    parser = argparse.ArgumentParser(description="엑셀 시트에서 특정 열 범위를 삭제합니다.")
    parser.add_argument("--file", default=str(DEFAULT_FILE), help="엑셀 파일 경로")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="시트 이름")
    parser.add_argument("--start_col", default=str(DEFAULT_START_COL), help="삭제 시작 열 (예: 3 또는 C)")
    parser.add_argument("--end_col", default=str(DEFAULT_END_COL), help="삭제 끝 열 (예: 5 또는 E)")
    args = parser.parse_args()

    delete_columns(Path(args.file), args.sheet, args.start_col, args.end_col)
    print(f"✅ 열 삭제 완료: {args.file} [{args.sheet}] {args.start_col}-{args.end_col}")


if __name__ == "__main__":
    main()
