"""
File: fill_sigmat_tabs.py
Version: v1.0.0
Role: 기존 엑셀의 종가/지수 데이터를 바탕으로 sigmat/isigmat(STD20) 시트를 채운다.
# 메모: v1.0.0 - 20일 표준편차(STD20) 전용 백필 패치 스크립트 추가
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
FILE_MAP_PATH = BASE_DIR / "stock_file_map.json"
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from extra_scores import (
    INDEX_TARGET_STEM,
    get_close_data,
    get_index_data,
    save_sigma_t_sheet,
)


def load_target_files() -> list[Path]:
    with FILE_MAP_PATH.open("r", encoding="utf-8") as fp:
        file_map = json.load(fp)
    return [BASE_DIR / filename for filename in file_map.values()]


def fill_sigmat_for_file(excel_path: Path) -> None:
    if not excel_path.exists():
        print(f"⚠ 파일이 없어 건너뜁니다: {excel_path.name}")
        return

    # 시트명은 sigmat/isigmat를 유지하지만, 의미는 20일 표준편차(STD20)다.
    print(f"\n=== STD20 패치 시작: {excel_path.name} ===")

    close_dates, close_stocks = get_close_data(excel_path)
    if close_dates and close_stocks:
        save_sigma_t_sheet(
            excel_path,
            close_dates,
            close_stocks,
            sheet_name="sigmat",
            window_std=20,
        )
    else:
        print("⚠ 종가 데이터가 없어 STD20 계산을 건너뜁니다.")

    if excel_path.stem == INDEX_TARGET_STEM:
        idx_dates, idx_stocks = get_index_data(excel_path)
        if idx_dates and idx_stocks:
            save_sigma_t_sheet(
                excel_path,
                idx_dates,
                idx_stocks,
                sheet_name="isigmat",
                window_std=20,
                name_header="업종명",
                code_header="업종코드",
            )
        else:
            print("⚠ 지수 데이터가 없어 지수 STD20 계산을 건너뜁니다.")

    print(f"=== STD20 패치 완료: {excel_path.name} ===")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="기존 엑셀 데이터로 sigmat/isigmat(STD20) 시트를 채웁니다."
    )
    parser.add_argument(
        "--file",
        dest="filename",
        help="대상 엑셀 파일명 또는 경로. 미지정 시 stock_file_map.json의 모든 파일 처리",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.filename:
        target_files = [Path(args.filename)]
        if not target_files[0].is_absolute():
            target_files[0] = BASE_DIR / target_files[0]
    else:
        target_files = load_target_files()

    for excel_path in target_files:
        fill_sigmat_for_file(excel_path)


if __name__ == "__main__":
    main()
