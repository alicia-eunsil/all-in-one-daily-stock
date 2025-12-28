"""
File: recompute_std_last10.py
Version: v1.1.0
Role: 최근 10일간 STD 값을 새로운 산식으로 재계산해 원본 파일을 덮어쓴다.
"""

from pathlib import Path
import numpy as np
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "KR_Stocks_Individual.xlsx"
STD_SHEET = "std"
WINDOW = 20
RECENT_DAYS = 10


def load_std_sheet(path: Path):
    df = pd.read_excel(path, sheet_name=STD_SHEET)
    if df.empty or len(df.columns) < 3:
        raise ValueError(f"{STD_SHEET} 시트에 데이터가 없습니다.")
    return df


def recompute_std(values: np.ndarray) -> float:
    mean = np.nanmean(values)
    diffs = values - mean
    return float(np.sqrt(np.nanmean(diffs ** 2)))


def update_recent_std(df: pd.DataFrame, date_columns):
    prices = df[date_columns].to_numpy(dtype=float)
    start_index = max(0, len(date_columns) - RECENT_DAYS)

    for idx_col in range(start_index, len(date_columns)):
        if idx_col < WINDOW - 1:
            continue
        window_values = prices[:, idx_col - WINDOW + 1: idx_col + 1]
        std_values = np.apply_along_axis(recompute_std, 1, window_values)
        df[date_columns[idx_col]] = std_values.round(2)

    return df


def main():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"{EXCEL_PATH} 파일이 없습니다.")

    df = load_std_sheet(EXCEL_PATH)
    date_cols = [c for c in df.columns if c not in ("종목명", "종목코드")]
    if len(date_cols) < WINDOW:
        raise ValueError(f"날짜 열이 {WINDOW}개 이상 필요합니다.")

    updated = update_recent_std(df.copy(), date_cols)

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        updated.to_excel(writer, sheet_name=STD_SHEET, index=False)

    print(f"✅ STD 시트 최신 {RECENT_DAYS}일 갱신 완료: {EXCEL_PATH.name}")


if __name__ == "__main__":
    main()
