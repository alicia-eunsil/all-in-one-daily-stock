"""
File: recompute_std_last10.py
Version: v1.1.0
Role: 최근 10일간 STD 값을 새로운 산식으로 재계산해 원본 파일을 덮어쓴다.
"""

from pathlib import Path
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from zipfile import BadZipFile
import shutil
from datetime import datetime

# helper
def normalize_code(value):
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s:
        return None
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 0:
        return s
    return digits.zfill(6)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "KR_Stocks_Individual.xlsx"
STD_SHEET = "std"
WINDOW = 20
RECENT_DAYS = 10


def load_std_sheet(path: Path):
    df = pd.read_excel(path, sheet_name=STD_SHEET, dtype={"종목코드": str})
    if "종목코드" in df.columns:
        df["종목코드"] = df["종목코드"].apply(normalize_code)
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


def format_date_label(value):
    if isinstance(value, (datetime, pd.Timestamp)):
        return int(value.strftime("%Y%m%d"))

    s = str(value).strip()
    if not s:
        return s

    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 8 and digits.isdigit():
        return int(digits)

    s = s.replace(".", "").replace("-", "").replace("/", "")
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 8 and digits.isdigit():
        return int(digits)

    return s


def normalize_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for col in df.columns:
        if col in ("종목명", "종목코드"):
            renamed[col] = col
        else:
            renamed[col] = format_date_label(col)
    return df.rename(columns=renamed)


def overwrite_std_sheet(path: Path, df: pd.DataFrame):
    """openpyxl로 STD 시트를 안전하게 덮어쓰기"""
    try:
        wb = load_workbook(path)
    except BadZipFile as e:
        raise RuntimeError(f"엑셀 파일이 손상되었습니다: {e}")

    # 기존 STD 시트 제거
    if STD_SHEET in wb.sheetnames:
        ws_old = wb[STD_SHEET]
        idx = wb.sheetnames.index(STD_SHEET)
        wb.remove(ws_old)
    else:
        idx = len(wb.sheetnames)

    ws = wb.create_sheet(STD_SHEET, idx)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    wb.save(path)
    wb.close()


def main():
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"{EXCEL_PATH} 파일이 없습니다.")

    df = load_std_sheet(EXCEL_PATH)
    date_cols = [c for c in df.columns if c not in ("종목명", "종목코드")]
    if len(date_cols) < WINDOW:
        raise ValueError(f"날짜 열이 {WINDOW}개 이상 필요합니다.")

    updated = update_recent_std(df.copy(), date_cols)
    updated = normalize_date_columns(updated)
    if "종목코드" in updated.columns:
        updated["종목코드"] = updated["종목코드"].apply(normalize_code)

    backup = EXCEL_PATH.with_suffix(".bak")
    shutil.copy2(EXCEL_PATH, backup)
    try:
        overwrite_std_sheet(EXCEL_PATH, updated)
    except Exception:
        shutil.move(backup, EXCEL_PATH)
        raise
    else:
        if backup.exists():
            backup.unlink()

    print(f"✅ STD 시트 최신 {RECENT_DAYS}일 갱신 완료: {EXCEL_PATH.name}")


if __name__ == "__main__":
    main()
