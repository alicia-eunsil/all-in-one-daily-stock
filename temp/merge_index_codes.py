import re
import shutil
from pathlib import Path
from typing import Optional

import pandas as pd


# ================== 설정 ==================
FILENAME = "KR_Stocks_Individual.xlsx"
SHEET_NAME = "지수"

# 지수 시트 컬럼 후보 (파일 버전 차이 대응)
NAME_COL_CANDIDATES = ["종목명", "업종명"]
CODE_COL_CANDIDATES = ["종목코드", "업종코드"]

# 안전용 백업 (원하면 True)
MAKE_BACKUP = False
BACKUP_SUFFIX = ".backup"
# =========================================


def clean_code(x) -> str:
    """숫자 코드 → 6자리 통일"""
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x).strip()
    if s == "" or s.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s2 = re.sub(r"[,\s]", "", s)
    if s2.isdigit():
        return s2.zfill(6)
    return s


def pick_existing_col(columns, candidates) -> Optional[str]:
    for c in candidates:
        if c in columns:
            return c
    return None


def choose_rep_index(grp: pd.DataFrame, code_col: str) -> int:
    """대표행: 6자리 코드 우선"""
    codes = grp[code_col].astype(str)

    m1 = codes.apply(lambda s: s.isdigit() and len(s) == 6)
    if m1.any():
        return grp[m1].index[0]

    m2 = codes.apply(lambda s: s.strip() != "" and s.lower() != "nan")
    if m2.any():
        return grp[m2].index[0]

    return grp.index[0]


def merge_index_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()

    name_col = pick_existing_col(df.columns, NAME_COL_CANDIDATES)
    code_col = pick_existing_col(df.columns, CODE_COL_CANDIDATES)

    if not name_col or not code_col:
        raise ValueError(
            f"지수 시트에서 종목명/종목코드 컬럼을 찾지 못했습니다.\n"
            f"현재 컬럼: {list(df.columns)}"
        )

    df[name_col] = df[name_col].astype(str).str.strip()
    df[code_col] = df[code_col].apply(clean_code)

    df = df[df[name_col].str.strip().ne("")]

    merged = []
    for _, grp in df.groupby(name_col, sort=False):
        if len(grp) == 1:
            merged.append(grp.iloc[0].to_dict())
            continue

        rep_idx = choose_rep_index(grp, code_col)
        rep = grp.loc[rep_idx].copy()

        for col in grp.columns:
            if col == name_col:
                continue
            v = rep[col]
            empty = (
                v is None
                or (isinstance(v, float) and pd.isna(v))
                or (isinstance(v, str) and v.strip() == "")
            )
            if empty:
                for cand in grp[col]:
                    if cand is None or (isinstance(cand, float) and pd.isna(cand)):
                        continue
                    if isinstance(cand, str) and cand.strip() == "":
                        continue
                    rep[col] = cand
                    break

        merged.append(rep.to_dict())

    return pd.DataFrame(merged)


def main():
    base_dir = Path(__file__).resolve().parent
    excel_path = base_dir / FILENAME

    if not excel_path.exists():
        raise FileNotFoundError(f"{excel_path} 파일이 없습니다.")

    if MAKE_BACKUP:
        backup = excel_path.with_name(f"{excel_path.stem}{BACKUP_SUFFIX}{excel_path.suffix}")
        if not backup.exists():
            shutil.copy2(excel_path, backup)

    xls = pd.ExcelFile(excel_path)
    if SHEET_NAME not in xls.sheet_names:
        raise ValueError(f"'지수' 시트를 찾을 수 없습니다.")

    index_df = pd.read_excel(excel_path, sheet_name=SHEET_NAME, dtype=object)
    merged_df = merge_index_df(index_df)

    tmp_path = excel_path.with_suffix(".tmp.xlsx")
    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        for sh in xls.sheet_names:
            if sh == SHEET_NAME:
                merged_df.to_excel(writer, sheet_name=sh, index=False)
            else:
                pd.read_excel(excel_path, sheet_name=sh, dtype=object).to_excel(
                    writer, sheet_name=sh, index=False
                )

    tmp_path.replace(excel_path)
    print("지수 시트 병합 완료 → 원본 파일 업데이트")


if __name__ == "__main__":
    main()
