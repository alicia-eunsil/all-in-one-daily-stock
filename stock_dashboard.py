"""
File: stock_dashboard.py
Version: v2.0.0
Role: 계산된 주가 지표와 원자료를 조회하는 Streamlit 대시보드.
"""

import streamlit as st
import subprocess
import sys
import pandas as pd
import openpyxl
from pathlib import Path
import bcrypt
from datetime import datetime, date, timedelta
import json  # 🔥 4개 엑셀 매핑용

# ======================================
# 페이지 설정 (최초 UI 출력 전에 호출)
# ======================================
st.set_page_config(page_title="주식 데이터 대시보드", page_icon="🚀", layout="wide")

# ======================================
# 0. 인증 (간단 비밀번호)
# ======================================
ACCESS_CODE_HASH = b"$2b$12$gDBpQYK.g938H.8cNwLeUu/VRidCP1GxqusJiEQzVnvaSrG4CBE6K"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 Access Required")
    st.write("Please enter the access code to open the dashboard.")

    with st.form("auth_form"):
        code = st.text_input("Enter access code", type="password")
        submitted = st.form_submit_button("Submit")

    if submitted:
        if bcrypt.checkpw(code.encode(), ACCESS_CODE_HASH):
            st.session_state["authenticated"] = True
            st.success("Access granted")
            st.rerun()
        else:
            st.error("Invalid code")

    st.stop()

# ======================================
# 1. 전역 상태 변수
# ======================================
if "run_update" not in st.session_state:
    st.session_state.run_update = False
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = True

# 🔥 종합 탭 날짜 확장용
if "show_days" not in st.session_state:
    st.session_state.show_days = 10  # 시작: 최근 10일

# 🔥 원자료 탭 날짜 확장용
if "show_days_raw" not in st.session_state:
    st.session_state.show_days_raw = 10  # 시작: 최근 10일

# 🔥 파일 선택 상태
JSON_PATH = "stock_file_map.json"
EXCEL_MAP = {}
if "selected_category" not in st.session_state:
    # JSON 로드해서 첫 번째 항목을 기본 선택값으로
    try:
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            _tmp_map = json.load(f)
        if isinstance(_tmp_map, dict) and _tmp_map:
            st.session_state.selected_category = list(_tmp_map.keys())[0]
        else:
            st.session_state.selected_category = None
    except:
        st.session_state.selected_category = None


# ======================================
# 2. 날짜/포맷 유틸 함수
# ======================================
def _to_datetime(v):
    """엑셀/문자열/숫자 등 다양한 형태의 날짜를 datetime으로 변환"""
    if isinstance(v, (datetime, date)):
        return datetime(v.year, v.month, v.day)

    if isinstance(v, (int, float)):
        iv = int(v)
        digits = str(iv)
        if len(digits) == 8 and digits.isdigit():
            try:
                return datetime.strptime(digits, "%Y%m%d")
            except:
                pass
        base = datetime(1899, 12, 30)
        try:
            return base + timedelta(days=iv)
        except:
            return None

    s = str(v).strip()
    if not s:
        return None

    for fmt in ("%Y-%m-%d", "%Y.%m.%d.", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except:
            pass

    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 8:
        try:
            return datetime.strptime(digits, "%Y%m%d")
        except:
            pass

    return None


def format_excel_date(v):
    """_to_datetime로 바꾼 날짜를 YYYY.MM.DD. 형식 문자열로 변환"""
    dt = _to_datetime(v)
    if dt:
        return dt.strftime("%Y.%m.%d.")
    s = str(v)
    s = s.replace("-", ".").replace("/", ".")
    if not s.endswith("."):
        s += "."
    return s


def _format_z_cell(v):
    val = pd.to_numeric(v, errors="coerce")
    if pd.isna(val):
        return "-"
    out = f"{val:.0f}"
    if val > 100:
        out += " 🔴"
    elif val < -100:
        out += " 🔵"
    return out


def _format_s_cell(v):
    val = pd.to_numeric(v, errors="coerce")
    if pd.isna(val):
        return "-"
    out = f"{val:.0f}"
    if abs(val - 100) < 0.1:
        out += " 🔴"
    elif abs(val - 0) < 0.1:
        out += " 🔵"
    return out


def _format_q_cell(v):
    val = pd.to_numeric(v, errors="coerce")
    if pd.isna(val):
        return "-"
    out = f"{val:.0f}"
    if val > 100:
        out += " 🔴"
    elif val < 25:
        out += " 🔵"
    return out


GAP_EMOJI_WINDOW = 20
def _apply_gap_emojis(df, columns, window=GAP_EMOJI_WINDOW):
    if not columns:
        return
    consider_cols = columns[-window:]
    for idx in df.index:
        numeric_vals = pd.to_numeric(df.loc[idx, consider_cols], errors="coerce").dropna()
        if numeric_vals.empty:
            continue
        max_val = numeric_vals.max()
        min_val = numeric_vals.min()
        for col in consider_cols:
            val = pd.to_numeric(df.at[idx, col], errors="coerce")
            if pd.isna(val):
                continue
            suffix = ""
            if val == max_val:
                suffix = " 🔴"
            elif val == min_val:
                suffix = " 🔵"
            df.at[idx, col] = f"{val:.0f}{suffix}"

# ======================================
# 3. 뷰 렌더링 함수들
# ======================================
def render_total_view(indicator_df, selected_labels, indicator_range_msg, total_days, index_df=None):
    """
    1️⃣ 종합 탭
    - 멀티헤더(날짜×지표) 구조
    - 맨 아래 평균 행
    - 그 아래 KOSPI/KOSDAQ/KOSPI200 행 추가
    """
    if indicator_df is None:
        st.warning("⚠️ 종합 데이터를 불러올 수 없습니다.")
        return

    st.markdown("### 🔍 필터 옵션 (종합)")
    c1, c2 = st.columns(2)
    with c1:
        search = st.text_input("🔎 종목명/종목코드 검색", key="search_total")
    with c2:
        sort_by = st.selectbox("정렬 기준", ["종목코드", "종목명"], key="sort_total")

    # 검색 적용
    df_f = indicator_df.copy()
    if search:
        df_f = df_f[
            df_f["종목명"].astype(str).str.contains(search, case=False) |
            df_f["종목코드"].astype(str).str.contains(search, case=False)
        ]

    df_f = df_f.sort_values(by=sort_by)

    st.info(indicator_range_msg)

    # --------------------------------------
    # 🔥 멀티헤더 생성 (1행: 날짜, 2행: 지표명)
    # --------------------------------------
    metrics = ["Z20", "Z60", "Z120", "S20", "S60", "S120", "GAP", "QUANT", "STD"]
    base_cols = ["종목코드", "종목명"]
    df_show = df_f[base_cols].copy()

    col_tuples = [("", "종목코드"), ("", "종목명")]

    # 날짜 × 지표 조합 생성 (값 없으면 '-' 처리)
    for lbl in selected_labels:
        for m in metrics:
            key = (lbl, m)
            if key in df_f.columns:
                df_show[(lbl, m)] = df_f[key]
            else:
                df_show[(lbl, m)] = "-"
            col_tuples.append((lbl, m))

    df_show.columns = pd.MultiIndex.from_tuples(col_tuples)

    # --------------------------------------
    # 🔥 평균 행 추가 (맨 마지막 행)
    # --------------------------------------
    avg_row = []
    for col in df_show.columns:
        if col == ("", "종목코드"):
            avg_row.append("AVG")
        elif col == ("", "종목명"):
            avg_row.append("평균")
        else:
            lbl, m = col
            key = (lbl, m)
            if key in df_f.columns:
                s = pd.to_numeric(df_f[key], errors="coerce")
                avg_val = s.mean(skipna=True)
                avg_row.append(f"{avg_val:.2f}")
            else:
                avg_row.append(None)

    df_show.loc[len(df_show)] = avg_row  # 평균 행 추가

    # --------------------------------------
    # Z/S/Q/GAP 포맷 이모지 적용
    # --------------------------------------
    for lbl in selected_labels:
        for m in ["Z20", "Z60", "Z120"]:
            col = (lbl, m)
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(_format_z_cell)

        for m in ["S20", "S60", "S120"]:
            col = (lbl, m)
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(_format_s_cell)

        for metric in ["GAP", "STD"]:
            col = (lbl, metric)
            if col in df_show.columns:
                df_show[col] = pd.to_numeric(df_show[col], errors="coerce")

        for m in ["QUANT"]:
            col = (lbl, m)
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(_format_q_cell)

    gap_columns_total = [(lbl, "GAP") for lbl in selected_labels if (lbl, "GAP") in df_show.columns]
    _apply_gap_emojis(df_show, gap_columns_total)

    # --------------------------------------
    # 🔽 지수(KOSPI/KOSDAQ/KOSPI200) 행 추가
    # --------------------------------------
    if index_df is not None and not index_df.empty:
        for _, idx_row in index_df.iterrows():
            new_row_vals = []
            used_dates = set()  # 같은 날짜에 한 번만 값 넣기 위한 기록

            for col in df_show.columns:
                if col == ("", "종목코드"):
                    new_row_vals.append(idx_row.get("업종코드", ""))
                elif col == ("", "종목명"):
                    new_row_vals.append(idx_row.get("업종명", ""))
                else:
                    lbl, m = col
                    if lbl not in used_dates:
                        val = idx_row.get(lbl, None)
                        new_row_vals.append(val if pd.notna(val) else "")
                        used_dates.add(lbl)
                    else:
                        new_row_vals.append("")

            df_show.loc[len(df_show)] = new_row_vals  # 지수 행 추가

    # 인덱스 설정 (종목코드·종목명)
    df_show = df_show.set_index([("", "종목코드"), ("", "종목명")])

    st.dataframe(
        df_show,
        width="stretch",
        height=600,
    )

    # 🔥 과거 확장 버튼 (종합)
    if st.button("⬅ 과거 10일 더보기(종합)", disabled=(total_days <= st.session_state.show_days)):
        st.session_state.show_days = min(st.session_state.show_days + 10, total_days)
        st.rerun()


def render_metric_view(indicator_df, selected_labels):
    """
    2️⃣ 지표별 탭:
    - 1열: 종목코드
    - 2열: 종목명
    - 이후: 날짜별 선택 지표값
    """
    st.subheader("📈 지표 선택")

    if indicator_df is None or len(indicator_df) == 0:
        st.warning("⚠️ 지표별 데이터를 불러올 수 없습니다.")
        return

    metric_options = ["Z20", "Z60", "Z120",
                      "S20", "S60", "S120",
                      "GAP", "QUANT", "STD"]

    # 실제 존재하는 지표만
    available = []
    for m in metric_options:
        if any(((lbl, m) in indicator_df.columns) for lbl in selected_labels):
            available.append(m)

    if not available:
        st.error("indicator_df에 S/Z/GAP/QUANT/STD 관련 컬럼이 없습니다.")
        st.write("현재 indicator_df.columns 예시:", list(indicator_df.columns)[:20])
        return

    metric = st.selectbox("지표를 선택하세요", available, index=0)

    # -------------------------
    # DF 구성 (종목코드, 종목명 + 날짜별 값)
    # -------------------------
    df_metric_numeric = indicator_df[["종목코드", "종목명"]].copy()

    for lbl in selected_labels:
        col_key = (lbl, metric)
        if col_key in indicator_df.columns:
            df_metric_numeric[lbl] = indicator_df[col_key]
        else:
            df_metric_numeric[lbl] = None

    # 값 포맷팅
    def _format_plain(v):
        val = pd.to_numeric(v, errors="coerce")
        if pd.isna(val):
            return "-"
        return f"{val:.0f}"

    def _format_std_cell(v):
        val = pd.to_numeric(v, errors="coerce")
        if pd.isna(val):
            return "-"
        return f"{val:.2f}"

    if metric == "STD":
        formatter = _format_std_cell
    elif metric.startswith("S"):
        formatter = _format_s_cell
    elif metric.startswith("Z"):
        formatter = _format_z_cell
    elif metric == "QUANT":
        formatter = _format_q_cell
    else:
        formatter = _format_plain

    df_metric_display = df_metric_numeric.copy()
    for lbl in selected_labels:
        if lbl in df_metric_display.columns:
            df_metric_display[lbl] = df_metric_display[lbl].apply(formatter)

    # 🔍 필터 + 정렬
    st.markdown("### 🔍 필터 옵션 (지표별)")
    c1, c2 = st.columns(2)
    with c1:
        search_metric = st.text_input("🔎 종목명/종목코드 검색", key="search_metric")
    with c2:
        sort_metric = st.selectbox("정렬 기준", ["종목코드", "종목명"], key="sort_metric")

    df_filtered_display = df_metric_display.copy()
    df_filtered_numeric = df_metric_numeric.copy()
    if search_metric:
        mask = (
            df_filtered_display["종목명"].astype(str).str.contains(search_metric, case=False)
            | df_filtered_display["종목코드"].astype(str).str.contains(search_metric, case=False)
        )
        df_filtered_display = df_filtered_display[mask]
        df_filtered_numeric = df_filtered_numeric[mask]

    sort_order = df_filtered_display.sort_values(by=sort_metric).index
    df_filtered_display = df_filtered_display.loc[sort_order].reset_index(drop=True)
    df_filtered_numeric = df_filtered_numeric.loc[sort_order].reset_index(drop=True)

    if metric == "GAP":
        gap_columns_metric = [lbl for lbl in selected_labels if lbl in df_filtered_display.columns]
        _apply_gap_emojis(df_filtered_display, gap_columns_metric)

    # 평균 행 추가
    avg_row = {"종목코드": "AVG", "종목명": "평균"}
    for lbl in selected_labels:
        if lbl not in df_filtered_numeric.columns:
            continue
        col_numeric = pd.to_numeric(df_filtered_numeric[lbl], errors="coerce")
        mean_val = col_numeric.mean(skipna=True)
        if pd.isna(mean_val):
            avg_row[lbl] = "-"
        else:
            avg_row[lbl] = formatter(mean_val)

    df_filtered = df_filtered_display.copy()
    df_filtered.loc[len(df_filtered)] = avg_row

    # 날짜 범위 안내
    if selected_labels:
        oldest_label = selected_labels[0]
        latest_label = selected_labels[-1]
        st.info(
            f"📅 지표별 표시 범위: **{oldest_label} ~ {latest_label}** "
            f"(최근 {len(selected_labels)}일)"
        )

    # 테이블 출력
    st.markdown(f"### 📋 {metric} · 추이")

    column_config = {
        "종목코드": st.column_config.TextColumn("종목코드", width="small", pinned="left"),
        "종목명": st.column_config.TextColumn("종목명", width="small", pinned="left"),
    }
    for lbl in selected_labels:
        if lbl in df_filtered.columns:
            column_config[lbl] = st.column_config.TextColumn(lbl)

    st.dataframe(
        df_filtered,
        width="stretch",
        height=600,
        hide_index=True,
        column_config=column_config,
    )

    # 🔥 과거 확장 버튼 (지표별)
    global total_days
    if st.button("⬅ 과거 10일 더보기(지표별)", disabled=(total_days <= st.session_state.show_days)):
        st.session_state.show_days = min(st.session_state.show_days + 10, total_days)
        st.rerun()


def render_raw_view(close_df, close_range_msg, total_close_days):
    """
    3️⃣ 원자료(종가) 탭
    - 종목코드/종목명 + 날짜별 종가
    """
    if close_df is None:
        st.warning("⚠️ 원자료(종가) 데이터를 불러올 수 없습니다.")
        return

    st.markdown("### 🔍 필터 옵션 (원자료)")
    r1, r2 = st.columns(2)
    with r1:
        search_raw = st.text_input("🔎 종목명/종목코드 검색", key="search_raw")
    with r2:
        sort_raw = st.selectbox("정렬 기준", ["종목코드", "종목명"], key="sort_raw")

    df_raw = close_df.copy()

    if search_raw:
        df_raw = df_raw[
            df_raw["종목코드"].astype(str).str.contains(search_raw, case=False) |
            df_raw["종목명"].astype(str).str.contains(search_raw, case=False)
        ]

    df_raw = df_raw.sort_values(by=sort_raw)

    st.info(close_range_msg)

    # 날짜 컬럼 추출
    date_cols = [c for c in df_raw.columns if c not in ["종목코드", "종목명"]]

    # 컬럼 순서 고정
    df_raw = df_raw[["종목코드", "종목명"] + date_cols]

    # 날짜 컬럼 숫자 변환
    column_config = {
        "종목코드": st.column_config.TextColumn("종목코드", width="small", pinned="left"),
        "종목명": st.column_config.TextColumn("종목명", width="small", pinned="left"),
    }
    for c in date_cols:
        df_raw[c] = pd.to_numeric(df_raw[c], errors="coerce")
        has_decimal = df_raw[c].dropna().apply(lambda v: not float(v).is_integer()).any()
        number_format = "%.2f" if has_decimal else "%.0f"
        column_config[c] = st.column_config.NumberColumn(c, format=number_format)

    st.dataframe(
        df_raw,
        width="stretch",
        height=600,
        hide_index=True,
        column_config=column_config,
    )

    # 🔥 과거 확장 버튼 (원자료)
    if st.button("⬅ 과거 10일 더보기(종가)", disabled=(total_close_days <= st.session_state.show_days_raw)):
        st.session_state.show_days_raw = min(st.session_state.show_days_raw + 10, total_close_days)
        st.rerun()


# ======================================
# 4. 엑셀 파일 매핑(JSON) 로드
# ======================================
try:
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        EXCEL_MAP = json.load(f)
    if not isinstance(EXCEL_MAP, dict) or not EXCEL_MAP:
        st.error("stock_files.json 형식이 잘못되었거나 비어 있습니다.")
        st.stop()
except FileNotFoundError:
    st.error("stock_files.json 파일을 찾을 수 없습니다.")
    st.stop()
except Exception as e:
    st.error(f"stock_files.json 읽기 오류: {e}")
    st.stop()

categories = list(EXCEL_MAP.keys())
if not categories:
    st.error("stock_file_map.json 에 항목이 없습니다.")
    st.stop()

if st.session_state.selected_category not in categories:
    st.session_state.selected_category = categories[0]

# ======================================
# 5. 상단: 네 개 카테고리 선택 버튼 (라디오, 가로)
# ======================================
st.markdown("### 📂 조회할 주식 그룹 선택")

selected_category = st.radio(
    "주식 그룹",
    categories,
    index=categories.index(st.session_state.selected_category),
    horizontal=True,
    label_visibility="collapsed",
)
st.session_state.selected_category = selected_category
selected_filename = EXCEL_MAP[selected_category]
excel_path = Path(selected_filename)

#st.markdown(f"#### 현재: `{selected_category}`")

# ======================================
# 6. 사이드바: 선택 파일 다운로드 + 전체 갱신 버튼
# ======================================
with st.sidebar:
    st.markdown("### 📁 현재 선택 파일")
    st.write(f"`{selected_filename}`")

    if excel_path.exists():
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📥 선택 파일 다운로드",
                data=f,
                file_name=excel_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel",
            )
    else:
        st.warning(f"`{selected_filename}` 파일이 아직 생성되지 않았습니다.")

    st.markdown("---")
    if st.button("🔄 네 개 파일 전체 데이터 갱신"):
        st.session_state.run_update = True

# ======================================
# 7. 데이터 갱신 실행 (외부 스크립트 호출)
# ======================================
if st.session_state.run_update:
    with st.sidebar:
        st.subheader("진행 상황")
        status_box = st.empty()

    scripts = [
        ("run_all_scores.py", "4개 엑셀 S/Z + GAP/QUANT/STD 계산"),
    ]

    for idx, (sc, desc) in enumerate(scripts):
        status_box.info(f"{desc} 실행 중...")
        try:
            result = subprocess.run(
                [sys.executable, sc], capture_output=True, text=True, timeout=1800
            )
            if result.returncode == 0:
                st.sidebar.success(f"{desc} 완료")
            else:
                st.sidebar.error(f"{desc} 실패")
                st.sidebar.code(result.stderr[:500])
        except Exception as e:
                st.sidebar.error(f"{desc} 오류 발생: {e}")

        status_box.success(f"{desc} 완료")

    st.session_state.data_loaded = True
    st.session_state.run_update = False
    st.rerun()

# ======================================
# 8. 선택된 엑셀 파일 로드
# ======================================
if not excel_path.exists():
    st.error(f"`{selected_filename}` 파일을 찾지 못했습니다. "
             "왼쪽의 '네 개 파일 전체 데이터 갱신' 버튼으로 먼저 데이터를 생성해 주세요.")
    st.stop()

excel_file = excel_path
wb = openpyxl.load_workbook(excel_file, data_only=True)

# ======================================
# 9. 종목 정보 로딩 (종목 시트)
# ======================================
stock_info = {}
if "종목" in wb.sheetnames:
    ws = wb["종목"]
    for r in ws.iter_rows(min_row=2, max_col=2):
        name = r[0].value
        code = r[1].value
        if code and name:
            stock_info[code] = name

# ======================================
# 10. 종합(Z20/Z60/S/GAP/QUANT/STD) 데이터 로딩
# ======================================
sheet_names = ["z20", "z60", "z120", "s20", "s60", "s120", "gap", "quant", "std"]

base_ws = None
for s in sheet_names:
    if s in wb.sheetnames:
        base_ws = wb[s]
        break

indicator_df = None
indicator_date_infos = []
total_days = 0
selected_labels = []
indicator_range_msg = ""

if base_ws:
    max_col = base_ws.max_column

    # 기준 시트에서 날짜 헤더 수집 (1행, 3열~)
    for col in range(3, max_col + 1):
        raw = base_ws.cell(row=1, column=col).value
        if raw is None:
            continue
        dt = _to_datetime(raw)
        label = format_excel_date(raw)
        indicator_date_infos.append((col, raw, dt, label))

    indicator_date_infos = sorted(
        indicator_date_infos,
        key=lambda x: (x[2] is None, x[2] or datetime.min)
    )

    total_days = len(indicator_date_infos)

    show_days = min(st.session_state.show_days, total_days)
    start_idx = total_days - show_days
    selected_infos = indicator_date_infos[start_idx:]
    selected_labels = [lbl for _, _, _, lbl in selected_infos]

    oldest_label = selected_infos[0][3]
    latest_label = selected_infos[-1][3]
    indicator_range_msg = (
        f"📅 종합 표시 범위: **{oldest_label} ~ {latest_label}** "
        f"(최근 {show_days}일 / 전체 {total_days}일)"
    )

    # 종목별 데이터 딕셔너리 구성
    data_dict = {code: {"종목코드": code, "종목명": name}
                 for code, name in stock_info.items()}

    # 시트별 데이터 채우기 (날짜 문자열 기준 매칭)
    for s in sheet_names:
        if s not in wb.sheetnames:
            continue

        ws = wb[s]
        max_row_s = ws.max_row
        max_col_s = ws.max_column

        label_to_col = {}
        for col in range(3, max_col_s + 1):
            raw = ws.cell(row=1, column=col).value
            if raw is None:
                continue
            lbl = format_excel_date(raw)
            label_to_col[lbl] = col

        for r in range(2, max_row_s + 1):
            code = ws.cell(row=r, column=2).value
            if code not in data_dict:
                continue

            for lbl in selected_labels:
                col_idx = label_to_col.get(lbl)
                if col_idx is None:
                    val = None
                else:
                    val = ws.cell(row=r, column=col_idx).value

                data_dict[code][(lbl, s.upper())] = val

    indicator_df = pd.DataFrame.from_dict(data_dict, orient="index").reset_index(drop=True)

else:
    indicator_df = None

# ======================================
# 11. 원자료(종가) 데이터 로딩
# ======================================
close_df = None
close_date_infos = []
total_close_days = 0
close_range_msg = ""

if "종가" in wb.sheetnames:
    ws = wb["종가"]
    max_col_c = ws.max_column

    # 날짜 헤더
    for col in range(3, max_col_c + 1):
        raw = ws.cell(row=1, column=col).value
        if raw is None:
            continue

        dt = _to_datetime(raw)

        if dt is None:
            digits = "".join(ch for ch in str(raw) if ch.isdigit())
            if len(digits) == 8:
                dt = datetime.strptime(digits, "%Y%m%d")

        if dt is None:
            continue

        label = dt.strftime("%Y.%m.%d.")
        close_date_infos.append((col, raw, dt, label))

    close_date_infos = sorted(
        close_date_infos,
        key=lambda x: (x[2] is None, x[2] or datetime.min)
    )

    total_close_days = len(close_date_infos)

    show_raw = min(st.session_state.show_days_raw, total_close_days)
    start_idx = total_close_days - show_raw
    selected_close_infos = close_date_infos[start_idx:]

    oldest_label = selected_close_infos[0][3]
    latest_label = selected_close_infos[-1][3]

    close_range_msg = (
        f"📅 종가 표시 범위: **{oldest_label} ~ {latest_label}** "
        f"(최근 {show_raw}일 / 전체 {total_close_days}일)"
    )

    close_dict = {code: {"종목명": name, "종목코드": code}
                  for code, name in stock_info.items()}

    max_row_c = ws.max_row

    for r in range(2, max_row_c + 1):
        code = ws.cell(row=r, column=2).value
        if code not in close_dict:
            continue

        for col_idx, raw, dt, label in selected_close_infos:
            val = ws.cell(row=r, column=col_idx).value
            close_dict[code][label] = val

    close_df = pd.DataFrame.from_dict(close_dict, orient="index").reset_index(drop=True)

    # 컬럼 이름을 yyyy.mm.dd. 형식으로 통일
    rename_map = {}
    for col in close_df.columns:
        if col in ["종목코드", "종목명"]:
            continue
        rename_map[col] = format_excel_date(col)

    close_df = close_df.rename(columns=rename_map)

# ======================================
# 12. 지수(KOSPI/KOSDAQ/KOSPI200) 데이터 로딩
# ======================================
index_df = None

if "지수" in wb.sheetnames and indicator_df is not None and selected_labels:
    ws_idx = wb["지수"]
    max_col_i = ws_idx.max_column

    index_date_infos = []
    for col in range(3, max_col_i + 1):
        raw = ws_idx.cell(row=1, column=col).value
        if raw is None:
            continue

        dt = _to_datetime(raw)
        if dt is None:
            digits = "".join(ch for ch in str(raw) if ch.isdigit())
            if len(digits) == 8:
                dt = datetime.strptime(digits, "%Y%m%d")
        if dt is None:
            continue

        label = dt.strftime("%Y.%m.%d.")
        index_date_infos.append((col, raw, dt, label))

    label_to_col_idx = {label: col for col, raw, dt, label in index_date_infos}

    index_rows = []
    max_row_i = ws_idx.max_row

    seen_codes = set()
    for r in range(2, max_row_i + 1):
        name = ws_idx.cell(row=r, column=1).value
        code = ws_idx.cell(row=r, column=2).value
        if not name or not code:
            continue

        code_str = str(code).strip()
        if code_str.isdigit():
            code_str = code_str.zfill(6)
        if code_str in seen_codes:
            continue
        seen_codes.add(code_str)

        row_dict = {
            "업종명": str(name),
            "업종코드": code_str,
        }

        for lbl in selected_labels:
            col_idx = label_to_col_idx.get(lbl)
            if col_idx is None:
                val = None
            else:
                val = ws_idx.cell(row=r, column=col_idx).value
            row_dict[lbl] = val

        index_rows.append(row_dict)

    if index_rows:
        index_df = pd.DataFrame(index_rows)

# ======================================
# 13. 엑셀 파일 닫기
# ======================================
wb.close()

# ======================================
# 14. 탭 구성 및 렌더링
# ======================================
tab_total, tab_metric, tab_raw = st.tabs(["1️⃣ 종합", "2️⃣ 지표별", "3️⃣ 원자료"])

with tab_total:
    if indicator_df is None:
        st.warning("⚠️ 종합 데이터를 불러올 수 없습니다.")
    else:
        render_total_view(
            indicator_df,
            selected_labels,
            indicator_range_msg,
            total_days,
            index_df=index_df,
        )

with tab_metric:
    if indicator_df is None:
        st.warning("⚠️ 지표별 데이터를 불러올 수 없습니다.")
    else:
        render_metric_view(indicator_df, selected_labels)

with tab_raw:
    if close_df is None:
        st.warning("⚠️ 원자료(종가) 데이터를 불러올 수 없습니다.")
    else:
        render_raw_view(close_df, close_range_msg, total_close_days)

st.markdown("---")
st.caption("Created by Alicia")
