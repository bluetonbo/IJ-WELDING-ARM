# -*- coding: utf-8 -*-
"""
WELDING-ARM-KW : 용접 암 측정 데이터(DATA-KW.xlsx) -> 전장 범위 내 일자-X / 휨-X Min·Max 추출기
------------------------------------------------------------
- 측정 데이터 엑셀(2행: 항목 머리글, 3행~: 시료별 측정값)을 업로드
- '전장' 측정값이 지정한 범위(기본 274.8 ~ 275.2, 경계 포함) 안에 있는 시료만 골라서
  '일자 X'와 '휨 X'의 Min / Max (측정값과 편차 둘 다)와 해당 시료(파일명, 엑셀 행)를 찾는다
- 엑셀 하단의 MIN / MAX / AVG 요약 행, '0.8 미만' 같은 구분 행은 시료가 아니므로
  파일명 패턴(기본: '.xls' 또는 'QMS' 포함)으로 걸러내서 계산에서 제외한다
- UI/로그인/임시비밀번호 인프라/사이드바 업로드/상단 탭 구조는 WELDING-ARM-KW 원본과
  동일하게 유지한다 (차콜+앰버 테마)
"""

import io
import json
import base64
from datetime import datetime, timedelta
from collections import Counter

import re as _re
import streamlit as st
import pandas as pd
from openpyxl.utils import get_column_letter

# ============================================================
# 0. 기본 설정 / 스타일 (BA-CAPA 콘솔과 동일한 테마 + 로그인 방식)
# ============================================================
st.set_page_config(page_title="WELDING-ARM-KW | 전장 범위 Min/Max 분석", layout="wide", initial_sidebar_state="expanded")

OWNER_PASSWORD = "nt1234"

# ILJIN BEARING ART CI 로고 (PNG, base64 인라인 - 별도 파일 배포 없이 앱 하나로 동작)
ILJIN_LOGO_B64 = "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KYR7HY/Hi4p6UDkqFqIoKMh1OA6oIGOB4EEGcSAg4IQaIAhCwASIQAjJ+aO6TUI63atq19q11trv53n6eZRetepNp7O712/Xr2oI6+cAa9JdH8WZj/R9VusX9X3O/noTzpxbOoTUxPrPA+8uHaNGt+DMdn2d0fq1gH/0dc7+6//rgvV7AVf3dU4p5WrA4MzTpYOIdN2iRV3/NVZERJpshdIBRERERCrYG/hk6RAiIiIiUi8VqkRERKQt3of100qHEBEREZH6qFAlIiIibXIu1m9ROoSIiIiI1EOFKhEREWmTZwEXYP0apYOIiIiISH4qVImIiEjbbAWcUzqEiIiIiOSnQpWIiIi00RFYf1LpECIiIiKSlwpVIiIi0lafxvrdS4cQERERkXxUqBIREZG2WhH4HtavXzqIiIiIiOShQpWIiIi02XOA72L9xNJBRERERKR3KlSJiIhI2+0GnFY6hIiIiIj0ToUqERER6YITsP7I0iFEREREpDcqVImIiEhXnIP125QOISIiIiLpVKgSERGRrlgNuADrJ5UOIiIiIiJpVKgSERGRLnkhcB7WD5UOIiIiIiLxVKgSERGRrpkKvL90CBERERGJp0KViIiIdNEnsP6VpUOIiIiISBwVqkRERKSLVgC+g/Ublg4iIiIiItWpUCUiIiJdtQ7wfaxfuXQQEREREalGhSoRERHpsh2B00uHEBEREZFqViwdQERERJbyNDCrdIgxbABMLB0i0tux/gacOa90EBEREREZmwpVIiIizTILZzYuHWK5rJ8BvKR0jARfxvpbcebm0kFEREREZPnU+iciIiKDYBXgB1g/uXQQEREREVk+FapERPpjbukAIsImwDexXr//iIiIiDSUflETEekHZxaVjiAiALwa+HDpECIiIiIyOhWqREREZNCcjPUHlA4hIiIiIstSoUpEREQGzRDw/7F+k9JBRERERGRpKlSJiIjIIJpM2Fx91dJBRERERGQxFapERERkUL0UOLN0CBERERFZTIUqERERGWRvxPq9S4cQERERkUCFKhERERl0a5YOICIiIiKBClUiIiIiIiIiItIIKlSJiIiIiIiIiEgjqFAlIiIiIiIiIiKNsGLpACIiIiIi0hxDQ0OlI7TaokWLSkcQEWk1ragSEREREREREZFGUKFKREREREREREQaQYUqERERERERERFpBBWq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"

BG_DARK = "#000000"
SURFACE = "#000000"
SURFACE_ALT = "#000000"
BORDER = "#3a3a3a"
BORDER_STRONG = "#454545"
TEXT_MAIN = "#f2f2f2"
TEXT_MUTED = "#9a9a9a"
ACCENT_BLUE = "#f59e0b"
DANGER = "#ef4444"
WARNING = "#fb923c"
SUCCESS = "#10b981"

CUSTOM_CSS = f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Inter:wght@400;500;600;700&display=swap');

    html {{ font-size: 90% !important; }}
    .stApp {{ background-color: {BG_DARK} !important; color: {TEXT_MAIN} !important; font-family: 'Inter', sans-serif; }}
    [data-testid="stAppViewContainer"] .main .block-container {{ max-width: 100% !important; padding-left: 2rem !important; padding-right: 2rem !important; }}

    h1, h2, h3, h4 {{ font-family: 'Inter', sans-serif; font-weight: 600 !important; color: {TEXT_MAIN} !important; }}
    div[data-testid="stMetricValue"] {{ color: {ACCENT_BLUE}; }}

    .block-title {{
        color: {ACCENT_BLUE}; font-size: 0.9rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.04em; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 1px solid {BORDER_STRONG};
    }}
    .warn-banner {{
        background-color: {SURFACE}; border-left: 4px solid {WARNING}; padding: 10px 14px;
        border-radius: 4px; margin-bottom: 10px; font-size: 14px; color: {TEXT_MAIN};
    }}
    .kpi-card {{ background-color: {SURFACE}; border-radius: 6px; padding: 14px 16px; border: 1px solid {BORDER}; }}
    .kpi-label {{ font-size: 12px; color: {TEXT_MUTED}; margin: 0 0 4px 0; }}
    .kpi-value {{ font-size: 24px; font-weight: 600; margin: 0; color: {TEXT_MAIN}; }}

    .glass-card {{
        background: rgba(0,0,0,0.75); backdrop-filter: blur(6px);
        border: 1px solid {BORDER_STRONG}; border-radius: 12px;
    }}

    div.stButton > button, div.stDownloadButton > button {{
        height: 2.6rem !important; background: {SUCCESS} !important; color: #ffffff !important;
        border-radius: 4px !important; border: none !important; font-weight: 600; width: 100%;
    }}
    div.stButton > button:hover {{ opacity: 0.9; }}

    label, .stTextInput label, .stSelectbox label, .stFileUploader label, [data-testid="stWidgetLabel"] p {{ color: {TEXT_MUTED} !important; }}
    div[data-testid="stSelectbox"] > div > div, div[data-baseweb="select"] > div {{ background-color: #000000 !important; border-color: {BORDER_STRONG} !important; }}
    div[data-testid="stSelectbox"] span, div[data-baseweb="select"] span {{ color: {TEXT_MAIN} !important; }}
    [data-testid="stFileUploaderDropzone"] {{ background-color: {SURFACE} !important; border: 1px solid {BORDER} !important; }}
    [data-testid="stAlert"] p, [data-testid="stAlert"] * {{ color: {TEXT_MAIN} !important; }}
    .stCaption, [data-testid="stCaptionContainer"] {{ color: {TEXT_MUTED} !important; }}
    [data-testid="stExpander"] {{ border: 1.5px solid {ACCENT_BLUE} !important; border-radius: 8px !important; background: {SURFACE} !important; margin-top: 8px; margin-bottom: 8px; box-shadow: 0 0 10px rgba(245,158,11,0.10); }}
    [data-testid="stExpander"] summary {{ font-weight: 700 !important; color: {ACCENT_BLUE} !important; }}

    /* 사이드바 콘솔 스타일 (BA-CAPA와 동일) */
    section[data-testid="stSidebar"] {{ background-color: {SURFACE_ALT} !important; border-right: 1px solid {BORDER}; min-width: 300px !important; }}
    section[data-testid="stSidebar"] * {{ color: {TEXT_MAIN} !important; }}
    section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.2); }}
    section[data-testid="stSidebar"] div.stButton > button {{
        background-color: rgba(255,255,255,0.10); color: white; border: 1px solid rgba(255,255,255,0.3);
    }}

    /* 상단 탭 스타일 (BA-CAPA와 동일) */
    .stTabs [data-baseweb="tab-list"] {{ border-bottom: 2px solid {BORDER_STRONG}; gap: 5px; flex-wrap: wrap; }}
    .stTabs [data-baseweb="tab"], .stTabs button[data-baseweb="tab"], .stTabs [role="tab"] {{
        background-color: #000000 !important; border: 1px solid {BORDER_STRONG} !important;
        border-bottom: none !important; border-radius: 6px 6px 0 0 !important;
        color: {TEXT_MAIN} !important; font-weight: 700 !important; opacity: 1 !important;
        padding: 7px 12px !important; font-size: 0.82rem !important;
    }}
    .stTabs [data-baseweb="tab"] p {{ font-size: 0.82rem !important; }}
    .stTabs [data-baseweb="tab"] * {{ color: {TEXT_MAIN} !important; opacity: 1 !important; }}
    .stTabs [aria-selected="true"], .stTabs button[aria-selected="true"] {{
        background-color: #5c3d0a44 !important; border-color: {ACCENT_BLUE} !important; color: {ACCENT_BLUE} !important;
    }}
    .stTabs [aria-selected="true"] * {{ color: {ACCENT_BLUE} !important; }}
    [data-baseweb="tab"] {{ opacity: 1 !important; }}

    /* OCR 진행 상황 모달(팝업) 스타일 - 다크네이비 알약형 배지 */
    .progress-pill {{
        background: linear-gradient(180deg, #1e2a44, #111a2e);
        border: 1px solid #2d3f5c; border-radius: 999px;
        padding: 9px 18px; text-align: center; font-size: 0.85rem;
        color: #93c5fd; box-shadow: 0 4px 14px rgba(0,0,0,0.45);
        margin: 4px 0 4px 0;
    }}
    .progress-pill b {{ color: #60a5fa; }}
    .progress-pill .pct {{ color: #60a5fa; font-weight: 700; }}
    .progress-pill.done {{
        background: linear-gradient(180deg, #14351f, #0d2416);
        border: 1px solid #1f5c33; color: #86efac;
    }}
    .progress-pill.done b {{ color: {SUCCESS}; }}

    /* "처리 중" 박스는 st.dialog가 아니라 _run_with_progress()가 직접 그리는 div라서,
       그 스타일은 여기(전역 CSS)가 아니라 그 함수 안에서 인라인으로 정의한다 - 배경(바깥
       클릭해도 안 닫힘) 포함 전부 그 함수 책임. */
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ============================================================
# 0-1. 로그인 / Google Sheets 임시비밀번호 인프라 (BA-CAPA와 동일한 방식 재사용)
# ============================================================
TEMP_PWD_WORKSHEET = "temp_pwd_store"


def _gsheet_available() -> bool:
    return "gcp_service_account" in st.secrets and "temp_pwd_sheet_id" in st.secrets


def _get_gspread_client():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
        return gspread.authorize(creds)
    except Exception:
        return None


def _get_temp_pwd_worksheet():
    if not _gsheet_available():
        return None
    gc = _get_gspread_client()
    if gc is None:
        return None
    import gspread
    sh = gc.open_by_key(st.secrets["temp_pwd_sheet_id"])
    try:
        ws = sh.worksheet(TEMP_PWD_WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=TEMP_PWD_WORKSHEET, rows=200, cols=3)
        ws.update([["password", "expires", "created"]])
    return ws


def load_temp_pwds() -> dict:
    try:
        ws = _get_temp_pwd_worksheet()
        if ws is None:
            return {}
        records = ws.get_all_records()
        st.session_state["_sheets_last_error"] = None
        result = {}
        for row in records:
            pwd = str(row.get("password", "")).strip()
            if not pwd:
                continue
            exp = row.get("expires")
            cre = row.get("created")
            result[pwd] = {
                "expires": datetime.fromisoformat(exp) if exp else None,
                "created": datetime.fromisoformat(cre) if cre else datetime.now(),
            }
        return result
    except Exception as e:
        st.session_state["_sheets_last_error"] = f"[로드 실패] {type(e).__name__}: {e}"
        return {}


def save_temp_pwds(pwd_dict: dict) -> bool:
    """임시 비번 목록을 Google Sheets에 저장.
    중요: 먼저 새 데이터를 기록해 성공을 확인한 뒤에만 남은 이전 행을 지웁니다."""
    try:
        ws = _get_temp_pwd_worksheet()
        if ws is None:
            return False
        rows = [["password", "expires", "created"]]
        for pwd, info in pwd_dict.items():
            exp = info.get("expires")
            cre = info.get("created")
            rows.append([
                pwd,
                exp.isoformat() if isinstance(exp, datetime) else (exp or ""),
                cre.isoformat() if isinstance(cre, datetime) else (cre or str(datetime.now())),
            ])

        ws.update(values=rows, range_name="A1")

        existing_row_count = ws.row_count
        last_needed_row = len(rows)
        if existing_row_count > last_needed_row:
            ws.batch_clear([f"A{last_needed_row + 1}:C{existing_row_count}"])

        st.session_state["_sheets_last_error"] = None
        return True
    except Exception as e:
        st.session_state["_sheets_last_error"] = f"[저장 실패] {type(e).__name__}: {e}"
        return False


def check_temp_pwd(p: str) -> bool:
    fresh = load_temp_pwds()
    st.session_state["temp_pwd_list"] = fresh
    info = fresh.get(p)
    if info is None:
        return False
    if info["expires"] is None:
        return True
    return datetime.now() < info["expires"]


def login_gate():
    if st.session_state.get("authenticated"):
        return True
    if "temp_pwd_list" not in st.session_state:
        st.session_state["temp_pwd_list"] = load_temp_pwds() if _gsheet_available() else {}

    _, center, _ = st.columns([1, 1.8, 1])
    with center:
        st.markdown("<div style='height:60px;'></div>", unsafe_allow_html=True)
        st.markdown(
            f"""<div class='glass-card' style='padding:26px 32px; margin-top:12px;'>
                <div style='text-align:center; margin-bottom:14px;'>
                    <img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:17px;'/>
                </div>
                <div style='color:{ACCENT_BLUE}; font-size:2.16rem; font-weight:700; letter-spacing:0.04em; text-align:center; margin-bottom:10px; line-height:1.25;'>WELDING ARM 측정 데이터 분석</div>
                <div style='text-align:center; margin:0 0 4px 0;'>
                    <span style='display:inline-block; color:{TEXT_MAIN}; font-size:1.35rem; font-weight:600; line-height:1.4; border-bottom:2px solid {ACCENT_BLUE}; padding-bottom:10px;'>전장 범위 내 일자-X · 휨-X Min / Max 추출</span>
                </div>
                <div style='color:{TEXT_MUTED}; font-size:0.82rem; text-align:center; margin-top:12px;'>V2.0</div>
            </div>""",
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        pw_col, btn_col = st.columns([4, 1])
        with pw_col:
            pw = st.text_input("비밀번호", type="password", key="login_pw",
                                label_visibility="collapsed", placeholder="비밀번호 입력")
        with btn_col:
            clicked = st.button("시스템 접속", type="primary", use_container_width=True)
        if clicked:
            if pw == OWNER_PASSWORD:
                st.session_state["authenticated"] = True
                st.session_state["is_owner"] = True
                st.rerun()
            elif check_temp_pwd(pw):
                st.session_state["authenticated"] = True
                st.session_state["is_owner"] = False
                st.session_state["logged_temp_pwd"] = pw
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
    st.stop()
# ============================================================
# 1. 분석 설정값 / 데이터 처리 함수
# ============================================================

DEFAULT_LEN_LOW = 274.8
DEFAULT_LEN_HIGH = 275.2

# 머리글 텍스트(2행)에서 각 항목의 '측정값 열'을 찾는 정규식. 값 열 바로 오른쪽 열이 '편차' 열이다
# (예: '1 일자 X 119' 열 + 그 옆 빈 머리글 열 = 편차). 머리글 표기가 달라지면 오너 전용
# '분석 설정 관리' 탭에서 정규식만 고치면 된다.
DEFAULT_COLUMN_RULES = {
    "일자-X": r"일자\s*X",
    "휨-X": r"휨\s*X",
    "전장": r"^\s*전장",
}
# 시료 행으로 인정할 파일명 패턴 (MIN/MAX/AVG, '0.8 미만' 등 요약·구분 행 제외용)
DEFAULT_NAME_PATTERN = r"\.xls|QMS"

TARGET_KEYS = ("일자-X", "휨-X")  # Min/Max를 구할 항목
_SHORT_ID_RX = _re.compile(r"QMS_\d+")


def _short_id(name) -> str:
    m = _SHORT_ID_RX.search(str(name))
    return m.group(0) if m else str(name)


@st.cache_data(show_spinner=False)
def _list_sheets(file_bytes: bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def _load_raw_sheet(file_bytes: bytes, sheet_name):
    """머리글 위치를 코드에서 직접 찾으므로 header=None 으로 있는 그대로 읽는다."""
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)


def locate_columns(raw: pd.DataFrame, rules: dict):
    """'파일명' 머리글 행과, rules 의 각 항목에 해당하는 값 열/편차 열 위치를 찾는다."""
    header_row = None
    for i in range(min(len(raw), 15)):
        if str(raw.iat[i, 0]).strip() == "파일명":
            header_row = i
            break
    if header_row is None:
        raise ValueError("A열에서 '파일명' 머리글 행을 찾지 못했습니다. 업로드한 엑셀 형식을 확인해주세요.")

    header = raw.iloc[header_row]
    cols, missing = {}, []
    for key, pattern in rules.items():
        rx = _re.compile(pattern)
        found = None
        for c in range(1, raw.shape[1]):
            v = header.iat[c]
            if isinstance(v, str) and rx.search(v.replace("\n", " ")):
                dev = c + 1 if (c + 1 < raw.shape[1] and pd.isna(header.iat[c + 1])) else None
                found = {"value_col": c, "dev_col": dev, "header": v.strip()}
                break
        if found is None:
            missing.append(key)
        else:
            cols[key] = found
    if missing:
        raise ValueError(f"머리글에서 다음 항목을 찾지 못했습니다: {', '.join(missing)} (분석 설정 관리 탭의 정규식을 확인해주세요)")
    return header_row, cols


def prepare_dataset(raw: pd.DataFrame, rules: dict, name_pattern: str):
    """시료 행만 골라 '엑셀 행 / 파일명 / 시료 / 일자-X / 휨-X / 전장 (+ 각 편차)' 표로 만든다."""
    header_row, cols = locate_columns(raw, rules)
    body = raw.iloc[header_row + 1:]

    out = pd.DataFrame({
        "엑셀 행": body.index + 1,  # raw 의 0부터 시작하는 index -> 실제 엑셀 행 번호
        "파일명": body.iloc[:, 0].astype(str).values,
    })
    out["시료"] = out["파일명"].map(_short_id)
    for key, info in cols.items():
        out[key] = pd.to_numeric(body.iloc[:, info["value_col"]], errors="coerce").values
        if info["dev_col"] is not None:
            out[f"{key} 편차"] = pd.to_numeric(body.iloc[:, info["dev_col"]], errors="coerce").values

    key_cols = list(cols.keys())
    is_sample = out["파일명"].str.contains(name_pattern, regex=True, na=False)
    has_number = out[key_cols].notna().any(axis=1)

    info = {
        "header_row_excel": header_row + 1,
        "n_excluded_non_sample": int((~is_sample & has_number).sum()),  # MIN/MAX/AVG 등 숫자가 있는 비시료 행
        "n_incomplete": int((is_sample & out[key_cols].isna().any(axis=1)).sum()),  # 시료인데 값이 빠진 행
    }
    df = out[is_sample].dropna(subset=key_cols).reset_index(drop=True)
    info["n_samples"] = len(df)
    return df, cols, info


def analyze_range(df: pd.DataFrame, low: float, high: float):
    """전장이 [low, high] (경계 포함, 입력 순서 무관) 안에 있는 시료에서
    일자-X / 휨-X 각각의 측정값·편차 Min/Max 와 해당 시료를 찾는다."""
    lo, hi = (low, high) if low <= high else (high, low)
    subset = df[(df["전장"] >= lo) & (df["전장"] <= hi)].sort_values("전장").reset_index(drop=True)

    rows = []
    if not subset.empty:
        for key in TARGET_KEYS:
            for label, col in (("측정값", key), ("편차", f"{key} 편차")):
                if col not in subset.columns:
                    continue
                i_min, i_max = subset[col].idxmin(), subset[col].idxmax()
                rows.append({
                    "항목": key, "구분": label,
                    "Min": subset.at[i_min, col], "Min 시료": subset.at[i_min, "시료"], "Min 엑셀 행": int(subset.at[i_min, "엑셀 행"]),
                    "Max": subset.at[i_max, col], "Max 시료": subset.at[i_max, "시료"], "Max 엑셀 행": int(subset.at[i_max, "엑셀 행"]),
                    "Max-Min": subset.at[i_max, col] - subset.at[i_min, col],
                })
    return pd.DataFrame(rows), subset, lo, hi


def build_result_xlsx(summary: pd.DataFrame, subset: pd.DataFrame, params: dict) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="Min_Max 요약", index=False)
        subset.to_excel(w, sheet_name="범위 내 데이터", index=False)
        pd.DataFrame(list(params.items()), columns=["조건", "값"]).to_excel(w, sheet_name="분석 조건", index=False)
        for ws in w.book.worksheets:
            for idx, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
                header = str(col[0].value or "")
                ws.column_dimensions[get_column_letter(idx)].width = 60 if header == "파일명" else max(12, len(header) + 4)
    return bio.getvalue()


# ============================================================
# 2. UI (원본과 동일한 로그인 → 사이드바 입력 → 상단 탭 구조)
# ============================================================
login_gate()

if "column_rules" not in st.session_state:
    st.session_state["column_rules"] = dict(DEFAULT_COLUMN_RULES)
if "name_pattern" not in st.session_state:
    st.session_state["name_pattern"] = DEFAULT_NAME_PATTERN


def _run_with_progress(key_prefix: str, work_fn):
    """진행 상황을 화면 정중앙의 고정 위치 박스(어두운 배경 포함)로 보여주면서 work_fn을
    실행한다.

    일부러 Streamlit 공식 st.dialog를 쓰지 않는다. st.dialog는 배경(바깥 영역)을
    클릭하면 자동으로 닫혀버리는 동작이 내장돼 있고, 이건 브라우저/Streamlit 내부 JS가
    하는 일이라 CSS(pointer-events 등)로 막으려는 시도가 두 번 다 실패했다
    (2026-09-20). 대신 BA-ROLLER 코드에서 검증된 방식대로, 진짜 모달 API 없이 이 함수가
    직접 전체화면 배경 div + 박스 div를 그린다 - 이 div들에는애초에 '클릭하면 닫기' 핸들러
    자체를 달지 않으므로, 바깥을 클릭해도 그냥 아무 일도 안 일어난다(닫히는 동작 자체가
    존재하지 않음). 박스는 남은 작업이 끝나고 우리가 명시적으로 지울 때만 사라진다.

    work_fn은 update_progress(현재_단계_라벨) 콜백 하나를 인자로 받는 함수여야 하고, 그
    안에서 실제 작업을 하다가 단계가 바뀔 때마다 update_progress(다음 라벨)를 불러
    진행률/문구를 갱신하면 된다 (전체 단계 수는 work_fn이 스스로 알고 있으므로,
    update_progress 호출을 '현재까지 몇 번째'로 세어 자동으로 %를 계산한다).

    끝나면 결과(성공 시 반환값, 실패 시 예외 메시지+트레이스백)를
    session_state[f"{key_prefix}_dialog_result"] / f"{key_prefix}_dialog_error"에 저장하고
    박스를 지운 뒤 rerun한다 - 그 다음 리런에서, 버튼을 그린 코드 바로 아래가 이 값을 읽어
    지금까지처럼 st.success/st.error로 보여주면 된다 (호출부는 예전 st.dialog 버전과 완전히
    동일하게 쓰면 되고 바꿀 필요가 없다)."""
    backdrop_cls = f"_bacapa_backdrop_{key_prefix}"
    box_cls = f"{backdrop_cls}_box"
    spin_cls = f"{backdrop_cls}_spin"
    # 주의: Streamlit의 st.markdown은 HTML을 넣기 전에 마크다운 파서를 먼저 거치는데,
    # 마크다운 문법상 '줄 앞에 공백 4칸 이상 들여쓰기'는 코드블록으로 해석돼서 HTML이
    # 아니라 글자 그대로("<div class=...>") 화면에 찍혀버린다 (2026-09-20 실제 발생 -
    # Python 코드 스타일대로 들여써서 넣은 게 원인). 그래서 여기 들어가는 HTML/CSS 문자열은
    # 반드시 줄마다 앞쪽 공백을 없앤 채로(한 줄씩 이어붙이는 방식으로) 만든다.
    st.markdown(
        f'<style>'
        f'.{backdrop_cls} {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; '
        f'background: rgba(8,8,8,0.82); z-index: 99998; }}'
        f'.{box_cls} {{ position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); '
        f'z-index: 99999; background: #131d34; border-radius: 14px; '
        f'box-shadow: 0 20px 60px rgba(0,0,0,0.7); width: 340px; max-width: 92vw; '
        f'box-sizing: border-box; padding: 22px 24px; }}'
        f'@keyframes {spin_cls} {{ to {{ transform: rotate(360deg); }} }}'
        f'.{spin_cls} {{ display: inline-block; width: 15px; height: 15px; margin-right: 8px; '
        f'border: 3px solid rgba(255,255,255,0.25); border-top-color: {ACCENT_BLUE}; '
        f'border-radius: 50%; animation: {spin_cls} 0.8s linear infinite; '
        f'vertical-align: middle; flex-shrink: 0; }}'
        f'</style>'
        f'<div class="{backdrop_cls}"></div>',
        unsafe_allow_html=True,
    )

    box_ph = st.empty()

    def render(label: str, pct: int):
        box_ph.markdown(
            f'<div class="{box_cls}">'
            f'<div style="font-weight:700;color:{TEXT_MAIN};font-size:1rem;margin-bottom:10px;">처리 중</div>'
            f'<div style="display:flex;align-items:center;color:{TEXT_MAIN};font-size:0.9rem;">'
            f'<span class="{spin_cls}"></span>{label}</div>'
            f'<div style="background:rgba(255,255,255,0.12);border-radius:6px;height:8px;overflow:hidden;margin-top:14px;">'
            f'<div style="background:{ACCENT_BLUE};height:100%;width:{pct}%;transition:width .3s;"></div></div>'
            f'<div style="text-align:right;font-size:0.78rem;color:{TEXT_MUTED};margin-top:4px;">{pct}%</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    calls = {"n": 0}

    def update_progress(label: str):
        calls["n"] += 1
        # 총 단계 수를 미리 알 수 없으니, 부드럽게 100%에 가까워지도록 점근적으로 채운다.
        pct = min(95, int(100 * (1 - 1 / (1 + calls["n"]))))
        render(label, pct)

    update_progress("처리를 시작합니다...")
    try:
        result = work_fn(update_progress)
        render("완료되었습니다", 100)
        st.session_state[f"{key_prefix}_dialog_result"] = result
        st.session_state[f"{key_prefix}_dialog_error"] = None
    except Exception as e:
        import traceback
        st.session_state[f"{key_prefix}_dialog_result"] = None
        st.session_state[f"{key_prefix}_dialog_error"] = (str(e), traceback.format_exc())
    box_ph.empty()
    st.rerun()


def render_header():
    res = st.session_state.get("mm_result")
    status_txt = "분석 완료" if res else "대기 중"
    status_color = SUCCESS if res else TEXT_MUTED
    range_txt = f"{res['lo']:g} ~ {res['hi']:g}" if res else "-"
    st.markdown(
        f"""<div class='glass-card' style='padding:18px 22px;margin-bottom:14px;display:flex;align-items:center;gap:18px;'>
            <img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:13px;flex-shrink:0;'/>
            <div style='width:1px;height:52px;background:{BORDER_STRONG};flex-shrink:0;'></div>
            <div>
                <div style='color:{ACCENT_BLUE};font-size:2.04rem;font-weight:700;letter-spacing:0.02em;margin-bottom:4px;line-height:1.2;'>WELDING ARM 측정 데이터 분석</div>
                <div style='display:flex;align-items:baseline;gap:10px;margin-bottom:2px;'>
                    <span style='color:{TEXT_MAIN};font-size:1.3rem;font-weight:700;border-bottom:2px solid {ACCENT_BLUE};padding-bottom:2px;'>전장 범위 내 일자-X · 휨-X Min / Max 추출</span>
                    <span style='color:{TEXT_MUTED};font-size:0.85rem;'>V1.0</span>
                </div>
                <div style='color:{TEXT_MUTED};font-size:0.85rem;'>
                    상태: <span style='color:{status_color};font-weight:600;'>{status_txt}</span>
                    &nbsp;·&nbsp; 최근 전장 범위: <b style='color:{TEXT_MAIN};'>{range_txt}</b>
                </div>
            </div>
        </div>""",
        unsafe_allow_html=True,
    )


render_header()

# ------------------------------------------------------------------
# 사이드바: 데이터 업로드 + 전장 범위 입력 + 실행 (+ 오너 전용 임시비밀번호 관리)
# ------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"<div style='margin-bottom:6px;'><img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:20px;'/></div>", unsafe_allow_html=True)
    st.caption("용접 암 측정 데이터 → 전장 범위 내 일자-X · 휨-X Min/Max 추출")
    st.markdown("---")

    with st.expander("STEP 1 · 데이터 업로드 및 Min/Max 분석", expanded=True):
        st.markdown('<div class="block-title">STEP 1-1. 측정 데이터 엑셀 업로드 — 파일명 예: DATA-KW</div>', unsafe_allow_html=True)
        data_file = st.file_uploader("측정 데이터 엑셀 업로드", type=["xlsx"], label_visibility="collapsed", key="data_uploader")

        data_bytes, sheet_name = None, None
        if data_file is not None:
            data_bytes = data_file.getvalue()
            try:
                sheets = _list_sheets(data_bytes)
                sheet_name = sheets[0] if len(sheets) == 1 else st.selectbox("시트 선택", sheets, key="data_sheet")
            except Exception as e:
                st.error(f"엑셀을 여는 중 오류가 발생했습니다 (열기 암호가 걸린 파일은 열 수 없습니다): {e}")
                data_bytes = None

        st.markdown('<div class="block-title">STEP 1-2. 전장 범위 (mm)</div>', unsafe_allow_html=True)
        rc1, rc2 = st.columns(2)
        len_low = rc1.number_input("하한", value=DEFAULT_LEN_LOW, step=0.05, format="%.3f", key="len_low")
        len_high = rc2.number_input("상한", value=DEFAULT_LEN_HIGH, step=0.05, format="%.3f", key="len_high")
        st.caption("하한/상한 경계값도 범위에 포함됩니다. 순서를 바꿔 입력해도 자동으로 정렬됩니다.")

        if data_bytes is None:
            st.info("측정 데이터 엑셀을 업로드하면 실행할 수 있습니다.")

        if st.button("Min/Max 분석 실행", type="primary", disabled=data_bytes is None):
            _rules = dict(st.session_state["column_rules"])
            _name_pat = st.session_state["name_pattern"]
            _file_name = data_file.name

            def _analyze_work(update_progress):
                update_progress("엑셀 파일을 읽는 중...")
                raw = _load_raw_sheet(data_bytes, sheet_name)
                update_progress("일자-X / 휨-X / 전장 열을 찾고 시료 행만 추리는 중...")
                df, cols, info = prepare_dataset(raw, _rules, _name_pat)
                update_progress("전장 범위로 필터링하고 Min/Max 계산하는 중...")
                summary, subset, lo, hi = analyze_range(df, len_low, len_high)

                update_progress("결과 파일 만드는 중...")
                params = {
                    "데이터 파일": _file_name, "시트": sheet_name,
                    "전장 하한": lo, "전장 상한": hi, "경계 포함": "예",
                    "유효 시료 수": info["n_samples"], "범위 내 시료 수": len(subset),
                }
                st.session_state["mm_result"] = {
                    "summary": summary, "subset": subset, "lo": lo, "hi": hi,
                    "cols": cols, "info": info, "file": _file_name, "sheet": sheet_name,
                    "xlsx": build_result_xlsx(summary, subset, params) if not subset.empty else None,
                }

                update_progress("처리 이력 기록하는 중...")
                hist = st.session_state.setdefault("process_history", [])
                rec = {"시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "데이터 파일": _file_name,
                       "전장 범위": f"{lo:g} ~ {hi:g}", "범위 내 시료 수": len(subset)}
                if not summary.empty:
                    for key in TARGET_KEYS:
                        r = summary[(summary["항목"] == key) & (summary["구분"] == "측정값")].iloc[0]
                        rec[f"{key} Min ~ Max"] = f"{r['Min']:.4f} ~ {r['Max']:.4f}"
                hist.insert(0, rec)
                st.session_state.setdefault("range_counter", Counter()).update([f"{lo:g} ~ {hi:g}"])
                return {"n_in": len(subset), "lo": lo, "hi": hi}

            _run_with_progress("analyze", _analyze_work)

        if st.session_state.get("analyze_dialog_result") is not None or \
                st.session_state.get("analyze_dialog_error") is not None:
            _err = st.session_state.pop("analyze_dialog_error", None)
            _res = st.session_state.pop("analyze_dialog_result", None)
            if _err:
                _msg, _tb = _err
                st.error(f"분석 중 오류가 발생했습니다: {_msg}")
                with st.expander("상세 오류"):
                    st.code(_tb)
            elif _res["n_in"] == 0:
                st.warning(f"전장 {_res['lo']:g} ~ {_res['hi']:g} 범위 안에 들어오는 시료가 없습니다. 범위를 확인해주세요.")
            else:
                st.success(f"전장 {_res['lo']:g} ~ {_res['hi']:g} 범위 내 {_res['n_in']}건으로 Min/Max를 계산했습니다. 위쪽 'Tab 1' 에서 확인하세요.")

    if st.session_state.get("is_owner"):
        st.markdown("---")
        with st.expander("임시 비밀번호 관리 (오너 전용)", expanded=True):
            if not _gsheet_available():
                st.caption("Google Sheets 연동 미설정 - secrets에 gcp_service_account/temp_pwd_sheet_id 필요")
            else:
                last_err = st.session_state.get("_sheets_last_error")
                if last_err:
                    st.error(last_err)

                pwd_list = st.session_state.get("temp_pwd_list", {})
                st.caption(f"현재 발급된 임시 비밀번호 {len(pwd_list)}개")
                for p, info in list(pwd_list.items()):
                    exp = info.get("expires")
                    exp_txt = exp.strftime("%Y-%m-%d") if exp else "무기한"
                    c1, c2 = st.columns([3, 1])
                    c1.write(f"`{p}` (만료: {exp_txt})")
                    if c2.button("삭제", key=f"del_{p}"):
                        pwd_list.pop(p, None)
                        success = save_temp_pwds(pwd_list)
                        if success:
                            st.session_state["temp_pwd_list"] = load_temp_pwds()
                            st.rerun()
                        else:
                            st.error("삭제 내용 저장에 실패했습니다. API 연결 상태를 확인해주세요.")

                st.markdown("---")
                new_tp = st.text_input("새 임시 비밀번호", key="new_temp_pwd")
                exp_option = st.selectbox("만료 기간", ["무기한", "1일 (24시간)", "7일", "30일"])
                if st.button("발급"):
                    if not new_tp.strip():
                        st.warning("비밀번호를 입력해 주십시오.")
                    elif new_tp == OWNER_PASSWORD:
                        st.error("오너 비밀번호와 동일한 값은 사용할 수 없습니다.")
                    else:
                        now_time = datetime.now()
                        if exp_option == "1일 (24시간)":
                            exp_date = now_time + timedelta(days=1)
                        elif exp_option == "7일":
                            exp_date = now_time + timedelta(days=7)
                        elif exp_option == "30일":
                            exp_date = now_time + timedelta(days=30)
                        else:
                            exp_date = None

                        pwd_list[new_tp.strip()] = {"expires": exp_date, "created": now_time}
                        success = save_temp_pwds(pwd_list)
                        if success:
                            st.session_state["temp_pwd_list"] = load_temp_pwds()
                            st.success("발급되었습니다.")
                            st.rerun()
                        else:
                            st.error("저장에 실패했습니다. 구글 시트 권한이나 API 설정을 확인해주세요.")

    st.markdown("---")
    if st.button("로그아웃"):
        st.session_state["authenticated"] = False
        st.rerun()


# ------------------------------------------------------------------
# 상단 탭
# ------------------------------------------------------------------
if st.session_state.get("is_owner"):
    NAV_ITEMS = ["Tab 1. 전장 범위 Min/Max 분석", "처리 이력 · 통계", "분석 설정 관리", "참고 / 한계 사항"]
    tab_main, tab_history, tab_sync, tab_help = st.tabs(NAV_ITEMS)
else:
    (tab_main,) = st.tabs(["Tab 1. 전장 범위 Min/Max 분석"])
    tab_history = tab_sync = tab_help = None


def _kpi(label, value, sub="", color=None):
    style = f' style="color:{color}"' if color else ""
    sub_html = f'<p class="kpi-label" style="margin-top:4px;">{sub}</p>' if sub else ""
    return f'<div class="kpi-card"><p class="kpi-label">{label}</p><p class="kpi-value"{style}>{value}</p>{sub_html}</div>'


with tab_main:
    res = st.session_state.get("mm_result")
    if res is None:
        st.info("왼쪽 사이드바 STEP 1에서 측정 데이터 엑셀을 업로드하고 전장 범위를 지정한 뒤 'Min/Max 분석 실행'을 눌러주세요.")
    else:
        summary, subset, info = res["summary"], res["subset"], res["info"]
        n_all, n_in = info["n_samples"], len(subset)
        pct = (n_in / n_all * 100) if n_all else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(_kpi("전장 범위 (mm)", f"{res['lo']:g} ~ {res['hi']:g}"), unsafe_allow_html=True)
        k2.markdown(_kpi("유효 시료 수", f"{n_all}개"), unsafe_allow_html=True)
        k3.markdown(_kpi("범위 내 시료 수", f"{n_in}개", color=SUCCESS if n_in else WARNING), unsafe_allow_html=True)
        k4.markdown(_kpi("범위 내 비율", f"{pct:.1f}%"), unsafe_allow_html=True)

        notes = []
        if info["n_excluded_non_sample"]:
            notes.append(f"MIN/MAX/AVG 등 요약·구분 행 {info['n_excluded_non_sample']}개는 시료가 아니므로 제외")
        if info["n_incomplete"]:
            notes.append(f"값이 비어 있는 시료 행 {info['n_incomplete']}개 제외")
        if notes:
            st.caption(" · ".join(notes))

        if summary.empty:
            st.warning("범위 안에 시료가 없어 Min/Max를 계산하지 못했습니다.")
        else:
            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            st.markdown('<div class="block-title">① 일자-X / 휨-X Min · Max (측정값 기준)</div>', unsafe_allow_html=True)
            vals = summary[summary["구분"] == "측정값"].set_index("항목")
            c1, c2, c3, c4 = st.columns(4)
            for col_holder, key, kind in ((c1, "일자-X", "Min"), (c2, "일자-X", "Max"), (c3, "휨-X", "Min"), (c4, "휨-X", "Max")):
                r = vals.loc[key]
                col_holder.markdown(
                    _kpi(f"{key} {kind}", f"{r[kind]:.4f}", f"{r[f'{kind} 시료']} (엑셀 {r[f'{kind} 엑셀 행']}행)"),
                    unsafe_allow_html=True,
                )

            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            st.markdown('<div class="block-title">② 상세 표 (측정값 · 편차)</div>', unsafe_allow_html=True)
            num_cfg = {c: st.column_config.NumberColumn(format="%.4f") for c in ("Min", "Max", "Max-Min")}
            st.dataframe(summary, use_container_width=True, hide_index=True, column_config=num_cfg)

            with st.expander("범위 내 시료 전체 보기 (전장 오름차순)"):
                st.dataframe(subset, use_container_width=True, hide_index=True)

            if res["cols"]:
                with st.expander("인식된 열 정보"):
                    st.dataframe(
                        pd.DataFrame([
                            {"항목": k, "인식한 머리글": v["header"], "값 열": get_column_letter(v["value_col"] + 1),
                             "편차 열": get_column_letter(v["dev_col"] + 1) if v["dev_col"] is not None else "-"}
                            for k, v in res["cols"].items()
                        ]),
                        use_container_width=True, hide_index=True,
                    )

            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            if res.get("xlsx"):
                st.download_button(
                    "전장범위_MinMax_결과.xlsx 다운로드",
                    data=res["xlsx"],
                    file_name="전장범위_MinMax_결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

if tab_history is not None:
  with tab_history:
    st.markdown('<div class="block-title">① 처리 이력</div>', unsafe_allow_html=True)
    st.caption("이번 접속 세션 동안 'Min/Max 분석 실행'을 누른 기록입니다 (브라우저를 새로고침하거나 앱이 재시작되면 초기화됩니다).")
    history = st.session_state.get("process_history", [])
    if not history:
        st.info("아직 처리 이력이 없습니다. 왼쪽 사이드바에서 파일을 업로드하고 실행해보세요.")
    else:
        h1, h2, h3 = st.columns(3)
        h1.markdown(_kpi("이번 세션 분석 횟수", f"{len(history)}회"), unsafe_allow_html=True)
        h2.markdown(_kpi("최근 전장 범위", history[0]["전장 범위"]), unsafe_allow_html=True)
        h3.markdown(_kpi("최근 범위 내 시료 수", f"{history[0]['범위 내 시료 수']}개"), unsafe_allow_html=True)
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(history), use_container_width=True, hide_index=True)
        if st.button("이력 초기화"):
            st.session_state["process_history"] = []
            st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="block-title">② 자주 쓰는 전장 범위 (누적 통계)</div>', unsafe_allow_html=True)
    counter = st.session_state.get("range_counter", Counter())
    if not counter:
        st.info("아직 집계된 통계가 없습니다.")
    else:
        stat_df = pd.DataFrame(counter.most_common(20), columns=["전장 범위", "분석 횟수"])
        c_left, c_right = st.columns([1, 1])
        with c_left:
            st.dataframe(stat_df, use_container_width=True, hide_index=True)
        with c_right:
            st.bar_chart(stat_df.set_index("전장 범위"))
        if st.button("통계 초기화"):
            st.session_state["range_counter"] = Counter()
            st.rerun()

if tab_sync is not None:
  with tab_sync:
    st.markdown('<div class="block-title">분석 설정 관리 (오너 전용)</div>', unsafe_allow_html=True)
    st.caption(
        "엑셀 2행 머리글 텍스트에서 각 항목의 열을 찾는 정규식과, 시료 행으로 인정할 파일명 패턴입니다. "
        "엑셀의 머리글 표기가 바뀌어 열을 못 찾을 때 여기서 정규식만 고치면 됩니다 (코드 수정 불필요)."
    )
    new_rules = {}
    for key, pat in st.session_state["column_rules"].items():
        new_rules[key] = st.text_input(f"'{key}' 머리글 정규식", value=pat, key=f"rule_{key}")
    new_name_pat = st.text_input("시료 행 파일명 패턴 (정규식)", value=st.session_state["name_pattern"], key="rule_name_pattern")
    if st.button("분석 설정 저장"):
        try:
            for p in list(new_rules.values()) + [new_name_pat]:
                _re.compile(p)
            st.session_state["column_rules"] = {k: v.strip() for k, v in new_rules.items()}
            st.session_state["name_pattern"] = new_name_pat.strip()
            st.success("저장되었습니다. 다음 분석부터 적용됩니다.")
        except _re.error as e:
            st.error(f"정규식 오류: {e}")

    if data_bytes is not None:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">현재 업로드 파일에서 열 인식 테스트</div>', unsafe_allow_html=True)
        try:
            _raw = _load_raw_sheet(data_bytes, sheet_name)
            _hr, _cols = locate_columns(_raw, st.session_state["column_rules"])
            st.dataframe(
                pd.DataFrame([
                    {"항목": k, "인식한 머리글": v["header"], "값 열": get_column_letter(v["value_col"] + 1),
                     "편차 열": get_column_letter(v["dev_col"] + 1) if v["dev_col"] is not None else "-"}
                    for k, v in _cols.items()
                ]),
                use_container_width=True, hide_index=True,
            )
        except Exception as e:
            st.error(str(e))

if tab_help is not None:
  with tab_help:
    st.markdown("""
- **계산 기준**: '전장' 측정값(예: 274.93)이 하한 ≤ 전장 ≤ 상한(기본 274.8 ~ 275.2, 경계 포함)인 시료만 대상으로,
  '일자 X'와 '휨 X'의 **측정값**과 **편차**(측정값 − 기준값) 각각의 Min / Max 를 구합니다.
  기준값은 엑셀 머리글 기준 일자 X = 119, 휨 X = 156.16 입니다 (편차는 엑셀 자체 계산값을 그대로 사용).
- **제외되는 행**: 엑셀 하단의 `MIN / MAX / AVG`, `0.8 미만`, `0.8 이상`, `합격 / 불합격` 같은 요약·구분 행은 시료가 아니므로
  파일명 패턴(기본 `.xls` 또는 `QMS` 포함)으로 걸러냅니다. 이 행들의 일자 X 값은 부호가 반대(+119 부근)인 절댓값 통계라
  섞이면 Max 가 크게 틀어지므로 반드시 제외해야 합니다.
- **열 찾기**: 2행 머리글 텍스트(`일자 X`, `휨 X`, `전장`)로 열을 찾고, 값 열 바로 오른쪽 열을 편차 열로 봅니다.
  열 위치가 바뀌어도 동작하지만 머리글 표기가 바뀌면 '분석 설정 관리' 탭에서 정규식을 고쳐야 합니다.
- **한계**: 한 번에 시트 하나만 분석합니다. 시료 하나에 값이 하나라도 비어 있으면(일자 X / 휨 X / 전장) 그 행은 제외합니다.
- '처리 이력 · 통계' 탭은 이번 접속 세션 동안의 기록만 보여줍니다.
""")