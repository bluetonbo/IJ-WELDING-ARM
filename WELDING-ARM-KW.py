# -*- coding: utf-8 -*-
"""
WELDING-ARM-KW : 용접 암 측정 데이터(DATA-KW.xlsx) -> 타겟 스펙을 만족하는 시료의 전장 / 일자-Y Min·Max 추출기
------------------------------------------------------------
- 타겟 항목(단축 위치도 0~1.0, 장축 위치도 0~1.0, 단축 윤곽도 0~0.8)을 '모두 동시에' 만족하는
  시료만 골라서, 변수 항목(전장, 일자-Y)의 Min / Max (측정값과 편차 둘 다)와 해당 시료
  (파일명, 엑셀 행)를 찾는다. 스펙 경계값은 포함한다.
- 타겟 항목 열은 1행(구분 행)의 '단축 위치도 / 장축 위치도 / 단축 윤곽도' 라벨 위치로,
  변수 항목 열은 2행 머리글('전장', '3-1 일자 Y')로 찾는다
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
st.set_page_config(page_title="WELDING-ARM-KW | 스펙 만족 시료 Min/Max 분석", layout="wide", initial_sidebar_state="expanded")

OWNER_PASSWORD = "nt1234"

# ILJIN Corp CI 로고 (PNG, base64 인라인 - 별도 파일 배포 없이 앱 하나로 동작)
ILJIN_LOGO_B64 = "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"

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
    .kpi-card {{ background-color: {SURFACE}; border-radius: 6px; padding: 14px 16px; border: 1px solid {BORDER}; min-height: 118px; box-sizing: border-box; }}
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
# 0-1. 로그인 (관리자: OWNER_PASSWORD / 사용자: secrets.toml 의 user_password)
#   - 사용자 비밀번호는 secrets.toml 에 `user_password = "..."` 로 넣는다 (여러 개면 리스트).
#   - 관리자가 그 줄을 지우거나 바꾸면 이후 사용자 로그인은 막힌다 (관리자 로그인은 항상 가능).
#     이미 로그인해서 쓰고 있는 사용자는 브라우저를 새로고침/재접속할 때부터 막힌다.
#   - secrets.toml 자체가 없거나 user_password 가 비어 있으면 사용자 로그인은 전부 거부된다.
# ============================================================
import hmac


def _user_passwords() -> list:
    try:
        cfg = st.secrets.get("user_password", "")
    except Exception:
        return []
    if isinstance(cfg, str):
        cfg = [cfg]
    try:
        return [str(x).strip() for x in cfg if str(x).strip()]
    except TypeError:
        return []


def check_user_pwd(p: str) -> bool:
    p = (p or "").strip()
    if not p:
        return False
    return any(hmac.compare_digest(p.encode("utf-8"), c.encode("utf-8")) for c in _user_passwords())


def login_gate():
    if st.session_state.get("authenticated"):
        return True

    _, center, _ = st.columns([1, 1.8, 1])
    with center:
        st.markdown("<div style='height:60px;'></div>", unsafe_allow_html=True)
        st.markdown(
            f"""<div class='glass-card' style='padding:26px 32px; margin-top:12px;'>
                <div style='text-align:center; margin-bottom:14px;'>
                    <img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:26px;'/>
                </div>
                <div style='color:{ACCENT_BLUE}; font-size:2.16rem; font-weight:700; letter-spacing:0.04em; text-align:center; margin-bottom:10px; line-height:1.25;'>WELDING ARM 측정 데이터 분석</div>
                <div style='text-align:center; margin:0 0 4px 0;'>
                    <span style='display:inline-block; color:{TEXT_MAIN}; font-size:1.35rem; font-weight:600; line-height:1.4; border-bottom:2px solid {ACCENT_BLUE}; padding-bottom:10px;'>타겟 스펙 만족 시료의 전장 · 일자-Y Min / Max 추출</span>
                </div>
                <div style='color:{TEXT_MUTED}; font-size:0.82rem; text-align:center; margin-top:12px;'>V1.0</div>
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
            elif check_user_pwd(pw):
                st.session_state["authenticated"] = True
                st.session_state["is_owner"] = False
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")
    st.stop()
# ============================================================
# 1. 분석 설정값 / 데이터 처리 함수
# ============================================================

# 타겟 항목: 1행(구분 행)에 적힌 라벨 위치로 열을 찾는다 (2행 머리글은 '4 일자 119'처럼 구분이 안 됨).
DEFAULT_TARGET_RULES = {
    "단축 위치도": r"단축\s*위치도",
    "장축 위치도": r"장축\s*위치도",
    "단축 윤곽도": r"단축\s*윤곽도",
}
# 타겟 스펙 (하한, 상한) - 경계 포함. 사이드바에서 바꿀 수 있다.
DEFAULT_TARGET_SPECS = {
    "단축 위치도": (0.0, 1.0),
    "장축 위치도": (0.0, 1.0),
    "단축 윤곽도": (0.0, 0.8),
}
# 변수 항목: 2행 머리글 텍스트로 '측정값 열'을 찾고, 바로 오른쪽 열(머리글 비어 있음)을 '편차' 열로 본다.
# '일자-Y'는 '3-1 일자 Y 219.66' 열이다 ('3-2 휨 Y'와는 다른 항목).
DEFAULT_VARIABLE_RULES = {
    "전장": r"^\s*전장",
    "일자-Y": r"일자\s*Y",
}
# 시료 행으로 인정할 파일명 패턴 (MIN/MAX/AVG, '0.8 미만' 등 요약·구분 행 제외용)
DEFAULT_NAME_PATTERN = r"\.xls|QMS"

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


def locate_columns(raw: pd.DataFrame, var_rules: dict, target_rules: dict):
    """'파일명' 머리글 행, 변수 항목(값 열/편차 열), 타겟 항목(값 열)의 위치를 찾는다."""
    header_row = None
    for i in range(min(len(raw), 15)):
        if str(raw.iat[i, 0]).strip() == "파일명":
            header_row = i
            break
    if header_row is None:
        raise ValueError("A열에서 '파일명' 머리글 행을 찾지 못했습니다. 업로드한 엑셀 형식을 확인해주세요.")

    header = raw.iloc[header_row]
    ncols = raw.shape[1]
    cols, missing = {}, []

    for key, pattern in var_rules.items():
        rx = _re.compile(pattern)
        found = None
        for c in range(1, ncols):
            v = header.iat[c]
            if isinstance(v, str) and rx.search(v.replace("\n", " ")):
                dev = c + 1 if (c + 1 < ncols and pd.isna(header.iat[c + 1])) else None
                found = {"value_col": c, "dev_col": dev, "header": v.strip(), "kind": "변수"}
                break
        if found is None:
            missing.append(key)
        else:
            cols[key] = found

    for key, pattern in target_rules.items():
        rx = _re.compile(pattern)
        found = None
        # 머리글 바로 위 행(구분 행)부터 위로, 마지막으로 머리글 행 자체도 확인한다.
        for r in list(range(header_row - 1, -1, -1)) + [header_row]:
            for c in range(1, ncols):
                v = raw.iat[r, c]
                if isinstance(v, str) and rx.search(v.replace("\n", " ")):
                    found = {"value_col": c, "dev_col": None, "header": v.strip(), "kind": "타겟"}
                    break
            if found:
                break
        if found is None:
            missing.append(key)
        else:
            cols[key] = found

    if missing:
        raise ValueError(f"머리글에서 다음 항목을 찾지 못했습니다: {', '.join(missing)} (분석 설정 관리 탭의 정규식을 확인해주세요)")
    return header_row, cols


def prepare_dataset(raw: pd.DataFrame, var_rules: dict, target_rules: dict, name_pattern: str):
    """시료 행만 골라 '엑셀 행 / 파일명 / 시료 / 전장 / 일자-Y (+편차) / 타겟 3종' 표로 만든다."""
    header_row, cols = locate_columns(raw, var_rules, target_rules)
    body = raw.iloc[header_row + 1:]

    out = pd.DataFrame({
        "엑셀 행": body.index + 1,  # raw 의 0부터 시작하는 index -> 실제 엑셀 행 번호
        "파일명": body.iloc[:, 0].astype(str).values,
    })
    out["시료"] = out["파일명"].map(_short_id)
    for key in list(var_rules) + list(target_rules):
        info = cols[key]
        out[key] = pd.to_numeric(body.iloc[:, info["value_col"]], errors="coerce").values
        if info["dev_col"] is not None:
            out[f"{key} 편차"] = pd.to_numeric(body.iloc[:, info["dev_col"]], errors="coerce").values

    key_cols = list(var_rules) + list(target_rules)
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


def analyze_spec(df: pd.DataFrame, specs: dict, var_keys):
    """타겟 항목 스펙(하한, 상한 - 경계 포함)을 '전부 동시에' 만족하는 시료만 골라서,
    변수 항목(var_keys)의 측정값/편차 Min/Max 와 해당 시료를 찾는다."""
    mask = pd.Series(True, index=df.index)
    status = []
    for key, (lo, hi) in specs.items():
        lo, hi = (lo, hi) if lo <= hi else (hi, lo)
        ok = (df[key] >= lo) & (df[key] <= hi)
        status.append({
            "타겟 항목": key, "스펙": f"{lo:g} ~ {hi:g}",
            "단독 합격 수": int(ok.sum()), "단독 합격률(%)": (ok.mean() * 100) if len(df) else 0.0,
            "전체 실측 Min": df[key].min() if len(df) else None, "전체 실측 Max": df[key].max() if len(df) else None,
        })
        mask &= ok
    subset = df[mask].sort_values(list(var_keys)[0]).reset_index(drop=True)
    status_df = pd.DataFrame(status)

    rows = []
    if not subset.empty:
        for key in var_keys:
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
    return pd.DataFrame(rows), subset, status_df


def build_result_xlsx(summary: pd.DataFrame, status_df: pd.DataFrame, subset: pd.DataFrame, params: dict) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="Min_Max 요약", index=False)
        status_df.to_excel(w, sheet_name="타겟 합격 현황", index=False)
        subset.to_excel(w, sheet_name="합격 시료 데이터", index=False)
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

if "var_rules" not in st.session_state:
    st.session_state["var_rules"] = dict(DEFAULT_VARIABLE_RULES)
if "target_rules" not in st.session_state:
    st.session_state["target_rules"] = dict(DEFAULT_TARGET_RULES)
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
    pass_txt = f"{len(res['subset'])}개 / {res['info']['n_samples']}개" if res else "-"
    st.markdown(
        f"""<div class='glass-card' style='padding:18px 22px;margin-bottom:14px;display:flex;align-items:center;gap:18px;'>
            <img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:20px;flex-shrink:0;'/>
            <div style='width:1px;height:52px;background:{BORDER_STRONG};flex-shrink:0;'></div>
            <div>
                <div style='color:{ACCENT_BLUE};font-size:2.04rem;font-weight:700;letter-spacing:0.02em;margin-bottom:4px;line-height:1.2;'>WELDING ARM 측정 데이터 분석</div>
                <div style='display:flex;align-items:baseline;gap:10px;margin-bottom:2px;'>
                    <span style='color:{TEXT_MAIN};font-size:1.3rem;font-weight:700;border-bottom:2px solid {ACCENT_BLUE};padding-bottom:2px;'>타겟 스펙 만족 시료의 전장 · 일자-Y Min / Max 추출</span>
                    <span style='color:{TEXT_MUTED};font-size:0.85rem;'>V1.0</span>
                </div>
                <div style='color:{TEXT_MUTED};font-size:0.85rem;'>
                    상태: <span style='color:{status_color};font-weight:600;'>{status_txt}</span>
                    &nbsp;·&nbsp; 스펙 합격 시료: <b style='color:{TEXT_MAIN};'>{pass_txt}</b>
                </div>
            </div>
        </div>""",
        unsafe_allow_html=True,
    )


render_header()

# ------------------------------------------------------------------
# 사이드바: 측정 데이터 업로드 (분석 조건 입력/실행/결과는 각 분석 탭에서)
# ------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"<div style='margin-bottom:6px;'><img src='data:image/png;base64,{ILJIN_LOGO_B64}' style='height:31px;'/></div>", unsafe_allow_html=True)
    st.caption("용접 암 측정 데이터 업로드 → 분석은 위쪽 탭에서")
    st.markdown("---")

    with st.expander("STEP 1 · 측정 데이터 업로드", expanded=True):
        st.markdown('<div class="block-title">측정 데이터 엑셀 업로드 — 파일명 예: DATA-KW</div>', unsafe_allow_html=True)
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

        if data_bytes is None:
            st.info("측정 데이터 엑셀을 업로드하면 위쪽 분석 탭에서 실행할 수 있습니다.")
        else:
            st.caption("업로드 완료 — 분석 조건 입력과 실행은 위쪽 Tab 1에서 합니다.")

    st.markdown("---")
    if st.button("로그아웃"):
        st.session_state["authenticated"] = False
        st.rerun()


# ------------------------------------------------------------------
# 상단 탭
# ------------------------------------------------------------------
if st.session_state.get("is_owner"):
    NAV_ITEMS = ["Tab 1. 스펙 만족 시료 Min/Max 분석", "Tab 2. 스펙 합격 시료 전체 보기", "Tab 3. 인식된 열 정보",
                 "처리 이력 · 통계", "분석 설정 관리", "참고 / 한계 사항"]
    tab_main, tab_samples, tab_cols, tab_history, tab_sync, tab_help = st.tabs(NAV_ITEMS)
else:
    tab_main, tab_samples, tab_cols = st.tabs(
        ["Tab 1. 스펙 만족 시료 Min/Max 분석", "Tab 2. 스펙 합격 시료 전체 보기", "Tab 3. 인식된 열 정보"])
    tab_history = tab_sync = tab_help = None


def _kpi(label, value, sub="", color=None):
    style = f' style="color:{color}"' if color else ""
    sub_html = f'<p class="kpi-label" style="margin-top:4px;">{sub}</p>' if sub else ""
    return f'<div class="kpi-card"><p class="kpi-label">{label}</p><p class="kpi-value"{style}>{value}</p>{sub_html}</div>'


def _cols_table(cols):
    return pd.DataFrame([
        {"구분": v["kind"], "항목": k, "인식한 머리글": v["header"], "값 열": get_column_letter(v["value_col"] + 1),
         "편차 열": get_column_letter(v["dev_col"] + 1) if v["dev_col"] is not None else "-"}
        for k, v in cols.items()
    ])


with tab_main:
    st.markdown('<div class="block-title">분석 조건 · 타겟 항목 스펙 (3개 모두 만족)</div>', unsafe_allow_html=True)
    specs = {}
    _tcols = st.columns(len(st.session_state["target_rules"]))
    for _holder, _key in zip(_tcols, st.session_state["target_rules"]):
        with _holder:
            _d_lo, _d_hi = DEFAULT_TARGET_SPECS.get(_key, (0.0, 1.0))
            st.markdown(f"<div style='font-size:0.9rem;color:{TEXT_MAIN};font-weight:600;'>{_key}</div>", unsafe_allow_html=True)
            _c1, _c2 = st.columns(2)
            _lo = _c1.number_input("하한", value=float(_d_lo), step=0.1, format="%.3f", key=f"spec_lo_{_key}")
            _hi = _c2.number_input("상한", value=float(_d_hi), step=0.1, format="%.3f", key=f"spec_hi_{_key}")
            specs[_key] = (_lo, _hi)
    st.caption("하한/상한 경계값도 합격에 포함됩니다. 변수 항목은 전장, 일자-Y 입니다.")

    if data_bytes is None:
        st.info("왼쪽 사이드바 STEP 1에서 측정 데이터 엑셀을 먼저 업로드해주세요.")

    if st.button("Min/Max 분석 실행", type="primary", disabled=data_bytes is None, key="spec_run"):
        _var_rules = dict(st.session_state["var_rules"])
        _target_rules = dict(st.session_state["target_rules"])
        _name_pat = st.session_state["name_pattern"]
        _file_name = data_file.name
        _specs = dict(specs)

        def _analyze_work(update_progress):
            update_progress("엑셀 파일을 읽는 중...")
            raw = _load_raw_sheet(data_bytes, sheet_name)
            update_progress("타겟/변수 열을 찾고 시료 행만 추리는 중...")
            df, cols, info = prepare_dataset(raw, _var_rules, _target_rules, _name_pat)
            update_progress("3개 타겟 스펙을 모두 만족하는 시료를 고르는 중...")
            summary, subset, status_df = analyze_spec(df, _specs, list(_var_rules))

            update_progress("결과 파일 만드는 중...")
            spec_txt = " / ".join(f"{k} {min(v):g}~{max(v):g}" for k, v in _specs.items())
            params = {
                "데이터 파일": _file_name, "시트": sheet_name, "타겟 스펙(경계 포함)": spec_txt,
                "변수 항목": ", ".join(_var_rules), "유효 시료 수": info["n_samples"], "스펙 합격 시료 수": len(subset),
            }
            st.session_state["mm_result"] = {
                "summary": summary, "subset": subset, "status": status_df, "specs": _specs, "spec_txt": spec_txt,
                "var_keys": list(_var_rules), "cols": cols, "info": info, "file": _file_name, "sheet": sheet_name,
                "xlsx": build_result_xlsx(summary, status_df, subset, params) if not subset.empty else None,
            }

            update_progress("처리 이력 기록하는 중...")
            hist = st.session_state.setdefault("process_history", [])
            rec = {"시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "구분": "타겟 스펙 합격",
                   "데이터 파일": _file_name, "조건": spec_txt, "시료 수": len(subset)}
            if not summary.empty:
                for key in _var_rules:
                    r = summary[(summary["항목"] == key) & (summary["구분"] == "측정값")].iloc[0]
                    rec[f"{key} Min ~ Max"] = f"{r['Min']:.4f} ~ {r['Max']:.4f}"
            hist.insert(0, rec)
            st.session_state.setdefault("spec_counter", Counter()).update([spec_txt])
            return {"n_in": len(subset), "spec_txt": spec_txt}

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
            st.warning("3개 타겟 스펙을 모두 만족하는 시료가 없습니다. 스펙 값을 확인해주세요.")
        else:
            st.success(f"3개 타겟 스펙을 모두 만족하는 시료 {_res['n_in']}건으로 Min/Max를 계산했습니다. 아래 결과를 확인하세요.")

    st.markdown("<hr>", unsafe_allow_html=True)
    res = st.session_state.get("mm_result")
    if res is None:
        st.info("아직 분석 결과가 없습니다. 위에서 타겟 스펙을 확인하고 'Min/Max 분석 실행'을 눌러주세요.")
    else:
        summary, subset, status_df, info = res["summary"], res["subset"], res["status"], res["info"]
        n_all, n_in = info["n_samples"], len(subset)
        pct = (n_in / n_all * 100) if n_all else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(_kpi("합격 조건", "3개 타겟 동시 만족", "단축·장축 위치도, 단축 윤곽도"), unsafe_allow_html=True)
        k2.markdown(_kpi("유효 시료 수", f"{n_all}개", "엑셀 시료 행 기준"), unsafe_allow_html=True)
        k3.markdown(_kpi("스펙 합격 시료 수", f"{n_in}개", "3개 타겟 스펙 모두 만족", color=SUCCESS if n_in else WARNING), unsafe_allow_html=True)
        k4.markdown(_kpi("합격률", f"{pct:.1f}%", "합격 시료 ÷ 유효 시료"), unsafe_allow_html=True)

        notes = []
        if info["n_excluded_non_sample"]:
            notes.append(f"MIN/MAX/AVG 등 요약·구분 행 {info['n_excluded_non_sample']}개는 시료가 아니므로 제외")
        if info["n_incomplete"]:
            notes.append(f"값이 비어 있는 시료 행 {info['n_incomplete']}개 제외")
        if notes:
            st.caption(" · ".join(notes))

        if summary.empty:
            st.warning("3개 타겟 스펙을 모두 만족하는 시료가 없어 Min/Max를 계산하지 못했습니다.")
        else:
            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            st.markdown('<div class="block-title">① 전장 / 일자-Y Min · Max (측정값 기준)</div>', unsafe_allow_html=True)
            vals = summary[summary["구분"] == "측정값"].set_index("항목")
            card_cols = st.columns(2 * len(res["var_keys"]))
            idx = 0
            for key in res["var_keys"]:
                for kind in ("Min", "Max"):
                    r = vals.loc[key]
                    card_cols[idx].markdown(
                        _kpi(f"{key} {kind}", f"{r[kind]:.4f}", f"{r[f'{kind} 시료']} (엑셀 {r[f'{kind} 엑셀 행']}행)"),
                        unsafe_allow_html=True,
                    )
                    idx += 1

            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            st.markdown('<div class="block-title">② 상세 표 (측정값 · 편차)</div>', unsafe_allow_html=True)
            num_cfg = {c: st.column_config.NumberColumn(format="%.4f") for c in ("Min", "Max", "Max-Min")}
            st.dataframe(summary, use_container_width=True, hide_index=True, column_config=num_cfg)

        st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">③ 타겟 항목별 합격 현황</div>', unsafe_allow_html=True)
        st.caption("'단독 합격'은 그 항목 하나만 봤을 때의 합격 수이고, 위 Min/Max는 세 항목을 모두 만족한 시료만 대상으로 합니다.")
        stat_cfg = {c: st.column_config.NumberColumn(format="%.4f") for c in ("전체 실측 Min", "전체 실측 Max")}
        stat_cfg["단독 합격률(%)"] = st.column_config.NumberColumn(format="%.1f")
        st.dataframe(status_df, use_container_width=True, hide_index=True, column_config=stat_cfg)

        if res.get("xlsx"):
            st.markdown("<div style='margin-top:16px;'></div>", unsafe_allow_html=True)
            st.download_button(
                "스펙만족_MinMax_결과.xlsx 다운로드",
                data=res["xlsx"],
                file_name="스펙만족_MinMax_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

with tab_samples:
    _r = st.session_state.get("mm_result")
    if _r is None:
        st.info("아직 분석 결과가 없습니다. Tab 1에서 타겟 스펙을 확인하고 'Min/Max 분석 실행'을 눌러주세요.")
    elif _r["subset"].empty:
        st.warning("3개 타겟 스펙을 모두 만족하는 시료가 없습니다.")
    else:
        st.markdown('<div class="block-title">스펙 합격 시료 전체 보기 (전장 오름차순)</div>', unsafe_allow_html=True)
        st.caption(f"3개 타겟 스펙을 모두 만족하는 시료 {len(_r['subset'])}개 (유효 시료 {_r['info']['n_samples']}개 중) · 조건: {_r['spec_txt']}")
        st.dataframe(_r["subset"], use_container_width=True, hide_index=True)

with tab_cols:
    _r = st.session_state.get("mm_result")
    if _r is None:
        st.info("아직 분석 결과가 없습니다. Tab 1에서 'Min/Max 분석 실행'을 누르면 인식된 열 정보가 표시됩니다.")
    else:
        st.markdown('<div class="block-title">인식된 열 정보</div>', unsafe_allow_html=True)
        st.caption("엑셀 머리글에서 찾은 변수(전장, 일자-Y)와 타겟(위치도, 윤곽도) 항목의 열 위치입니다.")
        st.dataframe(_cols_table(_r["cols"]), use_container_width=True, hide_index=True)

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
        h2.markdown(_kpi("최근 분석 시료 수", f"{history[0]['시료 수']}개"), unsafe_allow_html=True)
        h3.markdown(_kpi("최근 분석 파일", history[0]["데이터 파일"]), unsafe_allow_html=True)
        st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(history), use_container_width=True, hide_index=True)
        if st.button("이력 초기화"):
            st.session_state["process_history"] = []
            st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="block-title">② 자주 쓰는 분석 조건 (누적 통계)</div>', unsafe_allow_html=True)
    counter = st.session_state.get("spec_counter", Counter())
    if not counter:
        st.info("아직 집계된 통계가 없습니다.")
    else:
        stat_df = pd.DataFrame(counter.most_common(20), columns=["분석 조건", "분석 횟수"])
        c_left, c_right = st.columns([1, 1])
        with c_left:
            st.dataframe(stat_df, use_container_width=True, hide_index=True)
        with c_right:
            st.bar_chart(stat_df.set_index("분석 조건"))
        if st.button("통계 초기화"):
            st.session_state["spec_counter"] = Counter()
            st.rerun()

if tab_sync is not None:
  with tab_sync:
    st.markdown('<div class="block-title">분석 설정 관리 (오너 전용)</div>', unsafe_allow_html=True)
    st.caption(
        "타겟 항목은 엑셀 1행(구분 행)의 라벨, 변수 항목은 2행 머리글 텍스트에서 열을 찾는 정규식입니다. "
        "시료 행으로 인정할 파일명 패턴도 여기서 바꿉니다. 표기가 바뀌어 열을 못 찾을 때 정규식만 고치면 됩니다 (코드 수정 불필요)."
    )
    new_var = {k: st.text_input(f"변수 '{k}' 머리글 정규식", value=p, key=f"vrule_{k}")
               for k, p in st.session_state["var_rules"].items()}
    new_tgt = {k: st.text_input(f"타겟 '{k}' 라벨 정규식", value=p, key=f"trule_{k}")
               for k, p in st.session_state["target_rules"].items()}
    new_name_pat = st.text_input("시료 행 파일명 패턴 (정규식)", value=st.session_state["name_pattern"], key="rule_name_pattern")
    if st.button("분석 설정 저장"):
        try:
            for p in list(new_var.values()) + list(new_tgt.values()) + [new_name_pat]:
                _re.compile(p)
            st.session_state["var_rules"] = {k: v.strip() for k, v in new_var.items()}
            st.session_state["target_rules"] = {k: v.strip() for k, v in new_tgt.items()}
            st.session_state["name_pattern"] = new_name_pat.strip()
            st.success("저장되었습니다. 다음 분석부터 적용됩니다.")
        except _re.error as e:
            st.error(f"정규식 오류: {e}")

    if data_bytes is not None:
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="block-title">현재 업로드 파일에서 열 인식 테스트</div>', unsafe_allow_html=True)
        try:
            _raw = _load_raw_sheet(data_bytes, sheet_name)
            _hr, _cols = locate_columns(_raw, st.session_state["var_rules"], st.session_state["target_rules"])
            st.dataframe(_cols_table(_cols), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(str(e))

if tab_help is not None:
  with tab_help:
    st.markdown("""
- **계산 기준**: 타겟 항목 3개(단축 위치도 0~1.0, 장축 위치도 0~1.0, 단축 윤곽도 0~0.8, 경계 포함)를 **모두 동시에** 만족하는
  시료만 대상으로, 변수 항목 **전장**과 **일자-Y**의 측정값과 편차(엑셀 자체 계산값) 각각의 Min / Max 를 구합니다.
  타겟 스펙 값은 Tab 1 상단에서 바꿀 수 있습니다.
- **타겟은 '값' 열 기준**: 위치도/윤곽도 열의 값(예: 0.1128)이 스펙 범위 안인지 봅니다. 바로 옆 열의 편차값
  (예: -0.8872)은 쓰지 않습니다.
- **일자-Y**는 2행 머리글 `3-1 일자 Y 219.66` 열입니다 (`3-2 휨 Y`는 다른 항목이라 포함하지 않습니다).
- **제외되는 행**: 엑셀 하단의 `MIN / MAX / AVG`, `0.8 미만`, `0.8 이상`, `합격 / 불합격` 같은 요약·구분 행은 시료가 아니므로
  파일명 패턴(기본 `.xls` 또는 `QMS` 포함)으로 걸러냅니다.
- **열 찾기**: 타겟은 1행(구분 행) 라벨 위치, 변수는 2행 머리글 텍스트로 찾습니다. 열 위치가 바뀌어도 동작하지만
  표기가 바뀌면 '분석 설정 관리' 탭에서 정규식을 고쳐야 합니다.
- **한계**: 한 번에 시트 하나만 분석합니다. 시료 하나에 값이 하나라도 비어 있으면 그 행은 제외합니다.
- '처리 이력 · 통계' 탭은 이번 접속 세션 동안의 기록만 보여줍니다.
""")
