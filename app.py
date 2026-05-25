"""
성만교회 주차장 차량 관리 시스템
Church Vehicle Management System
(Google Sheets)
"""

import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# ─────────────────────────────────────────────
# 설정값
# ─────────────────────────────────────────────
COLUMNS = ["이름", "전화번호", "차량번호", "소속부서", "등록일시"]
DEPARTMENTS = ["여성1교구", "여성2교구", "여성3교구", "여성4교구", "여성5교구",
                "남성1교구", "남성2교구", "남성3교구", "남성4교구", "남성5교구",
                "청년부"]

# ─────────────────────────────────────────────
# 페이지 기본 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="성만교회 차량 관리",
    page_icon="🚗",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# 커스텀 CSS 스타일
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ── Google Font ── */
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Pretendard', 'Noto Sans KR', sans-serif !important;
}

/* ── 배경 : 밝은 회색 ── */
.stApp {
    background-color: #f9fafb;
}

/* ── 메인 컨테이너 ── */
.main .block-container {
    padding: 0 0 4rem 0 !important;
    max-width: 500px;
    margin: 0 auto;
}

/* ── Header ── */
.header-wrap {
    background-color: #ffffff;
    border-bottom: 1px solid #f3f4f6;
    box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    padding: 0.75rem 1rem;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.75rem;
    margin-bottom: 0;
}
.header-wrap img {
    height: 2.5rem;
    width: auto;
}
.header-wrap .divider {
    height: 1.5rem;
    width: 1px;
    background-color: #e5e7eb;
}
.header-wrap span {
    font-size: 0.875rem;
    font-weight: 500;
    color: #4b5563;
}

/* ── Tab Navigation (알약 형태) ── */
.stTabs [data-baseweb="tab-list"] {
    background: #f3f4f6 !important;
    border-radius: 0.75rem !important;
    padding: 0.25rem !important;
    margin: 1rem 1rem 1.5rem 1rem !important;
    border: none !important;
    display: flex !important;
    gap: 0 !important;
}
.stTabs [data-baseweb="tab"] {
    color: #6b7280 !important;
    font-weight: 500 !important;
    font-size: 0.875rem !important;
    padding: 0.625rem 0 !important;
    border: none !important;
    border-radius: 0.5rem !important;
    transition: all 0.2s ease;
    flex: 1 !important;
    text-align: center !important;
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
}
.stTabs [aria-selected="true"] {
    background-color: #ffffff !important;
    color: #C8192C !important;
    box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05) !important;
}
.stTabs [data-baseweb="tab-panel"] {
    padding: 0 1rem;
}

/* ── 폼 & 텍스트 ── */
.stTextInput label,
.stSelectbox label {
    color: #374151 !important;
    font-weight: 500 !important;
    font-size: 0.875rem !important;
    margin-bottom: 0.5rem !important;
}

/* ── 입력창 ── */
.stTextInput > div > div > input,
.stSelectbox > div > div {
    background-color: #f9fafb !important;
    border: 1px solid #e5e7eb !important;
    border-radius: 0.75rem !important; /* xl */
    color: #111827 !important;
    font-size: 1rem !important;
    padding: 0.875rem 1rem !important; /* py-3.5 px-4 */
    transition: all 0.2s ease;
}
.stTextInput > div > div > input::placeholder {
    color: #9ca3af !important;
}
.stTextInput > div > div > input:focus,
.stSelectbox > div > div:focus-within {
    background-color: #ffffff !important;
    border-color: #C8192C !important;
    box-shadow: 0 0 0 2px rgba(200, 25, 44, 0.2) !important;
    outline: none !important;
}

/* ── 폼 컨테이너 (차량 등록 탭) ── */
[data-testid="stForm"] {
    background: #ffffff;
    border-radius: 1rem; /* 2xl */
    padding: 1.25rem; /* p-5 */
    border: 1px solid #f3f4f6;
    box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
}

/* ── 등록 버튼 ── */
.stButton > button, [data-testid="stFormSubmitButton"] > button {
    background-color: #C8192C !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 1rem !important; /* 2xl */
    font-weight: 600 !important;
    font-size: 1rem !important;
    padding: 1rem !important; /* py-4 */
    width: 100% !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
    background-color: #A81526 !important;
}
.stButton > button:active, [data-testid="stFormSubmitButton"] > button:active {
    transform: scale(0.98) !important;
}

/* ── 성공/경고/에러 메시지 ── */
.stSuccess {
    background: #fef2f2 !important;
    border: 1px solid #fca5a5 !important;
    border-left: 4px solid #C8192C !important;
    border-radius: 0.75rem !important;
    color: #7f1d1d !important;
}

.stWarning {
    background: #fffbeb !important;
    border: 1px solid #fcd34d !important;
    border-left: 4px solid #f59e0b !important;
    border-radius: 0.75rem !important;
}

.stError {
    background: #fef2f2 !important;
    border: 1px solid #fca5a5 !important;
    border-left: 4px solid #ef4444 !important;
    border-radius: 0.75rem !important;
}

/* ── 검색 결과 카드 ── */
.result-card {
    background: #ffffff;
    border: 1px solid #f3f4f6;
    border-radius: 1rem; /* 2xl */
    padding: 1.25rem; /* p-5 */
    margin-bottom: 1rem;
    box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
}
.result-card .rc-header {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    margin-bottom: 1rem;
}
.result-card .rc-info-wrap {
    display: flex;
    align-items: center;
    gap: 0.75rem;
}
.result-card .rc-icon-box {
    width: 2.5rem; /* 10 */
    height: 2.5rem; /* 10 */
    background-color: rgba(200, 25, 44, 0.1);
    border-radius: 0.75rem; /* xl */
    display: flex;
    align-items: center;
    justify-content: center;
    color: #C8192C;
}
.result-card .rc-plate {
    font-size: 1.125rem; /* lg */
    font-weight: 700;
    color: #111827;
    margin: 0;
    line-height: 1.2;
}
.result-card .rc-name {
    font-size: 0.875rem; /* sm */
    color: #6b7280;
    margin: 0;
    margin-top: 0.25rem;
}
.result-card .rc-dept {
    padding: 0.25rem 0.75rem; /* py-1 px-3 */
    background-color: rgba(200, 25, 44, 0.1);
    color: #C8192C;
    font-size: 0.75rem; /* xs */
    font-weight: 500;
    border-radius: 9999px; /* full */
}
.result-card .rc-dept.shared-badge {
    background-color: #e0f2fe;
    color: #0369a1;
    border: 1px solid #bae6fd;
    margin-left: 0.5rem;
}

/* ── 액션 버튼 ── */
.action-btns {
    display: flex;
    gap: 0.75rem;
    padding-top: 0.5rem;
}
.btn-call, .btn-sms {
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.5rem;
    padding: 0.75rem 0; /* py-3 */
    border-radius: 0.75rem; /* xl */
    font-weight: 500;
    font-size: 1rem;
    text-decoration: none !important;
    transition: all 0.2s ease;
}
.btn-call {
    background-color: #f3f4f6; /* gray-100 */
    color: #374151 !important; /* gray-700 */
}
.btn-call:hover {
    background-color: #e5e7eb; /* gray-200 */
}
.btn-call:active {
    transform: scale(0.98);
}
.btn-sms {
    background-color: #C8192C;
    color: #ffffff !important;
}
.btn-sms:hover {
    background-color: #A81526;
}
.btn-sms:active {
    transform: scale(0.98);
}

/* ── 검색 결과 없음 / Empty State ── */
.empty-state {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 4rem 0; /* py-16 */
    text-align: center;
}
.empty-state .es-icon-wrap {
    width: 5rem; /* 20 */
    height: 5rem; /* 20 */
    border-radius: 9999px; /* full */
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 1rem;
}
.es-icon-wrap.bg-red {
    background-color: rgba(200, 25, 44, 0.1);
    color: rgba(200, 25, 44, 0.4);
}
.es-icon-wrap.bg-gray {
    background-color: #f3f4f6;
    color: #d1d5db;
}
.empty-state .es-title {
    color: #6b7280; /* gray-500 */
    font-weight: 500;
    font-size: 1rem;
    margin: 0;
}
.empty-state .es-desc {
    color: #9ca3af; /* gray-400 */
    font-size: 0.875rem; /* sm */
    margin: 0;
    margin-top: 0.25rem;
}

/* ── 카운트 배지 ── */
.count-text {
    font-size: 0.875rem; /* sm */
    color: #6b7280; /* gray-500 */
    margin-bottom: 1rem;
}
.count-text span {
    font-weight: 600;
    color: #C8192C;
}

/* ── 기타 UI ── */
.form-title {
    font-size: 1.125rem; /* lg */
    font-weight: 600;
    color: #111827; /* gray-900 */
    display: flex;
    align-items: center;
    gap: 0.5rem;
    margin-bottom: 1.25rem;
}
.form-title svg {
    color: #C8192C;
    width: 1.25rem;
    height: 1.25rem;
}

/* ── 검색창 레이블 숨기기 ── */
.search-input label {
    display: none;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 구글 시트 연동 함수
# ─────────────────────────────────────────────
def load_data() -> pd.DataFrame:
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(worksheet="시트1", ttl=0)
        
        if df.empty and len(df.columns) == 0:
            return pd.DataFrame(columns=COLUMNS)
            
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        
        df = df[COLUMNS]
        df = df.astype(str)
        df = df.fillna("")
        df = df.replace("nan", "")
        def fix_phone(p):
            p = str(p).strip()
            if p.endswith(".0"):
                p = p[:-2]
            p = re.sub(r"\D", "", p)
            if p.startswith("10"):
                p = "0" + p
            return p
            
        df["전화번호"] = df["전화번호"].apply(fix_phone)
        return df
    except Exception as e:
        st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
        return pd.DataFrame(columns=COLUMNS)

def save_data(df: pd.DataFrame):
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        conn.update(worksheet="시트1", data=df)
        st.cache_data.clear()
    except Exception as e:
        st.error(f"데이터를 저장하는 중 오류가 발생했습니다: {e}")

# ─────────────────────────────────────────────
# 유틸리티 함수
# ─────────────────────────────────────────────
def format_phone(p: str) -> str:
    p = re.sub(r"\D", "", str(p))
    if len(p) == 11:
        return f"{p[:3]}-{p[3:7]}-{p[7:]}"
    elif len(p) == 10:
        if p.startswith("02"):
            return f"{p[:2]}-{p[2:6]}-{p[6:]}"
        else:
            return f"{p[:3]}-{p[3:6]}-{p[6:]}"
    return p

def validate_phone(p: str) -> bool:
    digits = re.sub(r"\D", "", p)
    return len(digits) in [10, 11]

def validate_plate(p: str) -> bool:
    p = str(p).strip()
    digits = re.sub(r"\D", "", p)
    return len(digits) >= 4

def is_phone_duplicate(df: pd.DataFrame, phone: str) -> bool:
    if df.empty:
        return False
    digits = re.sub(r"\D", "", phone)
    return any(re.sub(r"\D", "", str(x)) == digits for x in df["전화번호"])

def is_plate_duplicate(df: pd.DataFrame, plate: str) -> bool:
    if df.empty:
        return False
    p_clean = re.sub(r"\s", "", str(plate))
    return any(re.sub(r"\s", "", str(x)) == p_clean for x in df["차량번호"])

def search_by_plate(df: pd.DataFrame, query: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=COLUMNS)
    q_clean = re.sub(r"\s", "", str(query))
    if not q_clean:
        return pd.DataFrame(columns=COLUMNS)
    
    def match_plate(x):
        x_clean = re.sub(r"\s", "", str(x))
        if len(q_clean) == 4 and q_clean.isdigit():
            return x_clean.endswith(q_clean)
        return q_clean in x_clean
        
    mask = df["차량번호"].apply(match_plate)
    return df[mask]

def make_tel_link(phone: str) -> str:
    digits = re.sub(r"\D", "", str(phone))
    return f"tel:{digits}"

def make_sms_link(phone: str, body: str) -> str:
    import urllib.parse
    digits = re.sub(r"\D", "", str(phone))
    encoded_body = urllib.parse.quote(body)
    return f"sms:{digits}?body={encoded_body}"

# ─────────────────────────────────────────────
# 세션 상태 초기화
# ─────────────────────────────────────────────
if "show_duplicate_dialog" not in st.session_state:
    st.session_state.show_duplicate_dialog = False
if "confirm_duplicate_plate" not in st.session_state:
    st.session_state.confirm_duplicate_plate = False
if "pending_registration" not in st.session_state:
    st.session_state.pending_registration = None
if "show_success_dialog" not in st.session_state:
    st.session_state.show_success_dialog = False
if "success_message" not in st.session_state:
    st.session_state.success_message = ""

# ─────────────────────────────────────────────
# 팝업 다이얼로그: 중복 차량 처리
# ─────────────────────────────────────────────
@st.dialog("⚠️ 차량번호 중복 알림")
def show_duplicate_plate_dialog():
    pending = st.session_state.pending_registration
    if not pending:
        return
        
    st.warning(f"이미 등록된 차량번호({pending['plate']})입니다.")
    st.markdown(
        "가족 등 **공동 담당자**로 추가하시겠습니까?<br>"
        "**'추가 등록'**을 눌러 계속하세요."
    )
    col1, col2 = st.columns(2)
    with col1:
        if st.button("추가 등록 ✅", type="primary", use_container_width=True):
            st.session_state.confirm_duplicate_plate = True
            st.session_state.show_duplicate_dialog = False
            st.rerun()
    with col2:
        if st.button("취소", use_container_width=True):
            st.session_state.pending_registration = None
            st.session_state.show_duplicate_dialog = False
            st.rerun()

# ─────────────────────────────────────────────
# 팝업 다이얼로그: 등록 완료
# ─────────────────────────────────────────────
@st.dialog("✅ 등록 완료")
def show_success_dialog():
    st.markdown(st.session_state.success_message)
    if st.button("확인", type="primary", use_container_width=True):
        st.session_state.show_success_dialog = False
        st.rerun()


# ─────────────────────────────────────────────
# 앱 시작
# ─────────────────────────────────────────────

# ── 헤더
st.markdown("""
<div class="header-wrap">
    <img src="https://hebbkx1anhila5yf.public.blob.vercel-storage.com/image-dp6YSXmd5R3GoHrI9qJJkSCyV48QaO.png" alt="성만교회 로고">
    <div class="divider"></div>
    <span>차량 관리 시스템</span>
</div>
""", unsafe_allow_html=True)

# ── 탭
tab1, tab2 = st.tabs(["🔍 차량 검색 및 알림", "➕ 차량 등록"])


# ══════════════════════════════════════════════
# 탭 1 : 차량 검색 및 알림 (순서 변경됨)
# ══════════════════════════════════════════════
with tab1:
    SMS_BODY = "성만교회 주차장입니다. 차량 이동 부탁드립니다."

    query = st.text_input(
        "차량번호 검색",
        placeholder="차량번호 뒤 4자리를 입력하세요",
        label_visibility="collapsed",
    )

    if query.strip():
        df_data = load_data()
        results = search_by_plate(df_data, query.strip())

        if results.empty:
            st.markdown(
                f"""
<div class="empty-state">
<div class="es-icon-wrap bg-gray">
<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9C18.7 10.6 16 10 16 10s-1.3-1.4-2.2-2.3c-.5-.4-1.1-.7-1.8-.7H5c-.6 0-1.1.4-1.4.9l-1.4 2.9A3.7 3.7 0 0 0 2 12v4c0 .6.4 1 1 1h2"/><circle cx="7" cy="17" r="2"/><path d="M9 17h6"/><circle cx="17" cy="17" r="2"/></svg>
</div>
<p class="es-title">검색 결과가 없습니다</p>
<p class="es-desc">다른 번호로 검색해 보세요</p>
</div>
                """,
                unsafe_allow_html=True,
            )
        else:
            # 동일 차량번호 그룹화 (공동 담당자 처리)
            plate_groups: dict = {}
            for _, row in results.iterrows():
                p_key = re.sub(r"\s", "", str(row.get("차량번호", ""))).upper()
                plate_groups.setdefault(p_key, []).append(row)

            total = len(results)
            st.markdown(
                f"<div class='count-text'>검색 결과 <span>{total}</span>건</div>",
                unsafe_allow_html=True,
            )

            for _, owners in plate_groups.items():
                owner_count = len(owners)
                for i, row in enumerate(owners):
                    owner_name  = str(row.get("이름", ""))
                    owner_dept  = str(row.get("소속부서", ""))
                    owner_plate = str(row.get("차량번호", ""))
                    owner_phone = str(row.get("전화번호", ""))

                    tel_link = make_tel_link(owner_phone)
                    sms_link = make_sms_link(owner_phone, f"[성만교회 차량 이동 요청] 안녕하세요. 차량번호 {owner_plate} 이동 부탁드립니다.")

                    shared_badge_html = ""
                    if owner_count > 1:
                        shared_badge_html = f'<span class="rc-dept shared-badge">공동담당자 {i+1}/{owner_count}</span>'

                    st.markdown(
                        f"""
<div class="result-card">
<div class="rc-header">
<div class="rc-info-wrap">
<div class="rc-icon-box">
<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9C18.7 10.6 16 10 16 10s-1.3-1.4-2.2-2.3c-.5-.4-1.1-.7-1.8-.7H5c-.6 0-1.1.4-1.4.9l-1.4 2.9A3.7 3.7 0 0 0 2 12v4c0 .6.4 1 1 1h2"/><circle cx="7" cy="17" r="2"/><path d="M9 17h6"/><circle cx="17" cy="17" r="2"/></svg>
</div>
<div>
<p class="rc-plate">{owner_plate}</p>
<p class="rc-name">{owner_name}</p>
</div>
</div>
<div>
<span class="rc-dept">{owner_dept}</span>
{shared_badge_html}
</div>
</div>
<div class="action-btns">
<a href="{tel_link}" class="btn-call">
<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg>
전화 걸기
</a>
<a href="{sms_link}" class="btn-sms">
<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
이동 요청 문자
</a>
</div>
</div>
                        """,
                        unsafe_allow_html=True,
                    )
    else:
        st.markdown(
            """
<div class="empty-state">
<div class="es-icon-wrap bg-red">
<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></svg>
</div>
<p class="es-title">차량을 검색해 주세요</p>
<p class="es-desc">차량번호 뒤 4자리로 검색할 수 있습니다</p>
</div>
            """,
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════
# 탭 2 : 차량 등록 (순서 변경됨)
# ══════════════════════════════════════════════
with tab2:
    # ── 팝업에서 '추가 등록' 확인 후 실제 등록 처리
    if st.session_state.confirm_duplicate_plate and st.session_state.pending_registration:
        pending = st.session_state.pending_registration
        df_current = load_data()
        new_row = {
            "이름": pending["name"],
            "전화번호": pending["phone"],
            "차량번호": pending["plate"],
            "소속부서": pending["dept"],
            "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        df_new = pd.concat([df_current, pd.DataFrame([new_row])], ignore_index=True)
        save_data(df_new)
        name_registered = pending["name"]
        st.session_state.pending_registration = None
        st.session_state.confirm_duplicate_plate = False
        st.session_state.success_message = f"'{name_registered}' 님이 공동 담당자로 추가 등록되었습니다."
        st.session_state.show_success_dialog = True
        st.rerun()

    # ── 차량번호 중복 팝업 표시
    if st.session_state.show_duplicate_dialog:
        show_duplicate_plate_dialog()

    # ── 등록 완료 팝업 표시
    if st.session_state.show_success_dialog:
        show_success_dialog()

    st.markdown(
        """
<div class="form-title">
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9C18.7 10.6 16 10 16 10s-1.3-1.4-2.2-2.3c-.5-.4-1.1-.7-1.8-.7H5c-.6 0-1.1.4-1.4.9l-1.4 2.9A3.7 3.7 0 0 0 2 12v4c0 .6.4 1 1 1h2"/><circle cx="7" cy="17" r="2"/><path d="M9 17h6"/><circle cx="17" cy="17" r="2"/></svg>
차량 정보 등록
</div>
        """, unsafe_allow_html=True
    )

    with st.form("register_form", clear_on_submit=True):
        name  = st.text_input("이름", placeholder="이름을 입력하세요")
        phone = st.text_input("전화번호", placeholder="010-0000-0000")
        plate = st.text_input("차량번호 전체", placeholder="12가 3456")
        dept  = st.selectbox("소속 교구/부서", options=DEPARTMENTS)
        submitted = st.form_submit_button("등록하기")

    if submitted:
        name  = name.strip()
        phone = phone.strip()
        plate = plate.strip()

        errors = []
        if not name:
            errors.append("이름을 입력해 주세요.")
        if not phone:
            errors.append("전화번호를 입력해 주세요.")
        elif not validate_phone(phone):
            errors.append("전화번호 형식이 올바르지 않습니다. (10~11자리 숫자)")
        if not plate:
            errors.append("차량번호를 입력해 주세요.")
        elif not validate_plate(plate):
            errors.append("차량번호가 너무 짧습니다. 전체 차량번호를 입력해 주세요.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            df_current = load_data()

            if is_phone_duplicate(df_current, phone):
                st.warning("이미 등록된 전화번호입니다.")
            elif is_plate_duplicate(df_current, plate):
                # 동일 차량번호 감지 → 팝업으로 추가 등록 여부 확인
                st.session_state.pending_registration = {
                    "name": name,
                    "phone": re.sub(r"\D", "", phone),
                    "plate": plate,
                    "dept": dept,
                }
                st.session_state.show_duplicate_dialog = True
                st.rerun()
            else:
                new_row = {
                    "이름": name,
                    "전화번호": re.sub(r"\D", "", phone),
                    "차량번호": plate,
                    "소속부서": dept,
                    "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }
                df_new = pd.concat(
                    [df_current, pd.DataFrame([new_row])],
                    ignore_index=True,
                )
                save_data(df_new)
                st.session_state.success_message = f"'{name}' 님의 차량이 성공적으로 등록되었습니다."
                st.session_state.show_success_dialog = True
                st.rerun()

