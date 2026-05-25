"""
성만교회 주차장 차량 관리 시스템
Church Vehicle Management System
(Google Sheets 연동 버전)
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
DEPARTMENTS = ["여성1교구", "남성1교구", "여성2교구", "남성2교구", "여성3교구", "남성3,4교구", "여성4교구", "청년부"]

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
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700;900&display=swap');

/* ── 전체 기본 설정 ── */
html, body, [class*="css"] {
    font-family: 'Noto Sans KR', sans-serif;
}

/* ── 배경 : 밝은 회색 ── */
.stApp {
    background-color: #f4f4f5;
}

/* ── 메인 컨테이너 ── */
.main .block-container {
    padding: 0 0 4rem 0;
    max-width: 500px;
    margin: 0 auto;
    background-color: #f4f4f5;
}

/* ── 헤더 ── */
.header-wrap {
    background-color: #C8192C;
    padding: 2rem 1.5rem 1.8rem;
    text-align: center;
    margin-bottom: 1.6rem;
}

.header-wrap .church-name {
    color: #ffffff;
    font-size: 0.78rem;
    font-weight: 500;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    margin: 0 0 0.5rem 0;
    opacity: 0.85;
}

.header-wrap h1 {
    color: #ffffff;
    font-size: 1.45rem;
    font-weight: 700;
    margin: 0 0 0.3rem 0;
    line-height: 1.35;
    letter-spacing: -0.01em;
}

.header-wrap .sub {
    color: rgba(255,255,255,0.7);
    font-size: 0.76rem;
    margin: 0;
    letter-spacing: 0.02em;
}

.header-divider {
    width: 32px;
    height: 2px;
    background: rgba(255,255,255,0.5);
    margin: 0.8rem auto 0.9rem;
    border-radius: 2px;
}

/* ── 탭 ── */
.stTabs [data-baseweb="tab-list"] {
    background: #ffffff;
    border-bottom: 2px solid #e5e7eb;
    padding: 0 1.2rem;
    gap: 0;
    margin-bottom: 1.4rem;
}

.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border-radius: 0 !important;
    color: #9ca3af !important;
    font-weight: 500 !important;
    font-size: 0.88rem !important;
    padding: 0.85rem 1.2rem !important;
    border-bottom: 2px solid transparent !important;
    margin-bottom: -2px !important;
    transition: all 0.2s ease;
}

.stTabs [aria-selected="true"] {
    color: #C8192C !important;
    border-bottom: 2px solid #C8192C !important;
    font-weight: 700 !important;
}

/* ── 탭 패딩 영역 ── */
.stTabs [data-baseweb="tab-panel"] {
    padding: 0 1.2rem;
}

/* ── 섹션 카드 ── */
.section-card {
    background: #ffffff;
    border-radius: 12px;
    padding: 1.4rem 1.3rem;
    margin-bottom: 1rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07), 0 2px 12px rgba(0,0,0,0.04);
    border: 1px solid #f0f0f0;
}

/* ── 폼 레이블 ── */
.stTextInput label,
.stSelectbox label {
    color: #374151 !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.01em !important;
}

/* ── 입력창 ── */
.stTextInput > div > div > input {
    background: #ffffff !important;
    border: 1.5px solid #d1d5db !important;
    border-radius: 8px !important;
    color: #111827 !important;
    font-family: 'Noto Sans KR', sans-serif !important;
    font-size: 0.9rem !important;
    padding: 0.6rem 0.85rem !important;
    transition: border-color 0.18s ease, box-shadow 0.18s ease;
}

.stTextInput > div > div > input::placeholder {
    color: #9ca3af !important;
}

.stTextInput > div > div > input:focus {
    border-color: #C8192C !important;
    box-shadow: 0 0 0 3px rgba(200,25,44,0.1) !important;
    outline: none !important;
}

/* ── 셀렉트박스 ── */
.stSelectbox > div > div {
    background: #ffffff !important;
    border: 1.5px solid #d1d5db !important;
    border-radius: 8px !important;
    color: #111827 !important;
    font-family: 'Noto Sans KR', sans-serif !important;
}

.stSelectbox > div > div:focus-within {
    border-color: #C8192C !important;
    box-shadow: 0 0 0 3px rgba(200,25,44,0.1) !important;
}

/* ── 등록 버튼 ── */
.stButton > button {
    background-color: #C8192C !important;
    background-image: none !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    padding: 0.72rem 1.5rem !important;
    width: 100% !important;
    transition: background-color 0.18s ease, box-shadow 0.18s ease !important;
    box-shadow: 0 2px 8px rgba(200,25,44,0.3) !important;
    font-family: 'Noto Sans KR', sans-serif !important;
    letter-spacing: 0.02em !important;
}

.stButton > button:hover {
    background-color: #a81424 !important;
    box-shadow: 0 4px 16px rgba(200,25,44,0.38) !important;
}

.stButton > button:active {
    background-color: #8f0f1d !important;
    box-shadow: none !important;
}

/* ── 성공/경고/에러 메시지 ── */
.stSuccess {
    background: #fef2f2 !important;
    border: 1px solid #fca5a5 !important;
    border-left: 4px solid #C8192C !important;
    border-radius: 8px !important;
    color: #7f1d1d !important;
}

.stWarning {
    background: #fffbeb !important;
    border: 1px solid #fcd34d !important;
    border-left: 4px solid #f59e0b !important;
    border-radius: 8px !important;
}

.stError {
    background: #fef2f2 !important;
    border: 1px solid #fca5a5 !important;
    border-left: 4px solid #ef4444 !important;
    border-radius: 8px !important;
}

/* ── 검색 결과 카드 ── */
.result-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-top: 3px solid #C8192C;
    border-radius: 10px;
    padding: 1.4rem 1.3rem 1.2rem;
    margin-bottom: 0.85rem;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    animation: fadeInUp 0.3s ease;
}

@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(8px); }
    to   { opacity: 1; transform: translateY(0); }
}

.result-card .r-label {
    color: #6b7280;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 0.25rem;
}

.result-card .r-plate {
    display: inline-block;
    background: #C8192C;
    color: #ffffff;
    font-size: 1.3rem;
    font-weight: 900;
    padding: 0.35rem 1rem;
    border-radius: 6px;
    letter-spacing: 0.05em;
    margin-bottom: 1rem;
}

.result-card .r-name {
    color: #111827;
    font-size: 1.05rem;
    font-weight: 700;
    display: inline;
}

.result-card .r-dept {
    display: inline-block;
    background: #f3f4f6;
    color: #374151;
    font-size: 0.75rem;
    font-weight: 500;
    padding: 0.2rem 0.65rem;
    border-radius: 20px;
    margin-left: 0.5rem;
    border: 1px solid #e5e7eb;
}

/* ── 액션 버튼 ── */
.action-btns {
    display: flex;
    gap: 0.6rem;
    margin-top: 1.1rem;
}

.btn-call {
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.35rem;
    padding: 0.72rem 0.5rem;
    background: #C8192C;
    color: #ffffff !important;
    font-weight: 700;
    font-size: 0.85rem;
    border-radius: 8px;
    text-decoration: none !important;
    box-shadow: 0 2px 8px rgba(200,25,44,0.28);
    transition: background-color 0.18s ease;
    font-family: 'Noto Sans KR', sans-serif;
}

.btn-call:hover {
    background: #a81424;
    color: #ffffff !important;
    text-decoration: none !important;
}

.btn-sms {
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.35rem;
    padding: 0.72rem 0.5rem;
    background: #ffffff;
    color: #C8192C !important;
    font-weight: 700;
    font-size: 0.85rem;
    border-radius: 8px;
    text-decoration: none !important;
    border: 1.5px solid #C8192C;
    transition: background-color 0.18s ease;
    font-family: 'Noto Sans KR', sans-serif;
}

.btn-sms:hover {
    background: #fef2f2;
    color: #C8192C !important;
    text-decoration: none !important;
}

/* ── 카운트 배지 ── */
.count-badge {
    display: inline-block;
    background: #fef2f2;
    border: 1px solid #fca5a5;
    color: #C8192C;
    font-size: 0.72rem;
    padding: 0.15rem 0.6rem;
    border-radius: 20px;
    margin-left: 0.4rem;
    font-weight: 700;
}

/* ── 검색 결과 없음 ── */
.no-result {
    text-align: center;
    color: #9ca3af;
    padding: 2.5rem 1rem;
    font-size: 0.88rem;
    line-height: 1.7;
}

.no-result .icon {
    font-size: 2rem;
    display: block;
    margin-bottom: 0.5rem;
    opacity: 0.4;
}

/* ── 구분선 ── */
.divider {
    border: none;
    border-top: 1px solid #e5e7eb;
    margin: 1rem 0;
}

/* ── 푸터 ── */
.footer {
    text-align: center;
    color: #9ca3af;
    font-size: 0.7rem;
    padding: 1.5rem 1rem 0.5rem;
    border-top: 1px solid #e5e7eb;
    margin-top: 1.5rem;
}

/* ── 안내 텍스트 ── */
.guide-text {
    color: #6b7280;
    font-size: 0.8rem;
    margin-bottom: 1rem;
    line-height: 1.6;
}

/* ── 검색창 레이블 숨기기 ── */
.search-input label {
    display: none;
}

/* ── 모바일 반응형 ── */
@media (max-width: 480px) {
    .header-wrap h1 {
        font-size: 1.2rem;
    }
    .result-card .r-plate {
        font-size: 1.1rem;
    }
    .stTabs [data-baseweb="tab-panel"] {
        padding: 0 0.8rem;
    }
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 구글 시트 연동 함수
# ─────────────────────────────────────────────
def load_data() -> pd.DataFrame:
    try:
        # ttl=0으로 설정하여 매번 최신 데이터를 불러옵니다.
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(worksheet="시트1", ttl=0)
        
        # 데이터프레임 정리 (빈 문자열 처리 등)
        if df.empty and len(df.columns) == 0:
            return pd.DataFrame(columns=COLUMNS)
            
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        
        df = df[COLUMNS]
        df = df.astype(str)
        df = df.fillna("")
        df = df.replace("nan", "")
        # 전화번호 0 복구 (구글 시트가 0을 지웠을 경우 대비)
        def fix_phone(p):
            p = str(p).strip()
            if p.endswith(".0"):
                p = p[:-2]
            p = re.sub(r"\D", "", p)
            if len(p) in (9, 10) and not p.startswith("0"):
                p = "0" + p
            return p
        
        df["전화번호"] = df["전화번호"].apply(fix_phone)
        
        # 빈 데이터 행 제거
        df = df[df["전화번호"].str.strip() != ""]
        return df
    except Exception as e:
        # 연결 실패나 파일이 비어있는 경우 빈 데이터프레임 반환
        return pd.DataFrame(columns=COLUMNS)

def save_data(df: pd.DataFrame):
    conn = st.connection("gsheets", type=GSheetsConnection)
    conn.update(worksheet="시트1", data=df)
    st.cache_data.clear() # 저장 후 캐시 비우기


# ─────────────────────────────────────────────
# 유효성 검사 및 유틸리티 함수
# ─────────────────────────────────────────────
def is_duplicate(df: pd.DataFrame, phone: str, plate: str) -> bool:
    phone_clean = re.sub(r"\D", "", phone)
    plate_clean = re.sub(r"\s", "", plate).upper()
    for _, row in df.iterrows():
        existing_phone = re.sub(r"\D", "", str(row.get("전화번호", "")))
        existing_plate = re.sub(r"\s", "", str(row.get("차량번호", ""))).upper()
        if existing_phone == phone_clean or existing_plate == plate_clean:
            return True
    return False

def validate_phone(phone: str) -> bool:
    cleaned = re.sub(r"\D", "", phone)
    return len(cleaned) in (10, 11)

def validate_plate(plate: str) -> bool:
    cleaned = re.sub(r"\s", "", plate)
    return len(cleaned) >= 5

def search_by_plate(df: pd.DataFrame, query: str) -> pd.DataFrame:
    q = re.sub(r"\s", "", query).upper()
    mask = df["차량번호"].apply(
        lambda x: q in re.sub(r"\s", "", str(x)).upper()
    )
    return df[mask].reset_index(drop=True)

def make_tel_link(phone: str) -> str:
    cleaned = re.sub(r"\D", "", phone)
    return f"tel:{cleaned}"

def make_sms_link(phone: str, body: str) -> str:
    from urllib.parse import quote
    cleaned = re.sub(r"\D", "", phone)
    encoded_body = quote(body, safe="")
    return f"sms:{cleaned}?body={encoded_body}"


# ─────────────────────────────────────────────
# 앱 시작
# ─────────────────────────────────────────────

# ── 헤더
st.markdown("""
<div class="header-wrap">
    <p class="church-name">Seongman Church</p>
    <h1>성만교회<br>차량 관리 시스템</h1>
    <div class="header-divider"></div>
    <p class="sub">Vehicle Management</p>
</div>
""", unsafe_allow_html=True)

# ── 탭
tab1, tab2 = st.tabs(["  차량 등록  ", "  차량 검색 및 알림  "])


# ══════════════════════════════════════════════
# 탭 1 : 차량 등록
# ══════════════════════════════════════════════
with tab1:
    df_all = load_data()
    car_count = len(df_all) if not df_all.empty else 0

    st.markdown(
        f"<p class='guide-text'>차량 정보를 정확히 입력 후 등록해 주세요."
        f" 현재 등록된 차량 <span class='count-badge'>{car_count}대</span></p>",
        unsafe_allow_html=True,
    )

    with st.form("register_form", clear_on_submit=True):
        name  = st.text_input("이름", placeholder="홍길동")
        phone = st.text_input("전화번호", placeholder="01012345678 (숫자만 입력하셔도 됩니다)")
        plate = st.text_input("차량번호 전체", placeholder="123가4567  또는  12가3456")
        dept  = st.selectbox("소속 부서", options=DEPARTMENTS)
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
            if is_duplicate(df_current, phone, plate):
                st.warning("이미 등록된 전화번호 또는 차량번호입니다.")
            else:
                new_row = {
                    "이름": name,
                    "전화번호": re.sub(r"\D", "", phone), # 숫자만 깔끔하게 저장
                    "차량번호": plate,
                    "소속부서": dept,
                    "등록일시": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }
                df_new = pd.concat(
                    [df_current, pd.DataFrame([new_row])],
                    ignore_index=True,
                )
                save_data(df_new)
                st.success(f"'{name}' 님의 차량이 성공적으로 등록되었습니다.")
                st.rerun() # 새로고침하여 카운트 갱신


# ══════════════════════════════════════════════
# 탭 2 : 차량 검색 및 알림
# ══════════════════════════════════════════════
with tab2:
    SMS_BODY = "성만교회 주차장입니다. 차량 이동 부탁드립니다."

    st.markdown(
        "<p class='guide-text'>차량번호 뒤 4자리 숫자 또는 전체 번호를 입력하세요."
        "<br>예) '4567' 입력 시 '123가4567' 차량이 검색됩니다.</p>",
        unsafe_allow_html=True,
    )

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
                <div class="no-result">
                    <strong style="color:#374151; display:block; margin-bottom:0.3rem;">검색 결과가 없습니다</strong>
                    '{query}' 에 해당하는 차량이 등록되어 있지 않습니다.
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f"<p class='guide-text' style='margin-bottom:0.8rem;'>"
                f"검색 결과 <span class='count-badge'>{len(results)}건</span></p>",
                unsafe_allow_html=True,
            )

            for _, row in results.iterrows():
                owner_name  = str(row.get("이름", ""))
                owner_dept  = str(row.get("소속부서", ""))
                owner_plate = str(row.get("차량번호", ""))
                owner_phone = str(row.get("전화번호", ""))

                tel_link = make_tel_link(owner_phone)
                sms_link = make_sms_link(owner_phone, SMS_BODY)

                st.markdown(
                    f"""
                    <div class="result-card">
                        <div class="r-label">차량번호</div>
                        <div style="margin-bottom:0.9rem;">
                            <span class="r-plate">{owner_plate}</span>
                        </div>
                        <div class="r-label">차주 정보</div>
                        <div style="margin-bottom:0.2rem;">
                            <span class="r-name">{owner_name}</span>
                            <span class="r-dept">{owner_dept}</span>
                        </div>
                        <div class="action-btns">
                            <a href="{tel_link}" class="btn-call">전화 걸기</a>
                            <a href="{sms_link}" class="btn-sms">이동 요청 문자</a>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
    else:
        st.markdown(
            """
            <div class="no-result">
                위 검색창에 차량번호를 입력하면<br>차주 정보와 연락 버튼이 나타납니다.
            </div>
            """,
            unsafe_allow_html=True,
        )

# ── 푸터
st.markdown(
    "<div class='footer'>성만교회 청년부 &middot; 차량 관리 시스템 v1.0</div>",
    unsafe_allow_html=True,
)
