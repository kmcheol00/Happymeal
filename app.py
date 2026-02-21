# =============================================================
# 사내 점심 주문 집계 앱 (Streamlit + Google Sheets)
#
# ── 로컬 실행 ──
#   pip install streamlit pandas streamlit-gsheets-connection
#   streamlit run app.py
#
# ── Streamlit Cloud 배포 (24시간 무료 서비스) ──
#   1. GitHub에 이 저장소를 push
#   2. https://share.streamlit.io 에서 앱 배포
#   3. App Settings > Secrets 에 아래 secrets.toml 내용 붙여넣기
#   4. Google Sheets를 서비스 계정 이메일에 편집자 권한으로 공유
#
# ── secrets.toml 위치 (로컬) ──
#   프로젝트 루트/.streamlit/secrets.toml
#
# ── .streamlit/secrets.toml 예시 ──
# [connections.gsheets]
# spreadsheet = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID"
# worksheet = "Sheet1"
# type = "service_account"
# project_id = "your-project-id"
# private_key_id = "..."
# private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
# client_email = "your-service-account@your-project.iam.gserviceaccount.com"
# client_id = "..."
# auth_uri = "https://accounts.google.com/o/oauth2/auth"
# token_uri = "https://oauth2.googleapis.com/token"
# =============================================================

import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit_gsheets import GSheetsConnection

# ─────────────────────────────────────────────
# 페이지 기본 설정 (모바일 친화적 centered 레이아웃)
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="🍱 점심 주문 취합",
    page_icon="🍱",
    layout="centered",
)

# ─────────────────────────────────────────────
# 전역 상수 정의
# ─────────────────────────────────────────────
COLUMNS = ["이름", "메뉴", "식당", "주문시간"]  # Google Sheets 컬럼 순서

# 식당별 메뉴 목록 (식당명: [메뉴1, 메뉴2, ...])
RESTAURANTS: dict[str, list[str]] = {
    "한솥도시락": ["제육볶음도시락", "참치마요도시락", "순살치킨도시락"],
    "김밥천국":   ["참치김밥", "치즈라면", "돈까스", "비빔밥"],
    "맘스터치":   ["싸이버거", "불싸이버거", "맘스오리지널"],
    "본죽":       ["전복죽", "참치야채죽", "소고기죽", "닭죽"],
}

# ─────────────────────────────────────────────
# Google Sheets 커넥션 초기화
# credentials는 .streamlit/secrets.toml 에서 자동으로 읽어옴
# ─────────────────────────────────────────────
conn = st.connection("gsheets", type=GSheetsConnection)

# ─────────────────────────────────────────────
# 커스텀 CSS 주입
# 모바일 화면에서 버튼이 가득 채워지고,
# 카드에 테두리·그림자가 적용되도록 스타일 설정
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ── 전체 폰트 크기 ── */
html, body, [class*="css"] {
    font-size: 16px;
}

/* ── 버튼 full-width 및 스타일 ── */
div.stButton > button {
    width: 100%;
    padding: 0.6rem 0.4rem;
    font-size: 0.95rem;
    font-weight: 600;
    border-radius: 10px;
    border: 1.5px solid #4A90D9;
    background-color: #EAF4FF;
    color: #1A5FA8;
    transition: background-color 0.2s;
}
div.stButton > button:hover {
    background-color: #C2DCFF;
}

/* ── 메뉴 카드 스타일 ── */
.menu-card {
    background: #FFFFFF;
    border: 1px solid #DEE2E6;
    border-radius: 12px;
    padding: 10px;
    margin-bottom: 14px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    text-align: center;
}
.menu-card img {
    border-radius: 8px;
    width: 100%;
    object-fit: cover;
}
.menu-name {
    font-size: 0.9rem;
    font-weight: 700;
    margin: 8px 0 6px 0;
    color: #333;
}

/* ── 식당 탭 헤더 ── */
.restaurant-header {
    font-size: 1.2rem;
    font-weight: 800;
    color: #1A5FA8;
    margin: 16px 0 10px 0;
    border-left: 4px solid #4A90D9;
    padding-left: 10px;
}

/* ── 섹션 구분선 ── */
.section-divider {
    border: none;
    border-top: 2px dashed #DEE2E6;
    margin: 24px 0;
}

/* ── 초기화 버튼 (빨간색 계열) ── */
div.delete-btn > button {
    background-color: #FFF0F0;
    border-color: #E74C3C;
    color: #C0392B;
}
div.delete-btn > button:hover {
    background-color: #FDDEDE;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# session_state 초기화 (앱 최초 로드 시 한 번만 실행)
# ─────────────────────────────────────────────
if "user_name" not in st.session_state:
    st.session_state["user_name"] = ""          # 입력된 사용자 이름

if "last_order_msg" not in st.session_state:
    st.session_state["last_order_msg"] = None   # 마지막 주문 완료 메시지


# ─────────────────────────────────────────────
# 헬퍼 함수: Google Sheets에서 주문 목록 불러오기
# ttl=5: 5초 캐시로 API 호출 횟수를 줄이면서 근실시간 반영
# ─────────────────────────────────────────────
def load_orders_from_gsheet() -> pd.DataFrame:
    """Google Sheets에서 주문 데이터를 읽어온다.
    시트가 비어 있거나 연결 실패 시 빈 DataFrame 반환."""
    try:
        df = conn.read(
            usecols=COLUMNS,
            ttl=5,  # 5초마다 최신 데이터 반영
        )
        # Google Sheets는 빈 행을 포함하는 경우가 있어 제거
        df = df.dropna(how="all").reset_index(drop=True)
        return df
    except Exception:
        # 시트가 완전히 비어 있거나 컬럼이 없을 경우
        return pd.DataFrame(columns=COLUMNS)


# ─────────────────────────────────────────────
# 헬퍼 함수: 주문 1건을 Google Sheets에 저장
# 기존 데이터를 모두 읽은 뒤 새 행을 추가하여 전체를 덮어씀
# ─────────────────────────────────────────────
def save_order_to_gsheet(name: str, menu: str, restaurant: str) -> bool:
    """새 주문을 Google Sheets에 추가한다. 성공 시 True, 실패 시 False 반환."""
    try:
        existing_df = load_orders_from_gsheet()
        new_row = pd.DataFrame([{
            "이름": name,
            "메뉴": menu,
            "식당": restaurant,
            "주문시간": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        # 기존 데이터 뒤에 새 행을 붙여서 전체 시트를 업데이트
        updated_df = pd.concat([existing_df, new_row], ignore_index=True)
        conn.update(data=updated_df)
        return True
    except Exception:
        st.error("⚠️ Google Sheets 연결에 실패했습니다. 잠시 후 다시 시도해주세요.")
        return False


# ─────────────────────────────────────────────
# 헬퍼 함수: 특정 사용자의 주문만 Google Sheets에서 삭제
# conn.clear() 후 update()하여 하단 찌꺼기 행을 방지
# ─────────────────────────────────────────────
def delete_my_order_from_gsheet(name: str) -> bool:
    """해당 이름의 주문 행을 모두 삭제한다. 성공 시 True, 실패 시 False 반환."""
    try:
        existing_df = load_orders_from_gsheet()
        updated_df = existing_df[existing_df["이름"] != name].reset_index(drop=True)
        conn.clear()
        conn.update(data=updated_df)
        return True
    except Exception:
        st.error("⚠️ Google Sheets 연결에 실패했습니다. 잠시 후 다시 시도해주세요.")
        return False


# ─────────────────────────────────────────────
# 앱 상단: 타이틀 & 안내 문구
# ─────────────────────────────────────────────
st.title("🍱 점심 주문 취합")
st.caption("팀원 모두가 같은 URL로 접속해 메뉴를 선택하세요.")

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 섹션 1: 이름 입력
# session_state에 저장하여 페이지 재렌더링 후에도 유지됨
# ─────────────────────────────────────────────
st.subheader("👤 이름 입력")
name_input = st.text_input(
    label="이름",
    value=st.session_state["user_name"],
    placeholder="예: 김민철, 김지훈",
    label_visibility="collapsed",
)
# 텍스트 입력값을 session_state에 즉시 반영
st.session_state["user_name"] = name_input.strip()

# 마지막 주문 완료/경고 메시지 출력 (버튼 클릭 직후에만 표시)
if st.session_state["last_order_msg"]:
    msg_type, msg_text = st.session_state["last_order_msg"]
    if msg_type == "success":
        st.success(msg_text)
    elif msg_type == "warning":
        st.warning(msg_text)
    # 한 번 표시 후 초기화 (다음 렌더링에서 중복 표시 방지)
    st.session_state["last_order_msg"] = None

# 이름을 입력했을 때만 '내 주문 취소' 버튼 노출
if st.session_state["user_name"]:
    st.markdown('<div class="delete-btn">', unsafe_allow_html=True)
    if st.button("❌ 내 주문 취소하기"):
        cancel_name = st.session_state["user_name"]
        if delete_my_order_from_gsheet(cancel_name):
            st.session_state["last_order_msg"] = (
                "success",
                f"🗑️ {cancel_name}님의 주문이 취소되었습니다.",
            )
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 섹션 2: 식당별 메뉴 그리드
# 식당마다 st.tabs로 구분하고, 메뉴를 2열 그리드로 배치
# ─────────────────────────────────────────────
st.subheader("🍽️ 메뉴 선택")

tab_labels = list(RESTAURANTS.keys())
tabs = st.tabs(tab_labels)

for tab, (restaurant, menus) in zip(tabs, RESTAURANTS.items()):
    with tab:
        # 메뉴 2열 그리드 배치
        cols = st.columns(2)
        for idx, menu in enumerate(menus):
            col = cols[idx % 2]
            with col:
                # 메뉴 카드 (이미지 + 이름 + 주문 버튼)
                st.markdown('<div class="menu-card">', unsafe_allow_html=True)

                # picsum.photos: seed로 메뉴마다 고정된 이미지 사용
                image_url = f"https://picsum.photos/seed/{menu}/300/200"
                st.image(image_url, caption=menu, use_container_width=True)

                st.markdown('</div>', unsafe_allow_html=True)

                # 주문 버튼 — 고유 key로 각 버튼을 구분
                btn_key = f"order_{restaurant}_{menu}"
                if st.button(f"🛒 {menu} 담기", key=btn_key):
                    current_name = st.session_state["user_name"]
                    if not current_name:
                        # 이름 미입력 시 경고 메시지 세팅 후 재실행
                        st.session_state["last_order_msg"] = (
                            "warning", "⚠️ 이름을 먼저 입력해주세요!"
                        )
                        st.rerun()
                    else:
                        # Google Sheets에 주문 저장 — 성공 시에만 성공 메시지 세팅 후 재실행
                        if save_order_to_gsheet(current_name, menu, restaurant):
                            st.session_state["last_order_msg"] = (
                                "success",
                                f"✅ {current_name}님의 [{restaurant}] {menu} 주문이 완료되었습니다!",
                            )
                            st.rerun()
                        # 실패 시 st.error는 save_order_to_gsheet 내부에서 이미 표시됨

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 섹션 3: 주문 현황 (전체 목록 + 집계)
# 새로고침 버튼 클릭 시 st.rerun()으로 최신 데이터를 다시 읽어옴
# ─────────────────────────────────────────────
col_title, col_refresh = st.columns([3, 1])
with col_title:
    st.subheader("📋 주문 현황")
with col_refresh:
    # 수동 새로고침 버튼: 클릭 시 Google Sheets를 다시 읽어 최신 주문 반영
    if st.button("🔄 새로고침"):
        st.rerun()

orders_df = load_orders_from_gsheet()

if orders_df.empty:
    st.info("아직 주문이 없습니다. 위에서 메뉴를 선택해주세요!")
else:
    # ── 3-1. 전체 주문 목록 ──
    st.markdown("#### 전체 주문 목록")
    st.dataframe(
        orders_df,
        use_container_width=True,
        hide_index=True,
    )

    # ── 3-2. 메뉴별 집계 ──
    st.markdown("#### 🍜 메뉴별 총 주문 수")
    summary = (
        orders_df
        .groupby(["식당", "메뉴"], sort=False)
        .size()
        .reset_index(name="주문 수")
        .sort_values("주문 수", ascending=False)
    )
    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True,
    )

