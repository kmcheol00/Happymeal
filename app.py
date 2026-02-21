# =============================================================
# 사내 점심 주문 집계 앱 (Streamlit + Google Sheets)
#
# pip install streamlit pandas streamlit-gsheets-connection Pillow
# Streamlit Cloud 배포 시 menu/ 폴더를 GitHub에 함께 push할 것
# 절대경로 하드코딩 금지: Path(__file__).parent 사용
#
#   pip install streamlit pandas streamlit-gsheets-connection
# ── 로컬 실행 ──
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

import io
import json as _json
import time as _time
import os as _os

# #region agent debug log helper
def _dbg(loc, msg, data=None, hyp=None):
    _e = {"sessionId": "b17557", "timestamp": int(_time.time() * 1000),
          "location": loc, "message": msg, "data": data or {}, "hypothesisId": hyp}
    _line = _json.dumps(_e, ensure_ascii=False) + "\n"
    for _log in [
        r"C:\Users\mincheol\OneDrive\keti\TEM\TES\Tem_defect_project\debug-b17557.log",
        r"C:\Users\mincheol\AppData\Local\Temp\debug-b17557.log",
    ]:
        try:
            with open(_log, "a", encoding="utf-8") as _f:
                _f.write(_line)
        except Exception:
            pass
# #endregion

# #region agent log: app startup context (H-A, H-C, H-D)
# 식당별 메뉴 목록 (식당명: [메뉴1, 메뉴2, ...])
RESTAURANTS: dict[str, list[str]] = {
    "한솥도시락": ["제육볶음도시락", "참치마요도시락", "순살치킨도시락"],
    "김밥천국":   ["참치김밥", "치즈라면", "돈까스", "비빔밥"],
    "맘스터치":   ["싸이버거", "불싸이버거", "맘스오리지널"],
    "본죽":       ["전복죽", "참치야채죽", "소고기죽", "닭죽"],
# 모바일 화면에서 버튼이 가득 채워지고,
# 카드에 테두리·그림자가 적용되도록 스타일 설정
}

_dbg("app.py:top", "앱 시작", {
    "__name__": __name__,
    "pid": _os.getpid(),
    "STREAMLIT_GUARD": _os.environ.get("_ST_GUARD", "없음"),
    "cwd": _os.getcwd(),
}, hyp="A,C,D")
# #endregion

_dbg("app.py:before_st_import", "streamlit import 시도 전", {}, hyp="A")
from streamlit_gsheets import GSheetsConnection
import streamlit as st
_dbg("app.py:after_st_import", "streamlit import 성공", {}, hyp="A")

# #region agent log: individual imports (H-E)
_dbg("app.py:before_pandas", "pandas import 시도", {}, hyp="E")
import pandas as pd
_dbg("app.py:after_pandas", "pandas import 성공", {}, hyp="E")

from datetime import datetime
from pathlib import Path

_dbg("app.py:before_PIL", "PIL import 시도", {}, hyp="E")
from PIL import Image
_dbg("app.py:after_PIL", "PIL import 성공", {}, hyp="E")
# #endregion

# #region agent log: gsheets import (H-B)
_dbg("app.py:before_gsheets_import", "streamlit_gsheets import 시도 전", {}, hyp="B")
try:
    from streamlit_gsheets import GSheetsConnection
    _dbg("app.py:gsheets_import_ok", "streamlit_gsheets import 성공", {}, hyp="B")
except Exception as _e:
    _dbg("app.py:gsheets_import_fail", "streamlit_gsheets import 실패", {"error": str(_e)}, hyp="B")
    raise
# #endregion

# ─────────────────────────────────────────────
# 페이지 기본 설정 (모바일 친화적 centered 레이아웃)
# ─────────────────────────────────────────────
# #region agent log: set_page_config (H-A)
_dbg("app.py:before_set_page_config", "st.set_page_config 호출 직전 - Streamlit 컨텍스트 확인", {}, hyp="A")
try:
    st.set_page_config(
        page_title="🍱 점심 주문 취합",
        page_icon="🍱",
        layout="centered",
    )
    _dbg("app.py:set_page_config_ok", "st.set_page_config 성공", {}, hyp="A")
except Exception as _e:
    _dbg("app.py:set_page_config_fail", "st.set_page_config 예외 발생", {"error": str(_e), "type": type(_e).__name__}, hyp="A")
    raise
# #endregion

# ─────────────────────────────────────────────
# 전역 상수 정의
# ─────────────────────────────────────────────
COLUMNS = ["이름", "메뉴", "식당", "주문시간"]  # Google Sheets 컬럼 순서
MENU_DIR = Path(__file__).parent                # app.py가 menu/ 안에 위치하므로 parent가 곧 메뉴 루트
AUTO_CLEAR_HOURS = 3                            # 주문 자동 만료 시간(시)

# #region agent log: path & secrets check (H-B, H-D)
_secrets_path = MENU_DIR / ".streamlit" / "secrets.toml"
_dbg("app.py:constants", "경로 및 secrets 확인", {
    "MENU_DIR": str(MENU_DIR),
    "secrets_exists": _secrets_path.exists(),
    "secrets_path": str(_secrets_path),
}, hyp="B,D")
# #endregion

# ─────────────────────────────────────────────
# Google Sheets 커넥션 초기화
# credentials는 .streamlit/secrets.toml 에서 자동으로 읽어옴
# ─────────────────────────────────────────────
# #region agent log: gsheets connection (H-C)
_dbg("app.py:before_conn", "st.connection 호출 직전", {}, hyp="C")
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
    _dbg("app.py:conn_ok", "st.connection 성공", {}, hyp="C")
except Exception as _e:
    _dbg("app.py:conn_fail", "st.connection 실패", {"error": str(_e), "type": type(_e).__name__}, hyp="C")
if "confirm_delete" not in st.session_state:
    st.session_state["confirm_delete"] = False  # 초기화 확인 단계 플래그

    raise
# #endregion

# ─────────────────────────────────────────────
# 커스텀 CSS 주입
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
    padding: 0.3rem 0.2rem;
    font-size: 0.6rem;
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
    padding: 6px;
    margin-bottom: 10px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    text-align: center;
}
        # 시트가 완전히 비어 있거나 컬럼이 없을 경우
.menu-card img {
    border-radius: 8px;
    width: 100%;
    object-fit: cover;
}
.menu-name {
    font-size: 0.6rem;
# ─────────────────────────────────────────────
# 헬퍼 함수: Google Sheets 전체 초기화
# 헤더만 남기고 빈 DataFrame으로 덮어씀

# ─────────────────────────────────────────────
def reset_gsheet() -> None:
    """Google Sheets의 모든 주문 데이터를 삭제한다."""
    try:
        empty_df = pd.DataFrame(columns=COLUMNS)
        conn.update(data=empty_df)
    except Exception:
        st.error("⚠️ Google Sheets 연결에 실패했습니다. 잠시 후 다시 시도해주세요.")
    font-weight: 700;
    margin: 6px 0 4px 0;
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

/* ── 4열 그리드 모바일 강제 유지 ── */
[data-testid="column"] {
    min-width: 0 !important;
    flex: 1 1 0% !important;
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

if "_last_auto_clear" not in st.session_state:
    st.session_state["_last_auto_clear"] = None  # auto_clear 마지막 실행 시각


# ─────────────────────────────────────────────
# 헬퍼 함수: Google Sheets에서 주문 목록 불러오기
# ttl=5: 5초 캐시로 API 호출 횟수를 줄이면서 근실시간 반영
# ─────────────────────────────────────────────
def load_orders_from_gsheet() -> pd.DataFrame:
    """Google Sheets에서 주문 데이터를 읽어온다.
    시트가 비어 있거나 연결 실패 시 빈 DataFrame 반환."""
    try:
        df = conn.read(ttl=5)
        df = df.dropna(how="all").reset_index(drop=True)
        existing_cols = [c for c in COLUMNS if c in df.columns]
        if not existing_cols:
            return pd.DataFrame(columns=COLUMNS)
        return df[existing_cols].reindex(columns=COLUMNS)
    except Exception:
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
# 헬퍼 함수: menu/ 폴더를 스캔하여 식당-이미지경로 딕셔너리 반환
# 앱 시작 시 1회만 실행 (cache_data)
# ─────────────────────────────────────────────
@st.cache_data
def scan_restaurants() -> dict[str, list[str]]:
    """menu/ 하위 폴더명=식당명, 이미지 파일명=메뉴명으로 구조를 스캔한다."""
    if not MENU_DIR.is_dir():
        st.error(f"⚠️ 식당 이미지 폴더를 찾을 수 없습니다. ({MENU_DIR})")
        st.stop()
    result: dict[str, list[str]] = {}
    for folder in sorted(MENU_DIR.iterdir()):
st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
        if not folder.is_dir():
            continue
        images = sorted(
            str(p) for p in folder.iterdir()
            if p.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp"}
        )
        if images:
            result[folder.name] = images
    return result


# ─────────────────────────────────────────────
# 헬퍼 함수: 이미지를 300x300 정사각형으로 중앙 크롭하여 JPEG bytes 반환
# 한글 경로 대응: Image.open(str(path))
# ─────────────────────────────────────────────
def load_square_image(path: Path) -> bytes:
    """이미지를 300×300 중앙 크롭 JPEG bytes로 반환한다. 실패 시 회색 플레이스홀더."""
    try:
        img = Image.open(str(path)).convert("RGB")
        w, h = img.size
        side = min(w, h)
        left = (w - side) // 2
        top = (h - side) // 2
        img = img.crop((left, top, left + side, top + side)).resize((300, 300))
        buf = io.BytesIO()
        img.save(buf, format="JPEG")
        return buf.getvalue()
    except Exception:
        buf = io.BytesIO()
        Image.new("RGB", (300, 300), (200, 200, 200)).save(buf, format="JPEG")
        return buf.getvalue()


# ─────────────────────────────────────────────
# 헬퍼 함수: AUTO_CLEAR_HOURS 경과 주문 자동 삭제
# 불필요한 API 호출 방지: 삭제할 행이 있을 때만 update
# ─────────────────────────────────────────────
def auto_clear_old_orders() -> None:
    """주문 시간이 AUTO_CLEAR_HOURS 이상 경과한 행을 자동으로 삭제한다."""
    df = load_orders_from_gsheet()
    if df.empty:
        return
    now = datetime.now()
    df["주문시간"] = pd.to_datetime(df["주문시간"])
    fresh_df = df[now - df["주문시간"] < pd.Timedelta(hours=AUTO_CLEAR_HOURS)]
    fresh_df = fresh_df.reset_index(drop=True)
    if len(fresh_df) < len(df):
        conn.update(data=fresh_df)


# ─────────────────────────────────────────────
# 앱 시작 시 오래된 주문 자동 정리 (5분에 1회 제한)
# ─────────────────────────────────────────────
_now = datetime.now()
if (
    st.session_state["_last_auto_clear"] is None
    or (_now - st.session_state["_last_auto_clear"]).total_seconds() > 300
):
    auto_clear_old_orders()
    st.session_state["_last_auto_clear"] = _now

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


# ─────────────────────────────────────────────
# 섹션 4: 전체 초기화
# 실수 방지를 위해 2단계 확인 (session_state 플래그 활용)
# ─────────────────────────────────────────────
st.subheader("🗑️ 전체 초기화")

if not st.session_state["confirm_delete"]:
    # 1단계: 초기화 버튼 (빨간 계열 스타일)
    st.markdown('<div class="delete-btn">', unsafe_allow_html=True)
    if st.button("🗑️ 전체 주문 초기화"):
        # 확인 단계로 진입
        st.session_state["confirm_delete"] = True
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
else:
    # 2단계: 정말 삭제할지 확인하는 단계
    st.warning("⚠️ 정말로 모든 주문을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.")
    col_yes, col_no = st.columns(2)

    with col_yes:
        st.markdown('<div class="delete-btn">', unsafe_allow_html=True)
        if st.button("✅ 예, 삭제합니다"):
            # Google Sheets를 빈 DataFrame으로 덮어써서 초기화
            reset_gsheet()
            # 확인 플래그 초기화
            st.session_state["confirm_delete"] = False
            st.success("모든 주문이 초기화되었습니다.")
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col_no:
        if st.button("❌ 취소"):
            # 삭제 취소 — 확인 플래그만 리셋
            st.session_state["confirm_delete"] = False
            st.rerun()

            )
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 섹션 2: 식당별 메뉴 그리드
# 식당마다 st.tabs로 구분하고, 메뉴를 4열 그리드로 배치
# ─────────────────────────────────────────────
st.subheader("🍽️ 메뉴 선택")

restaurants = scan_restaurants()
tab_labels = list(restaurants.keys())
tabs = st.tabs(tab_labels)

for tab, (restaurant, img_paths) in zip(tabs, restaurants.items()):
    with tab:
        # 메뉴 4열 그리드 배치
        cols = st.columns(4)
        for idx, img_path_str in enumerate(img_paths):
            img_p = Path(img_path_str)
            menu_name = img_p.stem
            col = cols[idx % 4]
            with col:
                # 메뉴 카드 (이미지 + 주문 버튼)
                st.markdown('<div class="menu-card">', unsafe_allow_html=True)
                st.image(
                    load_square_image(img_p),
                    caption=menu_name,
                    use_container_width=True,
                )
                st.markdown('</div>', unsafe_allow_html=True)

                # 주문 버튼 — 고유 key로 각 버튼을 구분
                btn_key = f"order_{restaurant}_{menu_name}"
                if st.button(f"🛒 {menu_name}", key=btn_key):
                    current_name = st.session_state["user_name"]
                    if not current_name:
                        # 이름 미입력 시 경고 메시지 세팅 후 재실행
                        st.session_state["last_order_msg"] = (
                            "warning", "⚠️ 이름을 먼저 입력해주세요!"
                        )
                        st.rerun()
                    else:
                        # Google Sheets에 주문 저장 — 성공 시에만 성공 메시지 세팅 후 재실행
                        if save_order_to_gsheet(current_name, menu_name, restaurant):
                            st.session_state["last_order_msg"] = (
                                "success",
                                f"✅ {current_name}님의 [{restaurant}] {menu_name} 주문이 완료되었습니다!",
                            )
                            st.rerun()
                        # 실패 시 st.error는 save_order_to_gsheet 내부에서 이미 표시됨

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    subprocess.run(
        [sys.executable, "-m", "streamlit", "run", __file__] + sys.argv[1:]
    )

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

    subprocess.run(
        [sys.executable, "-m", "streamlit", "run", __file__] + sys.argv[1:]
    )
if __name__ == "__main__":
    import subprocess
    import sys

    # #region agent log: __main__ block triggered (H-C)
    _dbg("app.py:__main__", "__main__ 블록 진입", {
        "GUARD_env": _os.environ.get("_ST_GUARD", "없음"),
        "will_launch": not bool(_os.environ.get("_ST_GUARD")),
        "pid": _os.getpid(),
    }, hyp="C")
    # #endregion

    if not _os.environ.get("_ST_GUARD"):
        env = _os.environ.copy()
        env["_ST_GUARD"] = "1"
        subprocess.run(
            [sys.executable, "-m", "streamlit", "run", __file__] + sys.argv[1:],
            env=env,
        )
