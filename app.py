import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from streamlit_gsheets import GSheetsConnection

# ── 상수 ──────────────────────────────────────────────────────────────────────
RESTAURANTS = ['경성카츠', '광교뚝배기', '바비든든', '포포420']
IMAGE_EXTS  = {'.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp'}
COLUMNS     = ['이름', '메뉴', '식당', '주문시간']
KST         = ZoneInfo('Asia/Seoul')
BASE_DIR    = Path(__file__).parent

# ── Google Sheets 연결 ─────────────────────────────────────────────────────────
conn = st.connection("gsheets", type=GSheetsConnection)


# ── 시간 헬퍼 ──────────────────────────────────────────────────────────────────
def kst_now() -> datetime:
    """서버 시간대와 무관하게 한국 시간(KST)을 반환합니다."""
    return datetime.now(KST)


def today_kst_str() -> str:
    return kst_now().strftime('%Y-%m-%d')


# ── 이미지 헬퍼 ────────────────────────────────────────────────────────────────
def get_menu_images(restaurant: str) -> list[Path]:
    """식당 폴더 내 이미지 파일 목록을 반환합니다 (OS 독립적 경로 처리)."""
    folder = BASE_DIR / restaurant
    if not folder.exists():
        return []
    return sorted(
        [f for f in folder.iterdir() if f.suffix.lower() in IMAGE_EXTS],
        key=lambda p: p.stem,
    )


# ── Google Sheets CRUD ─────────────────────────────────────────────────────────
def load_all_data() -> pd.DataFrame:
    """캐시 없이 시트 전체 데이터를 읽어옵니다."""
    try:
        df = conn.read(ttl=0)
        if df is None or df.empty:
            return pd.DataFrame(columns=COLUMNS)
        # 헤더가 일치하지 않는 경우를 방어적으로 처리
        if not all(c in df.columns for c in COLUMNS):
            df.columns = COLUMNS[: len(df.columns)]
        df = df[COLUMNS].dropna(how='all').reset_index(drop=True)
        df['주문시간'] = df['주문시간'].astype(str)
        return df
    except Exception:
        return pd.DataFrame(columns=COLUMNS)


def load_today_data() -> pd.DataFrame:
    """오늘(KST) 날짜의 주문만 필터링하여 반환합니다."""
    df = load_all_data()
    today = today_kst_str()
    return df[df['주문시간'].str.startswith(today)].reset_index(drop=True)


def append_order(name: str, menu: str, restaurant: str) -> None:
    """새 주문 1행을 시트 맨 끝에 안전하게 추가합니다.

    전체 덮어쓰기(read→concat→update) 대신 gspread의 append_row를 직접 사용하여
    동시 주문 시 데이터가 유실되는 Race Condition을 방지합니다.
    """
    now_str = kst_now().strftime('%Y-%m-%d %H:%M:%S')
    client    = conn._instance
    sheet     = client.open_by_url(st.secrets["connections"]["gsheets"]["spreadsheet"])
    worksheet = sheet.sheet1
    worksheet.append_row([name, menu, restaurant, now_str])


def cancel_my_orders(name: str) -> None:
    """오늘(KST) 날짜 기준으로 해당 이름의 주문을 모두 삭제합니다."""
    df = load_all_data()
    if df.empty:
        return
    today = today_kst_str()
    keep = ~((df['이름'] == name) & (df['주문시간'].str.startswith(today)))
    conn.update(data=df[keep].reset_index(drop=True))


# ── 페이지 설정 ────────────────────────────────────────────────────────────────
st.set_page_config(page_title='팀 점심 메뉴 취합', layout='wide')
st.title('팀 점심 메뉴 취합 시스템')

# ── 이름 입력 ──────────────────────────────────────────────────────────────────
name = st.text_input('이름을 입력하세요', placeholder='예) 홍길동')
st.divider()

# ── 식당 탭 및 메뉴 선택 ────────────────────────────────────────────────────────
tabs = st.tabs(RESTAURANTS)

for tab_idx, (tab, restaurant) in enumerate(zip(tabs, RESTAURANTS)):
    with tab:
        images = get_menu_images(restaurant)
        if not images:
            st.warning(f'**{restaurant}** 폴더에 메뉴 이미지가 없습니다.')
            continue

        cols = st.columns(4)
        for img_idx, img_path in enumerate(images):
            menu_name = img_path.stem
            with cols[img_idx % 4]:
                st.image(str(img_path), use_container_width=True)
                st.caption(f'**{menu_name}**')
                if st.button('선택', key=f'order_{tab_idx}_{img_idx}'):
                    if not name.strip():
                        st.error('이름을 먼저 입력해 주세요.')
                    else:
                        with st.spinner('주문 중...'):
                            append_order(name.strip(), menu_name, restaurant)
                        st.success(f'**{menu_name}** 주문 완료!')
                        st.rerun()

st.divider()

# ── 내 주문 전체 취소 ──────────────────────────────────────────────────────────
st.subheader('내 주문 관리')
if st.button('내 주문 전체 취소 (오늘 날짜 기준)', type='secondary'):
    if not name.strip():
        st.error('이름을 입력해 주세요.')
    else:
        with st.spinner('취소 처리 중...'):
            cancel_my_orders(name.strip())
        st.success(f'**{name.strip()}** 님의 오늘 주문이 모두 취소되었습니다.')
        st.rerun()

st.divider()

# ── 팀원 주문 현황 ─────────────────────────────────────────────────────────────
st.subheader('팀원 주문 현황')

# 버튼 클릭 시 데이터를 session_state에 저장.
# 표를 그리는 로직은 버튼 밖에 분리하여 탭 전환 등 리렌더링 후에도 유지됩니다.
if st.button('주문 현황 보기 (새로고침)', type='primary'):
    with st.spinner('데이터를 불러오는 중...'):
        st.session_state['today_df'] = load_today_data()

if 'today_df' in st.session_state:
    today_df = st.session_state['today_df']
    if today_df.empty:
        st.info('오늘 주문 내역이 없습니다.')
    else:
        st.caption(f'기준 날짜 (KST) : {today_kst_str()}  |  총 {len(today_df)}건')

        col_left, col_right = st.columns(2)

        with col_left:
            st.markdown('#### 전체 주문 내역')
            display = today_df[['이름', '식당', '메뉴', '주문시간']].copy()
            display.index = range(1, len(display) + 1)
            st.dataframe(display, use_container_width=True)

        with col_right:
            st.markdown('#### 메뉴별 주문 집계')
            summary = (
                today_df
                .groupby(['식당', '메뉴'], sort=False)
                .agg(수량=('이름', 'count'), 주문자=('이름', lambda x: ', '.join(x)))
                .reset_index()
                .sort_values(['식당', '수량'], ascending=[True, False])
            )
            summary.index = range(1, len(summary) + 1)
            st.dataframe(summary, use_container_width=True)
