from __future__ import annotations

import json
import html as html_lib
import os
import urllib.parse
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "ERICA_restaurants_geocoded_google.xlsx"
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
FAVORITES_FILE = DATA_DIR / "favorites.json"
REVIEWS_FILE = DATA_DIR / "reviews.json"
USER_RESTAURANTS_FILE = DATA_DIR / "user_restaurants.json"
PROFILE_FILE = DATA_DIR / "profile.json"

st.set_page_config(
    page_title="에리카 뭐먹지",
    page_icon="🍚",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Minimal blue tech style
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;600;700;800&display=swap');
:root { --ink:#172b45; --muted:#718096; --line:#e5eaf1; --soft:#f4f7fb; --navy:#244a78; --blue:#3978c6; --pale:#edf4ff; }
html, body, [class*="css"] { font-family:'Pretendard','Apple SD Gothic Neo',sans-serif; }
.stApp { background:#fff; color:var(--ink); }
.block-container { max-width:1160px; padding-top:1.35rem; padding-bottom:3.5rem; }
header[data-testid="stHeader"] { background:rgba(255,255,255,.96); }
footer {visibility:hidden;}
.brand {font-size:1.55rem;font-weight:800;letter-spacing:-.055em;color:#172b45 !important;-webkit-text-fill-color:#172b45 !important;opacity:1 !important;margin:0;}
.brand .dot {color:var(--blue) !important;-webkit-text-fill-color:var(--blue) !important;opacity:1 !important;}
/* 상단 브랜드 문구의 색상이 다른 스타일에 덮이지 않도록 고정 */
.stApp .block-container p.brand {color:#172b45 !important;-webkit-text-fill-color:#172b45 !important;opacity:1 !important;}
.brand-sub {font-size:.85rem;color:var(--muted);margin:.2rem 0 1.15rem;}
.hero {background:#f5f8fd;border:1px solid #e1eaf7;border-radius:18px;padding:27px 29px;margin:.45rem 0 1.2rem;}
.hero-kicker {font-size:.72rem;font-weight:700;letter-spacing:.12em;color:#6682a5;margin-bottom:.55rem;}
.hero-title {font-size:2rem;line-height:1.25;font-weight:800;letter-spacing:-.06em;color:#172b45;margin:0 0 .45rem;}
.hero-copy {font-size:.94rem;color:#64748b;margin:0;}
.section-heading {font-size:1.12rem;font-weight:800;letter-spacing:-.04em;color:#172b45;margin:.5rem 0 .8rem;}
.helper {font-size:.82rem;color:#718096;}
div[data-testid="stTabs"] [data-baseweb="tab-list"] {gap:1rem;border-bottom:1px solid #e8edf4;}
div[data-testid="stTabs"] button {
    font-weight:700;
    font-size:1.25rem;
    color:#64748b;
}
/* 실제 상단 탭 라벨에도 글자 크기 적용: 색상은 기존 설정 유지 */
div[data-testid="stTabs"] button[data-baseweb="tab"] *,
div[data-testid="stTabs"] button[data-baseweb="tab"] {
    font-size:1.25rem !important;
}
div[data-testid="stTabs"] button[aria-selected="true"] {color:#244a78;border-bottom-color:#3978c6;}
.stButton button, .stLinkButton a, .stFormSubmitButton button {border-radius:10px;font-weight:650;}
.stButton button[kind="primary"], .stFormSubmitButton button[kind="primary"] {background:#2f6fb8;border-color:#2f6fb8;color:white;}
.stButton button[kind="primary"]:hover, .stFormSubmitButton button[kind="primary"]:hover {background:#245b99;border-color:#245b99;}
div[data-testid="stMetric"] {background:#f6f8fb;border:1px solid #e8edf3;border-radius:13px;padding:12px 14px;}
div[data-testid="stMetricLabel"] {color:#718096;}
div[data-testid="stMetricValue"] {color:#172b45;font-weight:750;}
.restaurant-card {border:1px solid #e5eaf1;border-radius:14px;padding:15px 16px;background:#fff;margin-bottom:10px;}
.card-title {font-size:1.02rem;font-weight:750;color:#172b45;margin:0 0 .25rem;}
.pill {display:inline-block;padding:4px 9px;border-radius:99px;background:#edf4ff;color:#315f91;font-size:.73rem;font-weight:650;margin-right:4px;}
div[data-testid="stExpander"] {border-color:#e5eaf1;border-radius:12px;}
input, textarea {border-radius:9px !important;}
/* Streamlit widgets: blue accent instead of default red */
.stSlider [data-baseweb="slider"] div[role="slider"] {background:#3978c6 !important;border-color:#3978c6 !important;}
.stSlider [data-baseweb="slider"] div[data-testid="stTickBar"] {background:#dce8f7 !important;}
hr {border-color:#edf0f4;}
/* Force a consistent blue accent on Streamlit's native controls */
:root { accent-color:#3978c6 !important; }
[data-testid="stSlider"] [role="slider"] {background:#3978c6 !important;border-color:#3978c6 !important;}
[data-testid="stSlider"] [data-baseweb="slider"] > div > div {background:#3978c6 !important;}
[data-testid="stCheckbox"] input, [data-testid="stRadio"] input {accent-color:#3978c6 !important;}
[data-testid="stCheckbox"] label, [data-testid="stRadio"] label {color:#26364b !important;}
[data-testid="stCheckbox"] label, [data-testid="stCheckbox"] label p, [data-testid="stCheckbox"] [data-testid="stMarkdownContainer"] {background:transparent !important; box-shadow:none !important;}
[data-testid="stFormSubmitButton"] button, [data-testid="stButton"] button[kind="primary"] {background:#2f6fb8 !important;border-color:#2f6fb8 !important;color:#fff !important;}
[data-testid="stFormSubmitButton"] button:hover, [data-testid="stButton"] button[kind="primary"]:hover {background:#245b99 !important;border-color:#245b99 !important;}
</style>
""", unsafe_allow_html=True)

CATEGORY_MAP = {
    "한식": "한식", "닭갈비": "한식", "삼계탕": "한식", "고기요리": "한식",
    "순대, 순댓국": "한식", "순댓국": "한식", "찌개, 전골": "한식", "냉면": "한식",
    "칼국수": "한식", "감자탕": "한식", "치킨, 닭강정": "한식",
    "양식": "양식", "이탈리아음식": "양식", "피자": "양식",
    "샌드위치": "패스트푸드/간편식", "햄버거": "패스트푸드/간편식",
    "핫도그": "패스트푸드/간편식",
    "분식": "분식", "종합분식": "분식", "떡볶이": "분식",
    "중식": "중식", "마라탕": "중식",
    "일식": "일식", "돈가스": "일식",
    "베트남음식": "세계음식", "아시아음식": "세계음식", "인도음식": "세계음식",
    "멕시코음식": "세계음식", "멕시코, 남미음식": "세계음식",
}
CATEGORY_ORDER = ["한식", "양식", "패스트푸드/간편식", "분식", "중식", "일식", "세계음식", "기타"]
PRIORITIES = {
    "균형 추천": (.30, .25, .25, .20),
    "가성비 우선": (.50, .20, .15, .15),
    "거리 우선": (.20, .50, .15, .15),
    "인기·평점 우선": (.15, .15, .35, .35),
}


def read_json(path: Path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return default


def write_json(path: Path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def won(value) -> str:
    return "정보 없음" if pd.isna(value) else f"{int(value):,}원"


def naver_link(name: str) -> str:
    query = urllib.parse.quote(f"안산 상록구 {name}")
    return f"https://map.naver.com/p/search/{query}"


@st.cache_data(show_spinner="식당 데이터를 불러오는 중입니다...")
def load_restaurants(path: str) -> pd.DataFrame:
    frame = pd.read_excel(path, sheet_name="식당데이터")
    required = ["상호", "주소", "업종", "대표 메뉴", "대표메뉴 가격(원)",
                "대표메뉴 1인분 가격(원)", "도보 거리(m)", "도보 시간(분)",
                "리뷰 수", "별점", "위도", "경도"]
    missing = [col for col in required if col not in frame.columns]
    if missing:
        raise ValueError("엑셀에 필요한 열이 없습니다: " + ", ".join(missing))
    numeric = ["대표메뉴 가격(원)", "대표메뉴 1인분 가격(원)", "도보 거리(m)",
               "도보 시간(분)", "리뷰 수", "별점", "위도", "경도"]
    # Excel에 숫자가 "9,900"처럼 문자열로 저장된 경우 쉼표 때문에
    # pd.to_numeric(errors="coerce")가 전부 NaN으로 바뀌는 문제를 방지합니다.
    for col in numeric:
        cleaned = (
            frame[col]
            .astype("string")
            .str.replace(",", "", regex=False)
            .str.replace("원", "", regex=False)
            .str.strip()
        )
        frame[col] = pd.to_numeric(cleaned, errors="coerce")
    frame["대분류"] = frame["업종"].astype(str).map(CATEGORY_MAP).fillna("기타")
    frame["식당ID"] = frame["상호"].astype(str) + "|" + frame["주소"].astype(str)
    return frame


def recommend(df: pd.DataFrame, budget: int, total_minutes: int, party: int,
              categories: list[str], priority: str, solo_only: bool = False) -> pd.DataFrame:
    result = df.copy()

    if categories:
        result = result[result["대분류"].isin(categories)]

    # 혼밥 전용 옵션: 대표 메뉴 가격이 1인분 가격보다 높은 2인 이상 메뉴를 제외합니다.
    if solo_only:
        result = result[
            result["대표메뉴 가격(원)"].notna()
            & result["대표메뉴 1인분 가격(원)"].notna()
            & (result["대표메뉴 가격(원)"] <= result["대표메뉴 1인분 가격(원)"])
        ]

    # 예산은 1인당 기준이며, 1인분 가격이 없을 때 대표 메뉴 가격을 보조값으로 사용합니다.
    per_person = result["대표메뉴 1인분 가격(원)"].fillna(result["대표메뉴 가격(원)"])
    result = result[per_person.notna() & (per_person <= budget)].copy()
    result["_비교가격"] = per_person.loc[result.index]

    # 왕복 이동 시간만 고려합니다.
    result = result[
        result["도보 시간(분)"].notna()
        & (result["도보 시간(분)"] * 2 <= total_minutes)
    ]
    if result.empty:
        return result

    max_dist = max(float(df["도보 거리(m)"].max()), 1)
    max_reviews = max(float(np.log1p(df["리뷰 수"].fillna(0).max())), 1)
    global_mean_rating = df["별점"].mean()
    if pd.isna(global_mean_rating):
        global_mean_rating = 3.5

    result["가성비"] = (1 - result["_비교가격"] / max(budget, 1)).clip(0, 1)
    result["접근성"] = (
        1 - result["도보 거리(m)"].fillna(max_dist) / max_dist
    ).clip(0, 1)
    result["인기도"] = (
        np.log1p(result["리뷰 수"].fillna(0)) / max_reviews
    ).clip(0, 1)

    # 베이지안 평점: 리뷰가 적을수록 전체 평균에 가깝게 보정합니다.
    prior_weight = 300
    review_count = result["리뷰 수"].fillna(0).clip(lower=0)
    raw_rating = result["별점"].fillna(global_mean_rating)
    bayes_rating = (
        review_count * raw_rating + prior_weight * global_mean_rating
    ) / (review_count + prior_weight)
    result["보정별점"] = bayes_rating.round(2)
    result["평점"] = (bayes_rating / 5.0).clip(0, 1)

    weights = PRIORITIES[priority]
    score_columns = ["가성비", "접근성", "인기도", "평점"]
    result["추천점수"] = (
        100 * sum(w * result[col] for w, col in zip(weights, score_columns))
    ).round(1)

    return result.sort_values(["추천점수", "리뷰 수"], ascending=False)


def render_google_map(map_df: pd.DataFrame, height: int = 560):
    try:
        key = st.secrets.get("google_maps", {}).get("api_key", "")
    except Exception:
        key = ""
    key = key or os.getenv("GOOGLE_MAPS_JS_API_KEY", "")
    if not key:
        st.info("Google Maps API 키를 설정하면 실제 지도가 표시됩니다.")
        st.code('[google_maps]\napi_key = "YOUR_MAPS_JAVASCRIPT_API_KEY"', language="toml")
        return
    valid = map_df.dropna(subset=["위도", "경도"]).copy()
    if valid.empty:
        st.warning("표시할 좌표 데이터가 없습니다.")
        return
    places = []
    for _, row in valid.iterrows():
        places.append({
            "name": str(row["상호"]), "menu": str(row["대표 메뉴"]),
            "category": str(row["대분류"]), "price": won(row["대표메뉴 1인분 가격(원)"]),
            "walk": "정보 없음" if pd.isna(row["도보 시간(분)"]) else f"{int(row['도보 시간(분)'])}분",
            "rating": "정보 없음" if pd.isna(row["별점"]) else f"{float(row['별점']):.1f}",
            "address": str(row["주소"]), "lat": float(row["위도"]), "lng": float(row["경도"]),
        })
    payload = json.dumps(places, ensure_ascii=False).replace("</", "<\\/")
    key_json = json.dumps(key)
    html = r'''<!doctype html><html><head><meta charset="utf-8"><style>
      html,body,#map{height:100%;width:100%;margin:0;padding:0;font-family:Arial,sans-serif}
      #map{border-radius:14px;overflow:hidden}.gm-style-iw{font-size:13px}
    </style></head><body><div id="map"></div><script>
      const API_KEY=__KEY__; const PLACES=__PLACES__;
      function esc(v){return String(v??"").replace(/[&<>\"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'\"':"&quot;","'":"&#39;"}[c]));}
      async function initMap(){
        try{
          const {Map}=await google.maps.importLibrary("maps");
          const center={lat:PLACES.reduce((a,p)=>a+p.lat,0)/PLACES.length,lng:PLACES.reduce((a,p)=>a+p.lng,0)/PLACES.length};
          const map=new Map(document.getElementById("map"),{center,zoom:16,mapTypeControl:false,streetViewControl:false,fullscreenControl:true,gestureHandling:"greedy"});
          const info=new google.maps.InfoWindow(); const bounds=new google.maps.LatLngBounds();
          PLACES.forEach(p=>{
            const position={lat:p.lat,lng:p.lng}; bounds.extend(position);
            const marker=new google.maps.Marker({map,position,title:p.name});
            marker.addListener("click",()=>{
              info.setContent(`<div style="min-width:190px;line-height:1.6;padding:2px"><b style="font-size:15px">${esc(p.name)}</b><br><span>${esc(p.category)} · ${esc(p.menu)}</span><br><b>${esc(p.price)}</b><br>도보 ${esc(p.walk)} · ★ ${esc(p.rating)}<br><small>${esc(p.address)}</small></div>`);
              info.open({map,anchor:marker});
            });
          });
          if(PLACES.length>1) map.fitBounds(bounds);
        }catch(e){document.getElementById("map").innerHTML="<div style='padding:20px;color:#a33'>Google 지도를 불러오지 못했습니다. 브라우저 콘솔 오류를 확인해 주세요.<br><small>"+String(e)+"</small></div>";}
      }
      window.initMap=initMap;
      const script=document.createElement("script");
      script.src="https://maps.googleapis.com/maps/api/js?key="+encodeURIComponent(API_KEY)+"&v=weekly&loading=async&callback=initMap";
      script.async=true; script.onerror=()=>{document.getElementById("map").innerHTML="Google Maps 스크립트를 불러오지 못했습니다. API 키와 결제 설정을 확인해 주세요.";};
      document.head.appendChild(script);
    </script></body></html>'''
    html = html.replace("__KEY__", key_json).replace("__PLACES__", payload)
    components.html(html, height=height, scrolling=False)
    st.caption(f"좌표가 포함된 식당 {len(valid)}곳 · 마커를 선택하면 메뉴와 가격을 볼 수 있습니다.")


if not DATA_FILE.exists():
    st.error(f"좌표 엑셀 파일을 찾을 수 없습니다: {DATA_FILE.name}. app.py와 같은 폴더에 넣어주세요.")
    st.stop()
try:
    restaurants = load_restaurants(str(DATA_FILE))
except Exception as exc:
    st.error(f"식당 엑셀을 읽는 중 문제가 발생했습니다: {exc}")
    st.stop()

favorites = read_json(FAVORITES_FILE, [])
reviews = read_json(REVIEWS_FILE, [])
user_restaurants = read_json(USER_RESTAURANTS_FILE, [])
profile = read_json(PROFILE_FILE, {"nickname": "나의 프로필", "intro": "오늘도 맛있는 한 끼를 찾아보세요."})

st.markdown('<p class="brand">에리카 뭐먹지<span class="dot">.</span></p><p class="brand-sub">한양대 ERICA 주변, 오늘의 식사를 더 합리적으로</p>', unsafe_allow_html=True)
tabs = st.tabs(["홈", "즐겨찾기", "맛집 등록", "마이페이지"])

# HOME
with tabs[0]:
    st.markdown('<div class="hero"><div class="hero-kicker">ERICA MEAL FINDER</div><div class="hero-title">오늘 뭐 먹지?</div><p class="hero-copy">예산과 시간, 취향에 맞는 식당을 간단하게 찾아보세요.</p></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-heading">식사 조건</div>', unsafe_allow_html=True)
    with st.form("recommend_form"):
        c1, c2, c3 = st.columns(3)
        budget = c1.slider("1인당 예산", 3500, 30000, 10000, 500, format="%d원")
        available = c2.slider("왕복 이동 시간", 25, 150, 100, 5, format="%d분")
        party = c3.selectbox("식사 인원", [1, 2, 3, 4], index=0, format_func=lambda x: f"{x}명")
        solo_only = st.checkbox(
            "혼밥 가능 식당만 보기",
            value=False,
            help="대표 메뉴 가격과 1인분 가격을 비교해 2인 이상 메뉴를 제외합니다."
        )
        st.markdown("**음식 종류**")
        st.caption("원하는 음식 종류를 선택해 주세요. 여러 개를 선택할 수 있습니다.")
        available_categories = [c for c in CATEGORY_ORDER if c in restaurants["대분류"].unique()]
        selected_categories = []
        category_cols = st.columns(4)
        for i, category in enumerate(available_categories):
            with category_cols[i % 4]:
                if st.checkbox(category, value=True, key=f"home_cat_{category}"):
                    selected_categories.append(category)
        st.markdown("**추천 기준**")
        priority = st.radio("추천 기준", list(PRIORITIES), horizontal=True, label_visibility="collapsed")
        submitted = st.form_submit_button("✦  추천받기", type="primary", use_container_width=True)
    if submitted:
        st.session_state["recommendation_results"] = recommend(
            restaurants, budget, available, party, selected_categories, priority, solo_only
        )
        st.session_state["recommendation_params"] = {
            "budget": budget, "available": available, "party": party,
            "categories": selected_categories, "priority": priority,
            "solo_only": solo_only,
        }
    results = st.session_state.get("recommendation_results")
    st.markdown('<div class="section-heading">추천 결과</div>', unsafe_allow_html=True)
    if results is None:
        st.info("식사 조건을 선택한 뒤 ‘추천받기’를 눌러주세요.")
    elif results.empty:
        st.info("조건에 맞는 식당이 없습니다. 1인당 예산, 이동 시간 또는 음식 종류를 조정해 보세요.")
    else:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("조건에 맞는 식당", f"{len(results)}곳")
        m2.metric("평균 1인분 가격", won(results["대표메뉴 1인분 가격(원)"].mean()))
        m3.metric("평균 편도 도보", f"{results['도보 시간(분)'].mean():.1f}분")
        m4.metric("추천 1위", str(results.iloc[0]["상호"]))
        st.markdown('<div class="section-heading">추천 식당 TOP 5</div>', unsafe_allow_html=True)
        top5 = results.head(5)
        card_cols = st.columns(2)
        for idx, (_, row) in enumerate(top5.iterrows()):
            with card_cols[idx % 2]:
                with st.container(border=True):
                    st.markdown(
                        f'<div style="display:flex;align-items:center;gap:9px;margin-bottom:8px">'
                        f'<span style="background:#eaf2ff;color:#245b99;border-radius:8px;padding:5px 9px;font-weight:800">{idx+1}위</span>'
                        f'<span style="font-size:.78rem;color:#64748b">{row["대분류"]} · 추천 {row["추천점수"]}점</span></div>',
                        unsafe_allow_html=True
                    )
                    st.markdown(f'<div class="card-title">{row["상호"]}</div>', unsafe_allow_html=True)
                    st.write(str(row["대표 메뉴"]))
                    st.markdown(f"**{won(row['_비교가격'])} / 1인**")
                    raw_rating_text = "원본 평점 정보 없음" if pd.isna(row["별점"]) else f"원본 ★ {row['별점']:.1f}"
                    rating = f"보정 ★ {row['보정별점']:.2f}"
                    walk = "정보 없음" if pd.isna(row["도보 시간(분)"]) else f"도보 {int(row['도보 시간(분)'])}분"
                    review_count = 0 if pd.isna(row["리뷰 수"]) else int(row["리뷰 수"])
                    st.caption(f"{rating} · {raw_rating_text} · 리뷰 {review_count:,}개 · {walk}")
                    rid = row["식당ID"]
                    if st.button("♥ 즐겨찾기 해제" if rid in favorites else "♡ 즐겨찾기 추가",
                                 key=f"homefav_{idx}_{rid}", use_container_width=True):
                        favorites = [x for x in favorites if x != rid] if rid in favorites else favorites + [rid]
                        write_json(FAVORITES_FILE, favorites)
                        st.rerun()
                    st.link_button("네이버 지도에서 보기", naver_link(row["상호"]), use_container_width=True)

        if len(results) > 5:
            st.markdown('<div class="section-heading">전체 추천 순위 · TOP 15</div>', unsafe_allow_html=True)
            ranked = results.head(15).copy().reset_index(drop=True)
            ranked.insert(0, "순위", range(1, len(ranked) + 1))
            ranked["1인분 가격"] = ranked["_비교가격"].map(won)
            ranked["도보"] = ranked["도보 시간(분)"].map(
                lambda x: "정보 없음" if pd.isna(x) else f"{int(x)}분"
            )
            ranked["평점 표시"] = ranked["보정별점"].map(
                lambda x: "정보 없음" if pd.isna(x) else f"보정 ★ {float(x):.2f}"
            )
            st.dataframe(
                ranked[["순위", "상호", "대분류", "대표 메뉴", "1인분 가격", "도보", "평점 표시", "리뷰 수", "추천점수"]],
                use_container_width=True, hide_index=True
            )

        else:
            st.caption(f"조건에 맞는 식당 {len(results)}곳을 찾았습니다. 현재 조건에서는 TOP 15까지 표시할 결과가 없습니다.")

        with st.expander("세부 추천 점수 산출 내역 보기"):
            detail_df = results[
                ["상호", "추천점수", "가성비", "접근성", "인기도", "평점", "보정별점", "리뷰 수"]
            ].copy()
            for col in ["가성비", "접근성", "인기도", "평점"]:
                detail_df[col] = (detail_df[col] * 100).round(1)
            detail_df = detail_df.rename(columns={
                "가성비": "가성비 점수(%)",
                "접근성": "접근성 점수(%)",
                "인기도": "인기도 점수(%)",
                "평점": "보정 평점 점수(%)",
                "보정별점": "베이지안 보정 별점",
            })
            st.caption("각 항목은 0~100점으로 환산한 값이며, 최종 추천점수에는 선택한 추천 기준의 가중치가 적용됩니다.")
            st.dataframe(detail_df, use_container_width=True, hide_index=True)


# FAVORITES: map first, saved restaurants below
with tabs[1]:
    st.markdown('<div class="section-heading">즐겨찾기</div>', unsafe_allow_html=True)
    st.caption("ERICA 주변 식당 지도를 먼저 확인하고, 저장한 식당을 아래에서 관리할 수 있습니다.")

    st.markdown('<div class="section-heading">주변 지도</div>', unsafe_allow_html=True)
    map_data = restaurants.dropna(subset=["위도", "경도"]).copy()
    map_categories = [c for c in CATEGORY_ORDER if c in map_data["대분류"].unique()]
    selected_map_categories = st.multiselect(
        "표시할 음식 종류", map_categories, default=map_categories,
        key="favorite_map_categories"
    )
    map_data = map_data[map_data["대분류"].isin(selected_map_categories)]
    render_google_map(map_data, height=520)

    st.markdown("---")
    st.markdown('<div class="section-heading">저장한 식당</div>', unsafe_allow_html=True)
    favorite_df = restaurants[restaurants["식당ID"].isin(favorites)].copy()
    if favorite_df.empty:
        st.info("아직 저장한 식당이 없습니다. 홈에서 마음에 드는 식당을 저장해 보세요.")
    else:
        st.caption(f"총 {len(favorite_df)}곳을 즐겨찾기에 저장했습니다.")
        for _, row in favorite_df.iterrows():
            with st.container(border=True):
                left, right = st.columns([5, 1])
                left.markdown(f"**{row['상호']}** · {row['대표 메뉴']}")
                left.caption(f"{row['대분류']} · {won(row['대표메뉴 1인분 가격(원)'])} · {row['주소']}")
                if right.button("삭제", key=f"remove_{row['식당ID']}"):
                    favorites = [x for x in favorites if x != row["식당ID"]]
                    write_json(FAVORITES_FILE, favorites)
                    st.rerun()
                left.link_button("네이버 지도", naver_link(row["상호"]))


# ADD RESTAURANT
with tabs[2]:
    st.markdown('<div class="section-heading">맛집 등록</div>', unsafe_allow_html=True)
    st.caption("직접 알고 있는 식당 정보를 추가할 수 있습니다. 등록 정보는 data 폴더에 저장됩니다.")
    with st.form("add_restaurant", clear_on_submit=True):
        name = st.text_input("식당 이름 *")
        address = st.text_input("주소 *")
        menu = st.text_input("대표 메뉴")
        category = st.selectbox("음식 종류", ["한식", "양식", "패스트푸드/간편식", "분식", "중식", "일식", "세계음식", "기타"])
        price = st.number_input("1인분 가격", min_value=0, max_value=100000, value=10000, step=500)
        submitted_restaurant = st.form_submit_button("식당 등록", type="primary")
        if submitted_restaurant:
            if not name.strip() or not address.strip():
                st.error("식당 이름과 주소는 필수입니다.")
            else:
                user_restaurants.append({"name": name.strip(), "address": address.strip(), "menu": menu.strip(), "category": category, "price": int(price)})
                write_json(USER_RESTAURANTS_FILE, user_restaurants)
                st.success("식당을 등록했습니다.")
    if user_restaurants:
        st.markdown("#### 내가 등록한 식당")
        st.dataframe(pd.DataFrame(user_restaurants), use_container_width=True, hide_index=True)

# MY PAGE
with tabs[3]:
    st.markdown('<div class="section-heading">마이페이지</div>', unsafe_allow_html=True)

    profile_left, profile_right = st.columns([1, 3])
    with profile_left:
        st.markdown(
            """<div style="width:112px;height:112px;border-radius:50%;background:#edf4ff;
            display:flex;align-items:center;justify-content:center;font-size:3.2rem;
            border:1px solid #dce8f7;margin:0 auto;">🍚</div>""",
            unsafe_allow_html=True
        )
    with profile_right:
        safe_nickname = html_lib.escape(str(profile.get("nickname", "나의 프로필")))
        safe_intro = html_lib.escape(str(profile.get("intro", "오늘도 맛있는 한 끼를 찾아보세요.")))
        st.markdown(
            f"""<div style="padding:8px 0;">
            <div style="font-size:.78rem;font-weight:750;letter-spacing:.08em;color:#6682a5;">ERICA MEAL PROFILE</div>
            <div style="font-size:1.55rem;font-weight:800;color:#1f3858;margin:.2rem 0;">{safe_nickname}</div>
            <div style="color:#718096;font-size:.92rem;">{safe_intro}</div>
            </div>""",
            unsafe_allow_html=True
        )

    # 프로필 아바타와 입력 폼 사이에 여백을 두어 테두리와 겹쳐 보이지 않게 합니다.
    st.markdown('<div style="height:28px;"></div>', unsafe_allow_html=True)

    with st.form("profile_form"):
        nickname = st.text_input("닉네임", value=profile.get("nickname", "나의 프로필"), max_chars=20)
        intro = st.text_input("한 줄 소개", value=profile.get("intro", "오늘도 맛있는 한 끼를 찾아보세요."), max_chars=60)
        save_profile = st.form_submit_button("프로필 저장", type="primary")
        if save_profile:
            profile = {"nickname": nickname.strip() or "나의 프로필",
                       "intro": intro.strip() or "오늘도 맛있는 한 끼를 찾아보세요."}
            write_json(PROFILE_FILE, profile)
            st.success("프로필을 저장했습니다.")
            st.rerun()

    st.markdown("#### 나의 활동")
    a, b, c = st.columns(3)
    a.metric("전체 식당", f"{len(restaurants)}곳")
    b.metric("즐겨찾기", f"{len(favorites)}곳")
    c.metric("직접 등록", f"{len(user_restaurants)}곳")

    st.markdown("#### 데이터 정보")
    st.write(f"- 좌표가 포함된 식당: {restaurants.dropna(subset=['위도', '경도']).shape[0]}곳")
    st.write(f"- 기준 데이터: `{DATA_FILE.name}`")
    st.caption("좌표는 주소 기준점이며 실제 식당 출입구와 다를 수 있습니다.")
    st.caption("즐겨찾기와 직접 등록한 식당, 프로필은 프로젝트의 data 폴더에 저장됩니다.")
    if st.button("즐겨찾기 전체 초기화"):
        write_json(FAVORITES_FILE, [])
        st.success("즐겨찾기를 초기화했습니다.")
        st.rerun()


st.markdown("<hr><div style='text-align:center;color:#9aa5b1;font-size:.78rem'>ERICA · 더 합리적인 식사 선택</div>", unsafe_allow_html=True)
