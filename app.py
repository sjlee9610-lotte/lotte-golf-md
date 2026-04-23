"""
app.py — 롯데 석주 주간 골프 MD 뉴스 알림 (Gemini 무료 운용 버전)

실행 전 준비:
  pip install streamlit google-genai openpyxl requests beautifulsoup4 pandas
  export GEMINI_API_KEY="AIza..."
  # 선택:
  # export GEMINI_MODEL="gemini-2.5-flash"
  streamlit run app.py
"""

import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from google import genai

# ══════════════════════════════════════════════════════════════
# 설정
# ══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="롯데 석주 주간 골프 MD 뉴스 알림",
    page_icon="⛳",
    layout="wide",
)

MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

NEWS_FILE = Path("weekly_news.json")
STORE_FILE = Path("store_profiles.xlsx")

CATEGORY_ICON = {
    "골프 브랜드": "🏌️",
    "골프 경기": "🏆",
    "골프장 현황": "⛳",
    "기타 이슈": "📰",
}

PRIORITY_ORDER = {"HIGH": 0, "MID": 1, "LOW": 2}
PRIORITY_STYLE = {
    "HIGH": ("🔴 HIGH", "#fee2e2", "#dc2626"),
    "MID": ("🟡 MID", "#fef3c7", "#ca8a04"),
    "LOW": ("⚪ LOW", "#f1f5f9", "#64748b"),
}
PRIORITY_COLOR = {"HIGH": "#dc2626", "MID": "#d97706", "LOW": "#94a3b8"}

CRAWL_QUERIES = [
    "골프 브랜드 신제품 2026",
    "골프 장비 용품 트렌드",
    "KLPGA KPGA 대회 결과",
    "국내 골프장 현황 2026",
    "골프 소비 트렌드 2026",
]

# ══════════════════════════════════════════════════════════════
# 유틸
# ══════════════════════════════════════════════════════════════

def get_api_key() -> str | None:
    """환경변수 또는 Streamlit secrets 에서 Gemini API 키 조회."""
    api_key = os.getenv("GEMINI_API_KEY")
    if api_key:
        return api_key

    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
        if api_key:
            return api_key
    except Exception:
        pass

    return None


def ensure_api_config():
    """앱 시작 시 API 키 누락을 조기 감지."""
    api_key = get_api_key()
    if not api_key:
        st.error(
            "Gemini API 키가 설정되지 않았습니다.\n\n"
            "로컬 실행:\n"
            "```bash\n"
            "export GEMINI_API_KEY='AIza...'\n"
            "streamlit run app.py\n"
            "```\n\n"
            "Streamlit Cloud:\n"
            "`Settings > Secrets`에 아래를 추가하세요.\n"
            "```toml\n"
            "GEMINI_API_KEY = 'AIza...'\n"
            "```"
        )
        st.stop()


def safe_int(value: Any, default: int = 0) -> int:
    """엑셀 NaN/문자열/실수 방어용 안전 정수 변환."""
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass

    if value is None:
        return default

    try:
        return int(float(str(value).replace(",", "").strip()))
    except Exception:
        return default


def safe_str(value: Any, default: str = "") -> str:
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    return str(value).strip()


def get_client():
    api_key = get_api_key()
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY가 설정되지 않았습니다.")
    return genai.Client(api_key=api_key)


def _extract_text_from_response(response) -> str:
    """Gemini 응답에서 텍스트를 최대한 안전하게 추출."""
    # 가장 쉬운 경로
    text = getattr(response, "text", None)
    if text:
        return text.strip()

    # 후보 파싱 시도
    candidates = getattr(response, "candidates", None)
    if not candidates:
        raise ValueError("Gemini 응답에서 candidates를 찾을 수 없습니다.")

    parts_text = []
    for cand in candidates:
        content = getattr(cand, "content", None)
        if not content:
            continue
        parts = getattr(content, "parts", None) or []
        for part in parts:
            part_text = getattr(part, "text", None)
            if part_text:
                parts_text.append(part_text)

    text = "\n".join(parts_text).strip()
    if not text:
        raise ValueError("Gemini 응답에서 텍스트를 찾을 수 없습니다.")
    return text


def _parse_json(text: str):
    """모델 응답에서 JSON만 최대한 안전하게 추출."""
    text = text.strip()

    fence_match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.DOTALL)
    if fence_match:
        text = fence_match.group(1).strip()

    if not (text.startswith("{") or text.startswith("[")):
        obj_match = re.search(r"(\{.*\}|\[.*\])", text, re.DOTALL)
        if obj_match:
            text = obj_match.group(1).strip()

    return json.loads(text)


def ask_gemini_json(prompt: str):
    """Gemini에 프롬프트를 보내 JSON 응답을 파싱해서 반환."""
    client = get_client()

    strict_prompt = (
        prompt
        + "\n\n중요: 설명, 서론, 코드블록 없이 JSON만 반환하세요."
    )

    response = client.models.generate_content(
        model=MODEL,
        contents=strict_prompt,
    )

    text = _extract_text_from_response(response)
    return _parse_json(text)


def show_ai_error(feature_name: str, err: Exception):
    st.warning(f"{feature_name} 중 오류가 발생했습니다: {err}")


# ══════════════════════════════════════════════════════════════
# 데이터 로더
# ══════════════════════════════════════════════════════════════

@st.cache_data(ttl=60)
def load_news() -> dict:
    """weekly_news.json 로드. 없으면 빈 구조 반환."""
    if NEWS_FILE.exists():
        try:
            with open(NEWS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass

    return {
        "period": "데이터 없음",
        "골프 브랜드": [],
        "골프 경기": [],
        "골프장 현황": [],
        "기타 이슈": [],
    }


@st.cache_data(ttl=60)
def load_store_profiles() -> list[dict]:
    """store_profiles.xlsx 로드. 없으면 하드코딩 기본값 사용."""
    if STORE_FILE.exists():
        try:
            df = pd.read_excel(STORE_FILE)
            stores = []
            for _, row in df.iterrows():
                stores.append(
                    {
                        "name": safe_str(row.get("점포명", "")),
                        "annualSales": f"{safe_int(row.get('연간골프매출(억)', 0))}억",
                        "customers": f"{safe_int(row.get('구매고객수(명)', 0)):,}명",
                        "avgTicket": f"{safe_int(row.get('평균객단가(만원)', 0))}만원",
                        "vipRatio": safe_int(row.get("우수고객구성비(%)", 0)),
                        "avgAge": safe_int(row.get("평균연령(세)", 0)),
                        "trait": safe_str(row.get("점포특성", "")),
                    }
                )
            stores = [s for s in stores if s["name"]]
            if stores:
                return stores
        except Exception as e:
            st.warning(f"점포 엑셀 로드 중 오류가 발생했습니다. 기본값을 사용합니다. ({e})")

    return [
        {"name": "잠실점", "annualSales": "350억", "customers": "20,000명", "avgTicket": "80만원", "vipRatio": 45, "avgAge": 56, "trait": "롯데월드몰 이용가능"},
        {"name": "본점", "annualSales": "250억", "customers": "15,000명", "avgTicket": "90만원", "vipRatio": 50, "avgAge": 47, "trait": "외국인 고객 많음"},
        {"name": "부산본점", "annualSales": "250억", "customers": "18,000명", "avgTicket": "80만원", "vipRatio": 40, "avgAge": 51, "trait": "바닷가"},
        {"name": "인천점", "annualSales": "200억", "customers": "12,000명", "avgTicket": "80만원", "vipRatio": 35, "avgAge": 43, "trait": "상권단독"},
        {"name": "동탄점", "annualSales": "150억", "customers": "10,000명", "avgTicket": "80만원", "vipRatio": 40, "avgAge": 39, "trait": "젊은점포"},
        {"name": "노원점", "annualSales": "130억", "customers": "8,000명", "avgTicket": "60만원", "vipRatio": 32, "avgAge": 54, "trait": "포켓상권"},
        {"name": "영등포점", "annualSales": "100억", "customers": "7,000명", "avgTicket": "40만원", "vipRatio": 30, "avgAge": 60, "trait": "경쟁열위"},
        {"name": "광복점", "annualSales": "80억", "customers": "6,000명", "avgTicket": "40만원", "vipRatio": 28, "avgAge": 55, "trait": "점포는큼"},
        {"name": "미아점", "annualSales": "40억", "customers": "4,000명", "avgTicket": "30만원", "vipRatio": 21, "avgAge": 65, "trait": "행사효율 좋음"},
        {"name": "건대점", "annualSales": "40억", "customers": "4,000명", "avgTicket": "20만원", "vipRatio": 15, "avgAge": 43, "trait": "골프 부진"},
    ]


def save_news(data: dict):
    """뉴스 데이터를 weekly_news.json에 저장."""
    with open(NEWS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    load_news.clear()


# ══════════════════════════════════════════════════════════════
# 뉴스 크롤러
# ══════════════════════════════════════════════════════════════

def crawl_naver_news(query: str, max_items: int = 3) -> list[dict]:
    """네이버 뉴스 검색 결과 크롤링."""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    url = f"https://search.naver.com/search.naver?where=news&query={requests.utils.quote(query)}&sort=1"

    try:
        resp = requests.get(url, headers=headers, timeout=8)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        items = []

        for news_item in soup.select("ul.list_news > li.bx")[:max_items]:
            title_el = news_item.select_one("a.news_tit")
            press_el = news_item.select_one("a.info.press")
            desc_el = news_item.select_one("a.dsc_txt_wrap") or news_item.select_one("div.news_dsc")

            if not title_el:
                continue

            items.append(
                {
                    "title": title_el.get_text(strip=True),
                    "url": title_el.get("href", ""),
                    "source": press_el.get_text(strip=True) if press_el else "네이버뉴스",
                    "desc": desc_el.get_text(strip=True) if desc_el else "",
                }
            )
        return items

    except Exception:
        return []


def crawl_all_news() -> list[dict]:
    """전체 쿼리 크롤링 후 중복 제거."""
    all_items = []
    seen_titles = set()

    for query in CRAWL_QUERIES:
        for item in crawl_naver_news(query, max_items=3):
            key = item["title"][:30]
            if key not in seen_titles:
                seen_titles.add(key)
                all_items.append(item)
        time.sleep(0.5)

    return all_items


# ══════════════════════════════════════════════════════════════
# Gemini API
# ══════════════════════════════════════════════════════════════

def analyze_news_card(title: str, raw_text: str, category: str) -> dict:
    prompt = f"""당신은 롯데백화점 골프 MD(상품기획자)의 주간 리포트 작성을 돕는 전문 어시스턴트입니다.

카테고리: {category}
제목: {title}

뉴스 본문:
{raw_text}

반드시 아래 JSON 형식만 반환하세요.

{{
  "priority": "HIGH 또는 MID 또는 LOW",
  "oneLine": "핵심을 20자 이내로 압축한 한 줄 요약",
  "summary": "MD가 알아야 할 핵심 내용을 2~3문장으로 요약. 수치·브랜드명 포함.",
  "insight": "롯데백화점 골프 매장 운영·발주·프로모션에 미치는 함의를 1~2문장으로.",
  "actions": ["구체적 실행 액션 1", "구체적 실행 액션 2", "구체적 실행 액션 3"]
}}

priority 기준: HIGH=즉각 매출·발주 영향 / MID=중기 영향 / LOW=참고용"""

    result = ask_gemini_json(prompt)
    if not isinstance(result, dict):
        raise ValueError("뉴스 분석 응답 형식이 올바르지 않습니다.")
    return result


def classify_and_analyze_crawled(items: list[dict]) -> list[dict]:
    """크롤링된 뉴스 목록을 카테고리 분류 + 개별 분석."""
    batch_text = "\n".join(
        f"{i+1}. 제목: {it['title']}\n   본문: {it['desc'][:200]}"
        for i, it in enumerate(items)
    )

    prompt = f"""당신은 롯데백화점 골프 MD 어시스턴트입니다.
아래 골프 뉴스들을 카테고리로 분류하고 각각 MD 관점으로 분석해 JSON을 반환하세요.

카테고리: 골프 브랜드 / 골프 경기 / 골프장 현황 / 기타 이슈

뉴스 목록:
{batch_text}

반드시 아래 JSON 배열만 반환하세요.

[
  {{
    "index": 1,
    "category": "카테고리명",
    "priority": "HIGH 또는 MID 또는 LOW",
    "oneLine": "20자 이내 한 줄 요약",
    "summary": "2~3문장 요약. 수치·브랜드명 포함.",
    "insight": "MD 관점 함의 1~2문장.",
    "actions": ["액션1", "액션2"]
  }}
]"""

    result = ask_gemini_json(prompt)
    if not isinstance(result, list):
        raise ValueError("뉴스 일괄 분석 응답 형식이 올바르지 않습니다.")
    return result


def generate_store_insight(store: dict, news_items: list) -> list:
    news_summary = "\n".join(
        f"- [{it.get('priority', 'MID')}] {it.get('title', '')}: {it.get('oneLine', '')}"
        for it in news_items
    )

    prompt = f"""당신은 롯데백화점 골프 MD 전문 컨설턴트입니다.

## 점포 정보
- 점포명: {store['name']}
- 연간 골프 매출: {store['annualSales']}
- 구매 고객수: {store['customers']}
- 평균 객단가: {store['avgTicket']}
- 우수고객 구성비: {store['vipRatio']}%
- 고객 평균 연령: {store['avgAge']}세
- 점포 특성: {store['trait']}

## 이번 주 주요 뉴스
{news_summary}

반드시 아래 JSON 배열만 반환하세요.

[
  {{
    "title": "실행안 제목 (점포명 포함, 20자 이내)",
    "idea": "구체적 실행 방법 2~3문장.",
    "reason": "점포 데이터(수치 인용)와 뉴스 트렌드를 연결한 근거 2문장.",
    "score": 7
  }}
]

작성 원칙: 점포 특성 반영 / 뉴스 → 실행 연결 / 실현 가능한 백화점 MD 행사 수준"""

    results = ask_gemini_json(prompt)
    if not isinstance(results, list):
        raise ValueError("점포 인사이트 응답 형식이 올바르지 않습니다.")

    results = [r for r in results if isinstance(r, dict)]
    results.sort(key=lambda x: x.get("score", 0), reverse=True)
    return results[:3]


def extract_weekly_keywords(all_news: list) -> list:
    if not all_news:
        return []

    news_text = "\n".join(
        f"{it.get('title', '')} / {it.get('oneLine', '')} / {it.get('summary', '')}"
        for it in all_news
    )

    prompt = f"""아래 골프 업계 주간 뉴스에서 MD가 주목해야 할 핵심 키워드 Top 10을 추출하세요.

{news_text}

반드시 아래 JSON 배열만 반환하세요.

[{{"keyword": "키워드", "count": 3, "trend": "up"}}]

count: 1~5 정수 / trend: "up" "same" "down" 중 하나"""

    result = ask_gemini_json(prompt)
    if not isinstance(result, list):
        raise ValueError("키워드 추출 응답 형식이 올바르지 않습니다.")
    return result


# ══════════════════════════════════════════════════════════════
# 채택 리포트 생성
# ══════════════════════════════════════════════════════════════

def build_report_markdown(selected_cards: list, period: str) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    lines = [
        "# 롯데 석주 주간 골프 MD 리포트",
        f"**기간**: {period}  |  **생성일**: {now}  |  **채택 뉴스**: {len(selected_cards)}건",
        "",
        "---",
        "",
    ]

    for i, card in enumerate(selected_cards, 1):
        lines += [
            f"## {i}. {card['title']}",
            f"**날짜**: {card['date']}  |  **출처**: {card['source']}  |  **우선순위**: {card['priority']}",
            "",
            f"**핵심 요약**: {card['oneLine']}",
            "",
            card["summary"],
            "",
            f"**MD 인사이트**: {card['insight']}",
            "",
            f"**실행 액션**: {' / '.join(card['actions'])}",
            "",
            "---",
            "",
        ]

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════
# Rule-based 폴백 인사이트
# ══════════════════════════════════════════════════════════════

def build_store_insights_fallback(store, selected_cards, data) -> list:
    source = selected_cards if selected_cards else [
        it for items in data.values() for it in items if isinstance(items, list)
    ]

    ideas = []
    seen = set()

    for item in source:
        title = f'{store["name"]} 맞춤 기획전'
        if title not in seen:
            seen.add(title)
            ideas.append(
                {
                    "title": title,
                    "idea": f'{item.get("title", "주요 뉴스")} 트렌드를 반영해 {store["name"]}에서 시즌 기획전을 운영합니다.',
                    "reason": f'{store["name"]}의 특성({store["trait"]})과 평균 연령 {store["avgAge"]}세, 객단가 {store["avgTicket"]}를 고려한 제안입니다.',
                    "score": store["vipRatio"],
                }
            )
        if len(ideas) == 3:
            break

    return ideas


# ══════════════════════════════════════════════════════════════
# API 설정 확인
# ══════════════════════════════════════════════════════════════

ensure_api_config()

# ══════════════════════════════════════════════════════════════
# Session State 초기화
# ══════════════════════════════════════════════════════════════

defaults = {
    "page": "news",
    "category": "골프 브랜드",
    "selected_cards": [],
    "expanded_id": None,
    "ai_results": {},
    "keywords": [],
    "store_ai_cache": {},
    "crawl_log": "",
}

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── 데이터 로드 ────────────────────────────────────────────────
raw_data = load_news()
PERIOD = raw_data.get("period", "")
DATA = {k: v for k, v in raw_data.items() if k != "period"}
STORES = load_store_profiles()

if not DATA:
    DATA = {
        "골프 브랜드": [],
        "골프 경기": [],
        "골프장 현황": [],
        "기타 이슈": [],
    }

# ══════════════════════════════════════════════════════════════
# 전역 스타일
# ══════════════════════════════════════════════════════════════

st.markdown("""
<style>
.block-container{max-width:1280px;padding-top:1rem;padding-bottom:3rem}
html,body,[data-testid="stAppViewContainer"]{background:#f0fdf4}
#MainMenu{visibility:hidden!important}
header{visibility:hidden!important;height:0!important}
[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"]{display:none!important}
.stApp>header{display:none!important}
footer{visibility:hidden!important}

.header-bar{display:flex;align-items:center;justify-content:space-between;padding:18px 26px;margin-bottom:14px;background:linear-gradient(135deg,#064e3b,#065f46);border-radius:20px;box-shadow:0 4px 24px rgba(6,78,59,.28)}
.header-left{display:flex;align-items:center;gap:14px}
.header-icon{font-size:38px;line-height:1}
.header-title{font-size:26px;font-weight:800;color:white;letter-spacing:-.3px}
.header-sub{font-size:16px;color:#6ee7b7;margin-top:3px;font-weight:500}
.header-right{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.period-badge{font-size:15px;font-weight:600;color:#a7f3d0;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);border-radius:20px;padding:6px 16px}
.adopted-badge{font-size:15px;font-weight:700;color:#064e3b;background:#d1fae5;border:1px solid #6ee7b7;border-radius:20px;padding:6px 16px}
.adopted-badge-none{font-size:15px;color:#a7f3d0;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);border-radius:20px;padding:6px 16px}

.active-bar{height:3px;background:#059669;border-radius:2px;margin-top:-6px;margin-bottom:10px}
.inactive-bar{height:3px;margin-top:-6px;margin-bottom:10px}

.card-wrap{border:1px solid #d1fae5;border-radius:18px;background:white;margin-bottom:16px;box-shadow:0 2px 12px rgba(6,78,59,.07);overflow:hidden}
.card-top-bar{height:4px;width:100%}
.card-body{padding:20px 22px 16px}
.card-meta{display:flex;justify-content:space-between;align-items:center}
.muted{font-size:15px;color:#94a3b8;font-weight:500}
.card-title{margin-top:8px;font-size:22px;font-weight:800;color:#064e3b;line-height:1.4}
.badge{display:inline-block;font-size:14px;font-weight:700;border-radius:8px;padding:5px 12px}
.badge-HIGH{background:#fee2e2;color:#dc2626}
.badge-MID{background:#fef3c7;color:#d97706}
.badge-LOW{background:#f1f5f9;color:#64748b}
.one-line{margin-top:10px;font-size:16px;font-weight:700;color:#065f46;display:flex;align-items:center;gap:8px}
.one-line-bar{display:inline-block;width:4px;height:18px;background:#10b981;border-radius:2px;flex-shrink:0}
.summary{margin-top:8px;font-size:16px;color:#374151;line-height:1.9}
.box{margin-top:14px;border:1px solid #d1fae5;background:#f0fdf4;border-radius:12px;padding:12px 14px}
.small-title{font-size:13px;font-weight:800;color:#6ee7b7;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px}
.link-chip{display:inline-flex;align-items:center;gap:4px;font-size:15px;color:#065f46;background:white;border:1px solid #a7f3d0;border-radius:8px;padding:6px 12px;margin:0 5px 5px 0;text-decoration:none}
.link-chip:hover{background:#ecfdf5;border-color:#6ee7b7}
.insight-section{margin-top:14px;padding:18px 20px;background:linear-gradient(135deg,#f0fdf4,#ecfdf5);border:1px solid #a7f3d0;border-radius:14px}
.insight-label{font-size:13px;font-weight:800;color:#6ee7b7;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px}
.insight-text{font-size:16px;color:#1f2937;line-height:1.85}
.action-tag{display:inline-block;font-size:15px;color:#065f46;background:#d1fae5;border:1px solid #6ee7b7;border-radius:8px;padding:6px 14px;margin:3px 4px 3px 0;font-weight:600}
.ai-badge-inline{font-size:11px;background:#dbeafe;color:#1d4ed8;border-radius:5px;padding:2px 8px;font-weight:600;margin-left:6px}

.store-header{background:linear-gradient(135deg,#064e3b,#065f46);border-radius:16px;padding:20px 22px;margin-bottom:14px;color:white}
.store-name{font-size:26px;font-weight:800}
.store-sub{font-size:16px;color:#6ee7b7;margin-top:4px;font-weight:500}
.store-trait{display:inline-block;font-size:15px;font-weight:600;color:#a7f3d0;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);border-radius:20px;padding:5px 14px;margin-top:10px}
.stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.stat-card{border:1px solid #d1fae5;border-radius:12px;padding:14px 16px;background:white;box-shadow:0 1px 3px rgba(6,78,59,.06)}
.stat-label{font-size:13px;color:#059669;font-weight:700;text-transform:uppercase;letter-spacing:.5px}
.stat-value{margin-top:5px;font-weight:800;font-size:20px;color:#064e3b}
.prog-row{margin-bottom:12px}
.prog-header{display:flex;justify-content:space-between;margin-bottom:5px}
.prog-label{font-size:13px;font-weight:700;color:#059669;text-transform:uppercase;letter-spacing:.5px}
.prog-val{font-size:16px;font-weight:800;color:#064e3b}
.prog-bg{background:#d1fae5;border-radius:99px;height:8px;overflow:hidden}
.prog-fill-blue{height:100%;border-radius:99px;background:linear-gradient(90deg,#059669,#34d399)}
.prog-fill-amber{height:100%;border-radius:99px;background:linear-gradient(90deg,#d97706,#fbbf24)}

.idea-box{border-radius:16px;background:white;border:1px solid #d1fae5;margin-top:14px;box-shadow:0 2px 10px rgba(6,78,59,.07);overflow:hidden}
.idea-box-head{display:flex;align-items:center;gap:14px;padding:14px 18px;background:#f0fdf4;border-bottom:1px solid #d1fae5}
.idea-num-circle{width:38px;height:38px;border-radius:50%;flex-shrink:0;background:linear-gradient(135deg,#065f46,#10b981);color:white;font-size:17px;font-weight:800;display:flex;align-items:center;justify-content:center}
.idea-title{font-weight:800;font-size:17px;color:#064e3b}
.idea-box-body{padding:18px 20px}
.idea-body{font-size:16px;color:#1f2937;line-height:1.9}
.idea-sep{height:1px;background:#d1fae5;margin:12px 0}
.idea-reason-label{font-size:13px;font-weight:800;color:#6ee7b7;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px}
.idea-reason{font-size:16px;color:#374151;line-height:1.85}

.compare-table{width:100%;border-collapse:collapse;font-size:14px}
.compare-table th{background:#064e3b;color:white;padding:10px 12px;text-align:center;font-weight:700;white-space:nowrap}
.compare-table td{padding:9px 12px;text-align:center;border-bottom:1px solid #d1fae5;color:#1f2937}
.compare-table tr:nth-child(even) td{background:#f0fdf4}
.compare-table tr:hover td{background:#ecfdf5}
.bar-cell{display:flex;align-items:center;gap:6px;justify-content:center}
.mini-bar-bg{background:#d1fae5;border-radius:4px;height:6px;width:80px;overflow:hidden;flex-shrink:0}
.mini-bar-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,#059669,#34d399)}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 헤더
# ══════════════════════════════════════════════════════════════

adopted_count = len(st.session_state.selected_cards)
adopted_html = (
    f'<span class="adopted-badge">✔ 채택 {adopted_count}건</span>'
    if adopted_count > 0
    else '<span class="adopted-badge-none">채택 없음</span>'
)

st.markdown(f"""
<div class="header-bar">
  <div class="header-left">
    <span class="header-icon">⛳</span>
    <div>
      <div class="header-title">롯데 석주 주간 골프 MD 뉴스</div>
      <div class="header-sub">골프 MD 주간 트렌드 및 점포 실행안 리포트</div>
    </div>
  </div>
  <div class="header-right">
    <span class="period-badge">📅 {PERIOD}</span>
    {adopted_html}
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 상단 네비게이션
# ══════════════════════════════════════════════════════════════

nav1, nav2, nav3, _ = st.columns([1.8, 2.2, 2.0, 5])

with nav1:
    if st.button("📰 뉴스 분석", use_container_width=True, key="nav_news"):
        st.session_state.page = "news"
    st.markdown(
        '<div class="active-bar"></div>' if st.session_state.page == "news" else '<div class="inactive-bar"></div>',
        unsafe_allow_html=True,
    )

with nav2:
    if st.button("🏪 점포 인사이트", use_container_width=True, key="nav_insight"):
        st.session_state.page = "insight"
    st.markdown(
        '<div class="active-bar"></div>' if st.session_state.page == "insight" else '<div class="inactive-bar"></div>',
        unsafe_allow_html=True,
    )

with nav3:
    if st.button("📊 점포 비교", use_container_width=True, key="nav_compare"):
        st.session_state.page = "compare"
    st.markdown(
        '<div class="active-bar"></div>' if st.session_state.page == "compare" else '<div class="inactive-bar"></div>',
        unsafe_allow_html=True,
    )

# ══════════════════════════════════════════════════════════════
# 뉴스 분석 페이지
# ══════════════════════════════════════════════════════════════

if st.session_state.page == "news":
    top_c1, top_c2, top_c3, top_c4 = st.columns([2, 2, 2, 2])

    with top_c1:
        if st.button("🔄 이번 주 뉴스 수집", use_container_width=True, key="btn_crawl"):
            with st.spinner("네이버 골프 뉴스 수집 중..."):
                raw_items = crawl_all_news()

            if not raw_items:
                st.warning("뉴스 수집 결과가 없습니다. 잠시 후 다시 시도해주세요.")
            else:
                try:
                    with st.spinner(f"{len(raw_items)}건 Gemini 분석 중..."):
                        analyzed = classify_and_analyze_crawled(raw_items)

                    today = datetime.now().strftime("%m.%d")
                    new_data = {
                        "period": f"{today} 수집",
                        "골프 브랜드": [],
                        "골프 경기": [],
                        "골프장 현황": [],
                        "기타 이슈": [],
                    }

                    for i, result in enumerate(analyzed):
                        if i >= len(raw_items):
                            break

                        src = raw_items[i]
                        cat = result.get("category", "기타 이슈")
                        if cat not in new_data:
                            cat = "기타 이슈"

                        new_data[cat].append(
                            {
                                "id": f"crawled-{i}",
                                "date": today,
                                "source": src["source"],
                                "title": src["title"],
                                "priority": result.get("priority", "MID"),
                                "oneLine": result.get("oneLine", ""),
                                "summary": result.get("summary", src["desc"][:150]),
                                "headlines": [{"title": src["title"], "url": src["url"]}],
                                "insight": result.get("insight", ""),
                                "actions": result.get("actions", []),
                            }
                        )

                    save_news(new_data)
                    st.session_state.crawl_log = f"✅ {len(raw_items)}건 수집 완료 ({datetime.now().strftime('%H:%M')})"
                    st.rerun()

                except Exception as e:
                    show_ai_error("뉴스 일괄 분석", e)

    with top_c2:
        if st.button("✦ 키워드 추출", use_container_width=True, key="btn_keywords"):
            flat = [it for items in DATA.values() for it in items if isinstance(items, list)]
            try:
                with st.spinner("키워드 분석 중..."):
                    st.session_state.keywords = extract_weekly_keywords(flat)
                st.rerun()
            except Exception as e:
                show_ai_error("키워드 추출", e)

    with top_c3:
        if st.session_state.selected_cards:
            report_md = build_report_markdown(st.session_state.selected_cards, PERIOD)
            st.download_button(
                label="📥 채택 리포트 다운로드",
                data=report_md.encode("utf-8"),
                file_name=f"golf_md_report_{datetime.now().strftime('%Y%m%d')}.md",
                mime="text/markdown",
                use_container_width=True,
                key="btn_download",
            )
        else:
            st.button("📥 채택 리포트 다운로드", use_container_width=True, disabled=True, key="btn_download_dis")

    with top_c4:
        if st.session_state.selected_cards:
            if st.button("🗑 채택 초기화", use_container_width=True, key="btn_reset"):
                st.session_state.selected_cards = []
                st.rerun()

    if st.session_state.crawl_log:
        st.success(st.session_state.crawl_log)

    if st.session_state.keywords:
        TREND_ICON = {"up": "↑", "same": "→", "down": "↓"}
        TREND_COLOR = {"up": "#16a34a", "same": "#6b7280", "down": "#dc2626"}

        badges = ""
        for kw in st.session_state.keywords[:10]:
            icon = TREND_ICON.get(kw.get("trend", "same"), "→")
            color = TREND_COLOR.get(kw.get("trend", "same"), "#6b7280")
            size = 13 + kw.get("count", 1)
            badges += (
                f'<span style="display:inline-block;margin:3px 4px;padding:5px 12px;'
                f'border-radius:20px;background:#f0fdf4;border:1px solid #a7f3d0;'
                f'font-size:{size}px;font-weight:600;color:#064e3b;">'
                f'{kw.get("keyword","")} <span style="color:{color};font-size:11px;">{icon}</span></span>'
            )

        st.markdown(
            f'<div style="margin-bottom:14px;"><span style="font-size:13px;font-weight:700;'
            f'color:#059669;margin-right:8px;">이번 주 키워드</span>{badges}</div>',
            unsafe_allow_html=True,
        )

    cat_cols = st.columns(4)
    for i, cat in enumerate(DATA.keys()):
        with cat_cols[i]:
            icon = CATEGORY_ICON.get(cat, "")
            count = len(DATA.get(cat, []))
            if st.button(f"{icon} {cat} ({count})", use_container_width=True, key=f"cat_{cat}"):
                st.session_state.category = cat
                st.session_state.expanded_id = None

            st.markdown(
                '<div class="active-bar"></div>' if st.session_state.category == cat else '<div class="inactive-bar"></div>',
                unsafe_allow_html=True,
            )

    if st.session_state.category not in DATA:
        st.session_state.category = list(DATA.keys())[0]

    cards = sorted(
        DATA.get(st.session_state.category, []),
        key=lambda x: PRIORITY_ORDER.get(x.get("priority", "LOW"), 99),
    )

    if not cards:
        st.info("이 카테고리에 뉴스가 없습니다. 상단 '이번 주 뉴스 수집' 버튼을 눌러주세요.")

    for card in cards:
        ai = st.session_state.ai_results.get(card["id"])
        dp = ai["priority"] if ai else card["priority"]
        dl = ai["oneLine"] if ai else card["oneLine"]
        ds = ai["summary"] if ai else card["summary"]
        di = ai["insight"] if ai else card["insight"]
        da = ai["actions"] if ai else card["actions"]

        border_color = PRIORITY_COLOR.get(dp, "#94a3b8")
        ai_badge = '<span class="ai-badge-inline">✦ AI</span>' if ai else ""
        chips_html = "".join(
            f'<a class="link-chip" href="{h["url"]}" target="_blank" rel="noreferrer">↗ {h["title"]}</a>'
            for h in card.get("headlines", [])
        )

        st.markdown(f"""
<div class="card-wrap">
  <div class="card-top-bar" style="background:{border_color};"></div>
  <div class="card-body">
    <div class="card-meta">
      <span class="muted">{card["date"]} · {card["source"]}{ai_badge}</span>
      <span class="badge badge-{dp}">{PRIORITY_STYLE.get(dp, ('','',''))[0]}</span>
    </div>
    <div class="card-title">{card["title"]}</div>
    <div class="one-line"><span class="one-line-bar"></span>{dl}</div>
    <div class="summary">{ds}</div>
    <div class="box"><div class="small-title">관련 기사</div>{chips_html}</div>
  </div>
""", unsafe_allow_html=True)

        b1, b2, b3, _ = st.columns([1.4, 1.4, 1.4, 4])

        with b1:
            if st.button("✦ 재분석" if ai else "✦ AI 분석", key=f"ai_{card['id']}", use_container_width=True):
                try:
                    with st.spinner("Gemini 분석 중..."):
                        result = analyze_news_card(
                            card["title"],
                            f"{card.get('summary', '')} {card.get('oneLine', '')}",
                            st.session_state.category,
                        )
                    st.session_state.ai_results[card["id"]] = result
                    st.rerun()
                except Exception as e:
                    show_ai_error("뉴스 카드 재분석", e)

        with b2:
            exp_label = "인사이트 ▲" if st.session_state.expanded_id == card["id"] else "인사이트 ▼"
            if st.button(exp_label, key=f"expand_{card['id']}", use_container_width=True):
                st.session_state.expanded_id = None if st.session_state.expanded_id == card["id"] else card["id"]

        if st.session_state.expanded_id == card["id"]:
            action_tags = "".join(f'<span class="action-tag">{a}</span>' for a in da)

            st.markdown(f"""
<div class="insight-section">
  <div style="display:flex;gap:24px;flex-wrap:wrap;">
    <div style="flex:1;min-width:220px;">
      <div class="insight-label">인사이트</div>
      <div class="insight-text">{di}</div>
    </div>
    <div style="flex:1;min-width:180px;">
      <div class="insight-label">실행 액션</div>
      <div style="margin-top:6px;">{action_tags}</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

            with b3:
                already = any(it["id"] == card["id"] for it in st.session_state.selected_cards)
                if st.button(
                    "✔ 채택됨" if already else "실행안 채택",
                    key=f"select_{card['id']}",
                    disabled=already,
                    use_container_width=True,
                ):
                    merged = {**card}
                    if ai:
                        merged.update(
                            {
                                "priority": dp,
                                "oneLine": dl,
                                "summary": ds,
                                "insight": di,
                                "actions": da,
                            }
                        )
                    st.session_state.selected_cards.append(merged)
                    st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 점포 인사이트 페이지
# ══════════════════════════════════════════════════════════════

elif st.session_state.page == "insight":
    selected_name = st.selectbox("점포 선택", [s["name"] for s in STORES], index=0)
    store = next(s for s in STORES if s["name"] == selected_name)

    st.markdown(f"""
<div class="store-header">
  <div class="store-name">{store["name"]}</div>
  <div class="store-sub">연간 골프 매출 {store["annualSales"]} · 구매고객 {store["customers"]} · 객단가 {store["avgTicket"]}</div>
  <span class="store-trait">📍 {store["trait"]}</span>
</div>
""", unsafe_allow_html=True)

    left, right = st.columns([1.1, 2.2])

    with left:
        age_pct = min(store["avgAge"], 70) / 70 * 100
        st.markdown(f"""
<div style="font-size:12px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;">점포 지표</div>
<div class="stat-grid">
  <div class="stat-card"><div class="stat-label">연간 골프 매출</div><div class="stat-value">{store["annualSales"]}</div></div>
  <div class="stat-card"><div class="stat-label">구매고객수</div><div class="stat-value">{store["customers"]}</div></div>
  <div class="stat-card"><div class="stat-label">객단가</div><div class="stat-value">{store["avgTicket"]}</div></div>
  <div class="stat-card"><div class="stat-label">평균 연령</div><div class="stat-value">{store["avgAge"]}세</div></div>
</div>
<div style="margin-top:14px;">
  <div class="prog-row">
    <div class="prog-header"><span class="prog-label">우수고객 구성비</span><span class="prog-val">{store["vipRatio"]}%</span></div>
    <div class="prog-bg"><div class="prog-fill-blue" style="width:{store['vipRatio']}%;"></div></div>
  </div>
  <div class="prog-row">
    <div class="prog-header"><span class="prog-label">평균 연령</span><span class="prog-val">{store["avgAge"]}세</span></div>
    <div class="prog-bg"><div class="prog-fill-amber" style="width:{age_pct:.0f}%;"></div></div>
  </div>
</div>
""", unsafe_allow_html=True)

    with right:
        cache_key = f"{selected_name}_{'_'.join(c['id'] for c in st.session_state.selected_cards)}"
        adopted_info = f" — 채택 뉴스 {adopted_count}건 기반" if adopted_count > 0 else " — 전체 뉴스 기반"

        title_col, ai_col = st.columns([3, 1.4])

        with title_col:
            st.markdown(
                f'<div style="font-size:15px;font-weight:800;color:#0f172a;margin-bottom:8px;">점포 맞춤 제안'
                f'<span style="font-size:13px;color:#64748b;font-weight:400;">{adopted_info}</span></div>',
                unsafe_allow_html=True,
            )

        with ai_col:
            if st.button("✦ AI 생성", key="btn_store_ai", use_container_width=True):
                source_items = (
                    st.session_state.selected_cards
                    if st.session_state.selected_cards
                    else [it for items in DATA.values() for it in items if isinstance(items, list)]
                )

                try:
                    with st.spinner(f"{store['name']} 맞춤 인사이트 생성 중..."):
                        insights = generate_store_insight(store, source_items)
                    st.session_state.store_ai_cache[cache_key] = insights
                    st.rerun()
                except Exception as e:
                    show_ai_error("점포 인사이트 생성", e)

        insights = st.session_state.store_ai_cache.get(cache_key)
        ai_generated = insights is not None
        if insights is None:
            insights = build_store_insights_fallback(store, st.session_state.selected_cards, DATA)

        for idx, idea in enumerate(insights, 1):
            ai_mark = (
                '<span style="font-size:11px;background:#dbeafe;color:#1d4ed8;border-radius:5px;'
                'padding:2px 7px;font-weight:600;margin-left:6px;">AI</span>'
                if ai_generated else ""
            )

            st.markdown(f"""
<div class="idea-box">
  <div class="idea-box-head">
    <div class="idea-num-circle">{idx}</div>
    <div class="idea-title">{idea["title"]}{ai_mark}</div>
  </div>
  <div class="idea-box-body">
    <div class="idea-body">{idea["idea"]}</div>
    <div class="idea-sep"></div>
    <div class="idea-reason-label">근거</div>
    <div class="idea-reason">{idea["reason"]}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 점포 비교 페이지
# ══════════════════════════════════════════════════════════════

elif st.session_state.page == "compare":
    st.markdown(
        '<div style="font-size:20px;font-weight:800;color:#064e3b;margin-bottom:16px;">전체 점포 비교</div>',
        unsafe_allow_html=True,
    )

    sort_options = {
        "연간 매출 ↓": ("annualSales_raw", True),
        "객단가 ↓": ("avgTicket_raw", True),
        "우수고객 비율 ↓": ("vipRatio", True),
        "평균 연령 ↓": ("avgAge", True),
        "평균 연령 ↑": ("avgAge", False),
    }

    sort_key = st.selectbox("정렬 기준", list(sort_options.keys()), index=0, key="sort_compare")
    sort_col, sort_desc = sort_options[sort_key]

    stores_for_sort = []
    for s in STORES:
        s2 = dict(s)
        s2["annualSales_raw"] = int(re.sub(r"[^\d]", "", s["annualSales"]) or 0)
        s2["avgTicket_raw"] = int(re.sub(r"[^\d]", "", s["avgTicket"]) or 0)
        stores_for_sort.append(s2)

    stores_sorted = sorted(stores_for_sort, key=lambda x: x.get(sort_col, 0), reverse=sort_desc)

    max_sales = max(s["annualSales_raw"] for s in stores_sorted) or 1
    max_ticket = max(s["avgTicket_raw"] for s in stores_sorted) or 1
    max_vip = max(s["vipRatio"] for s in stores_sorted) or 1

    rows_html = ""
    for rank, s in enumerate(stores_sorted, 1):
        sales_pct = s["annualSales_raw"] / max_sales * 100
        ticket_pct = s["avgTicket_raw"] / max_ticket * 100
        vip_pct = s["vipRatio"] / max_vip * 100

        rows_html += f"""
<tr>
  <td style="font-weight:700;color:#064e3b;">{rank}</td>
  <td style="font-weight:800;color:#064e3b;">{s["name"]}</td>
  <td>
    <div class="bar-cell">
      <div class="mini-bar-bg"><div class="mini-bar-fill" style="width:{sales_pct:.0f}%;"></div></div>
      {s["annualSales"]}
    </div>
  </td>
  <td>{s["customers"]}</td>
  <td>
    <div class="bar-cell">
      <div class="mini-bar-bg"><div class="mini-bar-fill" style="width:{ticket_pct:.0f}%;background:linear-gradient(90deg,#d97706,#fbbf24);"></div></div>
      {s["avgTicket"]}
    </div>
  </td>
  <td>
    <div class="bar-cell">
      <div class="mini-bar-bg"><div class="mini-bar-fill" style="width:{vip_pct:.0f}%;background:linear-gradient(90deg,#7c3aed,#a78bfa);"></div></div>
      {s["vipRatio"]}%
    </div>
  </td>
  <td>{s["avgAge"]}세</td>
  <td><span style="font-size:13px;background:#f0fdf4;border:1px solid #a7f3d0;border-radius:6px;padding:3px 9px;color:#065f46;">{s["trait"]}</span></td>
</tr>
"""

    st.markdown(f"""
<table class="compare-table">
  <thead>
    <tr>
      <th>#</th><th>점포</th><th>연간 매출</th><th>구매고객수</th>
      <th>객단가</th><th>우수고객 비율</th><th>평균 연령</th><th>특성</th>
    </tr>
  </thead>
  <tbody>{rows_html}</tbody>
</table>
""", unsafe_allow_html=True)

    st.markdown(
        '<div style="margin-top:12px;font-size:13px;color:#94a3b8;">💡 점포명을 클릭하면 점포 인사이트 페이지로 이동합니다.</div>',
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("전체 연간 매출 합계", f"{sum(s['annualSales_raw'] for s in stores_sorted)}억")
    with col2:
        avg_vip = sum(s['vipRatio'] for s in stores_sorted) / len(stores_sorted)
        st.metric("평균 우수고객 비율", f"{avg_vip:.1f}%")
    with col3:
        avg_age = sum(s['avgAge'] for s in stores_sorted) / len(stores_sorted)
        st.metric("전체 평균 고객 연령", f"{avg_age:.1f}세")
