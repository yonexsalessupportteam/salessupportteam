import os
import json
import re
import time
import hashlib
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def parse_sheet_date(date_str):
    """구글 시트 '작성일' 컬럼 값을 date 객체로 파싱. 실패하면 None."""
    if not date_str:
        return None
    date_str = str(date_str).strip()
    if not date_str:
        return None
    # 구글 시트가 날짜를 시리얼넘버(숫자)로 줄 수도 있음
    try:
        serial = float(date_str)
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=serial)).date()
    except ValueError:
        pass
    for fmt in ('%Y-%m-%d', '%Y.%m.%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S', '%Y.%m.%d.'):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    return None

# CS 키워드 3단계 표 (정성 점수 CS코멘트, 기본 20점)
# CS 대리점평가 구글시트의 '키워드' 컬럼에서 담당자가 직접 선택한 값을 이 표와 매칭해 기계적으로 채점.
# 여러 등급이 섞여 있으면 가장 심각한 등급(고위험>중위험>저위험) 하나만 기준으로 감점.
# 키워드 목록은 추후 바뀔 수 있음 — 여기만 수정하면 전체 채점에 반영됨.
KEYWORD_TIERS = {
    '고위험': ['연락두절', '약속불이행', '클레임방치', '폐업징후', '허위접수', '타사이탈', '연락안됨', 'A/S 처리불량'],
    '중위험': ['연락지연', '가끔약속어김', '클레임/매출감소', '응대느림', '불만', 'A/S규정숙지'],
    '저위험': ['협조적', '응대원활', '약속이행', '클레임없음', '신뢰', '칭찬'],
}
KEYWORD_DEDUCT = {'고위험': 20, '중위험': 10, '저위험': 0}

# 매출규모 기준 (나중에 활성화)
SALES_GRADE = {
    '용품': {'상': 400_000_000, '중': 130_000_000},
    '의류': {'상': 190_000_000, '중':  50_000_000},
}
SALES_SCORE = {'상': 20, '중': 10, '하': 5}


def get_sheets_client():
    creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS', '')
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def fetch_sheet_data():
    spreadsheet_id = os.environ.get('SPREADSHEET_ID', '')
    client = get_sheets_client()
    return client.open_by_key(spreadsheet_id).sheet1.get_all_records()


def parse_month_header(header):
    """'26.04', '2026.04월', '2026-04' 등의 헤더에서 (year, month) 추출. 실패하면 None."""
    m = re.search(r'(\d{2,4})[.\-](\d{1,2})', str(header))
    if not m:
        return None
    year, month = int(m.group(1)), int(m.group(2))
    if year < 100:
        year += 2000
    if not (1 <= month <= 12):
        return None
    return (year, month)


def fetch_sales_tab(tab_name, name_col='매장명'):
    """'용품_3개월 매출' / '의류_3개월 매출' 같은 탭에서 최근 3개월 합계를 계산.
    탭 구조: 매장명 컬럼 + 월별 매출 컬럼(예: 26.04, 26.05, 26.06 ...).
    가장 최근 3개 월 컬럼을 자동 판별해 합산. 반환: {대리점명: 3개월합계}"""
    spreadsheet_id = os.environ.get('SPREADSHEET_ID', '')
    client = get_sheets_client()
    try:
        ws = client.open_by_key(spreadsheet_id).worksheet(tab_name)
    except Exception as e:
        print(f"'{tab_name}' 탭 읽기 실패: {e}")
        return {}

    values = ws.get_all_values()
    if not values:
        return {}
    header = values[0]

    try:
        name_idx = header.index(name_col)
    except ValueError:
        print(f"'{tab_name}' 탭에서 '{name_col}' 컬럼을 찾지 못함")
        return {}

    month_cols = []  # (year, month, col_idx)
    for idx, h in enumerate(header):
        parsed = parse_month_header(h)
        if parsed:
            month_cols.append((parsed[0], parsed[1], idx))
    if not month_cols:
        print(f"'{tab_name}' 탭에서 월별 컬럼을 찾지 못함")
        return {}

    # 최근 3개 월만 사용 (연/월 기준 정렬 후 마지막 3개)
    month_cols.sort(key=lambda x: (x[0], x[1]))
    recent_cols = month_cols[-3:]

    result = {}
    for row in values[1:]:
        if len(row) <= name_idx:
            continue
        name = ' '.join(str(row[name_idx]).split())
        if not name:
            continue
        total = 0
        for _, _, idx in recent_cols:
            if idx < len(row):
                total += parse_amount(row[idx])
        result[name] = total
    return result


# ───────────────────────────────────────────
# CS 점수 계산 (20점 만점)
# 메모 없음 → 20점 (이슈 없음으로 간주)
# 메모 있음 → AI(Gemini)가 메모 텍스트를 읽고 직접 판단한 위험단계(assessed_risk)를 기준으로 배점
#             (고정 키워드 리스트와의 문자열 일치가 아니라, 자유 서술 메모의 맥락을 보고 판단)
# ───────────────────────────────────────────
# ───────────────────────────────────────────
# CS 점수 계산 (20점 만점) — 구글시트 '키워드' 컬럼 값을 KEYWORD_TIERS와 매칭해 기계적으로 채점
# 키워드 없음 → 20점 / 저위험(긍정) → 20점 / 중위험 → 10점 / 고위험 → 0점
# ───────────────────────────────────────────
def score_cs_from_keywords(keyword_str):
    """반환: (점수, 매칭된 등급 또는 빈 문자열). 여러 등급이 섞여 있으면 가장 심각한 등급 하나만 기준으로 감점."""
    if not keyword_str or not keyword_str.strip():
        return 20, ''
    text = keyword_str
    for tier in ('고위험', '중위험', '저위험'):
        if any(kw in text for kw in KEYWORD_TIERS[tier]):
            return 20 - KEYWORD_DEDUCT[tier], tier
    return 20, ''  # 표에 없는 값 - 감점 없이 만점 처리


def score_partnership(p_goods, p_clothing):
    """파트너십 점수 (30점 만점, 용품 15 + 의류 15). 컬럼에 값이 있으면(공백 아니면) 위반으로 처리."""
    goods_score = 0 if str(p_goods).strip() else 15
    cloth_score = 0 if str(p_clothing).strip() else 15
    return goods_score + cloth_score


# ───────────────────────────────────────────
# 매출규모 감점 (3개월 매출 합계, 용품 10점 + 의류 10점 = 20점 만점)
# 구글시트 '매출3개월_용품' / '매출3개월_의류' 컬럼(각 사업부 3개월 합계 금액)을 기준으로
# 3구간 감점 방식 채점 (10점 만점에서 감점).
# 근거: 채권관리_배포자료(실측 3개월 합계, 114개 매장, 4~6월 기준)
# ───────────────────────────────────────────
SALES_DEDUCT_BRACKETS_GOODS = [
    (100_000_000, 0),
    (20_000_000,  4),
]  # 그 미만은 10점 감점
SALES_DEDUCT_BRACKETS_CLOTHING = [
    (50_000_000, 0),
    (4_000_000,  4),
]  # 그 미만은 10점 감점


def parse_amount(val):
    """'12,000,000' 같은 문자열/숫자를 금액(float)으로 변환. 실패하면 0."""
    s = str(val).strip().replace(',', '').replace('원', '')
    if not s:
        return 0
    try:
        return float(s)
    except ValueError:
        return 0


def deduct_sales(amount, brackets):
    """3개월 매출 합계 구간 감점 (10점 만점에서 감점할 점수 반환)."""
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        return 10
    for threshold, deduct in brackets:
        if amount >= threshold:
            return deduct
    return 10


def score_sales_tier_goods(amount):
    return 10 - deduct_sales(amount, SALES_DEDUCT_BRACKETS_GOODS)


def score_sales_tier_clothing(amount):
    return 10 - deduct_sales(amount, SALES_DEDUCT_BRACKETS_CLOTHING)


_quota_exhausted = False  # 일일 인사이트(Gemini) 호출 중 하루 할당량 초과가 확인되면 True



# ───────────────────────────────────────────
# 일일 AI 인사이트 코멘트 ("AI가 분석한 오늘의 대리점 인사이트" 팝업용)
# 하루 한 번만 Gemini를 호출해 경고/기회/검토 대상 대리점에 대한 코멘트를 생성.
# 같은 날 빌드가 여러 번 돌아도(파일 재업로드 등) insight_cache.json의 날짜를 확인해 재호출하지 않음.
# ───────────────────────────────────────────
INSIGHT_CACHE_FILE = 'insight_cache.json'


def get_kst_today_str():
    kst = timezone(timedelta(hours=9))
    return datetime.now(kst).strftime('%Y-%m-%d')


def load_insight_cache():
    try:
        with open(INSIGHT_CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_insight_cache(cache):
    try:
        with open(INSIGHT_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  ⚠️ 일일 인사이트 캐시 저장 실패: {e}")


def generate_daily_insights(warn_list, opp_list, review_list, api_key):
    """warn_list/opp_list/review_list: [{'name':..., 'total_score':..., 'worst_risk':..., 'ai_assessed_risk':...}, ...]
    반환: {대리점명: 코멘트} - 오늘 이미 생성된 적 있으면 캐시를 그대로 재사용."""
    today = get_kst_today_str()
    cache = load_insight_cache()
    if cache.get('date') == today and cache.get('comments'):
        print(f"  📦 오늘({today}) 일일 인사이트 이미 생성됨 - 캐시 재사용 ({len(cache['comments'])}건)")
        return cache['comments']

    if not api_key:
        print("  ⚠️ GEMINI_API_KEY 없음 - 일일 인사이트 코멘트 생성 건너뜀")
        return {}

    global _quota_exhausted
    if _quota_exhausted:
        print("  ⚠️ Gemini 할당량 소진 상태 - 일일 인사이트 코멘트 생성 건너뜀")
        return {}

    def block(store, tag):
        return (f"[{tag}] name: \"{store['name']}\"\n"
                f"- 종합점수: {round(store['total_score'])}점\n"
                f"- 등급: {store['worst_risk']}\n"
                f"- CS 키워드 등급: {store.get('ai_assessed_risk') or '-'}")

    blocks = ([block(s, '경고') for s in warn_list]
              + [block(s, '기회') for s in opp_list]
              + [block(s, '검토') for s in review_list])
    if not blocks:
        return {}

    stores_text = '\n\n'.join(blocks)
    prompt = f"""당신은 스포츠용품 유통사의 대리점 채권 리스크 담당 분석가입니다.
아래는 오늘 대시보드 상단 'AI가 분석한 오늘의 대리점 인사이트'에 노출될 대리점 목록입니다.
각 대리점 태그(경고/기회/검토)에 맞게, 영업 담당자가 한눈에 왜 이 대리점이 이 카테고리에 포함됐는지 이해할 수 있도록 코멘트를 작성하세요.

{stores_text}

각 대리점마다 1문장(간결하게, 가능하면 수치 근거 포함)으로 작성하세요.
- 경고 태그: 왜 위험한지와 어떤 조치가 필요한지
- 기회 태그: 왜 안정적인지와 어떤 점을 유지하면 좋은지
- 검토 태그: AI 판단과 기계적 등급이 왜 다를 수 있는지에 대한 추정

반드시 아래 JSON 배열 형식으로만 출력하세요. 다른 텍스트 일절 금지:
[{{"name": "대리점명(위와 정확히 동일하게)", "comment": "코멘트 한 문장"}}, ...]"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": min(4096, 150 * len(blocks) + 200)},
    }
    try:
        res = requests.post(url, json=body, timeout=60)
        if res.status_code == 429:
            err = res.json().get('error', {}) if res.content else {}
            if err.get('status') == 'RESOURCE_EXHAUSTED':
                _quota_exhausted = True
            print("  ⚠️ 일일 인사이트 생성 실패 - Gemini 한도 초과")
            return {}
        data = res.json()
        if res.status_code != 200:
            err = data.get('error', {})
            print(f"  ⚠️ 일일 인사이트 Gemini 오류: HTTP {res.status_code} / {err.get('message','')[:200]}")
            return {}
        candidates = data.get('candidates', [])
        if not candidates:
            print("  ⚠️ 일일 인사이트 응답에 candidates 없음")
            return {}
        raw = candidates[0]['content']['parts'][0]['text'].strip()
        raw = raw.strip('`').replace('json\n', '', 1).strip()
        parsed_list = json.loads(raw)
        comments = {}
        for item in parsed_list:
            name = ' '.join(str(item.get('name', '')).split())
            comment = str(item.get('comment', '')).strip()
            if name and comment:
                comments[name] = comment
        save_insight_cache({'date': today, 'comments': comments})
        print(f"  ✅ 일일 인사이트 코멘트 생성 완료 ({len(comments)}건, {today})")
        return comments
    except Exception as e:
        print(f"  ⚠️ 일일 인사이트 생성 오류: {type(e).__name__}: {e}")
        return {}


# ───────────────────────────────────────────
# 메인: CS 데이터 fetch
# ───────────────────────────────────────────
def fetch_cs_data(store_debt_map={}):
    """CS 대리점평가 구글시트(NO./대리점명/작성자/키워드/작성일/파트너십_용품/파트너십_의류)에서
    담당자가 선택한 '키워드'를 KEYWORD_TIERS 표와 매칭해 기계적으로 CS 점수를 산정.
    (Gemini 자유판단은 더 이상 사용하지 않음 - 시트가 자유 서술 메모 대신 고정 키워드 선택 방식으로 바뀌었기 때문)"""
    try:
        records = fetch_sheet_data()
    except Exception as e:
        print(f"구글 시트 읽기 실패: {e}")
        return {}

    today = datetime.now().date()
    cur_year, cur_month = today.year, today.month
    skipped_old = 0

    # 별도 탭('용품_3개월 매출' / '의류_3개월 매출')에서 최근 3개월 합계 조회
    try:
        sales_3m_goods_tab = fetch_sales_tab('용품_3개월 매출')
    except Exception as e:
        print(f"용품 3개월 매출 탭 조회 실패: {e}")
        sales_3m_goods_tab = {}
    try:
        sales_3m_clothing_tab = fetch_sales_tab('의류_3개월 매출')
    except Exception as e:
        print(f"의류 3개월 매출 탭 조회 실패: {e}")
        sales_3m_clothing_tab = {}

    # CS 시트 병합 (대리점명 기준)
    # - 키워드: 이번 달 작성건만 반영 (매달 리셋)
    # - 파트너십(용품/의류): 작성일과 무관하게 항상 최신 값 반영 (상태값이라 월 필터 미적용)
    merged = {}
    for row in records:
        name    = ' '.join(str(row.get('대리점명', '')).split())
        keyword = str(row.get('키워드', '')).strip()
        p_goods = str(row.get('파트너십_용품', '')).strip()
        p_cloth = str(row.get('파트너십_의류', '')).strip()
        written = parse_sheet_date(row.get('작성일', ''))
        if not name:
            continue
        sales_3m_goods    = parse_amount(row.get('매출3개월_용품', ''))
        sales_3m_clothing = parse_amount(row.get('매출3개월_의류', ''))
        if name not in merged:
            merged[name] = {'keywords': [], 'p_goods': '', 'p_clothing': '',
                             'sales_3m_goods': 0, 'sales_3m_clothing': 0}
        # 파트너십/3개월매출은 월 필터 없이 항상 최신 행 값으로 덮어씀
        if p_goods:
            merged[name]['p_goods'] = p_goods
        if p_cloth:
            merged[name]['p_clothing'] = p_cloth
        if sales_3m_goods:
            merged[name]['sales_3m_goods'] = sales_3m_goods
        if sales_3m_clothing:
            merged[name]['sales_3m_clothing'] = sales_3m_clothing
        # 키워드는 이번 달 작성분만 반영 (과거 기록 누적 방지)
        if written and (written.year != cur_year or written.month != cur_month):
            skipped_old += 1
            continue
        if keyword:
            merged[name]['keywords'].append(keyword)

    if skipped_old:
        print(f"  이번 달({cur_year}-{cur_month:02d}) 이전 작성 키워드 {skipped_old}건 제외 (파트너십은 반영됨)")

    # '용품_3개월 매출'/'의류_3개월 매출' 탭 값을 우선 반영 (있으면 덮어씀, 탭에만 있는 대리점은 새로 추가)
    for name, total in sales_3m_goods_tab.items():
        if name not in merged:
            merged[name] = {'keywords': [], 'p_goods': '', 'p_clothing': '',
                             'sales_3m_goods': 0, 'sales_3m_clothing': 0}
        merged[name]['sales_3m_goods'] = total
    for name, total in sales_3m_clothing_tab.items():
        if name not in merged:
            merged[name] = {'keywords': [], 'p_goods': '', 'p_clothing': '',
                             'sales_3m_goods': 0, 'sales_3m_clothing': 0}
        merged[name]['sales_3m_clothing'] = total

    result = {}
    for name, data in merged.items():
        keyword_str = ', '.join(data['keywords'])
        p_goods = data['p_goods']
        p_clothing = data['p_clothing']
        sales_3m_goods    = data.get('sales_3m_goods', 0)
        sales_3m_clothing = data.get('sales_3m_clothing', 0)

        cs_score, tier = score_cs_from_keywords(keyword_str)
        if keyword_str:
            if tier:
                comment = f"선택 키워드 '{keyword_str}' → {tier} 등급으로 CS {cs_score}점 반영"
            else:
                comment = f"선택 키워드 '{keyword_str}'가 기준표에 없어 감점 없이 처리(20점)"
            print(f"  {name}: 키워드 '{keyword_str}' → {tier or '미분류'} / CS {cs_score}점")
        else:
            comment = ''

        partnership_score = score_partnership(p_goods, p_clothing)
        sales_score_goods    = score_sales_tier_goods(sales_3m_goods)
        sales_score_clothing = score_sales_tier_clothing(sales_3m_clothing)

        result[name] = {
            'score':                  cs_score,
            'partnership_score':      partnership_score,
            'sales_score':            sales_score_goods + sales_score_clothing,
            'sales_score_goods':      sales_score_goods,
            'sales_score_clothing':   sales_score_clothing,
            'sales_3m_goods':         sales_3m_goods,
            'sales_3m_clothing':      sales_3m_clothing,
            'p_goods':                p_goods,
            'p_clothing':             p_clothing,
            'keywords':               keyword_str,
            'memo':                   '',
            'ai_comment':             comment,
            'ai_assessed_risk':       tier,
            'ai_mismatch':            False,
            'ai_mismatch_direction':  '',
        }

    return result
