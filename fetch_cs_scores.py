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

# 법인 접두어("주식회사"/"(주)"/"㈜") 유무 차이로 이름 매칭이 실패하는 것을 막기 위한 정규화.
# ERP 원본 데이터 자체가 "주식회사 나이스패밀리"(띄어쓰기 있음)와 "주식회사동우스포츠"(띄어쓰기 없음)처럼
# 표기가 일관되지 않아, 접두어 뒤 공백 유무와 무관하게 제거해야 한다.
# "_특판", "_단독" 같은 지점 구분 접미어는 서로 다른 거래선을 가리키므로 절대 제거하지 않는다.
_CORP_PREFIXES = ['주식회사', '(주)', '㈜']
_CORP_SUFFIXES = ['주식회사']


def normalize_dealer_name_for_matching(name):
    """대리점명에서 법인 접두/접미어만 제거한 매칭용 키를 반환 (표시용 이름은 원본 그대로 유지)."""
    name = ' '.join(str(name).split())
    for prefix in _CORP_PREFIXES:
        if name.startswith(prefix):
            name = name[len(prefix):].strip()
            break
    for suffix in _CORP_SUFFIXES:
        if name.endswith(suffix):
            name = name[:-len(suffix)].strip()
            break
    return name

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

# CS 키워드 3단계 표 (정성 점수 CS코멘트, 기본 30점 - 월 누적 방식)
# CS 대리점평가 구글시트의 '키워드' 컬럼에서 담당자가 직접 선택한 값을 이 표와 매칭.
# 한 행(=한 상담 건)에 여러 등급 키워드가 섞여 있으면 그 행에서는 가장 심각한 등급 하나만 인정(중복집계 방지).
# 이번 달에 작성된 행들을 모두 모아 등급별 건수를 세고, 건당 아래 배점을 누적 적용:
#   고위험 -5점/건, 중위험 -2점/건, 저위험(우수) +1점/건 - 최종 점수는 0~30점 사이로 clamp.
# 키워드 목록은 추후 바뀔 수 있음 — 여기만 수정하면 전체 채점에 반영됨.
KEYWORD_TIERS = {
    '고위험': ['허위안내', 'AS접수거부', '임의판정', '폭언·욕설', '보증서조작', 'AS규정위반', '클레임', '컴플레인', '한국소비자원신고', '법적조치', 'SNS게시'],
    '중위험': ['AS기준오안내', '접수지연', '응대미흡', '불만', '불편', '실망', '불친절', '개선요청'],
    '저위험': ['친절', '신속처리', '적극협조', '고객칭찬', '만족도우수', '응대우수', '규정준수', '정확한안내', '클레임없음'],
}
CS_BASE_SCORE = 30
KEYWORD_POINTS = {'고위험': -5, '중위험': -2, '저위험': 1}  # 건당 가감점 (저위험은 우수 보너스)

PARTNERSHIP_BASE_SCORE = 30  # 위반 1건(용품/의류 각 컬럼에 값이 있으면 1건)당 -1점, 월 누적

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


def fetch_monthly_series(tab_name, name_col='매장명', year=None):
    """'용품_3개월 매출' / '의류_3개월 매출' 탭에서 지정 연도(year)에 해당하는 월별 컬럼을
    전부(3개월 제한 없이) 읽어서 {대리점명: {월(1~12): 금액}} 형태로 반환.
    fetch_sales_tab()은 최근 3개월 합계만 필요할 때, 이 함수는 월별 상세 화면처럼
    개별 월 값이 다 필요할 때 사용. 시트에 실제로 입력된 월만 키로 존재한다
    (미입력 월은 아예 키가 없음 - 프론트에서 '실적 미입력'으로 구분하는 데 사용)."""
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

    month_cols = []  # (month, col_idx) - year 필터링됨
    for idx, h in enumerate(header):
        parsed = parse_month_header(h)
        if parsed and (year is None or parsed[0] == year):
            month_cols.append((parsed[1], idx))
    if not month_cols:
        return {}

    result = {}
    for row in values[1:]:
        if len(row) <= name_idx:
            continue
        name = ' '.join(str(row[name_idx]).split())
        if not name:
            continue
        monthly = {}
        for month, idx in month_cols:
            if idx < len(row):
                monthly[month] = parse_amount(row[idx])
        result[name] = monthly
    return result


def get_recent_months(tab_name, name_col='매장명', count=3):
    """매출 탭(예: '용품_3개월 매출') 헤더에서 최근 N개월의 (year, month)를 판별해서 반환.
    fetch_sales_tab과 동일한 컬럼 판별 로직을 재사용하되 헤더만 읽는다(목표 raw 데이터의
    3개월 구간을 매출 탭과 동일하게 맞추기 위함)."""
    spreadsheet_id = os.environ.get('SPREADSHEET_ID', '')
    client = get_sheets_client()
    try:
        ws = client.open_by_key(spreadsheet_id).worksheet(tab_name)
        header = ws.row_values(1)
    except Exception as e:
        print(f"'{tab_name}' 탭 헤더 조회 실패: {e}")
        return []

    month_cols = []
    for h in header:
        parsed = parse_month_header(h)
        if parsed:
            month_cols.append(parsed)
    month_cols.sort(key=lambda x: (x[0], x[1]))
    return month_cols[-count:]


_TARGET_RAW_FILES = {
    '용품': 'target_goods_raw.json',
    '의류': 'target_clothing_raw.json',
}
_target_raw_cache = {}


def load_target_raw(category):
    """매출목표달성 배점용 연간 목표 raw 데이터(용품_목표매출달성.xls/의류_목표매출달성.xls를
    미리 파싱해 커밋해둔 target_goods_raw.json/target_clothing_raw.json)를 로드.
    목표는 매 사이클 바뀌지 않는 고정값이라 구글시트가 아니라 저장소에 정적 파일로 둔다.
    형식: {대리점명: {'code': 'D00001', 'monthly': [1월,2월,...,12월]}}"""
    if category in _target_raw_cache:
        return _target_raw_cache[category]
    fname = _TARGET_RAW_FILES.get(category)
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), fname)
    try:
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"'{fname}' 로드 실패: {e}")
        data = {}
    _target_raw_cache[category] = data
    return data


def target_3m_from_raw(category, recent_months):
    """load_target_raw()로 읽은 연간 목표에서 recent_months(get_recent_months() 결과)에
    해당하는 월만 합산 - 매출3개월 탭과 동일한 최근 3개월 구간을 사용. 반환: {대리점명: 3개월목표합}"""
    raw = load_target_raw(category)
    months = [m for (_, m) in recent_months if 1 <= m <= 12]
    result = {}
    for name, entry in raw.items():
        monthly = entry.get('monthly', [])
        result[name] = sum(monthly[m - 1] for m in months if m - 1 < len(monthly))
    return result


# ───────────────────────────────────────────
# CS 점수 계산 (20점 만점)
# 메모 없음 → 20점 (이슈 없음으로 간주)
# 메모 있음 → AI(Gemini)가 메모 텍스트를 읽고 직접 판단한 위험단계(assessed_risk)를 기준으로 배점
#             (고정 키워드 리스트와의 문자열 일치가 아니라, 자유 서술 메모의 맥락을 보고 판단)
# ───────────────────────────────────────────
# CS 점수 계산 (30점 만점, 월 누적) — 구글시트 '키워드' 컬럼 값을 KEYWORD_TIERS와 매칭
# 1) 행(상담 건) 단위로 최고 등급 하나만 분류 → 2) 이번 달 전체 행에 대해 등급별 건수를 세고
# 고위험 -5점/건, 중위험 -2점/건, 저위험(우수) +1점/건을 30점에서 가감(0~30 clamp)
# ───────────────────────────────────────────
def classify_keyword_tier(keyword_str):
    """한 행(상담 건)의 키워드 문자열에서 가장 심각한 등급 하나만 반환 ('고위험'/'중위험'/'저위험'/'').
    저위험(긍정) 문구 중 일부('클레임없음' 등)가 고/중위험 키워드('클레임')를 부분 문자열로 포함하는 경우가 있어,
    저위험 문구를 먼저 제거한 나머지 텍스트에서 고/중위험을 검사해 오탐을 막는다.
    대소문자 차이(예: 'as규정위반' vs 'AS규정위반')로 매칭이 빠지는 것을 막기 위해 대문자로 통일 후 비교한다."""
    if not keyword_str or not keyword_str.strip():
        return ''
    text = re.sub(r'\s+', '', keyword_str).upper()  # 띄어쓰기/대소문자 차이로 매칭이 빠지는 것을 방지
    matched_low = any(kw.upper() in text for kw in KEYWORD_TIERS['저위험'])
    stripped = text
    for kw in KEYWORD_TIERS['저위험']:
        stripped = stripped.replace(kw.upper(), '')
    for tier in ('고위험', '중위험'):
        if any(kw.upper() in stripped for kw in KEYWORD_TIERS[tier]):
            return tier
    return '저위험' if matched_low else ''  # 표에 없는 값 - 미분류(카운트 안 함)


def score_cs_cumulative(count_high, count_mid, count_low):
    """이번 달 누적 등급별 건수로 CS 점수 계산 (30점 기준, 0~30 clamp)."""
    raw = CS_BASE_SCORE + count_high * KEYWORD_POINTS['고위험'] + count_mid * KEYWORD_POINTS['중위험'] + count_low * KEYWORD_POINTS['저위험']
    return min(30, max(0, raw))


def score_partnership_cumulative(violations):
    """이번 달 누적 파트너십 위반 건수(용품+의류 합산)로 점수 계산 (30점 기준, 위반 1건당 -1점, 0~30 clamp)."""
    return min(PARTNERSHIP_BASE_SCORE, max(0, PARTNERSHIP_BASE_SCORE - violations))


# ───────────────────────────────────────────
# 매출규모 감점 (3개월 매출 합계, 용품 5점 + 의류 5점 = 10점 만점)
# (기존 20점(10+10) 만점이었으나, 매출목표 달성 배점(10점=5+5) 신설로 130점 체계 내에서
#  재배분 — 매출규모 20→10점으로 축소. 브라켓 감점폭·컷라인은 그대로 두고 최종 점수만
#  1/2로 스케일)
# 구글시트 '매출3개월_용품' / '매출3개월_의류' 컬럼(각 사업부 3개월 합계 금액)을 기준으로
# 5구간 감점 방식 채점 (5점 만점에서 감점).
# 근거: 매출_규모_기준.xlsx 용품_2안/의류_2안 시트 (회사 실제 영업매출=부가세제외 계산 기준,
# 최근 6개 분기 실측 합산 분포로 5분위(quintile) 재산정)
#
# ⚠️ ERP 매출 데이터는 부가세 포함 금액이지만, 회사 기준은 부가세 제외(영업매출)이므로
# 아래 컷라인은 부가세 제외(영업매출) 기준 금액임 — 부가세 포함 원본 금액은
# VAT_RATE로 나눠서(영업매출 환산) 비교함 (SALES_VAT_RATE, deduct_sales 참고).
# ───────────────────────────────────────────
SALES_VAT_RATE = 1.1  # 부가세 10% 가정 (부가세포함액 = 영업매출 × 1.1)
SALES_DEDUCT_BRACKETS_GOODS = [
    (130_000_000, 0),
    (70_000_000,  2),
    (35_000_000,  4),
    (20_000_000,  7),
]  # 2,000만 미만은 -10점 (영업매출/부가세 제외 기준)
SALES_DEDUCT_BRACKETS_CLOTHING = [
    (80_000_000, 0),
    (35_000_000, 2),
    (15_000_000, 4),
    (5_000_000,  7),
]  # 500만 미만은 -10점 (영업매출/부가세 제외 기준)


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
    """3개월 매출 합계(부가세 포함 원본) 구간 감점 (10점 만점에서 감점할 점수 반환).
    ERP 원본 금액은 부가세 포함이라, 컷라인(부가세 제외 기준)과 비교하기 전에
    SALES_VAT_RATE로 나눠 영업매출(부가세 제외)로 환산한다."""
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        return 10
    net_amount = amount / SALES_VAT_RATE
    for threshold, deduct in brackets:
        if net_amount >= threshold:
            return deduct
    return 10


def score_sales_tier_goods(amount):
    return (10 - deduct_sales(amount, SALES_DEDUCT_BRACKETS_GOODS)) / 2


def score_sales_tier_clothing(amount):
    return (10 - deduct_sales(amount, SALES_DEDUCT_BRACKETS_CLOTHING)) / 2


# ───────────────────────────────────────────
# 매출목표 달성 배점 (3개월 목표 합계 대비 3개월 매출 합계, 용품 5점 + 의류 5점 = 10점 만점)
# '용품_목표3개월' / '의류_목표3개월' 구글시트 탭(매출3개월 탭과 동일한 구조: 매장명 + 월별
# 컬럼)에서 fetch_sales_tab()으로 최근 3개월 목표 합계를 자동 산출해 비교한다.
# 목표 raw 파일(용품_목표매출달성.xls/의류_목표매출달성.xls) 기준 목표값은 부가세 제외
# (영업매출) 금액이므로, ERP 매출(부가세 포함)은 SALES_VAT_RATE로 나눠 환산 후 비교.
# 달성(영업매출 환산 매출 ≥ 목표) 5점 / 미달성 0점. 목표가 0/미입력이면 판정 불가로 보고
# 감점 없이 5점(달성) 처리.
# ───────────────────────────────────────────
def score_target_achieve(sales_3m_amount, target_3m_amount):
    try:
        target = float(target_3m_amount)
    except (TypeError, ValueError):
        target = 0
    if target <= 0:
        return 5
    try:
        net_sales = float(sales_3m_amount) / SALES_VAT_RATE
    except (TypeError, ValueError):
        net_sales = 0
    return 5 if net_sales >= target else 0


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


def save_insight_cache(updates):
    """updates에 담긴 키만 기존 캐시에 병합해서 저장 (덮어쓰지 않음).
    예: generate_daily_insights가 comments/flagged만 저장해도 opp_shown_names 등
    다른 곳에서 저장한 값이 사라지지 않도록 함."""
    try:
        cache = load_insight_cache()
        cache.update(updates)
        with open(INSIGHT_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  ⚠️ 일일 인사이트 캐시 저장 실패: {e}")


def _fallback_daily_selection(warn_pool, opp_pool, review_pool):
    """Gemini 호출 없이(키 없음/할당량 소진/파싱 실패) 안전하게 쓸 결정론적 대체 선택.
    각 풀은 이미 심각도/점수 기준으로 정렬되어 들어오므로 그냥 앞에서 3개."""
    return {
        'warn': [s['name'] for s in warn_pool[:3]],
        'opp': [s['name'] for s in opp_pool[:3]],
        'review': [s['name'] for s in review_pool[:3]],
    }


def generate_daily_insights(warn_pool, opp_pool, review_pool, candidate_list, api_key, prev_selection=None):
    """warn_pool/opp_pool/review_pool: 점수 기준으로 넉넉히(최대 10개) 추린 후보 목록
    [{'name':..., 'total_score':..., 'worst_risk':..., 'ai_assessed_risk':...}, ...].
    이 함수가 각 풀 안에서 '오늘 대시보드에 보여줄 최대 3곳'을 AI가 직접 선택한다
    (기존에는 점수 정렬 + 하드코딩 로테이션으로 기계적으로 뽑았으나, 이제는 AI 판단에 맡김).
    candidate_list: [{'name':..., 'total_score':..., 'worst_risk':..., 'keywords':..., 'count_high':, 'count_mid':, 'count_low':}, ...]
    - 아직 위기/경계/관리 등급은 아니지만 이번 달 CS 키워드가 입력된 대리점들. AI가 키워드 내용을 읽고
      기계적 등급에는 안 잡혔지만 위험 신호가 보이는 곳을 추가로 골라낸다('AI 추가 발견').
    prev_selection: {'warn':[...], 'opp':[...], 'review':[...]} 어제 노출된 목록(참고용 - 강제 배제 아님).
    반환: {'selected': {'warn':[...],'opp':[...],'review':[...]}, 'comments': {대리점명: 코멘트}, 'flagged': {대리점명: 사유}}
    - 오늘 이미 생성된 적 있으면 캐시를 그대로 재사용."""
    today = get_kst_today_str()
    cache = load_insight_cache()
    if cache.get('date') == today and cache.get('selected'):
        sel = cache.get('selected', {})
        n_c, n_f = len(cache.get('comments', {})), len(cache.get('flagged', {}))
        print(f"  📦 오늘({today}) 일일 인사이트 이미 생성됨 - 캐시 재사용 "
              f"(경고 {len(sel.get('warn',[]))}·기회 {len(sel.get('opp',[]))}·검토 {len(sel.get('review',[]))}, "
              f"코멘트 {n_c}건, AI 추가 발견 {n_f}건)")
        return {'selected': sel, 'comments': cache.get('comments', {}), 'flagged': cache.get('flagged', {})}

    fallback_selected = _fallback_daily_selection(warn_pool, opp_pool, review_pool)

    if not api_key:
        print("  ⚠️ GEMINI_API_KEY 없음 - AI 선택/코멘트 생성 건너뜀 (점수순 상위 3개로 대체)")
        return {'selected': fallback_selected, 'comments': {}, 'flagged': {}}

    global _quota_exhausted
    if _quota_exhausted:
        print("  ⚠️ Gemini 할당량 소진 상태 - AI 선택/코멘트 생성 건너뜀 (점수순 상위 3개로 대체)")
        return {'selected': fallback_selected, 'comments': {}, 'flagged': {}}

    def pool_block(store, tag):
        # 등급은 반드시 종합점수(130점) 기준 score_grade를 써야 함 - worst_risk는 담보초과율 기반
        # 개별 채권 리스크라 종합점수와 불일치할 수 있음(예: 76점인데 worst_risk만 '위기'인 경우)
        return (f"[{tag}후보] name: \"{store['name']}\"\n"
                f"- 종합점수: {round(store['total_score'])}점\n"
                f"- 등급: {store['score_grade']}\n"
                f"- CS 키워드 등급: {store.get('ai_assessed_risk') or '-'}")

    def candidate_block(store):
        return (f"[후보] name: \"{store['name']}\"\n"
                f"- 종합점수: {round(store['total_score'])}점\n"
                f"- 등급: {store['score_grade']}\n"
                f"- 이번 달 CS 키워드: {store.get('keywords') or '-'}\n"
                f"- 고위험 {store.get('count_high',0)}건 · 중위험 {store.get('count_mid',0)}건 · 우수 {store.get('count_low',0)}건")

    if not warn_pool and not opp_pool and not review_pool and not candidate_list:
        return {'selected': fallback_selected, 'comments': {}, 'flagged': {}}

    warn_text = '\n\n'.join(pool_block(s, '경고') for s in warn_pool) if warn_pool else '(해당 없음)'
    opp_text = '\n\n'.join(pool_block(s, '기회') for s in opp_pool) if opp_pool else '(해당 없음)'
    review_text = '\n\n'.join(pool_block(s, '검토') for s in review_pool) if review_pool else '(해당 없음)'
    candidates_text = '\n\n'.join(candidate_block(s) for s in candidate_list) if candidate_list else '(해당 없음)'

    prev = prev_selection or {}
    prev_text = (f"어제 노출: 경고 {prev.get('warn', []) or '없음'} / "
                 f"기회 {prev.get('opp', []) or '없음'} / 검토 {prev.get('review', []) or '없음'}")

    prompt = f"""당신은 스포츠용품 유통사의 대리점 채권 리스크 담당 분석가입니다.
아래는 오늘 대시보드 상단 'AI가 분석한 오늘의 대리점 인사이트'에 노출할 대리점을 고르기 위한 후보 목록입니다.
경고/기회/검토 각 카테고리에서, 아래 후보 중 오늘 담당자가 가장 먼저 봐야 할 곳을 최대 3곳씩 직접 선택하세요
(후보가 3곳 미만이면 있는 만큼만 선택). 반드시 해당 카테고리 후보 목록 안에서만 골라야 합니다.

[경고 후보] (위기/경계/관리 등급)
{warn_text}

[기회 후보] (적정 등급, 고득점)
{opp_text}

[검토 후보] (AI 판단과 기계적 등급이 불일치)
{review_text}

{prev_text}
- 위 '어제 노출' 목록은 참고만 하세요. 오늘도 여전히 가장 눈에 띄는 곳이면 반복 선택해도 되고,
  비슷한 수준의 다른 후보가 있다면 다양하게 섞어서 보여주는 것도 좋습니다. 담당자에게 매번 같은
  대리점만 반복해서 보여주는 것은 피하되, 억지로 순서를 바꾸지는 마세요 — 정말 그날 상황에 맞게 판단하세요.

각 선택된 대리점마다, 담당자가 한눈에 왜 이 대리점이 이 카테고리에 포함됐는지 이해할 수 있도록 코멘트도 함께 작성하세요.
줄바꿈 없이 1문장(간결하게, 가능하면 수치 근거 포함)으로 작성하세요.
어조는 "~하세요/~하십시오" 같은 지시나 명령이 아니라, "~해보는 건 어떨까요/~하면 좋을 것 같습니다/~을 권장합니다" 같은
부드러운 권유·제안 톤으로 작성하세요. "즉시", "강도 높은 조사", "조사가 필요합니다" 같이 단정적이고 강한 표현은 피하고,
"확인해보시길 권장합니다", "한 번 살펴보시면 좋을 것 같습니다" 처럼 담당자의 판단을 존중하는 완곡한 표현을 쓰세요.
- 경고: 왜 위험한지와 어떤 조치를 고려해보면 좋을지
- 기회: 왜 안정적인지와 어떤 점을 유지하면 좋을지
- 검토: AI 판단과 기계적 등급이 왜 다를 수 있는지에 대한 추정

---

아래는 아직 종합 등급상 위험 등급(위기/경계/관리)에는 속하지 않지만, 이번 달 CS 키워드가 입력된 '추가 발견' 후보 대리점 목록입니다.
키워드 내용을 실제로 읽고, 건수가 적더라도 표현 수위가 심각하거나(예: 법적조치, 한국소비자원신고, 반복되는 불만 등)
앞으로 등급이 나빠질 조짐이 있다고 판단되는 대리점만 최대 3곳까지 골라주세요. 단순히 건수가 있다는 이유만으로 고르지 말고,
정말 주의가 필요하다고 판단될 때만 고르세요. 우려되는 곳이 없으면 빈 배열로 두세요.
사유(reason)도 위와 같은 부드러운 권유 톤으로 작성하고, "조사가 필요합니다" 같은 단정적 표현 대신
"한 번 확인해보시면 좋을 것 같습니다" 같은 완곡한 표현을 쓰세요.

{candidates_text}

반드시 아래 JSON 형식으로만 출력하세요. 다른 텍스트 일절 금지. 값 안에 줄바꿈 문자를 절대 넣지 마세요:
{{"selected": {{"warn": ["대리점명", ...최대3], "opp": ["대리점명", ...최대3], "review": ["대리점명", ...최대3]}},
 "comments": [{{"name": "대리점명(위와 정확히 동일하게)", "comment": "코멘트 한 문장"}}, ...],
 "flagged": [{{"name": "추가 발견 후보 목록 중 대리점명(정확히 동일하게)", "reason": "왜 주의가 필요한지 한 문장"}}, ...]}}"""

    total_blocks = len(warn_pool) + len(opp_pool) + len(review_pool) + len(candidate_list)
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        # maxOutputTokens: Gemini 2.5 Flash는 내부 "생각(thinking)" 토큰도 이 한도 안에서 소모하므로
        # 넉넉하게 잡지 않으면 실제 텍스트(JSON)가 다 나오기 전에 잘릴 수 있음 (Unterminated string 오류의 원인).
        # thinkingBudget: 0으로 생각 토큰 자체를 꺼서 전체 예산을 응답 텍스트에만 쓰도록 함.
        "generationConfig": {
            "maxOutputTokens": min(8192, 200 * total_blocks + 500),
            "thinkingConfig": {"thinkingBudget": 0},
        },
    }
    try:
        res = None
        for attempt in range(3):
            res = requests.post(url, json=body, timeout=60)
            if res.status_code in (503, 500, 502, 504):
                # 모델 일시 과부하 등 - 잠깐 대기 후 재시도 (최대 2회)
                if attempt < 2:
                    wait = 15 * (attempt + 1)
                    print(f"  ⏳ 일일 인사이트 생성 - Gemini 서버 일시 오류(HTTP {res.status_code}), {wait}초 후 재시도")
                    time.sleep(wait)
                    continue
            break
        if res.status_code == 429:
            err = res.json().get('error', {}) if res.content else {}
            if err.get('status') == 'RESOURCE_EXHAUSTED':
                _quota_exhausted = True
            print("  ⚠️ 일일 인사이트 생성 실패 - Gemini 한도 초과 (점수순 상위 3개로 대체)")
            return {'selected': fallback_selected, 'comments': {}, 'flagged': {}}
        data = res.json()
        if res.status_code != 200:
            err = data.get('error', {})
            print(f"  ⚠️ 일일 인사이트 Gemini 오류: HTTP {res.status_code} / {err.get('message','')[:200]} (점수순 상위 3개로 대체)")
            return {'selected': fallback_selected, 'comments': {}, 'flagged': {}}
        candidates_resp = data.get('candidates', [])
        if not candidates_resp:
            print("  ⚠️ 일일 인사이트 응답에 candidates 없음 (점수순 상위 3개로 대체)")
            return {'selected': fallback_selected, 'comments': {}, 'flagged': {}}
        raw = candidates_resp[0]['content']['parts'][0]['text'].strip()
        raw = raw.strip('`').replace('json\n', '', 1).strip()
        # strict=False: Gemini가 값 안에 이스케이프 안 된 줄바꿈(제어문자)을 넣는 경우가 있어
        # 기본 json 파서(strict=True)는 이를 오류로 처리함 - strict=False로 완화해서 그대로 허용
        parsed = json.loads(raw, strict=False)

        # 선택 결과 검증 - 각 카테고리는 반드시 해당 풀 안에서만 골라야 함(할루시네이션 방지),
        # 부족하면(개수 미달/전부 무효) 그 카테고리는 fallback으로 보충
        pool_names = {
            'warn': {s['name'] for s in warn_pool},
            'opp': {s['name'] for s in opp_pool},
            'review': {s['name'] for s in review_pool},
        }
        selected = {}
        raw_selected = parsed.get('selected', {})
        for cat in ('warn', 'opp', 'review'):
            names = [' '.join(str(n).split()) for n in raw_selected.get(cat, [])]
            valid = [n for n in names if n in pool_names[cat]]
            if len(valid) < min(3, len(pool_names[cat])):
                # 부족한 만큼 fallback(점수순 상위)에서 중복 없이 보충
                for n in fallback_selected[cat]:
                    if n not in valid:
                        valid.append(n)
                    if len(valid) >= min(3, len(pool_names[cat])):
                        break
            selected[cat] = valid[:3]

        comments = {}
        for item in parsed.get('comments', []):
            name = ' '.join(str(item.get('name', '')).split())
            comment = str(item.get('comment', '')).strip()
            if name and comment:
                comments[name] = comment
        candidate_names = {' '.join(s['name'].split()) for s in candidate_list}
        flagged = {}
        for item in parsed.get('flagged', []):
            name = ' '.join(str(item.get('name', '')).split())
            reason = str(item.get('reason', '')).strip()
            # 후보 목록에 없는 이름을 Gemini가 지어내는 경우를 대비해 검증
            if name and reason and name in candidate_names:
                flagged[name] = reason
        save_insight_cache({'date': today, 'selected': selected, 'comments': comments, 'flagged': flagged})
        print(f"  ✅ 일일 인사이트 AI 선택 완료 (경고 {len(selected['warn'])}·기회 {len(selected['opp'])}·검토 {len(selected['review'])}, "
              f"코멘트 {len(comments)}건, AI 추가 발견 {len(flagged)}건, {today})")
        return {'selected': selected, 'comments': comments, 'flagged': flagged}
    except Exception as e:
        print(f"  ⚠️ 일일 인사이트 생성 오류: {type(e).__name__}: {e} (점수순 상위 3개로 대체)")
        return {'selected': fallback_selected, 'comments': {}, 'flagged': {}}



# ───────────────────────────────────────────
# 메인: CS 데이터 fetch
# ───────────────────────────────────────────
def fetch_cs_data(store_debt_map={}):
    """CS 대리점평가 구글시트(NO./대리점명/작성자/키워드/작성일/파트너십_용품/파트너십_의류)에서
    이번 달에 작성된 행들을 모아 대리점별로 누적 집계해 CS/파트너십 점수를 산정.
    - CS: 행(상담 건)마다 키워드의 최고 등급 하나만 인정 → 등급별 건수 누적(고위험 -5/건, 중위험 -2/건, 저위험(우수) +1/건, 30점 기준 0~30 clamp)
    - 파트너십: 파트너십_용품/파트너십_의류 컬럼에 값이 있는 행 = 위반 1건, 용품+의류 합산 위반 건수만큼 30점에서 1점씩 차감
    - 시트에 입력된 대리점명이 ERP 등록명과 법인 접두어(주식회사/(주)/㈜) 유무만 다른 경우, 매 행을 읽는
      시점에 바로 ERP 등록명으로 정규화해서 병합한다 (한 대리점이 여러 달에 걸쳐 다른 표기로 입력돼도
      항상 하나의 키로 누적되도록 - 병합을 나중에 하면 이미 존재하는 빈 기본값 항목에 자리를 뺏길 수 있음)
    (Gemini 자유판단은 사용하지 않음 - 시트가 자유 서술 메모 대신 고정 키워드 선택 방식으로 바뀌었기 때문)"""
    try:
        records = fetch_sheet_data()
    except Exception as e:
        print(f"구글 시트 읽기 실패: {e}")
        return {}

    today = datetime.now().date()
    cur_year, cur_month = today.year, today.month
    skipped_old = 0

    # ERP 등록명 정규화 인덱스 (법인 접두어 유무와 무관하게 매칭용) - store_debt_map이 있을 때만
    erp_names = set(store_debt_map.keys())
    norm_to_erp = {}
    for erp_name in erp_names:
        norm_to_erp.setdefault(normalize_dealer_name_for_matching(erp_name), []).append(erp_name)

    def resolve_erp_name(raw_name):
        """시트 입력명을 가능하면 ERP 등록명으로 정규화해서 반환 (유일하게 매칭될 때만; 모호하면 원본 그대로)."""
        if not raw_name or raw_name in erp_names:
            return raw_name
        candidates = norm_to_erp.get(normalize_dealer_name_for_matching(raw_name), [])
        return candidates[0] if len(candidates) == 1 else raw_name

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

    # 매출목표 달성 "월별 상세" 화면용: 3개월 제한 없이 올해(cur_year)에 입력된 월별 실적을 전부 조회.
    # 시트에 없는 월은 dict에 키 자체가 없음 (프론트에서 '실적 미입력'으로 구분).
    try:
        sales_monthly_goods_tab = fetch_monthly_series('용품_3개월 매출', year=cur_year)
    except Exception as e:
        print(f"용품 월별 매출 조회 실패: {e}")
        sales_monthly_goods_tab = {}
    try:
        sales_monthly_clothing_tab = fetch_monthly_series('의류_3개월 매출', year=cur_year)
    except Exception as e:
        print(f"의류 월별 매출 조회 실패: {e}")
        sales_monthly_clothing_tab = {}

    # 매출목표 달성 배점(5+5=10점)용 목표값: 목표는 매 사이클 바뀌지 않는 고정값이라
    # 구글시트 대신 저장소에 커밋해둔 target_goods_raw.json/target_clothing_raw.json(연간
    # 월별 목표, 용품_목표매출달성.xls/의류_목표매출달성.xls를 미리 파싱한 것)을 사용.
    # '용품_3개월 매출' 탭 헤더에서 현재 활성 3개월을 판별해 그 월들만 목표에서 합산 -
    # 실제 매출 3개월 구간과 항상 동일한 달을 비교하게 된다.
    try:
        recent_months = get_recent_months('용품_3개월 매출')
    except Exception as e:
        print(f"최근 3개월 판별 실패: {e}")
        recent_months = []
    if not recent_months:
        # 폴백: 판별 실패 시 KST 기준 현재월 이전 3개월(당월 제외)을 사용
        recent_months = [(cur_year, ((cur_month - i - 1) % 12) + 1) for i in range(3)][::-1]
    target_3m_goods_tab = target_3m_from_raw('용품', recent_months)
    target_3m_clothing_tab = target_3m_from_raw('의류', recent_months)

    def new_entry():
        return {
            'keywords': [], 'count_high': 0, 'count_mid': 0, 'count_low': 0,
            'p_goods_count': 0, 'p_clothing_count': 0,
            'p_goods_latest': '', 'p_clothing_latest': '',
            'sales_3m_goods': 0, 'sales_3m_clothing': 0,
            'target_3m_goods': 0, 'target_3m_clothing': 0,
            'sales_monthly_goods': {}, 'sales_monthly_clothing': {},
        }

    # CS 시트 병합 (대리점명 기준) - 키워드/파트너십 둘 다 "이번 달 작성분만" 누적 집계 (매달 리셋)
    merged = {}
    for row in records:
        name    = resolve_erp_name(' '.join(str(row.get('대리점명', '')).split()))
        keyword = str(row.get('키워드', '')).strip()
        p_goods = str(row.get('파트너십_용품', '')).strip()
        p_cloth = str(row.get('파트너십_의류', '')).strip()
        written = parse_sheet_date(row.get('작성일', ''))
        if not name:
            continue
        sales_3m_goods    = parse_amount(row.get('매출3개월_용품', ''))
        sales_3m_clothing = parse_amount(row.get('매출3개월_의류', ''))
        if name not in merged:
            merged[name] = new_entry()
        # 3개월매출은 월 필터 없이 항상 최신 행 값으로 덮어씀 (상태값)
        if sales_3m_goods:
            merged[name]['sales_3m_goods'] = sales_3m_goods
        if sales_3m_clothing:
            merged[name]['sales_3m_clothing'] = sales_3m_clothing

        # 키워드/파트너십은 이번 달 작성분만 반영 (과거 기록 누적 방지)
        if written and (written.year != cur_year or written.month != cur_month):
            skipped_old += 1
            continue

        if keyword:
            merged[name]['keywords'].append(keyword)
            tier = classify_keyword_tier(keyword)  # 이 행(상담 건)의 최고 등급 하나만 인정
            if tier == '고위험':
                merged[name]['count_high'] += 1
            elif tier == '중위험':
                merged[name]['count_mid'] += 1
            elif tier == '저위험':
                merged[name]['count_low'] += 1
        if p_goods:
            merged[name]['p_goods_count'] += 1
            merged[name]['p_goods_latest'] = p_goods
        if p_cloth:
            merged[name]['p_clothing_count'] += 1
            merged[name]['p_clothing_latest'] = p_cloth

    if skipped_old:
        print(f"  이번 달({cur_year}-{cur_month:02d}) 이전 작성 CS/파트너십 기록 {skipped_old}건 제외 (매출3개월은 반영됨)")

    # '용품_3개월 매출'/'의류_3개월 매출' 탭 값을 우선 반영 (있으면 덮어씀, 탭에만 있는 대리점은 새로 추가)
    # 여기도 ERP 등록명으로 정규화해서 위 키워드/파트너십 집계와 같은 키로 합쳐지게 한다
    for name, total in sales_3m_goods_tab.items():
        name = resolve_erp_name(name)
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['sales_3m_goods'] = total
    for name, total in sales_3m_clothing_tab.items():
        name = resolve_erp_name(name)
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['sales_3m_clothing'] = total
    for name, total in target_3m_goods_tab.items():
        name = resolve_erp_name(name)
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['target_3m_goods'] = total
    for name, total in target_3m_clothing_tab.items():
        name = resolve_erp_name(name)
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['target_3m_clothing'] = total
    for name, monthly in sales_monthly_goods_tab.items():
        name = resolve_erp_name(name)
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['sales_monthly_goods'] = monthly
    for name, monthly in sales_monthly_clothing_tab.items():
        name = resolve_erp_name(name)
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['sales_monthly_clothing'] = monthly

    target_goods_raw = load_target_raw('용품')
    target_clothing_raw = load_target_raw('의류')

    result = {}
    for name, data in merged.items():
        keyword_str = ', '.join(data['keywords'])
        count_high, count_mid, count_low = data['count_high'], data['count_mid'], data['count_low']
        p_goods_count, p_clothing_count = data['p_goods_count'], data['p_clothing_count']
        sales_3m_goods    = data.get('sales_3m_goods', 0)
        sales_3m_clothing = data.get('sales_3m_clothing', 0)
        target_3m_goods    = data.get('target_3m_goods', 0)
        target_3m_clothing = data.get('target_3m_clothing', 0)

        cs_score = score_cs_cumulative(count_high, count_mid, count_low)
        if keyword_str:
            comment = (f"이번 달 고위험 {count_high}건(-{count_high*5}점) · 중위험 {count_mid}건(-{count_mid*2}점) · "
                       f"우수 {count_low}건(+{count_low}점) → CS {cs_score}점")
            worst_tier = '고위험' if count_high else ('중위험' if count_mid else ('저위험' if count_low else ''))
            print(f"  {name}: 키워드 '{keyword_str}' → 고위험{count_high}/중위험{count_mid}/우수{count_low} / CS {cs_score}점")
        else:
            comment = ''
            worst_tier = ''

        p_violations = p_goods_count + p_clothing_count
        partnership_score = score_partnership_cumulative(p_violations)
        sales_score_goods    = score_sales_tier_goods(sales_3m_goods)
        sales_score_clothing = score_sales_tier_clothing(sales_3m_clothing)
        target_score_goods    = score_target_achieve(sales_3m_goods, target_3m_goods)
        target_score_clothing = score_target_achieve(sales_3m_clothing, target_3m_clothing)
        sales_monthly_goods    = data.get('sales_monthly_goods', {})
        sales_monthly_clothing = data.get('sales_monthly_clothing', {})
        target_monthly_goods    = target_goods_raw.get(name, {}).get('monthly', [])
        target_monthly_clothing = target_clothing_raw.get(name, {}).get('monthly', [])

        result[name] = {
            'score':                  cs_score,
            'partnership_score':      partnership_score,
            'sales_score':            sales_score_goods + sales_score_clothing,
            'sales_score_goods':      sales_score_goods,
            'sales_score_clothing':   sales_score_clothing,
            'sales_3m_goods':         sales_3m_goods,
            'sales_3m_clothing':      sales_3m_clothing,
            'target_score':           target_score_goods + target_score_clothing,
            'target_score_goods':     target_score_goods,
            'target_score_clothing':  target_score_clothing,
            'target_3m_goods':        target_3m_goods,
            'target_3m_clothing':     target_3m_clothing,
            'sales_monthly_goods':    sales_monthly_goods,
            'sales_monthly_clothing': sales_monthly_clothing,
            'target_monthly_goods':   target_monthly_goods,
            'target_monthly_clothing': target_monthly_clothing,
            'p_goods':                data['p_goods_latest'],
            'p_clothing':             data['p_clothing_latest'],
            'p_goods_count':          p_goods_count,
            'p_clothing_count':       p_clothing_count,
            'keywords':               keyword_str,
            'memo':                   '',
            'ai_comment':             comment,
            'ai_assessed_risk':       worst_tier,
            'ai_mismatch':            False,
            'ai_mismatch_direction':  '',
            'count_high':             count_high,
            'count_mid':              count_mid,
            'count_low':              count_low,
        }

    return result
