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
    저위험 문구를 먼저 제거한 나머지 텍스트에서 고/중위험을 검사해 오탐을 막는다."""
    if not keyword_str or not keyword_str.strip():
        return ''
    text = re.sub(r'\s+', '', keyword_str)  # 띄어쓰기 차이로 매칭이 빠지는 것을 방지 (예: "정확한 안내" vs "정확한안내")
    matched_low = any(kw in text for kw in KEYWORD_TIERS['저위험'])
    stripped = text
    for kw in KEYWORD_TIERS['저위험']:
        stripped = stripped.replace(kw, '')
    for tier in ('고위험', '중위험'):
        if any(kw in stripped for kw in KEYWORD_TIERS[tier]):
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


def generate_daily_insights(warn_list, opp_list, review_list, candidate_list, api_key):
    """warn_list/opp_list/review_list: [{'name':..., 'total_score':..., 'worst_risk':..., 'ai_assessed_risk':...}, ...]
    candidate_list: [{'name':..., 'total_score':..., 'worst_risk':..., 'keywords':..., 'count_high':, 'count_mid':, 'count_low':}, ...]
    - 아직 위기/경계/관리 등급은 아니지만 이번 달 CS 키워드가 입력된 대리점들. AI가 키워드 내용을 읽고
      기계적 등급에는 안 잡혔지만 위험 신호가 보이는 곳을 추가로 골라낸다('AI 추가 발견').
    반환: {'comments': {대리점명: 코멘트}, 'flagged': {대리점명: 사유}} - 오늘 이미 생성된 적 있으면 캐시를 그대로 재사용."""
    today = get_kst_today_str()
    cache = load_insight_cache()
    if cache.get('date') == today and (cache.get('comments') or cache.get('flagged')):
        n_c, n_f = len(cache.get('comments', {})), len(cache.get('flagged', {}))
        print(f"  📦 오늘({today}) 일일 인사이트 이미 생성됨 - 캐시 재사용 (코멘트 {n_c}건, AI 추가 발견 {n_f}건)")
        return {'comments': cache.get('comments', {}), 'flagged': cache.get('flagged', {})}

    if not api_key:
        print("  ⚠️ GEMINI_API_KEY 없음 - 일일 인사이트 코멘트 생성 건너뜀")
        return {'comments': {}, 'flagged': {}}

    global _quota_exhausted
    if _quota_exhausted:
        print("  ⚠️ Gemini 할당량 소진 상태 - 일일 인사이트 코멘트 생성 건너뜀")
        return {'comments': {}, 'flagged': {}}

    def block(store, tag):
        return (f"[{tag}] name: \"{store['name']}\"\n"
                f"- 종합점수: {round(store['total_score'])}점\n"
                f"- 등급: {store['worst_risk']}\n"
                f"- CS 키워드 등급: {store.get('ai_assessed_risk') or '-'}")

    def candidate_block(store):
        return (f"[후보] name: \"{store['name']}\"\n"
                f"- 종합점수: {round(store['total_score'])}점\n"
                f"- 등급: {store['worst_risk']}\n"
                f"- 이번 달 CS 키워드: {store.get('keywords') or '-'}\n"
                f"- 고위험 {store.get('count_high',0)}건 · 중위험 {store.get('count_mid',0)}건 · 우수 {store.get('count_low',0)}건")

    blocks = ([block(s, '경고') for s in warn_list]
              + [block(s, '기회') for s in opp_list]
              + [block(s, '검토') for s in review_list])
    candidate_blocks = [candidate_block(s) for s in candidate_list]
    if not blocks and not candidate_blocks:
        return {'comments': {}, 'flagged': {}}

    stores_text = '\n\n'.join(blocks) if blocks else '(해당 없음)'
    candidates_text = '\n\n'.join(candidate_blocks) if candidate_blocks else '(해당 없음)'
    prompt = f"""당신은 스포츠용품 유통사의 대리점 채권 리스크 담당 분석가입니다.
아래는 오늘 대시보드 상단 'AI가 분석한 오늘의 대리점 인사이트'에 노출될 대리점 목록입니다.
각 대리점 태그(경고/기회/검토)에 맞게, 영업 담당자가 한눈에 왜 이 대리점이 이 카테고리에 포함됐는지 이해할 수 있도록 코멘트를 작성하세요.

{stores_text}

각 대리점마다 줄바꿈 없이 1문장(간결하게, 가능하면 수치 근거 포함)으로 작성하세요.
어조는 담당자에게 "~하세요/~하십시오" 같은 지시나 명령이 아니라, "~해보는 건 어떨까요/~하면 좋을 것 같습니다/~을 권장합니다" 같은
부드러운 권유·제안 톤으로 작성하세요. "즉시", "강도 높은 조사", "조사가 필요합니다" 같이 단정적이고 강한 표현은 피하고,
"확인해보시길 권장합니다", "한 번 살펴보시면 좋을 것 같습니다" 처럼 담당자의 판단을 존중하는 완곡한 표현을 쓰세요.
- 경고 태그: 왜 위험한지와 어떤 조치를 고려해보면 좋을지
- 기회 태그: 왜 안정적인지와 어떤 점을 유지하면 좋을지
- 검토 태그: AI 판단과 기계적 등급이 왜 다를 수 있는지에 대한 추정

---

아래는 아직 종합 등급상 위험 등급(위기/경계/관리)에는 속하지 않지만, 이번 달 CS 키워드가 입력된 '후보' 대리점 목록입니다.
키워드 내용을 실제로 읽고, 건수가 적더라도 표현 수위가 심각하거나(예: 법적조치, 한국소비자원신고, 반복되는 불만 등)
앞으로 등급이 나빠질 조짐이 있다고 판단되는 대리점만 최대 3곳까지 골라주세요. 단순히 건수가 있다는 이유만으로 고르지 말고,
정말 주의가 필요하다고 판단될 때만 고르세요. 우려되는 곳이 없으면 빈 배열로 두세요.
사유(reason)도 위와 같은 부드러운 권유 톤으로 작성하고, "조사가 필요합니다" 같은 단정적 표현 대신
"한 번 확인해보시면 좋을 것 같습니다" 같은 완곡한 표현을 쓰세요.

{candidates_text}

반드시 아래 JSON 형식으로만 출력하세요. 다른 텍스트 일절 금지. 값 안에 줄바꿈 문자를 절대 넣지 마세요:
{{"comments": [{{"name": "대리점명(위와 정확히 동일하게)", "comment": "코멘트 한 문장"}}, ...],
 "flagged": [{{"name": "후보 목록 중 대리점명(정확히 동일하게)", "reason": "왜 주의가 필요한지 한 문장"}}, ...]}}"""

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        # maxOutputTokens: Gemini 2.5 Flash는 내부 "생각(thinking)" 토큰도 이 한도 안에서 소모하므로
        # 넉넉하게 잡지 않으면 실제 텍스트(JSON)가 다 나오기 전에 잘릴 수 있음 (Unterminated string 오류의 원인).
        # thinkingBudget: 0으로 생각 토큰 자체를 꺼서 전체 예산을 응답 텍스트에만 쓰도록 함.
        "generationConfig": {
            "maxOutputTokens": min(8192, 200 * (len(blocks) + len(candidate_blocks)) + 500),
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
            print("  ⚠️ 일일 인사이트 생성 실패 - Gemini 한도 초과")
            return {'comments': {}, 'flagged': {}}
        data = res.json()
        if res.status_code != 200:
            err = data.get('error', {})
            print(f"  ⚠️ 일일 인사이트 Gemini 오류: HTTP {res.status_code} / {err.get('message','')[:200]}")
            return {'comments': {}, 'flagged': {}}
        candidates_resp = data.get('candidates', [])
        if not candidates_resp:
            print("  ⚠️ 일일 인사이트 응답에 candidates 없음")
            return {'comments': {}, 'flagged': {}}
        raw = candidates_resp[0]['content']['parts'][0]['text'].strip()
        raw = raw.strip('`').replace('json\n', '', 1).strip()
        # strict=False: Gemini가 값 안에 이스케이프 안 된 줄바꿈(제어문자)을 넣는 경우가 있어
        # 기본 json 파서(strict=True)는 이를 오류로 처리함 - strict=False로 완화해서 그대로 허용
        parsed = json.loads(raw, strict=False)
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
        save_insight_cache({'date': today, 'comments': comments, 'flagged': flagged})
        print(f"  ✅ 일일 인사이트 코멘트 생성 완료 (코멘트 {len(comments)}건, AI 추가 발견 {len(flagged)}건, {today})")
        return {'comments': comments, 'flagged': flagged}
    except Exception as e:
        print(f"  ⚠️ 일일 인사이트 생성 오류: {type(e).__name__}: {e}")
        return {'comments': {}, 'flagged': {}}


# ───────────────────────────────────────────
# 메인: CS 데이터 fetch
# ───────────────────────────────────────────
def fetch_cs_data(store_debt_map={}):
    """CS 대리점평가 구글시트(NO./대리점명/작성자/키워드/작성일/파트너십_용품/파트너십_의류)에서
    이번 달에 작성된 행들을 모아 대리점별로 누적 집계해 CS/파트너십 점수를 산정.
    - CS: 행(상담 건)마다 키워드의 최고 등급 하나만 인정 → 등급별 건수 누적(고위험 -5/건, 중위험 -2/건, 저위험(우수) +1/건, 30점 기준 0~30 clamp)
    - 파트너십: 파트너십_용품/파트너십_의류 컬럼에 값이 있는 행 = 위반 1건, 용품+의류 합산 위반 건수만큼 30점에서 1점씩 차감
    (Gemini 자유판단은 사용하지 않음 - 시트가 자유 서술 메모 대신 고정 키워드 선택 방식으로 바뀌었기 때문)"""
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

    def new_entry():
        return {
            'keywords': [], 'count_high': 0, 'count_mid': 0, 'count_low': 0,
            'p_goods_count': 0, 'p_clothing_count': 0,
            'p_goods_latest': '', 'p_clothing_latest': '',
            'sales_3m_goods': 0, 'sales_3m_clothing': 0,
        }

    # CS 시트 병합 (대리점명 기준) - 키워드/파트너십 둘 다 "이번 달 작성분만" 누적 집계 (매달 리셋)
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
    for name, total in sales_3m_goods_tab.items():
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['sales_3m_goods'] = total
    for name, total in sales_3m_clothing_tab.items():
        if name not in merged:
            merged[name] = new_entry()
        merged[name]['sales_3m_clothing'] = total

    result = {}
    for name, data in merged.items():
        keyword_str = ', '.join(data['keywords'])
        count_high, count_mid, count_low = data['count_high'], data['count_mid'], data['count_low']
        p_goods_count, p_clothing_count = data['p_goods_count'], data['p_clothing_count']
        sales_3m_goods    = data.get('sales_3m_goods', 0)
        sales_3m_clothing = data.get('sales_3m_clothing', 0)

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

        result[name] = {
            'score':                  cs_score,
            'partnership_score':      partnership_score,
            'sales_score':            sales_score_goods + sales_score_clothing,
            'sales_score_goods':      sales_score_goods,
            'sales_score_clothing':   sales_score_clothing,
            'sales_3m_goods':         sales_3m_goods,
            'sales_3m_clothing':      sales_3m_clothing,
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
