import os
import json
import re
import time
import hashlib
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

AI_CACHE_FILE = 'ai_cache.json'

def load_ai_cache():
    """이전 실행에서 저장한 AI 분석 결과 캐시를 불러옴. 메모가 안 바뀐 대리점은 재호출하지 않기 위함."""
    try:
        with open(AI_CACHE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_ai_cache(cache):
    try:
        with open(AI_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  ⚠️ AI 캐시 저장 실패: {e}")

def memo_hash(memo):
    return hashlib.md5(memo.encode('utf-8')).hexdigest()

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

# CS 키워드 정의
RISK_KEYWORDS = {
    'high': ['연락두절', '약속불이행', '클레임다발', '폐업징후', '허위접수', '재고과다', '타사이탈', '무단온라인', '연락안됨', '잠수'],
    'mid':  ['연락지연', '가끔약속어김', '클레임', '재고증가', '매출감소', '응대느림', '불만', 'A/S규정숙지'],
    'low':  ['협조적', '응대원활', '약속이행', '클레임없음', '재고적정', '매출안정', '신뢰', '칭찬']
}

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
CS_RISK_SCORE = {'적정': 20, '주의': 15, '경계': 10, '위기': 5}


def classify_cs_risk_fallback(memo):
    """Gemini 판단이 없을 때(API 실패 등)의 최후 대체 위험단계 추정.
    고정 키워드 매칭으로 대략적인 위험단계('적정'/'경계'/'위기')를 반환."""
    text = memo.lower()
    if any(kw in text for kw in RISK_KEYWORDS['high']):
        return '위기'
    if any(kw in text for kw in RISK_KEYWORDS['mid']):
        return '경계'
    return '적정'


def score_cs(memo, cs_risk):
    """CS 점수 계산 (20점 만점). cs_risk는 이 대리점의 CS 위험단계('적정'/'주의'/'경계'/'위기')."""
    # 메모 없으면 만점
    if not memo or not memo.strip():
        return 20

    if cs_risk in CS_RISK_SCORE:
        return CS_RISK_SCORE[cs_risk]

    # 위험단계를 판단하지 못한 예외적인 경우 최후의 키워드 기반 추정
    return CS_RISK_SCORE[classify_cs_risk_fallback(memo)]


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


# ───────────────────────────────────────────
# ───────────────────────────────────────────
# AI 키워드 추출 + 리스크 분석 (Gemini, 1회 호출로 통합)
# ───────────────────────────────────────────
RISK_ORDER = ['적정', '주의', '경계', '위기']  # 심각도 오름차순 (관리/해당없음은 별도 취급)

_quota_exhausted = False  # 이번 실행 중 하루 할당량 초과가 한 번이라도 확인되면 True (이후 재시도 없이 바로 대체 분석 사용)


def extract_keywords_rule_based(memo):
    """메모 텍스트에서 회사 키워드 목록을 직접 문자열 매칭으로 추출.
    Gemini 호출 여부와 무관하게 항상 동작 (CS 점수 감점의 근거가 되므로 API 의존 없이 안정적이어야 함)."""
    if not memo:
        return []
    all_kw = RISK_KEYWORDS['high'] + RISK_KEYWORDS['mid'] + RISK_KEYWORDS['low']
    return [kw for kw in all_kw if kw in memo]


def generate_rule_based_comment(store_name, memo, keywords, mechanical_risk, debt_info):
    """Gemini 호출 없이 숫자+키워드만으로 리스크 분석 코멘트를 생성.
    Gemini가 실패하거나 API 키가 없을 때도 AI 위험탐지 분석 박스가 항상 채워지도록 하는 대체 로직."""
    collateral = debt_info.get('collateral', 0)
    receivable = debt_info.get('receivable', 0)
    ratio = debt_info.get('ratio', 0)
    ratio_pct = int(ratio * 100)
    collection_days = max(debt_info.get('collection_days_clothing', 0), debt_info.get('collection_days_goods', 0))

    high_hits = [k for k in keywords if k in RISK_KEYWORDS['high']]
    mid_hits = [k for k in keywords if k in RISK_KEYWORDS['mid']]
    low_hits = [k for k in keywords if k in RISK_KEYWORDS['low']]

    # 리스크 요인
    risk_parts = []
    if collateral == 0 and receivable > 0:
        risk_parts.append(f"담보 없이 채권 {receivable:,}원이 발생한 상태로 회수 안전장치가 없습니다.")
    elif collateral:
        if ratio_pct > 0:
            risk_parts.append(f"담보 대비 채권 초과율 {ratio_pct}%로 담보금액을 채권잔액이 초과하고 있습니다.")
    if collection_days and collection_days > 60:
        risk_parts.append(f"회수일이 {collection_days}일로 기준(60일)을 초과했습니다.")
    if high_hits:
        risk_parts.append(f"CS 메모에 고위험 키워드({', '.join(high_hits)})가 확인됩니다.")
    if mid_hits:
        risk_parts.append(f"중위험 키워드({', '.join(mid_hits)})가 확인됩니다.")
    if not risk_parts:
        risk_parts.append("현재 수치상 뚜렷한 위험 요인은 확인되지 않습니다.")
    risk_line = ' '.join(risk_parts)

    # 긍정 요인
    good_parts = []
    if collateral and ratio_pct <= 0:
        good_parts.append(f"담보금액이 채권잔액보다 커서({abs(ratio_pct)}% 여유) 회수 안전성이 확보되어 있습니다.")
    if collection_days and collection_days <= 60:
        good_parts.append(f"회수일 {collection_days}일로 기준 이내를 유지하고 있습니다.")
    if low_hits:
        good_parts.append(f"CS상 긍정 신호({', '.join(low_hits)})가 확인됩니다.")
    good_line = ' '.join(good_parts) if good_parts else "해당 없음."

    # 권고사항
    rec_parts = []
    if mechanical_risk in ('위기', '관리'):
        rec_parts.append("담보 보강 또는 채권 회수 조치를 즉시 진행하고, 담당 영업사원의 대면 확인이 필요합니다.")
    elif mechanical_risk == '경계':
        rec_parts.append("담보대비 초과율과 회수 현황을 주 단위로 재점검하고 다음 달 매출 추이를 함께 모니터링하세요.")
    elif high_hits or mid_hits:
        rec_parts.append("CS 이슈 재발 여부를 담당 영업사원이 직접 확인하고 후속 조치 결과를 메모에 남겨주세요.")
    else:
        rec_parts.append("현재 특별한 조치는 필요하지 않으며, 정기 모니터링을 유지하세요.")
    rec_line = ' '.join(rec_parts)

    return f"🔴 리스크 요인: {risk_line}\n🟢 긍정 요인: {good_line}\n📋 권고사항: {rec_line}\n(※ 이 분석은 자동 규칙 기반으로 생성되었습니다)"


# ───────────────────────────────────────────
# AI 키워드 추출 + 리스크 분석 (Gemini, 여러 대리점을 한 번에 묶어서 호출)
# RPD(일일 요청수) 한도가 20으로 매우 낮아서(gemini-2.5-flash, 무료 등급),
# 대리점 1개당 1회 호출이 아니라 BATCH_SIZE개씩 묶어 1회 호출로 처리해 하루 처리 가능 대리점 수를 늘림.
# ───────────────────────────────────────────
BATCH_SIZE = 8  # 한 번의 Gemini 호출에 묶어서 보낼 대리점 수


def gemini_analyze_batch(stores, api_key):
    """stores: [{'name':..., 'memo':..., 'debt_info':...}, ...] (최대 BATCH_SIZE개)
    반환: {대리점명: (keywords, comment, assessed_risk)} — 판단 성공한 대리점만 포함.
    기계적 등급(classify_risk)은 프롬프트에 알려주지 않고, AI가 원본 숫자만으로 스스로 판단하게 함."""
    global _quota_exhausted
    if _quota_exhausted or not stores:
        return {}

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    all_keywords = RISK_KEYWORDS['high'] + RISK_KEYWORDS['mid'] + RISK_KEYWORDS['low']
    keyword_list = ', '.join(all_keywords)

    store_blocks = []
    for i, s in enumerate(stores, 1):
        debt_info = s.get('debt_info', {})
        collateral = debt_info.get('collateral', 0)
        receivable = debt_info.get('receivable', 0)
        ratio      = debt_info.get('ratio', 0)
        ratio_pct  = int(ratio * 100)
        col_str    = f"{collateral:,}원 ({collateral//10000:,}만원)" if collateral else "없음"
        rec_str    = f"{receivable:,}원 ({receivable//10000:,}만원)" if receivable else "없음"
        excess     = receivable - collateral
        excess_str = f"{excess:,}원 ({excess//10000:,}만원)" if excess > 0 else "초과 없음"
        store_blocks.append(f"""[대리점 {i}] name: "{s['name']}"
- 담보금액: {col_str}
- 채권잔액: {rec_str}
- 담보초과액: {excess_str}
- 담보대비 채권비율: {ratio_pct}%
- CS 특이사항 메모: {s['memo']}""")
    stores_text = '\n\n'.join(store_blocks)

    prompt = f"""당신은 스포츠용품 유통사의 대리점 채권·CS 리스크 전담 분석가입니다.
아래에 여러 대리점의 원본 데이터가 나열되어 있습니다. 각 대리점마다 회사가 정한 등급 기준을 참고하여 위험도를 직접 판단하세요.
(주의: 아래엔 최종 위험단계를 알려주지 않습니다 — 반드시 각 대리점별로 스스로 판단하세요)

{stores_text}

[참고: 회사 등급 기준 (담보대비 채권비율)]
- 적정: 60% 이하 / 주의: 60~100% / 경계: 100~150% / 위기: 150% 초과
- 단, 이 기준은 참고용이며 CS 메모의 심각도(연락두절, 폐업징후 등)를 종합적으로 고려해 기계적 기준보다 더 엄격하게(또는 완화해서) 판단해도 됩니다. 판단 근거를 반드시 밝히세요.

각 대리점마다 아래 작업을 수행하세요:
작업1) 해당 대리점 메모에 해당하는 키워드를 아래 목록에서만 골라 쉼표로 구분해 추출 (해당 없으면 빈 문자열)
사용 가능한 키워드 목록: {keyword_list}

작업2) 당신이 직접 판단한 위험단계를 "적정","주의","경계","위기" 중 하나로 선택

작업3) 영업팀 관리자가 즉시 활용할 수 있는 전문 분석 보고 작성:
🔴 리스크 요인: (채권 수치와 CS 이슈를 근거로 구체적 위험 요소 서술. 수치 반드시 포함.)
🟢 긍정 요인: (안정적 요소나 완화 요인이 있으면 서술. 없으면 "해당 없음".)
📋 권고사항: (담당 영업사원이 취해야 할 즉각적 조치를 1~2가지 구체적으로 제시.)
- 각 항목은 1~2문장으로 간결하되 수치 근거 포함, 불필요한 인사말/서론 없이 바로 본문 시작
- 만약 작업2에서 기계적 기준(비율 구간)과 다르게 판단했다면, 리스크 요인 첫 문장에 왜 다르게 판단했는지 명시

반드시 아래 JSON 배열 형식으로만, 대리점 순서대로 모두 출력하세요. 다른 텍스트 일절 금지:
[{{"name": "대리점명(위에 준 name 값과 정확히 동일하게)", "keywords": "쉼표로 구분된 키워드 또는 빈 문자열", "assessed_risk": "적정|주의|경계|위기 중 하나", "comment": "작업3 분석 보고 전문"}}, ...]"""

    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": min(8192, 250 * len(stores) + 200)},
    }
    store_names = [s['name'] for s in stores]
    try:
        res = None
        for attempt in range(2):
            res = requests.post(url, json=body, timeout=60)
            if res.status_code == 429:
                if attempt == 0:
                    print(f"  ⏳ Gemini 순간 한도 대기 (배치 {len(stores)}건) - 15초 후 1회 재시도")
                    time.sleep(15)
                    continue
                err = res.json().get('error', {}) if res.content else {}
                if err.get('status') == 'RESOURCE_EXHAUSTED':
                    _quota_exhausted = True
                    print(f"  ⚠️ Gemini 하루 할당량 소진 확인 (배치 {len(stores)}건) - 이후 대리점은 재시도 없이 대체 분석 사용")
            break
        data = res.json()
        if res.status_code != 200:
            err = data.get('error', {})
            print(f"  ⚠️ Gemini API 오류 (배치 {store_names}): HTTP {res.status_code} / {err.get('status','')} / {err.get('message','')[:200]}")
            return {}
        candidates = data.get('candidates', [])
        if not candidates:
            print(f"  ⚠️ Gemini 응답에 candidates 없음 (배치 {store_names}): {json.dumps(data, ensure_ascii=False)[:300]}")
            return {}
        finish_reason = candidates[0].get('finishReason', '')
        if finish_reason and finish_reason not in ('STOP', 'MAX_TOKENS'):
            print(f"  ⚠️ Gemini 응답 비정상 종료 (배치 {store_names}): finishReason={finish_reason}")
        raw = candidates[0]['content']['parts'][0]['text'].strip()
        raw = raw.strip('`').replace('json\n', '', 1).strip()
        parsed_list = json.loads(raw)
        if not isinstance(parsed_list, list):
            print(f"  ⚠️ Gemini 배치 응답이 배열이 아님 (배치 {store_names})")
            return {}

        results = {}
        used_idx = set()
        # 1차: name 값이 정확히 일치하는 항목끼리 매칭
        for item in parsed_list:
            item_name = ' '.join(str(item.get('name', '')).split())
            if item_name in store_names and item_name not in results:
                results[item_name] = item
        # 2차: 이름 매칭이 안 된 대리점이 남아있고 응답 개수가 동일하면 순서대로 매칭 (최후 안전장치)
        if len(results) < len(stores) and len(parsed_list) == len(stores):
            for name, item in zip(store_names, parsed_list):
                if name not in results:
                    results[name] = item

        out = {}
        for name, item in results.items():
            keywords = str(item.get('keywords', '')).strip()
            valid = [k.strip() for k in keywords.split(',') if k.strip() in all_keywords]
            comment = str(item.get('comment', '')).strip()
            assessed_risk = str(item.get('assessed_risk', '')).strip()
            if assessed_risk not in RISK_ORDER:
                assessed_risk = ''
            if comment:
                out[name] = (', '.join(valid), comment, assessed_risk)
        missing = [n for n in store_names if n not in out]
        if missing:
            print(f"  ⚠️ Gemini 배치 응답에서 판단 누락: {missing}")
        return out
    except Exception as e:
        print(f"  ⚠️ Gemini 배치 오류 ({store_names}): {type(e).__name__}: {e}")
        return {}


# ───────────────────────────────────────────
# 메인: CS 데이터 fetch
# ───────────────────────────────────────────
def fetch_cs_data(store_debt_map={}):
    api_key = os.environ.get('GEMINI_API_KEY', '')
    if not api_key:
        print("  ⚠️ GEMINI_API_KEY가 비어있음 - 모든 대리점의 AI 분석이 건너뛰어집니다. GitHub Secrets 설정을 확인하세요.")

    ai_cache = load_ai_cache()
    cache_hits = 0
    cache_calls = 0

    try:
        records = fetch_sheet_data()
    except Exception as e:
        print(f"구글 시트 읽기 실패: {e}")
        return {}

    today = datetime.now().date()
    cur_year, cur_month = today.year, today.month
    skipped_old_memo = 0

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
    # - 특이사항메모: 이번 달 작성건만 반영 (매달 리셋)
    # - 파트너십(용품/의류): 작성일과 무관하게 항상 최신 값 반영 (상태값이라 월 필터 미적용)
    merged = {}
    for row in records:
        name      = ' '.join(str(row.get('대리점명', '')).split())
        memo      = str(row.get('특이사항메모', '')).strip()
        p_goods   = str(row.get('파트너십_용품', '')).strip()
        p_cloth   = str(row.get('파트너십_의류', '')).strip()
        written   = parse_sheet_date(row.get('작성일', ''))
        if not name:
            continue
        sales_3m_goods    = parse_amount(row.get('매출3개월_용품', ''))
        sales_3m_clothing = parse_amount(row.get('매출3개월_의류', ''))
        if name not in merged:
            merged[name] = {'memos': [], 'p_goods': '', 'p_clothing': '',
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
        # 메모는 이번 달 작성분만 반영 (과거 메모 누적 방지)
        if written and (written.year != cur_year or written.month != cur_month):
            skipped_old_memo += 1
            continue
        if memo:
            merged[name]['memos'].append(memo)

    if skipped_old_memo:
        print(f"  이번 달({cur_year}-{cur_month:02d}) 이전 작성 메모 {skipped_old_memo}건 제외 (파트너십은 반영됨)")

    # '용품_3개월 매출'/'의류_3개월 매출' 탭 값을 우선 반영 (있으면 덮어씀, 탭에만 있는 대리점은 새로 추가)
    for name, total in sales_3m_goods_tab.items():
        if name not in merged:
            merged[name] = {'memos': [], 'p_goods': '', 'p_clothing': '',
                             'sales_3m_goods': 0, 'sales_3m_clothing': 0}
        merged[name]['sales_3m_goods'] = total
    for name, total in sales_3m_clothing_tab.items():
        if name not in merged:
            merged[name] = {'memos': [], 'p_goods': '', 'p_clothing': '',
                             'sales_3m_goods': 0, 'sales_3m_clothing': 0}
        merged[name]['sales_3m_clothing'] = total

    # ── 1단계: 캐시 확인, 새로 분석이 필요한 대리점만 모으기 ──
    analysis = {}  # name -> {'keywords', 'comment', 'assessed_risk', 'ai_judged'}
    to_call = []   # 캐시 미스: [{'name', 'memo', 'debt_info'}, ...]
    for name, data in merged.items():
        memo = ' / '.join(data['memos'])
        if not memo:
            continue
        debt_info = store_debt_map.get(name, {})
        rule_keywords = ', '.join(extract_keywords_rule_based(memo))
        if api_key:
            h = memo_hash(memo)
            cached = ai_cache.get(name)
            if cached and cached.get('memo_hash') == h and cached.get('assessed_risk') in RISK_ORDER:
                analysis[name] = {
                    'keywords': rule_keywords,
                    'comment': cached.get('comment', ''),
                    'assessed_risk': cached.get('assessed_risk', ''),
                    'ai_judged': True,
                }
                cache_hits += 1
                print(f"  {name} AI 분석 (캐시 재사용)")
                continue
        to_call.append({'name': name, 'memo': memo, 'debt_info': debt_info, 'rule_keywords': rule_keywords})

    # ── 2단계: 캐시 미스 대리점을 BATCH_SIZE씩 묶어 Gemini 호출 (RPD 20 한도 대응) ──
    for i in range(0, len(to_call), BATCH_SIZE):
        batch = to_call[i:i + BATCH_SIZE]
        if not api_key:
            break
        cache_calls += 1
        batch_results = gemini_analyze_batch(
            [{'name': s['name'], 'memo': s['memo'], 'debt_info': s['debt_info']} for s in batch],
            api_key,
        )
        for s in batch:
            name = s['name']
            if name in batch_results:
                keywords, comment, assessed_risk = batch_results[name]
                analysis[name] = {
                    'keywords': s['rule_keywords'],  # 감점 근거 키워드는 항상 규칙 기반 추출값 사용
                    'comment': comment,
                    'assessed_risk': assessed_risk,
                    'ai_judged': assessed_risk in RISK_ORDER,
                }
                print(f"  {name} AI 분석: Gemini 성공 (배치)")
        if not _quota_exhausted and i + BATCH_SIZE < len(to_call):
            time.sleep(13)  # RPM 5 안전 마진 확보 (배치 단위이므로 호출 자체는 훨씬 적음)

    # ── 3단계: 점수 계산 ──
    result = {}
    for name, data in merged.items():
        memo      = ' / '.join(data['memos'])
        p_goods   = data['p_goods']
        p_clothing = data['p_clothing']
        sales_3m_goods    = data.get('sales_3m_goods', 0)
        sales_3m_clothing = data.get('sales_3m_clothing', 0)

        debt_info = store_debt_map.get(name, {})
        mechanical_risk = debt_info.get('risk', '')

        if memo:
            a = analysis.get(name)
            rule_keywords_list = extract_keywords_rule_based(memo)
            rule_keywords = ', '.join(rule_keywords_list)
            keywords = rule_keywords

            if a:
                comment, assessed_risk, ai_judged = a['comment'], a['assessed_risk'], a['ai_judged']
            else:
                comment, assessed_risk, ai_judged = '', '', False
                print(f"  {name} AI 분석: Gemini 미판단 (쿼터 소진 등)")

            # CS 점수 산정용 위험단계: Gemini가 실제로 메모를 읽고 판단했으면 그 결과를 그대로 사용.
            # Gemini가 실패/미판단이면 점수만은 키워드 기반으로 대체 추정 (표시용 assessed_risk와는 별개).
            cs_risk = assessed_risk if ai_judged else classify_cs_risk_fallback(memo)

            if not comment:
                comment = generate_rule_based_comment(name, memo, rule_keywords_list, mechanical_risk, debt_info)
                if not assessed_risk:
                    # 표시(미스매치 비교)용 assessed_risk는 기존과 동일하게 채권등급으로 대체
                    assessed_risk = mechanical_risk if mechanical_risk in RISK_ORDER else ''
                print(f"  {name} AI 분석: 규칙 기반 대체 생성 사용")

            if api_key:
                ai_cache[name] = {'memo_hash': memo_hash(memo), 'keywords': rule_keywords, 'comment': comment, 'assessed_risk': assessed_risk if ai_judged else ''}
        else:
            keywords, comment, assessed_risk, cs_risk = '', '', '', ''

        # AI 자체 판단과 기계적 등급(classify_risk) 불일치 체크
        # (관리/해당없음 등 4단계 스케일 밖의 등급은 비교 대상에서 제외)
        ai_mismatch = False
        ai_mismatch_direction = ''  # 'severe'=AI가 더 심각하게 판단, 'mild'=AI가 더 완화해서 판단
        if assessed_risk and mechanical_risk in RISK_ORDER:
            if assessed_risk != mechanical_risk:
                ai_mismatch = True
                if RISK_ORDER.index(assessed_risk) > RISK_ORDER.index(mechanical_risk):
                    ai_mismatch_direction = 'severe'
                else:
                    ai_mismatch_direction = 'mild'

        # CS 점수 (20점 만점, 메모 없으면 20점)
        cs_score = score_cs(memo, cs_risk)

        # 파트너십 점수 (30점 만점)
        partnership_score = score_partnership(p_goods, p_clothing)

        # 매출규모 점수 (3개월 합계, 용품10+의류10=20점 만점)
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
            'keywords':               keywords,
            'memo':                   memo,
            'ai_comment':             comment,
            'ai_assessed_risk':       assessed_risk,
            'ai_mismatch':            ai_mismatch,
            'ai_mismatch_direction':  ai_mismatch_direction,
        }
        if ai_mismatch:
            arrow = '⬆️더 심각' if ai_mismatch_direction == 'severe' else '⬇️더 완화'
            print(f"  ⚠️ {name}: AI판단({assessed_risk}) vs 기계적등급({mechanical_risk}) 불일치 {arrow}")
        print(f"  {name}: CS {cs_score}점 / 파트너십 {partnership_score}점 완료")

    save_ai_cache(ai_cache)
    print(f"  📦 AI 캐시: 재사용 {cache_hits}건 / 신규 배치 호출 {cache_calls}건 (대리점 {len(to_call)}곳을 배치 {BATCH_SIZE}개씩 묶어 처리)")

    return result
