import os
import json
import re
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

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
# 메모 있음 → 20점 기본에서 키워드 감점
# ───────────────────────────────────────────
def score_cs(keywords_str, memo):
    """CS 점수 계산 (20점 만점)"""
    # 메모 없으면 만점
    if not memo or not memo.strip():
        return 20

    text = f"{keywords_str} {memo}".lower()
    score = 20  # 기본 20점

    for kw in RISK_KEYWORDS['high']:
        if kw in text:
            score -= 20   # 고위험 -20점
    for kw in RISK_KEYWORDS['mid']:
        if kw in text:
            score -= 10   # 중위험 -10점
    # 저위험(RISK_KEYWORDS['low'])은 감점 없음

    return max(0, score)  # 중복 감점 누적, 만점(20점) 구조상 최대 감점은 자동으로 -20점(=0점)까지로 제한됨


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


def gemini_analyze(store_name, memo, api_key, debt_info={}):
    """키워드 추출 + AI 자체 위험도 판단 + 분석 보고를 한 번의 API 호출로 처리.
    기계적 등급(classify_risk)은 프롬프트에 알려주지 않고, AI가 원본 숫자만으로 스스로 판단하게 함."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    all_keywords = RISK_KEYWORDS['high'] + RISK_KEYWORDS['mid'] + RISK_KEYWORDS['low']
    keyword_list = ', '.join(all_keywords)

    collateral = debt_info.get('collateral', 0)
    receivable = debt_info.get('receivable', 0)
    ratio      = debt_info.get('ratio', 0)
    ratio_pct  = int(ratio * 100)
    col_str    = f"{collateral:,}원 ({collateral//10000:,}만원)" if collateral else "없음"
    rec_str    = f"{receivable:,}원 ({receivable//10000:,}만원)" if receivable else "없음"
    excess     = receivable - collateral
    excess_str = f"{excess:,}원 ({excess//10000:,}만원)" if excess > 0 else "초과 없음"

    prompt = f"""당신은 스포츠용품 유통사의 대리점 채권·CS 리스크 전담 분석가입니다.
아래 원본 데이터만 보고, 회사가 정한 등급 기준을 참고하여 이 대리점의 위험도를 직접 판단하세요.
(주의: 아래엔 최종 위험단계를 알려주지 않습니다 — 반드시 스스로 판단하세요)

[채권 현황 - 원본 수치]
- 담보금액: {col_str}
- 채권잔액: {rec_str}
- 담보초과액: {excess_str}
- 담보대비 채권비율: {ratio_pct}%

[CS 특이사항 메모]
{memo}

[참고: 회사 등급 기준 (담보대비 채권비율)]
- 적정: 60% 이하 / 주의: 60~100% / 경계: 100~150% / 위기: 150% 초과
- 단, 이 기준은 참고용이며 CS 메모의 심각도(연락두절, 폐업징후 등)를 종합적으로 고려해 기계적 기준보다 더 엄격하게(또는 완화해서) 판단해도 됩니다. 판단 근거를 반드시 밝히세요.

작업1) 위 메모에 해당하는 키워드를 아래 목록에서만 골라 쉼표로 구분해 추출 (해당 없으면 빈 문자열)
사용 가능한 키워드 목록: {keyword_list}

작업2) 당신이 직접 판단한 위험단계를 "적정","주의","경계","위기" 중 하나로 선택

작업3) 영업팀 관리자가 즉시 활용할 수 있는 전문 분석 보고 작성:
🔴 리스크 요인: (채권 수치와 CS 이슈를 근거로 구체적 위험 요소 서술. 수치 반드시 포함.)
🟢 긍정 요인: (안정적 요소나 완화 요인이 있으면 서술. 없으면 "해당 없음".)
📋 권고사항: (담당 영업사원이 취해야 할 즉각적 조치를 1~2가지 구체적으로 제시.)
- 각 항목은 1~2문장으로 간결하되 수치 근거 포함, 불필요한 인사말/서론 없이 바로 본문 시작
- 만약 작업2에서 기계적 기준(비율 구간)과 다르게 판단했다면, 리스크 요인 첫 문장에 왜 다르게 판단했는지 명시

반드시 아래 JSON 형식으로만 출력하세요. 다른 텍스트 일절 금지:
{{"keywords": "쉼표로 구분된 키워드 또는 빈 문자열", "assessed_risk": "적정|주의|경계|위기 중 하나", "comment": "작업3 분석 보고 전문"}}"""

    body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": 700},
    }
    try:
        res = requests.post(url, json=body, timeout=30)
        data = res.json()
        candidates = data.get('candidates', [])
        if not candidates:
            return '', '', ''
        raw = candidates[0]['content']['parts'][0]['text'].strip()
        raw = raw.strip('`').replace('json\n', '', 1).strip()
        parsed = json.loads(raw)
        keywords = parsed.get('keywords', '').strip()
        valid = [k.strip() for k in keywords.split(',') if k.strip() in all_keywords]
        comment = parsed.get('comment', '').strip()
        assessed_risk = parsed.get('assessed_risk', '').strip()
        if assessed_risk not in RISK_ORDER:
            assessed_risk = ''
        return ', '.join(valid), comment, assessed_risk
    except Exception as e:
        print(f"Gemini 오류 ({store_name}): {e}")
        return '', '', ''


# ───────────────────────────────────────────
# 메인: CS 데이터 fetch
# ───────────────────────────────────────────
def fetch_cs_data(store_debt_map={}):
    api_key = os.environ.get('GEMINI_API_KEY', '')

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

    result = {}
    for name, data in merged.items():
        memo      = ' / '.join(data['memos'])
        p_goods   = data['p_goods']
        p_clothing = data['p_clothing']
        sales_3m_goods    = data.get('sales_3m_goods', 0)
        sales_3m_clothing = data.get('sales_3m_clothing', 0)

        # AI 키워드 추출 + 리스크 분석 (메모 있을 때만, 1회 호출로 통합)
        debt_info = store_debt_map.get(name, {})
        mechanical_risk = debt_info.get('risk', '')
        if api_key and memo:
            keywords, comment, assessed_risk = gemini_analyze(name, memo, api_key, debt_info)
            print(f"  {name} 키워드 추출: {keywords if keywords else '없음'}")
        else:
            keywords, comment, assessed_risk = '', '', ''

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
        cs_score = score_cs(keywords, memo)

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

    return result
