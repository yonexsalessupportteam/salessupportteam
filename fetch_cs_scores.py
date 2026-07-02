import os
import json
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
            score -= 5   # 고위험 -5점
    for kw in RISK_KEYWORDS['mid']:
        if kw in text:
            score -= 3   # 중위험 -3점

    return max(0, score)


def score_partnership(p_goods, p_clothing):
    """파트너십 점수 (20점 만점). 컬럼에 값이 있으면(공백 아니면) 위반으로 처리."""
    goods_score = 0 if str(p_goods).strip() else 10
    cloth_score = 0 if str(p_clothing).strip() else 10
    return goods_score + cloth_score


# ───────────────────────────────────────────
# AI 키워드 자동 추출 (Gemini)
# ───────────────────────────────────────────
def extract_keywords_with_gemini(store_name, memo, api_key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    all_keywords = RISK_KEYWORDS['high'] + RISK_KEYWORDS['mid'] + RISK_KEYWORDS['low']
    keyword_list = ', '.join(all_keywords)

    prompt = f"""아래는 대리점 CS 특이사항 메모입니다.
메모 내용을 분석하여 해당하는 키워드만 골라 쉼표로 구분해서 출력하세요.

사용 가능한 키워드 목록: {keyword_list}

메모: {memo}

규칙:
- 반드시 위 키워드 목록에 있는 것만 사용
- 메모에 해당하는 키워드만 골라서 쉼표로 구분
- 키워드 외 다른 텍스트 일절 출력 금지
- 해당 키워드 없으면 빈 문자열 출력"""

    body = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        res = requests.post(url, json=body, timeout=15)
        data = res.json()
        candidates = data.get('candidates', [])
        if candidates:
            keywords = candidates[0]['content']['parts'][0]['text'].strip()
            valid = [k.strip() for k in keywords.split(',') if k.strip() in all_keywords]
            return ', '.join(valid)
        return ''
    except Exception as e:
        print(f"키워드 추출 오류 ({store_name}): {e}")
        return ''


def gemini_analyze(store_name, keywords, memo, api_key, debt_info={}):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    collateral = debt_info.get('collateral', 0)
    receivable = debt_info.get('receivable', 0)
    ratio      = debt_info.get('ratio', 0)
    risk       = debt_info.get('risk', '정보없음')
    ratio_pct  = int(ratio * 100)
    col_str    = f"{collateral:,}원 ({collateral//10000:,}만원)" if collateral else "없음"
    rec_str    = f"{receivable:,}원 ({receivable//10000:,}만원)" if receivable else "없음"
    excess     = receivable - collateral
    excess_str = f"{excess:,}원 ({excess//10000:,}만원)" if excess > 0 else "초과 없음"

    prompt = f"""당신은 스포츠용품 유통사의 대리점 채권·CS 리스크 전담 분석가입니다.
아래 데이터를 바탕으로 영업팀 관리자가 즉시 활용할 수 있는 전문 분석 보고를 작성해주세요.

[채권 현황]
- 담보금액: {col_str}
- 채권잔액: {rec_str}
- 담보초과액: {excess_str}
- 담보대비 채권비율: {ratio_pct}%
- 위험단계: {risk}

[CS 현황]
- 키워드: {keywords if keywords else '없음'}
- 특이사항 메모: {memo if memo else '없음'}

다음 형식으로 반드시 작성하세요:

🔴 리스크 요인: (채권 수치와 CS 이슈를 근거로 구체적 위험 요소 서술. 수치 반드시 포함.)
🟢 긍정 요인: (안정적 요소나 완화 요인이 있으면 서술. 없으면 "해당 없음".)
📋 권고사항: (담당 영업사원이 취해야 할 즉각적 조치를 1~2가지 구체적으로 제시.)

조건:
- 각 항목은 1~2문장으로 간결하되 수치 근거 포함
- 불필요한 인사말, 서론 없이 바로 본문 시작"""

    body = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        res = requests.post(url, json=body, timeout=30)
        data = res.json()
        candidates = data.get('candidates', [])
        if candidates:
            return candidates[0]['content']['parts'][0]['text'].strip()
        return ''
    except Exception as e:
        print(f"Gemini 오류: {e}")
        return ''


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
    skipped_old = 0

    # CS 시트 병합 (대리점명 기준, 이번 달 작성건만)
    merged = {}
    for row in records:
        name      = ' '.join(str(row.get('대리점명', '')).split())
        memo      = str(row.get('특이사항메모', '')).strip()
        p_goods   = str(row.get('파트너십_용품', '')).strip()
        p_cloth   = str(row.get('파트너십_의류', '')).strip()
        written   = parse_sheet_date(row.get('작성일', ''))
        if not name:
            continue
        # 작성일이 있는데 이번 달이 아니면 제외 (과거 메모 누적 방지)
        if written and (written.year != cur_year or written.month != cur_month):
            skipped_old += 1
            continue
        if name not in merged:
            merged[name] = {'memos': [], 'p_goods': '', 'p_clothing': ''}
        if memo:
            merged[name]['memos'].append(memo)
        if p_goods:
            merged[name]['p_goods'] = p_goods
        if p_cloth:
            merged[name]['p_clothing'] = p_cloth

    if skipped_old:
        print(f"  이번 달({cur_year}-{cur_month:02d}) 이전 작성 건 {skipped_old}개 제외")

    result = {}
    for name, data in merged.items():
        memo      = ' / '.join(data['memos'])
        p_goods   = data['p_goods']
        p_clothing = data['p_clothing']

        # AI 키워드 자동 추출 (메모 있을 때만)
        if api_key and memo:
            keywords = extract_keywords_with_gemini(name, memo, api_key)
            print(f"  {name} 키워드 추출: {keywords if keywords else '없음'}")
        else:
            keywords = ''

        # CS 점수 (20점 만점, 메모 없으면 20점)
        cs_score = score_cs(keywords, memo)

        # 파트너십 점수 (20점 만점)
        partnership_score = score_partnership(p_goods, p_clothing)

        # Gemini AI 분석 (메모 있을 때만)
        debt_info = store_debt_map.get(name, {})
        if api_key and memo:
            comment = gemini_analyze(name, keywords, memo, api_key, debt_info)
        else:
            comment = ''

        result[name] = {
            'score':             cs_score,
            'partnership_score': partnership_score,
            'p_goods':           p_goods,
            'p_clothing':        p_clothing,
            'keywords':          keywords,
            'memo':              memo,
            'ai_comment':        comment,
        }
        print(f"  {name}: CS {cs_score}점 / 파트너십 {partnership_score}점 완료")

    return result
