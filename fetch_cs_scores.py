import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets'
]

RISK_KEYWORDS = {
    'high': ['연락두절', '약속불이행', '클레임다발', '폐업징후', '허위접수', '재고과다', '타사이탈', '무단온라인', '연락안됨', '잠수'],
    'mid':  ['연락지연', '가끔약속어김', '클레임', '재고증가', '매출감소', '응대느림', '불만', 'A/S규정숙지', '연락두절'],
    'low':  ['협조적', '응대원활', '약속이행', '클레임없음', '재고적정', '매출안정', '신뢰', '칭찬']
}

# ───────────────────────────────────────────
# 매출규모 배점 기준
# ───────────────────────────────────────────
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


def get_sheet():
    spreadsheet_id = os.environ.get('SPREADSHEET_ID', '')
    client = get_sheets_client()
    return client.open_by_key(spreadsheet_id).sheet1


def fetch_sheet_data():
    return get_sheet().get_all_records()


# ───────────────────────────────────────────
# 매출규모 스프레드시트 읽기
# 구조: A(매장코드) B(매장명) C(담당자) D~O(25.06~26.05 월별매출)
#       용품/의류 각각 별도 시트
# ───────────────────────────────────────────
def _parse_sales_sheet(ws, category):
    """
    시트에서 직전 12개월 매출 합산 후 상/중/하 배점 반환
    반환: { '매장명': {'매출': int, '규모': str, '배점': int} }
    """
    all_values = ws.get_all_values()
    if len(all_values) < 2:
        return {}

    header = all_values[0]  # 1행: 매장, 매장명, 담당자, 25.06, 25.07, ...

    # 직전 12개월 범위 (오늘 KST 기준)
    today = datetime.utcnow() + timedelta(hours=9)
    target_months = set()
    y, m = today.year, today.month
    for _ in range(12):
        m -= 1
        if m == 0:
            m = 12
            y -= 1
        target_months.add(f"{y}.{m:02d}")

    # 헤더에서 해당 월 컬럼 인덱스 추출
    # 헤더 형식: '25.06', '25.07' ... → '2025.06' 으로 변환해서 매칭
    month_col_indices = []
    for idx, h in enumerate(header):
        h = str(h).strip()
        # '25.06' 형태 → '2025.06'
        parts = h.split('.')
        if len(parts) == 2 and len(parts[0]) == 2 and parts[1].isdigit():
            normalized = f"20{parts[0]}.{parts[1].zfill(2)}"
            if normalized in target_months:
                month_col_indices.append(idx)

    thresholds = SALES_GRADE[category]

    def get_grade(sales):
        if sales >= thresholds['상']:
            return '상'
        elif sales >= thresholds['중']:
            return '중'
        else:
            return '하'

    result = {}
    for row in all_values[1:]:
        if not row or not str(row[1]).strip():
            continue
        name = str(row[1]).strip()  # B열: 매장명

        total = 0
        for idx in month_col_indices:
            if idx < len(row):
                try:
                    val = str(row[idx]).replace(',', '').strip()
                    total += int(val) if val else 0
                except ValueError:
                    pass

        grade = get_grade(total)
        result[name] = {
            '매출': total,
            '규모': grade,
            '배점': SALES_SCORE[grade],
        }
        print(f"  [{category}] {name}: {total:,}원 → {grade} {SALES_SCORE[grade]}점")

    return result


def fetch_sales_grade_data():
    """
    매출규모 전용 스프레드시트(용품/의류 시트 별도)에서 배점 반환
    반환: { '매장명': {'용품_규모': str, '용품_배점': int, '의류_규모': str, '의류_배점': int} }
    """
    sales_id = os.environ.get('SALES_SPREADSHEET_ID', '')
    if not sales_id:
        print("SALES_SPREADSHEET_ID 환경변수 없음 — 매출규모 배점 생략")
        return {}

    try:
        client = get_sheets_client()
        sh = client.open_by_key(sales_id)
    except Exception as e:
        print(f"매출규모 스프레드시트 열기 실패: {e}")
        return {}

    # 용품 시트
    try:
        goods_ws = sh.worksheet('용품')
        goods_data = _parse_sales_sheet(goods_ws, '용품')
    except Exception as e:
        print(f"용품 시트 읽기 실패: {e}")
        goods_data = {}

    # 의류 시트
    try:
        clothing_ws = sh.worksheet('의류')
        clothing_data = _parse_sales_sheet(clothing_ws, '의류')
    except Exception as e:
        print(f"의류 시트 읽기 실패: {e}")
        clothing_data = {}

    # 매장명 기준으로 합산
    all_names = set(goods_data.keys()) | set(clothing_data.keys())
    result = {}
    for name in all_names:
        g = goods_data.get(name, {'규모': '하', '배점': 5})
        c = clothing_data.get(name, {'규모': '하', '배점': 5})
        result[name] = {
            '용품_규모': g['규모'],
            '용품_배점': g['배점'],
            '의류_규모': c['규모'],
            '의류_배점': c['배점'],
        }

    return result


# ───────────────────────────────────────────
# AI 키워드 자동 추출 (Gemini)
# ───────────────────────────────────────────
def extract_keywords_with_gemini(store_name, memo, api_key):
    """메모 텍스트에서 CS 리스크 키워드를 AI가 자동 추출"""
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
            # 키워드 목록에 있는 것만 필터링
            valid = [k.strip() for k in keywords.split(',')
                     if k.strip() in all_keywords]
            return ', '.join(valid)
        return ''
    except Exception as e:
        print(f"키워드 추출 오류 ({store_name}): {e}")
        return ''


def score_from_keywords(keywords_str, memo):
    text = f"{keywords_str} {memo}".lower()
    score = 50
    for kw in RISK_KEYWORDS['high']:
        if kw in text:
            score -= 15
    for kw in RISK_KEYWORDS['mid']:
        if kw in text:
            score -= 7
    for kw in RISK_KEYWORDS['low']:
        if kw in text:
            score += 10
    return max(0, min(100, score))


def score_partnership(p_goods, p_clothing):
    goods_score  = 0 if str(p_goods).strip()    in ['1', 'TRUE', 'true', '미준수', 'Y', 'y'] else 10
    cloth_score  = 0 if str(p_clothing).strip()  in ['1', 'TRUE', 'true', '미준수', 'Y', 'y'] else 10
    return goods_score + cloth_score


def gemini_analyze(store_name, keywords, memo, api_key, debt_info={}):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    collateral = debt_info.get('collateral', 0)
    receivable = debt_info.get('receivable', 0)
    ratio = debt_info.get('ratio', 0)
    risk = debt_info.get('risk', '정보없음')
    ratio_pct = int(ratio * 100)

    col_str = f"{collateral:,}원 ({collateral//10000:,}만원)" if collateral else "없음"
    rec_str = f"{receivable:,}원 ({receivable//10000:,}만원)" if receivable else "없음"
    excess = receivable - collateral
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

다음 형식으로 반드시 작성하세요 (각 항목은 줄바꿈으로 구분):

🔴 리스크 요인: (채권 수치와 CS 이슈를 근거로 구체적 위험 요소 서술. 수치 반드시 포함.)
🟢 긍정 요인: (안정적 요소나 완화 요인이 있으면 서술. 없으면 "해당 없음".)
📋 권고사항: (담당 영업사원이 취해야 할 즉각적·단기적 조치를 1~2가지 구체적으로 제시.)

조건:
- 각 항목은 1~2문장으로 간결하되 수치 근거를 포함할 것
- 실무자가 바로 행동할 수 있는 수준의 구체성
- 불필요한 인사말, 서론 없이 바로 본문 시작"""

    body = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        res = requests.post(url, json=body, timeout=30)
        data = res.json()
        candidates = data.get('candidates', [])
        if candidates:
            return candidates[0]['content']['parts'][0]['text'].strip()
        return f"키워드 [{keywords}] 기반 분석: 복합 이슈 확인 필요"
    except Exception as e:
        print(f"Gemini 오류: {e}")
        return f"키워드 [{keywords}] 기반 분석: 복합 이슈 확인 필요"


# ───────────────────────────────────────────
# 메인: CS 데이터 fetch
# ───────────────────────────────────────────
def fetch_cs_data(store_debt_map={}):
    api_key = os.environ.get('GEMINI_API_KEY', '')

    # CS 평가 시트
    try:
        records = fetch_sheet_data()
    except Exception as e:
        print(f"구글 시트 읽기 실패: {e}")
        return {}

    # 매출규모 시트
    print("\n매출규모 시트 읽는 중...")
    sales_data = fetch_sales_grade_data()

    # CS 시트 병합
    merged = {}
    for row in records:
        name = ' '.join(str(row.get('대리점명', '')).split())
        memo = str(row.get('특이사항메모', '')).strip()
        p_goods    = str(row.get('파트너십_용품', '')).strip()
        p_clothing = str(row.get('파트너십_의류', '')).strip()
        if not name:
            continue
        if name not in merged:
            merged[name] = {'memos': [], 'p_goods': '', 'p_clothing': ''}
        if memo:
            merged[name]['memos'].append(memo)
        if p_goods:
            merged[name]['p_goods'] = p_goods
        if p_clothing:
            merged[name]['p_clothing'] = p_clothing

    result = {}
    for name, data in merged.items():
        memo = ' / '.join(data['memos'])
        p_goods = data['p_goods']
        p_clothing = data['p_clothing']

        # AI 키워드 자동 추출
        if api_key and memo:
            keywords = extract_keywords_with_gemini(name, memo, api_key)
            print(f"  {name} 키워드 추출: {keywords if keywords else '없음'}")
        else:
            keywords = ''

        cs_score = score_from_keywords(keywords, memo)
        partnership_score = score_partnership(p_goods, p_clothing)

        debt_info = store_debt_map.get(name, {})

        if api_key and memo:
            comment = gemini_analyze(name, keywords, memo, api_key, debt_info)
        else:
            comment = ''

        # 매출규모 배점
        sales_info = sales_data.get(name, {})

        result[name] = {
            'score': cs_score,
            'partnership_score': partnership_score,
            'p_goods': p_goods,
            'p_clothing': p_clothing,
            'keywords': keywords,
            'memo': memo,
            'ai_comment': comment,
            # 정량평가 — 매출규모
            'goods_sales_grade':  sales_info.get('용품_규모', ''),
            'goods_sales_score':  sales_info.get('용품_배점', 0),
            'clothing_sales_grade': sales_info.get('의류_규모', ''),
            'clothing_sales_score': sales_info.get('의류_배점', 0),
        }
        print(f"  {name}: CS {cs_score}점 / 파트너십 {partnership_score}점 완료")

    return result
