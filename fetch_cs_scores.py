import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials

# 읽기+쓰기 권한
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets'
]

RISK_KEYWORDS = {
    'high': ['연락두절', '약속불이행', '클레임다발', '폐업징후', '허위접수', '재고과다', '타사이탈', '무단온라인', '연락안됨', '잠수'],
    'mid':  ['연락지연', '가끔약속어김', '클레임', '재고증가', '매출감소', '응대느림', '불만'],
    'low':  ['협조적', '응대원활', '약속이행', '클레임없음', '재고적정', '매출안정', '신뢰']
}

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
    """본사 파트너십 점수 (20점 만점): 용품 10점 + 의류 10점, 미준수 시 0점"""
    goods_score  = 0  if str(p_goods).strip()    in ['1', 'TRUE', 'true', '미준수', 'Y', 'y'] else 10
    cloth_score  = 0  if str(p_clothing).strip()  in ['1', 'TRUE', 'true', '미준수', 'Y', 'y'] else 10
    return goods_score + cloth_score

def gemini_analyze(store_name, keywords, memo, api_key, debt_info={}):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
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
        res = requests.post(url, json=body, timeout=15)
        data = res.json()
        candidates = data.get('candidates', [])
        if candidates:
            return candidates[0]['content']['parts'][0]['text'].strip()
        print(f"Gemini 전체 응답: {data}")
        return f"키워드 [{keywords}] 기반 분석: 복합 이슈 확인 필요"
    except Exception as e:
        print(f"Gemini 오류: {e}")
        return f"키워드 [{keywords}] 기반 분석: 복합 이슈 확인 필요"

def fetch_cs_data(store_debt_map={}):
    api_key = os.environ.get('GEMINI_API_KEY', '')
    try:
        records = fetch_sheet_data()
    except Exception as e:
        print(f"구글 시트 읽기 실패: {e}")
        return {}

    # 같은 대리점 여러 행 → 키워드/메모/파트너십 합산
    merged = {}
    for row in records:
        name = ' '.join(str(row.get('대리점명', '')).split())
        keywords = str(row.get('키워드', '')).strip()
        memo = str(row.get('특이사항메모', '')).strip()
        p_goods   = str(row.get('파트너십_용품', '')).strip()
        p_clothing = str(row.get('파트너십_의류', '')).strip()
        if not name:
            continue
        if name not in merged:
            merged[name] = {'keywords': [], 'memos': [], 'p_goods': '', 'p_clothing': ''}
        if keywords:
            merged[name]['keywords'].append(keywords)
        if memo:
            merged[name]['memos'].append(memo)
        # 한 번이라도 미준수 체크되면 미준수로 처리
        if p_goods:
            merged[name]['p_goods'] = p_goods
        if p_clothing:
            merged[name]['p_clothing'] = p_clothing

    result = {}
    for name, data in merged.items():
        keywords = ', '.join(data['keywords'])
        memo = ' / '.join(data['memos'])
        p_goods = data['p_goods']
        p_clothing = data['p_clothing']

        cs_score = score_from_keywords(keywords, memo)
        partnership_score = score_partnership(p_goods, p_clothing)

        debt_info = store_debt_map.get(name, {})
        if api_key and (keywords or memo):
            comment = gemini_analyze(name, keywords, memo, api_key, debt_info)
        elif keywords or memo:
            comment = f"키워드 분석: {keywords}"
        else:
            comment = ""

        result[name] = {
            'score': cs_score,
            'partnership_score': partnership_score,
            'p_goods': p_goods,
            'p_clothing': p_clothing,
            'keywords': keywords,
            'memo': memo,
            'ai_comment': comment
        }
        print(f"  {name}: CS {cs_score}점 / 파트너십 {partnership_score}점 분석 완료")
    return result
