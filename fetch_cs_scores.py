import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
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
def fetch_sheet_data():
    spreadsheet_id = os.environ.get('SPREADSHEET_ID', '')
    client = get_sheets_client()
    sheet = client.open_by_key(spreadsheet_id).sheet1
    records = sheet.get_all_records()
    return records
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
def gemini_analyze(store_name, keywords, memo, api_key, debt_info={}):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
    collateral = debt_info.get('collateral', 0)
    receivable = debt_info.get('receivable', 0)
    ratio = debt_info.get('ratio', 0)
    risk = debt_info.get('risk', '정보없음')
    ratio_pct = int(ratio * 100)
    col_str = f"{collateral//10000}백만" if collateral else "없음"
    rec_str = f"{receivable//10000}백만" if receivable else "없음"
    prompt = f"""당신은 스포츠용품 대리점 채권·CS 리스크 분석 전문가입니다.
아래 데이터를 바탕으로 실무적인 리스크 분석을 작성해주세요.

[채권 현황]
- 담보금액: {col_str} / 채권잔액: {rec_str} / 초과율: {ratio_pct}% / 위험단계: {risk}

[CS 현황]
- 키워드: {keywords}
- 특이사항: {memo}

담보 수치와 CS 이슈를 함께 언급하며 2~3문장으로 간결하게 작성하세요.
구체적 숫자를 포함하고 마지막엔 조치사항을 제시하세요.
예시: "담보(180백만) 대비 채권(342백만)이 187% 초과 상태로 담보가 부족합니다. 채권 위기단계와 CS 클레임 반복, 담당자 연락두절이 복합적으로 작용하고 있어 즉각적인 현장 방문 및 담보 추가설정 검토가 필요합니다." """
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

    # 같은 대리점 여러 행 → 키워드/메모 합산
    merged = {}
    for row in records:
        name = ' '.join(str(row.get('대리점명', '')).split())
        keywords = str(row.get('키워드', '')).strip()
        memo = str(row.get('특이사항메모', '')).strip()
        if not name:
            continue
        if name not in merged:
            merged[name] = {'keywords': [], 'memos': []}
        if keywords:
            merged[name]['keywords'].append(keywords)
        if memo:
            merged[name]['memos'].append(memo)

    result = {}
    for name, data in merged.items():
        keywords = ', '.join(data['keywords'])
        memo = ' / '.join(data['memos'])
        score = score_from_keywords(keywords, memo)
        debt_info = store_debt_map.get(name, {})
        if api_key:
            comment = gemini_analyze(name, keywords, memo, api_key, debt_info)
        else:
            comment = f"키워드 분석: {keywords}"
        result[name] = {
            'score': score,
            'keywords': keywords,
            'memo': memo,
            'ai_comment': comment
        }
        print(f"  {name}: {score}점 분석 완료")
    return result
