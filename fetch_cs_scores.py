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

def gemini_analyze(store_name, keywords, memo, api_key):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
    prompt = f"""
당신은 스포츠용품 대리점 CS 리스크 분석 전문가입니다.
아래 대리점 정보를 바탕으로 CS 위험도를 분석해주세요.

대리점명: {store_name}
키워드: {keywords}
특이사항: {memo}

다음 형식으로 간결하게 분석해주세요 (3문장 이내):
- 주요 위험 신호 파악
- 현재 CS 상태 평가
- 권고 조치 사항
"""
    body = {"contents": [{"parts": [{"text": prompt}]}]}
    try:
        res = requests.post(url, json=body, timeout=15)
        data = res.json()
        return data['candidates'][0]['content']['parts'][0]['text'].strip()
    except Exception:
        return f"{keywords} 키워드 기반 분석: 담당자 확인 필요"

def fetch_cs_data():
    api_key = os.environ.get('GEMINI_API_KEY', '')
    try:
        records = fetch_sheet_data()
    except Exception as e:
        print(f"구글 시트 읽기 실패: {e}")
        return {}

    result = {}
    for row in records:
        name = str(row.get('대리점명', '')).strip()
        keywords = str(row.get('키워드', '')).strip()
        memo = str(row.get('특이사항메모', '')).strip()
        if not name:
            continue
        score = score_from_keywords(keywords, memo)
        if api_key:
            comment = gemini_analyze(name, keywords, memo, api_key)
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
