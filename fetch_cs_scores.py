"""
구글 시트에서 CS 평가 데이터를 읽어 대리점별 점수를 계산하는 모듈

평가기준:
- 기본점수: 50점
- 가점 (각 +10점): 항목 1~5번
- 감점 (각 -5점): 항목 6~15번
- 등급: S(90~100) / A(70~89) / B(50~69) / C(30~49) / D(0~29)
"""

import os
import json
import gspread
from google.oauth2.service_account import Credentials

# 평가항목 번호별 분류
GAIN_ITEMS = {1, 2, 3, 4, 5}   # 가점 (+10점)
LOSS_ITEMS = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15}  # 감점 (-5점)

# 감점 시 등급별 카운트 기준
LOSS_COUNT_THRESHOLD = {
    'A': 5,  # A등급: 5회 이상 카운트시 -1회
    'B': 3,  # B등급: 3회 이상 카운트시 -1회
    'C': 1,  # C등급: 1회 이상 카운트시 -1회
}

def get_cs_grade(score):
    """점수 → CS 등급 변환"""
    if score >= 90: return 'S'
    elif score >= 70: return 'A'
    elif score >= 50: return 'B'
    elif score >= 30: return 'C'
    else: return 'D'

def get_cs_score_value(grade):
    """CS 등급 → 영업 점수 변환"""
    mapping = {'S': 5, 'A': 4, 'B': 3, 'C': 2, 'D': 1}
    return mapping.get(grade, 3)

def fetch_cs_data():
    """구글 시트에서 CS 평가 데이터를 읽어 대리점별 점수 계산"""
    try:
        creds_json = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        spreadsheet_id = os.environ.get('SPREADSHEET_ID')

        if not creds_json or not spreadsheet_id:
            print("구글 시트 환경변수 없음 - CS 점수 기본값 사용")
            return {}

        creds_dict = json.loads(creds_json)
        scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)

        sheet = client.open_by_key(spreadsheet_id)
        ws = sheet.worksheet('데이터입력')
        rows = ws.get_all_records()

        # 대리점별 가점/감점 집계
        store_data = {}
        for row in rows:
            name = str(row.get('대리점명', '')).strip()
            grade = str(row.get('대리점등급', 'B')).strip().upper()
            item_no = int(row.get('평가번호', 0))
            count = int(row.get('건수', 1))

            if not name or item_no == 0:
                continue

            if name not in store_data:
                store_data[name] = {
                    'grade': grade,
                    'gain_counts': {},   # 가점 항목별 카운트
                    'loss_counts': {},   # 감점 항목별 카운트
                }

            if item_no in GAIN_ITEMS:
                store_data[name]['gain_counts'][item_no] = \
                    store_data[name]['gain_counts'].get(item_no, 0) + count
            elif item_no in LOSS_ITEMS:
                store_data[name]['loss_counts'][item_no] = \
                    store_data[name]['loss_counts'].get(item_no, 0) + count

        # 대리점별 최종 점수 계산
        cs_scores = {}
        for name, data in store_data.items():
            grade = data['grade']
            base_score = 50

            # 가점 계산 (항목당 카운트 1회 이상이면 +10점)
            gain_score = 0
            for item_no, cnt in data['gain_counts'].items():
                if cnt >= 1:
                    gain_score += 10

            # 감점 계산 (등급별 카운트 기준 초과시 -5점)
            threshold = LOSS_COUNT_THRESHOLD.get(grade, 3)
            loss_score = 0
            for item_no, cnt in data['loss_counts'].items():
                if item_no == 11:  # 본사 응대 태도는 1회도 무조건 감점
                    if cnt >= 1:
                        loss_score -= 5
                else:
                    if cnt >= threshold:
                        loss_score -= 5

            final_score = max(0, min(100, base_score + gain_score + loss_score))
            cs_grade = get_cs_grade(final_score)
            cs_score_value = get_cs_score_value(cs_grade)

            cs_scores[name] = {
                'score': final_score,
                'grade': cs_grade,
                'score_value': cs_score_value,
                'gain': gain_score,
                'loss': loss_score,
            }

        print(f"CS 점수 계산 완료: {len(cs_scores)}개 대리점")
        return cs_scores

    except Exception as e:
        print(f"구글 시트 연동 오류: {e}")
        return {}
