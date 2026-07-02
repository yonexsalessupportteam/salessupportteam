"""
채권 위험도 대시보드 자동 생성 스크립트

총점 구조 (최대 100점):
  정량점수  기본 50점 - 회수일 감점(최대 15) - 담보대비채권 감점(최대 15), 하한 20점
  + CS 코멘트  최대 20점
  + 파트너십   최대 30점 (용품 15 + 의류 15)
  = 최대 100점
"""

import pandas as pd
import numpy as np
import json
import sys
from datetime import datetime, timezone, timedelta
from fetch_cs_scores import fetch_cs_data

RAW_FILES = {
    '의류': 'clothing_raw.xls',
    '용품': 'goods_raw.xls',
}

DEPT_TABS = ['영업1팀', '영업2팀', 'E-BIZ팀']

RISK_THRESHOLDS = {
    'safe_max': 0.6,
    'caution_max': 1.0,
    'warning_max': 1.5,
    'danger_max': 2.0,
}

MIN_RECEIVABLE_THRESHOLD = 500_000
MIN_DISPLAY_THRESHOLD = 100_000


# ───────────────────────────────────────────
# 감점 기준
# ───────────────────────────────────────────

def deduct_collection_days(days):
    """회수일 감점 (60일 기준 초과시 감점, 최대 25점, 카테고리 1개만 있는 대리점용)"""
    try:
        days = float(days)
    except (TypeError, ValueError):
        return 0
    if days <= 60:
        return 0
    elif days <= 80:
        return 10
    elif days <= 90:
        return 15
    else:
        return 25


def deduct_collection_days_half(days):
    """회수일 감점 (최대 15점, 의류+용품 둘 다 있는 대리점의 카테고리별 감점용).
    둘을 합산했을 때 최대 25점(15+15=30→25 캡)이 되도록 절반 스케일."""
    try:
        days = float(days)
    except (TypeError, ValueError):
        return 0
    if days <= 60:
        return 0
    elif days <= 80:
        return 5
    elif days <= 90:
        return 10
    else:
        return 15


def deduct_collateral_ratio(collateral, receivable):
    """담보대비 채권잔액 감점 (최대 25점, 카테고리 1개만 있는 대리점용). classify_risk() 등급 기준(60/100/150%)과 동일하게 정렬."""
    # 무담보 & 채권 없음 → 감점 없음
    if collateral == 0 and receivable <= 0:
        return 0
    # 무담보 & 채권 있음 (=관리 등급) → 최대 감점
    if collateral == 0 and receivable > 0:
        return 25
    ratio = receivable / collateral * 100
    if ratio <= RISK_THRESHOLDS['safe_max'] * 100:      # 적정 (≤60%)
        return 0
    elif ratio <= RISK_THRESHOLDS['caution_max'] * 100:  # 주의 (60~100%)
        return 10
    elif ratio <= RISK_THRESHOLDS['warning_max'] * 100:  # 경계 (100~150%)
        return 15
    else:                                                # 위기 (150% 초과)
        return 25


def deduct_collateral_ratio_half(collateral, receivable):
    """담보대비 채권잔액 감점 (최대 15점, 의류+용품 둘 다 있는 대리점의 카테고리별 감점용).
    둘을 합산했을 때 최대 25점(15+15=30→25 캡)이 되도록 절반 스케일."""
    if collateral == 0 and receivable <= 0:
        return 0
    if collateral == 0 and receivable > 0:
        return 15
    ratio = receivable / collateral * 100
    if ratio <= RISK_THRESHOLDS['safe_max'] * 100:
        return 0
    elif ratio <= RISK_THRESHOLDS['caution_max'] * 100:
        return 5
    elif ratio <= RISK_THRESHOLDS['warning_max'] * 100:
        return 10
    else:
        return 15



def classify_risk(collateral, receivable, ratio):
    if abs(receivable) < MIN_RECEIVABLE_THRESHOLD:
        return '해당없음'
    if collateral == 0 and receivable > 0:
        return '관리'
    if ratio <= RISK_THRESHOLDS['safe_max']:
        return '적정'
    elif ratio <= RISK_THRESHOLDS['caution_max']:
        return '주의'
    elif ratio <= RISK_THRESHOLDS['warning_max']:
        return '경계'
    else:
        return '위기'


def process_raw(filepath):
    xl = pd.read_excel(filepath, engine='xlrd', sheet_name=None, header=None)
    df = xl['export']
    data = df.iloc[1:].copy()
    data.columns = range(len(data.columns))

    mask = data[4].astype(str).str.match(r'^D\d+$', na=False)
    store_data = data[mask].copy()

    result = pd.DataFrame()
    result['code']            = store_data[4].astype(str)
    result['name']            = store_data[5].astype(str)
    result['salesperson']     = store_data[3].astype(str).str.strip()
    result['dept_code']       = store_data[0].astype(str)
    result['dept_name']       = store_data[1].astype(str)
    result['collateral']      = pd.to_numeric(store_data[6],  errors='coerce').fillna(0)
    result['receivable']      = pd.to_numeric(store_data[13], errors='coerce').fillna(0)
    result['sales']           = pd.to_numeric(store_data[11], errors='coerce').fillna(0)
    result['collection']      = pd.to_numeric(store_data[12], errors='coerce').fillna(0)
    result['collection_days'] = pd.to_numeric(store_data[14], errors='coerce').fillna(0)

    result['excess'] = result['receivable'] - result['collateral']
    result['ratio']  = np.where(
        result['collateral'] > 0,
        result['receivable'] / result['collateral'], 0.0
    )
    result['risk'] = result.apply(
        lambda r: classify_risk(r['collateral'], r['receivable'], r['ratio']), axis=1
    )

    # 감점 계산
    result['deduct_collection'] = result['collection_days'].apply(deduct_collection_days)
    result['deduct_collection_half'] = result['collection_days'].apply(deduct_collection_days_half)
    result['deduct_collateral'] = result.apply(
        lambda r: deduct_collateral_ratio(r['collateral'], r['receivable']), axis=1
    )
    result['deduct_collateral_half'] = result.apply(
        lambda r: deduct_collateral_ratio_half(r['collateral'], r['receivable']), axis=1
    )

    return result


def build_group_data(full_sub):
    # 요약(담보/채권 합계 등)은 전체 매장 기준으로 계산
    full_sub = full_sub.copy()

    # 매장 리스트에는 소액(채권 10만원 이하) 매장은 숨김
    sub = full_sub[full_sub['receivable'].abs() > MIN_DISPLAY_THRESHOLD].copy()
    stores = []
    for _, r in sub.iterrows():
        stores.append({
            'code':             r['code'],
            'name':             r['name'],
            'salesperson':      r['salesperson'],
            'collateral':       int(r['collateral']),
            'receivable':       int(r['receivable']),
            'excess':           int(r['excess']),
            'ratio':            round(float(r['ratio']), 2),
            'risk':             r['risk'],
            'collection_days':  int(r['collection_days']),
            'deduct_collection': int(r['deduct_collection']),
            'deduct_collection_half': int(r['deduct_collection_half']),
            'deduct_collateral': int(r['deduct_collateral']),
            'deduct_collateral_half': int(r['deduct_collateral_half']),
        })
    stores = sorted(stores, key=lambda x: -x['receivable'])

    summary = {
        'total_collateral': int(full_sub['collateral'].sum()),
        'total_receivable': int(full_sub['receivable'].sum()),
        'total_excess':     int(full_sub[full_sub['excess'] > 0]['excess'].sum()),
        'risk_counts':      {k: int(v) for k, v in full_sub['risk'].value_counts().to_dict().items()}
    }

    sp_summary = {}
    for sp, g in full_sub.groupby('salesperson'):
        sp_summary[sp] = {
            'collateral': int(g['collateral'].sum()),
            'receivable': int(g['receivable'].sum()),
            'excess':     int(g[g['excess'] > 0]['excess'].sum()),
            'stores':     len(g)
        }
    return {'stores': stores, 'summary': summary, 'by_salesperson': sp_summary}


def build_category_dashboard(filepath):
    result = process_raw(filepath)
    dashboard = {}
    for dept in DEPT_TABS:
        sub = result[result['dept_name'] == dept]
        dashboard[dept] = build_group_data(sub)
    return dashboard


def get_update_timestamp():
    kst = timezone(timedelta(hours=9))
    now = datetime.now(kst)
    weekdays = ['월', '화', '수', '목', '금', '토', '일']
    weekday = weekdays[now.weekday()]
    return f"{now.year}.{now.month:02d}.{now.day:02d}({weekday}) {now.hour:02d}:{now.minute:02d}"


def sanitize_text(text):
    if not text:
        return text
    return (text
            .replace('<', '〈')
            .replace('>', '〉')
            .replace('"', '"')
            .replace("'", "'"))


def generate_html(clothing_dash, goods_dash, cs_scores, output_path='index.html'):
    clothing_raw = json.dumps(clothing_dash, ensure_ascii=False)
    goods_raw    = json.dumps(goods_dash, ensure_ascii=False)
    update_date  = get_update_timestamp()

    with open('template.html', encoding='utf-8') as f:
        template = f.read()

    cs_scores = {' '.join(k.split()): v for k, v in cs_scores.items()}

    for name, data in cs_scores.items():
        for field in ['memo', 'ai_comment', 'keywords']:
            if field in data and data[field]:
                data[field] = sanitize_text(data[field])

    html = (template
            .replace('__CLOTHING_DATA__', clothing_raw)
            .replace('__GOODS_DATA__',    goods_raw)
            .replace('__CS_DATA__',       json.dumps(cs_scores, ensure_ascii=False))
            .replace('__UPDATE_DATE__',   update_date))

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"생성 완료: {output_path} (업데이트 시각: {update_date})")


def main():
    print("raw 파일 처리 시작...")
    clothing_dash = build_category_dashboard(RAW_FILES['의류'])
    goods_dash    = build_category_dashboard(RAW_FILES['용품'])

    print("\n=== 의류 ===")
    for dept, d in clothing_dash.items():
        s = d['summary']
        print(f"  {dept}: {len(d['stores'])}개 / 채권 {s['total_receivable']:,} / 초과 {s['total_excess']:,}")

    print("\n=== 용품 ===")
    for dept, d in goods_dash.items():
        s = d['summary']
        print(f"  {dept}: {len(d['stores'])}개 / 채권 {s['total_receivable']:,} / 초과 {s['total_excess']:,}")

    # 대리점별 감점 정보 수집
    # - 의류/용품 둘 다 있는 대리점: 카테고리별 감점(각 최대 8점)을 합산, 합계 최대 15점 캡 (회수일/담보 각각)
    # - 한 카테고리만 있는 대리점: 해당 카테고리 감점을 원래 스케일(최대 15점) 그대로 사용
    store_debt_map = {}
    for cat, dash in [('의류', clothing_dash), ('용품', goods_dash)]:
        for dept, data in dash.items():
            for s in data['stores']:
                name = ' '.join(s['name'].split())
                entry = store_debt_map.setdefault(name, {
                    'collateral': 0, 'receivable': 0, 'ratio': 0, 'risk': '해당없음',
                    'deduct_collateral_clothing': 0, 'deduct_collateral_goods': 0,
                    'deduct_collateral_single': 0,
                    'deduct_collection_clothing': 0, 'deduct_collection_goods': 0,
                    'deduct_collection_single': 0,
                    'collection_days_clothing': 0, 'collection_days_goods': 0,
                    'cats_seen': set(),
                })
                entry['cats_seen'].add(cat)
                entry['collateral'] = s['collateral']
                entry['receivable'] = s['receivable']
                entry['ratio']      = s['ratio']
                entry['risk']       = s['risk']
                entry['cat']        = cat
                if cat == '의류':
                    entry['deduct_collateral_clothing'] = s['deduct_collateral_half']
                    entry['deduct_collection_clothing']  = s['deduct_collection_half']
                    entry['collection_days_clothing']    = s['collection_days']
                else:
                    entry['deduct_collateral_goods'] = s['deduct_collateral_half']
                    entry['deduct_collection_goods']  = s['deduct_collection_half']
                    entry['collection_days_goods']    = s['collection_days']
                entry['deduct_collateral_single'] = s['deduct_collateral']
                entry['deduct_collection_single']  = s['deduct_collection']

    for name, entry in store_debt_map.items():
        if len(entry['cats_seen']) >= 2:
            # 의류+용품 둘 다 있음 → 각 카테고리 감점(최대 8점씩) 합산, 합계 최대 15점
            entry['deduct_collateral'] = min(25, entry['deduct_collateral_clothing'] + entry['deduct_collateral_goods'])
            entry['deduct_collection'] = min(25, entry['deduct_collection_clothing'] + entry['deduct_collection_goods'])
            entry['collection_days'] = max(entry['collection_days_clothing'], entry['collection_days_goods'])
        else:
            # 카테고리 1개뿐 → 원래 스케일(최대 15점) 그대로, 표시용 분해값도 채워줌
            entry['deduct_collateral'] = entry['deduct_collateral_single']
            entry['deduct_collection'] = entry['deduct_collection_single']
            if '의류' in entry['cats_seen']:
                entry['deduct_collateral_clothing'] = entry['deduct_collateral_single']
                entry['deduct_collateral_goods'] = 0
                entry['deduct_collection_clothing'] = entry['deduct_collection_single']
                entry['deduct_collection_goods'] = 0
                entry['collection_days'] = entry['collection_days_clothing']
            else:
                entry['deduct_collateral_goods'] = entry['deduct_collateral_single']
                entry['deduct_collateral_clothing'] = 0
                entry['deduct_collection_goods'] = entry['deduct_collection_single']
                entry['deduct_collection_clothing'] = 0
                entry['collection_days'] = entry['collection_days_goods']
        del entry['cats_seen']
        del entry['deduct_collateral_single']
        del entry['deduct_collection_single']

    cs_scores = fetch_cs_data(store_debt_map)

    # 감점 정보 병합
    for name, debt in store_debt_map.items():
        if name in cs_scores:
            cs_scores[name]['deduct_collection']          = debt.get('deduct_collection', 0)
            cs_scores[name]['deduct_collection_clothing']  = debt.get('deduct_collection_clothing', 0)
            cs_scores[name]['deduct_collection_goods']     = debt.get('deduct_collection_goods', 0)
            cs_scores[name]['deduct_collateral']           = debt.get('deduct_collateral', 0)
            cs_scores[name]['deduct_collateral_clothing']  = debt.get('deduct_collateral_clothing', 0)
            cs_scores[name]['deduct_collateral_goods']     = debt.get('deduct_collateral_goods', 0)
            cs_scores[name]['collection_days']             = debt.get('collection_days', 0)
            cs_scores[name]['collection_days_clothing']    = debt.get('collection_days_clothing', 0)
            cs_scores[name]['collection_days_goods']       = debt.get('collection_days_goods', 0)
        else:
            cs_scores[name] = {
                'score': 20, 'partnership_score': 30,
                'p_goods': '', 'p_clothing': '',
                'keywords': '', 'memo': '', 'ai_comment': '',
                'deduct_collection':          debt.get('deduct_collection', 0),
                'deduct_collection_clothing': debt.get('deduct_collection_clothing', 0),
                'deduct_collection_goods':    debt.get('deduct_collection_goods', 0),
                'deduct_collateral':          debt.get('deduct_collateral', 0),
                'deduct_collateral_clothing': debt.get('deduct_collateral_clothing', 0),
                'deduct_collateral_goods':    debt.get('deduct_collateral_goods', 0),
                'collection_days':            debt.get('collection_days', 0),
                'collection_days_clothing':   debt.get('collection_days_clothing', 0),
                'collection_days_goods':      debt.get('collection_days_goods', 0),
            }

    generate_html(clothing_dash, goods_dash, cs_scores)


if __name__ == '__main__':
    try:
        main()
    except FileNotFoundError as e:
        print(f"파일을 찾을 수 없습니다: {e}")
        print("clothing_raw.xls, goods_raw.xls 파일이 저장소 루트에 있는지 확인하세요.")
        sys.exit(1)
