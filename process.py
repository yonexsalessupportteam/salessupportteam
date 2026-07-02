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
    """회수일 감점 (60일 기준 초과시 감점, 최대 15점)"""
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
    """담보대비 채권잔액 감점 (최대 15점)"""
    # 무담보 & 채권 없음 → 감점 없음
    if collateral == 0 and receivable <= 0:
        return 0
    # 무담보 & 채권 있음 → 최대 감점
    if collateral == 0 and receivable > 0:
        return 15
    ratio = receivable / collateral * 100
    if ratio <= 50:
        return 0
    elif ratio <= 100:
        return 5
    elif ratio <= 150:
        return 10
    elif ratio <= 200:
        return 12
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
    result['deduct_collateral'] = result.apply(
        lambda r: deduct_collateral_ratio(r['collateral'], r['receivable']), axis=1
    )

    return result


def build_group_data(sub):
    sub = sub[sub['receivable'].abs() > MIN_DISPLAY_THRESHOLD].copy()
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
            'deduct_collateral': int(r['deduct_collateral']),
        })
    stores = sorted(stores, key=lambda x: -x['receivable'])

    summary = {
        'total_collateral': int(sub['collateral'].sum()),
        'total_receivable': int(sub['receivable'].sum()),
        'total_excess':     int(sub[sub['excess'] > 0]['excess'].sum()),
        'risk_counts':      {k: int(v) for k, v in sub['risk'].value_counts().to_dict().items()}
    }

    sp_summary = {}
    for sp, g in sub.groupby('salesperson'):
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
    store_debt_map = {}
    for cat, dash in [('의류', clothing_dash), ('용품', goods_dash)]:
        for dept, data in dash.items():
            for s in data['stores']:
                name = ' '.join(s['name'].split())
                if name not in store_debt_map:
                    store_debt_map[name] = {
                        'collateral':        s['collateral'],
                        'receivable':        s['receivable'],
                        'ratio':             s['ratio'],
                        'risk':              s['risk'],
                        'cat':               cat,
                        'deduct_collection': s['deduct_collection'],
                        'deduct_collateral': s['deduct_collateral'],
                        'collection_days':   s['collection_days'],
                    }
                else:
                    # 의류+용품 둘 다 있으면 나쁜 쪽 감점 기준
                    store_debt_map[name]['deduct_collection'] = max(
                        store_debt_map[name]['deduct_collection'], s['deduct_collection'])
                    store_debt_map[name]['deduct_collateral'] = max(
                        store_debt_map[name]['deduct_collateral'], s['deduct_collateral'])

    cs_scores = fetch_cs_data(store_debt_map)

    # 감점 정보 병합
    for name, debt in store_debt_map.items():
        if name in cs_scores:
            cs_scores[name]['deduct_collection'] = debt.get('deduct_collection', 0)
            cs_scores[name]['deduct_collateral']  = debt.get('deduct_collateral', 0)
            cs_scores[name]['collection_days']    = debt.get('collection_days', 0)
        else:
            cs_scores[name] = {
                'score': 20, 'partnership_score': 30,
                'p_goods': '', 'p_clothing': '',
                'keywords': '', 'memo': '', 'ai_comment': '',
                'deduct_collection': debt.get('deduct_collection', 0),
                'deduct_collateral': debt.get('deduct_collateral', 0),
                'collection_days':   debt.get('collection_days', 0),
            }

    generate_html(clothing_dash, goods_dash, cs_scores)


if __name__ == '__main__':
    try:
        main()
    except FileNotFoundError as e:
        print(f"파일을 찾을 수 없습니다: {e}")
        print("clothing_raw.xls, goods_raw.xls 파일이 저장소 루트에 있는지 확인하세요.")
        sys.exit(1)
