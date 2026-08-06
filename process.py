"""
채권 위험도 대시보드 자동 생성 스크립트

총점 구조 (최대 130점):
  정량점수  최대 70점
    = 채권건전성 50점 (기본 50 - 회수일 감점(최대25) - 담보대비채권 감점(최대25))
    + 매출규모  20점 (3개월 매출 합계 구간, fetch_cs_scores.score_sales_tier)
  + CS 코멘트  최대 30점 (월 누적: 고위험 -5점/건, 중위험 -2점/건, 우수 +1점/건)
  + 파트너십   최대 30점 (월 누적: 위반 1건당 -1점, 용품+의류 합산)
  = 최대 130점
"""

import pandas as pd
import numpy as np
import json
import sys
import os
from datetime import datetime, timezone, timedelta
from fetch_cs_scores import fetch_cs_data, generate_daily_insights

RAW_FILES = {
    '의류': 'clothing_raw.xls',
    '용품': 'goods_raw.xls',
}

CREDIT_GRACE_FILES = {
    '의류': '의류 수주여신.xls',
    '용품': '용품 수주여신.xls',
}

DEPT_TABS = ['영업1팀', '영업2팀', 'E-BIZ팀']

RISK_THRESHOLDS = {
    'safe_max': 0.0,    # 초과율 ≤0% → 적정
    'caution_max': 0.3,  # 0~30% → 주의
    'warning_max': 1.2,  # 30~120% → 경계
    'danger_max': 1.5,   # 120% 초과 → 위기
}

MIN_RECEIVABLE_THRESHOLD = 500_000
MIN_DISPLAY_THRESHOLD = 100_000


# ───────────────────────────────────────────
# 감점 기준
# ───────────────────────────────────────────

def deduct_collection_days(days):
    """회수일 감점 (60일 기준 초과시 감점, 카테고리당 최대 25점).
    의류+용품 둘 다 있는 대리점은 카테고리별로 이 값을 각각 구해 합산한 뒤 25점으로 캡한다."""
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


def deduct_collateral_ratio(collateral, receivable):
    """담보대비 초과율 감점 (카테고리당 최대 25점). classify_risk() 등급 기준(0/30/120%)과 동일하게 정렬.
    초과율 = (채권잔액-담보)/담보. 의류+용품 둘 다 있는 대리점은 카테고리별로 이 값을 각각 구해 합산한 뒤 25점으로 캡한다."""
    # 무담보 & 채권 없음 → 감점 없음
    if collateral == 0 and receivable <= 0:
        return 0
    # 무담보 & 채권 있음 (=관리 등급) → 최대 감점
    if collateral == 0 and receivable > 0:
        return 25
    excess_rate = (receivable - collateral) / collateral * 100
    if excess_rate <= RISK_THRESHOLDS['safe_max'] * 100:      # 적정 (초과율 ≤0%)
        return 0
    elif excess_rate <= RISK_THRESHOLDS['caution_max'] * 100:  # 주의 (0~30%)
        return 10
    elif excess_rate <= RISK_THRESHOLDS['warning_max'] * 100:  # 경계 (30~120%)
        return 15
    else:                                                       # 위기 (120% 초과)
        return 25



def classify_risk(collateral, receivable):
    if abs(receivable) < MIN_RECEIVABLE_THRESHOLD:
        return '해당없음'
    if collateral == 0:
        return '관리' if receivable > 0 else '적정'
    excess_rate = (receivable - collateral) / collateral
    if excess_rate <= RISK_THRESHOLDS['safe_max']:
        return '적정'
    elif excess_rate <= RISK_THRESHOLDS['caution_max']:
        return '주의'
    elif excess_rate <= RISK_THRESHOLDS['warning_max']:
        return '경계'
    else:
        return '위기'


def load_credit_grace(filepath):
    """수주여신(여신유예금액) 로더. ERP '여신현황' 계열 리포트 전용 컬럼 배치(코드=col1,이름=col2,여신유예금액=col34)를
    사용하므로 clothing_raw.xls/goods_raw.xls(채권현황 리포트, 코드=col4)와 컬럼 배치가 다르다.
    파일이 없으면(수주 기간이 아니면) 빈 dict를 반환해 나머지 로직에 영향 없게 한다."""
    if not os.path.exists(filepath):
        return {}
    xl = pd.read_excel(filepath, engine='xlrd', sheet_name=None, header=None)
    df = xl['export']
    data = df.iloc[1:].copy()
    data.columns = range(len(data.columns))

    mask = data[1].astype(str).str.match(r'^D\d+$', na=False)
    store_data = data[mask].copy()

    names = store_data[2].astype(str).apply(lambda x: ' '.join(x.split()))
    grace = pd.to_numeric(store_data[34], errors='coerce').fillna(0)

    result = {}
    for name, g in zip(names, grace):
        result[name] = result.get(name, 0) + float(g)
    return result


def process_raw(filepath, credit_grace_map=None):
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

    # 수주여신(여신유예금액): 수주 기간 중 담보 위에 추가로 열어주는 여신 한도.
    # 담보대비초과율/감점 계산에는 "담보+수주여신"을 유효 담보로 반영해, 정상 승인된 여신 확대를
    # 위험(채권 급증)으로 오인하지 않도록 한다. 원 담보값은 'collateral_base'로 따로 보존해 화면에 분리 표시한다.
    credit_grace_map = credit_grace_map or {}
    name_norm = store_data[5].astype(str).apply(lambda x: ' '.join(x.split()))
    result['credit_grace']  = name_norm.map(credit_grace_map).fillna(0)
    result['collateral_base'] = result['collateral']
    result['collateral']    = result['collateral_base'] + result['credit_grace']

    result['excess'] = result['receivable'] - result['collateral']
    result['ratio']  = np.where(
        result['collateral'] > 0,
        (result['receivable'] - result['collateral']) / result['collateral'], 0.0
    )
    result['risk'] = result.apply(
        lambda r: classify_risk(r['collateral'], r['receivable']), axis=1
    )

    # 감점 계산 (카테고리당 최대 25점 - 의류+용품 둘 다 있으면 합산 후 25점 캡은 store_debt_map 병합 단계에서 처리)
    result['deduct_collection'] = result['collection_days'].apply(deduct_collection_days)
    result['deduct_collateral'] = result.apply(
        lambda r: deduct_collateral_ratio(r['collateral'], r['receivable']), axis=1
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
            'collateral_base':  int(r['collateral_base']),
            'credit_grace':     int(r['credit_grace']),
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
        'total_collateral': int(full_sub['collateral'].sum()),
        'total_credit_grace': int(full_sub['credit_grace'].sum()),
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


def build_category_dashboard(filepath, credit_grace_map=None):
    result = process_raw(filepath, credit_grace_map=credit_grace_map)
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
        for field in ['memo', 'ai_comment', 'keywords', 'insight_comment', 'ai_flag_comment']:
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
    print("수주여신 파일 확인 중...")
    credit_grace = {}
    for cat, fp in CREDIT_GRACE_FILES.items():
        credit_grace[cat] = load_credit_grace(fp)
        if credit_grace[cat]:
            print(f"  {cat} 수주여신: {len(credit_grace[cat])}개 매장, 합계 {sum(credit_grace[cat].values()):,.0f}원")
        else:
            print(f"  {cat} 수주여신: 파일 없음 또는 데이터 없음 (수주 기간이 아니면 정상)")

    clothing_dash = build_category_dashboard(RAW_FILES['의류'], credit_grace_map=credit_grace['의류'])
    goods_dash    = build_category_dashboard(RAW_FILES['용품'], credit_grace_map=credit_grace['용품'])

    print("\n=== 의류 ===")
    for dept, d in clothing_dash.items():
        s = d['summary']
        print(f"  {dept}: {len(d['stores'])}개 / 채권 {s['total_receivable']:,} / 초과 {s['total_excess']:,}")

    print("\n=== 용품 ===")
    for dept, d in goods_dash.items():
        s = d['summary']
        print(f"  {dept}: {len(d['stores'])}개 / 채권 {s['total_receivable']:,} / 초과 {s['total_excess']:,}")

    # 대리점별 감점 정보 수집 (채권건전성 50점 = 회수일 최대25 + 담보 최대25)
    # 카테고리별(의류/용품) 감점을 각각 구해 합산하고, 회수일/담보 각각 25점으로 캡한다.
    # (카테고리가 1개뿐이면 다른 쪽은 0이라 자연히 그 카테고리 값 그대로 반영됨)
    store_debt_map = {}
    for cat, dash in [('의류', clothing_dash), ('용품', goods_dash)]:
        for dept, data in dash.items():
            for s in data['stores']:
                name = ' '.join(s['name'].split())
                entry = store_debt_map.setdefault(name, {
                    'collateral': 0, 'receivable': 0, 'ratio': 0, 'risk': '해당없음',
                    'clothing_risk': '해당없음', 'goods_risk': '해당없음',
                    'deduct_collateral_clothing': 0, 'deduct_collateral_goods': 0,
                    'deduct_collection_clothing': 0, 'deduct_collection_goods': 0,
                    'collection_days_clothing': 0, 'collection_days_goods': 0,
                    'credit_grace_clothing': 0, 'credit_grace_goods': 0,
                })
                entry['collateral'] += s['collateral']
                entry['receivable'] += s['receivable']
                entry['risk']       = s['risk']
                entry['cat']        = cat
                if cat == '의류':
                    entry['clothing_risk'] = s['risk']
                    entry['deduct_collateral_clothing'] = s['deduct_collateral']
                    entry['deduct_collection_clothing']  = s['deduct_collection']
                    entry['collection_days_clothing']    = s['collection_days']
                    entry['credit_grace_clothing']        = s['credit_grace']
                else:
                    entry['goods_risk'] = s['risk']
                    entry['deduct_collateral_goods'] = s['deduct_collateral']
                    entry['deduct_collection_goods']  = s['deduct_collection']
                    entry['collection_days_goods']    = s['collection_days']
                    entry['credit_grace_goods']        = s['credit_grace']

    for name, entry in store_debt_map.items():
        entry['deduct_collateral'] = min(25, entry['deduct_collateral_clothing'] + entry['deduct_collateral_goods'])
        entry['deduct_collection'] = min(25, entry['deduct_collection_clothing'] + entry['deduct_collection_goods'])
        entry['collection_days'] = max(entry['collection_days_clothing'], entry['collection_days_goods'])
        entry['ratio'] = (entry['receivable'] - entry['collateral']) / entry['collateral'] if entry['collateral'] > 0 else 0.0

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
            cs_scores[name]['credit_grace_clothing']       = debt.get('credit_grace_clothing', 0)
            cs_scores[name]['credit_grace_goods']          = debt.get('credit_grace_goods', 0)
        else:
            cs_scores[name] = {
                'score': 30, 'partnership_score': 30,
                'sales_score': 0, 'sales_score_goods': 0, 'sales_score_clothing': 0,
                'sales_3m_goods': 0, 'sales_3m_clothing': 0,
                'p_goods': '', 'p_clothing': '', 'p_goods_count': 0, 'p_clothing_count': 0,
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
                'credit_grace_clothing':      debt.get('credit_grace_clothing', 0),
                'credit_grace_goods':         debt.get('credit_grace_goods', 0),
            }

    # ── 대리점별 종합점수·등급 계산 (template.html buildAiDash의 계산과 동일한 공식) ──
    # "AI가 분석한 오늘의 대리점 인사이트" 팝업(경고/기회/검토)에 노출할 대리점을 선정하기 위함
    risk_order = ['위기', '경계', '주의', '적정', '관리', '해당없음']

    def combine_worst_risk(risks):
        w = '해당없음'
        for r in risks:
            if r and risk_order.index(r) < risk_order.index(w):
                w = r
        return w

    all_store_totals = []
    for name, debt in store_debt_map.items():
        cs = cs_scores.get(name, {})
        cs_score = min(30, max(0, cs.get('score', 30)))
        partnership_score = min(30, max(0, cs.get('partnership_score', 30)))
        sales_score_goods = min(10, max(0, cs.get('sales_score_goods', 0)))
        sales_score_clothing = min(10, max(0, cs.get('sales_score_clothing', 0)))
        sales_score = sales_score_goods + sales_score_clothing
        deduct_collection = cs.get('deduct_collection', 0)
        deduct_collateral = cs.get('deduct_collateral', 0)
        health_score = max(0, 50 - deduct_collection - deduct_collateral)
        quant_score = min(70, health_score + sales_score)
        total_score = min(130, quant_score + cs_score + partnership_score)
        worst_risk = combine_worst_risk([debt.get('clothing_risk'), debt.get('goods_risk')])
        all_store_totals.append({
            'name': name,
            'total_score': total_score,
            'worst_risk': worst_risk,
            'ai_assessed_risk': cs.get('ai_assessed_risk', ''),
            'ai_mismatch': cs.get('ai_mismatch', False),
            'ai_mismatch_direction': cs.get('ai_mismatch_direction', ''),
            'keywords': cs.get('keywords', ''),
            'count_high': cs.get('count_high', 0),
            'count_mid': cs.get('count_mid', 0),
            'count_low': cs.get('count_low', 0),
        })

    warn_list = sorted(
        [s for s in all_store_totals if s['worst_risk'] in ('위기', '경계', '관리')],
        key=lambda s: s['total_score']
    )[:3]
    opp_list = sorted(
        [s for s in all_store_totals if s['worst_risk'] != '관리' and s['total_score'] >= 80],
        key=lambda s: -s['total_score']
    )[:3]
    review_list = sorted(
        [s for s in all_store_totals if s['ai_mismatch']],
        key=lambda s: (0 if s['ai_mismatch_direction'] == 'severe' else 1)
    )[:3]
    # 'AI 추가 발견' 후보: 아직 위기/경계/관리 등급은 아니지만, 이번 달 CS 키워드가 입력됐고
    # 고위험/중위험 건이 하나라도 있는 대리점 (완전히 깨끗한 곳은 검토할 필요 없어 제외)
    candidate_list = sorted(
        [s for s in all_store_totals
         if s['worst_risk'] not in ('위기', '경계', '관리') and s['keywords']
         and (s['count_high'] > 0 or s['count_mid'] > 0)],
        key=lambda s: (-s['count_high'], -s['count_mid'])
    )[:15]

    api_key = os.environ.get('GEMINI_API_KEY', '')
    daily_insights = generate_daily_insights(warn_list, opp_list, review_list, candidate_list, api_key)
    for name, comment in daily_insights.get('comments', {}).items():
        if name in cs_scores:
            cs_scores[name]['insight_comment'] = comment
    for name, reason in daily_insights.get('flagged', {}).items():
        if name in cs_scores:
            cs_scores[name]['ai_flagged'] = True
            cs_scores[name]['ai_flag_comment'] = reason

    generate_html(clothing_dash, goods_dash, cs_scores)


if __name__ == '__main__':
    try:
        main()
    except FileNotFoundError as e:
        print(f"파일을 찾을 수 없습니다: {e}")
        print("clothing_raw.xls, goods_raw.xls 파일이 저장소 루트에 있는지 확인하세요.")
        sys.exit(1)
