"""
채권 위험도 대시보드 자동 생성 스크립트

총점 구조 (최대 130점):
  정량점수  최대 70점
    = 채권건전성 50점 (회수일 감점(카테고리당 최대15, 의류+용품 합산 최대30) + 담보대비채권 감점(카테고리당 최대10, 의류+용품 합산 최대20)을 50에서 차감)
    + 매출규모  10점 (3개월 매출 합계 구간, fetch_cs_scores.score_sales_tier, 용품5+의류5)
    + 매출목표 달성 10점 (3개월 목표합계 대비 3개월 매출합계 달성여부, fetch_cs_scores.score_target_achieve, 용품5+의류5 — 달성 5점/미달성 0점)
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
from fetch_cs_scores import fetch_cs_data, generate_daily_insights, load_insight_cache, save_insight_cache, get_kst_today_str

RAW_FILES = {
    '의류': 'clothing_raw.xls',
    '용품': 'goods_raw.xls',
}

DEPT_TABS = ['영업1팀', '영업2팀', 'E-BIZ팀']

RISK_THRESHOLDS = {
    'safe_max': 0.0,    # 초과율 ≤0% → 적정
    'caution_max': 0.3,  # 0~30% → 주의
    'warning_max': 1.2,  # 30~120% → 경계
    'danger_max': 1.5,   # 120% 초과 → 위기
}

# (폐기됨) 과거엔 성수기(2~10월, 수주기간)에 완화된 컷라인을 별도 적용했으나,
# 수주여신 부여 자체가 회사 입장에서 감수하는 리스크이므로 그 기간에도 동일하게
# 엄격기준으로 평가해야 한다는 결정에 따라 계절 구분 없이 RISK_THRESHOLDS를
# 연중 동일하게 적용하는 것으로 변경.
MIN_RECEIVABLE_THRESHOLD = 500_000
MIN_DISPLAY_THRESHOLD = 100_000
ACTIVE_THRESHOLDS = RISK_THRESHOLDS


# ───────────────────────────────────────────
# 감점 기준
# ───────────────────────────────────────────

def deduct_collection_days(days):
    """회수일 감점 (60일 기준 초과시 감점, 카테고리당 최대 15점 - 의류/용품 각각 독립 채점).
    의류+용품 둘 다 있는 대리점은 카테고리별로 이 값을 각각 구해 그대로 합산(최대 30점)한다."""
    try:
        days = float(days)
    except (TypeError, ValueError):
        return 0
    if days <= 60:
        return 0
    elif days <= 80:
        return 6
    elif days <= 90:
        return 9
    else:
        return 15


def deduct_collateral_ratio(collateral, receivable):
    """담보대비 초과율 감점 (카테고리당 최대 10점 - 의류/용품 각각 독립 채점). classify_risk() 등급 기준과 동일하게 정렬.
    초과율 = (채권잔액-담보)/담보. 의류+용품 둘 다 있는 대리점은 카테고리별로 이 값을 각각 구해 그대로 합산(최대 20점)한다.
    연중 동일 엄격기준(ACTIVE_THRESHOLDS=RISK_THRESHOLDS) 적용 — 수주기간 완화 없음."""
    # 무담보 & 채권 없음 → 감점 없음
    if collateral == 0 and receivable <= 0:
        return 0
    # 무담보 & 채권 있음 (=관리 등급) → 최대 감점
    if collateral == 0 and receivable > 0:
        return 10
    excess_rate = (receivable - collateral) / collateral * 100
    if excess_rate <= ACTIVE_THRESHOLDS['safe_max'] * 100:      # 적정
        return 0
    elif excess_rate <= ACTIVE_THRESHOLDS['caution_max'] * 100:  # 주의
        return 4
    elif excess_rate <= ACTIVE_THRESHOLDS['warning_max'] * 100:  # 경계
        return 6
    else:                                                       # 위기
        return 10



def classify_risk(collateral, receivable):
    """소액 채권(50만원 미만)은 주의/경계/위기 등급의 노이즈를 막기 위해 '해당없음' 처리하되,
    담보가 충분해 '적정'인 경우와 무담보 상태라 '관리'인 경우는 소액이어도 그대로 노출한다
    (둘 다 노이즈가 아니라 실제 담보 상태를 정확히 보여줘야 하는 정보이므로).
    연중 동일 엄격기준(ACTIVE_THRESHOLDS=RISK_THRESHOLDS) 적용 — 수주기간 완화 없음."""
    if collateral == 0:
        grade = '관리' if receivable > 0 else '적정'
    else:
        excess_rate = (receivable - collateral) / collateral
        if excess_rate <= ACTIVE_THRESHOLDS['safe_max']:
            grade = '적정'
        elif excess_rate <= ACTIVE_THRESHOLDS['caution_max']:
            grade = '주의'
        elif excess_rate <= ACTIVE_THRESHOLDS['warning_max']:
            grade = '경계'
        else:
            grade = '위기'
    if grade not in ('적정', '관리') and abs(receivable) < MIN_RECEIVABLE_THRESHOLD:
        return '해당없음'
    return grade


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
        (result['receivable'] - result['collateral']) / result['collateral'], 0.0
    )
    result['risk'] = result.apply(
        lambda r: classify_risk(r['collateral'], r['receivable']), axis=1
    )

    # 감점 계산 (회수일 카테고리당 최대15/합산최대30, 담보 카테고리당 최대10/합산최대20 - 의류+용품 합산은 store_debt_map 병합 단계에서 처리)
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


def generate_html(clothing_dash, goods_dash, cs_scores, opp_shown_names, output_path='index.html'):
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
            .replace('__OPP_SHOWN__',     json.dumps(opp_shown_names, ensure_ascii=False))
            .replace('__UPDATE_DATE__',   update_date))

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"생성 완료: {output_path} (업데이트 시각: {update_date})")


def main():
    print("raw 파일 처리 시작...")
    print(f"현재 적용 담보대비 초과율 컷라인 (연중 동일): "
          f"적정≤{ACTIVE_THRESHOLDS['safe_max']*100:.0f}% / "
          f"주의~{ACTIVE_THRESHOLDS['caution_max']*100:.0f}% / "
          f"경계~{ACTIVE_THRESHOLDS['warning_max']*100:.0f}% / 위기 초과")

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

    # 대리점별 감점 정보 수집 (채권건전성 50점 = 회수일 최대30(의류15+용품15) + 담보 최대20(의류10+용품10))
    # 카테고리별(의류/용품) 감점을 각각 구해 합산하고, 회수일은 30점, 담보는 20점으로 캡한다.
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
                else:
                    entry['goods_risk'] = s['risk']
                    entry['deduct_collateral_goods'] = s['deduct_collateral']
                    entry['deduct_collection_goods']  = s['deduct_collection']
                    entry['collection_days_goods']    = s['collection_days']

    for name, entry in store_debt_map.items():
        entry['deduct_collateral'] = min(20, entry['deduct_collateral_clothing'] + entry['deduct_collateral_goods'])
        entry['deduct_collection'] = min(30, entry['deduct_collection_clothing'] + entry['deduct_collection_goods'])
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
        else:
            cs_scores[name] = {
                'score': 30, 'partnership_score': 30,
                'sales_score': 0, 'sales_score_goods': 0, 'sales_score_clothing': 0,
                'sales_3m_goods': 0, 'sales_3m_clothing': 0,
                'target_score': 0, 'target_score_goods': 0, 'target_score_clothing': 0,
                'target_3m_goods': 0, 'target_3m_clothing': 0,
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

    # 종합점수(130점) 기준 등급 - 대시보드에 실제로 표시되는 등급과 반드시 일치해야 함
    # (담보초과율 기반 worst_risk와는 별개 체계이므로 절대 혼용하지 않음 - template.html getRiskClass와 동일 공식)
    def score_grade(total_score, worst_risk):
        if worst_risk == '관리':
            return '관리'
        if total_score >= 116:
            return '적정'
        if total_score >= 101:
            return '주의'
        if total_score >= 90:
            return '경계'
        return '위기'

    all_store_totals = []
    for name, debt in store_debt_map.items():
        cs = cs_scores.get(name, {})
        cs_score = min(30, max(0, cs.get('score', 30)))
        partnership_score = min(30, max(0, cs.get('partnership_score', 30)))
        sales_score_goods = min(5, max(0, cs.get('sales_score_goods', 0)))
        sales_score_clothing = min(5, max(0, cs.get('sales_score_clothing', 0)))
        sales_score = sales_score_goods + sales_score_clothing
        target_score_goods = min(5, max(0, cs.get('target_score_goods', 0)))
        target_score_clothing = min(5, max(0, cs.get('target_score_clothing', 0)))
        target_score = target_score_goods + target_score_clothing
        deduct_collection = cs.get('deduct_collection', 0)
        deduct_collateral = cs.get('deduct_collateral', 0)
        health_score = max(0, 50 - deduct_collection - deduct_collateral)
        quant_score = min(70, health_score + sales_score + target_score)
        total_score = min(130, quant_score + cs_score + partnership_score)
        # worst_risk: 담보초과율 기반 개별(의류/용품) 채권 리스크 - "채권 리스크" 진단 표시용, 종합등급과는 다른 지표
        worst_risk = combine_worst_risk([debt.get('clothing_risk'), debt.get('goods_risk')])
        all_store_totals.append({
            'name': name,
            'total_score': total_score,
            'worst_risk': worst_risk,
            'score_grade': score_grade(total_score, worst_risk),
            'ai_assessed_risk': cs.get('ai_assessed_risk', ''),
            'ai_mismatch': cs.get('ai_mismatch', False),
            'ai_mismatch_direction': cs.get('ai_mismatch_direction', ''),
            'keywords': cs.get('keywords', ''),
            'count_high': cs.get('count_high', 0),
            'count_mid': cs.get('count_mid', 0),
            'count_low': cs.get('count_low', 0),
        })

    warn_list = sorted(
        [s for s in all_store_totals if s['score_grade'] in ('위기', '경계', '관리')],
        key=lambda s: s['total_score']
    )[:3]
    # 기회 요소 로테이션: 어제까지 노출됐던 대리점은 후순위로 밀어서, 만점 근처에 고정된
    # 대리점이 매일 반복 노출되는 문제를 방지 (점수 자체는 그대로 반영, 순서만 조정)
    today_str = get_kst_today_str()
    opp_rotation_cache = load_insight_cache()
    if opp_rotation_cache.get('opp_shown_date') == today_str:
        prev_opp_names = set()  # 같은 날 재빌드는 로테이션 재적용 안 함 (오늘 목록 유지)
    else:
        prev_opp_names = set(opp_rotation_cache.get('opp_shown_names', []))

    opp_pool = sorted(
        [s for s in all_store_totals if s['score_grade'] != '관리' and s['total_score'] >= 116],
        key=lambda s: -s['total_score']
    )
    opp_fresh = [s for s in opp_pool if s['name'] not in prev_opp_names]
    opp_repeat = [s for s in opp_pool if s['name'] in prev_opp_names]
    opp_list = (opp_fresh + opp_repeat)[:3]

    if opp_rotation_cache.get('opp_shown_date') != today_str:
        save_insight_cache({
            'opp_shown_date': today_str,
            'opp_shown_names': [s['name'] for s in opp_list],
        })
    review_list = sorted(
        [s for s in all_store_totals if s['ai_mismatch']],
        key=lambda s: (0 if s['ai_mismatch_direction'] == 'severe' else 1)
    )[:3]
    # 'AI 추가 발견' 후보: 아직 위기/경계/관리 등급은 아니지만, 이번 달 CS 키워드가 입력됐고
    # 고위험/중위험 건이 하나라도 있는 대리점 (완전히 깨끗한 곳은 검토할 필요 없어 제외)
    candidate_list = sorted(
        [s for s in all_store_totals
         if s['score_grade'] not in ('위기', '경계', '관리') and s['keywords']
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

    generate_html(clothing_dash, goods_dash, cs_scores, [s['name'] for s in opp_list])


if __name__ == '__main__':
    try:
        main()
    except FileNotFoundError as e:
        print(f"파일을 찾을 수 없습니다: {e}")
        print("clothing_raw.xls, goods_raw.xls 파일이 저장소 루트에 있는지 확인하세요.")
        sys.exit(1)
