"""
채권 위험도 대시보드 자동 생성 스크립트

사용법:
  python process.py

전제조건:
  - 저장소 루트에 clothing_raw.xls (의류), goods_raw.xls (용품) 파일이 있어야 함
  - 두 파일 모두 ERP에서 export한 동일한 컬럼 구조 (부서/영업사원/매장/담보금액/.../미수채권잔액)

출력:
  - index.html (저장소 루트에 생성/덮어쓰기)
"""

import pandas as pd
import numpy as np
import json
import sys
from datetime import datetime, timezone, timedelta
from fetch_cs_scores import fetch_cs_data

# ──────────────────────────────────────────────
# 설정값 (회사 정책에 맞게 이 부분만 수정하면 됨)
# ──────────────────────────────────────────────

RAW_FILES = {
    '의류': 'clothing_raw.xls',
    '용품': 'goods_raw.xls',
}

DEPT_TABS = ['영업1팀', '영업2팀', 'E-BIZ팀']

# 위험도 판정 기준 (채권잔액 / 담보금액 비율)
# 회사 정책 확정되면 이 숫자만 바꾸면 전체 반영됨
RISK_THRESHOLDS = {
    'safe_max': 0.6,      # 이 이하면 '적정'
    'caution_max': 1.0,   # 이 이하면 '주의' (위 초과는 일단 '경계')
    'warning_max': 1.5,   # 이 이하면 '경계'
    'danger_max': 2.0,    # 이 이하면 '위기' (이 초과도 '위기')
}

MIN_RECEIVABLE_THRESHOLD = 500_000   # 이 미만 채권은 '해당없음' 처리
MIN_DISPLAY_THRESHOLD = 100_000      # 대시보드에 표시할 최소 채권 규모


def classify_risk(collateral, receivable, ratio):
    """위험도 5단계 분류: 적정/주의/경계/위기/관리"""
    if abs(receivable) < MIN_RECEIVABLE_THRESHOLD:
        return '해당없음'
    if collateral == 0 and receivable > 0:
        return '관리'  # 담보없음
    if ratio <= RISK_THRESHOLDS['safe_max']:
        return '적정'
    elif ratio <= RISK_THRESHOLDS['caution_max']:
        return '주의'
    elif ratio <= RISK_THRESHOLDS['warning_max']:
        return '경계'
    else:
        return '위기'


def process_raw(filepath):
    """ERP raw export(.xls)를 읽어 매장(D코드) 단위 DataFrame으로 변환"""
    xl = pd.read_excel(filepath, engine='xlrd', sheet_name=None, header=None)
    df = xl['export']
    data = df.iloc[1:].copy()
    data.columns = range(len(data.columns))

    # D코드(매장)만 필터링 — F코드(협회/단체/개인 등)는 제외
    mask = data[4].astype(str).str.match(r'^D\d+$', na=False)
    store_data = data[mask].copy()

    result = pd.DataFrame()
    result['code'] = store_data[4].astype(str)
    result['name'] = store_data[5].astype(str)
    result['salesperson'] = store_data[3].astype(str).str.strip()
    result['dept_code'] = store_data[0].astype(str)
    result['dept_name'] = store_data[1].astype(str)
    result['collateral'] = pd.to_numeric(store_data[6], errors='coerce').fillna(0)
    result['receivable'] = pd.to_numeric(store_data[13], errors='coerce').fillna(0)
    result['sales'] = pd.to_numeric(store_data[11], errors='coerce').fillna(0)
    result['collection'] = pd.to_numeric(store_data[12], errors='coerce').fillna(0)

    result['excess'] = result['receivable'] - result['collateral']
    result['ratio'] = np.where(result['collateral'] > 0,
                                result['receivable'] / result['collateral'], 0.0)
    result['risk'] = result.apply(
        lambda r: classify_risk(r['collateral'], r['receivable'], r['ratio']), axis=1
    )
    return result


def build_group_data(sub):
    """부서별 서브셋을 대시보드용 JSON 구조로 변환"""
    sub = sub[sub['receivable'].abs() > MIN_DISPLAY_THRESHOLD].copy()
    stores = []
    for _, r in sub.iterrows():
        stores.append({
            'code': r['code'], 'name': r['name'], 'salesperson': r['salesperson'],
            'collateral': int(r['collateral']), 'receivable': int(r['receivable']),
            'excess': int(r['excess']), 'ratio': round(float(r['ratio']), 2),
            'risk': r['risk']
        })
    stores = sorted(stores, key=lambda x: -x['receivable'])

    summary = {
        'total_collateral': int(sub['collateral'].sum()),
        'total_receivable': int(sub['receivable'].sum()),
        'total_excess': int(sub[sub['excess'] > 0]['excess'].sum()),
        'risk_counts': {k: int(v) for k, v in sub['risk'].value_counts().to_dict().items()}
    }

    sp_summary = {}
    for sp, g in sub.groupby('salesperson'):
        sp_summary[sp] = {
            'collateral': int(g['collateral'].sum()),
            'receivable': int(g['receivable'].sum()),
            'excess': int(g[g['excess'] > 0]['excess'].sum()),
            'stores': len(g)
        }
    return {'stores': stores, 'summary': summary, 'by_salesperson': sp_summary}


def build_category_dashboard(filepath):
    """raw 파일 하나(용품 또는 의류)를 부서별 4탭 dashboard dict로 변환"""
    result = process_raw(filepath)
    dashboard = {}
    for dept in DEPT_TABS:
        sub = result[result['dept_name'] == dept]
        dashboard[dept] = build_group_data(sub)
    return dashboard


def get_update_timestamp():
    """한국 시간(KST) 기준 업데이트 시각 문자열 생성"""
    kst = timezone(timedelta(hours=9))
    now = datetime.now(kst)
    weekdays = ['월', '화', '수', '목', '금', '토', '일']
    weekday = weekdays[now.weekday()]
    return f"{now.year}.{now.month:02d}.{now.day:02d}({weekday}) {now.hour:02d}:{now.minute:02d}"


def generate_html(clothing_dash, goods_dash, cs_scores, output_path='index.html'):
    """최종 index.html 생성"""
    clothing_raw = json.dumps(clothing_dash, ensure_ascii=False)
    goods_raw = json.dumps(goods_dash, ensure_ascii=False)
    cs_raw = json.dumps(cs_scores, ensure_ascii=False)
    update_date = get_update_timestamp()

    with open('template.html', encoding='utf-8') as f:
        template = f.read()

    html = (template
            .replace('__CLOTHING_DATA__', clothing_raw)
            .replace('__GOODS_DATA__', goods_raw)
            .replace('__CS_DATA__', json.dumps(cs_scores, ensure_ascii=False))
            .replace('__UPDATE_DATE__', update_date))

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"생성 완료: {output_path} (업데이트 시각: {update_date})")


def main():
    print("raw 파일 처리 시작...")
    clothing_dash = build_category_dashboard(RAW_FILES['의류'])
    goods_dash = build_category_dashboard(RAW_FILES['용품'])

    print("\n=== 의류 ===")
    for dept, d in clothing_dash.items():
        s = d['summary']
        print(f"  {dept}: {len(d['stores'])}개 / 채권 {s['total_receivable']:,} / 초과 {s['total_excess']:,}")

    print("\n=== 용품 ===")
    for dept, d in goods_dash.items():
        s = d['summary']
        print(f"  {dept}: {len(d['stores'])}개 / 채권 {s['total_receivable']:,} / 초과 {s['total_excess']:,}")

    cs_scores = fetch_cs_data()
    generate_html(clothing_dash, goods_dash, cs_scores)


if __name__ == '__main__':
    try:
        main()
    except FileNotFoundError as e:
        print(f"파일을 찾을 수 없습니다: {e}")
        print("clothing_raw.xls, goods_raw.xls 파일이 저장소 루트에 있는지 확인하세요.")
        sys.exit(1)
