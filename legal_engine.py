# -*- coding: utf-8 -*-
"""
legal_engine.py — 법적쇠퇴진단 (도시재생법 시행령 방식, 규칙 기반 o/x)
=====================================================================
복합쇠퇴지수(Z/T·가중치)와 별개. 표준화 없이 임계값 판정.

부문·규칙 (골든본 '전주시 법적쇠퇴진단' 시트 수식에서 추출):
  · 인구: 30년 최다 대비 증감률 ≤ -20% → o(S)  OR  최근5년 3년연속감소 → o(T)
  · 산업: 총사업체수 증감률 < -5% → o(AT)       OR  최근5년 3년연속감소 → o(AU)
  · 물리: 노후건축물(2004이전)/전체 ≥ 50% → o(AY)
  · 부합개수(AZ) = (S|T) + (AT|AU) + (AY)   (부문별 1점, max 3)
  · 쇠퇴지역(BA) = 부합개수 ≥ 2

입력 raw = decline_engine 과 동일 dict (to_in, to_fa, ho_yr 사용).
집계구/행정동 레벨 모두 지원(level='jgu'|'dong').
"""
import numpy as np
import pandas as pd
import config as C
import decline_engine as E


def _yearly(raw_df, key):
    """단위 × 연도 합계표(wide)."""
    g = E.aggregate_dong_year_code(raw_df, key)
    return E._pivot_year(g, key=key)


def _ox(cond_series):
    return cond_series.map(lambda b: 'o' if bool(b) else 'x')


def _consec_decline(pivot_years, n_last=5, min_run=3):
    """최근 n_last개 연도열에서 min_run개 연속 시점 감소가 있으면 'o'.
    골든 수식 OR(AND(Q<P,P<O),…) 등가(연속 감소 시점 수 = min_run)."""
    if pivot_years is None or pivot_years.shape[1] == 0:
        return pd.Series('x', index=[] if pivot_years is None else pivot_years.index)
    cols = sorted(pivot_years.columns)[-n_last:]
    sub = pivot_years[cols]

    def check(vals):
        run = 1
        for i in range(1, len(vals)):
            a, b = vals[i - 1], vals[i]
            if pd.notna(a) and pd.notna(b) and b < a:
                run += 1
                if run >= min_run:
                    return 'o'
            else:
                run = 1
        return 'x'

    return sub.apply(lambda r: check(list(r.values)), axis=1)


def run_legal(raw, level='dong'):
    """법적쇠퇴진단 → DataFrame(단위 index × 판정열).
    열: 인구증감률, 인구감소해당, 인구연속감소, 사업체증감률, 사업체감소해당,
        사업체연속감소, 노후건축물비율, 물리노후해당, 부합개수, 쇠퇴지역."""
    key = '집계구' if level == 'jgu' else '행정동코드'

    pop = _yearly(raw['to_in'], key)          # 인구
    biz = _yearly(raw['to_fa'], key)          # 총사업체수

    R = E._growth_vs_max(pop, C.YEAR_POP_LATEST)     # 인구 증감률(%)
    S = _ox(R <= -20)                                 # ≤ -20% → o
    T = _consec_decline(pop)                          # 인구 연속감소

    AS = E._growth_vs_max(biz, C.YEAR_BIZ_LATEST)     # 사업체 증감률(%)
    AT = _ox(AS < -5)                                  # < -5% → o
    AU = _consec_decline(biz)                          # 사업체 연속감소

    AX = E.derive_old_housing_ratio(raw['ho_yr'], key)  # 노후/전체 *100
    AY = _ox(AX >= 50)                                   # ≥ 50% → o

    # 공통 인덱스(합집합)로 정렬
    idx = pop.index.union(biz.index).union(AX.index)
    R, S, T = R.reindex(idx).fillna(0.0), S.reindex(idx).fillna('x'), T.reindex(idx).fillna('x')
    AS, AT, AU = AS.reindex(idx).fillna(0.0), AT.reindex(idx).fillna('x'), AU.reindex(idx).fillna('x')
    AX, AY = AX.reindex(idx).fillna(0.0), AY.reindex(idx).fillna('x')

    pop_hit = ((S == 'o') | (T == 'o')).astype(int)
    biz_hit = ((AT == 'o') | (AU == 'o')).astype(int)
    phy_hit = (AY == 'o').astype(int)
    AZ = pop_hit + biz_hit + phy_hit
    BA = _ox(AZ >= 2)

    return pd.DataFrame({
        '인구증감률': R, '인구감소해당': S, '인구연속감소': T,
        '사업체증감률': AS, '사업체감소해당': AT, '사업체연속감소': AU,
        '노후건축물비율': AX, '물리노후해당': AY,
        '부합개수': AZ, '쇠퇴지역': BA,
    }, index=idx)
