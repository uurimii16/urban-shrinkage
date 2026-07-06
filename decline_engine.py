# -*- coding: utf-8 -*-
"""
decline_engine.py — 복합쇠퇴진단 계산 엔진 (순수 함수)
======================================================
입력 : 분류별 원시 long DataFrame  (컬럼 = 연도 | 집계구 | CODE | 값 | 행정동코드)
       · 값의 결측(N/A)은 NaN 으로 들어와 있어야 함
       · 행정동코드 = LEFT(집계구, 8)  (금암동 보정은 원시데이터 단계에서 이미 반영)
출력 : 행정동(34개) × 12지표의 값·Z·T·부문·종합

엑셀 등가:
  · groupby(행정동코드).sum(min_count=1)  ≡  SUMIFS (결측 제외)
  · 표준화 mean  ≡ AGGREGATE(1,6)=AVERAGE(오류무시)
  · 표준화 std   ≡ AGGREGATE(8,6)=STDEV.P(모표준편차, 오류무시)
  · IFERROR(…,0) ≡ 분모 0 / 최댓값 0 일 때 결과 0

의존성: pandas, numpy 만 사용 (Jenks도 numpy로 자체 구현).
"""
import numpy as np
import pandas as pd
import config as C


# ══════════════════════════════════════════════════════════════
#  ① 집계구 → 행정동 집계
# ══════════════════════════════════════════════════════════════
def aggregate_dong_year_code(df, key='행정동코드'):
    """(집계키, 연도, code숫자) 별 값 합계. 결측은 합산에서 제외(min_count=1).
    key='행정동코드'(기본, 행정동 레벨) 또는 '집계구'(집계구 레벨)."""
    d = df.copy()
    # code 숫자부분 = 코드 끝 3자리 (in_age_004→4, cp_bem_017→17 …)
    d['code_no'] = d['CODE'].str[-3:].astype(int)
    g = (d.groupby([key, '연도', 'code_no'])['값']
           .sum(min_count=1)          # 전부 NaN인 그룹은 NaN 유지
           .reset_index())
    return g


def _pivot_year(df_dong, value_filter=None, key='행정동코드'):
    """집계키 × 연도 합계표(wide). value_filter: code_no 집합(None=전체)."""
    d = df_dong
    if value_filter is not None:
        d = d[d['code_no'].isin(value_filter)]
    p = (d.groupby([key, '연도'])['값']
           .sum(min_count=1)
           .unstack('연도'))
    return p


# ══════════════════════════════════════════════════════════════
#  ② 파생지표 (12개)  — 행정동 index Series 반환
# ══════════════════════════════════════════════════════════════
def _growth_vs_max(pivot_years, latest_year):
    """(당해값 - 전체연도 최댓값) / 최댓값 * 100.  최댓값 0 → 0 (IFERROR)."""
    latest = pivot_years.get(latest_year)
    if latest is None:                     # 기준연도 데이터 부재 → 0 (원시 불완전 방어)
        return pd.Series(0.0, index=pivot_years.index)
    mx = pivot_years.max(axis=1)
    out = (latest - mx) / mx * 100.0
    out = out.where(mx != 0, 0.0)     # max=0 → 0
    return out.fillna(0.0)


def derive_population_change(raw_to_in, key='행정동코드'):
    """인구변화율 = (2024 인구 - 30년 최다인구)/최다 *100."""
    g = aggregate_dong_year_code(raw_to_in, key)
    piv = _pivot_year(g, key=key)             # to_in_001 단일코드
    return _growth_vs_max(piv, C.YEAR_POP_LATEST)


def derive_age_indicators(raw_in_age, key='행정동코드'):
    """노년부양비 · 경제활동인구비율 · 소멸위험지수 (2024, in_age)."""
    g = aggregate_dong_year_code(raw_in_age, key)
    g = g[g['연도'] == C.YEAR_POP_LATEST]
    # 집계키 × 연령코드 합계
    m = (g.groupby([key, 'code_no'])['값'].sum(min_count=1)
           .unstack('code_no'))

    def s(codes):
        cols = [c for c in codes if c in m.columns]
        return m[cols].sum(axis=1, min_count=1).fillna(0.0)

    pop_total   = s(C.AGE_TOTAL_ALL)     # 001~021
    pop_working = s(C.AGE_WORKING)       # 004~013 (15~64)
    pop_elderly = s(C.AGE_ELDERLY)       # 014~021 (65+)
    fem_20_39   = s(C.AGE_FEM_20_39)     # 065~068

    def safe_div(num, den, scale=1.0):
        r = num / den * scale
        return r.where(den != 0, 0.0).fillna(0.0)   # IFERROR→0

    노년부양비   = safe_div(pop_elderly, pop_working, 100.0)
    경제활동비율 = safe_div(pop_working, pop_total,  100.0)
    소멸위험지수 = safe_div(fem_20_39,   pop_elderly, 1.0)
    return 노년부양비, 경제활동비율, 소멸위험지수


def derive_business_change(raw_to_fa, key='행정동코드'):
    """총사업체수 증감률 (to_fa_010, 2023 대비 전체 최다)."""
    g = aggregate_dong_year_code(raw_to_fa, key)
    piv = _pivot_year(g, key=key)
    return _growth_vs_max(piv, C.YEAR_BIZ_LATEST)


def derive_employment_changes(raw_cp_bem, key='행정동코드'):
    """총종사자·제조·고차·도소매·음식숙박 종사자수 증감률 (cp_bem)."""
    g = aggregate_dong_year_code(raw_cp_bem, key)
    out = {}
    for ind in ['총종사자', '제조업', '고차산업', '도소매', '음식숙박']:
        # 연도별 차수분기 코드로 필터하며 집계키×연도 합계 만들기
        parts = []
        for yr, sub in g.groupby('연도'):
            codes = C.industry_codes(ind, yr)          # None=전체합
            if codes is not None:
                sub = sub[sub['code_no'].isin(codes)]
            s = sub.groupby(key)['값'].sum(min_count=1)
            s.name = yr
            parts.append(s)
        if not parts:                                   # cp_bem 원시가 빈 경우 방어(빈 지표)
            out[ind] = pd.Series(dtype=float)
            continue
        piv = pd.concat(parts, axis=1)                 # 집계키 × 연도
        out[ind] = _growth_vs_max(piv, C.YEAR_BIZ_LATEST)
    return out


def derive_old_housing_ratio(raw_ho_yr, key='행정동코드'):
    """노후건축물비율 = Σ(2004이전 코드1~4) / Σ(전체) *100."""
    g = aggregate_dong_year_code(raw_ho_yr, key)
    g = g[g['연도'] == C.YEAR_POP_LATEST]
    m = g.groupby([key, 'code_no'])['값'].sum(min_count=1).unstack('code_no')
    old = m[[c for c in C.HO_YR_OLD_CODES if c in m.columns]].sum(axis=1, min_count=1).fillna(0.0)
    tot = m.sum(axis=1, min_count=1).fillna(0.0)
    r = (old / tot * 100.0)
    return r.where(tot != 0, 0.0).fillna(0.0)


def derive_small_housing_ratio(raw_ho_ar, key='행정동코드'):
    """소형주택비율 = Σ(60㎡이하 코드1~3) / Σ(전체) *100."""
    g = aggregate_dong_year_code(raw_ho_ar, key)
    g = g[g['연도'] == C.YEAR_POP_LATEST]
    m = g.groupby([key, 'code_no'])['값'].sum(min_count=1).unstack('code_no')
    small = m[[c for c in C.HO_AR_SMALL_CODES if c in m.columns]].sum(axis=1, min_count=1).fillna(0.0)
    tot = m.sum(axis=1, min_count=1).fillna(0.0)
    r = (small / tot * 100.0)
    return r.where(tot != 0, 0.0).fillna(0.0)


def derive_indicators(raw, key='행정동코드'):
    """
    6개 원시 long(dict) → (집계키) × 12지표 값 DataFrame.
    raw 키: to_in, in_age, to_fa, cp_bem, ho_yr, ho_ar
    key='행정동코드'(기본) 또는 '집계구'.
    """
    노년, 경제, 소멸 = derive_age_indicators(raw['in_age'], key)
    emp = derive_employment_changes(raw['cp_bem'], key)
    cols = {
        '인구변화율':      derive_population_change(raw['to_in'], key),
        '노년부양비':      노년,
        '경제활동인구비율': 경제,
        '소멸위험지수':    소멸,
        '총사업체수증감률': derive_business_change(raw['to_fa'], key),
        '총종사자수증감률': emp['총종사자'],
        '제조업증감률':    emp['제조업'],
        '고차산업증감률':  emp['고차산업'],
        '도소매증감률':    emp['도소매'],
        '음식숙박증감률':  emp['음식숙박'],
        '노후건축물비율':  derive_old_housing_ratio(raw['ho_yr'], key),
        '소형주택비율':    derive_small_housing_ratio(raw['ho_ar'], key),
    }
    df = pd.DataFrame(cols)
    return df[C.IND_IDS]        # 지표 순서 고정


# ══════════════════════════════════════════════════════════════
#  ③ 표준화 (Z, T)
# ══════════════════════════════════════════════════════════════
def standardize(values, sign):
    """
    Z = (값 - 평균) / 모표준편차 ,  T = Z*sign + 50
    · 평균/표준편차는 결측(NaN) 무시 (AGGREGATE 오류무시 등가)
    · std=0 이면 Z=0
    """
    v = values.astype(float)
    mean = np.nanmean(v.values)
    std  = np.nanstd(v.values, ddof=0)          # 모표준편차 STDEV.P
    if std == 0 or np.isnan(std):
        z = pd.Series(0.0, index=v.index)
    else:
        z = (v - mean) / std
    z = z.fillna(0.0)
    t = z * sign + 50.0
    return z, t, mean, std


def build_scores(ind_values):
    """지표값 DF → 값/Z/T DF (지표별 컬럼) + 통계(mean,std) dict."""
    zt = {}
    stats = {}
    for ind in C.IND_IDS:
        z, t, mean, std = standardize(ind_values[ind], C.SIGN[ind])
        zt[(ind, '값')] = ind_values[ind]
        zt[(ind, 'Z')]  = z
        zt[(ind, 'T')]  = t
        stats[ind] = {'mean': mean, 'std': std}
    out = pd.DataFrame(zt)
    out.columns = pd.MultiIndex.from_tuples(out.columns)
    return out, stats


# ══════════════════════════════════════════════════════════════
#  ④ 부문·종합
# ══════════════════════════════════════════════════════════════
def composite(scores):
    """부문별 Σ(T×가중치) 와 종합 Σ(부문). 반환: DataFrame(행정동 × 부문+종합)."""
    res = {}
    for sec in C.SECTORS:
        s = pd.Series(0.0, index=scores.index)
        for ind in C.IND_IDS:
            if C.SECTOR_OF[ind] == sec:
                s = s + scores[(ind, 'T')] * C.WEIGHT[ind]
        res[sec] = s
    df = pd.DataFrame(res)
    df['종합'] = df[C.SECTORS].sum(axis=1)
    return df


# ══════════════════════════════════════════════════════════════
#  ⑤ 등급 (선택: Jenks / Quantile / Pretty) — 10등급 기본
#     ※ 등급은 데이터에서 컷을 재산정하므로 골든본(고정컷)과 다를 수 있음(정상)
# ══════════════════════════════════════════════════════════════
def _jenks_breaks(data, n_classes):
    """
    Natural Breaks(Jenks) 경계 — numpy 자체 구현(동적계획법).
    데이터가 크지 않을 때(수백 개) 정확한 최적 분류를 제공.
    반환: 길이 n_classes+1 의 경계값 리스트.
    """
    v = np.sort(np.asarray(data, dtype=float))
    n = len(v)
    if n == 0:
        return []
    n_classes = min(n_classes, n)
    # 동적계획표
    mat1 = np.zeros((n + 1, n_classes + 1))
    mat2 = np.full((n + 1, n_classes + 1), np.inf)
    mat1[0, :] = 1
    mat2[0, :] = 0
    for i in range(1, n + 1):
        s1 = s2 = w = 0.0
        for l in range(1, i + 1):
            i3 = i - l + 1
            val = v[i3 - 1]
            s2 += val * val
            s1 += val
            w += 1
            variance = s2 - (s1 * s1) / w
            i4 = i3 - 1
            if i4 != 0:
                for j in range(2, n_classes + 1):
                    if mat2[i, j] >= variance + mat2[i4, j - 1]:
                        mat1[i, j] = i3
                        mat2[i, j] = variance + mat2[i4, j - 1]
        mat1[i, 1] = 1
        mat2[i, 1] = s2 - (s1 * s1) / w
    # 경계 역추적
    breaks = [0.0] * (n_classes + 1)
    breaks[n_classes] = v[-1]
    breaks[0] = v[0]
    k = n
    for j in range(n_classes, 1, -1):
        idx = int(mat1[k, j]) - 2
        breaks[j - 1] = v[idx]
        k = int(mat1[k, j]) - 1
    return breaks


def assign_grades(values, n_classes=10, method='jenks'):
    """
    종합점수 Series → 등급(1=쇠퇴심함 … n=양호).  method: jenks|quantile|pretty
    (골든본은 높은 종합점수=쇠퇴심함=1등급. 여기서는 값이 클수록 1등급.)
    """
    v = values.astype(float)
    x = v.dropna().values
    if method == 'quantile':
        qs = np.quantile(x, np.linspace(0, 1, n_classes + 1))
        breaks = list(qs)
    elif method == 'pretty':
        lo, hi = np.min(x), np.max(x)
        breaks = list(np.linspace(lo, hi, n_classes + 1))
    else:  # jenks
        breaks = _jenks_breaks(x, n_classes)
    breaks = np.unique(np.asarray(breaks, dtype=float))
    # 값 → 구간번호(낮은구간=1). 높은 종합점수가 쇠퇴 → 등급 반전(1=최고점)
    binno = np.digitize(v.values, breaks[1:-1], right=False) + 1
    grade = (len(breaks) - binno)   # 반전: 큰 값 → 작은 등급숫자
    return pd.Series(grade, index=v.index)


# ══════════════════════════════════════════════════════════════
#  전체 파이프라인 한 번에
# ══════════════════════════════════════════════════════════════
def run(raw, mapping=None, grade_method='jenks', n_classes=10, level='dong'):
    """원시 dict → (지표값·Z·T DF, 부문·종합 DF, 등급 Series, 통계 dict).
    level='dong'(행정동, 기본) 또는 'jgu'(집계구). 집계 레벨만 바뀌고 산식은 동일."""
    key = '집계구' if level == 'jgu' else '행정동코드'
    ind_values = derive_indicators(raw, key)
    scores, stats = build_scores(ind_values)
    comp = composite(scores)
    grades = assign_grades(comp['종합'], n_classes, grade_method)
    return scores, comp, grades, stats
