# -*- coding: utf-8 -*-
"""
code_audit.py — 코드표 안전장치 (③)
=====================================
기본 12지표는 SGIS 항목코드를 '의미'가 아니라 '위치(positional)'로 합산한다.
예) 경제활동가능인구 = in_age 코드 004~013 을 그대로 더함. 코드체계가 바뀌면
    (예: 5세계급 구간 재편) 조용히 틀린 값이 나온다.

이 모듈은 '엔진이 실제로 더할 코드'가 로드된 원시자료에 실제로 존재하는지 점검해
누락 코드를 경고로 알려준다(계산 자체는 그대로 — 표시·검증용).

핵심: config 의 코드 상수(AGE_WORKING 등)와 industry_codes() 를 그대로 신뢰원으로 사용.
"""
from __future__ import annotations

import pandas as pd

import config as C


# (bucket, 점검라벨, 기대코드 리스트, 연도종류 'pop'|'biz')
_CHECKS = [
    ("to_in",  "총인구(001)",                 [1],              "pop"),
    ("in_age", "총인구 연령합(001~021)",        C.AGE_TOTAL_ALL,  "pop"),
    ("in_age", "경제활동가능 15~64세(004~013)", C.AGE_WORKING,    "pop"),
    ("in_age", "65세+ 노년(014~021)",          C.AGE_ELDERLY,    "pop"),
    ("in_age", "20~39세 여성(065~068)",        C.AGE_FEM_20_39,  "pop"),
    ("to_fa",  "총괄사업체수(010)",             [10],             "biz"),
    ("ho_yr",  "노후주택 2004이전(001~004)",    C.HO_YR_OLD_CODES, "pop"),
    ("ho_ar",  "소형주택 60㎡이하(001~003)",    C.HO_AR_SMALL_CODES, "pop"),
]

_INDUSTRY = ["제조업", "도소매", "음식숙박", "고차산업"]

# 총인구 5세계급 '총' 계열의 알려진 코드 우주: 001~021(연령) + 022(미상).
# 남(031~052)·여(061~082)는 별도 계열이라 [1,30] 구간으로 '총' 계열만 격리 점검한다.
# 이 구간에 022를 넘는 코드가 새로 보이면 = SGIS 코드체계 변경 가능성 →
# AGE_TOTAL_ALL(001~021) 분모에서 조용히 누락될 위험.
_AGE_TOTAL_KNOWN = set(range(1, 23))     # 001~022
_AGE_TOTAL_RANGE = range(1, 31)          # 총 계열 격리 구간(남 031 이전)


def _present_codes(raw, bucket, year=None):
    """해당 분류(·연도)에 실제 존재하는 code_no(코드 끝 3자리 정수) 집합."""
    df = raw.get(bucket)
    if df is None or len(df) == 0 or "CODE" not in df.columns:
        return set()
    d = df
    if year is not None and "연도" in d.columns:
        yr = pd.to_numeric(d["연도"], errors="coerce")
        d = d[yr == year]
    out = set()
    for c in d["CODE"].astype(str).str[-3:]:
        try:
            out.add(int(c))
        except ValueError:
            continue
    return out


def _fmt(codes):
    return ", ".join(f"{c:03d}" for c in codes)


def audit(raw, pop_year=None, biz_year=None):
    """원시 dict → (경고 메시지 list, 상세 DataFrame).
    pop_year/biz_year 미지정 시 config 기본값(기준연도) 사용."""
    pop_year = int(pop_year if pop_year is not None else C.YEAR_POP_LATEST)
    biz_year = int(biz_year if biz_year is not None else C.YEAR_BIZ_LATEST)

    warns, rows = [], []

    def add_row(bucket, label, status, missing):
        rows.append({"분류": bucket, "점검 항목": label, "상태": status,
                     "누락 코드": _fmt(missing) if missing else ""})

    for bucket, label, codes, kind in _CHECKS:
        yr = pop_year if kind == "pop" else biz_year
        present = _present_codes(raw, bucket, yr)
        if not present:
            add_row(bucket, label, "데이터없음", [])
            continue
        missing = sorted(set(codes) - present)
        if missing:
            add_row(bucket, label, "⚠ 누락", missing)
            warns.append(f"{bucket} · {label}: {yr}년 코드 {_fmt(missing)} 없음 "
                         f"→ 해당 지표 과소계산 위험(코드체계 확인).")
        else:
            add_row(bucket, label, "정상", [])

    # 산업(cp_bem) — 기준연도 차수(8/9/10차)에 맞는 코드가 실제 있는지
    present_biz = _present_codes(raw, "cp_bem", biz_year)
    for ind in _INDUSTRY:
        codes = C.industry_codes(ind, biz_year)
        if codes is None:
            continue
        label = f"{ind} 종사자({biz_year}년 차수코드)"
        if not present_biz:
            add_row("cp_bem", label, "데이터없음", [])
            continue
        missing = sorted(set(codes) - present_biz)
        if missing:
            add_row("cp_bem", label, "⚠ 누락", missing)
            warns.append(f"cp_bem · {ind}: {biz_year}년 코드 {_fmt(missing)} 없음 "
                         f"→ 과소계산 위험(8/9/10차 코드 확인).")
        else:
            add_row("cp_bem", label, "정상", [])

    # 예상 밖 코드(총인구 연령계열) — 새 코드가 생기면 비율 분모에서 조용히 누락됨
    present_age = _present_codes(raw, "in_age", pop_year)
    unexpected = sorted(c for c in present_age
                        if c in _AGE_TOTAL_RANGE and c not in _AGE_TOTAL_KNOWN)
    if unexpected:
        add_row("in_age", "총인구 연령계열 예상 밖 코드", "⚠ 신규", unexpected)
        warns.append(f"in_age · 총인구 연령계열에 미등록 코드 {_fmt(unexpected)} 발견 "
                     f"→ SGIS 코드체계 변경 가능성. AGE_TOTAL_ALL(001~021) 분모에서 "
                     f"누락되어 비율이 과대계산될 수 있음(config 확인).")

    return warns, pd.DataFrame(rows, columns=["분류", "점검 항목", "상태", "누락 코드"])


def assumptions_table(raw, pop_year=None, biz_year=None):
    """엔진이 각 지표를 계산할 때 '어떤 코드를 더한다고 가정'하는지 사람이 읽는 표.
    데이터에 코드 '의미(라벨)'가 없으므로, 이 표로 사람이 눈으로 코드↔뜻을 검증한다.
    반환: DataFrame[지표, 계산(합산 코드), 코드의 뜻, 기준연도, 데이터 상태]."""
    pop_year = int(pop_year if pop_year is not None else C.YEAR_POP_LATEST)
    biz_year = int(biz_year if biz_year is not None else C.YEAR_BIZ_LATEST)
    pres_pop_in = _present_codes(raw, "in_age", pop_year)
    pres_pop_toin = _present_codes(raw, "to_in", pop_year)
    pres_pop_hoyr = _present_codes(raw, "ho_yr", pop_year)
    pres_fa = _present_codes(raw, "to_fa", biz_year)
    pres_biz = _present_codes(raw, "cp_bem", biz_year)

    def _stat(codes, present):
        if not present:
            return "데이터없음"
        miss = sorted(set(codes) - present)
        return "정상" if not miss else "⚠ 누락 " + _fmt(miss)

    rows = []

    def add(ind, codes, present, meaning, year):
        rows.append({"지표": ind, "합산 코드": _fmt(sorted(codes)),
                     "코드의 뜻": meaning, "기준연도": year,
                     "데이터 상태": _stat(codes, present)})

    add("인구변화율", [1], pres_pop_toin, "총인구(to_in 001), 기준연도 대비 30년 최다와 비교", pop_year)
    add("노년부양비", C.AGE_ELDERLY, pres_pop_in, "65세+(014~021) ÷ 15~64세(004~013) ×100", pop_year)
    add("경제활동인구비율", C.AGE_WORKING, pres_pop_in, "15~64세(004~013) ÷ 총인구(001~021) ×100", pop_year)
    add("소멸위험지수", C.AGE_FEM_20_39, pres_pop_in, "20~39세 여성(065~068) ÷ 65세+(014~021)", pop_year)
    add("총사업체수증감률", [10], pres_fa, "총괄사업체수(to_fa 010) 전체기간 최다 대비", biz_year)
    add("총종사자수증감률", sorted(pres_biz) or [0], pres_biz, "cp_bem 전체 코드 합(차수 무관 전체합)", biz_year)
    for ind in _INDUSTRY:
        codes = C.industry_codes(ind, biz_year)
        if ind == "도소매":
            meaning = "cp_bem 전 차수 007 고정"
        else:
            cha = "8차(≤2005년)" if biz_year <= 2005 else "9·10차(2006년~)"
            meaning = f"cp_bem {cha} 코드 합"
        add(f"{ind}증감률", codes, pres_biz, meaning, biz_year)
    add("노후건축물비율", C.HO_YR_OLD_CODES, pres_pop_hoyr, "2004년 이전 건축(001~004) ÷ 전체 건축연도", pop_year)

    return pd.DataFrame(rows, columns=["지표", "합산 코드", "코드의 뜻", "기준연도", "데이터 상태"])
