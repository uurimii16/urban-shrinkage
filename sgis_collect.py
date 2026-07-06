# -*- coding: utf-8 -*-
"""
sgis_collect.py — SGIS OpenAPI 시계열 자동수집 (하이브리드의 API 파트)
======================================================================
연도범위 + 지역(시군구)을 주면 SGIS OpenAPI로 아래 '시계열' 원시를 수집해
loader/엔진이 그대로 먹는 raw dict(4열+행정동코드)로 돌려준다.

담당 버킷(시계열 증감률용):
  · to_in   : 인구총괄 tot_ppltn        → to_in_001
  · to_fa   : 사업체 corp_cnt(총사업체) → to_fa_010
  · cp_bem  : 산업 대분류별 종사자수     → cp_bem_{대분류순번:03d}
              (대분류 순번 = 엔진 config.industry_codes 코드와 차수별 일치 — 실측 확인)

성연령(in_age)·건축연도(ho_yr)·연면적(ho_ar)은 API가 정밀도/코드가 불명확하여
이 수집기 대상이 아니다 → 그 2024 스냅샷은 파일 업로드로 보완(하이브리드).

adm_cd는 8자리(=행정동코드)로 나오므로 loader의 8자리 처리와 그대로 맞물린다.
표준 라이브러리만 사용(pip 불필요).
"""
from __future__ import annotations
import urllib.parse
import urllib.request
import json

import pandas as pd

BASE = "https://sgisapi.kostat.go.kr/OpenAPI3"


def _get(path, params, timeout=30):
    q = urllib.parse.urlencode(params)
    with urllib.request.urlopen(f"{BASE}/{path}?{q}", timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def authenticate(consumer_key, consumer_secret):
    j = _get("auth/authentication.json",
             {"consumer_key": consumer_key, "consumer_secret": consumer_secret})
    tok = (j.get("result") or {}).get("accessToken")
    if not tok:
        raise RuntimeError(f"SGIS 인증 실패: {j.get('errMsg') or j}")
    return tok


def class_deg_of(year: int) -> str:
    """산업분류 차수: 2005이전=8, 2006~2015=9, 2016이후=10 (config.industry_codes와 동일 경계)."""
    return "8" if year <= 2005 else ("9" if year <= 2015 else "10")


def _num(v):
    """'12,345'/'N/A'/'' → float 또는 None."""
    if v is None:
        return None
    s = str(v).replace(",", "").strip()
    if s in ("", "N/A", "-", "null", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _rows_from(result, code, year, out):
    """API result(list) → out 리스트에 (연도, 집계구, CODE, 값, 행정동코드) 추가.
    adm_cd(8자리)를 집계구·행정동코드 양쪽에 넣는다(행정동 단위 데이터)."""
    if not isinstance(result, list):
        return
    for r in result:
        adm = str(r.get("adm_cd", "")).strip()
        val = _num(r.get(_FIELD[code_kind(code)]))
        if not adm or val is None:
            continue
        out.append({"연도": year, "집계구": adm, "CODE": code, "값": val, "행정동코드": adm})


# CODE → 어떤 응답필드를 값으로 쓰는가
_FIELD = {"pop": "tot_ppltn", "corp": "corp_cnt", "worker": "tot_worker"}


def code_kind(code):
    if code.startswith("to_in"):
        return "pop"
    if code.startswith("to_fa"):
        return "corp"
    return "worker"


def collect_raw(token, sigungu_codes, years, progress=None):
    """token + 시군구코드 리스트 + 연도 리스트 → 엔진 raw dict(to_in/to_fa/cp_bem).
    progress(msg): 진행상황 콜백(선택)."""
    sigungu_codes = [str(s).strip() for s in sigungu_codes]
    years = sorted(int(y) for y in years)
    rows = []
    ind_cache = {}   # class_deg -> [대분류 letter,...] (순서=코드순번)

    def log(m):
        if progress:
            progress(m)

    for y in years:
        deg = class_deg_of(y)
        if deg not in ind_cache:
            res = _get("stats/industrycode.json",
                       {"accessToken": token, "class_deg": deg}).get("result") or []
            letters = [str(r.get("class_code")) for r in res
                       if len(str(r.get("class_code", ""))) == 1]
            ind_cache[deg] = letters
        letters = ind_cache[deg]

        for sg in sigungu_codes:
            base = {"accessToken": token, "adm_cd": sg, "low_search": "1"}
            # 총인구 → to_in_001
            try:
                r = _get("stats/population.json", {**base, "year": str(y)}).get("result")
                _rows_from(r, "to_in_001", y, rows)
            except Exception as e:
                log(f"  [경고] {y} {sg} 인구: {e}")
            # 총사업체수 → to_fa_010  (class_code 없이 총계)
            try:
                r = _get("stats/company.json", {**base, "year": str(y)}).get("result")
                _rows_from(r, "to_fa_010", y, rows)
            except Exception as e:
                log(f"  [경고] {y} {sg} 사업체: {e}")
            # 산업 대분류별 종사자수 → cp_bem_{순번:03d}
            for i, letter in enumerate(letters, 1):
                try:
                    r = _get("stats/company.json",
                             {**base, "year": str(y), "class_code": letter}).get("result")
                    _rows_from(r, f"cp_bem_{i:03d}", y, rows)
                except Exception as e:
                    log(f"  [경고] {y} {sg} 종사자 {letter}: {e}")
            log(f"  {y}년 {sg} 수집 완료")

    df = pd.DataFrame(rows, columns=["연도", "집계구", "CODE", "값", "행정동코드"])
    # 엔진 raw dict(6버킷) — 없는 버킷은 빈 DataFrame(파일 업로드로 채움)
    raw = {}
    for b in ("to_in", "in_age", "to_fa", "cp_bem", "ho_yr", "ho_ar"):
        sub = df[df["CODE"].str.startswith(b)].reset_index(drop=True) if len(df) else \
            pd.DataFrame(columns=["연도", "집계구", "CODE", "값", "행정동코드"])
        raw[b] = sub
    return raw


def collect(consumer_key, consumer_secret, sigungu_codes, year_from, year_to, progress=None):
    """상위 진입점: 키+지역+연도범위 → raw dict."""
    token = authenticate(consumer_key, consumer_secret)
    years = list(range(int(year_from), int(year_to) + 1))
    return collect_raw(token, sigungu_codes, years, progress=progress)


if __name__ == "__main__":
    import sys
    key, secret = sys.argv[1], sys.argv[2]
    raw = collect(key, secret, ["35011", "35012"], 2022, 2023, progress=print)
    for b, v in raw.items():
        print(f"{b:7s}: {len(v):>6,}행  코드종류={sorted(v['CODE'].unique())[:6] if len(v) else []}")
