# -*- coding: utf-8 -*-
"""
batch_build.py — 전국 시군구 '디폴트' 배치 빌드
=================================================
여러 시군구의 원시데이터(raw dict)를 **시군구코드(집계구 앞 5자리)로 쪼개**,
각 시군구를 **독립적으로** 표준화·진단해(그 시군구 행정동/집계구 안에서 Z·T)
시군구별 xlsx 한 개씩 만들고, 전체를 하나의 zip으로 묶는다.

핵심 설계(사용자 확정):
  · 표준화 기준 = 시군구별 독립 (골든본 전주와 동일 철학)
  · 지표/가중치 = 기본값(config) 그대로, 연도만 선택
  · 산출 = 집계구 + 행정동(집계구 합) — export.build_integrated_workbook 재사용

의존: pandas, config, decline_engine, legal_engine, export (모두 순수/기존 모듈).
"""
from __future__ import annotations

import io
import re
import zipfile

import pandas as pd

import config as C
import decline_engine as E
import legal_engine as LG
import export

BUCKETS = ["to_in", "in_age", "to_fa", "cp_bem", "ho_yr", "ho_ar"]
_EMPTY_COLS = ["연도", "집계구", "CODE", "값", "행정동코드"]


def _sgg5(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str[:5]


def list_sigungu(raw: dict) -> list[str]:
    """raw dict 안에 존재하는 시군구코드(5자리) 목록(정렬)."""
    codes = set()
    for b in BUCKETS:
        df = raw.get(b)
        if df is not None and len(df) and "집계구" in df.columns:
            codes.update(_sgg5(df["집계구"]).tolist())
    return sorted(c for c in codes if len(c) == 5 and c.isdigit())


def split_raw_by_sigungu(raw: dict) -> dict[str, dict]:
    """raw dict → {시군구코드: raw_subset(6버킷)}. 각 버킷은 그 시군구 행만."""
    out: dict[str, dict] = {}
    for sgg in list_sigungu(raw):
        sub = {}
        for b in BUCKETS:
            df = raw.get(b)
            if df is not None and len(df) and "집계구" in df.columns:
                m = _sgg5(df["집계구"]) == sgg
                sub[b] = df[m].copy()
            else:
                sub[b] = pd.DataFrame(columns=_EMPTY_COLS)
        out[sgg] = sub
    return out


def _ensure_buckets(raw_sub: dict) -> dict:
    for b in BUCKETS:
        if b not in raw_sub or raw_sub[b] is None:
            raw_sub[b] = pd.DataFrame(columns=_EMPTY_COLS)
    return raw_sub


def build_one_workbook(raw_sub: dict, *, name_map=None, method="jenks", n_classes=10,
                       decimals=2, final_only=True, selected_years=None):
    """한 시군구 raw_subset → openpyxl Workbook (기본 지표·가중치).
    각 시군구를 독립 표준화(그 시군구 단위집합 안에서 Z·T)."""
    raw_sub = _ensure_buckets(raw_sub)
    dong = E.run(raw_sub, level="dong", grade_method=method, n_classes=int(n_classes))
    jgu = E.run(raw_sub, level="jgu", grade_method=method, n_classes=int(n_classes))
    legal_dong = LG.run_legal(raw_sub, level="dong")
    legal_jgu = LG.run_legal(raw_sub, level="jgu")
    wb = export.build_integrated_workbook(
        raw_sub, selected_years=selected_years, name_map=name_map,
        dong_res=dong[:3], jgu_res=jgu[:3],
        legal_dong=legal_dong, legal_jgu=legal_jgu,
        n_classes=int(n_classes), method=method,
        decimals=decimals, pivot_level="both", final_only=final_only)
    n_decl = int((legal_dong["쇠퇴지역"] == "o").sum()) if len(legal_dong) else 0
    stats = {"n_dong": len(dong[0]), "n_jgu": len(jgu[0]), "n_decl": n_decl}
    return wb, stats


def _safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(s)).strip() or "sigungu"


def build_batch_zip(raw: dict, *, sigungu=None, name_map=None, sido_name_map=None,
                    method="jenks", n_classes=10, decimals=2, final_only=True,
                    selected_years=None, year_pop=None, year_biz=None, progress=None):
    """여러 시군구 raw → (zip_bytes, 요약 DataFrame).
    sigungu: 처리할 시군구코드 리스트(None=raw 안 전체).
    year_pop/year_biz: 기준연도(엔진 전역에 설정). None이면 config 현재값 유지.
    progress(done, total, sgg): 진행 콜백(선택)."""
    if year_pop is not None:
        C.YEAR_POP_LATEST = int(year_pop)
    if year_biz is not None:
        C.YEAR_BIZ_LATEST = int(year_biz)

    # 엔진에 넘기기 전 선택 연도로 먼저 필터(메인 파이프라인과 동일)
    if selected_years:
        import sheet_builder as _sb
        raw = _sb.filter_raw_years(raw, selected_years)

    parts = split_raw_by_sigungu(raw)
    codes = [c for c in (sigungu or list(parts.keys())) if c in parts]
    total = len(codes)
    rows = []
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, sgg in enumerate(codes):
            try:
                wb, stats = build_one_workbook(
                    parts[sgg], name_map=name_map, method=method, n_classes=n_classes,
                    decimals=decimals, final_only=final_only, selected_years=selected_years)
                wbuf = io.BytesIO(); wb.save(wbuf); wbuf.seek(0)
                sname = (sido_name_map or {}).get(sgg[:2], "")
                fname = _safe_name(f"{sgg}_{sname}_쇠퇴진단.xlsx")
                zf.writestr(fname, wbuf.getvalue())
                rows.append({"시군구코드": sgg, "시도": sname, "파일": fname,
                             "행정동수": stats["n_dong"], "집계구수": stats["n_jgu"],
                             "법적쇠퇴행정동": stats["n_decl"], "상태": "OK"})
            except Exception as e:  # 한 시군구 실패가 전체를 막지 않게 격리
                rows.append({"시군구코드": sgg, "시도": (sido_name_map or {}).get(sgg[:2], ""),
                             "파일": "", "행정동수": 0, "집계구수": 0,
                             "법적쇠퇴행정동": 0, "상태": f"실패: {e}"})
            if progress:
                progress(i + 1, total, sgg)
        summary = pd.DataFrame(rows)
        # 요약 CSV도 zip에 포함(엑셀에서 바로 열림, cp949)
        try:
            zf.writestr("_요약.csv", summary.to_csv(index=False).encode("utf-8-sig"))
        except Exception:
            pass
    buf.seek(0)
    return buf.getvalue(), summary
