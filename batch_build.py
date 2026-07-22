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
import os
import re
import zipfile

import pandas as pd

import config as C
import decline_engine as E
import legal_engine as LG
import export
import template_export as TE   # save_wb(캐시 주입 저장) — build_batch_zip·stream_sigungu_templates에서 사용

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


# 특별시·광역시(시도코드): 자치구·군을 통째로 '그 시' 하나로 묶는다.
GWANGYEOK = {"11", "21", "22", "23", "24", "25", "26"}


def city_key(sgg5: str) -> str:
    """시군구코드(5자리) → '시 단위' 그룹키.
    - 특별/광역시(11·21·22·23·24·25·26): 자치구 통째로 시도 단위(예 서울 11xxx→'11000').
    - 그 외(도·세종): 일반구(끝자리≠0)를 부모 시로(끝자리→0). 예 전주 35011·35012→'35010'.
      단독 시/군(끝자리 0)은 자기 자신 유지. 세종 29010→'29010'."""
    sd = sgg5[:2]
    if sd in GWANGYEOK:
        return sd + "000"
    return sgg5[:4] + "0"


def split_raw_by_city(raw: dict) -> dict[str, dict]:
    """raw dict → {시 단위 그룹키: raw_subset}. 각 '시' 안 집계구를 다 모아
    그 시 전체(행정동 집합)에서 표준화하게 한다(광역시 통째·일반구는 시로 합침)."""
    out: dict[str, dict] = {}
    for b in BUCKETS:
        df = raw.get(b)
        if df is None or not len(df) or "집계구" not in df.columns:
            continue
        keys = _sgg5(df["집계구"]).map(city_key)
        for ck, sub in df.groupby(keys):
            out.setdefault(ck, {})[b] = sub.copy()
    # 빈 버킷 채우기
    for ck in out:
        for b in BUCKETS:
            out[ck].setdefault(b, pd.DataFrame(columns=_EMPTY_COLS))
    return out


def _ensure_buckets(raw_sub: dict) -> dict:
    for b in BUCKETS:
        if b not in raw_sub or raw_sub[b] is None:
            raw_sub[b] = pd.DataFrame(columns=_EMPTY_COLS)
    return raw_sub


def build_one_workbook(raw_sub: dict, *, name_map=None, method="jenks", n_classes=10,
                       decimals=2, final_only=False, formula_mode=True, selected_years=None):
    """한 시군구 raw_subset → openpyxl Workbook (기본 지표·가중치).
    각 시군구를 독립 표준화(그 시군구 단위집합 안에서 Z·T).
    기본값: 전체 시트(final_only=False) + 함수 포함(formula_mode=True) + 집계구·행정동(both).
    지표·가중치·방향·라벨·부문은 config 기본값(export가 None→C.* 로 채움)."""
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
        decimals=decimals, pivot_level="both", final_only=final_only,
        formula_mode=formula_mode)
    n_decl = int((legal_dong["쇠퇴지역"] == "o").sum()) if len(legal_dong) else 0
    stats = {"n_dong": len(dong[0]), "n_jgu": len(jgu[0]), "n_decl": n_decl}
    return wb, stats


def build_one_template(raw_sub: dict, *, indicators, custom_df=None, recipes=None,
                       admin_path=None, result_only=False):
    """한 시군구 raw_subset → 정본 양식 Workbook.
    ③설정 가중치·지표목록(indicators=template_export.indicators_from_cfg(cfg))을 그대로 쓰고,
    base+커스텀+계산식 집계구 지표값을 채운다. 각 시군구는 그 시군구 집계구 집합 안에서
    독립 표준화(정본 종합의 Z/T가 그 파일 자기 행들로 함수 계산).
    result_only=True: **결과만(계산방법 + 복합쇠퇴진단 종합 2시트)** — 원시 6시트(15만 행 종사자 등)를
      빼서 메모리를 대폭 줄임. 큰 지역도 안 터지고, 지표값·Z·T·부문·종합은 그대로 함수+계산값."""
    import template_export as TE
    import custom_indicators as CI
    import recipe_engine as RE
    raw_sub = _ensure_buckets(raw_sub)
    ids = [i[0] for i in indicators]
    scores = E.run(raw_sub, level="jgu")[0]          # 집계구 base scores((지표,'값'/'Z'/'T'))
    idx = scores.index
    if custom_df is not None and len(custom_df):
        scores = CI.combine_scores(scores, CI.build_scores(custom_df, idx, "jgu"))
    if recipes:
        scores = CI.combine_scores(scores, RE.build_recipe_scores(recipes, raw_sub, "jgu", idx))
    values = TE.values_from_scores(scores, ids)       # base+커스텀+계산식 지표값
    if result_only:                                   # 결과만 2시트(원시시트 제외) — 가벼움
        wb = TE.build_composite_workbook(values=values, indicators=indicators, admin_path=admin_path)
    else:
        wb = TE.build_full_workbook(raw_sub, values=values, indicators=indicators,
                                    admin_path=admin_path)   # 원본 9시트 전부
    n_dong = len({str(c)[:8] for c in idx})
    stats = {"n_dong": n_dong, "n_jgu": len(idx), "n_decl": 0}
    return wb, stats


def _safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", str(s)).strip() or "sigungu"


def build_batch_zip(raw: dict, *, sigungu=None, name_map=None, sido_name_map=None,
                    method="jenks", n_classes=10, decimals=2, final_only=False,
                    formula_mode=True, selected_years=None, year_pop=None, year_biz=None,
                    out_dir=None, progress=None, group="sigungu",
                    template_mode=False, indicators=None, custom_df=None, recipes=None,
                    admin_path=None, result_only=False):
    """여러 시군구 raw → (zip_bytes, 요약 DataFrame).
    sigungu: 처리할 시군구코드 리스트(None=raw 안 전체).
    year_pop/year_biz: 기준연도(엔진 전역에 설정). None이면 config 현재값 유지.
    out_dir: 지정하면 각 시군구 xlsx를 그 폴더에 '착착' 저장(로컬 실행용). zip은 항상 반환.
    progress(done, total, sgg): 진행 콜백(선택).
    template_mode=True: 각 시군구를 ③설정 가중치(indicators)로 '정본 양식'(계산방법+복합종합)으로
      산출(build_one_template). custom_df/recipes로 커스텀·계산식 지표값도 반영. False면 기존 구양식."""
    if out_dir:
        try:                              # 저장 폴더는 부가기능 — 권한/경로 실패해도 zip 산출은 계속
            os.makedirs(out_dir, exist_ok=True)
        except Exception:
            out_dir = None
    if year_pop is not None:
        C.YEAR_POP_LATEST = int(year_pop)
    if year_biz is not None:
        C.YEAR_BIZ_LATEST = int(year_biz)

    # 엔진에 넘기기 전 선택 연도로 먼저 필터(메인 파이프라인과 동일)
    if selected_years:
        import sheet_builder as _sb
        raw = _sb.filter_raw_years(raw, selected_years)

    parts = split_raw_by_city(raw) if group == "city" else split_raw_by_sigungu(raw)
    codes = [c for c in (sigungu or list(parts.keys())) if c in parts]
    total = len(codes)
    rows = []
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, sgg in enumerate(codes):
            try:
                if template_mode:
                    wb, stats = build_one_template(
                        parts[sgg], indicators=indicators, custom_df=custom_df,
                        recipes=recipes, admin_path=admin_path, result_only=result_only)
                else:
                    wb, stats = build_one_workbook(
                        parts[sgg], name_map=name_map, method=method, n_classes=n_classes,
                        decimals=decimals, final_only=final_only, formula_mode=formula_mode,
                        selected_years=selected_years)
                wbuf = io.BytesIO(); TE.save_wb(wb, wbuf); wbuf.seek(0)
                sname = (sido_name_map or {}).get(sgg[:2], "")
                fname = _safe_name(f"{sgg}_{sname}_쇠퇴진단.xlsx")
                zf.writestr(fname, wbuf.getvalue())
                if out_dir:      # 로컬 저장 폴더에 착착 저장(실패해도 zip엔 이미 들어감)
                    try:
                        with open(os.path.join(out_dir, fname), "wb") as fh:
                            fh.write(wbuf.getvalue())
                    except Exception:
                        pass
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
        # 요약 CSV도 zip에 포함(엑셀에서 바로 열림, utf-8-sig)
        try:
            csv_bytes = summary.to_csv(index=False).encode("utf-8-sig")
            zf.writestr("_요약.csv", csv_bytes)
            if out_dir:
                with open(os.path.join(out_dir, "_요약.csv"), "wb") as fh:
                    fh.write(csv_bytes)
        except Exception:
            pass
    buf.seek(0)
    return buf.getvalue(), summary


def _iter_named_bytes(files):
    """[(name, bytes), ...] → (내부파일명, bytes) 제너레이터. .zip은 자동 해제해 내부 csv/txt를 펼침."""
    for name, data in files or []:
        if str(name).lower().endswith(".zip"):
            try:
                zf = zipfile.ZipFile(io.BytesIO(data))
            except Exception:
                continue
            for zi in zf.namelist():
                if zi.endswith("/") or not zi.lower().endswith((".csv", ".txt")):
                    continue
                try:
                    yield (os.path.basename(zi) or zi), zf.read(zi)
                except Exception:
                    continue
        else:
            yield name, data


# ════════════════════════════════════════════════════════════════════════════
# 청크(chunk) 처리용 — 웹에서 시군구를 몇 개씩 나눠 빌드해 타임아웃 방지
#   spool_files: 다운로드 파일을 시군구별 pkl로 스풀(누적 가능). 메모리 안전.
#   build_sigungu_from_pkls: 스풀된 pkl → 시군구 1개 xlsx(정본 또는 구양식).
# app_v2가 이 둘을 이용해 '준비(스풀) → K개씩 빌드 → zip' 파이프라인을 st.rerun으로 굴린다.
# ════════════════════════════════════════════════════════════════════════════
def spool_files(files, tmp, *, sigungu=None, progress=None, sgg_parts=None, year_meta=None):
    """파일들을 읽어 시군구코드(집계구 앞5)별 pkl로 저장. {sgg:[pkl경로]} 반환.
    sgg_parts를 넘기면 그 dict에 '누적'(파일을 하나씩 다운로드하며 반복 호출 가능 → 메모리 안전).
    year_meta(dict)를 넘기면 인구/사업체 기준연도 최댓값을 감지해 {'pop':int,'biz':int}로 채운다.
    progress(n, None, name, 'spool')."""
    import loader as L
    if sgg_parts is None:
        sgg_parts = {}
    nfiles = 0
    for name, data in _iter_named_bytes(files):
        if L._is_corrupt(name):
            continue
        df = L._read_csv_bytes(data)
        if df is None or len(df) == 0:
            continue
        nfiles += 1
        if progress:
            progress(nfiles, None, name, "spool")
        if year_meta is not None:            # 파일명 아닌 CODE 접두어로 부문 판별 → 최신연도 감지
            try:
                ymax = pd.to_numeric(df["연도"], errors="coerce").max()
                if pd.notna(ymax):
                    b = L._bucket_of(str(df["CODE"].iloc[0]).strip())
                    key = "biz" if b in ("to_fa", "cp_bem") else "pop"
                    year_meta[key] = max(year_meta.get(key, 0), int(ymax))
            except Exception:
                pass
        df["집계구"] = df["집계구"].astype(str).str.strip()
        df = df[df["집계구"].str.len().isin([8, 14])]     # 롤업행(2·5자리) 제거
        if len(df) == 0:
            continue
        df["_sgg"] = df["집계구"].str[:5]
        for sgg, sub in df.groupby("_sgg"):
            if sigungu and sgg not in sigungu:
                continue
            d = os.path.join(tmp, sgg)
            os.makedirs(d, exist_ok=True)
            p = os.path.join(d, f"{len(sgg_parts.get(sgg, []))}.pkl")
            sub.drop(columns="_sgg").to_pickle(p)
            sgg_parts.setdefault(sgg, []).append(p)
        del df
    return sgg_parts


def build_sigungu_from_pkls(sgg, pkls, *, template_mode=True, indicators=None,
                            custom_df=None, recipes=None, admin_path=None,
                            name_map=None, sido_name_map=None, selected_years=None,
                            method="jenks", n_classes=10, decimals=2,
                            final_only=False, formula_mode=True, result_only=False):
    """스풀된 pkl들 → 시군구 1개 xlsx. (fname, xlsx_bytes, 요약row) 반환.
    실패해도 예외 대신 (None, None, 실패row)로 반환(청크 루프가 안 끊기게)."""
    import loader as L
    import sheet_builder as SB
    sname = (sido_name_map or {}).get(sgg[:2], "")
    try:
        frames = [pd.read_pickle(p) for p in pkls]
        raw_sub = L._build_raw_from_rows(frames)
        del frames
        if selected_years:
            raw_sub = SB.filter_raw_years(raw_sub, selected_years)
        if template_mode:
            wb, stats = build_one_template(raw_sub, indicators=indicators,
                                           custom_df=custom_df, recipes=recipes,
                                           admin_path=admin_path, result_only=result_only)
        else:
            wb, stats = build_one_workbook(raw_sub, name_map=name_map, method=method,
                                           n_classes=n_classes, decimals=decimals,
                                           final_only=final_only, formula_mode=formula_mode,
                                           selected_years=selected_years)
        del raw_sub
        wbuf = io.BytesIO(); TE.save_wb(wb, wbuf); wbuf.seek(0)
        fname = _safe_name(f"{sgg}_{sname}_쇠퇴진단.xlsx")
        row = {"시군구코드": sgg, "시도": sname, "파일": fname,
               "집계구수": stats.get("n_jgu", 0), "상태": "OK"}
        return fname, wbuf.getvalue(), row
    except Exception as e:
        return None, None, {"시군구코드": sgg, "시도": sname, "파일": "",
                            "집계구수": 0, "상태": f"실패: {e}"}


def stream_sigungu_templates(files, *, indicators, custom_df=None, recipes=None,
                             admin_path=None, sido_name_map=None, selected_years=None,
                             sigungu=None, out_dir=None, progress=None, tmp_dir=None):
    """다운로드 파일들을 '시군구별로 하나씩' 처리해 정본 양식(계산방법+복합종합) xlsx를 만들어 zip.
    전체 raw를 한 번에 메모리에 안 올려 큰 지역(서울 25구 등)도 안전.
      1단계(스풀): 파일 하나씩 읽어 → 시군구코드(집계구 앞5)로 쪼개 → 디스크에 저장 → 파일 버림.
      2단계(처리): 시군구 하나씩만 불러와 → 정본 산출 → zip 추가 → 임시파일 삭제 → 진행바.
    files: [(name, bytes), ...] (zip 자동 해제). 반환: (zip_bytes, 요약 DataFrame)."""
    import shutil
    import tempfile
    import loader as L
    import sheet_builder as SB

    tmp = tmp_dir or tempfile.mkdtemp(prefix="sgg_stream_")
    if out_dir:   # 로컬 저장 폴더는 '있으면 좋은' 부가기능 → 실패해도 zip 산출은 계속
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception:
            out_dir = None

    # ── 1단계: 스풀 (파일 하나씩 → 시군구별 pkl append) ──
    sgg_parts: dict[str, list[str]] = {}
    nfiles = 0
    for name, data in _iter_named_bytes(files):
        if L._is_corrupt(name):
            continue
        df = L._read_csv_bytes(data)
        if df is None or len(df) == 0:
            continue
        nfiles += 1
        if progress:
            progress(nfiles, None, name, "spool")   # 정제 진행(총 개수는 미정 → None)
        df["집계구"] = df["집계구"].astype(str).str.strip()
        df = df[df["집계구"].str.len().isin([8, 14])]      # 롤업행(2·5자리) 제거
        if len(df) == 0:
            continue
        df["_sgg"] = df["집계구"].str[:5]
        for sgg, sub in df.groupby("_sgg"):
            if sigungu and sgg not in sigungu:
                continue
            d = os.path.join(tmp, sgg)
            os.makedirs(d, exist_ok=True)
            p = os.path.join(d, f"{len(sgg_parts.get(sgg, []))}.pkl")
            sub.drop(columns="_sgg").to_pickle(p)
            sgg_parts.setdefault(sgg, []).append(p)
        del df

    codes = sorted(sgg_parts.keys())
    if sigungu:
        codes = [c for c in codes if c in sigungu]

    # ── 2단계: 시군구 하나씩 정본 산출 ──
    rows, total = [], len(codes)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, sgg in enumerate(codes):
            try:
                frames = [pd.read_pickle(p) for p in sgg_parts[sgg]]
                raw_sub = L._build_raw_from_rows(frames)
                del frames
                if selected_years:
                    raw_sub = SB.filter_raw_years(raw_sub, selected_years)
                wb, stats = build_one_template(raw_sub, indicators=indicators,
                                               custom_df=custom_df, recipes=recipes,
                                               admin_path=admin_path)
                del raw_sub
                wbuf = io.BytesIO(); TE.save_wb(wb, wbuf); wbuf.seek(0)
                sname = (sido_name_map or {}).get(sgg[:2], "")
                fname = _safe_name(f"{sgg}_{sname}_쇠퇴진단.xlsx")
                zf.writestr(fname, wbuf.getvalue())
                if out_dir:
                    try:
                        with open(os.path.join(out_dir, fname), "wb") as fh:
                            fh.write(wbuf.getvalue())
                    except Exception:
                        pass
                rows.append({"시군구코드": sgg, "시도": sname, "파일": fname,
                             "집계구수": stats.get("n_jgu", 0), "상태": "OK"})
            except Exception as e:   # 한 시군구 실패가 전체를 막지 않게 격리
                rows.append({"시군구코드": sgg, "시도": (sido_name_map or {}).get(sgg[:2], ""),
                             "파일": "", "집계구수": 0, "상태": f"실패: {e}"})
            finally:
                for p in sgg_parts.get(sgg, []):
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            if progress:
                progress(i + 1, total, sgg, "build")
        summary = pd.DataFrame(rows)
        try:
            csv_bytes = summary.to_csv(index=False).encode("utf-8-sig")
            zf.writestr("_요약.csv", csv_bytes)
            if out_dir:
                try:
                    with open(os.path.join(out_dir, "_요약.csv"), "wb") as fh:
                        fh.write(csv_bytes)
                except Exception:
                    pass
        except Exception:
            pass
    shutil.rmtree(tmp, ignore_errors=True)
    buf.seek(0)
    return buf.getvalue(), summary
