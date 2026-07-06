# -*- coding: utf-8 -*-
"""
loader.py — 원시 SGIS CSV 폴더 → 정제 long dict (Part A)
=========================================================
흩어진 cp949 CSV(헤더없음, 4열: 연도,집계구코드,항목코드,값)를 읽어
엔진 입력(golden_io 와 동일 스키마: 연도|집계구|CODE|값|행정동코드) 6분류로 정제.

확정 사실(원시 파일 실측):
  · 산업 종사자수는 전 차수 cp_bem 접두어(8/9/10차 코드번호만 다름 → config.industry_codes 처리)
  · 총사업체수 = 독립코드 to_fa_010("총괄사업체수" 파일). 산업별 사업체수(cp_bnu)는 미사용.
  · 총인구 = to_in_001,  성연령 = in_age_*,  주택 = ho_yr_* / ho_ar_*
  · 집계구코드 = 14자리. (행정구역)35 롤업(2자리)·꼬리공백(15자)은 strip 후 14자리만 채택.
  · 손상파일 35011_2023년_성연령별인구.csv 제외(엔진 미사용 연도이나 오염 방지).
"""
import io
import os
import glob
import zipfile
import numpy as np
import pandas as pd
import config as C
from golden_io import to_num

# CODE → 분류 버킷 라우팅 규칙 (정확매칭 또는 접두어)
_EXACT = {'to_in_001': 'to_in', 'to_fa_010': 'to_fa'}
_PREFIX = [('in_age_', 'in_age'), ('cp_bem_', 'cp_bem'),
           ('ho_yr_', 'ho_yr'), ('ho_ar_', 'ho_ar')]
BUCKETS = ['to_in', 'in_age', 'to_fa', 'cp_bem', 'ho_yr', 'ho_ar']


def _bucket_of(code: str):
    if code in _EXACT:
        return _EXACT[code]
    for pre, b in _PREFIX:
        if code.startswith(pre):
            return b
    return None


def _is_corrupt(path: str) -> bool:
    n = os.path.basename(path)
    return ('2023' in n) and ('성연령' in n)


def find_csvs(folders):
    """폴더(들) 재귀 순회 → SGIS 원시 파일(csv/txt) 경로 리스트(손상파일 제외)."""
    if isinstance(folders, (str, os.PathLike)):
        folders = [folders]
    out = []
    for fol in folders:
        for ext in ('*.csv', '*.CSV', '*.txt', '*.TXT'):
            out += glob.glob(os.path.join(fol, '**', ext), recursive=True)
    out = sorted(set(out))
    return [p for p in out if not _is_corrupt(p)]


def _read_one(path):
    """cp949/euc-kr/utf-8 4열 CSV/TXT → DataFrame(연도,집계구,CODE,값)."""
    for enc in ('cp949', 'euc-kr', 'utf-8-sig', 'utf-8'):
        for sep in (',', '^', '\t'):
            try:
                return pd.read_csv(path, header=None, encoding=enc, dtype=str,
                                   names=['연도', '집계구', 'CODE', '값'],
                                   usecols=[0, 1, 2, 3], sep=sep,
                                   skip_blank_lines=True)
            except Exception:
                continue
    return pd.DataFrame(columns=['연도', '집계구', 'CODE', '값'])


def _read_csv_bytes(data: bytes, column_map=None):
    """업로드 CSV/TXT bytes → DataFrame(연도,집계구,CODE,값).
    column_map이 없으면 SGIS 기본 4열(0,1,2,3)을 사용한다."""
    column_map = column_map or {'연도': 0, '집계구': 1, 'CODE': 2, '값': 3}
    usecols = [column_map[k] for k in ['연도', '집계구', 'CODE', '값']]
    names = ['연도', '집계구', 'CODE', '값']
    for enc in ('cp949', 'euc-kr', 'utf-8-sig', 'utf-8'):
        for sep in (',', '^', '\t'):
            try:
                df = pd.read_csv(io.BytesIO(data), header=None, encoding=enc, dtype=str,
                                 usecols=usecols, sep=sep, skip_blank_lines=True)
                df = df.rename(columns=dict(zip(usecols, names)))[names]
                return df
            except Exception:
                continue
    return pd.DataFrame(columns=names)


def summarize_rows(df, source_name):
    if df is None or df.empty:
        return []
    d = df.copy()
    d['연도'] = pd.to_numeric(d['연도'], errors='coerce')
    d['집계구'] = d['집계구'].astype(str).str.strip()
    d['CODE'] = d['CODE'].astype(str).str.strip()
    d = d[d['집계구'].str.len() >= 5]
    d['시군구코드'] = d['집계구'].str[:5]
    rows = []
    for sigungu, sub in d.groupby('시군구코드'):
        years = sorted(sub['연도'].dropna().astype(int).unique().tolist())
        rows.append({
            '파일': source_name,
            '시군구코드': sigungu,
            '행수': len(sub),
            '연도': f"{min(years)}~{max(years)}" if years else '',
            '연도목록': ', '.join(map(str, years[:20])),
            '항목수': sub['CODE'].nunique(),
            '집계구수': sub['집계구'].nunique(),
        })
    return rows


def summarize_folders(folders):
    rows = []
    for path in find_csvs(folders):
        rows.extend(summarize_rows(_read_one(path), os.path.basename(path)))
    return pd.DataFrame(rows)


def _expand_uploads(files):
    """업로드 목록 → [(파일명, bytes)]. **zip은 자동으로 풀어** 내부 csv/txt를 펼친다.
    (SGIS 자료제공 다운로드가 zip이라 압축 안 풀고 그대로 올려도 되게)."""
    out = []
    for f in files or []:
        name = getattr(f, 'name', '업로드파일')
        data = f.getvalue() if hasattr(f, 'getvalue') else bytes(f)
        if name.lower().endswith('.zip'):
            try:
                zf = zipfile.ZipFile(io.BytesIO(data))
            except Exception:
                continue
            for zi in zf.namelist():
                if zi.endswith('/'):
                    continue
                low = zi.lower()
                if low.endswith(('.csv', '.txt')):
                    inner = os.path.basename(zi) or zi
                    try:
                        out.append((inner, zf.read(zi)))
                    except Exception:
                        continue
        else:
            out.append((name, data))
    return out


def summarize_uploaded_files(files, column_map=None):
    rows = []
    for name, data in _expand_uploads(files):
        if _is_corrupt(name):
            continue
        rows.extend(summarize_rows(_read_csv_bytes(data, column_map), name))
    return pd.DataFrame(rows)


def _build_raw_from_rows(frames, mapping=None):
    """정제 전 row DataFrame 목록 → 엔진 raw dict."""
    if not frames:
        raise ValueError('원시 CSV/TXT를 읽지 못했습니다. 폴더 안에 .csv/.txt 파일이 있는지, SGIS 4열 형식인지 확인하세요.')
    allrows = pd.concat(frames, ignore_index=True)

    allrows['집계구'] = allrows['집계구'].astype(str).str.strip()
    allrows['CODE'] = allrows['CODE'].astype(str).str.strip()
    # 지원 공간단위: 집계구(14자리) 또는 행정동(8자리).
    #   2·5자리(시도·시군구 롤업 행)는 버린다. 단위는 '분류(bucket)별로' 정한다 —
    #   성연령은 행정동(8)만, 산업은 집계구(14)만 받은 경우처럼 폴더마다 단위가 다를 수 있어
    #   전역 단일 단위로 자르면 한쪽이 통째로 사라진다(예: 성연령 누락). 아래 분류별 처리로 해결.
    allrows['_len'] = allrows['집계구'].str.len()
    allrows = allrows[allrows['_len'].isin([8, 14])]
    allrows['bucket'] = allrows['CODE'].map(_bucket_of)
    allrows = allrows[allrows['bucket'].notna()]
    if allrows.empty:
        raise ValueError('읽은 파일에서 지원 공간단위(집계구 14자리 또는 행정동 8자리)의 '
                         'SGIS 항목코드를 찾지 못했습니다. 원시 4열 형식(연도, 집계구, CODE, 값)인지 확인하세요.')

    allrows['연도'] = pd.to_numeric(allrows['연도'], errors='coerce')
    allrows = allrows[allrows['연도'].notna()]
    allrows['연도'] = allrows['연도'].astype(int)
    allrows['값'] = allrows['값'].map(to_num)

    raw = {}
    for b in BUCKETS:
        sub = allrows[allrows['bucket'] == b].copy()
        # 이 분류에 집계구(14)가 있으면 14 우선(더 세밀하고 행정동으로 정확히 롤업).
        # 14가 섞여 있으면 8자리(행정동 소계행)는 버려 이중집계를 막는다. 14가 전혀 없으면 8 사용.
        unit = 14 if (len(sub) and (sub['_len'] == 14).any()) else 8
        if len(sub):
            sub = sub[sub['_len'] == unit]
        sub = (sub[['연도', '집계구', 'CODE', '값']]
               .drop_duplicates(subset=['연도', '집계구', 'CODE'], keep='first')
               .reset_index(drop=True))
        if unit == 8:
            sub['행정동코드'] = sub['집계구']                 # 이미 행정동 단위
        elif mapping:
            sub['행정동코드'] = sub['집계구'].map(lambda g: mapping.get(g, g[:8]))
        else:
            sub['행정동코드'] = sub['집계구'].str[:8]
        raw[b] = sub
    return raw


def _zip_members(folders):
    """폴더 안(또는 경로 자체)의 .zip을 열어 내부 csv/txt를 (파일명, bytes)로 반환.
    → 압축 안 풀고 zip 그대로 폴더에 둬도(또는 zip 경로를 직접 줘도) 읽힘."""
    if isinstance(folders, (str, os.PathLike)):
        folders = [folders]
    zips = []
    for fol in folders:
        fol = str(fol)
        if fol.lower().endswith('.zip') and os.path.isfile(fol):
            zips.append(fol)
        elif os.path.isdir(fol):
            for ext in ('*.zip', '*.ZIP'):
                zips += glob.glob(os.path.join(fol, '**', ext), recursive=True)
    out = []
    for zp in sorted(set(zips)):
        try:
            with open(zp, 'rb') as fh:
                zf = zipfile.ZipFile(io.BytesIO(fh.read()))
        except Exception:
            continue
        for zi in zf.namelist():
            if zi.endswith('/') or not zi.lower().endswith(('.csv', '.txt')):
                continue
            inner = os.path.basename(zi) or zi
            if _is_corrupt(inner):
                continue
            try:
                out.append((inner, zf.read(zi)))
            except Exception:
                continue
    return out


def load_raw_from_folders(folders, mapping=None, verbose=False):
    """원시 CSV/TXT 폴더(들) → dict[str, DataFrame] (엔진 입력 스키마).
    폴더 안(또는 경로 자체)의 **zip도 자동 해제**. mapping: {집계구14: 행정동8} — 없으면 집계구[:8]."""
    frames = []
    for p in find_csvs(folders):
        df = _read_one(p)
        if len(df):
            frames.append(df)
    for name, data in _zip_members(folders):     # zip 내부 csv/txt
        df = _read_csv_bytes(data)
        if len(df):
            frames.append(df)
    if not frames:
        raise ValueError('입력에서 .csv/.txt(또는 zip 내부) 원시 파일을 찾지 못했습니다. 경로를 확인하세요.')
    raw = _build_raw_from_rows(frames, mapping=mapping)
    if verbose:
        for b, sub in raw.items():
            yrs = sorted(sub["연도"].unique())[:3] if len(sub) else []
            print(f'  {b:7s}: {len(sub):>7,}행  연도 {yrs}…')
    return raw


def load_raw_from_uploaded_files(files, mapping=None, column_map=None):
    """Streamlit 업로드 CSV 목록 → 엔진 raw dict.
    column_map: {'연도':0, '집계구':1, 'CODE':2, '값':3} 형태의 0-based 열 번호."""
    frames = []
    for name, data in _expand_uploads(files):
        if name and _is_corrupt(name):
            continue
        df = _read_csv_bytes(data, column_map=column_map)
        if len(df):
            frames.append(df)
    return _build_raw_from_rows(frames, mapping=mapping)


def coverage_warnings(raw):
    """원시 커버리지 점검 → 경고 메시지 리스트. 기준연도 부재 등."""
    warns = []
    def years(b):
        return set(raw[b]['연도'].unique()) if b in raw and len(raw[b]) else set()
    if C.YEAR_POP_LATEST not in years('to_in'):
        warns.append(f"총인구 {C.YEAR_POP_LATEST}년 데이터 없음 → 인구변화율 부정확.")
    if C.YEAR_POP_LATEST not in years('in_age'):
        warns.append(f"성연령 {C.YEAR_POP_LATEST}년 데이터 없음 → 인문사회 지표 부정확.")
    if C.YEAR_BIZ_LATEST not in years('to_fa'):
        yy = sorted(years('to_fa'))
        rng = f"{min(yy)}~{max(yy)}" if yy else "없음"
        warns.append(f"총사업체수 기준연도 {C.YEAR_BIZ_LATEST} 없음(보유 {rng}) "
                     f"→ 총사업체수 증감률 부정확(원본 총괄사업체수 파일 부족).")
    if C.YEAR_BIZ_LATEST not in years('cp_bem'):
        warns.append(f"종사자수 기준연도 {C.YEAR_BIZ_LATEST} 없음 → 산업 증감률 부정확.")
    if C.YEAR_POP_LATEST not in years('ho_yr'):
        warns.append(f"건축연도주택 {C.YEAR_POP_LATEST}년 없음 → 노후건축물비율 부정확.")
    # ho_ar(연면적)은 소형주택비율(기본지표에서 제거됨) 전용 → 없어도 경고하지 않음.
    #   '복제'로 소형주택비율을 되살린 경우에만 필요하며, 그땐 code_audit/지표별 경고로 안내됨.
    return warns
