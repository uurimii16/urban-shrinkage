# -*- coding: utf-8 -*-
"""
custom_indicators.py — 사용자 추가지표 로딩/표준화/가중합
=======================================================
추가지표 파일 형식:
  지표명 | 부문 | 방향 | 단위레벨 | 단위코드 | 값 | 내부가중치 | 설명(선택)
"""
from __future__ import annotations

import io

import pandas as pd

import config as C
import decline_engine as E


REQUIRED_COLUMNS = ["지표명", "부문", "방향", "단위레벨", "단위코드", "값", "내부가중치"]
SECTOR_ALIASES = {
    "인구": "인문사회",
    "인문": "인문사회",
    "인문사회": "인문사회",
    "산업": "산업경제",
    "산업경제": "산업경제",
    "물리": "물리환경",
    "물리환경": "물리환경",
}
LEVEL_ALIASES = {"행정동": "dong", "행정동코드": "dong", "dong": "dong",
                 "집계구": "jgu", "집계구코드": "jgu", "jgu": "jgu"}


def template_dataframe():
    return pd.DataFrame([
        ["부실건축물비율", "물리", "+", "행정동", "35011100", 12.3, 10, "값이 높을수록 쇠퇴"],
        ["부실건축물비율", "물리", "+", "행정동", "35011200", 8.7, 10, ""],
        ["상권활력지수", "산업", "-", "행정동", "35011100", 72.4, 10, "값이 낮을수록 쇠퇴"],
    ], columns=REQUIRED_COLUMNS + ["설명"])


def template_xlsx_bytes():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        template_dataframe().to_excel(writer, sheet_name="추가지표", index=False)
    buf.seek(0)
    return buf.getvalue()


def read_uploaded(file):
    if file is None:
        return pd.DataFrame(columns=REQUIRED_COLUMNS)
    name = getattr(file, "name", "").lower()
    data = file.getvalue()
    if name.endswith(".csv"):
        for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
            try:
                return pd.read_csv(io.BytesIO(data), encoding=enc, dtype=str)
            except Exception:
                continue
        raise ValueError("추가지표 CSV 인코딩을 읽지 못했습니다.")
    return pd.read_excel(io.BytesIO(data), sheet_name="추가지표", dtype=str)


def normalize(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=REQUIRED_COLUMNS + ["설명"])
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError("추가지표 파일 필수 열 누락: " + ", ".join(missing))
    d = df.copy()
    for c in REQUIRED_COLUMNS:
        d[c] = d[c].astype(str).str.strip()
    d = d[d["지표명"] != ""].copy()
    d["부문"] = d["부문"].map(lambda x: SECTOR_ALIASES.get(x, x))
    bad_sector = sorted(set(d.loc[~d["부문"].isin(C.SECTORS), "부문"]))
    if bad_sector:
        raise ValueError("추가지표 부문 값 오류: " + ", ".join(bad_sector))
    d["방향"] = d["방향"].replace({"높을수록쇠퇴": "+", "낮을수록쇠퇴": "-", "증가": "+", "감소": "-"})
    bad_dir = sorted(set(d.loc[~d["방향"].isin(["+", "-"]), "방향"]))
    if bad_dir:
        raise ValueError("추가지표 방향은 + 또는 - 여야 합니다: " + ", ".join(bad_dir))
    d["단위레벨"] = d["단위레벨"].map(lambda x: LEVEL_ALIASES.get(x, x))
    bad_level = sorted(set(d.loc[~d["단위레벨"].isin(["dong", "jgu"]), "단위레벨"]))
    if bad_level:
        raise ValueError("추가지표 단위레벨은 행정동 또는 집계구여야 합니다: " + ", ".join(bad_level))
    d["값"] = pd.to_numeric(d["값"], errors="coerce")
    d["내부가중치"] = pd.to_numeric(d["내부가중치"], errors="coerce")
    if d["값"].isna().any():
        raise ValueError("추가지표 값 열에 숫자가 아닌 값이 있습니다.")
    if d["내부가중치"].isna().any():
        raise ValueError("추가지표 내부가중치 열에 숫자가 아닌 값이 있습니다.")

    # 지표별 메타데이터는 모든 행에서 동일해야 한다.
    meta_cols = ["부문", "방향", "단위레벨", "내부가중치"]
    for ind, sub in d.groupby("지표명"):
        for col in meta_cols:
            vals = sub[col].dropna().unique()
            if len(vals) > 1:
                raise ValueError(f"추가지표 '{ind}'의 {col} 값이 행마다 다릅니다.")
    return d


def metadata(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["지표", "부문", "방향", "내부가중치"])
    rows = []
    for ind, sub in df.groupby("지표명", sort=False):
        first = sub.iloc[0]
        rows.append({
            "지표": ind,
            "부문": first["부문"],
            "방향": first["방향"],
            "내부가중치": float(first["내부가중치"]),
        })
    return pd.DataFrame(rows)


def build_scores(df, index, level):
    """추가지표 long → scores DataFrame(지표 × 값/Z/T)."""
    if df is None or df.empty:
        return pd.DataFrame(index=index)
    d = df[df["단위레벨"] == level].copy()
    if d.empty:
        return pd.DataFrame(index=index)
    out = {}
    for ind, sub in d.groupby("지표명", sort=False):
        ser = (sub.groupby("단위코드")["값"].mean()
                 .reindex(index)
                 .astype(float))
        sign = 10 if sub.iloc[0]["방향"] == "+" else -10
        z, t, _, _ = E.standardize(ser, sign)
        out[(ind, "값")] = ser
        out[(ind, "Z")] = z
        out[(ind, "T")] = t
    scores = pd.DataFrame(out, index=index)
    if len(scores.columns):
        scores.columns = pd.MultiIndex.from_tuples(scores.columns)
    return scores


def combine_scores(base_scores, custom_scores):
    if custom_scores is None or custom_scores.empty:
        return base_scores
    return pd.concat([base_scores, custom_scores], axis=1)


def composite(scores, indicator_ids, sector_of, weight):
    res = {}
    for sec in C.SECTORS:
        s = pd.Series(0.0, index=scores.index)
        for ind in indicator_ids:
            if sector_of.get(ind) == sec and (ind, "T") in scores.columns:
                s = s + scores[(ind, "T")] * float(weight.get(ind, 0.0))
        res[sec] = s
    df = pd.DataFrame(res)
    df["종합"] = df[C.SECTORS].sum(axis=1)
    return df
