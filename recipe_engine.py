# -*- coding: utf-8 -*-
"""
recipe_engine.py — 계산식(레시피) 기반 사용자 지표 엔진
=====================================================
기존 검증 엔진(decline_engine)의 12지표는 건드리지 않는다. 여기서는 사용자가
'표 한 줄'로 정의한 지표를 원시자료에서 계산해 값 Series 로 만든 뒤,
decline_engine 의 표준화(standardize)를 그대로 재사용해 값/Z/T 로 만든다.

레시피(dict) 스키마:
  name      : 지표 이름(고유)
  sector    : 부문(인문사회/산업경제/물리환경)
  direction : '+'(값↑=쇠퇴↑) | '-'(값↑=쇠퇴↓)  → T점수 방향부호 ±10
  type      : 'ratio'(비율) | 'sum'(합계) | 'growth'(증감률)
  category  : 원시 분류 키(in_age/to_in/to_fa/cp_bem/ho_yr/ho_ar)
  num_codes : 분자 코드 리스트(code_no) 또는 None(=카테고리 전체)
  den_codes : 분모 코드 리스트 또는 None(=전체). ratio 에서만 사용
  scale     : 곱하는 배율(비율 보통 100, 합계 1)
  year      : 'pop' | 'biz' | 정수(연도).  ratio/sum=해당연도, growth=기준연도

지표를 '함수'가 아니라 '데이터'로 두므로 UI 에서 자유롭게 추가·삭제·수정 가능.
"""
from __future__ import annotations

import pandas as pd

import config as C
import decline_engine as E


# 원시 분류 → 기본 기준연도군(인구계열 pop / 산업계열 biz)
CATEGORY_LABELS = {
    "in_age": "성연령인구", "to_in": "총인구", "to_fa": "총사업체수",
    "cp_bem": "종사자수", "ho_yr": "건축연도주택", "ho_ar": "연건평주택",
}
_BIZ_CATEGORIES = {"to_fa", "cp_bem"}
TYPE_LABELS = {"비율": "ratio", "합계": "sum", "증감률": "growth"}
TYPE_LABELS_INV = {v: k for k, v in TYPE_LABELS.items()}

# 기본 12지표의 '수식 표시'(읽기전용 참고용) — 계산은 decline_engine 이 담당
BUILTIN_FORMULAS = {
    "인구변화율": "(기준연도 총인구 − 30년 최댓값) ÷ 최댓값 × 100   [방향 −]",
    "노년부양비": "Σ(in_age 014~021, 65세+) ÷ Σ(in_age 004~013, 15~64세) × 100   [방향 +]",
    "경제활동인구비율": "Σ(in_age 004~013) ÷ Σ(in_age 001~021) × 100   [방향 −]",
    "소멸위험지수": "Σ(in_age 065~068, 20~39세女) ÷ Σ(in_age 014~021, 65세+)   [방향 −]",
    "총사업체수증감률": "(기준연도 to_fa_010 − 최댓값) ÷ 최댓값 × 100   [방향 −]",
    "총종사자수증감률": "(기준연도 cp_bem 전체 − 최댓값) ÷ 최댓값 × 100   [방향 −]",
    "제조업증감률": "cp_bem 제조업(8차004/9·10차003) 증감률   [방향 −]",
    "고차산업증감률": "cp_bem 고차산업(010,011,013~016,+017) 증감률   [방향 −]",
    "도소매증감률": "cp_bem 도소매(007) 증감률   [방향 −]",
    "음식숙박증감률": "cp_bem 음식숙박(8차008/9·10차009) 증감률   [방향 −]",
    "노후건축물비율": "Σ(ho_yr 001~004, 2004이전) ÷ Σ(ho_yr 전체) × 100   [방향 +]",
    "소형주택비율": "Σ(ho_ar 001~003, 60㎡이하) ÷ Σ(ho_ar 전체) × 100   [방향 +]",
}


# 기본 12지표의 '복제 씨앗' — 기본지표를 편집 가능한 계산식 사본으로 복제할 때 쓴다.
#   exact=True  : recipe 로 계산해도 decline_engine 원본과 값이 정확히 일치
#   exact=False : 산업 8차(연도<=2005) 코드분기를 recipe 가 재현 못 해 '근사'(최신 차수 코드 고정)
#   direction/sector/scale 은 config.INDICATORS 와 일치. num/den 은 config 상수와 동일.
BUILTIN_SEEDS = {
    "인구변화율":     {"sector": "인문사회", "direction": "-", "type": "growth", "category": "to_in",
                    "num_codes": None,        "den_codes": None, "scale": 100, "exact": True},
    "노년부양비":     {"sector": "인문사회", "direction": "+", "type": "ratio", "category": "in_age",
                    "num_codes": list(range(14, 22)), "den_codes": list(range(4, 14)), "scale": 100, "exact": True},
    "경제활동인구비율": {"sector": "인문사회", "direction": "-", "type": "ratio", "category": "in_age",
                    "num_codes": list(range(4, 14)),  "den_codes": list(range(1, 22)), "scale": 100, "exact": True},
    "소멸위험지수":   {"sector": "인문사회", "direction": "-", "type": "ratio", "category": "in_age",
                    "num_codes": [65, 66, 67, 68], "den_codes": list(range(14, 22)), "scale": 1, "exact": True},
    "총사업체수증감률": {"sector": "산업경제", "direction": "-", "type": "growth", "category": "to_fa",
                    "num_codes": [10],        "den_codes": None, "scale": 100, "exact": True},
    "총종사자수증감률": {"sector": "산업경제", "direction": "-", "type": "growth", "category": "cp_bem",
                    "num_codes": None,        "den_codes": None, "scale": 100, "exact": True},
    "제조업증감률":   {"sector": "산업경제", "direction": "-", "type": "growth", "category": "cp_bem",
                    "num_codes": [3],         "den_codes": None, "scale": 100, "exact": False},
    "고차산업증감률": {"sector": "산업경제", "direction": "-", "type": "growth", "category": "cp_bem",
                    "num_codes": [10, 11, 13, 14, 15, 16, 17], "den_codes": None, "scale": 100, "exact": False},
    "도소매증감률":   {"sector": "산업경제", "direction": "-", "type": "growth", "category": "cp_bem",
                    "num_codes": [7],         "den_codes": None, "scale": 100, "exact": True},
    "음식숙박증감률": {"sector": "산업경제", "direction": "-", "type": "growth", "category": "cp_bem",
                    "num_codes": [9],         "den_codes": None, "scale": 100, "exact": False},
    "노후건축물비율": {"sector": "물리환경", "direction": "+", "type": "ratio", "category": "ho_yr",
                    "num_codes": [1, 2, 3, 4], "den_codes": None, "scale": 100, "exact": True},
    "소형주택비율":   {"sector": "물리환경", "direction": "+", "type": "ratio", "category": "ho_ar",
                    "num_codes": [1, 2, 3],   "den_codes": None, "scale": 100, "exact": True},
}


# 학습용 샘플 계산식 3종(비율/증감률/합계) — 기본은 '사용 해제'로 들어감(실제 진단 불변).
SAMPLE_RECIPES = [
    {"name": "20~39세여성비율", "sector": "인문사회", "direction": "-", "type": "ratio",
     "category": "in_age", "num_codes": [65, 66, 67, 68], "den_codes": list(range(1, 22)),
     "num_text": "65-68", "den_text": "1-21", "scale": 100.0, "year": "pop"},
    {"name": "건설업종사자증감률", "sector": "산업경제", "direction": "-", "type": "growth",
     "category": "cp_bem", "num_codes": [6], "den_codes": None,
     "num_text": "6", "den_text": "전체", "scale": 100.0, "year": "biz"},
    {"name": "총종사자수(합계샘플)", "sector": "산업경제", "direction": "-", "type": "sum",
     "category": "cp_bem", "num_codes": None, "den_codes": None,
     "num_text": "전체", "den_text": "전체", "scale": 1.0, "year": "biz"},
]
SAMPLE_NAMES = [r["name"] for r in SAMPLE_RECIPES]


def sample_recipes():
    """샘플 계산식 3종의 새 사본(list[dict])."""
    return [dict(r, num_codes=(list(r["num_codes"]) if r["num_codes"] else None),
                den_codes=(list(r["den_codes"]) if r["den_codes"] else None)) for r in SAMPLE_RECIPES]


def blank_recipe(existing_names=()) -> dict:
    """편집기 '한 줄 추가'용 빈 계산식(기본값). 이름은 중복 안 되게."""
    base, n, k = "새 계산식", "새 계산식", 2
    taken = set(existing_names)
    while n in taken:
        n = f"{base}{k}"
        k += 1
    return {"name": n, "sector": C.SECTORS[0], "direction": "+", "type": "ratio",
            "category": "in_age", "num_codes": None, "den_codes": None,
            "num_text": "", "den_text": "", "scale": 100.0, "year": "pop"}


def seed_from_builtin(name: str, existing_names=()) -> dict:
    """기본지표 이름 → 편집 가능한 계산식 레시피 사본(dict). 이름 충돌 시 '(사본2)' 식으로 회피."""
    seed = BUILTIN_SEEDS.get(name)
    if seed is None:
        raise ValueError(f"복제할 수 없는 기본지표: {name}")
    base = f"{name} (사본)"
    new_name, k = base, 2
    taken = set(existing_names)
    while new_name in taken:
        new_name = f"{name} (사본{k})"
        k += 1
    return {
        "name": new_name, "sector": seed["sector"], "direction": seed["direction"],
        "type": seed["type"], "category": seed["category"],
        "num_codes": list(seed["num_codes"]) if seed["num_codes"] else None,
        "den_codes": list(seed["den_codes"]) if seed["den_codes"] else None,
        "num_text": _fmt_codes(seed["num_codes"]), "den_text": _fmt_codes(seed["den_codes"]),
        "scale": float(seed["scale"]), "year": default_year_for(seed["category"]),
    }


def default_year_for(category: str) -> str:
    return "biz" if category in _BIZ_CATEGORIES else "pop"


def parse_codes(text):
    """'4-13, 65,66' → [4,…,13,65,66].  '전체'/빈칸 → None(=카테고리 전체)."""
    if text is None:
        return None
    s = str(text).strip()
    if s == "" or s.lower() in ("전체", "all", "*"):
        return None
    out = []
    for part in s.replace(" ", "").split(","):
        if not part:
            continue
        token = part.replace("~", "-")
        if "-" in token:
            a, b = token.split("-")[:2]
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(token))
    return sorted(set(out))


def _fmt_codes(codes):
    if not codes:
        return "전체"
    return ",".join(f"{c:03d}" for c in codes)


def formula_text(recipe: dict) -> str:
    cat = recipe.get("category", "")
    scale = recipe.get("scale", 1) or 1
    typ = recipe.get("type", "ratio")
    if typ == "ratio":
        return f"Σ({cat} {_fmt_codes(recipe.get('num_codes'))}) ÷ Σ({cat} {_fmt_codes(recipe.get('den_codes'))}) × {scale:g}"
    if typ == "sum":
        return f"Σ({cat} {_fmt_codes(recipe.get('num_codes'))}) × {scale:g}"
    return f"({cat} {_fmt_codes(recipe.get('num_codes'))}) 증감률 = (기준연도 − 최댓값) ÷ 최댓값 × 100"


def resolve_year(recipe: dict) -> int:
    y = recipe.get("year", "pop")
    if y == "pop":
        return int(C.YEAR_POP_LATEST)
    if y == "biz":
        return int(C.YEAR_BIZ_LATEST)
    return int(y)


def eval_recipe(recipe: dict, raw: dict, level: str) -> pd.Series:
    """레시피 → 지역별(집계구/행정동) 값 Series. 분모 0/데이터 없음은 0 처리(IFERROR)."""
    key = "집계구" if level == "jgu" else "행정동코드"
    df = raw.get(recipe.get("category"))
    if df is None or len(df) == 0:
        return pd.Series(dtype=float)
    g = E.aggregate_dong_year_code(df, key)          # key | 연도 | code_no | 값
    typ = recipe.get("type", "ratio")
    scale = float(recipe.get("scale", 1) or 1)

    if typ == "growth":
        codes = recipe.get("num_codes")
        d = g if codes is None else g[g["code_no"].isin(codes)]
        piv = d.groupby([key, "연도"])["값"].sum(min_count=1).unstack("연도")
        return E._growth_vs_max(piv, resolve_year(recipe))

    ref = resolve_year(recipe)
    gg = g[g["연도"] == ref]
    if gg.empty:
        return pd.Series(0.0, index=sorted(g[key].unique()))
    m = gg.groupby([key, "code_no"])["값"].sum(min_count=1).unstack("code_no")

    def s(codes):
        cols = list(m.columns) if codes is None else [c for c in codes if c in m.columns]
        return m[cols].sum(axis=1, min_count=1).fillna(0.0) if cols else pd.Series(0.0, index=m.index)

    num = s(recipe.get("num_codes"))
    if typ == "sum":
        return num * scale
    den = s(recipe.get("den_codes"))
    r = num / den * scale
    return r.where(den != 0, 0.0).fillna(0.0)


def build_recipe_scores(recipes, raw, level, index) -> pd.DataFrame:
    """레시피 리스트 → 값/Z/T DataFrame(지표별 컬럼). decline_engine.standardize 재사용."""
    out = {}
    for rc in recipes:
        name = rc.get("name")
        if not name:
            continue
        val = eval_recipe(rc, raw, level).reindex(index).astype(float)
        sign = 10 if rc.get("direction", "+") == "+" else -10
        z, t, _, _ = E.standardize(val, sign)
        out[(name, "값")] = val
        out[(name, "Z")] = z
        out[(name, "T")] = t
    if not out:
        return pd.DataFrame(index=index)
    sc = pd.DataFrame(out, index=index)
    sc.columns = pd.MultiIndex.from_tuples(sc.columns)
    return sc


def recipes_to_df(recipes) -> pd.DataFrame:
    """ss.recipes(list[dict]) → 편집용 DataFrame."""
    rows = []
    for rc in recipes:
        rows.append({
            "이름": rc.get("name", ""),
            "부문": rc.get("sector", C.SECTORS[0]),
            "방향": rc.get("direction", "+"),
            "유형": TYPE_LABELS_INV.get(rc.get("type", "ratio"), "비율"),
            "카테고리": rc.get("category", "in_age"),
            "분자코드": rc.get("num_text", _fmt_codes(rc.get("num_codes")) if rc.get("num_codes") else "전체"),
            "분모코드": rc.get("den_text", _fmt_codes(rc.get("den_codes")) if rc.get("den_codes") else "전체"),
            "스케일": float(rc.get("scale", 100) or 0),
        })
    cols = ["이름", "부문", "방향", "유형", "카테고리", "분자코드", "분모코드", "스케일"]
    return pd.DataFrame(rows, columns=cols)


def df_to_recipes(df) -> tuple[list, list]:
    """편집 DataFrame → (recipes list[dict], 오류메시지 list). 이름 없는 행은 건너뜀."""
    recipes, errors = [], []
    seen = set()
    for i, row in df.iterrows():
        name = str(row.get("이름", "") or "").strip()
        if not name:
            continue
        if name in seen:
            errors.append(f"'{name}' 이름이 중복됩니다.")
            continue
        seen.add(name)
        sector = str(row.get("부문", "")).strip()
        if sector not in C.SECTORS:
            errors.append(f"'{name}': 부문이 잘못됨({sector}).")
            continue
        typ = TYPE_LABELS.get(str(row.get("유형", "비율")).strip(), "ratio")
        category = str(row.get("카테고리", "in_age")).strip()
        if category not in CATEGORY_LABELS:
            errors.append(f"'{name}': 카테고리가 잘못됨({category}).")
            continue
        try:
            num_codes = parse_codes(row.get("분자코드"))
            den_codes = parse_codes(row.get("분모코드"))
        except Exception:
            errors.append(f"'{name}': 코드 형식 오류(예: 4-13, 65,66).")
            continue
        try:
            scale = float(row.get("스케일", 100) or 0)
        except Exception:
            scale = 100.0
        recipes.append({
            "name": name, "sector": sector,
            "direction": "+" if str(row.get("방향", "+")).strip() != "-" else "-",
            "type": typ, "category": category,
            "num_codes": num_codes, "den_codes": den_codes,
            "num_text": str(row.get("분자코드", "") or "").strip(),
            "den_text": str(row.get("분모코드", "") or "").strip(),
            "scale": scale, "year": default_year_for(category),
        })
    return recipes, errors
