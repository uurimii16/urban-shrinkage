# -*- coding: utf-8 -*-
"""추가지표 + 수식 엑셀 경로 최소 검증."""
import pandas as pd

import config as C
import custom_indicators as CI
import decline_engine as E
import export
import legal_engine as LG


def _raw():
    cols = ["연도", "집계구", "CODE", "값", "행정동코드"]
    code_map = {
        "to_in": ["to_in_001"],
        "to_fa": ["to_fa_010"],
        "in_age": [f"in_age_{i:03d}" for i in list(range(1, 22)) + list(range(65, 69))],
        "cp_bem": [f"cp_bem_{i:03d}" for i in range(1, 18)],
        "ho_yr": [f"ho_yr_{i:03d}" for i in range(1, 11)],
        "ho_ar": [f"ho_ar_{i:03d}" for i in range(1, 10)],
    }
    out = {}
    for key, codes in code_map.items():
        rows = []
        for yr in [2020, 2021, 2022, 2023, 2024]:
            for gu, dong, base in [("35011100010001", "35011100", 10), ("35011200010001", "35011200", 20)]:
                for code in codes:
                    rows.append((yr, gu, code, float(base + yr - 2020), dong))
        out[key] = pd.DataFrame(rows, columns=cols)
    return out


def main():
    raw = _raw()
    custom = CI.normalize(pd.DataFrame([
        ["부실건축물비율", "물리", "+", "행정동", "35011100", 12.3, 10],
        ["부실건축물비율", "물리", "+", "행정동", "35011200", 8.7, 10],
    ], columns=CI.REQUIRED_COLUMNS))

    for ind in C.IND_IDS:
        C.WEIGHT[ind] = 1 / 12
    base = E.run(raw, level="dong")
    custom_scores = CI.build_scores(custom, base[0].index, "dong")
    scores = CI.combine_scores(base[0], custom_scores)
    indicator_ids = list(C.IND_IDS) + ["부실건축물비율"]
    sector_of = dict(C.SECTOR_OF)
    sector_of["부실건축물비율"] = "물리환경"
    weight = dict(C.WEIGHT)
    weight["부실건축물비율"] = 0.04
    sign_map = dict(C.SIGN)
    sign_map["부실건축물비율"] = 10
    comp = CI.composite(scores, indicator_ids, sector_of, weight)
    grades = E.assign_grades(comp["종합"])

    wb = export.build_integrated_workbook(
        raw,
        dong_res=(scores, comp, grades),
        legal_dong=LG.run_legal(raw),
        indicator_ids=indicator_ids,
        sector_of=sector_of,
        weight=weight,
        sign_map=sign_map,
        formula_mode=True,
    )
    final = wb["전주시 복합쇠퇴지수(행정동)"]
    if not str(final["C3"].value).startswith("="):
        raise AssertionError("수식 모드 최종표에 수식이 생성되지 않았습니다.")
    print("OK: custom indicators + formula workbook")


if __name__ == "__main__":
    main()
