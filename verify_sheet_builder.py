# -*- coding: utf-8 -*-
"""
verify_sheet_builder.py — 통합 DATA 시트 빌더 최소 회귀 검증
===========================================================
골든 파일 없이도 synthetic raw 데이터로 빌더/엔진/엑셀 생성 경로가
깨지지 않는지 확인한다.
"""
import io

import pandas as pd

import decline_engine as E
import export
import legal_engine as LG


def _sample_raw():
    cols = ["연도", "집계구", "CODE", "값", "행정동코드"]
    code_map = {
        "to_in": ["to_in_001"],
        "to_fa": ["to_fa_010"],
        "in_age": [f"in_age_{i:03d}" for i in list(range(1, 22)) + list(range(65, 69))],
        "cp_bem": [f"cp_bem_{i:03d}" for i in range(1, 18)],
        "ho_yr": [f"ho_yr_{i:03d}" for i in range(1, 11)],
        "ho_ar": [f"ho_ar_{i:03d}" for i in range(1, 10)],
    }
    units = [("35011100010001", "35011100", 10), ("35011200010001", "35011200", 20)]
    raw = {}
    for key, codes in code_map.items():
        rows = []
        for yr in [2020, 2021, 2022, 2023, 2024]:
            for gu, dong, base in units:
                for code in codes:
                    rows.append((yr, gu, code, float(base + (yr - 2020)), dong))
        raw[key] = pd.DataFrame(rows, columns=cols)
    return raw


def main():
    raw = _sample_raw()
    dong = E.run(raw, level="dong")
    jgu = E.run(raw, level="jgu")
    legal_dong = LG.run_legal(raw, level="dong")
    legal_jgu = LG.run_legal(raw, level="jgu")
    wb = export.build_integrated_workbook(
        raw,
        selected_years=[2020, 2021, 2022, 2023, 2024],
        dong_res=dong[:3],
        jgu_res=jgu[:3],
        legal_dong=legal_dong,
        legal_jgu=legal_jgu,
    )
    required = {
        "생성요약",
        "법적DATA_1.인구총괄(수정)",
        "전주시 법적쇠퇴진단(행정동)",
        "전주시 복합쇠퇴지수(행정동)",
    }
    missing = required - set(wb.sheetnames)
    if missing:
        raise AssertionError(f"필수 시트 누락: {sorted(missing)}")
    buf = io.BytesIO()
    wb.save(buf)
    if buf.tell() <= 0:
        raise AssertionError("워크북 저장 결과가 비어 있습니다.")
    print(f"OK: {len(wb.sheetnames)} sheets, {buf.tell():,} bytes")


if __name__ == "__main__":
    main()
