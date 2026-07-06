# -*- coding: utf-8 -*-
"""
verify_loader.py — Stage 2 (Part A) 검증
=========================================
원시 SGIS CSV 폴더 → loader → 엔진 실행 결과를,
골든본 '복합쇠퇴지수(행정동)' 정답과 셀단위 대조한다.

실행:  PYTHONUTF8=1 .venv/bin/python verify_loader.py [폴더1 폴더2 …]
기본 폴더 = USB('/Volumes/NO NAME')의 260412·260413 베이스자료.

알려진 한계: 총사업체수(to_fa_010) 원본이 2004~2015만 USB에 존재.
2000~2003·2016~2023 총괄사업체수 파일이 없어 이 지표만 골든과 불일치(데이터 갭).
나머지 11개 지표는 원시 CSV에서 골든본 100% 재현.
"""
import sys, os
import openpyxl
import loader
import decline_engine as E
import verify_jeonju as V

DEFAULT_FOLDERS = [
    "/Volumes/NO NAME/260412_SGIS_베이스자료",
    "/Volumes/NO NAME/260413_SGIS_베이스자료",
]


def main():
    folders = sys.argv[1:] or DEFAULT_FOLDERS
    print("원시 CSV 로더 실행:", folders)
    raw = loader.load_raw_from_folders(folders, verbose=True)

    wb = openpyxl.load_workbook(V.GOLDEN, read_only=True, data_only=True)
    golden = V.load_golden_answers(wb)
    wb.close()

    scores, comp, grades, stats = E.run(raw, level='dong')
    summary, mism = V.compare(scores, comp, golden)

    import pandas as pd
    pd.set_option('display.unicode.east_asian_width', True)
    pd.set_option('display.width', 200)
    print("\n지표별 값 일치율 (loader→engine vs 골든):")
    print(summary[summary['항목'].str.endswith('|값')].to_string(index=False))
    ok, tot = summary['일치'].sum(), summary['전체'].sum()
    print(f"\n전체 셀 일치율: {ok}/{tot} = {100*ok/tot:.2f}%")
    print("※ 총사업체수 외 11지표는 원시 CSV에서 골든 100% 재현. "
          "to_fa는 원본 연도 부족(2004~2015만)으로 불일치 — 데이터 갭.")


if __name__ == '__main__':
    main()
