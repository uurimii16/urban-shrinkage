# -*- coding: utf-8 -*-
"""
batch_run.py — 전국 시군구 배치 빌드 (터미널 백그라운드 일괄처리)
==================================================================
UI를 열지 않고 명령 한 줄로: 원시 SGIS 데이터 폴더/zip → 시군구별 진단 xlsx를
출력 폴더에 착착 저장(+ 전체 zip 1개). 229개처럼 대량·오래 걸리는 작업을
돌려놓고 자리를 떠도 되게 한다.

사용법(Windows PowerShell 예):
  python batch_run.py "D:\\원시데이터폴더_또는_zip" "D:\\쇠퇴진단_전국출력"
  python batch_run.py 입력 출력 --year-pop 2024 --year-biz 2023
  python batch_run.py 입력 출력 --final-only         # 최종 4시트만(값)
  python batch_run.py 입력 출력 --values             # 전체 시트지만 함수 없이 값으로

옵션:
  --year-pop N   인구·주택 기준연도(기본 2024)
  --year-biz N   산업 기준연도(기본 2023)
  --years a,b,c  사용할 연도만(기본: 데이터 전체)
  --final-only   최종 4시트만(법적·복합 × 행정동·집계구)
  --values       함수(수식) 없이 값으로만(기본은 역산 함수 포함)
  --decimals N   소수점 자릿수(기본 2)
  --sigungu a,b  특정 시군구코드만(기본: 전체)
"""
import argparse
import os
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import loader as L
import batch_build as BB
import dong_names


def main():
    ap = argparse.ArgumentParser(description="전국 시군구 배치 빌드(백그라운드 일괄처리)")
    ap.add_argument("input", help="원시 SGIS 데이터 폴더 또는 zip 경로")
    ap.add_argument("output", help="결과를 저장할 폴더 경로")
    ap.add_argument("--year-pop", type=int, default=2024)
    ap.add_argument("--year-biz", type=int, default=2023)
    ap.add_argument("--years", default="", help="쉼표구분 연도(예: 2020,2024). 비우면 전체")
    ap.add_argument("--final-only", action="store_true")
    ap.add_argument("--values", action="store_true", help="함수 없이 값으로")
    ap.add_argument("--decimals", type=int, default=2)
    ap.add_argument("--sigungu", default="", help="쉼표구분 시군구코드. 비우면 전체")
    args = ap.parse_args()

    print(f"[1/3] 원시데이터 로드: {args.input}")
    raw = L.load_raw_from_folders(args.input, verbose=True)
    sgg = BB.list_sigungu(raw)
    print(f"      인식된 시군구: {len(sgg)}곳 — {', '.join(sgg[:30])}{' …' if len(sgg) > 30 else ''}")

    sel_years = [int(y) for y in args.years.split(",") if y.strip()] or None
    only = [c.strip() for c in args.sigungu.split(",") if c.strip()] or None
    fmt = "최종4시트(값)" if args.final_only else ("전체(값)" if args.values else "전체(함수포함)")
    print(f"[2/3] 배치 빌드 시작 · 형식={fmt} · 기준연도 인구{args.year_pop}/산업{args.year_biz} "
          f"· 소수점{args.decimals}")

    t0 = time.time()

    def _cb(done, total, code):
        el = time.time() - t0
        eta = (el / done) * (total - done) if done else 0
        print(f"  {done}/{total}  {code}  ({el:5.0f}s 경과, 남은 예상 {eta:5.0f}s)", flush=True)

    zbytes, summary = BB.build_batch_zip(
        raw, sigungu=only, name_map=dong_names.default_name_map(),
        sido_name_map=dong_names.default_sido_map() or None, decimals=args.decimals,
        final_only=args.final_only, formula_mode=(not args.values),
        selected_years=sel_years, year_pop=args.year_pop, year_biz=args.year_biz,
        out_dir=args.output, progress=_cb)

    zip_path = os.path.join(args.output, "쇠퇴진단_전국배치.zip")
    with open(zip_path, "wb") as f:
        f.write(zbytes)

    ok = int((summary["상태"] == "OK").sum()) if len(summary) else 0
    print(f"[3/3] 완료 · 성공 {ok}/{len(summary)}곳 · {time.time() - t0:.0f}s")
    print(f"      저장 폴더: {args.output}")
    print(f"      전체 zip : {zip_path}")
    fails = summary[summary["상태"] != "OK"] if len(summary) else []
    if len(fails):
        print("      실패 시군구:")
        for _, r in fails.iterrows():
            print(f"        - {r['시군구코드']}: {r['상태']}")


if __name__ == "__main__":
    main()
