# -*- coding: utf-8 -*-
"""
batch_watch.py — 승인 자동 감시 → 다운로드 → 전국 배치 빌드 (무인 워처)
========================================================================
신청까지 해둔 뒤 이 스크립트를 돌려두면, SGIS '다운로드 목록'을 주기적으로
확인하다가 **승인 자료가 뜨는 순간 자동으로 전부 다운로드**하고 시군구별
디폴트 엑셀을 만들어 저장한 뒤 종료한다. 자리를 떠나도 됨(컴퓨터만 켜두면).

사용법(Windows PowerShell 예):
  1) 쿠키를 파일로 저장(메모장): 이 폴더에 cookie.txt (Copy as cURL 통째로 붙여넣기)
  2) python batch_watch.py "D:\\쇠퇴진단_전국출력"
     - 승인건이 하나라도 뜨고 '더 안 늘어나면' 자동 진행
  옵션:
     --cookie-file cookie.txt   쿠키 파일(기본 cookie.txt)
     --cookie "..."             쿠키를 직접 전달(파일 대신)
     --min-items N              최소 N건 승인되면 즉시 진행(기본: 안 정해두고 '안정되면')
     --interval 120             확인 주기(초, 기본 120)
     --stable 3                 승인 건수가 N회 연속 그대로면 '다 됐다'로 보고 진행(기본 3)
     --max-wait 7200            최대 대기(초, 기본 2시간). 넘으면 그때까지 것으로 진행
     --final-only / --values / --decimals N   (batch_run과 동일)

주의: SGIS 로그인 세션(쿠키)은 시간이 지나면 만료돼요. 오래 걸리면 중간에
      cookie.txt를 최신 쿠키로 덮어써 주세요(스크립트가 매 회차 다시 읽음).
"""
import argparse
import os
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import sgis_request as SR
import loader as L
import batch_build as BB
import dong_names


class _NB:
    def __init__(self, name, data):
        self.name, self._d = name, data

    def getvalue(self):
        return self._d


def _read_cookie(args):
    if args.cookie:
        return SR.extract_cookie(args.cookie)
    try:
        with open(args.cookie_file, "r", encoding="utf-8") as f:
            return SR.extract_cookie(f.read())
    except Exception:
        return ""


def main():
    ap = argparse.ArgumentParser(description="승인 감시 → 다운로드 → 전국 배치 빌드(무인)")
    ap.add_argument("output", help="결과 저장 폴더")
    ap.add_argument("--cookie-file", default="cookie.txt")
    ap.add_argument("--cookie", default="")
    ap.add_argument("--min-items", type=int, default=0)
    ap.add_argument("--interval", type=int, default=120)
    ap.add_argument("--stable", type=int, default=3)
    ap.add_argument("--max-wait", type=int, default=7200)
    ap.add_argument("--year-pop", type=int, default=0, help="0이면 데이터에서 자동")
    ap.add_argument("--year-biz", type=int, default=0)
    ap.add_argument("--final-only", action="store_true")
    ap.add_argument("--values", action="store_true")
    ap.add_argument("--decimals", type=int, default=2)
    args = ap.parse_args()

    print(f"[감시 시작] 저장={args.output} · 주기 {args.interval}s · "
          f"조건={'≥'+str(args.min_items)+'건' if args.min_items else f'{args.stable}회 연속 동일'} "
          f"· 최대대기 {args.max_wait}s")
    t0 = time.time()
    prev_count, stable_hits = -1, 0
    items = []

    while True:
        ck = _read_cookie(args)
        if not ck:
            print("  ! 쿠키를 못 읽음(cookie.txt 확인). 30초 후 재시도")
            time.sleep(30)
            continue
        try:
            items = SR.fetch_download_list(ck)
        except Exception as e:
            print(f"  ! 목록 조회 실패({e}). 쿠키 만료면 cookie.txt 갱신. {args.interval}s 후 재시도")
            time.sleep(args.interval)
            continue

        n = len(items)
        el = int(time.time() - t0)
        print(f"  [{el:5d}s] 승인 다운로드 가능: {n}건", flush=True)

        ready = False
        if args.min_items and n >= args.min_items:
            ready = True
            print(f"  → 최소 {args.min_items}건 도달 → 진행")
        elif n > 0:
            stable_hits = stable_hits + 1 if n == prev_count else 0
            if stable_hits >= args.stable:
                ready = True
                print(f"  → 승인 건수 {n} 안정({args.stable}회 연속) → 진행")
        prev_count = n

        if ready:
            break
        if el >= args.max_wait:
            if n > 0:
                print(f"  → 최대대기 초과, 현재 {n}건으로 진행")
                break
            print("  → 최대대기 초과했지만 승인 0건. 종료.")
            return
        time.sleep(args.interval)

    # ── 다운로드 → 로드 → 빌드 ──
    ck = _read_cookie(args)
    print(f"[다운로드] {len(items)}건…")
    files = []
    for it in items:
        try:
            files.append(_NB(f"{it['req_id']}.zip", SR.download_zip(ck, it["zippath"])))
        except Exception as e:
            print(f"  ! {it['req_id']} 실패: {e}")
    if not files:
        print("받은 자료 없음. 종료.")
        return
    raw = L.load_raw_from_uploaded_files(files)
    sgg = BB.list_sigungu(raw)
    yp = args.year_pop or (max(_yrs(raw, ["to_in", "in_age", "ho_yr", "ho_ar"])) or 2024)
    yb = args.year_biz or (max(_yrs(raw, ["to_fa", "cp_bem"])) or 2023)
    print(f"[빌드] 시군구 {len(sgg)}곳 · 기준연도 인구{yp}/산업{yb}")

    def _cb(done, total, code):
        print(f"  {done}/{total}  {code}", flush=True)

    zbytes, summary = BB.build_batch_zip(
        raw, name_map=dong_names.default_name_map(),
        sido_name_map=dict(getattr(SR, "SIDO_LIST", [])),
        decimals=args.decimals, final_only=args.final_only, formula_mode=(not args.values),
        selected_years=None, year_pop=yp, year_biz=yb, out_dir=args.output, progress=_cb)
    zip_path = os.path.join(args.output, "쇠퇴진단_전국배치.zip")
    with open(zip_path, "wb") as f:
        f.write(zbytes)
    ok = int((summary["상태"] == "OK").sum()) if len(summary) else 0
    print(f"[완료] 성공 {ok}/{len(summary)}곳 · {int(time.time()-t0)}s · 저장: {args.output}")


def _yrs(raw, keys):
    out = set()
    for k in keys:
        df = raw.get(k)
        if df is not None and len(df):
            out |= set(int(y) for y in df["연도"].dropna().unique())
    return out or {0}


if __name__ == "__main__":
    main()
