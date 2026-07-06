# -*- coding: utf-8 -*-
"""
sgis_probe2.py — SGIS OpenAPI 2차 탐침(코드표 실측)
====================================================
1차에서 확인: 인증OK / adm_cd 8자리(=행정동) / 총인구·사업체·종사자 총계는 2024까지.
2차 목표: 문서가 감춘 '코드값'을 실측으로 캐낸다.
  A) 성연령 searchpopulation: 어떤 age_type/gender 조합이 5세계급을 주나
  B) 주택 house: const_year(건축연도) 코드가 뭐뭐 있고 몇 년까지 주나
  C) 산업 종사자수: 대분류별(제조/도소매/음식숙박…) 종사자수를 어떻게 받나

실행:  python 쇠퇴진단엔진/sgis_probe2.py <consumer_key> <consumer_secret>
표준 라이브러리만 사용.
"""
from __future__ import annotations
import os, sys, json, urllib.parse, urllib.request

for _s in (sys.stdout, sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

BASE = "https://sgisapi.kostat.go.kr/OpenAPI3"
WANSAN = "35011"   # 완산구


def _get(path, params):
    q = urllib.parse.urlencode(params)
    with urllib.request.urlopen(f"{BASE}/{path}?{q}", timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def auth(key, secret):
    j = _get("auth/authentication.json", {"consumer_key": key, "consumer_secret": secret})
    tok = (j.get("result") or {}).get("accessToken")
    if not tok:
        print("[인증실패]", json.dumps(j, ensure_ascii=False)); sys.exit(1)
    print("[OK] 인증 성공\n")
    return tok


def one(path, params):
    """단일호출 → (errCd, result). low_search=0 → 완산구 1행."""
    try:
        j = _get(path, params)
    except Exception as e:
        return f"EXC:{type(e).__name__}", None
    return j.get("errCd"), j.get("result"), j.get("errMsg")


def main():
    key = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("SGIS_KEY", "")
    secret = sys.argv[2] if len(sys.argv) > 2 else os.environ.get("SGIS_SECRET", "")
    if not key or not secret:
        print("사용법: python 쇠퇴진단엔진/sgis_probe2.py <key> <secret>"); sys.exit(2)
    tok = auth(key, secret)
    base = {"accessToken": tok, "adm_cd": WANSAN, "low_search": "0"}

    # ── A) 성연령: age_type 스윕(2020) ──
    print("="*70, "\nA) 성연령 searchpopulation — age_type 스윕 (year=2020, gender=0)\n", "="*70)
    for at in list(range(0, 27)):
        err, res, msg = (one("stats/searchpopulation.json",
                             {**base, "year": "2020", "gender": "0", "age_type": str(at)}) + (None,))[:3]
        if err in (0, "0") and res:
            r0 = res[0] if isinstance(res, list) else res
            print(f"  age_type={at:>2}: OK 행수={len(res) if isinstance(res,list) else 1} "
                  f"필드={list(r0.keys()) if isinstance(r0,dict) else r0}  샘플={json.dumps(r0, ensure_ascii=False)[:200]}")
        else:
            print(f"  age_type={at:>2}: - ({err} {msg})")

    # 성연령 연도 커버리지(작동 age_type 하나로 2015~2024 확인)
    print("\n  [성연령 연도 커버리지] age_type=0, gender=0")
    for yr in (2015, 2016, 2018, 2020, 2021, 2022, 2023, 2024):
        err, res, msg = (one("stats/searchpopulation.json",
                             {**base, "year": str(yr), "gender": "0", "age_type": "0"}) + (None,))[:3]
        print(f"    {yr}: {'OK' if err in (0,'0') and res else '- '+str(msg)}")

    # ── B) 주택: const_year 스윕(2020) ──
    print("\n" + "="*70, "\nB) 주택 house.json — const_year 스윕 (year=2020)\n", "="*70)
    for cy in list(range(0, 13)):
        err, res, msg = (one("stats/house.json", {**base, "year": "2020", "const_year": str(cy)}) + (None,))[:3]
        if err in (0, "0") and res:
            r0 = res[0] if isinstance(res, list) else res
            print(f"  const_year={cy:>2}: OK house_cnt={r0.get('house_cnt') if isinstance(r0,dict) else r0}  필드={list(r0.keys()) if isinstance(r0,dict) else ''}")
        else:
            print(f"  const_year={cy:>2}: - ({err} {msg})")
    print("\n  [주택 연도 커버리지] const_year=1")
    for yr in (2005, 2010, 2015, 2020, 2023, 2024):
        err, res, msg = (one("stats/house.json", {**base, "year": str(yr), "const_year": "1"}) + (None,))[:3]
        print(f"    {yr}: {'OK' if err in (0,'0') and res else '- '+str(msg)}")

    # ── C) 산업: 대분류 종사자수 ──
    print("\n" + "="*70, "\nC) 산업 종사자수 — 산업분류 코드 + company.json class_code\n", "="*70)
    # 산업분류표(10차 대분류) 목록
    err, res, msg = (one("stats/industrycode.json", {"accessToken": tok, "class_deg": "10"}) + (None,))[:3]
    if err in (0, "0") and res:
        print("  [10차 산업 대분류 목록]")
        for r in (res if isinstance(res, list) else [])[:25]:
            print("   ", r.get("class_code"), r.get("class_nm"))
    else:
        print("  industrycode.json:", err, msg)
    # company.json에 class_code 넣어 대분류 종사자수 나오나(제조 C, 도소매 G, 숙박음식 I 가정)
    print("\n  [company.json + class_code 시도] year=2019")
    for cc in ("C", "G", "I", "J", "M"):
        err, res, msg = (one("stats/company.json", {**base, "year": "2019", "class_code": cc}) + (None,))[:3]
        r0 = (res[0] if isinstance(res, list) and res else res) if res else None
        print(f"    class_code={cc}: {'OK '+json.dumps(r0, ensure_ascii=False)[:150] if r0 else '- '+str(msg)}")
    print("\n  [company.json 연도 커버리지] (총계)")
    for yr in (2019, 2020, 2021, 2022, 2023, 2024):
        err, res, msg = (one("stats/company.json", {**base, "year": str(yr)}) + (None,))[:3]
        print(f"    {yr}: {'OK' if err in (0,'0') and res else '- '+str(msg)}")


if __name__ == "__main__":
    main()
