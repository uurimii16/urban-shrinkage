# -*- coding: utf-8 -*-
"""
sgis_probe.py — SGIS OpenAPI 탐침(1회용)
=========================================
목적: 본격 수집기 만들기 전에, SGIS가 '실제로' 뭘 어떤 형식으로 주는지 확인.
      - 인증(accessToken) 되는지
      - 지역코드 자릿수(읍면동 7 vs 8), 응답 필드명
      - 연도 커버리지(2024까지 주는지)
      - 성연령/건축연도/사업체 테마가 읍면동 단위로 나오는지

실행(둘 중 편한 것):
  python sgis_probe.py <consumer_key> <consumer_secret>
  # 또는 환경변수:  SGIS_KEY / SGIS_SECRET 설정 후  python sgis_probe.py

키 발급: https://sgis.kostat.go.kr/developer  → 로그인 → 인증키 발급(서비스 신청).
표준 라이브러리만 사용(pip 불필요).
"""
from __future__ import annotations

import os
import sys
import json
import urllib.parse
import urllib.request

# 윈도우 콘솔(cp949)에서도 한글·기호 출력이 깨지거나 죽지 않도록 UTF-8로 강제.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE = "https://sgisapi.kostat.go.kr/OpenAPI3"

# 전주시: 완산구 35011 · 덕진구 35012 (5자리 시군구코드)
JEONJU = {"완산구": "35011", "덕진구": "35012"}


def _get(url, params):
    q = urllib.parse.urlencode(params)
    full = f"{url}?{q}"
    with urllib.request.urlopen(full, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def auth(key, secret):
    j = _get(f"{BASE}/auth/authentication.json",
             {"consumer_key": key, "consumer_secret": secret})
    if str(j.get("errCd")) not in ("0", "None") and j.get("errCd") not in (0, None):
        print("[인증 실패]", json.dumps(j, ensure_ascii=False, indent=2))
        sys.exit(1)
    token = j.get("result", {}).get("accessToken")
    if not token:
        print("[인증 응답 이상]", json.dumps(j, ensure_ascii=False, indent=2))
        sys.exit(1)
    print("✔ 인증 성공 — accessToken 획득\n")
    return token


def probe(token, name, path, extra, years):
    """한 테마를 여러 연도로 찔러보고 응답 요약 출력."""
    print(f"\n{'='*70}\n■ {name}  ({path})\n{'='*70}")
    for yr in years:
        params = {"accessToken": token, "year": str(yr),
                  "adm_cd": JEONJU["완산구"], "low_search": "1", **extra}
        try:
            j = _get(f"{BASE}/{path}", params)
        except Exception as e:
            print(f"  {yr}: 요청오류 {type(e).__name__} {e}")
            continue
        err = j.get("errCd")
        result = j.get("result")
        if err not in (0, "0", None) or not result:
            print(f"  {yr}: 데이터없음/오류 errCd={err} msg={j.get('errMsg')}")
            continue
        n = len(result) if isinstance(result, list) else 1
        sample = result[0] if isinstance(result, list) and result else result
        keys = list(sample.keys()) if isinstance(sample, dict) else type(sample)
        print(f"  {yr}: OK  행수={n}  필드={keys}")
        if yr == years[-1] and isinstance(sample, dict):
            print("     └ 샘플1건:", json.dumps(sample, ensure_ascii=False)[:400])


def main():
    key = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("SGIS_KEY", "")
    secret = sys.argv[2] if len(sys.argv) > 2 else os.environ.get("SGIS_SECRET", "")
    if not key or not secret:
        print("사용법: python sgis_probe.py <consumer_key> <consumer_secret>")
        print("   또는 환경변수 SGIS_KEY / SGIS_SECRET 설정 후 실행")
        sys.exit(2)

    token = auth(key, secret)

    # 연도 커버리지 확인용(과거·최근·최신)
    yrs = [2000, 2015, 2020, 2023, 2024]

    # 1) 인구총괄(총인구) — to_in
    probe(token, "인구총괄(총인구) → to_in", "stats/population.json", {}, yrs)
    # 2) 성연령별 인구 — in_age (조건검색: 성별·연령대)
    probe(token, "성연령 인구 → in_age", "stats/searchpopulation.json",
          {"gender": "0", "age_type": "1"}, yrs)
    # 3) 주택(건축연도/연면적) — ho_yr / ho_ar
    probe(token, "주택 → ho_yr/ho_ar", "stats/house.json", {}, yrs)
    # 4) 사업체·종사자 — to_fa / cp_bem
    probe(token, "사업체·종사자 → to_fa/cp_bem", "stats/company.json", {}, yrs)

    print("\n\n※ 위 결과에서 확인할 것:")
    print("  · 각 테마의 최신 연도(2023/2024 나오나?)")
    print("  · adm_cd 자릿수(읍면동 코드 몇 자리인가) — 집계구[:8]=8자리와 비교")
    print("  · 성연령/건축연도/종사자 값이 어떤 필드로 나오는가(→ 항목코드 매핑 설계)")


if __name__ == "__main__":
    main()
