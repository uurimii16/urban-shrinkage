# -*- coding: utf-8 -*-
"""
verify_clone_seeds.py — 기본지표 '복제 씨앗'이 원본 계산을 재현하는지 검증
==========================================================================
decline_engine 의 12지표 값(원본)과, recipe_engine.BUILTIN_SEEDS 로 만든
계산식 사본의 eval_recipe 값(복제)을 행정동·집계구 두 레벨에서 셀단위 대조.
  · exact=True  지표: 최대오차 ~0 이어야 통과
  · exact=False 지표(제조/고차/음식숙박): 8차(<=2005) 코드분기 차이 → 오차 허용, 참고표시
데이터: data/golden.xlsx (verify_jeonju 의 원시 long 로더 재사용).
"""
import sys
import numpy as np
import openpyxl
import decline_engine as E
import recipe_engine as RE
from verify_jeonju import GOLDEN, load_raw_from_golden

wb = openpyxl.load_workbook(GOLDEN, data_only=True, read_only=True)
raw = load_raw_from_golden(wb)

print(f"원시 로드: " + ", ".join(f"{k}={len(v)}행" for k, v in raw.items()))
print("=" * 74)

ok_all = True
for level in ("dong", "jgu"):
    key = "집계구" if level == "jgu" else "행정동코드"
    orig = E.derive_indicators(raw, key)            # 현행 기본지표 값 DF(11개)
    idx = orig.index
    # 소형주택비율은 기본지표에서 빠졌지만 복제 씨앗은 남아있으므로 엔진 함수로 직접 참조 계산.
    extra_ref = {"소형주택비율": lambda: E.derive_small_housing_ratio(raw["ho_ar"], key)}
    print(f"\n[{('행정동' if level=='dong' else '집계구')} 레벨]  지역 {len(idx)}개")
    print(f"  {'지표':<16}{'구분':<6}{'최대절대오차':>14}   판정")
    for name, seed in RE.BUILTIN_SEEDS.items():
        recipe = RE.seed_from_builtin(name)
        got = RE.eval_recipe(recipe, raw, level).reindex(idx).astype(float)
        ref = orig[name] if name in orig.columns else extra_ref[name]()
        base = ref.reindex(idx).astype(float)
        diff = np.nanmax(np.abs(got.values - base.values)) if len(idx) else 0.0
        exact = seed["exact"]
        if exact:
            verdict = "✓ 일치" if diff < 1e-6 else "✗ 불일치!"
            ok_all = ok_all and diff < 1e-6
        else:
            verdict = f"≈ 근사(8차분기)"
        print(f"  {name:<16}{'정확' if exact else '근사':<6}{diff:>14.6g}   {verdict}")

print("\n" + "=" * 74)
print("결과:", "✅ exact 지표 전부 원본과 일치" if ok_all else "❌ exact 지표에 불일치 있음 — 씨앗 코드 점검 필요")
sys.exit(0 if ok_all else 1)
