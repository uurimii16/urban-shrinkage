# -*- coding: utf-8 -*-
"""step2 렌더 스모크 — 골든 원시 주입 후 ②설정 화면을 AppTest로 실제 렌더해 예외 없는지 확인."""
import openpyxl
from streamlit.testing.v1 import AppTest
from verify_jeonju import GOLDEN, load_raw_from_golden

wb = openpyxl.load_workbook(GOLDEN, data_only=True, read_only=True)
raw = load_raw_from_golden(wb)
years = sorted({int(y) for df in raw.values() for y in df["연도"].dropna().unique()})

at = AppTest.from_file("app_v2.py", default_timeout=120)
at.session_state["raw"] = raw
at.session_state["selected_years"] = years
at.session_state["step"] = 2
# 복제 시나리오까지 태우려고 계산식 사본 1개 미리 주입
import recipe_engine as RE
at.session_state["recipes"] = [RE.seed_from_builtin("노년부양비")]
at.run()

if at.exception:
    print("EXCEPTION:")
    for e in at.exception:
        print(" ", e.value)
    raise SystemExit(1)
print("STEP2 RENDER OK — 예외 없음")
print("  data_editor 수:", len(at.get("data_editor")) if hasattr(at, "get") else "n/a")
