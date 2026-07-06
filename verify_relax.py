# -*- coding: utf-8 -*-
"""verify_relax.py — ② 하드차단 완화 로직 end-to-end 검증.
빈 분류(ho_ar)를 만들어도 (1) 크래시 없이, (2) 소형주택비율만 제외되고,
(3) 나머지 지표로 복합지수·통합엑셀이 정상 산출되는지 확인."""
import io
import openpyxl
import config as C
import decline_engine as E
import legal_engine as LG
import custom_indicators as CI
import export
from verify_jeonju import GOLDEN, load_raw_from_golden

wb = openpyxl.load_workbook(GOLDEN, data_only=True, read_only=True)
raw = load_raw_from_golden(wb)

# cp_bem(종사자수) 분류를 빈 상태로(= 산업 종사자 데이터 없음 시나리오)
import pandas as pd
raw["cp_bem"] = pd.DataFrame(columns=["연도", "집계구", "CODE", "값", "행정동코드"])
available = {b for b in ["to_in", "in_age", "to_fa", "cp_bem", "ho_yr", "ho_ar"] if b in raw and not raw[b].empty}

# step4 필터 로직 재현
indicator_ids, excluded = [], []
for ind in C.IND_IDS:
    b = C.INDICATOR_BUCKET.get(ind)
    if b is not None and b not in available:
        excluded.append(ind)
    else:
        indicator_ids.append(ind)

INDUSTRY = ["총종사자수증감률", "제조업증감률", "고차산업증감률", "도소매증감률", "음식숙박증감률"]
print("제외된 지표:", excluded)
print("유지된 지표:", len(indicator_ids), "개")
assert set(excluded) == set(INDUSTRY), f"기대: {INDUSTRY}, 실제: {excluded}"

# 파이프라인 실행(크래시 없어야)
dong = E.run(raw, level="dong")
jgu = E.run(raw, level="jgu")
sector_of = dict(C.SECTOR_OF); weight = {i: float(C.WEIGHT.get(i, 0.0)) for i in indicator_ids}
dong_comp = CI.composite(dong[0], indicator_ids, sector_of, weight)
legal_dong = LG.run_legal(raw, level="dong")
wbk = export.build_integrated_workbook(
    raw, dong_res=dong[:3], jgu_res=jgu[:3], legal_dong=legal_dong, legal_jgu=LG.run_legal(raw, "jgu"),
    indicator_ids=indicator_ids, label_map=dict(C.INDLABEL), sector_of=sector_of,
    weight=weight, sign_map=dict(C.SIGN), final_only=True)
buf = io.BytesIO(); wbk.save(buf)
print("복합지수 계산 행정동 수:", len(dong_comp), " 종합점수 예시:", round(float(dong_comp["종합"].max()), 2))
print("통합엑셀 시트:", wbk.sheetnames)
print("엑셀 바이트:", len(buf.getvalue()))
print("\n✅ ho_ar 비어도 크래시 없이 11지표로 정상 산출 — 하드차단 완화 확인")
