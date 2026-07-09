# -*- coding: utf-8 -*-
"""template_export.py — 차장님 정본 양식(복합쇠퇴진단.xlsx) 그대로 산출.
9시트(집계구 단위): 계산방법 / 1 법적쇠퇴진단 종합 / 2 복합쇠퇴진단 종합 / 원시 6시트.
엔진 계산은 decline_engine·legal_engine 그대로 재사용, 출력 양식만 이 모듈이 담당.
지금은 계산방법 + 복합쇠퇴진단 종합 을 우선 구현(검증용). 법적·원시시트는 후속.
"""
import os
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side
import pandas as pd
import numpy as np

import decline_engine as E

# 12지표: (id, 헤더(R1), 부문, 방향부호, 기본최종가중치) — 계산방법 E24~E35에 대응
TEMPLATE_INDICATORS = [
    ('인구변화율',      '인구변화율\n(30년간 최다인구수 연도대비)',       '인문사회', -10, 0.220),
    ('노년부양비',      '2024년 노년부양비',                           '인문사회',  10, 0.058),
    ('경제활동인구비율','2024년 경제활동인구비율',                      '인문사회', -10, 0.113),
    ('소멸위험지수',    '2024년 소멸위험지수',                          '인문사회', -10, 0.009),
    ('총사업체수증감률','총사업체수 증감률\n(10년간 최다 사업체수 대비)', '산업경제', -10, 0.021),
    ('총종사자수증감률','총 종사자수 증감률\n(10년간 최다 종사자수 대비)','산업경제', -10, 0.028),
    ('제조업증감률',    '제조업 종사자수 증감률',                       '산업경제', -10, 0.057),
    ('고차산업증감률',  '고차산업종사자수 증감률',                      '산업경제', -10, 0.066),
    ('도소매증감률',    '도소매종사자수 증감률',                        '산업경제', -10, 0.024),
    ('음식숙박증감률',  '음식숙박업 종사자수 증감률',                    '산업경제', -10, 0.004),
    ('노후건축물비율',  '노후건축물비율',                              '물리환경',  10, 0.250),
    ('소형주택비율',    '19평 이하 소형주택비율',                       '물리환경',  10, 0.150),
]
SECTORS = ['인문사회', '산업경제', '물리환경']
CALC_SHEET = '계산방법(수식·알고리즘)'
# 지표 id → 정본 양식 헤더/부문/방향/가중치. 앱 cfg에 없는 값은 여기서 폴백.
_TEMPLATE_HDR = {iid: hdr for iid, hdr, *_ in TEMPLATE_INDICATORS}
_TEMPLATE_SEC = {iid: sec for iid, _h, sec, _s, _w in TEMPLATE_INDICATORS}
_TEMPLATE_SIGN = {iid: sign for iid, _h, _sec, sign, _w in TEMPLATE_INDICATORS}
_TEMPLATE_W = {iid: w for iid, _h, _sec, _s, w in TEMPLATE_INDICATORS}
# 전국 행정동명표(load_admin_names 기본 경로) — 이 파일 기준 상대경로.
DEFAULT_ADMIN_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  '행정구역코드', '행정구역코드_전국.xlsx')
_THIN = Side(style='thin', color='B0B0B0')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR_FONT = Font(bold=True, size=9)
_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def indicators_from_cfg(cfg):
    """app_v2의 ss.cfg → 정본 양식용 indicators 리스트.
    반환: [(id, 헤더, 부문, 방향부호, 최종가중치), ...]  (순서 = cfg['indicator_ids']).
    · 방향부호·최종가중치·부문 = ③설정에서 사용자가 정한 값(sign_map/final_weight/sector_of).
    · 헤더 = 정본 양식 헤더 우선, 없으면(커스텀 지표) cfg의 label_map.
    cfg가 비어 있으면 정본 기본 12지표를 그대로 돌려준다."""
    ids = list(cfg.get('indicator_ids') or [i[0] for i in TEMPLATE_INDICATORS])
    sign_map = cfg.get('sign_map', {})
    weight = cfg.get('final_weight', {})
    sector_of = cfg.get('sector_of', {})
    label_map = cfg.get('label_map', {})
    out = []
    for iid in ids:
        hdr = _TEMPLATE_HDR.get(iid) or label_map.get(iid, iid)
        sec = sector_of.get(iid) or _TEMPLATE_SEC.get(iid) or SECTORS[0]
        # cfg에 값이 있으면(0 포함) 그대로, 아예 없으면 정본 기본값으로 폴백
        sign = int(sign_map[iid]) if iid in sign_map else _TEMPLATE_SIGN.get(iid, -10)
        w = float(weight[iid]) if iid in weight else _TEMPLATE_W.get(iid, 0.0)
        out.append((iid, hdr, sec, sign, w))
    return out


def load_admin_names(path):
    """행정구역코드_전국.xlsx → (행정동코드8→행정동명, 시군구코드5→시군구명)."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb['전체'] if '전체' in wb.sheetnames else wb[wb.sheetnames[0]]
    dong, sgg = {}, {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        # 연번, 시도코드, 시도명, 시군구코드, 시군구명, 행정동코드, 행정동명
        sg, sgn, dc, dn = (str(row[3] or '').strip(), str(row[4] or '').strip(),
                           str(row[5] or '').strip(), str(row[6] or '').strip())
        if dc:
            dong[dc] = dn
        if sg:
            sgg[sg] = sgn
    wb.close()
    return dong, sgg


def compute_values(raw_sub, ids=None, key='집계구'):
    """집계구(또는 행정동)별 지표 값 DataFrame. ids 순서·구성으로 정렬.
    ids=None이면 정본 기본 12지표. 엔진이 못 만드는 지표(커스텀 등)는 NaN 열로 남고,
    종합시트의 Z/T 수식이 빈칸을 0으로 처리하므로 구조는 그대로 유지된다."""
    if ids is None:
        ids = [i[0] for i in TEMPLATE_INDICATORS]
    df = E.derive_indicators(raw_sub, key=key)            # 11지표(config)
    if '소형주택비율' in ids and '소형주택비율' not in df.columns and 'ho_ar' in raw_sub:
        df = df.copy()                                     # 소형주택 되살림(요청 시)
        df['소형주택비율'] = E.derive_small_housing_ratio(raw_sub['ho_ar'], key)
    return df.reindex(columns=list(ids))


def values_from_scores(scores, ids):
    """앱이 이미 산출한 scores(멀티인덱스 (지표,'값'/'Z'/'T')) → 정본 지표값 DataFrame.
    기본+커스텀+계산식 지표 모두 (지표,'값') 열을 갖고 있으므로 그대로 뽑아 쓴다.
    scores에 없는 지표는 NaN(빈칸)으로 남고 종합시트 함수가 0으로 처리."""
    cols = {}
    for iid in ids:
        cols[iid] = scores[(iid, '값')] if (iid, '값') in scores.columns else np.nan
    return pd.DataFrame(cols, index=scores.index)


# ── 계산방법 시트 ──────────────────────────────────────────────
def write_calc_method(ws, indicators=TEMPLATE_INDICATORS):
    """알고리즘 설명 + ④지표별 산식표. 방향부호(D24~)·최종가중치(E24~)는 여기서 주입."""
    doc = [
        '① 표준화 (Z점수 · T점수)',
        'Z = (지표값 − 평균) ÷ 모표준편차(STDEV.P)',
        '   · 평균·표준편차는 그 단위 집합(집계구) 전체에서 결측(N/A) 제외하고 계산',
        '   · 엑셀 등가:  평균=AVERAGE(범위) ,  모표준편차=STDEV.P(범위)',
        '   · 표준편차가 0이면(모두 같은 값) Z = 0 으로 처리',
        'T = Z × 방향부호 + 50',
        '   · 방향부호(±10): 값이 클수록 쇠퇴가 심하면 +10, 값이 클수록 양호하면 −10',
        '   · 즉 일반 T점수(=Z×10+50)에 쇠퇴 방향을 부호로 합친 형태 (평균 50 기준)',
        '',
        '② 가중치 · 부문점수 · 종합',
        '최종가중치 = (부문비율 ÷ 100) × (부문 내부비율 ÷ 100)',
        '   · 부문비율 : 인문사회/산업경제/물리환경 3부문 사이의 배분(합 100%)',
        '   · 내부비율 : 한 부문 안에서 지표들 사이의 배분(부문별 합 100%)',
        '부문점수 = Σ ( 지표T × 그 지표 최종가중치 )   (그 부문에 속한 지표만 합산)',
        '종합점수 = 인문사회 + 산업경제 + 물리환경  (세 부문점수의 합)',
        '   · 엑셀 등가:  가중T = T×가중치 ,  부문 = SUM(부문 가중T들) ,  종합 = SUM(부문들)',
        '',
        '③ 등급',
        '분류 방식 : Natural Breaks(Jenks) · 등분위 · 등간격 · 10등급',
        '종합점수가 클수록 쇠퇴가 심함 → 1등급(가장 쇠퇴) … 큰 등급숫자일수록 양호',
        '',
        '④ 지표별 산식 · 방향 · 최종가중치',
    ]
    for i, line in enumerate(doc, start=1):
        ws.cell(i, 1).value = line
    # 표 헤더(R23) + 지표행(R24~)
    hdr = ['지표', '부문', '산식', '방향부호', '최종가중치']
    for c, h in enumerate(hdr, start=1):
        cell = ws.cell(23, c); cell.value = h; cell.font = _HDR_FONT; cell.border = _BORDER
    for i, (iid, _hdr, sec, sign, w) in enumerate(indicators):
        r = 24 + i
        ws.cell(r, 1).value = iid
        ws.cell(r, 2).value = sec
        ws.cell(r, 3).value = ''          # 산식 설명(선택 — 후속에 채움)
        ws.cell(r, 4).value = sign        # D열 방향부호
        ws.cell(r, 5).value = w           # E열 최종가중치 ← 복합종합이 참조
        for c in range(1, 6):
            ws.cell(r, c).border = _BORDER
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['C'].width = 40


# ── 복합쇠퇴진단 종합 시트 ─────────────────────────────────────
def write_composite(ws, codes, values, dong_names, indicators=TEMPLATE_INDICATORS,
                    calc_sheet=CALC_SHEET):
    """codes: 집계구코드 리스트(정렬). values: df(집계구 index × 12지표). dong_names: 8자리→명."""
    n = len(codes)
    N = 2 + n                              # 데이터 마지막 행(3..N)
    q = lambda s: f"'{s}'"                 # 시트명 인용

    # ── R1/R2 헤더 ──
    ws.cell(1, 1).value = '행정동'; ws.cell(1, 2).value = '집계구'; ws.cell(1, 3).value = '행정동명'
    for i, (iid, hdr, sec, sign, w) in enumerate(indicators):
        vcol = 4 + i * 3
        ws.cell(1, vcol).value = hdr
        ws.merge_cells(start_row=1, start_column=vcol, end_row=1, end_column=vcol + 2)
        ws.cell(2, vcol).value = '지표값'; ws.cell(2, vcol + 1).value = 'Z점수'; ws.cell(2, vcol + 2).value = 'T점수'
    base = 4 + len(indicators) * 3          # 40 (AN)
    sec_cols = {}
    for s_i, sec in enumerate(SECTORS):
        wcol = base + s_i * 2
        ws.cell(1, wcol).value = sec
        ws.merge_cells(start_row=1, start_column=wcol, end_row=1, end_column=wcol + 1)
        ws.cell(2, wcol).value = '가중치적용'; ws.cell(2, wcol + 1).value = '에러수정'
        sec_cols[sec] = wcol
    total_col = base + len(SECTORS) * 2      # 46 (AT)
    ws.cell(1, total_col).value = '종합'; ws.cell(2, total_col).value = '가중치적용'
    ws.cell(1, total_col + 1).value = '등급'
    ws.cell(2, total_col + 1).value = '등분위'; ws.cell(2, total_col + 2).value = '등간격'
    ws.cell(2, total_col + 3).value = 'Natural Break'

    # 부문별 지표 인덱스
    sec_idx = {s: [i for i, ind in enumerate(indicators) if ind[2] == s] for s in SECTORS}

    # ── 데이터 행(R3~) ──
    for ridx, code in enumerate(codes):
        r = 3 + ridx
        dong8 = str(code)[:8]
        ws.cell(r, 1).value = dong8
        ws.cell(r, 2).value = str(code)
        ws.cell(r, 3).value = dong_names.get(dong8, '')
        # 지표값 + Z/T 수식
        for i, (iid, hdr, sec, sign, w) in enumerate(indicators):
            vcol = 4 + i * 3
            V = get_column_letter(vcol); Z = get_column_letter(vcol + 1); Tc = get_column_letter(vcol + 2)
            val = values.iloc[ridx][iid] if iid in values.columns else None
            ws.cell(r, vcol).value = (None if val is None or (isinstance(val, float) and np.isnan(val)) else float(val))
            rng = f"{V}$3:{V}${N}"
            if i == 0:
                ws.cell(r, vcol + 1).value = f"=({V}{r}-AVERAGE({rng}))/_xlfn.STDEV.P({rng})"
                ws.cell(r, vcol + 2).value = f"=({Z}{r}*{sign})+50"
            else:
                ws.cell(r, vcol + 1).value = f'=IF({V}{r}="",0,(({V}{r}-AVERAGE({rng}))/_xlfn.STDEV.P({rng})))'
                ws.cell(r, vcol + 2).value = f'=IF({V}{r}="",0,({Z}{r}*{sign})+50)'
        # 부문 가중치적용 / 에러수정
        for sec in SECTORS:
            wcol = sec_cols[sec]
            terms, terms_err = [], []
            for i in sec_idx[sec]:
                Tc = get_column_letter(4 + i * 3 + 2)
                ecell = f"{q(calc_sheet)}!$E${24 + i}"
                terms.append(f"({Tc}{r}*{ecell})")
                terms_err.append(f"IFERROR({Tc}{r}*{ecell}, 0)")
            ws.cell(r, wcol).value = "=SUM(" + ",".join(terms) + ")"
            ws.cell(r, wcol + 1).value = "=SUM(" + ",".join(terms_err) + ")"
        # 종합 = 부문 에러수정들의 합
        err_cols = [get_column_letter(sec_cols[s] + 1) + str(r) for s in SECTORS]
        ws.cell(r, total_col).value = "=SUM(" + ",".join(err_cols) + ")"
        # 등급(종합 AT열 기준)
        AT = get_column_letter(total_col)
        rngT = f"$AT$3:${AT}${N}".replace('AT', AT)  # AT range
        rngT = f"${AT}$3:${AT}${N}"
        ws.cell(r, total_col + 1).value = (
            f"=MIN(10,ROUNDUP(RANK({AT}{r},{rngT},0)/(COUNT({rngT})/10),0))")      # 등분위
        ws.cell(r, total_col + 2).value = (
            f"=MAX(1,MIN(10,ROUNDUP((MAX({rngT})-{AT}{r})/((MAX({rngT})-MIN({rngT}))/10),0)))")  # 등간격
        # Natural Break: 후속(경계+VLOOKUP). 지금은 등분위와 동일 자리 비움.
        ws.cell(r, total_col + 3).value = None

    # 표시서식 0.00 (지표값·Z·T·부문·종합)
    for r in range(3, N + 1):
        for c in range(4, total_col + 1):
            cell = ws.cell(r, c)
            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith('=')):
                cell.number_format = '0.00'
    # 헤더 스타일
    for r in (1, 2):
        for c in range(1, total_col + 4):
            cell = ws.cell(r, c); cell.font = _HDR_FONT; cell.alignment = _CENTER
    ws.freeze_panes = 'D3'
    return N


def build_composite_workbook(raw_sub=None, admin_path=None, indicators=TEMPLATE_INDICATORS,
                             values=None):
    """계산방법 + 복합쇠퇴진단 종합 2시트 wb (집계구 단위).
    indicators = TEMPLATE_INDICATORS(기본) 또는 indicators_from_cfg(cfg)로 만든
    ③설정 반영 목록. 값·계산방법·종합 셋 다 같은 목록으로 산출해 E열 참조가 일치한다.
    · values 주어지면(앱 scores에서 뽑은 base+custom+recipe 전체 지표값) 그걸 사용.
    · 없으면 raw_sub에서 기본지표만 계산(standalone·검증용)."""
    dong_names, _ = load_admin_names(admin_path or DEFAULT_ADMIN_PATH)
    ids = [i[0] for i in indicators]
    vals = values.reindex(columns=ids) if values is not None else compute_values(raw_sub, ids, key='집계구')
    vals = vals[vals.index.astype(str).str.len() == 14]      # 집계구(14자리)만
    codes = sorted(vals.index.astype(str))
    vals = vals.loc[codes]
    wb = openpyxl.Workbook()
    write_calc_method(wb.active, indicators)
    wb.active.title = CALC_SHEET
    ws2 = wb.create_sheet('2 복합쇠퇴진단 종합')
    write_composite(ws2, codes, vals, dong_names, indicators)
    return wb
