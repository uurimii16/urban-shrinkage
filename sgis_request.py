# -*- coding: utf-8 -*-
"""
sgis_request.py — SGIS 자료제공 '신청(제출)' 자동화 (승인 전까지만)
====================================================================
장바구니에 넣고 신청완료(`saveRequestData`) 하는 부분을 스크립트 1번으로.
※ 다운로드는 승인(약 10분) 후라 자동화 대상 아님 — 신청내역에서 수동 다운로드.

동작: 아래 CART(완산·덕진 × 6통계, 차수별 사업체 포함)를 한 신청으로 제출.
      SGIS 화면에서 '추가' 24번 누르던 걸 대체.

실행:
  1) 브라우저로 SGIS(sgis.mods.go.kr) 로그인 → 아무 요청이나 F12 Network에서
     'Copy as cURL' 해서, 그 안 -b "..." 의 쿠키 문자열을 통째로 복사.
  2) 먼저 미리보기(제출 안 함):
       python sgis_request.py "쿠키문자열"
  3) 진짜 신청(주의: 실제로 신청됨):
       python sgis_request.py "쿠키문자열" --submit
  4) 처음엔 1건만 시험 권장:
       python sgis_request.py "쿠키문자열" --submit --only-first

표준 라이브러리만 사용.
"""
from __future__ import annotations
import sys
import uuid
import re
import urllib.parse
import urllib.request

# ── 신청자 정보 기본값(개인정보는 비움 — 앱 ①화면에서 입력받아 override) ────────
#   공개 배포 대비: 이메일·연락처·회사·SGIS계정 등 개인정보는 코드에 두지 않는다.
APPLICANT = {
    "param_userkey": "",            # SGIS 로그인 계정(아이디)
    "sgis_census_req_sosok_nm": "민간",
    "sgis_census_req_sosok": "001005",
    "sgis_census_req_mokjuk_nm": "공간DB구축",
    "sgis_census_req_mokjuk": "002001",
    "sgis_census_req_company": "",   # 회사/소속명
    "sgis_census_req_tel_1": "",     # 연락처
    "sgis_census_req_tel_2": "",
    "sgis_census_req_tel_3": "",
    "email_id": "",                  # 이메일 아이디
    "email_addr": "naver.com",       # 이메일 도메인
    "email_addr_select": "naver.com",
    "sgis_census_req_goal": "복합쇠퇴진단",
    "sgis_census_req_kwaje": "도시재생 기초자료",
}

SIGUNGU = ["35011", "35012"]          # 완산구·덕진구
AREA_YEAR = "2025"                      # 집계구 경계연도
OUTPUT_TYPE = "1"                       # 1 = 집계구 CSV

# (data_id, detail_code, 대표연도, 라벨) — 대표연도는 해당 차수/자료의 기준연도
ITEMS = [
    ("0", "in_04", "2024", "총인구"),
    ("0", "in_02", "2024", "성연령별인구"),
    ("2", "ho_06", "2024", "건축년도별주택"),
    ("2", "ho_04", "2024", "연건평별주택"),
    # 총괄사업체수(to_fa) — 차수별
    ("3", "cp_05", "2005", "총괄사업체수(8차)"),
    ("3", "cp_06", "2015", "총괄사업체수(9차)"),
    ("3", "cp_09", "2020", "총괄사업체수(10차)"),
    ("3", "cp_27", "2024", "총괄사업체수(11차)"),
    # 대분류 종사자수(cp_bem) — 차수별
    ("3", "cp_01", "2005", "종사자수(8차)"),
    ("3", "cp_02", "2015", "종사자수(9차)"),
    ("3", "cp_08", "2020", "종사자수(10차)"),
    ("3", "cp_26", "2024", "종사자수(11차)"),
]

URL = "https://sgis.mods.go.kr/view/pss/saveRequestData"
OPT_URL = "https://sgis.mods.go.kr/view/pss/requestOptionData"
REQ_PAGE_URL = "https://sgis.mods.go.kr/view/pss/requestData"   # 세션 프라이밍용(신청 전 GET 필수)

# ── SGIS 실측 제공연도 ──
#   인구·주택(인구주택총조사): 2000·2005·2010·2015 + 2016~2024 매년 = 13개년
#     (2001~04·2006~09·2011~14는 조사를 안 해 '원래 없음')
#   사업체(전국사업체조사): 2000~2024 매년 = 25개년
YEARS_POP = [2000, 2005, 2010, 2015] + list(range(2016, 2025))
YEARS_BIZ = list(range(2000, 2025))

# 사업체 세부코드는 차수(연도)마다 다름 — 실측 매핑(연도→코드)
def _biz_jong(y):   # 대분류 종사자수 → cp_bem
    return "cp_01" if y <= 2005 else ("cp_02" if y <= 2016 else ("cp_08" if y <= 2023 else "cp_26"))
def _biz_chong(y):  # 총괄사업체수 → to_fa
    return "cp_05" if y <= 2005 else ("cp_06" if y <= 2016 else ("cp_09" if y <= 2023 else "cp_27"))

# ── 앱 체크박스용 항목 카탈로그 (라벨 → [(data_id, code, 연도),...]) ──
#   시계열(증감률)용은 '전 연도'를 각각 명시 신청 / 스냅샷용은 최신연도만.
ITEM_CATALOG = {
    "총인구(전연도)":        [("0", "in_04", str(y)) for y in YEARS_POP],
    "성연령별인구(연도선택)":  [("0", "in_02", str(y)) for y in YEARS_POP],
    "건축년도별주택(연도선택)": [("2", "ho_06", str(y)) for y in YEARS_POP],
    "연면적별주택(최신)":     [("2", "ho_04", "2024")],
    "종사자수(전연도 2000~2024)":   [("3", _biz_jong(y), str(y)) for y in YEARS_BIZ],
    "총괄사업체수(전연도 2000~2024)": [("3", _biz_chong(y), str(y)) for y in YEARS_BIZ],
    "가구총괄(최신)":        [("1", "ga_03", "2024")],
    "세대구성별가구(최신)":   [("1", "ga_02", "2024")],
    "주택유형별주택(최신)":   [("2", "ho_03", "2024")],
    "주택총괄(최신)":        [("2", "ho_05", "2024")],
}
# 쇠퇴진단 엔진이 실제 쓰는 기본 선택 = 복합쇠퇴 전체세트(필수 5종)
DEFAULT_CHECKED = ["총인구(전연도)", "성연령별인구(연도선택)", "건축년도별주택(연도선택)",
                   "종사자수(전연도 2000~2024)", "총괄사업체수(전연도 2000~2024)"]

# 항목 → (필수/선택, 어떤 지표에 쓰이는지 설명)
ITEM_META = {
    "총인구(전연도)":       ("필수", "인구변화율(인문사회) — 2000·05·10·15 + 2016~2024 매년(13개년)"),
    "성연령별인구(연도선택)": ("필수", "노년부양비·경제활동인구비율·소멸위험지수(인문사회) — 아래 ‘성연령 연도’에서 고른 해"),
    "건축년도별주택(연도선택)": ("필수", "노후건축물비율(물리환경) — 아래 ‘건축연도 연도’에서 고른 해"),
    "종사자수(전연도 2000~2024)":   ("필수", "총·제조·고차·도소매·음식숙박 종사자 증감률(산업경제) — 매년, 차수 자동"),
    "총괄사업체수(전연도 2000~2024)": ("필수", "총사업체수 증감률(산업경제) — 매년, 차수 자동"),
    "연면적별주택(최신)":    ("선택", "소형주택비율 — 기본에서 제외됨. ‘복제’로 되살릴 때만"),
    "가구총괄(최신)":       ("선택", "1인가구율 등 가구 기반 새 지표 만들 때"),
    "세대구성별가구(최신)":  ("선택", "세대구성 기반 새 지표 만들 때"),
    "주택유형별주택(최신)":  ("선택", "아파트비율 등 주택유형 새 지표 만들 때"),
    "주택총괄(최신)":       ("선택", "총주택수(분모) 등 새 지표 만들 때"),
}

# ── 부문별 개별 연도 선택 메타 ───────────────────────────────────────────
#   도메인키 → (라벨, 제공연도, 기본선택연도, 도움말)
#   증감률(총인구·사업체·종사자)=전 연도 기본. 스냅샷(성연령·건축)=최신만 기본이나
#   다른 해로 비교하려면 그 해를 추가 신청 가능(엔진은 기준연도 1개만 사용해 안전).
YEAR_DOMAINS = {
    "pop":  ("인문사회 · 총인구 연도", YEARS_POP, list(YEARS_POP),
             "인구변화율용(증감률). 여러 해가 있어야 정확 → 전 연도 권장."),
    "age":  ("인문사회 · 성연령별인구 연도", YEARS_POP, [2024],
             "노년부양비·경제활동·소멸위험용. 분석엔 기준연도 1개면 충분하지만, 다른 해로 비교하려면 그 해도 체크."),
    "biz":  ("산업경제 · 사업체/종사자 연도", YEARS_BIZ, list(YEARS_BIZ),
             "총사업체·종사자 증감률용(증감률). 전 연도 권장."),
    "phys": ("물리환경 · 건축연도별주택 연도", YEARS_POP, [2024],
             "노후건축물비율용. 기준연도 1개면 충분하지만 다른 해 비교하려면 추가."),
}
# 항목명 → 도메인키 (연도 선택 대상 항목)
YEAR_DOMAIN = {
    "총인구(전연도)": "pop",
    "성연령별인구(연도선택)": "age",
    "종사자수(전연도 2000~2024)": "biz",
    "총괄사업체수(전연도 2000~2024)": "biz",
    "건축년도별주택(연도선택)": "phys",
}
# 항상 최신 1개년만 신청하는 스냅샷 항목(연도 선택 대상 아님) — 선택 항목들
SNAPSHOT_ITEMS = {"연면적별주택(최신)", "가구총괄(최신)", "세대구성별가구(최신)",
                  "주택유형별주택(최신)", "주택총괄(최신)"}

# 앱 '전체 분석에 필요한 데이터' 설명(마크다운)
NEED_EXPLAIN = """**복합쇠퇴분석 = 11개 지표를 3부문(인문사회·산업경제·물리환경)으로 종합.**
아래 **5종만** 받으면 전체 결과표가 나옵니다.

| 부문 | 지표 | 필요한 SGIS 데이터 | 연도 |
|---|---|---|---|
| 인문사회 | 인구변화율 | **총인구** | 2000~2024 |
| 인문사회 | 노년부양비·경제활동인구비율·소멸위험지수 | **성연령별인구** | 최신 |
| 산업경제 | 총사업체수 증감률 | **총괄사업체수**(전차수) | 2000~2024 |
| 산업경제 | 총·제조·고차·도소매·음식숙박 종사자 증감률 | **종사자수**(전차수) | 2000~2024 |
| 물리환경 | 노후건축물비율 | **건축년도별주택** | 최신 |

- **증감률 지표**(인구·사업체·종사자)는 *전 기간 최댓값*과 비교 → **여러 해가 있어야** 정확 → 2000~2024 전체 수집.
- **성연령·건축년도**는 최신 1개년만 쓰지만 전 연도 같이 와도 무해.
- ❗ **총괄종사자수는 안 받아도 돼요** — 대분류 종사자수를 합해 총종사자를 계산합니다.
- **연면적·가구·주택유형**은 기본 분석엔 불필요 — 새 지표(부실건축물·1인가구율 등) 만들 때만."""


REQUIRED_COOKIES = ("JSESSIONID", "accessToken")


def extract_cookie(raw):
    """붙여넣은 쿠키/‘Copy as cURL’ 통째 문자열에서 JSESSIONID·accessToken만 추출."""
    parts = []
    for k in REQUIRED_COOKIES:
        m = re.search(k + r'=([^;^"\'\s]+)', raw)
        if m:
            parts.append(f"{k}={m.group(1)}")
    return "; ".join(parts)


def found_cookies(raw):
    """붙여넣은 문자열에 필수 쿠키가 각각 있는지 확인. {'JSESSIONID':bool,'accessToken':bool}.
    UI에서 어느 값이 잡혔는지 개별 표시하는 데 사용."""
    raw = raw or ""
    return {k: bool(re.search(k + r'=([^;^"\'\s]+)', raw)) for k in REQUIRED_COOKIES}


def make_cart(items, sigungu_codes, only_first=False):
    """items=[(data_id, code, year, label)] × 시군구 → cart 목록.
    sido_id = 대표연도 + 시도2자리(시군구 앞2자리)."""
    cart = []
    for sg in sigungu_codes:
        for it in items:
            did, det, yr = it[0], it[1], it[2]
            lab = it[3] if len(it) > 3 else det
            cart.append((did, det, yr, f"{yr}{sg[:2]}", sg, f"{lab}·{sg}"))
    return cart[:1] if only_first else cart


def prime_session(cookie):
    """★ 신청 등록의 필수 선행단계 ★
    saveRequestData 는 같은 세션이 먼저 requestData 페이지를 GET 해야만 신청을
    '신청내역에 등록'한다. 프라이밍 없이 POST 하면 200·응답 1/2 가 와도 **등록되지
    않는다**(=조용한 실패). 실측 확정(2026-07-08, 국내IP): 프라이밍 없이는 미등록,
    동일 요청도 GET requestData 직후엔 등록됨(JSESSIONID+accessToken 2쿠키로 충분).
    반환: 성공 여부(bool). 실패해도 신청은 시도한다."""
    try:
        req = urllib.request.Request(REQ_PAGE_URL, headers={
            "Cookie": cookie, "User-Agent": "Mozilla/5.0",
            "Referer": REQ_PAGE_URL})
        with urllib.request.urlopen(req, timeout=30) as r:
            r.read()
        return True
    except Exception:
        return False


def submit_cart(cookie, cart, applicant=None):
    """cart → saveRequestData 제출 → (status, response_text).
    applicant: 신청자 정보 override dict(앱 화면 입력값). None이면 APPLICANT 기본값.
    ※ 제출 전 prime_session()으로 세션을 준비해야 신청이 실제로 등록된다."""
    prime_session(cookie)                     # ★ 필수: 없으면 신청이 등록 안 됨
    boundary = "WebKitFormBoundary" + uuid.uuid4().hex[:16]
    body = build_body(cart, boundary, applicant=applicant)
    return submit(cookie, body, boundary)

# 집계구 통계 대분류: 0=인구, 1=가구, 2=주택, 3=사업체
DATA_CATS = {"0": "인구", "1": "가구", "2": "주택", "3": "사업체"}


def _post_option(cookie, data):
    """requestOptionData 로 POST → 응답 HTML 문자열(표준 라이브러리)."""
    req = urllib.request.Request(OPT_URL, data=urllib.parse.urlencode(data).encode(),
        headers={"Cookie": cookie, "X-Requested-With": "XMLHttpRequest",
                 "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                 "Referer": "https://sgis.mods.go.kr/view/pss/requestData",
                 "User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=25) as r:
        return r.read().decode("utf-8", "replace")


def _parse_options(html):
    """<option value="35011">완산구</option> → [("35011","완산구"), ...].
    빈 값·'00'·'선택하세요' 같은 플레이스홀더는 제외."""
    out = []
    for val, txt in re.findall(r'<option[^>]*value="([^"]*)"[^>]*>(.*?)</option>', html, re.S):
        val = val.strip()
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", txt)).strip()
        if not val or val in ("0", "00", "00000") or not txt or "선택" in txt or txt == "전체":
            continue
        out.append((val, txt))
    return out


# SGIS 자료제공 시도 코드(2자리). mode=6 응답은 '시군구' 전용이라 시도는 이 고정표를 쓴다.
# (실측 검증됨 2026-07-07: 각 코드로 fetch_sigungu_list 호출 시 정상 시군구 반환 —
#  서울25·부산16·대구9·인천10·광주5·대전5·울산5·세종1·경기44·강원18·충북14·충남16·
#  전북15·전남22·경북23·경남22·제주2)
SIDO_LIST = [
    ("11", "서울특별시"), ("21", "부산광역시"), ("22", "대구광역시"), ("23", "인천광역시"),
    ("24", "광주광역시"), ("25", "대전광역시"), ("26", "울산광역시"), ("29", "세종특별자치시"),
    ("31", "경기도"), ("32", "강원도"), ("33", "충청북도"), ("34", "충청남도"),
    ("35", "전북특별자치도"), ("36", "전라남도"), ("37", "경상북도"), ("38", "경상남도"),
    ("39", "제주특별자치도"),
]


def fetch_sido_list(cookie=None, year="2024"):
    """전국 시도 목록 [(2자리코드, 이름)].
    SGIS mode=6는 시군구 전용이라 시도는 고정표(SIDO_LIST)를 반환(네트워크 호출 없음).
    cookie/year 인자는 호출부 호환용."""
    return list(SIDO_LIST)


def fetch_sigungu_list(cookie, sido_code, year="2024"):
    """해당 시도의 시군구 목록 [(시군구5자리코드, 이름)].
    mode=6, year_sido=<year><시도2자리> (예: 202435=전북)."""
    html = _post_option(cookie, {
        "sgis_census_id": "1", "sgis_census_data_id": "0", "sgis_census_req_id": "",
        "year_sido": f"{year}{sido_code}", "census_output_area_year": AREA_YEAR,
        "inUse": "inUse1", "years": "years1", "mode": "6"})
    return _parse_options(html)


# ── 승인 자료 자동 다운로드 (신청→승인 후) ──────────────────────────────────
#   흐름(브라우저 분석으로 확정):
#     1) GET  /view/pss/downloadList        → 승인·다운로드 가능 목록(HTML)
#        각 행: 신청번호(req_id) · 자료명 · 생성일 · 만료일 · zip경로(/census/reqdoc/<번호>.zip)
#     2) POST /contents/include/download.jsp (filename=<zip경로>) → 실제 zip 바이트
#   ※ SGIS 정부서버라 국내 IP에서만 동작(해외=차단).
DL_LIST_URL = "https://sgis.mods.go.kr/view/pss/downloadList"
DL_FILE_URL = "https://sgis.mods.go.kr/contents/include/download.jsp"


def fetch_download_list(cookie):
    """다운로드 가능(승인완료) 목록 파싱 → [{req_id, label, date, expire, zippath}]."""
    req = urllib.request.Request(DL_LIST_URL, headers={
        "Cookie": cookie, "User-Agent": "Mozilla/5.0",
        "Referer": "https://sgis.mods.go.kr/view/index",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ko-KR,ko;q=0.9"})
    with urllib.request.urlopen(req, timeout=30) as r:
        html = r.read().decode("utf-8", "replace")
    m = re.search(r"<tbody.*?</tbody>", html, re.S | re.I)
    body = m.group(0) if m else html
    out = []
    for tr in re.findall(r"<tr.*?</tr>", body, re.S | re.I):
        oc = re.search(r"census_download\([^)]*?,\s*['\"]([^'\"]+\.zip)['\"]", tr)
        if not oc:
            continue
        zippath = oc.group(1)
        rid = re.search(r"sgis_census_req_id=(\d+)", tr)
        tds = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", td)).strip()
               for td in re.findall(r"<td.*?</td>", tr, re.S | re.I)]
        out.append({
            "req_id": (rid.group(1) if rid else (tds[0] if tds else "")),
            "label": tds[2] if len(tds) > 2 else "",
            "date": tds[3] if len(tds) > 3 else "",
            "expire": tds[4] if len(tds) > 4 else "",
            "zippath": zippath,
        })
    return out


def download_zip(cookie, zippath):
    """POST download.jsp (filename=zip경로) → zip 바이트. 실패 시 예외."""
    data = urllib.parse.urlencode({"filename": zippath, "path": ""}).encode()
    req = urllib.request.Request(DL_FILE_URL, data=data, headers={
        "Cookie": cookie, "User-Agent": "Mozilla/5.0",
        "Referer": "https://sgis.mods.go.kr/view/pss/downloadList",
        "Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=120) as r:
        blob = r.read()
    if blob[:2] != b"PK":
        raise ValueError(f"zip이 아님(응답 앞부분: {blob[:60]!r}) — 쿠키 만료/만료된 자료일 수 있음")
    return blob


def list_items(cookie):
    """SGIS에서 '선택 가능한 모든 세부항목 코드'를 조회해 출력.
    ITEMS에 넣을 (data_id, detail_code, year) 를 여기서 골라 쓰면 됨."""
    def opt(data):
        req = urllib.request.Request(OPT_URL, data=urllib.parse.urlencode(data).encode(),
            headers={"Cookie": cookie, "X-Requested-With": "XMLHttpRequest",
                     "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                     "Referer": "https://sgis.mods.go.kr/view/pss/requestData",
                     "User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.read().decode("utf-8", "replace")

    # 사업체는 차수가 연도마다 달라 대표연도 4개를 모두 훑는다.
    years_by_cat = {"0": ["2024"], "1": ["2024"], "2": ["2024"],
                    "3": ["2005", "2015", "2020", "2024"]}
    print("=== 선택 가능한 세부항목 (data_id / 코드 / 라벨) ===")
    for did, cat in DATA_CATS.items():
        seen = set()
        print(f"\n[{did}] {cat}")
        for yr in years_by_cat[did]:
            body = opt({"sgis_census_id": "1", "sgis_census_data_id": did,
                        "sgis_census_req_id": "", "sgis_census_year": yr,
                        "census_output_area_year": AREA_YEAR, "inUse": "inUse1",
                        "years": "years1", "mode": "4", "codeValue": ""})
            pairs = re.findall(r'value="((?!din_)[a-z]{2,3}_\d+)"[^>]*/?>\s*<label[^>]*>([^<]+)</label>', body)
            if not pairs:
                codes = re.findall(r'value="((?!din_)[a-z]{2,3}_\d+)"', body)
                labs = re.findall(r'<label[^>]*>([^<]+)</label>', body)[1:]
                pairs = list(zip(codes, labs))
            for c, l in pairs:
                if c in seen:
                    continue
                seen.add(c)
                print(f'   ("{did}", "{c}", "{yr}", "{l.strip()}"),')
    print('\n→ 위 줄을 복사해 sgis_request.py 의 ITEMS 리스트에 넣거나 빼세요.')


def build_cart(only_first=False):
    """CLI용: 상단 ITEMS × SIGUNGU 로 cart 생성."""
    return make_cart(ITEMS, SIGUNGU, only_first=only_first)


def _part(name, value):
    return (f'Content-Disposition: form-data; name="{name}"\r\n\r\n{value}\r\n')


def build_body(cart, boundary, applicant=None):
    """saveRequestData multipart 본문 생성(캡처 구조 미러링).
    applicant: 신청자 정보 override dict — 없는 키는 APPLICANT 기본값 사용."""
    A = {**APPLICANT, **(applicant or {})}
    B = f"------{boundary}"
    out = []

    def add(name, value):
        out.append(B + "\r\n" + _part(name, value))

    # 1) 신청자/공통
    add("param_userkey", A["param_userkey"])
    add("aT", "INS")
    add("sgis_census_req_id", "")
    add("old_census_file", "")
    add("inUse", "")
    add("years", "")
    add("sgis_census_req_sosok_nm", A["sgis_census_req_sosok_nm"])
    add("sgis_census_req_mokjuk_nm", A["sgis_census_req_mokjuk_nm"])
    add("sgis_census_req_sosok", A["sgis_census_req_sosok"])
    add("sgis_census_req_company", A["sgis_census_req_company"])
    add("sgis_census_req_tel_1", A["sgis_census_req_tel_1"])
    add("sgis_census_req_tel_2", A["sgis_census_req_tel_2"])
    add("sgis_census_req_tel_3", A["sgis_census_req_tel_3"])
    add("email_id", A["email_id"])
    add("email_addr", A["email_addr"])
    add("email_addr_select", A["email_addr_select"])
    add("sgis_census_req_mokjuk", A["sgis_census_req_mokjuk"])
    add("sgis_census_req_goal", A["sgis_census_req_goal"])
    add("sgis_census_req_kwaje", A["sgis_census_req_kwaje"])
    add("concur", "on")

    # 2) 대표(첫) 항목 — top-level
    d0 = cart[0]
    add("census_output_data_type", OUTPUT_TYPE)
    add("sgis_census_id", "1")
    add("sgis_census_data_id", d0[0])
    add("census_output_area_year", AREA_YEAR)
    add("sgis_census_year_id", d0[2])
    add("sgis_census_detail_data_id", d0[1])
    add("sgis_census_sido_id", d0[3])
    add("sgis_census_sigungu_id", d0[4])
    add("sgis_all", "on")

    # 3) 장바구니 전체 — _new 반복 블록
    for did, det, yr, sido, sg, lab in cart:
        add("cbox", "on")
        add("sgis_census_id_new", "1")
        add("sgis_census_data_id_new", did)
        add("sgis_census_year_id_new", yr)
        add("sgis_census_detail_data_id_new", det)
        add("census_output_data_type_new", OUTPUT_TYPE)
        add("sgis_census_req_map_level", "-")
        add("sgis_census_sido_id_new", sido)
        add("sgis_census_req_map_code", "-")
        add("sgis_census_sigungu_id_new", sg)
        add("census_output_area_dts_year_new", AREA_YEAR)

    out.append(B + "--\r\n")
    return "".join(out).encode("utf-8")


def submit(cookie, body, boundary):
    req = urllib.request.Request(URL, data=body, headers={
        "Cookie": cookie,
        "Content-Type": f"multipart/form-data; boundary=----{boundary}",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "https://sgis.mods.go.kr",
        "Referer": "https://sgis.mods.go.kr/view/pss/requestData",
        "User-Agent": "Mozilla/5.0",
    })
    with urllib.request.urlopen(req, timeout=90) as r:
        return r.status, r.read().decode("utf-8", "replace")


def main():
    if len(sys.argv) < 2:
        print('사용법: python sgis_request.py "쿠키문자열" [--submit] [--only-first]')
        sys.exit(2)
    cookie = sys.argv[1]
    do_submit = "--submit" in sys.argv
    only_first = "--only-first" in sys.argv

    if "--list" in sys.argv:
        list_items(cookie)
        return

    cart = build_cart(only_first=only_first)
    print(f"■ 장바구니 {len(cart)}건:")
    for c in cart:
        print(f"   - {c[5]}  (data_id={c[0]} detail={c[1]} year={c[2]} sido={c[3]} sgg={c[4]})")

    boundary = "WebKitFormBoundary" + uuid.uuid4().hex[:16]
    body = build_body(cart, boundary)
    print(f"\n본문 크기: {len(body):,} bytes")

    if not do_submit:
        print("\n[미리보기 모드] 실제 신청 안 함. 진짜 신청하려면 끝에 --submit 붙이세요.")
        return
    print("\n[신청 전송 중…]")
    prime_session(cookie)                     # ★ 필수: 없으면 신청이 등록 안 됨
    status, resp = submit(cookie, body, boundary)
    print("HTTP", status)
    print("응답:", resp[:800])
    print("\n→ SGIS '신청내역'에서 접수됐는지 확인하세요. 약 10분 뒤 승인되면 다운로드.")


if __name__ == "__main__":
    main()
