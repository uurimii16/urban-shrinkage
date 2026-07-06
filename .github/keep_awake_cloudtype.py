"""Cloudtype에 올린 Streamlit 앱을 주기적으로 깨우는 스크립트.

Cloudtype 무료 플랜은 '매일 1회 중지'(+미사용 절전)가 있다. Cloudtype는 요청이 오면
컨테이너를 다시 시작하므로, 단순 HTTP GET만 보내면 깨울 수 있다(브라우저 불필요).
콜드 스타트(재기동)에는 수십 초가 걸릴 수 있어 200이 올 때까지 잠깐 재시도한다.

URL은 환경변수/레포 Variable `CLOUDTYPE_URL` 로 주입.
"""
import os
import time
import urllib.request

URL = os.environ.get("CLOUDTYPE_URL", "").strip().rstrip("/")
if not URL:
    # 배포 전이라 아직 URL이 없으면 조용히 건너뜀(실패로 처리하지 않음).
    print("CLOUDTYPE_URL 미설정 - 건너뜀. 배포 후 repo Variable에 URL을 넣으세요.")
    raise SystemExit(0)

# 헬스 엔드포인트(가벼움) → 루트 순으로 시도
TARGETS = [URL + "/_stcore/health", URL + "/"]


def ping(u):
    req = urllib.request.Request(u, headers={"User-Agent": "keep-awake/1.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status


ok = False
for attempt in range(1, 9):          # 최대 8회(재시도 간 15초) ≈ 2분간 스핀업 대기
    for t in TARGETS:
        try:
            s = ping(t)
            print(f"[{attempt}] {t} -> {s}")
            if s == 200:
                ok = True
                break
        except Exception as e:
            print(f"[{attempt}] {t} -> 실패: {e}")
    if ok:
        break
    time.sleep(15)

if ok:
    print("앱 깨어남 확인 OK")
else:
    print("[경고] 200 응답을 못 받음 - Cloudtype 대시보드에서 앱 상태/자동시작 설정을 확인하세요.")
    raise SystemExit(1)
