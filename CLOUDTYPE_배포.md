# Cloudtype 배포 가이드 (국내 서버 → SGIS 신청 가능)

> **왜 Cloudtype?** Streamlit Cloud는 해외(미국) 서버라 SGIS 정부서버(sgis.mods.go.kr)가
> 해외 IP를 막아 신청이 타임아웃됩니다. Cloudtype는 **국내 서버**라 SGIS가 열려요.
> 동료들은 발급된 **URL 하나로 웹에서 신청까지** 할 수 있습니다. (코드 변경 없음)

## 준비물
- GitHub 리포: `uurimii16/urban-shrinkage` (이미 Public)
- 이 리포에 포함됨: `requirements.txt`, `.streamlit/config.toml`(프록시 설정), 엔트리 `app_v2.py`

## 배포 단계 (10분)
1. **https://cloudtype.io** 접속 → **GitHub로 로그인**(리포 접근 허용).
2. **새 프로젝트** → **GitHub** 선택 → 리포 `uurimii16/urban-shrinkage`, 브랜치 `main` 선택.
3. 템플릿은 **Python** (또는 Streamlit 프리셋이 있으면 그것) 선택.
4. **설정값 입력**:
   - **설치 명령어**: `pip install -r requirements.txt`
   - **시작 명령어**:
     ```
     streamlit run app_v2.py --server.port ${PORT} --server.address 0.0.0.0
     ```
   - **포트**: Cloudtype가 주는 `${PORT}` 자동사용(시작 명령에 이미 반영). 별도 지정 요구 시 `8501`.
   - **Python 버전**: 3.11 (선택 가능하면).
5. **배포하기** → 빌드 로그가 끝나면 **URL 발급**(예: `https://urban-shrinkage.cloudtype.app`).
6. 발급된 URL 접속 → ① 자료신청 화면에서 **쿠키 붙여넣기 → 📥 전국 지역목록 불러오기** 눌러
   시군구가 뜨면 성공(= 국내 IP에서 SGIS 열림). 동료들에게 이 URL 공유.

## 확인 포인트
- "📥 전국 지역목록 불러오기"가 **타임아웃 없이** 시도/시군구 목록을 채우면 정상.
- 여전히 타임아웃이면: 시작 명령의 포트/주소, requirements 설치 로그 확인.

## 참고
- 기존 Streamlit Cloud(`urban-shrinkage.streamlit.app`)는 분석·산출용으론 계속 사용 가능
  (단 SGIS 신청 단계만 해외 IP라 불가). 헷갈리면 Cloudtype URL 하나로 통일 권장.
- Cloudtype 무료 플랜은 미사용 시 절전될 수 있음 → 접속하면 다시 깨어남(수 초).
- 신청은 각자 **본인 SGIS 로그인 쿠키**가 필요(로그인 후 F12 → Copy as cURL → 붙여넣기).

## 매일 강제중지(404) 자동 복구 — API 자동 재배포
> Cloudtype 무료는 **하루 1회 강제 중지**가 있어, 그 시간대엔 URL이 404가 됩니다.
> 단순 접속(핑)으로는 안 켜지고 **재배포**해야 다시 켜져요.
> 이 UI엔 '자동배포 토글'이 없어서, **Cloudtype 공식 API**로 재배포를 자동화합니다.

**설정(딱 1번):**
1. Cloudtype **API 키 발급**: 오른쪽 위 스페이스 아이콘 → 스페이스 설정 → **인증** → **새 API 키 생성** → 복사.
2. GitHub **Secret 등록**: 리포 → Settings → Secrets and variables → **Actions** → **New repository secret**
   → Name `CLOUDTYPE_TOKEN`, Secret = 위 API 키 → Add secret.
3. 끝. `.github/workflows/cloudtype-redeploy.yml`이 **10분마다 헬스체크** →
   살아있으면 warm 유지(콜드스타트 방지), **죽어있으면 즉시 Cloudtype API로 재배포**(자동복구).
   → 강제중지돼도 10분 내 자동으로 다시 켜져 동료가 느린 순간을 거의 안 겪음.
   (project=urban-shrinkage, stage=main, workload=app 기준으로 박혀 있음)

**확인:** GitHub 리포 → Actions 탭 → "Cloudtype Auto Redeploy" → **Run workflow**(수동 실행) →
초록 체크(✓)로 끝나고 몇 분 뒤 URL이 열리면 성공. 실패(빨강)면 로그 열어 에러문구 확인
(대개 배포 스펙의 `name/ports/preset` 불일치 → 그 값만 맞추면 됨).

- 대안: 매일 손이 가는 게 싫고 확실히 하려면 **Cloudtype 유료**(월 몇천원)로 '하루 1회 중지' 자체가 사라짐.
