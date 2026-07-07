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

## 매일 강제중지(404) 자동 복구 — 자동 재배포
> Cloudtype 무료는 **하루 1회 강제 중지**가 있어, 그 시간대엔 URL이 404가 됩니다.
> 단순 접속(핑)으로는 안 켜지고 **재배포**해야 다시 켜져요. 이걸 GitHub Actions가 자동으로 합니다.

**설정(딱 1번):**
1. Cloudtype 대시보드 → 이 앱(프로젝트) 열기.
2. **설정에서 "자동배포(Auto Deploy)" 토글을 ON** 으로. (GitHub `main`에 push되면 자동 재배포)
3. 끝. 이후 `.github/workflows/cloudtype-redeploy.yml`이 **한국시간 09·13·17시에 빈 커밋을 push**
   → Cloudtype 자동 재배포 → 강제중지됐어도 다시 켜짐.

**확인:** GitHub 리포 → Actions 탭 → "Cloudtype Auto Redeploy" → **Run workflow**(수동 실행)로
한 번 돌려보고, 몇 분 뒤 URL이 열리면 성공. (자동배포 토글이 꺼져 있으면 push만 되고 재배포는 안 됨)

- 참고: 빈 커밋이 기록에 쌓이는 게 싫으면, 대신 Cloudtype **API키 + 공식 배포 액션** 방식으로
  바꿀 수 있음(토큰 발급 필요). 기능은 동일.
