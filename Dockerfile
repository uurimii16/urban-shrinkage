# Cloudtype 등에서 이 앱을 '있는 그대로' 빌드/실행하기 위한 설정.
# 이 파일이 있으면 시작명령·포트를 대시보드에서 직접 입력할 필요가 없다.
FROM python:3.11-slim

WORKDIR /app

RUN pip install --no-cache-dir --upgrade pip
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 콜드스타트(부팅) 중 브라우저 탭에 잠깐 뜨는 기본 제목 'Streamlit'을 우리 제목으로 교체.
# (앱이 뜨면 set_page_config 제목으로 바뀌지만, 로딩 몇 초 동안 이 index.html 제목이 보임)
RUN python -c "import streamlit,os; p=os.path.join(os.path.dirname(streamlit.__file__),'static','index.html'); h=open(p,encoding='utf-8').read(); h=h.replace('<title>Streamlit</title>','<title>쇠퇴진단 자동화 시스템</title>'); open(p,'w',encoding='utf-8').write(h); print('patched index.html title')"

COPY . .

# Cloudtype가 PORT 환경변수를 주입하면 그 포트로, 없으면 8501.
ENV PORT=8501
EXPOSE 8501

# 셸 형태 CMD여야 ${PORT} 가 실제 값으로 치환된다.
CMD streamlit run app_v2.py --server.port ${PORT} --server.address 0.0.0.0 --server.headless true
