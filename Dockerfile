# Cloudtype 등에서 이 앱을 '있는 그대로' 빌드/실행하기 위한 설정.
# 이 파일이 있으면 시작명령·포트를 대시보드에서 직접 입력할 필요가 없다.
FROM python:3.11-slim

WORKDIR /app

RUN pip install --no-cache-dir --upgrade pip
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Cloudtype가 PORT 환경변수를 주입하면 그 포트로, 없으면 8501.
ENV PORT=8501
EXPOSE 8501

# 셸 형태 CMD여야 ${PORT} 가 실제 값으로 치환된다.
CMD streamlit run app_v2.py --server.port ${PORT} --server.address 0.0.0.0 --server.headless true
