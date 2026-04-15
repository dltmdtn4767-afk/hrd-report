FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# uploads, output 디렉토리 생성
RUN mkdir -p uploads output

EXPOSE 8080

# Render는 $PORT 환경변수 주입 — 없으면 8080 사용
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}
