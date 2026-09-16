# Playwright 공식 이미지. 크로미움과 크로미움이 쓰는 시스템 라이브러리가 이미 들어있다.
# Render 기본 파이썬 환경에는 그 라이브러리들이 없어서 크로미움이 안 뜬다(그래서 도커로 배포한다).
# 태그의 v1.61.0 은 requirements.txt 의 playwright 버전과 반드시 맞춰야 한다.
FROM mcr.microsoft.com/playwright/python:v1.61.0-noble

WORKDIR /app

# 소스가 바뀌어도 라이브러리 설치는 캐시를 쓰도록 requirements 를 먼저 복사한다.
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PYTHONUNBUFFERED=1
# 이미지가 크로미움을 넣어둔 위치. app.py/neat_core.py 가 이 값을 그대로 존중한다.
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

EXPOSE 7860

# Render 는 PORT 환경변수로 포트를 넘겨준다. app.py 가 그 값을 읽는다.
CMD ["python", "app.py"]
