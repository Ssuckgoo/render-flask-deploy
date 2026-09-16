import os

from flask import Flask

app = Flask(__name__)


@app.route('/')
def home():
    return "24시간 웹 서비스가 작동 중입니다!", 200


if __name__ == '__main__':
    # 환경 변수 PORT가 있으면 그 값을, 없으면 5000번 포트를 사용
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)