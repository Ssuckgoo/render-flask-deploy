# Render 클라우드 웹 호스팅 24시간 구동 테스트

### Rendor는 무료사이트가 아니기 때문에, 15분 동안 트래픽이 없으면 자동으로 서버가 다운된다.     그렇기 때문에 uptimerobot 이라는 툴을 사용해 5분마다 테스트 트래픽을 보내서 서버가 다운되지 않도록 한다이건 온전히 테스트 용도이다.

### 사용 툴
* **GitHub**,
**Render**,
**uptimerobot**

### 소스 버전관리
* GitHub에 소스 업로드

### 소스 배포
1. Render.com 접속
2. Web Services -> New WebServie
3. 깃계정 연동 후, 프로젝트 선택
4. Deploy web service

### 타이머 세팅
1. https://uptimerobot.com 접속
2. New
3. 얼마나 자주 통신을 보낼 지 세팅
4. Create monitor
