import gradio as gr

# 1. 실행할 파이썬 함수 정의
def greet(name):
    return f"안녕하세요, {name}님!"

# 2. 인터페이스 생성 (함수, 입력 형태, 출력 형태 지정)
demo = gr.Interface(
    fn=greet, 
    inputs="text", 
    outputs="text"
)

# 3. 웹 서버 실행
demo.launch()
