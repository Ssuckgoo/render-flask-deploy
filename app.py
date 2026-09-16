import os
import gradio as gr

def greet(name):
    return f"안녕하세요, {name}님!"

demo = gr.Interface(
    fn=greet, 
    inputs="text", 
    outputs="text"
)

if __name__ == '__main__':
    # Render가 지정하는 환경변수 PORT 연결 필수
    port = int(os.environ.get('PORT', 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
