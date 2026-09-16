# -*- coding: utf-8 -*-
"""
NeaT 공공급식 입찰공고 수집 (웹 버전 / Gradio).

수집 로직은 콘솔·데스크톱 버전(이주희_공공급식조달자동화.py)을 그대로 가져다 쓰고,
화면만 웹으로 바꾼 것이다. 브라우저만 있으면 되니까 남에게 잠깐 보여주기 좋다.

[데스크톱(PySide6) 버전과 다른 점]
 1) 연도를 범위가 아니라 "한 해"만 고른다 (1월 ~ 12월).
 2) 공고가 여러 건일 때 물어보지 않는다.
    웹은 중간에 사용자에게 물어보고 기다리기가 어려워서, 조건에 맞는 첫 건을
    자동으로 쓰고 결과표 비고에 "여러 건 중 첫 건 사용"이라고 남긴다.

[실행]
    python 이주희_공공급식조달자동화_WEB.py
    -> http://127.0.0.1:7860

[남에게 잠깐 보여줄 때]
    아래 SHARE 를 True 로 바꾸면 gradio가 임시 공개주소(*.gradio.live)를 만들어준다.
    (72시간짜리 임시 주소다. 브라우저는 서버 PC에서 도니까 그 PC는 켜져 있어야 한다.)
"""

import os
import sys
import queue
import tempfile
import threading
import importlib.util
from datetime import datetime

import gradio as gr

# 수집 로직은 옆에 있는 콘솔 버전 파일을 그대로 불러다 쓴다.
_BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
CORE_MODULE_PATH = os.path.join(_BASE_DIR, "이주희_공공급식조달자동화.py")
_spec = importlib.util.spec_from_file_location("neat_collect_core", CORE_MODULE_PATH)
core = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(core)

from playwright.sync_api import sync_playwright


# 공개 임시주소(*.gradio.live)를 만들지 여부
SHARE = False
SERVER_PORT = 7860

TABLE_HEADERS = ["대상월", "공고번호", "기초가격", "낙찰예정가", "투찰률", "공고건명 / 비고"]


# ---------------------------------------------------------------------------
# 수집 (백그라운드 스레드 -> 큐 -> 화면)
# ---------------------------------------------------------------------------

def collect_in_thread(inst_nm, year, item_keyword, cancel_event, event_queue):
    """작업 스레드. 진행상황을 전부 큐에 넣고, 끝나면 ("done"/"error", ...)를 넣는다.

    Playwright의 sync API는 자기를 시작한 스레드에서만 쓸 수 있어서,
    sync_playwright() 자체를 이 스레드 안에서 연다.
    """
    def log(message):
        event_queue.put(("log", message))

    try:
        periods = core.build_periods_for_year(year)
        if not periods:
            event_queue.put(("error", "조회할 달이 없습니다. 연도를 확인해주세요."))
            return

        event_queue.put(("total", len(periods)))
        log("수요기관명 : %s" % inst_nm)
        log("수집연도   : %d년 (%s ~ %s, %d개월)"
            % (year, periods[0]["label"], periods[-1]["label"], len(periods)))
        log("품목 키워드 : %s" % item_keyword)
        log("")
        log("크로미움 확인 중...")
        core.ensure_chromium_installed()

        with sync_playwright() as playwright:
            log("사이트 접속 중... (넥사크로 로딩에 20초 정도 걸립니다)")
            browser, page = core.open_site(playwright, headless=True)
            try:
                if cancel_event.is_set():
                    event_queue.put(("cancelled", None))
                    return

                log("입찰공고 화면 여는 중...")
                core.open_bid_list_screen(page)
                log("입찰공고 화면 이동 완료\n")

                records = core.collect(
                    page, inst_nm, periods,
                    log=log,
                    choose_func=None,          # 웹에서는 되묻지 않고 첫 건을 쓴다
                    item_keyword=item_keyword,
                    stop_check=cancel_event.is_set,
                    on_record=lambda record: event_queue.put(("record", record)),
                )
            finally:
                browser.close()

        if cancel_event.is_set():
            event_queue.put(("cancelled", records))
            return

        found = [r for r in records if r["base_price"]]
        if not found:
            event_queue.put(("error", "수집된 데이터가 없습니다. 수요기관명이 정확한지 확인해주세요."))
            return

        # 웹에서는 서버 아무 데나 저장하면 안 되니 임시폴더에 만들고 다운로드로 내려준다.
        out_dir = tempfile.mkdtemp(prefix="neat_")
        file_path = core.save_excel(inst_nm, records, out_dir)

        log("")
        log("%d개월 중 %d개월 수집 완료" % (len(records), len(found)))
        log("엑셀 파일 준비 완료")
        event_queue.put(("done", (file_path, len(records), len(found))))

    except Exception as error:
        log("[오류] %s" % error)
        event_queue.put(("error", str(error)))


def record_to_row(record):
    rate = core.bid_rate(record)
    note = record["bid_nm"] or ""
    if record["note"]:
        note = ("%s  [%s]" % (note, record["note"])).strip()

    return [
        record["label"],
        record["bid_no"] or "-",
        format(record["base_price"], ",") if record["base_price"] else "-",
        format(record["plan_price"], ",") if record["plan_price"] else "-",
        ("%.5f" % rate) if rate is not None else "-",
        note or "-",
    ]


# ---------------------------------------------------------------------------
# 화면에서 부르는 함수 (제너레이터 : yield 할 때마다 화면이 갱신된다)
# ---------------------------------------------------------------------------

def start_collect(inst_nm, year, item_keyword, session):
    inst_nm = (inst_nm or "").strip()
    item_keyword = (item_keyword or "").strip() or core.DEFAULT_ITEM_KEYWORD
    year = int(year)

    if not inst_nm:
        yield (status_html("수요기관명을 입력해주세요.", "warn"), [], "",
               gr.update(visible=False),
               gr.update(interactive=True), gr.update(interactive=False))
        return

    cancel_event = threading.Event()
    session["cancel"] = cancel_event

    event_queue = queue.Queue()
    worker = threading.Thread(
        target=collect_in_thread,
        args=(inst_nm, year, item_keyword, cancel_event, event_queue),
        daemon=True,
    )
    worker.start()

    rows, log_lines, total, done = [], [], 0, 0

    # 버튼 상태만 먼저 바꿔서 "시작됐다"는 걸 바로 보여준다.
    yield (status_html("준비 중...", "run"), rows, "",
           gr.update(visible=False),
           gr.update(interactive=False), gr.update(interactive=True))

    while True:
        try:
            kind, payload = event_queue.get(timeout=0.4)
        except queue.Empty:
            if not worker.is_alive():
                break
            continue

        if kind == "log":
            log_lines.append(payload)
            yield (status_html("수집 중...  %d / %d" % (done, total) if total else "수집 중...", "run"),
                   rows, "\n".join(log_lines), gr.update(visible=False),
                   gr.update(interactive=False), gr.update(interactive=True))

        elif kind == "total":
            total = payload

        elif kind == "record":
            rows = rows + [record_to_row(payload)]
            done += 1
            yield (status_html("수집 중...  %d / %d" % (done, total), "run"),
                   rows, "\n".join(log_lines), gr.update(visible=False),
                   gr.update(interactive=False), gr.update(interactive=True))

        elif kind == "done":
            file_path, month_count, found_count = payload
            yield (status_html("완료 · %d개월 중 %d개월 수집" % (month_count, found_count), "ok"),
                   rows, "\n".join(log_lines), gr.update(value=file_path, visible=True),
                   gr.update(interactive=True), gr.update(interactive=False))
            return

        elif kind == "cancelled":
            yield (status_html("중지됨", "warn"), rows, "\n".join(log_lines), gr.update(visible=False),
                   gr.update(interactive=True), gr.update(interactive=False))
            return

        elif kind == "error":
            log_lines.append(payload)
            yield (status_html(payload, "warn"), rows, "\n".join(log_lines), gr.update(visible=False),
                   gr.update(interactive=True), gr.update(interactive=False))
            return

    yield (status_html("종료됨", "warn"), rows, "\n".join(log_lines), gr.update(visible=False),
           gr.update(interactive=True), gr.update(interactive=False))


def stop_collect(session):
    """중지 버튼. 작업 스레드에 깃발만 꽂으면 조회 중인 달까지만 하고 멈춘다."""
    cancel_event = session.get("cancel")
    if cancel_event is not None:
        cancel_event.set()
    return status_html("중지하는 중... (조회 중인 달까지만 마칩니다)", "warn")


def status_html(text, kind="idle"):
    colors = {
        "idle": ("#eef1f7", "#5b6473"),
        "run": ("#e7effc", "#2b55bb"),
        "ok": ("#e8f5ec", "#1d6b38"),
        "warn": ("#fdf1f2", "#a63a46"),
    }
    background, color = colors.get(kind, colors["idle"])
    return (
        '<div class="status-pill" style="background:%s;color:%s;">%s</div>'
        % (background, color, text)
    )


def year_hint(year):
    periods = core.build_periods_for_year(int(year))
    if not periods:
        return '<div class="hint">조회할 달이 없습니다</div>'

    minutes = max(1, round(len(periods) * 5 / 60))
    return ('<div class="hint">%s ~ %s · %d개월 · 약 %d분 소요</div>'
            % (periods[0]["label"], periods[-1]["label"], len(periods), minutes))


# ---------------------------------------------------------------------------
# 화면
# ---------------------------------------------------------------------------

CSS = """
.gradio-container { max-width: 1120px !important; background: #f4f6fb; }
footer { display: none !important; }

#header h1 { font-size: 24px; font-weight: 800; margin: 0 0 4px 0; color: #141824; }
#header p { font-size: 13px; color: #737d8c; margin: 0; }

.card {
    background: #ffffff !important;
    border: 1px solid #e3e8f0 !important;
    border-radius: 12px !important;
    padding: 16px 18px !important;
}
.card-title {
    font-size: 13px; font-weight: 700; color: #141824; margin-bottom: 10px;
}
.hint { font-size: 11.5px; color: #97a0ae; margin-top: 4px; }

.status-pill {
    display: inline-block; padding: 7px 14px; border-radius: 999px;
    font-size: 12.5px; font-weight: 700;
}

#log textarea {
    background: #1c2130 !important;
    color: #ccd4e0 !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: Consolas, D2Coding, monospace !important;
    font-size: 12px !important;
    line-height: 1.6 !important;
}
"""


def build_app():
    # gradio 6부터 theme/css 는 Blocks 가 아니라 launch() 에 넘긴다.
    with gr.Blocks(title="NeaT 공공급식 입찰공고 수집") as demo:
        session = gr.State({})

        gr.HTML(
            '<div id="header">'
            '<h1>NeaT 공공급식 입찰공고 수집</h1>'
            '<p>수요기관의 입찰공고를 한 달씩 조회해 기초가격·낙찰예정가를 모아 엑셀로 만듭니다.</p>'
            '</div>'
        )

        with gr.Group(elem_classes="card"):
            gr.HTML('<div class="card-title">조회 조건</div>')
            with gr.Row():
                inst_input = gr.Textbox(
                    label="수요기관명",
                    placeholder="예: 기산중학교",
                    info="사이트에 등록된 이름 그대로 입력하세요",
                    scale=3,
                )
                year_input = gr.Dropdown(
                    label="수집 연도",
                    choices=[str(y) for y in range(datetime.today().year, 2014, -1)],
                    value=str(datetime.today().year),
                    info="선택한 해의 1월~12월을 가져옵니다",
                    scale=1,
                )
                keyword_input = gr.Textbox(
                    label="품목 키워드",
                    value=core.DEFAULT_ITEM_KEYWORD,
                    info="공고가 여러 건일 때 고를 기준",
                    scale=1,
                )
            hint_html = gr.HTML(year_hint(datetime.today().year))

        with gr.Group(elem_classes="card"):
            gr.HTML('<div class="card-title">실행</div>')
            with gr.Row():
                run_button = gr.Button("수집 시작", variant="primary", scale=0)
                stop_button = gr.Button("중지", interactive=False, scale=0)
                status_box = gr.HTML(status_html("대기 중"))

        with gr.Group(elem_classes="card"):
            gr.HTML('<div class="card-title">수집 결과</div>')
            result_table = gr.Dataframe(
                headers=TABLE_HEADERS,
                datatype=["str"] * len(TABLE_HEADERS),
                column_count=(len(TABLE_HEADERS), "fixed"),
                value=[],
                wrap=True,
                interactive=False,
                show_label=False,
            )
            # 미리보기 상자(gr.File) 대신 누르면 바로 받아지는 버튼을 쓴다.
            download_button = gr.DownloadButton(
                "엑셀 파일 내려받기", variant="primary", visible=False,
            )

        with gr.Group(elem_classes="card"):
            gr.HTML('<div class="card-title">진행 로그</div>')
            log_box = gr.Textbox(
                value="", lines=12, max_lines=12, show_label=False,
                interactive=False, elem_id="log", autoscroll=True,
            )

        year_input.change(lambda y: year_hint(y), inputs=year_input, outputs=hint_html)

        run_event = run_button.click(
            fn=start_collect,
            inputs=[inst_input, year_input, keyword_input, session],
            outputs=[status_box, result_table, log_box, download_button, run_button, stop_button],
        )

        # 중지: 작업 스레드에 깃발을 꽂고(stop_collect), 화면 갱신 제너레이터도 멈춘다(cancels).
        stop_button.click(fn=stop_collect, inputs=session, outputs=status_box, cancels=[run_event])

    return demo


def main():
    app = build_app()
    app.queue()          # 제너레이터(진행상황 스트리밍)를 쓰려면 큐가 필요하다
    app.launch(
        theme=gr.themes.Soft(primary_hue="blue", neutral_hue="slate"),
        css=CSS,
        server_name="0.0.0.0",
        server_port=SERVER_PORT,
        share=SHARE,
        inbrowser=True,
    )


if __name__ == "__main__":
    main()


# if __name__ == '__main__':
#     # Render가 지정하는 환경변수 PORT 연결 필수
#     port = int(os.environ.get('PORT', 7860))
#     demo.launch(server_name="0.0.0.0", server_port=port)
