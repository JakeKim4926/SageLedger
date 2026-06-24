from __future__ import annotations

import struct
import time
from pathlib import Path

import win32clipboard
import win32con
from pywinauto import Application
from pywinauto.keyboard import send_keys

# PC 카카오톡 창 식별값 (버전에 따라 달라질 수 있어 한 곳에 모아둔다)
MAIN_TITLE = "카카오톡"
MAIN_CLASS = "EVA_Window_Dblclk"


class KakaoNotReady(Exception):
    """카카오톡이 실행/로그인되어 있지 않을 때."""


def _copy_text_to_clipboard(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def _copy_file_to_clipboard(path: Path) -> None:
    """파일을 '복사한 것처럼' 클립보드에 올린다 (CF_HDROP). Ctrl+V 로 첨부됨."""
    full = str(Path(path).resolve())
    # DROPFILES 구조체(20바이트) + 더블널 종료 utf-16 파일 목록
    header = struct.pack("Iiiii", 20, 0, 0, 0, 1)  # pFiles=20, pt(0,0), fNC=0, fWide=1
    files = (full + "\0\0").encode("utf-16-le")
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_HDROP, header + files)
    finally:
        win32clipboard.CloseClipboard()


def ensure_kakao_ready() -> Application:
    """선행 체크: 로그인된 카카오톡 메인 창이 떠 있어야 한다. 없으면 중단."""
    try:
        app = Application(backend="win32").connect(
            title=MAIN_TITLE, class_name=MAIN_CLASS, timeout=2
        )
    except Exception:
        raise KakaoNotReady(
            "카카오톡 메인 창을 찾지 못했습니다.\n"
            "PC 카카오톡을 실행하고 로그인한 뒤 다시 시도하세요."
        )
    return app


def open_room(room: str, timeout: float = 6.0):
    """메인 창 검색으로 방을 열고, 방 이름이 제목인 채팅창 핸들을 돌려준다.

    주의: 이 단계가 카카오톡 UI 변화에 가장 취약하다. 깨지면 검색창 컨트롤
    탐색(child_window) 부분을 사용자 PC에 맞게 조정해야 한다.
    """
    app = ensure_kakao_ready()
    main = app.window(title=MAIN_TITLE, class_name=MAIN_CLASS)
    main.set_focus()

    # 메인 창 검색 입력칸(첫 Edit 컨트롤)에 방 이름을 붙여넣고 Enter 로 최상단 결과 열기
    search = main.child_window(class_name="Edit", found_index=0)
    search.set_focus()
    send_keys("^a{BACKSPACE}")  # 기존 검색어 비우기
    _copy_text_to_clipboard(room)
    send_keys("^v")
    time.sleep(0.6)
    send_keys("{ENTER}")

    # 방 창(제목 == room)이 뜰 때까지 대기
    deadline = time.time() + timeout
    last_err: Exception | None = None
    while time.time() < deadline:
        try:
            room_app = Application(backend="win32").connect(title=room, timeout=1)
            return room_app.window(title=room)
        except Exception as e:  # 아직 안 열림
            last_err = e
            time.sleep(0.4)

    raise KakaoNotReady(
        f"'{room}' 방 창을 열지 못했습니다. 방 이름이 정확한지, 카카오톡 설정에서 "
        f"'채팅방을 새 창으로 열기'가 켜져 있는지 확인하세요. (원인: {last_err})"
    )


def send_file(room_win, path: Path) -> None:
    """채팅창에 파일을 붙여넣고 전송 확인창을 Enter 로 승인한다."""
    room_win.set_focus()  # 창이 활성화되면 입력칸에 커서가 간다
    _copy_file_to_clipboard(Path(path))
    send_keys("^v")
    time.sleep(0.8)  # 전송 확인 팝업 대기
    send_keys("{ENTER}")
    time.sleep(0.4)


def send_text(room_win, text: str) -> None:
    """채팅창에 텍스트 메시지를 붙여넣고 전송한다 (한글 안전하게 클립보드 사용)."""
    room_win.set_focus()
    _copy_text_to_clipboard(text)
    send_keys("^v")
    time.sleep(0.3)
    send_keys("{ENTER}")
    time.sleep(0.3)
