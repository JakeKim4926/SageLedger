from __future__ import annotations

import ctypes
import struct
import time
from ctypes import wintypes
from pathlib import Path

import win32clipboard
import win32con
import win32gui
from pywinauto.keyboard import send_keys

MAIN_CLASS = "EVA_Window_Dblclk"
MAIN_TITLE = "카카오톡"

_GMEM_MOVEABLE = 0x0002
_WM_DROPFILES = 0x0233

_k32 = ctypes.windll.kernel32
_k32.GlobalAlloc.restype = wintypes.HGLOBAL
_k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
_k32.GlobalLock.restype = ctypes.c_void_p
_k32.GlobalLock.argtypes = [wintypes.HGLOBAL]
_k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]


class KakaoNotReady(Exception):
    """카카오톡이 실행/로그인되어 있지 않거나 방을 열지 못했을 때."""


def _set_dpi_aware() -> None:
    # 다중 모니터/배율 환경에서 좌표·창 처리 일관성 확보 (Per-Monitor v2)
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        pass


def _copy_text(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def _make_dropfiles(path: Path, x: int, y: int) -> int:
    """파일 드롭(WM_DROPFILES)용 DROPFILES 구조를 전역 메모리에 만들어 핸들을 돌려준다."""
    full = str(Path(path).resolve())
    files = (full + "\0\0").encode("utf-16-le")
    # DROPFILES: pFiles=20, pt(x,y), fNC=0, fWide=1
    blob = struct.pack("Iiiii", 20, x, y, 0, 1) + files
    h = _k32.GlobalAlloc(_GMEM_MOVEABLE, len(blob))
    ptr = _k32.GlobalLock(h)
    ctypes.memmove(ptr, blob, len(blob))
    _k32.GlobalUnlock(h)
    return h


def _open_dialogs() -> set[int]:
    """현재 떠 있는 카카오톡 창(파일 전송 확인 창 등) hwnd 집합."""
    out: set[int] = set()

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == MAIN_CLASS:
            out.add(hwnd)
        return True

    win32gui.EnumWindows(cb, None)
    return out


def _main_hwnd() -> int:
    return win32gui.FindWindow(MAIN_CLASS, MAIN_TITLE)


def _focus(hwnd: int) -> None:
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # 이미 포그라운드면 에러(183) 가 날 수 있는데 무시해도 된다.
        pass
    time.sleep(0.4)


def ensure_kakao_ready() -> int:
    """선행 체크: 로그인된 카카오톡 메인 창이 떠 있어야 한다. 없으면 중단."""
    _set_dpi_aware()
    hwnd = _main_hwnd()
    if not hwnd:
        raise KakaoNotReady(
            "카카오톡 메인 창을 찾지 못했습니다.\n"
            "PC 카카오톡을 실행하고 로그인한 뒤 다시 시도하세요."
        )
    return hwnd


def open_room(room: str, timeout: float = 6.0) -> int:
    """Ctrl+F 검색으로 방을 열고, 방 이름이 제목인 채팅창 hwnd 를 돌려준다.

    같은 이름의 방이 여러 개면 검색 최상단(가장 최근 활성) 방이 열린다.
    """
    main = ensure_kakao_ready()
    _focus(main)

    send_keys("^f")          # 채팅 검색 진입
    time.sleep(0.6)
    _copy_text(room)         # 한글 안전하게 클립보드로 입력
    send_keys("^v")
    time.sleep(1.0)
    send_keys("{ENTER}")     # 최상단 결과 열기

    deadline = time.time() + timeout
    while time.time() < deadline:
        hwnd = win32gui.FindWindow(MAIN_CLASS, room)
        if hwnd:
            return hwnd
        time.sleep(0.3)

    raise KakaoNotReady(
        f"'{room}' 방 창을 열지 못했습니다. 방 이름이 PC 카카오톡 방 제목과 정확히 같은지 확인하세요."
    )


def send_file(room_hwnd: int, path: Path, timeout: float = 6.0) -> None:
    """방 창에 파일을 드롭(WM_DROPFILES)해 전송 확인 창을 띄우고 Enter 로 전송한다.

    PC 카카오톡은 Ctrl+V 로 파일 첨부가 되지 않아, 드래그드롭과 동일한
    WM_DROPFILES 메시지로 방 창에 파일을 떨어뜨린다.
    """
    _focus(room_hwnd)

    l, t, r, b = win32gui.GetClientRect(room_hwnd)
    before = _open_dialogs()
    win32gui.PostMessage(
        room_hwnd, _WM_DROPFILES, _make_dropfiles(Path(path), (r - l) // 2, (b - t) // 2), 0
    )

    deadline = time.time() + timeout
    dialog = None
    while time.time() < deadline:
        new = _open_dialogs() - before
        if new:
            dialog = next(iter(new))
            break
        time.sleep(0.2)

    if dialog is None:
        raise KakaoNotReady("파일 전송 확인 창이 뜨지 않았습니다. 다시 시도하세요.")

    win32gui.SetForegroundWindow(dialog)
    time.sleep(0.4)
    send_keys("{ENTER}")

    # 확인 창이 닫힐 때까지(=전송 시작) 대기. 다음 전송이 이전 확인 창을 잡는 레이스 방지.
    close_deadline = time.time() + timeout
    while time.time() < close_deadline:
        if not (win32gui.IsWindow(dialog) and win32gui.IsWindowVisible(dialog)):
            break
        time.sleep(0.2)
    time.sleep(0.6)


def send_text(room_hwnd: int, text: str) -> None:
    """채팅창에 텍스트 메시지를 붙여넣고 전송한다."""
    _focus(room_hwnd)
    _copy_text(text)
    send_keys("^v")
    time.sleep(0.3)
    send_keys("{ENTER}")
    time.sleep(0.3)
