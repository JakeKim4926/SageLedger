from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from SageLedger import kakao_sender
from SageLedger.config import list_groups, load_pw, load_room
from SageLedger.sheet_capture import capture_deposit_sheet

EXCEL_EXT = (".xlsx", ".xlsm", ".xls")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="장부 엑셀/메시지를 카카오톡 단톡방으로 전송")
    p.add_argument("group", nargs="?", help="모임 이름 (config/groups.yml 키, 예: fit, scm)")
    p.add_argument("--file", default=None, help="보낼 파일 (기본: output 의 최신 <group> 장부)")
    p.add_argument("--message", default=None, help="함께 보낼 텍스트 메시지")
    p.add_argument("--groups", default="config/groups.yml", help="모임 설정 yml 경로")
    p.add_argument(
        "--all",
        action="store_true",
        help="output 의 엑셀을 파일명에 든 모임키로 라우팅해 각 단톡방으로 전송",
    )
    p.add_argument(
        "--no-capture",
        action="store_true",
        help="--all 시 개인입금 캡처 이미지 없이 엑셀만 전송",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="방 찾기/열기까지만 하고 전송하지 않음 (안전 테스트)",
    )
    return p


def latest_output(group: str) -> Path | None:
    files = sorted(
        Path("output").glob(f"{group}_장부_자동작성_*"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return files[0] if files else None


def latest_excel_for(group: str) -> Path | None:
    """output 에서 파일명에 모임키(대소문자 무시)가 든 최신 엑셀 1개."""
    files = [
        p
        for p in Path("output").glob("*")
        if p.is_file()
        and p.suffix.lower() in EXCEL_EXT
        and group.lower() in p.name.lower()
    ]
    return max(files, key=lambda p: p.stat().st_mtime, default=None)


def elapsed_message(seconds: float) -> str:
    """마지막에 보낼 카카오톡 전송 소요시간 문구."""
    if seconds >= 60:
        m, s = divmod(int(round(seconds)), 60)
        body = f"{m}분 {s}초"
    else:
        body = f"{seconds:.1f}초"
    return f"SageLedger 자동화 소요시간: {body}"


def send_group(
    group: str,
    room: str,
    file_path: Path,
    capture: bool,
    groups_path: str,
) -> float:
    """한 모임의 단톡방으로 (캡처 이미지) → 엑셀 → 소요시간 순으로 보낸다.

    캡처(Excel)는 카카오 포커스 다툼을 피하려 전송 시작 전에 끝낸다.
    돌려주는 값은 카카오톡 전송에 걸린 시간(초).
    """
    image = None
    if capture:
        read_pw, write_pw = load_pw(group, groups_path)
        image = capture_deposit_sheet(file_path, read_pw, write_pw)

    start = time.time()
    room_hwnd = kakao_sender.open_room(room)
    if image is not None:
        kakao_sender.send_file(room_hwnd, image)
    kakao_sender.send_file(room_hwnd, file_path)
    elapsed = time.time() - start

    kakao_sender.send_text(room_hwnd, elapsed_message(elapsed))
    return elapsed


def run_all(groups_path: str, capture: bool = True) -> None:
    """output 의 엑셀을 모임키로 라우팅해 각 단톡방으로 전송한다.

    그룹별로 헤더 메시지 → 개인입금 캡처 이미지 → 장부 엑셀 순으로 보낸다.
    """
    plan: list[tuple[str, str, Path]] = []
    for group in list_groups(groups_path):
        file_path = latest_excel_for(group)
        if file_path is None:
            continue
        plan.append((group, load_room(group, groups_path), file_path))

    if not plan:
        sys.exit("output 에서 보낼 엑셀을 찾지 못했습니다.")

    print("전송 계획:" + ("  (메시지 + 개인입금 캡처 + 엑셀)" if capture else "  (메시지 + 엑셀)"))
    for group, room, file_path in plan:
        print(f"  - {group}: {file_path.name}  ->  '{room}'")

    ans = input("\n위 내용을 각 단톡방으로 전송할까요? [y/N] ").strip().lower()
    if ans != "y":
        print("취소했습니다.")
        return

    try:
        kakao_sender.ensure_kakao_ready()
    except kakao_sender.KakaoNotReady as e:
        sys.exit(str(e))

    for group, room, file_path in plan:
        elapsed = send_group(group, room, file_path, capture, groups_path)
        print(f"전송 완료: {group} -> {room}  (카톡 전송 {elapsed:.1f}초)")


def main() -> None:
    args = build_parser().parse_args()

    if args.all:
        run_all(args.groups, capture=not args.no_capture)
        return

    if not args.group:
        sys.exit("모임 이름(group) 또는 --all 중 하나가 필요합니다.")

    room = load_room(args.group, args.groups)

    # 선행 체크: 카카오톡 로그인/창 확인
    try:
        kakao_sender.ensure_kakao_ready()
    except kakao_sender.KakaoNotReady as e:
        sys.exit(str(e))

    # dry-run 은 방 열기만 검증 (보낼 파일 불필요)
    if args.dry_run:
        print(f"- 모임: {args.group}")
        print(f"- 단톡방: {room}")
        kakao_sender.open_room(room)
        print(f"\n[dry-run] '{room}' 방을 찾아 열었습니다. 전송은 하지 않았습니다.")
        return

    file_path = Path(args.file) if args.file else latest_output(args.group)
    if file_path is None:
        sys.exit(
            f"보낼 파일이 없습니다. output/{args.group}_장부_자동작성_* 도 없고 --file 도 없습니다."
        )
    if not file_path.exists():
        sys.exit(f"파일이 존재하지 않습니다: {file_path}")

    print(f"- 모임: {args.group}")
    print(f"- 단톡방: {room}")
    print(f"- 파일: {file_path}")
    if args.message:
        print(f"- 메시지: {args.message}")

    # 전송 직전 확인 1회
    ans = input(f"\n'{room}' 방에 위 내용을 전송할까요? [y/N] ").strip().lower()
    if ans != "y":
        print("취소했습니다.")
        return

    room_hwnd = kakao_sender.open_room(room)
    kakao_sender.send_file(room_hwnd, file_path)
    if args.message:
        kakao_sender.send_text(room_hwnd, args.message)
    print("전송 완료")


if __name__ == "__main__":
    main()
