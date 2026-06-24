from __future__ import annotations

import argparse
import sys
from pathlib import Path

from SageLedger import kakao_sender
from SageLedger.config import load_room


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="장부 엑셀/메시지를 카카오톡 단톡방으로 전송")
    p.add_argument("group", help="모임 이름 (config/groups.yml 키, 예: fit, scm)")
    p.add_argument("--file", default=None, help="보낼 파일 (기본: output 의 최신 <group> 장부)")
    p.add_argument("--message", default=None, help="함께 보낼 텍스트 메시지")
    p.add_argument("--groups", default="config/groups.yml", help="모임 설정 yml 경로")
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="방 찾기까지만 하고 전송하지 않음 (안전 테스트)",
    )
    return p


def latest_output(group: str) -> Path | None:
    files = sorted(
        Path("output").glob(f"{group}_장부_자동작성_*"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return files[0] if files else None


def main() -> None:
    args = build_parser().parse_args()

    room = load_room(args.group, args.groups)

    file_path = Path(args.file) if args.file else latest_output(args.group)
    if file_path is None:
        sys.exit(
            f"보낼 파일이 없습니다. output/{args.group}_장부_자동작성_* 도 없고 --file 도 없습니다."
        )
    if not file_path.exists():
        sys.exit(f"파일이 존재하지 않습니다: {file_path}")

    # 1) 선행 체크: 카카오톡 로그인/창 확인
    try:
        kakao_sender.ensure_kakao_ready()
    except kakao_sender.KakaoNotReady as e:
        sys.exit(str(e))

    print(f"- 모임: {args.group}")
    print(f"- 단톡방: {room}")
    print(f"- 파일: {file_path}")
    if args.message:
        print(f"- 메시지: {args.message}")

    if args.dry_run:
        kakao_sender.open_room(room)
        print(f"\n[dry-run] '{room}' 방을 찾아 열었습니다. 전송은 하지 않았습니다.")
        return

    # 2) 전송 직전 확인 1회
    ans = input(f"\n'{room}' 방에 위 내용을 전송할까요? [y/N] ").strip().lower()
    if ans != "y":
        print("취소했습니다.")
        return

    room_win = kakao_sender.open_room(room)
    kakao_sender.send_file(room_win, file_path)
    if args.message:
        kakao_sender.send_text(room_win, args.message)
    print("전송 완료")


if __name__ == "__main__":
    main()
