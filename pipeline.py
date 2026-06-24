from __future__ import annotations

import argparse

from app import generate_ledger
from kakao_send import send_group
from SageLedger.config import load_room


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="장부 자동 작성 후 개인입금 캡처와 함께 카카오톡 단톡방으로 전송"
    )
    p.add_argument("group", help="모임 이름 (config/groups.yml 키, 예: fit, scm)")
    p.add_argument("--groups", default="config/groups.yml", help="모임 설정 yml 경로")
    p.add_argument(
        "--no-capture",
        action="store_true",
        help="개인입금 캡처 이미지 없이 엑셀만 전송",
    )
    return p


def main() -> None:
    args = build_parser().parse_args()

    # 1) 장부 자동 작성 (Excel)
    result = generate_ledger(args.group, groups_path=args.groups)

    # 2) 카카오톡 전송 (헤더 → 캡처 이미지 → 엑셀 → 소요시간)
    room = load_room(args.group, args.groups)
    elapsed = send_group(
        args.group, room, result["output"], not args.no_capture, args.groups
    )

    print(f"\n카카오톡 전송 완료: {args.group} -> '{room}'")
    print(f"- 카톡 전송 소요시간: {elapsed:.1f}초")


if __name__ == "__main__":
    main()
