from __future__ import annotations

import argparse
import sys

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


def verify_balance(ledger_balance: int | None) -> None:
    """자동 입력 결과 잔액과 사용자가 확인한 실제 계좌 잔액이 일치하는지 검증.

    불일치/입력 오류/잔액 미확인 시 errmsg 출력 후 종료(전송 단계로 넘어가지 않음).
    """
    if ledger_balance is None:
        print("[중단] 장부 자동입력 잔액을 확인하지 못했습니다. 카카오톡 전송을 진행하지 않습니다.")
        sys.exit(1)

    raw = input(f"\n현재 실제 계좌 잔액을 입력하세요 (자동입력 결과: {ledger_balance:,}원): ")
    try:
        actual = int(raw.strip().replace(",", "").replace("원", ""))
    except ValueError:
        print("[중단] 잔액은 숫자로 입력하세요. 카카오톡 전송을 진행하지 않습니다.")
        sys.exit(1)

    if actual != ledger_balance:
        diff = ledger_balance - actual
        print(f"[중단] 잔액 불일치: 입력 {actual:,}원 ≠ 자동입력 {ledger_balance:,}원 "
              f"({diff:+,}원 차이). 카카오톡 전송을 진행하지 않습니다.")
        sys.exit(1)

    print(f"잔액 확인 정상: {actual:,}원. 전송을 진행합니다.")


def main() -> None:
    args = build_parser().parse_args()

    # 1) 장부 자동 작성 (Excel)
    result = generate_ledger(args.group, groups_path=args.groups)

    # 2) 현재 계좌 잔액 확인 게이트 (불일치 시 전송하지 않고 종료)
    verify_balance(result["ledger_balance"])

    # 3) 카카오톡 전송 (헤더 → 캡처 이미지 → 엑셀 → 소요시간)
    room = load_room(args.group, args.groups)
    elapsed = send_group(
        args.group, room, result["output"], not args.no_capture, args.groups
    )

    print(f"\n✅ {args.group.upper()} 장부 전송 완료 "
          f"(잔액 {result['ledger_balance']:,}원) · {elapsed:.1f}초")


if __name__ == "__main__":
    main()
