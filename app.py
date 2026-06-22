from __future__ import annotations

import argparse
from getpass import getpass
from pathlib import Path

from fit_ledger_automation.config_loader import load_yaml
from fit_ledger_automation.kakao_parser import parse_kakaobank_transactions
from fit_ledger_automation.ledger_writer import (
    load_ledger_workbook,
    save_workbook,
    write_personal_deposits,
    write_transaction_ledger,
)
from fit_ledger_automation.member_loader import load_members, members_from_ledger_personal_sheet
from fit_ledger_automation.models import LedgerWriteResult
from fit_ledger_automation.name_matcher import match_deposits
from fit_ledger_automation.report_writer import write_review_sheet
from fit_ledger_automation.validator import get_last_bank_balance, get_last_ledger_balance, validate_balances


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="F.I.T 카카오뱅크 장부 자동화")
    parser.add_argument("--kakao", required=True, help="카카오뱅크 거래내역 xlsx 경로")
    parser.add_argument("--ledger", required=True, help="기존 장부 xlsm/xlsx 경로")
    parser.add_argument("--members", default="config/members_template.csv", help="회원명단 csv/xlsx 경로")
    parser.add_argument("--mapping", default="config/sheet_mapping.yml", help="장부 시트 매핑 yml 경로")
    parser.add_argument("--category-rules", default="config/category_rules.yml", help="지출 분류 규칙 yml 경로")
    parser.add_argument("--output", required=True, help="결과 장부 저장 경로")
    parser.add_argument("--kakao-password", default=None, help="카카오뱅크 거래내역 비밀번호. 생략 시 필요할 때 입력")
    parser.add_argument("--ledger-password", default=None, help="장부 비밀번호. 생략 시 필요할 때 입력")
    parser.add_argument("--auto-write-review", action="store_true", help="검토필요 입금도 장부에 반영. 기본값은 미반영")
    parser.add_argument("--members-from-ledger", action="store_true", help="회원명단 파일 대신 장부 개인 입금 내역 헤더에서 회원명 추출")
    return parser


def main() -> None:
    args = build_parser().parse_args()

    kakao_password = args.kakao_password
    ledger_password = args.ledger_password

    # 비밀번호를 CLI 인자로 남기기 싫으면 입력 프롬프트 사용.
    if kakao_password is None:
        kakao_password = getpass("카카오뱅크 파일 비밀번호, 없으면 Enter: ") or None
    if ledger_password is None:
        ledger_password = getpass("장부 파일 비밀번호, 없으면 Enter: ") or None

    mapping = load_yaml(args.mapping)
    category_rules = load_yaml(args.category_rules)

    transactions = parse_kakaobank_transactions(args.kakao, password=kakao_password)
    wb = load_ledger_workbook(args.ledger, password=ledger_password)

    if args.members_from_ledger:
        members = members_from_ledger_personal_sheet(
            args.ledger,
            personal_sheet_name=mapping["personal_deposit_sheet"]["name"],
            member_header_row=int(mapping["personal_deposit_sheet"]["member_header_row"]),
        )
    else:
        members = load_members(args.members)

    matches = match_deposits(
        transactions,
        members,
        auto_confirm_review_amounts=args.auto_write_review,
    )

    result = LedgerWriteResult()
    result.final_bank_balance = get_last_bank_balance(transactions)
    result.written_personal_rows, written_tx_rows = write_personal_deposits(
        wb,
        matches,
        mapping,
        write_review_required=args.auto_write_review,
    )
    result.written_ledger_rows = write_transaction_ledger(
        wb,
        transactions,
        matches,
        mapping,
        category_rules,
        written_transaction_rows=written_tx_rows,
        write_review_required=args.auto_write_review,
    )
    result.final_ledger_balance = get_last_ledger_balance(wb, mapping)
    validate_balances(result)
    result.review_rows = write_review_sheet(wb, matches, result)

    output_path = save_workbook(wb, args.output)

    print("\n장부 자동화 완료")
    print(f"- 카카오뱅크 거래 건수: {len(transactions)}")
    print(f"- 회원 수: {len(members)}")
    print(f"- 개인 입금 내역 반영: {result.written_personal_rows}건")
    print(f"- 입출금 내역 반영: {result.written_ledger_rows}건")
    print(f"- 검토 필요: {result.review_rows}건")
    print(f"- 카카오뱅크 최종 잔액: {result.final_bank_balance:,}원" if result.final_bank_balance is not None else "- 카카오뱅크 최종 잔액: 확인불가")
    print(f"- 장부 최종 잔액: {result.final_ledger_balance:,}원" if result.final_ledger_balance is not None else "- 장부 최종 잔액: 확인불가")
    for message in result.validation_messages:
        print(f"- {message}")
    print(f"- 저장 위치: {output_path}")


if __name__ == "__main__":
    main()
