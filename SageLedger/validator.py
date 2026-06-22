from __future__ import annotations

from typing import Optional


def balance_message(bank_balance: Optional[int], ledger_balance: Optional[int]) -> str:
    if bank_balance is None:
        return "카카오뱅크 최종 잔액을 확인하지 못했습니다."
    if ledger_balance is None:
        return "장부 최종 잔액을 확인하지 못했습니다."
    if bank_balance == ledger_balance:
        return f"잔액 검증 정상: 카카오뱅크와 장부 모두 {bank_balance:,}원."
    diff = ledger_balance - bank_balance
    return (f"잔액 불일치: 장부 {ledger_balance:,}원 - 카카오뱅크 {bank_balance:,}원 "
            f"= {diff:,}원 차이. 중복/누락 또는 수기 회비정산과의 중복을 확인하세요.")
