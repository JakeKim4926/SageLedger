from __future__ import annotations

from typing import Optional

from .models import KakaoTransaction, LedgerWriteResult


def get_last_bank_balance(transactions: list[KakaoTransaction]) -> Optional[int]:
    if not transactions:
        return None
    return sorted(transactions, key=lambda tx: tx.traded_at)[-1].balance


def get_last_ledger_balance(wb, mapping: dict) -> Optional[int]:
    cfg = mapping["transaction_sheet"]
    ws = wb[cfg["name"]]
    balance_col = int(cfg["columns"]["balance"])
    date_col = int(cfg["columns"]["date"])
    amount_col = int(cfg["columns"]["amount"])
    for row_idx in range(ws.max_row, int(cfg["start_row"]) - 1, -1):
        if ws.cell(row_idx, date_col).value is not None and ws.cell(row_idx, amount_col).value is not None:
            value = ws.cell(row_idx, balance_col).value
            return int(value) if isinstance(value, (int, float)) else None
    return None


def validate_balances(result: LedgerWriteResult) -> None:
    if result.final_bank_balance is None:
        result.validation_messages.append("카카오뱅크 최종 잔액을 확인하지 못했습니다.")
        return
    if result.final_ledger_balance is None:
        result.validation_messages.append("장부 최종 잔액을 확인하지 못했습니다.")
        return
    if result.final_bank_balance == result.final_ledger_balance:
        result.validation_messages.append("잔액 검증 정상: 카카오뱅크 최종 잔액과 장부 최종 잔액이 일치합니다.")
    else:
        diff = result.final_ledger_balance - result.final_bank_balance
        result.validation_messages.append(
            f"잔액 불일치: 장부 잔액과 카카오뱅크 잔액 차이 {diff:,}원. 중복 반영 또는 누락을 확인하세요."
        )
