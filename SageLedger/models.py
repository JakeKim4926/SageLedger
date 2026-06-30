from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Optional


@dataclass(frozen=True)
class Member:
    name: str
    monthly_due: int
    aliases: tuple[str, ...] = field(default_factory=tuple)

    def match_names(self) -> set[str]:
        return {self.name, *self.aliases}


@dataclass(frozen=True)
class KakaoTransaction:
    row_no: int
    traded_at: datetime
    direction: str       # 입금 / 출금
    amount: int          # 입금 양수, 출금 음수
    balance: int
    transaction_type: str  # 일반입금 / 예금이자 / 일반이체 ...
    description: str
    memo: str = ""

    @property
    def income(self) -> int:
        return self.amount if self.amount > 0 else 0

    @property
    def expense(self) -> int:
        return -self.amount if self.amount < 0 else 0

    @property
    def year_month_key(self) -> str:
        return f"{self.traded_at.year:04d}-{self.traded_at.month:02d}"


@dataclass(frozen=True)
class CellWrite:
    """한 셀에 쓸 값. formula=True 면 value 를 수식으로 기록한다."""
    sheet: str
    row: int
    col: int
    value: Any
    formula: bool = False


@dataclass(frozen=True)
class MergeRange:
    """입출금 새 행에 적용할 셀 병합 범위 (한 행의 col_start..col_end)."""
    sheet: str
    row: int
    col_start: int
    col_end: int


@dataclass
class ReviewItem:
    when: Optional[datetime]
    name: str
    amount: int
    reason: str


@dataclass
class WritePlan:
    cells: list[CellWrite] = field(default_factory=list)
    merges: list[MergeRange] = field(default_factory=list)
    reviews: list[ReviewItem] = field(default_factory=list)
    personal_written: int = 0
    transaction_written: int = 0
    notes: list[str] = field(default_factory=list)
    # 검증용: 입출금 내역의 마지막 데이터 행과 합계 열(계산은 xlwings 가 수행).
    last_transaction_row: Optional[int] = None
    balance_col: Optional[int] = None
    bank_final_balance: Optional[int] = None
