from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Optional


class MatchStatus(str, Enum):
    NORMAL = "정상"
    REVIEW_REQUIRED = "검토필요"
    UNKNOWN_MEMBER = "회원미확인"
    DUPLICATE = "중복의심"
    SHORT_PAYMENT = "부족납부"
    NO_MEMBER_FILE = "회원명단없음"


@dataclass(frozen=True)
class Member:
    name: str
    monthly_due: int = 20_000
    status: str = "취직"
    aliases: tuple[str, ...] = field(default_factory=tuple)
    active: bool = True

    def match_names(self) -> set[str]:
        return {self.name, *self.aliases}


@dataclass(frozen=True)
class KakaoTransaction:
    row_no: int
    traded_at: datetime
    direction: str  # 입금 / 출금
    amount: int     # 입금 양수, 출금 음수
    balance: int
    transaction_type: str
    description: str
    memo: str = ""

    @property
    def income(self) -> int:
        return self.amount if self.amount > 0 else 0

    @property
    def expense(self) -> int:
        return abs(self.amount) if self.amount < 0 else 0

    @property
    def year_month_key(self) -> str:
        return f"{self.traded_at.year:04d}-{self.traded_at.month:02d}"


@dataclass
class MatchResult:
    transaction: KakaoTransaction
    extracted_name: str
    member: Optional[Member]
    status: MatchStatus
    reason: str
    auto_confirmed: bool = False


@dataclass
class LedgerWriteResult:
    written_personal_rows: int = 0
    written_ledger_rows: int = 0
    skipped_rows: int = 0
    review_rows: int = 0
    final_bank_balance: Optional[int] = None
    final_ledger_balance: Optional[int] = None
    validation_messages: list[str] = field(default_factory=list)
