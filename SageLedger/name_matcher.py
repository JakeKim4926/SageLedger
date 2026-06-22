from __future__ import annotations

import re
from collections import defaultdict
from typing import Iterable

from .models import KakaoTransaction, MatchResult, MatchStatus, Member

KOREAN_NAME_RE = re.compile(r"[가-힣]{2,5}")
MONTH_SUFFIX_RE = re.compile(r"[_\s-]?(\d{1,2}월|\d{1,2}개월|회비|모임비)")


def extract_depositor_name(description: str) -> str:
    """Extract a likely depositor name from KakaoBank description text.

    Examples:
    - 홍길동_5월 -> 홍길동
    - F.I.T 홍길동 회비 -> 홍길동
    """
    text = (description or "").strip()
    if not text:
        return ""

    first = re.split(r"[_/|,]", text)[0].strip()
    first = MONTH_SUFFIX_RE.sub("", first).strip()
    if KOREAN_NAME_RE.fullmatch(first):
        return first

    matches = KOREAN_NAME_RE.findall(text)
    return matches[0] if matches else text


def build_member_index(members: Iterable[Member]) -> dict[str, list[Member]]:
    index: dict[str, list[Member]] = defaultdict(list)
    for member in members:
        for name in member.match_names():
            index[name].append(member)
    return dict(index)


def match_deposits(
    transactions: Iterable[KakaoTransaction],
    members: Iterable[Member],
    auto_confirm_review_amounts: bool = False,
) -> list[MatchResult]:
    """Match KakaoBank income transactions to members.

    Only 일반입금 rows are treated as member deposit candidates.
    Interest and expenses are not returned here.
    """
    member_index = build_member_index(members)
    results: list[MatchResult] = []
    seen_member_month: set[tuple[str, str]] = set()

    for tx in transactions:
        if tx.direction != "입금" or tx.transaction_type != "일반입금":
            continue

        extracted = extract_depositor_name(tx.description)
        candidates = member_index.get(extracted, [])

        if not candidates:
            results.append(
                MatchResult(
                    transaction=tx,
                    extracted_name=extracted,
                    member=None,
                    status=MatchStatus.UNKNOWN_MEMBER,
                    reason="회원명단에서 입금자명을 찾지 못했습니다.",
                    auto_confirmed=False,
                )
            )
            continue

        if len(candidates) > 1:
            results.append(
                MatchResult(
                    transaction=tx,
                    extracted_name=extracted,
                    member=None,
                    status=MatchStatus.REVIEW_REQUIRED,
                    reason="동일 별칭/이름 후보가 여러 명입니다.",
                    auto_confirmed=False,
                )
            )
            continue

        member = candidates[0]
        month_key = tx.year_month_key
        duplicate_key = (member.name, month_key)

        if duplicate_key in seen_member_month:
            results.append(
                MatchResult(
                    transaction=tx,
                    extracted_name=extracted,
                    member=member,
                    status=MatchStatus.DUPLICATE,
                    reason="같은 회원의 같은 달 입금이 이미 존재합니다.",
                    auto_confirmed=False,
                )
            )
            continue

        seen_member_month.add(duplicate_key)

        if tx.amount == member.monthly_due:
            results.append(
                MatchResult(
                    transaction=tx,
                    extracted_name=extracted,
                    member=member,
                    status=MatchStatus.NORMAL,
                    reason="회원명과 기준 회비가 일치합니다.",
                    auto_confirmed=True,
                )
            )
        elif 0 < tx.amount < member.monthly_due:
            results.append(
                MatchResult(
                    transaction=tx,
                    extracted_name=extracted,
                    member=member,
                    status=MatchStatus.SHORT_PAYMENT,
                    reason=f"기준 회비 {member.monthly_due:,}원보다 적게 입금되었습니다.",
                    auto_confirmed=False,
                )
            )
        else:
            results.append(
                MatchResult(
                    transaction=tx,
                    extracted_name=extracted,
                    member=member,
                    status=MatchStatus.REVIEW_REQUIRED,
                    reason=f"기준 회비 {member.monthly_due:,}원과 입금액 {tx.amount:,}원이 다릅니다.",
                    auto_confirmed=auto_confirm_review_amounts,
                )
            )

    return results
