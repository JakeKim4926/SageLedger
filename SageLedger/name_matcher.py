from __future__ import annotations

import re
from collections import defaultdict
from typing import Iterable

from .models import KakaoTransaction, Member

KOREAN_NAME_RE = re.compile(r"[가-힣]{2,5}")
MONTH_SUFFIX_RE = re.compile(r"[_\s-]?(\d{1,2}월|\d{1,2}개월|회비|모임비|회칙)")


def extract_depositor_name(description: str) -> str:
    """카카오뱅크 '내용' 에서 입금자명을 추출한다.

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


def match_member_deposits(
    transactions: Iterable[KakaoTransaction],
    members: Iterable[Member],
) -> tuple[list[tuple[Member, KakaoTransaction]], list[tuple[KakaoTransaction, str]]]:
    """일반입금 거래를 회원과 매칭한다.

    반환: (매칭된 [(member, tx)], 매칭 실패 [(tx, 사유)])
    이자/출금 등 회비가 아닌 거래는 여기서 다루지 않는다.
    """
    index = build_member_index(members)
    matched: list[tuple[Member, KakaoTransaction]] = []
    unmatched: list[tuple[KakaoTransaction, str]] = []

    for tx in transactions:
        if tx.direction != "입금" or tx.transaction_type != "일반입금":
            continue

        name = extract_depositor_name(tx.description)
        candidates = index.get(name, [])

        if not candidates:
            unmatched.append((tx, f"회원명단에서 '{name}' 를 찾지 못함 (내용: {tx.description})"))
        elif len(candidates) > 1:
            unmatched.append((tx, f"'{name}' 동명/별칭 후보가 여러 명"))
        else:
            matched.append((candidates[0], tx))

    return matched, unmatched
