from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from SageLedger.config import load_group
from SageLedger.config_loader import load_yaml
from SageLedger.excel_writer import apply_plan
from SageLedger.kakao_parser import parse_kakaobank_transactions
from SageLedger.planner import build_plan
from SageLedger.validator import balance_message


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="모임 장부 자동화 (카카오뱅크 거래내역 -> 장부)")
    p.add_argument("group", help="모임 이름 (config/groups.yml 의 키, 예: fit, scm)")
    p.add_argument("--groups", default="config/groups.yml", help="모임 설정 yml 경로")
    p.add_argument("--env", default=".env", help=".env 경로")
    p.add_argument("--mapping", default="config/sheet_mapping.yml", help="시트 매핑 yml 경로")
    p.add_argument("--category-rules", default="config/category_rules.yml", help="지출 분류 규칙 yml")
    p.add_argument("--output", default=None, help="결과 장부 저장 경로 (기본: output/<group>_장부_자동작성_<시각><확장자>)")
    return p


def generate_ledger(
    group_name: str,
    groups_path: str = "config/groups.yml",
    env_path: str = ".env",
    mapping_path: str = "config/sheet_mapping.yml",
    category_rules_path: str = "config/category_rules.yml",
    output: str | None = None,
) -> dict:
    """장부를 자동 작성해 output 파일을 만들고 결과(dict) 를 돌려준다."""
    group = load_group(group_name, groups_path=groups_path, env_path=env_path)
    mapping = load_yaml(mapping_path)
    category_rules = load_yaml(category_rules_path)

    transactions = parse_kakaobank_transactions(group.kakao_path, password=group.kakao_pw)

    plan = build_plan(group, transactions, mapping, category_rules)

    if output:
        output_path = Path(output)
    else:
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = Path("output") / f"{group_name}_장부_자동작성_{stamp}{group.ledger_path.suffix}"

    result = apply_plan(group, plan, output_path)

    msg = balance_message(result["bank_balance"], result["ledger_balance"])

    print("\n장부 자동화 완료")
    print(f"- 모임: {group.name}  (회원 {len(group.members)}명)")
    print(f"- 카카오뱅크 거래 건수: {len(transactions)}")
    print(f"- 개인 입금 내역 반영: {plan.personal_written}칸")
    print(f"- 입출금 내역 반영: {plan.transaction_written}행")
    print(f"- 검토 필요: {len(plan.reviews)}건  (시트 '자동화_검토결과' 확인)")
    for note in plan.notes:
        print(f"- 참고: {note}")
    print(f"- {msg}")
    print(f"- 저장 위치: {result['output']}")
    return result


def main() -> None:
    args = build_parser().parse_args()
    generate_ledger(
        args.group,
        groups_path=args.groups,
        env_path=args.env,
        mapping_path=args.mapping,
        category_rules_path=args.category_rules,
        output=args.output,
    )


if __name__ == "__main__":
    main()
