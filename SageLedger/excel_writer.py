from __future__ import annotations

import shutil
from datetime import datetime
from numbers import Number
from pathlib import Path
from typing import Optional

from .config import GroupConfig
from .models import WritePlan

REVIEW_SHEET = "자동화_검토결과"


def apply_plan(group: GroupConfig, plan: WritePlan, output_path: str | Path) -> dict:
    """xlwings 로 원본 장부 사본을 열어 계획을 적용하고, 수식 재계산 후 저장한다.

    원본을 그대로 복사한 뒤 같은 파일에 저장하므로 읽기/쓰기 비밀번호가 보존된다.
    반환: {'output': Path, 'ledger_balance': int|None, 'bank_balance': int|None}
    """
    import xlwings as xw

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(group.ledger_path, output)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    ledger_balance: Optional[int] = None
    try:
        book = app.books.open(
            str(output.resolve()),
            password=group.read_pw,
            write_res_password=group.write_pw or group.read_pw,
            ignore_read_only_recommended=True,
            update_links=False,
        )

        for cw in plan.cells:
            cell = book.sheets[cw.sheet].range((cw.row, cw.col))
            if cw.formula:
                cell.formula = cw.value
            else:
                cell.value = cw.value

        _write_review_sheet(xw, book, plan)

        book.app.calculate()

        if plan.last_transaction_row and plan.balance_col:
            val = book.sheets[_tx_sheet_name(book, plan)].range(
                (plan.last_transaction_row, plan.balance_col)).value
            ledger_balance = (int(round(val))
                              if isinstance(val, Number) and not isinstance(val, bool)
                              else None)

        book.save()
        book.close()
    finally:
        app.quit()

    return {"output": output, "ledger_balance": ledger_balance,
            "bank_balance": plan.bank_final_balance}


def _tx_sheet_name(book, plan: WritePlan) -> str:
    # 입출금 셀 쓰기는 모두 같은 시트명을 쓰므로 첫 입출금 셀에서 시트명을 가져온다.
    for cw in plan.cells:
        if cw.col == plan.balance_col and cw.formula:
            return cw.sheet
    # 안전망: 회비/지출이 하나도 없으면 잔액은 확인하지 않는다.
    return book.sheets[0].name


def _write_review_sheet(xw, book, plan: WritePlan) -> None:
    if REVIEW_SHEET in [s.name for s in book.sheets]:
        book.sheets[REVIEW_SHEET].delete()
    ws = book.sheets.add(REVIEW_SHEET, after=book.sheets[len(book.sheets) - 1])

    ws.range("A1").value = "장부 자동화 검토 결과"
    ws.range("A1").font.bold = True
    ws.range("A2").value = f"생성시각: {datetime.now():%Y-%m-%d %H:%M:%S}"

    headers = ["거래일시", "회원/구분", "금액", "사유"]
    ws.range("A4").value = headers
    ws.range("A4:D4").font.bold = True

    r = 5
    for item in plan.reviews:
        ws.range((r, 1)).value = item.when.strftime("%Y-%m-%d %H:%M") if item.when else ""
        ws.range((r, 2)).value = item.name
        ws.range((r, 3)).value = item.amount
        ws.range((r, 4)).value = item.reason
        r += 1

    r += 1
    ws.range((r, 1)).value = "처리 요약"
    ws.range((r, 1)).font.bold = True
    for label, value in [
        ("개인 입금 내역 반영(칸)", plan.personal_written),
        ("입출금 내역 반영(행)", plan.transaction_written),
        ("검토 필요", len(plan.reviews)),
    ]:
        r += 1
        ws.range((r, 1)).value = label
        ws.range((r, 2)).value = value

    if plan.notes:
        r += 2
        ws.range((r, 1)).value = "참고"
        ws.range((r, 1)).font.bold = True
        for note in plan.notes:
            r += 1
            ws.range((r, 1)).value = note

    ws.range("A:A").column_width = 18
    ws.range("D:D").column_width = 50
