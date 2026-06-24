from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

from .config_loader import load_yaml

DEFAULT_DEPOSIT_SHEET = "개인 입금 내역"


def _deposit_sheet_name(mapping_path: str | Path = "config/sheet_mapping.yml") -> str:
    data = load_yaml(mapping_path)
    return (data.get("personal_deposit_sheet") or {}).get("name", DEFAULT_DEPOSIT_SHEET)


def capture_deposit_sheet(
    ledger_path: str | Path,
    read_pw: Optional[str] = None,
    write_pw: Optional[str] = None,
    mapping_path: str | Path = "config/sheet_mapping.yml",
) -> Path:
    """장부의 '개인 입금 내역' 시트를 PNG 로 캡처해 임시 파일 경로를 돌려준다.

    xlwings(Excel) 로 렌더링한다. visible=False 면 to_png 가 빈 이미지가 되므로
    창을 보이게 열고(끝나면 자동 종료) 시트의 used_range 를 그대로 내보낸다.
    """
    import xlwings as xw

    ledger = Path(ledger_path)
    out = Path(tempfile.gettempdir()) / f"{ledger.stem}_개인입금내역.png"
    want = _deposit_sheet_name(mapping_path)
    key = want.replace(" ", "")

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    try:
        book = app.books.open(
            str(ledger.resolve()),
            password=read_pw,
            write_res_password=write_pw or read_pw,
            ignore_read_only_recommended=True,
            update_links=False,
        )
        sheet = next((s for s in book.sheets if s.name.replace(" ", "") == key), None)
        if sheet is None:
            raise KeyError(f"'{want}' 시트를 찾지 못했습니다: {ledger.name}")
        sheet.used_range.to_png(str(out))
        book.close()
    finally:
        app.quit()
    return out
