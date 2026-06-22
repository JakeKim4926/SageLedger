from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook


def _can_open_with_openpyxl(path: Path) -> bool:
    try:
        wb = load_workbook(path, read_only=True, keep_vba=path.suffix.lower() == ".xlsm", data_only=True)
        wb.close()
        return True
    except Exception:
        return False


def ensure_readable_excel(path: str | Path, password: Optional[str] = None, suffix: Optional[str] = None) -> Path:
    """Return a readable Excel path.

    If the workbook is already readable, returns the original path.
    If it is password-encrypted, decrypts it to a temporary file using msoffcrypto-tool.
    """
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {src}")

    if _can_open_with_openpyxl(src):
        return src

    if not password:
        raise ValueError(f"암호화된 엑셀로 보입니다. 비밀번호가 필요합니다: {src.name}")

    try:
        import msoffcrypto
    except ImportError as exc:
        raise ImportError(
            "암호화 엑셀을 열려면 msoffcrypto-tool이 필요합니다. "
            "pip install msoffcrypto-tool 로 설치하세요."
        ) from exc

    suffix = suffix or src.suffix or ".xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_path = Path(tmp.name)
    tmp.close()

    with src.open("rb") as fp:
        office_file = msoffcrypto.OfficeFile(fp)
        office_file.load_key(password=password)
        with tmp_path.open("wb") as out:
            office_file.decrypt(out)

    if not _can_open_with_openpyxl(tmp_path):
        raise ValueError(f"복호화는 되었지만 openpyxl로 열 수 없습니다: {src.name}")

    return tmp_path
