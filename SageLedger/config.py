from __future__ import annotations

import glob
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

from .config_loader import load_yaml
from .models import Member

DEFAULT_DUES = {"취직": 20000, "미취직": 10000}


@dataclass
class GroupConfig:
    name: str
    ledger_path: Path
    kakao_path: Path
    read_pw: Optional[str]
    write_pw: Optional[str]
    kakao_pw: Optional[str]
    members: list[Member]


def _resolve_one(pattern: str, what: str) -> Path:
    matches = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)
    if not matches:
        raise FileNotFoundError(f"{what} 파일을 찾지 못했습니다: 패턴 {pattern!r}")
    return Path(matches[0])


def load_group(
    group_name: str,
    groups_path: str | Path = "config/groups.yml",
    env_path: str | Path = ".env",
) -> GroupConfig:
    if Path(env_path).exists():
        load_dotenv(env_path)

    data = load_yaml(groups_path)
    dues = {**DEFAULT_DUES, **(data.get("dues") or {})}
    groups = data.get("groups") or {}
    if group_name not in groups:
        raise KeyError(f"'{group_name}' 모임이 {groups_path} 에 없습니다. 사용 가능: {list(groups)}")

    g = groups[group_name]

    members: list[Member] = []
    for m in g["members"]:
        status = str(m.get("status", "취직")).strip()
        if status not in dues:
            raise ValueError(f"알 수 없는 status {status!r} (회원 {m.get('name')}). dues 에 정의 필요.")
        members.append(Member(name=str(m["name"]).strip(), monthly_due=int(dues[status])))

    return GroupConfig(
        name=group_name,
        ledger_path=_resolve_one(g["ledger_glob"], f"{group_name} 장부"),
        kakao_path=_resolve_one(g["kakao_glob"], f"{group_name} 카카오뱅크 거래내역"),
        read_pw=os.environ.get(g["read_pw_env"]) or None,
        write_pw=os.environ.get(g["write_pw_env"]) or None,
        kakao_pw=os.environ.get("KAKAO_PASSWORD") or None,
        members=members,
    )
