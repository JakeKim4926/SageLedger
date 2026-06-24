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


def list_groups(groups_path: str | Path = "config/groups.yml") -> list[str]:
    """설정에 정의된 모임 이름 목록."""
    data = load_yaml(groups_path)
    return list((data.get("groups") or {}).keys())


def load_pw(
    group_name: str,
    groups_path: str | Path = "config/groups.yml",
    env_path: str | Path = ".env",
) -> tuple[Optional[str], Optional[str]]:
    """모임의 (읽기, 쓰기) 비밀번호를 env 에서 읽는다 (입력 파일 불필요)."""
    if Path(env_path).exists():
        load_dotenv(env_path)
    data = load_yaml(groups_path)
    g = (data.get("groups") or {}).get(group_name)
    if not g:
        raise KeyError(f"'{group_name}' 모임이 {groups_path} 에 없습니다.")
    return (
        os.environ.get(g["read_pw_env"]) or None,
        os.environ.get(g["write_pw_env"]) or None,
    )


def load_room(group_name: str, groups_path: str | Path = "config/groups.yml") -> str:
    """모임의 카카오톡 단톡방 이름을 읽는다 (전송 전용, 입력 파일 없이도 동작)."""
    data = load_yaml(groups_path)
    groups = data.get("groups") or {}
    if group_name not in groups:
        raise KeyError(f"'{group_name}' 모임이 {groups_path} 에 없습니다. 사용 가능: {list(groups)}")
    room = (groups[group_name] or {}).get("kakao_room")
    if not room:
        raise KeyError(f"'{group_name}' 에 kakao_room 이 설정되지 않았습니다 ({groups_path}).")
    return str(room).strip()
