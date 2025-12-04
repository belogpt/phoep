"""Модели данных для контактов и групп телефонной книги."""
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Contact:
    """Контакт в телефонной книге."""
    group: str
    name: str
    office: str = ""
    mobile: str = ""
    other: str = ""
    photo: str = ""
    contact_id: Optional[int] = field(default=None)


@dataclass
class Group:
    """Группа (подразделение) телефонной книги."""
    name: str
    contact_count: int = 0
    order_index: int = 0
