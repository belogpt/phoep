"""Импорт и экспорт телефонной книги в формате Excel."""
import io
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional
import pandas as pd
from openpyxl import load_workbook

from .models import Contact
from .repository import save_contacts

COLUMNS = [
    'Department',
    'Name',
    'Office Number',
    'Mobile Number',
    'Other Number',
    'Head Portrait',
]


def export_to_excel(contacts: List[Contact]) -> bytes:
    """Возвращает байты Excel-файла со всеми контактами."""
    data = []
    for contact in contacts:
        data.append({
            'Department': contact.group,
            'Name': contact.name,
            'Office Number': contact.office,
            'Mobile Number': contact.mobile,
            'Other Number': contact.other,
            'Head Portrait': '',
        })
    df = pd.DataFrame(data, columns=COLUMNS)
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer.read()


def import_from_excel(file_stream) -> int:
    """Импортирует контакты из Excel, полностью перезаписывая XML. Возвращает количество контактов."""
    df = pd.read_excel(file_stream, engine=None)
    # нормализуем имена колонок
    columns_map = {c.lower(): c for c in df.columns}
    required = {
        'department': 'Department',
        'name': 'Name',
        'office number': 'Office Number',
        'mobile number': 'Mobile Number',
        'other number': 'Other Number',
        'head portrait': 'Head Portrait',
    }
    for key in required:
        if key not in columns_map:
            raise ValueError('Неверный формат столбцов Excel')
    contacts: List[Contact] = []
    for _, row in df.iterrows():
        department = str(row[columns_map['department']]).strip() if not pd.isna(row[columns_map['department']]) else ''
        name = str(row[columns_map['name']]).strip() if not pd.isna(row[columns_map['name']]) else ''
        office = '' if pd.isna(row[columns_map['office number']]) else str(row[columns_map['office number']]).strip()
        mobile = '' if pd.isna(row[columns_map['mobile number']]) else str(row[columns_map['mobile number']]).strip()
        other = '' if pd.isna(row[columns_map['other number']]) else str(row[columns_map['other number']]).strip()
        if not department or (not name and not office and not mobile and not other):
            continue
        contacts.append(Contact(group=department, name=name, office=office, mobile=mobile, other=other, photo=''))
    save_contacts(contacts)
    return len(contacts)


@dataclass
class RawContact:
    """Контакт, полученный из «сырой» таблицы отделов."""

    full_department_name: str
    full_name: str
    raw_row_data: List[str] = field(default_factory=list)
    internal_extension: Optional[str] = None


def _is_department_row(row) -> bool:
    """Пытается определить, является ли строка строкой отдела по цвету/заполненности."""

    for cell in row:
        fill = getattr(cell, "fill", None)
        if fill and fill.fill_type and getattr(fill.start_color, "rgb", None):
            color = fill.start_color.rgb
            if color and color not in {"00000000", "FFFFFFFF", "FF000000"}:
                return True

    first_value = row[0].value if row else None
    other_values = [c.value for c in row[1:]] if len(row) > 1 else []
    if first_value and all(v in (None, "") for v in other_values):
        return True
    return False


def _extract_internal_extension(values: List[str]) -> Optional[str]:
    """Ищет первый короткий номер (3–5 цифр) в наборе значений строки."""

    for raw in values:
        if raw is None:
            continue
        text = str(raw)
        for match in re.findall(r"[\d\s\+\-\(\)]+", text):
            digits = re.sub(r"[\s\-\+\(\)]", "", match)
            if digits.isdigit() and 3 <= len(digits) <= 5:
                return digits
    return None


def parse_raw_department_table(filepath: str) -> List[RawContact]:
    """Парсит «сырую» Excel-таблицу с отделами и сотрудниками."""

    if not os.path.exists(filepath):
        raise FileNotFoundError(filepath)

    wb = load_workbook(filepath)
    ws = wb.active
    current_department: Optional[str] = None
    raw_contacts: List[RawContact] = []

    for row in ws.iter_rows():
        values = [cell.value for cell in row]
        if any(values):
            if _is_department_row(row):
                current_department = str(values[0]).strip()
                continue

            full_name = str(values[0]).strip() if values and values[0] else ""
            if not full_name or not current_department:
                continue

            normalized_values = ["" if v is None else str(v) for v in values]
            internal_extension = _extract_internal_extension(normalized_values)
            raw_contacts.append(
                RawContact(
                    full_department_name=current_department,
                    full_name=full_name,
                    raw_row_data=normalized_values,
                    internal_extension=internal_extension,
                )
            )

    return raw_contacts


def normalize_raw_contacts(raw_contacts: List[RawContact], dept_alias_map: Dict[str, str]) -> List[Contact]:
    """Преобразует RawContact в модель Contact с использованием сокращений отделов."""

    normalized: List[Contact] = []
    for raw in raw_contacts:
        group = dept_alias_map.get(raw.full_department_name, raw.full_department_name)
        normalized.append(
            Contact(
                group=group,
                name=raw.full_name,
                office=raw.internal_extension or "",
                mobile="",
                other="",
                photo="",
            )
        )
    return normalized
