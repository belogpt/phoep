"""Импорт и экспорт телефонной книги в формате Excel."""
import io
from typing import List
import pandas as pd

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
            'Head Portrait': contact.photo,
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
        photo = '' if pd.isna(row[columns_map['head portrait']]) else str(row[columns_map['head portrait']]).strip()
        if not department or (not name and not office and not mobile and not other):
            continue
        contacts.append(Contact(group=department, name=name, office=office, mobile=mobile, other=other, photo=photo))
    save_contacts(contacts)
    return len(contacts)
