"""Работа с Config.cfg и XML-файлом телефонной книги Yealink."""
import configparser
import json
import os
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple

from .models import Contact, Group

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Config.cfg')
DEFAULT_REMOTE_DIR = '/app/data'
PHONEBOOK_FILENAME = os.environ.get("PHONEBOOK_FILENAME", "rem.xml")
MAX_GROUPS = 50
MAX_NAME_LENGTH = 99
MAX_PHONE_LENGTH = 32
REMOVE_EMPTY_GROUPS = True
GROUP_ORDER_FILENAME = 'group_order.json'
PREFIX_PATTERN = re.compile(r"^\s*(\d{2})\.\s+(.*)$")


def ensure_environment() -> None:
    """Создаёт Config.cfg и директорию данных при отсутствии."""
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_PATH):
        config['OutPutPath'] = {
            'RemotePhoneDir': DEFAULT_REMOTE_DIR,
            'LocalPhoneDir': DEFAULT_REMOTE_DIR,
        }
        _write_config(config)
    else:
        config.read(CONFIG_PATH)
    remote_dir = config['OutPutPath'].get('RemotePhoneDir', DEFAULT_REMOTE_DIR)
    os.makedirs(remote_dir, exist_ok=True)
    phonebook_path = os.path.join(remote_dir, PHONEBOOK_FILENAME)
    if not os.path.exists(phonebook_path):
        _write_empty_phonebook(phonebook_path)


def _get_group_order_path() -> str:
    remote_dir, _ = get_paths()
    os.makedirs(remote_dir, exist_ok=True)
    return os.path.join(remote_dir, GROUP_ORDER_FILENAME)


def load_group_order() -> Dict[str, int]:
    """Загружает порядок групп из JSON."""

    path = _get_group_order_path()
    if not os.path.exists(path):
        return {}
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {k: int(v) for k, v in data.items()}


def save_group_order(order_map: Dict[str, int]) -> None:
    """Сохраняет порядок групп в JSON."""

    path = _get_group_order_path()
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(order_map, f, ensure_ascii=False, indent=2)


def _split_prefixed_group(name: str) -> tuple[str, int | None]:
    match = PREFIX_PATTERN.match(name)
    if match:
        return match.group(2).strip(), int(match.group(1))
    return name, None


def _normalize_order_map(order_map: Dict[str, int], group_names: List[str]) -> Dict[str, int]:
    ordered_pairs = []
    for name in group_names:
        if name in order_map and isinstance(order_map[name], int):
            ordered_pairs.append((name, order_map[name]))

    ordered_pairs.sort(key=lambda x: (x[1], x[0]))
    max_value = max((val for _, val in ordered_pairs), default=0)
    missing = [name for name in group_names if name not in order_map]
    for name in sorted(missing):
        max_value += 1
        ordered_pairs.append((name, max_value))

    ordered_pairs.sort(key=lambda x: x[1])
    return {name: idx for idx, (name, _) in enumerate(ordered_pairs, start=1)}


def _write_empty_phonebook(path: str) -> None:
    """Создаёт пустой XML-файл телефонной книги."""
    root = ET.Element('YealinkIPPhoneBook')
    tree = ET.ElementTree(root)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tree.write(path, encoding='utf-8', xml_declaration=True)


def _write_config(config: configparser.ConfigParser) -> None:
    with open(CONFIG_PATH, 'w', encoding='utf-8') as cfg:
        config.write(cfg)


def load_config() -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH)
    return config


def save_remote_dir(new_dir: str) -> None:
    """Обновляет путь RemotePhoneDir в Config.cfg."""
    config = load_config()
    if 'OutPutPath' not in config:
        config['OutPutPath'] = {}
    config['OutPutPath']['RemotePhoneDir'] = new_dir
    if 'LocalPhoneDir' not in config['OutPutPath']:
        config['OutPutPath']['LocalPhoneDir'] = DEFAULT_REMOTE_DIR
    os.makedirs(new_dir, exist_ok=True)
    _write_config(config)
    phonebook_path = os.path.join(new_dir, PHONEBOOK_FILENAME)
    if not os.path.exists(phonebook_path):
        _write_empty_phonebook(phonebook_path)


def get_paths() -> Tuple[str, str]:
    """Возвращает (remote_dir, phonebook_path)."""
    config = load_config()
    remote_dir = config.get('OutPutPath', 'RemotePhoneDir', fallback=DEFAULT_REMOTE_DIR)
    phonebook_path = os.path.join(remote_dir, PHONEBOOK_FILENAME)
    return remote_dir, phonebook_path


def load_contacts() -> List[Contact]:
    """Читает XML и возвращает список контактов с присвоенными ID."""
    _, phonebook_path = get_paths()
    if not os.path.exists(phonebook_path):
        _write_empty_phonebook(phonebook_path)
    tree = ET.parse(phonebook_path)
    root = tree.getroot()
    contacts: List[Contact] = []
    derived_order: Dict[str, int] = {}
    cid = 0
    for position, menu in enumerate(root.findall('Menu'), start=1):
        raw_group_name = menu.get('Name', '')
        group_name, prefixed_order = _split_prefixed_group(raw_group_name)
        derived_order.setdefault(group_name, prefixed_order or position)
        for unit in menu.findall('Unit'):
            contact = Contact(
                group=group_name,
                name=unit.get('Name', ''),
                office=unit.get('Phone1', ''),
                mobile=unit.get('Phone2', ''),
                other=unit.get('Phone3', ''),
                photo=unit.get('default_photo', ''),
                contact_id=cid,
            )
            contacts.append(contact)
            cid += 1
    existing_order = load_group_order()
    if not existing_order:
        normalized = _normalize_order_map(derived_order, list(derived_order.keys())) if derived_order else {}
        if normalized:
            save_group_order(normalized)
    else:
        missing = [name for name in derived_order if name not in existing_order]
        if missing:
            merged = {**existing_order, **{name: derived_order[name] for name in missing}}
            save_group_order(_normalize_order_map(merged, list(merged.keys())))
    return contacts


def _validate_lengths(contact: Contact) -> None:
    if len(contact.group) == 0:
        raise ValueError('Группа обязательна')
    if len(contact.group) > MAX_NAME_LENGTH:
        raise ValueError('Название группы слишком длинное (макс 99)')
    if len(contact.name) == 0:
        raise ValueError('Имя обязательно')
    if len(contact.name) > MAX_NAME_LENGTH:
        raise ValueError('Имя слишком длинное (макс 99)')
    for number in [contact.office, contact.mobile, contact.other]:
        if len(number) > MAX_PHONE_LENGTH:
            raise ValueError('Номер слишком длинный (макс 32)')
    if len(contact.photo) > MAX_NAME_LENGTH:
        raise ValueError('Путь к фото слишком длинный (макс 99)')


def save_contacts(contacts: List[Contact], preserved_groups: List[str] | None = None) -> None:
    """Сохраняет список контактов в XML, соблюдая ограничения.

    preserved_groups позволяет оставить пустые группы, если нужно не удалять их автоматически.
    """
    groups = {}
    for contact in contacts:
        _validate_lengths(contact)
        if contact.group not in groups:
            if len(groups) >= MAX_GROUPS:
                raise ValueError('Превышено число групп (50)')
            groups[contact.group] = []
        groups[contact.group].append(contact)

    if preserved_groups:
        for g in preserved_groups:
            if g not in groups:
                if len(groups) >= MAX_GROUPS:
                    raise ValueError('Превышено число групп (50)')
                groups[g] = []

    _, phonebook_path = get_paths()
    group_names = list(groups.keys())
    order_map = _normalize_order_map(load_group_order(), group_names)
    if order_map:
        save_group_order(order_map)
    ordered_names = sorted(group_names, key=lambda n: (order_map.get(n, 10**6), n))
    root = ET.Element('YealinkIPPhoneBook')
    for idx, group_name in enumerate(ordered_names, start=1):
        group_contacts = groups[group_name]
        menu = ET.SubElement(root, 'Menu')
        prefix = f"{idx:02d}. "
        menu.set('Name', f"{prefix}{group_name}")
        for contact in group_contacts:
            unit = ET.SubElement(menu, 'Unit')
            unit.set('Name', contact.name)
            unit.set('default_photo', contact.photo)
            unit.set('Phone1', contact.office)
            unit.set('Phone2', contact.mobile)
            unit.set('Phone3', contact.other)
    tree = ET.ElementTree(root)
    os.makedirs(os.path.dirname(phonebook_path), exist_ok=True)
    tree.write(phonebook_path, encoding='utf-8', xml_declaration=True)


def get_groups_with_counts() -> List[Group]:
    contacts = load_contacts()
    counts = {}
    for contact in contacts:
        counts[contact.group] = counts.get(contact.group, 0) + 1
    order_map = _normalize_order_map(load_group_order(), list(counts.keys())) if counts else {}
    if order_map:
        save_group_order(order_map)
    groups = [
        Group(name=k, contact_count=v, order_index=order_map.get(k, 0))
        for k, v in counts.items()
    ]
    return sorted(groups, key=lambda g: (g.order_index or 10**6, g.name))


def update_contact(contact_id: int, updated: Contact) -> None:
    contacts = load_contacts()
    for idx, contact in enumerate(contacts):
        if contact.contact_id == contact_id:
            contacts[idx] = Contact(**{**updated.__dict__, 'contact_id': contact_id})
            save_contacts(contacts)
            return
    raise ValueError('Контакт не найден')


def add_contact(contact: Contact) -> None:
    contacts = load_contacts()
    contact.contact_id = max([c.contact_id for c in contacts], default=-1) + 1
    contacts.append(contact)
    save_contacts(contacts)


def delete_contact(contact_id: int) -> None:
    existing = load_contacts()
    preserved = list({c.group for c in existing}) if not REMOVE_EMPTY_GROUPS else None
    contacts = [c for c in existing if c.contact_id != contact_id]
    save_contacts(contacts, preserved_groups=preserved)


def rename_group(old_name: str, new_name: str) -> None:
    contacts = load_contacts()
    for contact in contacts:
        if contact.group == old_name:
            contact.group = new_name
    order_map = load_group_order()
    if old_name in order_map:
        order_map[new_name] = order_map.pop(old_name)
        order_map = _normalize_order_map(order_map, list(order_map.keys()))
        save_group_order(order_map)
    save_contacts(contacts)


def delete_group(group_name: str, with_contacts: bool) -> None:
    contacts = load_contacts()
    if with_contacts:
        contacts = [c for c in contacts if c.group != group_name]
    else:
        for c in contacts:
            if c.group == group_name:
                raise ValueError('Нельзя удалить группу с контактами без подтверждения')
    order_map = load_group_order()
    if group_name in order_map:
        order_map.pop(group_name, None)
    remaining_groups = list({c.group for c in contacts})
    if order_map:
        order_map = _normalize_order_map(order_map, remaining_groups)
        save_group_order(order_map)
    save_contacts(contacts)


def update_group_order(new_order: List[str]) -> None:
    """Обновляет порядок групп и перезаписывает XML."""

    seen = set()
    cleaned_order = []
    for name in new_order:
        stripped = name.strip()
        if stripped and stripped not in seen:
            cleaned_order.append(stripped)
            seen.add(stripped)

    current_contacts = load_contacts()
    existing_groups = list({c.group for c in current_contacts})
    for g in existing_groups:
        if g not in seen:
            cleaned_order.append(g)

    if cleaned_order:
        order_map = {name: idx for idx, name in enumerate(cleaned_order, start=1)}
        order_map = _normalize_order_map(order_map, cleaned_order)
        save_group_order(order_map)
        save_contacts(current_contacts, preserved_groups=existing_groups if not REMOVE_EMPTY_GROUPS else None)
