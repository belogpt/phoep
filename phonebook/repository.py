"""Работа с Config.cfg и XML-файлом телефонной книги Yealink."""
import os
import configparser
import xml.etree.ElementTree as ET
from typing import List, Tuple

from .models import Contact, Group

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Config.cfg')
DEFAULT_REMOTE_DIR = './data'
PHONEBOOK_FILENAME = 'RemotePhonebook.xml'
MAX_GROUPS = 50
MAX_NAME_LENGTH = 99
MAX_PHONE_LENGTH = 32
REMOVE_EMPTY_GROUPS = True


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
    cid = 0
    for menu in root.findall('Menu'):
        group_name = menu.get('Name', '')
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
    root = ET.Element('YealinkIPPhoneBook')
    for group_name, group_contacts in groups.items():
        menu = ET.SubElement(root, 'Menu')
        menu.set('Name', group_name)
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
    return [Group(name=k, contact_count=v) for k, v in sorted(counts.items())]


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
    save_contacts(contacts)


def delete_group(group_name: str, with_contacts: bool) -> None:
    contacts = load_contacts()
    if with_contacts:
        contacts = [c for c in contacts if c.group != group_name]
    else:
        for c in contacts:
            if c.group == group_name:
                raise ValueError('Нельзя удалить группу с контактами без подтверждения')
    save_contacts(contacts)
