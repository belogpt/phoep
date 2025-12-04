"""Маршруты Flask для работы с телефонной книгой, группами и настройками."""
import os
import re
from datetime import datetime
from typing import Dict, List
from flask import (
    Blueprint,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file,
    abort,
)
from io import BytesIO
from PIL import Image, UnidentifiedImageError

from .models import Contact
from . import repository
from .excel_io import export_to_excel, import_from_excel

phonebook_bp = Blueprint('phonebook', __name__)

PHOTOS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'static', 'photos')
PHOTO_MAX_SIZE = (100, 100)

BUILTIN_ICONS: List[tuple[str, str]] = [
    ("Нет иконки", ""),
    ("Семья (синяя)", "Resource:icon_family_b.png"),
    ("Семья (зелёная)", "Resource:icon_family_g.png"),
    ("Семья (оранжевая)", "Resource:icon_family_o.png"),
    ("Семья (красная)", "Resource:icon_family_r.png"),
    ("Семья (жёлтая)", "Resource:icon_family_y.png"),
]

os.makedirs(PHOTOS_DIR, exist_ok=True)


def _filter_contacts(contacts, group_filter, search):
    filtered = contacts
    if group_filter:
        filtered = [c for c in filtered if c.group == group_filter]
    if search:
        term = search.lower()
        filtered = [
            c for c in filtered
            if term in c.name.lower() or term in c.office.lower() or term in c.mobile.lower() or term in c.other.lower()
        ]
    return filtered


def make_default_photo_value(filename: str) -> str:
    """Возвращает значение default_photo для пользовательского файла."""
    return f"/static/photos/{filename}"


def _slugify(text: str) -> str:
    slug = re.sub(r'[^\w]+', '_', text, flags=re.UNICODE).strip('_').lower()
    return slug or 'photo'


def _process_photo_upload(file_storage, contact_name: str) -> str:
    """Обрабатывает загрузку изображения, нормализует и сохраняет PNG."""
    try:
        image = Image.open(file_storage.stream)
    except UnidentifiedImageError as exc:  # noqa: BLE001
        raise ValueError('Невозможно прочитать изображение. Загрузите корректный PNG или JPEG.') from exc

    image = image.convert('RGBA')
    image.thumbnail(PHOTO_MAX_SIZE, Image.Resampling.LANCZOS)

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    filename = f"{_slugify(contact_name)}_{timestamp}.png"
    full_path = os.path.join(PHOTOS_DIR, filename)
    image.save(full_path, format='PNG')
    return filename


def _list_uploaded_photos() -> List[str]:
    files = []
    if os.path.exists(PHOTOS_DIR):
        files = [f for f in os.listdir(PHOTOS_DIR) if f.lower().endswith('.png')]
    return sorted(files)


def _extract_photo_filename(photo_value: str) -> str | None:
    if not photo_value:
        return None
    if photo_value.startswith('/static/photos/'):
        return photo_value.split('/static/photos/', maxsplit=1)[1]
    if photo_value.startswith('Config:photos/'):
        return photo_value.split('Config:photos/', maxsplit=1)[1]
    return None


def _count_photo_usage() -> Dict[str, int]:
    usage: Dict[str, int] = {}
    contacts = repository.load_contacts()
    for contact in contacts:
        filename = _extract_photo_filename(contact.photo)
        if filename:
            usage[filename] = usage.get(filename, 0) + 1
    return usage


def _remove_photo_if_unused(filename: str) -> None:
    usage = _count_photo_usage()
    if usage.get(filename):
        return
    path = os.path.join(PHOTOS_DIR, filename)
    if os.path.exists(path):
        os.remove(path)


@phonebook_bp.route('/')
def index():
    try:
        contacts = repository.load_contacts()
    except Exception as exc:  # noqa: BLE001
        flash(f'Ошибка чтения телефонной книги: {exc}', 'danger')
        contacts = []
    group_filter = request.args.get('group', '')
    search = request.args.get('search', '')
    filtered_contacts = _filter_contacts(contacts, group_filter, search)
    groups = repository.get_groups_with_counts()
    return render_template(
        'index.html',
        contacts=filtered_contacts,
        groups=groups,
        group_filter=group_filter,
        search=search,
    )


@phonebook_bp.route('/contact/new', methods=['GET', 'POST'])
def new_contact():
    if request.method == 'POST':
        return _save_contact()
    groups = repository.get_groups_with_counts()
    return render_template(
        'edit_contact.html',
        contact=None,
        groups=groups,
        builtin_icons=BUILTIN_ICONS,
        uploaded_photos=_list_uploaded_photos(),
    )


@phonebook_bp.route('/contact/<int:contact_id>/edit', methods=['GET', 'POST'])
def edit_contact(contact_id: int):
    contacts = repository.load_contacts()
    contact = next((c for c in contacts if c.contact_id == contact_id), None)
    if not contact:
        abort(404)
    if request.method == 'POST':
        return _save_contact(contact_id)
    groups = repository.get_groups_with_counts()
    return render_template(
        'edit_contact.html',
        contact=contact,
        groups=groups,
        builtin_icons=BUILTIN_ICONS,
        uploaded_photos=_list_uploaded_photos(),
    )


@phonebook_bp.route('/contact/<int:contact_id>/delete', methods=['POST'])
def delete_contact(contact_id: int):
    try:
        repository.delete_contact(contact_id)
        flash('Контакт удалён', 'success')
    except Exception as exc:  # noqa: BLE001
        flash(f'Не удалось удалить контакт: {exc}', 'danger')
    return redirect(url_for('phonebook.index'))


@phonebook_bp.route('/groups')
def manage_groups():
    groups = repository.get_groups_with_counts()
    return render_template('manage_groups.html', groups=groups)


@phonebook_bp.route('/groups/rename', methods=['POST'])
def rename_group():
    old_name = request.form.get('old_name', '').strip()
    new_name = request.form.get('new_name', '').strip()
    if not new_name:
        flash('Новое имя группы не может быть пустым', 'danger')
    else:
        try:
            repository.rename_group(old_name, new_name)
            flash('Группа переименована', 'success')
        except Exception as exc:  # noqa: BLE001
            flash(f'Ошибка переименования: {exc}', 'danger')
    return redirect(url_for('phonebook.manage_groups'))


@phonebook_bp.route('/groups/delete', methods=['POST'])
def remove_group():
    group_name = request.form.get('group_name', '')
    mode = request.form.get('mode', 'keep')
    with_contacts = mode == 'delete_contacts'
    try:
        repository.delete_group(group_name, with_contacts=with_contacts)
        flash('Группа удалена', 'success')
    except Exception as exc:  # noqa: BLE001
        flash(f'Ошибка удаления группы: {exc}', 'danger')
    return redirect(url_for('phonebook.manage_groups'))


@phonebook_bp.route('/RemotePhonebook.xml')
def download_xml():
    _, phonebook_path = repository.get_paths()
    if not os.path.exists(phonebook_path):
        repository.ensure_environment()
    return send_file(
        phonebook_path,
        mimetype='application/xml',
        as_attachment=True,
        download_name=repository.PHONEBOOK_FILENAME,
    )


@phonebook_bp.route('/export/excel')
def export_excel():
    contacts = repository.load_contacts()
    data = export_to_excel(contacts)
    filename = f"phonebook_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(BytesIO(data), as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@phonebook_bp.route('/import/excel', methods=['POST'])
def import_excel():
    file = request.files.get('excel_file')
    if not file:
        flash('Не выбран файл для импорта', 'danger')
        return redirect(url_for('phonebook.index'))

    try:
        imported_count = import_from_excel(file)
        return render_template('import_result.html', count=imported_count)
    except Exception as exc:  # noqa: BLE001
        flash(f'Ошибка импорта: {exc}', 'danger')
        return redirect(url_for('phonebook.index'))


@phonebook_bp.route('/photos')
def manage_photos():
    usage = _count_photo_usage()
    photos = [
        {
            'filename': filename,
            'usage': usage.get(filename, 0),
            'url': make_default_photo_value(filename),
        }
        for filename in _list_uploaded_photos()
    ]
    return render_template('manage_photos.html', photos=photos)


@phonebook_bp.route('/photos/delete/<path:filename>', methods=['POST'])
def delete_photo(filename: str):
    safe_name = os.path.basename(filename)
    usage = _count_photo_usage()
    if usage.get(safe_name):
        flash(f'Нельзя удалить {safe_name}: используется {usage[safe_name]} контактов', 'danger')
        return redirect(url_for('phonebook.manage_photos'))
    path = os.path.join(PHOTOS_DIR, safe_name)
    if not os.path.exists(path):
        flash('Файл не найден', 'danger')
        return redirect(url_for('phonebook.manage_photos'))
    os.remove(path)
    flash('Изображение удалено', 'success')
    return redirect(url_for('phonebook.manage_photos'))


def _save_contact(contact_id=None):
    group = request.form.get('group', '').strip()
    name = request.form.get('name', '').strip()
    office = request.form.get('office', '').strip()
    mobile = request.form.get('mobile', '').strip()
    other = request.form.get('other', '').strip()
    photo_option = request.form.get('photo_option', 'none')
    builtin_photo = request.form.get('builtin_photo', '').strip()
    existing_photo = request.form.get('existing_photo', '').strip()
    previous_photo = None

    if contact_id is not None:
        contacts = repository.load_contacts()
        existing = next((c for c in contacts if c.contact_id == contact_id), None)
        if not existing:
            abort(404)
        previous_photo = existing.photo

    try:
        if photo_option == 'builtin':
            photo_value = builtin_photo
        elif photo_option == 'existing':
            photo_value = make_default_photo_value(existing_photo) if existing_photo else ''
        elif photo_option == 'upload':
            upload = request.files.get('photo_file')
            if not upload or upload.filename == '':
                raise ValueError('Не выбран файл для загрузки изображения')
            filename = _process_photo_upload(upload, name or group or 'contact')
            photo_value = make_default_photo_value(filename)
        else:
            photo_value = ''
    except ValueError as exc:  # noqa: BLE001
        flash(str(exc), 'danger')
        if contact_id is None:
            return redirect(url_for('phonebook.new_contact'))
        return redirect(url_for('phonebook.edit_contact', contact_id=contact_id))

    contact = Contact(group=group, name=name, office=office, mobile=mobile, other=other, photo=photo_value, contact_id=contact_id)
    try:
        if contact_id is None:
            repository.add_contact(contact)
            flash('Контакт добавлен', 'success')
        else:
            repository.update_contact(contact_id, contact)
            flash('Контакт обновлён', 'success')
    except Exception as exc:  # noqa: BLE001
        flash(f'Ошибка сохранения: {exc}', 'danger')
        if contact_id is None:
            return redirect(url_for('phonebook.new_contact'))
        return redirect(url_for('phonebook.edit_contact', contact_id=contact_id))

    if previous_photo and previous_photo != contact.photo:
        old_filename = _extract_photo_filename(previous_photo)
        new_filename = _extract_photo_filename(contact.photo)
        if old_filename and old_filename != new_filename:
            _remove_photo_if_unused(old_filename)
    return redirect(url_for('phonebook.index'))
