"""Маршруты Flask для работы с телефонной книгой, группами и настройками."""
import json
import os
from datetime import datetime
from flask import (
    Blueprint,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file,
    abort,
    session,
)
from io import BytesIO

from .models import Contact
from . import repository
from .excel_io import (
    export_to_excel,
    import_from_excel,
    normalize_raw_contacts,
    parse_raw_department_table,
)
from .departments_registry import (
    load_department_aliases,
    save_department_aliases,
    suggest_alias,
)

phonebook_bp = Blueprint('phonebook', __name__)
RAW_UPLOAD_SESSION_KEY = 'raw_import_file'


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
    group_prefixes = {
        g.name: f"{g.order_index:02d}. " if g.order_index else ''
        for g in groups
    }
    return render_template(
        'index.html',
        contacts=filtered_contacts,
        groups=groups,
        group_filter=group_filter,
        search=search,
        group_prefixes=group_prefixes,
    )


@phonebook_bp.route('/contacts/sort', methods=['POST'])
def sort_contacts_by_name():
    """Сортирует контакты по ФИО и сохраняет порядок в XML."""

    try:
        repository.sort_contacts_by_name()
        flash('Контакты отсортированы по ФИО', 'success')
    except Exception as exc:  # noqa: BLE001
        flash(f'Не удалось отсортировать контакты: {exc}', 'danger')
    return redirect(url_for('phonebook.index'))


@phonebook_bp.route('/contacts/reorder', methods=['GET', 'POST'])
def reorder_contacts():
    if request.method == 'POST':
        payload = request.get_json(silent=True)
        if payload and isinstance(payload.get('order'), list):
            order = payload['order']
        else:
            raw_order = request.form.get('order', '[]')
            try:
                parsed = json.loads(raw_order)
                order = parsed if isinstance(parsed, list) else []
            except json.JSONDecodeError:
                order = []
        try:
            repository.update_contact_order(order)
            flash('Порядок контактов обновлён', 'success')
        except Exception as exc:  # noqa: BLE001
            flash(f'Не удалось обновить порядок: {exc}', 'danger')
        return redirect(url_for('phonebook.reorder_contacts'))

    contacts = repository.load_contacts()
    groups = repository.get_groups_with_counts()
    group_prefixes = {
        g.name: f"{g.order_index:02d}. " if g.order_index else ''
        for g in groups
    }
    return render_template(
        'manage_contacts.html',
        contacts=contacts,
        group_prefixes=group_prefixes,
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


@phonebook_bp.route('/groups/reorder', methods=['POST'])
def reorder_groups():
    payload = request.get_json(silent=True)
    if payload and isinstance(payload.get('order'), list):
        order = [str(item) for item in payload['order']]
    else:
        raw_order = request.form.get('order', '[]')
        try:
            parsed = json.loads(raw_order)
            order = parsed if isinstance(parsed, list) else []
        except json.JSONDecodeError:
            order = [name for name in raw_order.split(',') if name]
    repository.update_group_order(order)
    flash('Порядок групп обновлён', 'success')
    return redirect(url_for('phonebook.manage_groups'))


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


def _save_raw_upload(file) -> str:
    remote_dir, _ = repository.get_paths()
    raw_dir = os.path.join(remote_dir, 'raw_uploads')
    os.makedirs(raw_dir, exist_ok=True)
    filename = f"raw_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    path = os.path.join(raw_dir, filename)
    file.save(path)
    session[RAW_UPLOAD_SESSION_KEY] = path
    return path


def _load_raw_contacts_from_session():
    path = session.get(RAW_UPLOAD_SESSION_KEY)
    if not path or not os.path.exists(path):
        return None, None
    raw_contacts = parse_raw_department_table(path)
    aliases = load_department_aliases()
    return raw_contacts, aliases


def _render_raw_preview(raw_contacts, alias_map):
    normalized = normalize_raw_contacts(raw_contacts, alias_map)
    missing_numbers = sum(1 for raw in raw_contacts if not raw.internal_extension)
    return render_template(
        'import_raw_preview.html',
        contacts=normalized,
        raw_contacts=raw_contacts,
        alias_map=alias_map,
        missing_numbers=missing_numbers,
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


@phonebook_bp.route('/import/raw', methods=['GET', 'POST'])
def import_raw_upload():
    if request.method == 'GET':
        return render_template('import_raw_upload.html')

    file = request.files.get('raw_excel')
    if not file:
        flash('Не выбран файл для импорта', 'danger')
        return redirect(url_for('phonebook.import_raw_upload'))

    try:
        path = _save_raw_upload(file)
        raw_contacts = parse_raw_department_table(path)
    except Exception as exc:  # noqa: BLE001
        flash(f'Ошибка разбора файла: {exc}', 'danger')
        return redirect(url_for('phonebook.import_raw_upload'))

    if not raw_contacts:
        flash(
            'В файле не найдено ни одного контакта. Убедитесь, что таблица соответствует шаблону (строки отделов в колонке B, под ними строки с номерами 1,2,3 и ФИО в колонке D).',
            'warning',
        )
        return redirect(url_for('phonebook.import_raw_upload'))

    departments = sorted({c.full_department_name for c in raw_contacts})
    aliases = load_department_aliases()
    suggestions = {d: aliases.get(d, suggest_alias(d)) for d in departments}
    return render_template(
        'import_raw_departments.html',
        departments=departments,
        suggestions=suggestions,
    )


@phonebook_bp.route('/import/raw/departments', methods=['POST'])
def import_raw_departments():
    raw_contacts, aliases = _load_raw_contacts_from_session()
    if raw_contacts is None:
        flash('Файл загрузки не найден. Пожалуйста, начните импорт заново.', 'danger')
        return redirect(url_for('phonebook.import_raw_upload'))

    names = request.form.getlist('department_names')
    values = request.form.getlist('aliases')
    new_aliases = dict(aliases)
    for dept, alias in zip(names, values):
        dept_name = dept.strip()
        if not dept_name:
            continue
        cleaned = alias.strip()
        new_aliases[dept_name] = cleaned or suggest_alias(dept_name)

    save_department_aliases(new_aliases)
    aliases = load_department_aliases()
    return _render_raw_preview(raw_contacts, aliases)


@phonebook_bp.route('/import/raw/confirm', methods=['POST'])
def import_raw_confirm():
    raw_contacts, aliases = _load_raw_contacts_from_session()
    if raw_contacts is None:
        flash('Файл загрузки не найден. Пожалуйста, начните импорт заново.', 'danger')
        return redirect(url_for('phonebook.import_raw_upload'))

    try:
        normalized = normalize_raw_contacts(raw_contacts, aliases)
        repository.save_contacts(normalized)
        session.pop(RAW_UPLOAD_SESSION_KEY, None)
        return render_template('import_result.html', count=len(normalized))
    except Exception as exc:  # noqa: BLE001
        flash(f'Ошибка импорта: {exc}', 'danger')
        return redirect(url_for('phonebook.import_raw_upload'))


@phonebook_bp.route('/departments/aliases', methods=['GET', 'POST'])
def manage_department_aliases():
    current_aliases = load_department_aliases()
    if request.method == 'POST':
        names = request.form.getlist('department_names')
        values = request.form.getlist('aliases')
        updated: dict[str, str] = {}
        for full, alias in zip(names, values):
            clean_name = full.strip()
            if not clean_name:
                continue
            updated[clean_name] = alias.strip() or suggest_alias(clean_name)
        save_department_aliases(updated)
        flash('Сокращения отделов сохранены', 'success')
        return redirect(url_for('phonebook.manage_department_aliases'))

    items = sorted(current_aliases.items())
    return render_template('departments_aliases.html', aliases=items)


def _save_contact(contact_id=None):
    group = request.form.get('group', '').strip()
    name = request.form.get('name', '').strip()
    office = request.form.get('office', '').strip()
    mobile = request.form.get('mobile', '').strip()
    other = request.form.get('other', '').strip()

    contact = Contact(group=group, name=name, office=office, mobile=mobile, other=other, photo='', contact_id=contact_id)
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
    return redirect(url_for('phonebook.index'))
