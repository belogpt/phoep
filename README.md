# Yealink Remote Phonebook Manager

Веб-приложение на Flask для работы с удалённой телефонной книгой Yealink. Позволяет редактировать XML-файл `RemotePhonebook.xml` через браузер, импортировать и экспортировать книгу в Excel, управлять группами и путями к файлам. Интерфейс и описания на русском языке. Предназначено для внутреннего сервера (Linux/РЕД ОС).

## Архитектура проекта
- **Backend:** Flask (Python 3) + стандартная библиотека `xml.etree.ElementTree` для чтения/записи XML.
- **Frontend:** HTML5/CSS/JavaScript с Bootstrap из CDN.
- **Хранение:** исходный файл `RemotePhonebook.xml`; Excel импорт/экспорт через `pandas`, `openpyxl`, `xlrd`.
- **Безопасность:** простая HTTP Basic Auth (логин/пароль через переменные окружения).

## Структура репозитория
```
app.py                   # точка входа
phonebook/
  __init__.py
  models.py              # модели контактов и групп
  repository.py          # работа с Config.cfg и XML
  excel_io.py            # логика Excel импорта/экспорта
  routes.py              # Flask-маршруты
templates/
  base.html, index.html, edit_contact.html, manage_groups.html,
  settings.html, import_result.html
static/
  styles.css, main.js
data/                    # каталог по умолчанию для RemotePhonebook.xml
Config.cfg               # пример конфигурации
Dockerfile
docker-compose.yml
.dockerignore
requirements.txt
```

## Установка без Docker
1. Требуется Python 3.11+ и `pip`.
2. Создайте окружение и установите зависимости:
   ```bash
   python -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```
3. Запуск приложения:
   ```bash
   python app.py
   ```
4. Порт задаётся переменной окружения `APP_PORT` (по умолчанию 5000). Приложение слушает `0.0.0.0`.
5. Basic Auth: переменные `BASIC_AUTH_USERNAME` и `BASIC_AUTH_PASSWORD` (по умолчанию `admin`/`admin`). Для безопасности измените их сразу после запуска.
6. Конфигурация путей:
   - Файл `Config.cfg` в корне проекта.
   - Секция `[OutPutPath]`, ключ `RemotePhoneDir` — каталог, где хранится `RemotePhonebook.xml`.
   - При первом запуске создаётся `Config.cfg` с
     ```
     RemotePhoneDir=./data
     LocalPhoneDir=./data
     ```
   - Каталог создаётся автоматически, файл телефонной книги тоже (пустой каркас).

## Использование
- Главная страница `/` — список контактов, фильтр по группе, поиск, добавление/редактирование/удаление контактов.
- Управление группами `/groups` — переименование и удаление групп (можно удалять вместе с контактами или запретить).
- Настройки `/settings` — изменение `RemotePhoneDir` (сохранение в `Config.cfg`).
- Скачивание XML: кнопка «Скачать XML» или прямой URL `http://<server>:5000/RemotePhonebook.xml` (этот же URL указывать в телефоне Yealink в разделе Directory → Remote Phone Book).

### Импорт/экспорт Excel
- Поддерживаются форматы `.xls` и `.xlsx`.
- Ожидаемые столбцы (без учёта регистра):
  - Department
  - Name
  - Office Number
  - Mobile Number
  - Other Number
  - Head Portrait
- Экспорт формирует файл с этими колонками.
- Импорт: **полностью перезаписывает** текущую телефонную книгу. Пустые строки или строки без имени и номеров пропускаются. Перед загрузкой отображается предупреждение.
- Результаты импорта: количество загруженных контактов или сообщение об ошибке (неверные заголовки, повреждённый файл и т.п.).

### Ограничения и валидация
- До 50 групп.
- Имя группы/контакта — до 99 символов.
- Номера — до 32 символов.
- Если группа пустая при сохранении контакта — запись не создаётся.
- После удаления контакта пустые группы удаляются (поведение включено по умолчанию в коде).

## Развёртывание в Docker (Linux/РЕД ОС)
1. Сборка образа:
   ```bash
   docker build -t yealink-phonebook .
   ```
2. Запуск одной командой (можно копировать как есть):
   ```bash
   docker build -t yealink-phonebook . && \
   docker run -d --name yealink-phonebook -p 5000:5000 -v $(pwd)/data:/app/data yealink-phonebook
   ```
   - `$(pwd)/data` — локальная папка для данных (Config.cfg и RemotePhonebook.xml сохраняются здесь вне контейнера).
   - Контейнер слушает порт 5000.
   - Переменные окружения можно передать флагами `-e BASIC_AUTH_USERNAME=admin -e BASIC_AUTH_PASSWORD=admin123 -e APP_PORT=5000`.
3. Docker Compose:
   ```bash
   docker-compose up -d
   ```
   `docker-compose.yml` монтирует `./data` в `/app/data` и пробрасывает порт `5000:5000`.

## Безопасность
- Приложение рассчитано на внутреннюю сеть. В Docker используется dev-сервер Flask; для публичного доступа следует поставить полноценный WSGI (gunicorn) и прокси.
- Смените стандартные логин/пароль Basic Auth сразу после развёртывания.

## Работа с Yealink
- В настройках телефона Yealink: Directory → Remote Phone Book.
- URL: `http://<ip_сервера>:5000/RemotePhonebook.xml` (учитывает Basic Auth, если настроена на телефоне).

## Дополнительно
- Все пути обрабатываются кросс-платформенно через `os.path`.
- При отсутствии `RemotePhonebook.xml` создаётся файл с каркасом:
  ```xml
  <?xml version="1.0" encoding="UTF-8"?>
  <YealinkIPPhoneBook>
  </YealinkIPPhoneBook>
  ```
- При первом старте создаётся каталог `./data`.
