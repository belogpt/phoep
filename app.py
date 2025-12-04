"""Точка входа Flask-приложения для управления удалённой телефонной книгой Yealink."""
import os
from functools import wraps
from flask import Flask, request, Response

from phonebook.routes import phonebook_bp
from phonebook.repository import ensure_environment


def create_app() -> Flask:
    """Создаёт и настраивает экземпляр Flask, регистрирует блюпринты."""
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret')

    # Инициализируем конфигурационный файл и папку с данными
    ensure_environment()

    username = os.environ.get('BASIC_AUTH_USERNAME', 'admin')
    password = os.environ.get('BASIC_AUTH_PASSWORD', 'admin')

    def check_auth(user: str, pwd: str) -> bool:
        return user == username and pwd == password

    def authenticate() -> Response:
        return Response(
            'Требуется авторизация', 401,
            {'WWW-Authenticate': 'Basic realm="Login Required"'}
        )

    def requires_auth(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            auth = request.authorization
            if not auth or not check_auth(auth.username, auth.password):
                return authenticate()
            return f(*args, **kwargs)
        return decorated

    # Подключаем блюпринт и защищаем все маршруты базовой авторизацией
    app.register_blueprint(phonebook_bp)

    for rule in app.url_map.iter_rules():
        if rule.endpoint == 'static':
            continue
        view_func = app.view_functions[rule.endpoint]
        app.view_functions[rule.endpoint] = requires_auth(view_func)

    return app


app = create_app()


if __name__ == '__main__':
    port = int(os.environ.get('APP_PORT', 5000))
    app.run(host='0.0.0.0', port=port)
