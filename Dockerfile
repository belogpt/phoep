# Официальный базовый образ Python
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV APP_PORT=5000
ENV BASIC_AUTH_USERNAME=admin
ENV BASIC_AUTH_PASSWORD=admin

VOLUME ["/app/data"]

EXPOSE 5000

CMD ["python", "app.py"]
