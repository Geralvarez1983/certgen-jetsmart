FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    git \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /tmp/certgen_data/uploads /tmp/certgen_data/modelos /tmp/certgen_data/data

ENV RENDER=true

EXPOSE 10000

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000", "--timeout", "300", "--workers", "1"]
