# Railway runtime for 师助AI: Python hosts the API; Node renders editable PPTX.
FROM node:20-bookworm-slim

RUN apt-get update \
    && apt-get install -y --no-install-recommends python3 python3-venv \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt package.json package-lock.json ./
RUN python3 -m venv /opt/venv \
    && /opt/venv/bin/pip install --no-cache-dir -r requirements.txt \
    && npm ci --omit=dev

COPY . ./

ENV PATH="/opt/venv/bin:${PATH}"
ENV PYTHONUNBUFFERED=1

CMD ["python", "ppt_api.py"]
