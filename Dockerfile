FROM node:24-alpine AS web-build
WORKDIR /build/web
COPY web/package.json web/package-lock.json ./
RUN npm ci --ignore-scripts
COPY web/ ./
RUN npm run build

FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    SHEETPILOT_WEB_DATA_DIR=/app/data_web
WORKDIR /app
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*
COPY pyproject.toml README.md ./
COPY src/ ./src/
COPY --from=web-build /build/web/dist ./web/dist
RUN pip install --no-cache-dir .
RUN useradd --create-home --uid 10001 sheetpilot && mkdir -p /app/data_web && chown -R sheetpilot:sheetpilot /app
USER sheetpilot
EXPOSE 8000
CMD ["uvicorn", "sheetpilot.web_api:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
