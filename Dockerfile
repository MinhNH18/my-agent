# ── Build stage ───────────────────────────────────────────────────────────────
FROM python:3.12-slim AS builder

WORKDIR /build
COPY requirements.txt .
RUN pip install --no-cache-dir --prefix=/install -r requirements.txt


# ── Runtime stage ─────────────────────────────────────────────────────────────
FROM python:3.12-slim

LABEL maintainer="Zalopay FP&A"
LABEL description="Contract Review Agent — Public REST API"

# Copy installed packages từ builder
COPY --from=builder /install /usr/local

WORKDIR /app
COPY agent.py api.py ./

# Biến môi trường (truyền khi chạy container, không hardcode ở đây)
ENV ANTHROPIC_API_KEY=""
ENV API_SECRET_KEY=""
ENV PORT=8000

EXPOSE 8000

# Dùng $PORT để tương thích Railway / Render (họ tự set PORT)
CMD uvicorn api:app --host 0.0.0.0 --port ${PORT}
