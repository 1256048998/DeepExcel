FROM python:3.12-slim-bookworm

# Non-root from the start. A container that only ever needs to read its own code
# and talk to Postgres has no reason to run as root.
RUN useradd --create-home --shell /usr/sbin/nologin deepexcel

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /srv

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY alembic.ini ./
COPY alembic ./alembic
COPY app ./app

USER deepexcel
EXPOSE 8000

# Schema migration is an explicit operational step, never something that happens
# implicitly on container start: an automatic migration during a rolling deploy
# runs concurrently in every replica.
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8000/healthz', timeout=4).status==200 else 1)"

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--proxy-headers", "--forwarded-allow-ips", "*"]
