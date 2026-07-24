# AfriCareer AI — Streamlit container for Hugging Face Spaces (Docker SDK) / Render / any container host
FROM python:3.11-slim

# System deps kept minimal; python-docx / PyPDF2 are pure-Python wheels
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_SERVER_ENABLECORS=false \
    STREAMLIT_SERVER_ENABLEXSRFPROTECTION=false \
    STREAMLIT_BROWSER_GATHERUSAGESTATS=false

WORKDIR /app

# Install deps first for better layer caching
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# App code (see .dockerignore — .env, analytics, and the .zipx are excluded)
COPY app.py .

# Hugging Face Spaces expects the app on port 7860 by default.
# The container runs as a non-root user (HF requirement / good practice).
RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app
USER appuser

EXPOSE 7860

HEALTHCHECK CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://localhost:7860/_stcore/health').status==200 else 1)" || exit 1

CMD ["streamlit", "run", "app.py", \
     "--server.port=7860", \
     "--server.address=0.0.0.0"]
