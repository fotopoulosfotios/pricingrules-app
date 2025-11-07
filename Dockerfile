FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# system deps για libmagic (αναγνώριση τύπου αρχείων)
RUN apt-get update && apt-get install -y --no-install-recommends \
    locales libmagic1 \
    && rm -rf /var/lib/apt/lists/* \
    && sed -i 's/# el_GR.UTF-8 UTF-8/el_GR.UTF-8 UTF-8/' /etc/locale.gen \
    && locale-gen

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]

