FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Install python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Cloud platforms inject PORT environment variable (default to 5050)
ENV PORT=5050
EXPOSE 5050

CMD exec gunicorn app:app --bind 0.0.0.0:${PORT:-5050} --workers 2 --timeout 120
