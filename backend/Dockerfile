# ---- Base image
FROM python:3.11-slim

# Faster installs & no cache
ENV PIP_NO_CACHE_DIR=1

# Workdir
WORKDIR /app

# Install dependencies (make sure uvicorn exists)
COPY requirements.txt /app/requirements.txt
RUN pip install --upgrade pip && pip install -r /app/requirements.txt
# safety net in case requirements.txt missed uvicorn
RUN python -c "import importlib; import sys; sys.exit(0 if importlib.util.find_spec('uvicorn') else 1)" \
  || pip install 'uvicorn[standard]'

# Copy the app code
COPY . /app

# Railway injects $PORT
ENV PORT=8000
EXPOSE 8000

# Start FastAPI
CMD sh -c "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000}"


