FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p models

RUN python -m piper.download_voices \
    --download-dir models \
    en_US-lessac-low

ENV PYTHONUNBUFFERED=1

EXPOSE 10000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "10000"]
