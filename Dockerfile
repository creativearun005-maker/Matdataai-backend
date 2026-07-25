FROM python:3.10-slim

# Tesseract + Hindi language pack + basic OpenCV runtime deps
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-hin \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Hugging Face Spaces expects the app to listen on port 7860
EXPOSE 7860
CMD ["uvicorn", "backend_api:app", "--host", "0.0.0.0", "--port", "7860"]
