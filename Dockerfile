# Cloud server option (any Linux VPS / Render / Railway / Fly.io): docker build -t permit-server . && docker run -p 8080:8080 -v permitdata:/data permit-server
FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-eng && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements-cloud.txt .
RUN pip install --no-cache-dir -r requirements-cloud.txt
COPY . .
# set PERMIT_ADMIN_PASSWORD (and optionally PERMIT_ADMIN_USER) on the host - the first admin is made at start
ENV PERMIT_SERVER_DATA=/data PYTHONUNBUFFERED=1
VOLUME /data
EXPOSE 8080
CMD ["python", "server.py", "--no-browser"]
