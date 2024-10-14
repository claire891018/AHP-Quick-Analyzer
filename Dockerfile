FROM python:3.9-slim

WORKDIR /app

COPY . .

# 安裝所需的 Python 套件
RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 8501 8000

RUN apt-get update && apt-get install -y supervisor
COPY supervisord.conf /etc/supervisor/conf.d/supervisord.conf

CMD ["supervisord", "-n"]
