FROM selenium/standalone-chrome:latest

USER root

ENV PYTHONIOENCODING=utf-8
ENV LANG=C.UTF-8
ENV LC_ALL=C.UTF-8

# Install Python and pip
RUN apt-get update && apt-get install -y python3 python3-pip \
    --no-install-recommends && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt --break-system-packages

COPY parser.py .
RUN mkdir -p /output

ENTRYPOINT ["python3", "-u", "parser.py"]