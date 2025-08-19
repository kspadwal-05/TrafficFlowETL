# SQL-only TrafficFlow ETL with MS Access support
FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive
WORKDIR /app

# System deps: Java for UCanAccess, ODBC for pyodbc, CA certs for HTTPS
RUN apt-get update && apt-get install -y --no-install-recommends \
    openjdk-17-jre-headless \
    curl unzip \
    unixodbc unixodbc-dev \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# UCanAccess jars (JDBC driver for Access)
ENV UCANACCESS_VERSION=5.0.1
RUN mkdir -p /opt/ucanaccess && \
    curl -fsSL -o /tmp/ucan.zip https://downloads.sourceforge.net/project/ucanaccess/UCanAccess-${UCANACCESS_VERSION}-bin.zip && \
    unzip -q /tmp/ucan.zip -d /opt && \
    mv /opt/UCanAccess-${UCANACCESS_VERSION}/* /opt/ucanaccess/ && \
    rm -rf /tmp/ucan.zip
ENV CLASSPATH=/opt/ucanaccess/ucanaccess.jar:/opt/ucanaccess/lib/*

# Python deps
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# App code and data
COPY src ./src
COPY legacy ./legacy
COPY data ./data
COPY scripts ./scripts

ENV PYTHONPATH=/app
CMD ["python", "-m", "src.etl_main"]
