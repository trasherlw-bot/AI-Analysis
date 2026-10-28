FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1 DATA_ROOT=/data PORT=8000 BASE_PATH=/hde-analysis
# natasha masks personal names but needs ~300 MB of RAM. It is OFF by default (small server).
# Turn on with INSTALL_NATASHA=true in .env, then: docker compose up -d --build
ARG INSTALL_NATASHA=false
RUN pip install --no-cache-dir requests openpyxl \
 && if [ "$INSTALL_NATASHA" = "true" ]; then pip install --no-cache-dir natasha; fi
WORKDIR /srv/app
COPY app/ /srv/app/
RUN mkdir -p /data
EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/api/health', timeout=3)"
CMD ["python", "server.py"]
