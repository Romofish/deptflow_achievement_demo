# syntax=docker/dockerfile:1.6
#
# DeptFlow Achievement Reporter — OpenShift-ready container image.
#
# Design notes (Novartis OpenShift / OCP):
#   * OpenShift runs containers with a random, non-root UID that belongs to the root
#     group (GID 0). The image therefore must NOT hardcode a USER uid, and all files
#     the app needs to read/write must be group-readable/writable for group 0.
#   * Streamlit listens on a non-privileged port (8080).
#   * /opt/app-root/src/outputs is created and made group-writable so report history
#     and generated PPTX files can be written by the arbitrary runtime UID. Mount a
#     PVC there for persistence.
#   * Health endpoint for liveness/readiness probes: GET /_stcore/health -> "ok".

############################
# Builder stage
############################
FROM registry.access.redhat.com/ubi9/python-311:latest AS builder

USER 0
WORKDIR /opt/app-root/src

COPY requirements.txt ./

# Build wheels into a local dir so the runtime stage stays slim and offline-installable.
RUN python -m pip install --upgrade pip wheel \
 && pip wheel --no-cache-dir --wheel-dir /tmp/wheels -r requirements.txt

############################
# Runtime stage
############################
FROM registry.access.redhat.com/ubi9/python-311:latest

LABEL org.opencontainers.image.title="DeptFlow Achievement Reporter" \
      org.opencontainers.image.description="Streamlit app: SharePoint achievement CSV -> dashboard + fixed 8-slide PPTX." \
      org.opencontainers.image.source="https://example.invalid/deptflow_achievement_demo" \
      io.openshift.expose-services="8080:http" \
      io.openshift.tags="streamlit,python,reporting"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    APP_HOME=/opt/app-root/src \
    PORT=8080 \
    STREAMLIT_SERVER_PORT=8080 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false \
    STREAMLIT_SERVER_ENABLE_CORS=false \
    STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION=true \
    HOME=/opt/app-root/src

WORKDIR ${APP_HOME}

# Install pre-built wheels from the builder.
USER 0
COPY --from=builder /tmp/wheels /tmp/wheels
COPY requirements.txt ./
RUN pip install --no-cache-dir --no-index --find-links=/tmp/wheels -r requirements.txt \
 && rm -rf /tmp/wheels

# Copy application source.
COPY app.py ./
COPY src ./src
COPY .streamlit ./.streamlit
COPY README.md Agent.md ./

# Create runtime-writable directories and grant group-0 ownership/permissions.
# This is the OpenShift "arbitrary UID" pattern: files are owned by root:0 with
# group-rwX so any random UID assigned by OCP can still read/write them.
RUN mkdir -p ${APP_HOME}/outputs ${APP_HOME}/sample_data ${APP_HOME}/.streamlit \
 && chgrp -R 0 ${APP_HOME} \
 && chmod -R g=u ${APP_HOME}

EXPOSE 8080

# Do NOT pin a numeric UID here. OpenShift will inject one. We only declare a
# non-root user; using a high UID keeps `docker run` (without OCP) also non-root.
USER 1001

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD python -c "import urllib.request,sys; \
sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8080/_stcore/health', timeout=3).read().strip()==b'ok' else 1)"

ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.port=8080", \
            "--server.address=0.0.0.0", \
            "--server.headless=true", \
            "--browser.gatherUsageStats=false"]
