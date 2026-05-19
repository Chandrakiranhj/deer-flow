# DeerFlow sidecar — production image for Railway / Render / any Docker host.
#
# Bundles:
#   • Python 3.12 (slim)
#   • uv for fast dep resolution
#   • Chromium headless (needed by build_pptx.py's HTML → PNG → slide pipeline)
#   • node 20 (dom-to-pptx CLI used by build_pptx.py)
#   • All Python deps from packages/harness/pyproject.toml
#
# Build context = the deer-flow/ directory (this Dockerfile's parent).

FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    UV_LINK_MODE=copy \
    PATH="/root/.local/bin:${PATH}"

# System deps:
#   • chromium for the html-ppt → screenshot pipeline (build_pptx.py)
#   • build-essential + libs for any wheels that compile from source
#   • curl for cloud health probes
#   • node + npm so dom-to-pptx can be installed by build_pptx.py via npx
RUN apt-get update && apt-get install -y --no-install-recommends \
        chromium \
        chromium-driver \
        fonts-liberation \
        fonts-noto-color-emoji \
        build-essential \
        curl \
        ca-certificates \
        git \
        nodejs \
        npm \
    && rm -rf /var/lib/apt/lists/*

# Install uv (Astral's Python package manager — what the project uses).
RUN curl -LsSf https://astral.sh/uv/install.sh | sh

WORKDIR /app

# Copy the whole deer-flow tree. Build context is the deer-flow/ directory.
COPY . /app

# Install the harness package and its app dependencies.
WORKDIR /app/backend
RUN /root/.local/bin/uv sync --frozen --no-dev || /root/.local/bin/uv sync --no-dev

# Tell the build scripts where Chromium lives so headless Chrome detection
# short-circuits straight to it instead of probing Windows/macOS paths.
ENV VEPIP_CHROME=/usr/bin/chromium \
    VEPIP_NODE=/usr/bin/node \
    DEER_FLOW_CONFIG_PATH=/app/config.yaml \
    PYTHONPATH=/app/backend

# DeerFlow's runtime state — threads, memory, reports — lives here.
# Mount a persistent volume here in Railway/Render so it survives restarts.
RUN mkdir -p /app/backend/.deer-flow

EXPOSE 8001

# Railway provides $PORT; default to 8001 for local docker runs.
CMD ["sh", "-c", "/root/.local/bin/uv run uvicorn app.gateway.app:app --host 0.0.0.0 --port ${PORT:-8001}"]
