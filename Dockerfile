# Multi-stage build for md2excel with Mermaid support
FROM python:3.12-slim AS builder

# Install uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

# Set working directory
WORKDIR /app

# Copy project files
COPY pyproject.toml ./
COPY md2excel ./md2excel

# Install dependencies
RUN uv pip install --system -e .

# Final stage
FROM python:3.12-slim

# Install system dependencies for Playwright, .NET (Spire.XLS), and Japanese fonts
RUN apt-get update && apt-get install -y \
    # .NET Core dependencies (required by Spire.XLS)
    libicu-dev \
    # Playwright/Chromium dependencies
    libnspr4 \
    libnss3 \
    libasound2t64 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libdbus-1-3 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxrandr2 \
    libgbm1 \
    libpango-1.0-0 \
    libcairo2 \
    libatspi2.0-0 \
    libxshmfence1 \
    # Japanese fonts
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    fonts-ipafont \
    fonts-ipaexfont \
    && rm -rf /var/lib/apt/lists/*

# Install uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

# Set working directory
WORKDIR /app

# Copy project files
COPY pyproject.toml ./
COPY md2excel ./md2excel

# Install Python dependencies
RUN uv pip install --system -e .

# Install Playwright browsers
RUN python -m playwright install chromium

# Create directory for input/output files
RUN mkdir -p /data
WORKDIR /data

# Set entrypoint
ENTRYPOINT ["md2excel"]
CMD ["--help"]
