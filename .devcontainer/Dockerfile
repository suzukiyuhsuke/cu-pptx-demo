FROM mcr.microsoft.com/devcontainers/python:3.11-bookworm

# --- FIX: remove yarn repo that breaks apt update ---
RUN rm -f /etc/apt/sources.list.d/yarn.list \
    && rm -f /etc/apt/trusted.gpg.d/yarn.gpg \
    && rm -f /usr/share/keyrings/yarnkey.gpg || true

# --- prerequisites ---
RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
    curl \
    gnupg \
  && rm -rf /var/lib/apt/lists/*

# --- Azure CLI (official Microsoft repo) ---
RUN mkdir -p /etc/apt/keyrings \
 && curl -fsSL https://packages.microsoft.com/keys/microsoft.asc \
    | gpg --dearmor -o /etc/apt/keyrings/microsoft.gpg \
 && echo "deb [arch=amd64 signed-by=/etc/apt/keyrings/microsoft.gpg] \
    https://packages.microsoft.com/repos/azure-cli/ bookworm main" \
    > /etc/apt/sources.list.d/azure-cli.list \
 && apt-get update \
 && apt-get install -y --no-install-recommends \
    azure-cli \
 && apt-get clean \
 && rm -rf /var/lib/apt/lists/*

# --- Azure Developer CLI (azd) ---
RUN curl -fsSL https://aka.ms/install-azd.sh | bash

# --- LibreOffice + fonts (Japanese-safe) ---
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-impress \
    fonts-noto-cjk \
    fontconfig \
    libgl1 \
  && apt-get clean \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /workspace
