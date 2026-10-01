#!/usr/bin/env bash
# Prepare the Playwright MCP server for a Claude Code on the web session.
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

VERSION="${PLAYWRIGHT_MCP_VERSION:-0.0.83}"

# 1. Pre-install the server so the first browser tool call doesn't pay for a
#    cold npx download inside the MCP startup timeout.
if ! command -v playwright-mcp >/dev/null 2>&1; then
  npm install -g "@playwright/mcp@${VERSION}" >/dev/null 2>&1
fi

# 2. Chromium reads its own NSS store, not /etc/ssl/certs, so it does not pick
#    up the session's HTTPS interception CA. Without this every https://
#    navigation fails with ERR_CERT_AUTHORITY_INVALID.
CA=/root/.ccr/agent-proxy-ca.crt
if [ -f "$CA" ]; then
  if ! command -v certutil >/dev/null 2>&1; then
    apt-get install -y -qq libnss3-tools >/dev/null 2>&1 || true
  fi
  if command -v certutil >/dev/null 2>&1; then
    mkdir -p "$HOME/.pki/nssdb"
    certutil -d "sql:$HOME/.pki/nssdb" -L -n ccr-agent-proxy-ca >/dev/null 2>&1 ||
      certutil -d "sql:$HOME/.pki/nssdb" -A -t "C,," -n ccr-agent-proxy-ca -i "$CA" >/dev/null 2>&1 ||
      true
  fi
fi

echo "Playwright MCP ready: browser tools available via the 'playwright' MCP server."
