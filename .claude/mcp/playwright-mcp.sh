#!/usr/bin/env bash
# Launcher for the Playwright MCP server (@playwright/mcp).
#
# Wrapped in a script rather than inlined into .mcp.json so the same committed
# config works both in a Claude Code on the web container (headless Chromium
# from the image, running as root) and on a developer machine (Playwright's own
# bundled browser, headed).
set -euo pipefail

VERSION="${PLAYWRIGHT_MCP_VERSION:-0.0.83}"
PROJECT_DIR="${CLAUDE_PROJECT_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)}"

args=(--isolated --output-dir "${PROJECT_DIR}/.playwright-mcp")

# No display in the container, and the agent runs as root, so Chromium's
# sandbox cannot start.
if [ "${CLAUDE_CODE_REMOTE:-}" = "true" ] || [ "$(id -u)" = "0" ]; then
  args+=(--headless --no-sandbox)
fi

# Reuse the Chromium from the image: the build that @playwright/mcp's bundled
# playwright-core expects cannot be downloaded here
# (PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1).
CHROMIUM="${PLAYWRIGHT_MCP_CHROMIUM:-/opt/pw-browsers/chromium}"
if [ -x "$CHROMIUM" ]; then
  args+=(--executable-path "$CHROMIUM")
fi

if command -v playwright-mcp >/dev/null 2>&1; then
  exec playwright-mcp "${args[@]}" "$@"
fi
exec npx -y "@playwright/mcp@${VERSION}" "${args[@]}" "$@"
