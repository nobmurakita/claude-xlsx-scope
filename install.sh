#!/usr/bin/env bash
set -euo pipefail

REPO="nobmurakita/cc-read-xlsx"
SKILL_NAME="cc-read-xlsx"
INSTALL_DIR="${HOME}/.claude/skills/${SKILL_NAME}"

# 最新リリースの zip URL を取得
echo "Fetching latest release..."
RELEASE_JSON="$(curl -fsSL "https://api.github.com/repos/${REPO}/releases/latest")"

ZIP_URL="$(echo "$RELEASE_JSON" | grep '"browser_download_url"' | grep '\.zip"' | head -1 | sed 's/.*"browser_download_url": "\(.*\)".*/\1/')"

if [ -z "$ZIP_URL" ]; then
  echo "Error: could not find zip asset in latest release." >&2
  exit 1
fi

echo "Downloading ${ZIP_URL}..."
TMP_ZIP="$(mktemp /tmp/cc-read-xlsx-XXXXXX.zip)"
curl -fsSL -o "$TMP_ZIP" "$ZIP_URL"

# インストール先を準備
mkdir -p "$INSTALL_DIR"

echo "Installing to ${INSTALL_DIR}..."
unzip -o "$TMP_ZIP" -d "$INSTALL_DIR"
rm -f "$TMP_ZIP"

# 実行権限を付与
chmod +x "${INSTALL_DIR}/scripts/cc-read-xlsx"
chmod +x "${INSTALL_DIR}/scripts/cc-read-xlsx-"* 2>/dev/null || true

echo "Done. Installed to ${INSTALL_DIR}"
