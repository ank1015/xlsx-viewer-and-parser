#!/usr/bin/env sh
set -eu

REPO="${HEYSNAP_XLSXL_REPO:-ank1015/xlsx-viewer-and-parser}"
VERSION="${HEYSNAP_XLSXL_VERSION:-latest}"
INSTALL_DIR="${HEYSNAP_XLSXL_INSTALL_DIR:-/usr/local/bin}"
BIN_NAME="heysnap-xlsxl"

need() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "Missing required command: $1" >&2
    exit 1
  fi
}

need curl
need tar
need uname

machine="$(uname -m)"
system="$(uname -s)"

case "$system:$machine" in
  Linux:x86_64|Linux:amd64)
    rid="linux-x64"
    ;;
  Linux:aarch64|Linux:arm64)
    rid="linux-arm64"
    ;;
  Darwin:x86_64|Darwin:amd64)
    rid="osx-x64"
    ;;
  Darwin:arm64|Darwin:aarch64)
    rid="osx-arm64"
    ;;
  *)
    echo "Unsupported platform: $system $machine" >&2
    exit 1
    ;;
esac

asset="${BIN_NAME}-${rid}.tar.gz"
if [ "$VERSION" = "latest" ]; then
  url="https://github.com/${REPO}/releases/latest/download/${asset}"
else
  url="https://github.com/${REPO}/releases/download/${VERSION}/${asset}"
fi

tmp_dir="$(mktemp -d)"
cleanup() {
  rm -rf "$tmp_dir"
}
trap cleanup EXIT INT TERM

echo "Downloading ${BIN_NAME} ${VERSION} for ${rid}..."
curl -fsSL "$url" -o "$tmp_dir/$asset"
tar -xzf "$tmp_dir/$asset" -C "$tmp_dir"

if [ ! -f "$tmp_dir/$BIN_NAME" ]; then
  echo "Archive did not contain $BIN_NAME" >&2
  exit 1
fi

mkdir -p "$INSTALL_DIR" 2>/dev/null || sudo mkdir -p "$INSTALL_DIR"
if [ -w "$INSTALL_DIR" ]; then
  install "$tmp_dir/$BIN_NAME" "$INSTALL_DIR/$BIN_NAME"
else
  sudo install "$tmp_dir/$BIN_NAME" "$INSTALL_DIR/$BIN_NAME"
fi

echo "Installed ${BIN_NAME} to ${INSTALL_DIR}/${BIN_NAME}"
"$INSTALL_DIR/$BIN_NAME" --help || true
