#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-0.1.0}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PROJECT="$ROOT_DIR/src/OpenXml.Cli/OpenXml.Cli.csproj"
ARTIFACT_DIR="$ROOT_DIR/artifacts/release/v$VERSION"
BIN_NAME="heysnap-xlsxl"
RIDS=("linux-x64" "linux-arm64")

rm -rf "$ARTIFACT_DIR"
mkdir -p "$ARTIFACT_DIR"

for rid in "${RIDS[@]}"; do
  publish_dir="$ARTIFACT_DIR/publish/$rid"
  package_dir="$ARTIFACT_DIR/package/$rid"

  dotnet publish "$PROJECT" \
    -c Release \
    -r "$rid" \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:PublishTrimmed=false \
    -p:DebugType=None \
    -p:DebugSymbols=false \
    -o "$publish_dir"

  mkdir -p "$package_dir"
  cp "$publish_dir/$BIN_NAME" "$package_dir/$BIN_NAME"
  chmod 755 "$package_dir/$BIN_NAME"

  cat > "$package_dir/README.txt" <<EOF
$BIN_NAME $VERSION

Run:
  ./$BIN_NAME xlsx file.xlsx
  ./$BIN_NAME xlsx-assets file.xlsx output-dir
EOF

  tar -C "$package_dir" -czf "$ARTIFACT_DIR/$BIN_NAME-$rid.tar.gz" .
done

cp "$ROOT_DIR/install.sh" "$ARTIFACT_DIR/install.sh"
chmod 755 "$ARTIFACT_DIR/install.sh"

(
  cd "$ARTIFACT_DIR"
  shasum -a 256 "$BIN_NAME"-linux-*.tar.gz install.sh > SHA256SUMS
)

echo "Release artifacts written to $ARTIFACT_DIR"
