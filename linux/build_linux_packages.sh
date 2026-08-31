#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-1.0.0}"
if [[ ! "$VERSION" =~ ^[0-9]+([.][0-9]+)*([+-][A-Za-z0-9.]+)?$ ]]; then
  VERSION="1.0.0"
fi

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
APP_BINARY="$PROJECT_ROOT/dist/Corpus Querier"
OUTPUT_DIR="$PROJECT_ROOT/linux-output"
APPDIR="$PROJECT_ROOT/build-linux/Corpus-Querier.AppDir"
DEB_ROOT="$PROJECT_ROOT/build-linux/deb-root"

if [[ ! -x "$APP_BINARY" ]]; then
  echo "Missing Linux executable: $APP_BINARY" >&2
  echo "Run PyInstaller before this script." >&2
  exit 1
fi

rm -rf "$PROJECT_ROOT/build-linux" "$OUTPUT_DIR"
mkdir -p "$OUTPUT_DIR"

# AppImage directory
mkdir -p \
  "$APPDIR/usr/bin" \
  "$APPDIR/usr/share/applications" \
  "$APPDIR/usr/share/icons/hicolor/scalable/apps"
install -m 755 "$APP_BINARY" "$APPDIR/usr/bin/corpus-querier"
install -m 755 "$PROJECT_ROOT/linux/AppRun" "$APPDIR/AppRun"
install -m 644 "$PROJECT_ROOT/linux/corpus-querier.desktop" \
  "$APPDIR/corpus-querier.desktop"
install -m 644 "$PROJECT_ROOT/linux/corpus-querier.desktop" \
  "$APPDIR/usr/share/applications/corpus-querier.desktop"
install -m 644 "$PROJECT_ROOT/assets/corpus-querier.svg" \
  "$APPDIR/corpus-querier.svg"
install -m 644 "$PROJECT_ROOT/assets/corpus-querier.svg" \
  "$APPDIR/usr/share/icons/hicolor/scalable/apps/corpus-querier.svg"

ARCH=x86_64 VERSION="$VERSION" \
  "$PROJECT_ROOT/appimagetool-x86_64.AppImage" --appimage-extract-and-run \
  "$APPDIR" "$OUTPUT_DIR/Corpus-Querier-x86_64.AppImage"
chmod +x "$OUTPUT_DIR/Corpus-Querier-x86_64.AppImage"

# Debian package
mkdir -p \
  "$DEB_ROOT/DEBIAN" \
  "$DEB_ROOT/usr/bin" \
  "$DEB_ROOT/usr/share/applications" \
  "$DEB_ROOT/usr/share/icons/hicolor/scalable/apps" \
  "$DEB_ROOT/usr/share/doc/corpus-querier"
install -m 755 "$APP_BINARY" "$DEB_ROOT/usr/bin/corpus-querier"
install -m 644 "$PROJECT_ROOT/linux/corpus-querier.desktop" \
  "$DEB_ROOT/usr/share/applications/corpus-querier.desktop"
install -m 644 "$PROJECT_ROOT/assets/corpus-querier.svg" \
  "$DEB_ROOT/usr/share/icons/hicolor/scalable/apps/corpus-querier.svg"
install -m 644 "$PROJECT_ROOT/LICENSE" \
  "$DEB_ROOT/usr/share/doc/corpus-querier/copyright"

sed "s/@VERSION@/$VERSION/g" "$PROJECT_ROOT/linux/control.template" \
  > "$DEB_ROOT/DEBIAN/control"
chmod 644 "$DEB_ROOT/DEBIAN/control"

dpkg-deb --build --root-owner-group "$DEB_ROOT" \
  "$OUTPUT_DIR/corpus-querier_${VERSION}_amd64.deb"

echo "Linux packages created in $OUTPUT_DIR"

