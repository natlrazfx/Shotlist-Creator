#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-2.1.15}"
APP_NAME="ShotlistCreator"

ICON_ARG=()
if [ -f "icon.icns" ]; then
  ICON_ARG=(--icon "icon.icns")
elif [ -f "icon.png" ]; then
  ICON_ARG=(--icon "icon.png")
fi

python3 -m PyInstaller --noconfirm --windowed --name "${APP_NAME}" \
  --add-data "assets:assets" \
  "${ICON_ARG[@]}" \
  ShotlistCreator.py

ROOT_DIR="dist/pkgroot"
rm -rf "${ROOT_DIR}"
mkdir -p "${ROOT_DIR}/Applications"
cp -R "dist/${APP_NAME}.app" "${ROOT_DIR}/Applications/${APP_NAME}.app"

pkgbuild \
  --root "${ROOT_DIR}" \
  --identifier "com.natlrazfx.shotlistcreator" \
  --version "${VERSION}" \
  --install-location "/" \
  "dist/${APP_NAME}-${VERSION}-macOS.pkg"

echo "Built: dist/${APP_NAME}-${VERSION}-macOS.pkg"
