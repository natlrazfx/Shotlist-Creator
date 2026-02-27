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

DMG_STAGING="dist/dmg"
DMG_NAME="${APP_NAME}-${VERSION}-macOS.dmg"
DMG_PATH="dist/${DMG_NAME}"

rm -rf "${DMG_STAGING}" "${DMG_PATH}"
mkdir -p "${DMG_STAGING}"

cp -R "dist/${APP_NAME}.app" "${DMG_STAGING}/${APP_NAME}.app"
ln -s /Applications "${DMG_STAGING}/Applications"

hdiutil create \
  -volname "${APP_NAME}" \
  -srcfolder "${DMG_STAGING}" \
  -ov \
  -format UDZO \
  "${DMG_PATH}"

echo "Built: ${DMG_PATH}"
