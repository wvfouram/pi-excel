#!/bin/bash
# Removes Excel.
SRC="$(cd "$(dirname "$0")" && pwd)"
if [ "$(id -u)" = "0" ]; then
  echo "Run this as your normal user (not with sudo)."
  exit 1
fi

# put back the apps that opened these file types before
if [ -f "$HOME/.config/pi-excel/old-defaults" ]; then
  while read -r t old; do
    [ -n "$old" ] && xdg-mime default "$old" "$t" 2>/dev/null || true
  done < "$HOME/.config/pi-excel/old-defaults"
  rm -f "$HOME/.config/pi-excel/old-defaults"
fi

sudo rm -rf /opt/pi-excel /usr/local/bin/pi-excel /usr/share/applications/pi-excel.desktop \
  /usr/share/icons/hicolor/scalable/apps/pi-excel.svg
sudo update-desktop-database -q /usr/share/applications 2>/dev/null || true
echo "Excel removed. Your settings are still in ~/.config/pi-excel (delete that folder to remove them)."
