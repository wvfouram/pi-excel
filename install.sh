#!/bin/bash
# Installer for Excel (Windows 11 style) on Raspberry Pi
set -e
SRC="$(cd "$(dirname "$0")" && pwd)"
if [ "$(id -u)" = "0" ]; then
  echo "Run this as your normal user (not with sudo). It will ask for sudo when needed."
  exit 1
fi

echo "==> Installing packages"
sudo apt-get update || echo "   (apt update had errors, continuing)"
if ! sudo apt-get install -y python3-gi python3-gi-cairo gir1.2-gtk-3.0 python3-openpyxl fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation xdg-utils; then
  for p in python3-gi python3-gi-cairo gir1.2-gtk-3.0 python3-openpyxl fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation xdg-utils; do
    sudo apt-get install -y "$p" || echo "   (could not install $p)"
  done
fi
sudo apt-get install -y fonts-open-sans || echo "   (Open Sans font not available, using system font)"
for p in ; do
  sudo apt-get install -y "$p" || echo "   ($p not available, skipping)"
done

echo "==> Installing app"
sudo install -d /opt/pi-excel
sudo install -m 644 "$SRC/pi_excel.py" /opt/pi-excel/pi_excel.py
printf '#!/bin/sh\nexec /usr/bin/python3 /opt/pi-excel/pi_excel.py "$@"\n' | sudo tee /usr/local/bin/pi-excel >/dev/null
sudo chmod 755 /usr/local/bin/pi-excel

echo "==> Adding menu entry and icon"
sudo install -D -m 644 "$SRC/pi-excel.svg" /usr/share/icons/hicolor/scalable/apps/pi-excel.svg
sudo gtk-update-icon-cache -f -t /usr/share/icons/hicolor >/dev/null 2>&1 || true
sudo tee /usr/share/applications/pi-excel.desktop >/dev/null <<'DESK'
[Desktop Entry]
Type=Application
Name=Excel
GenericName=Spreadsheet
Comment=Spreadsheets with formulas and charts, saved as .xlsx
Exec=pi-excel %F
Icon=pi-excel
Terminal=false
Categories=Office;Spreadsheet;GTK;
MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;text/csv;
StartupWMClass=pi-excel
StartupNotify=true
DESK
sudo update-desktop-database -q /usr/share/applications 2>/dev/null || true

echo "==> Making Excel the default app for its file types"
mkdir -p "$HOME/.config/pi-excel"
for t in application/vnd.openxmlformats-officedocument.spreadsheetml.sheet text/csv; do
  old=$(xdg-mime query default "$t" 2>/dev/null || true)
  if [ -n "$old" ] && [ "$old" != "pi-excel.desktop" ]; then
    echo "$t $old" >> "$HOME/.config/pi-excel/old-defaults"
  fi
  xdg-mime default pi-excel.desktop "$t" 2>/dev/null || true
done

# the file library comes from apt when Raspberry Pi OS has it, otherwise from pip
if ! /usr/bin/python3 -c "import openpyxl" 2>/dev/null; then
  echo "==> Installing openpyxl with pip (not available from apt)"
  sudo apt-get install -y python3-pip || true
  sudo /usr/bin/python3 -m pip install --break-system-packages openpyxl || echo "   (could not install openpyxl)"
fi

echo
echo "Done. Open Excel from the menu or run: pi-excel"
