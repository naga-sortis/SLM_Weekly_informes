#!/usr/bin/env bash
# SLM Weekly Informes — remove the systemd service. The DATA IS KEPT (database, backups and settings).
#
#   sudo bash deploy/uninstall-service.sh
set -euo pipefail
SERVICE_NAME="${SERVICE_NAME:-slm-informes}"
INSTALL_DIR="${INSTALL_DIR:-/opt/slm-weekly-informes}"
[ "$(id -u)" -eq 0 ] || { echo "Run it with sudo." >&2; exit 1; }
systemctl disable --now "$SERVICE_NAME-backup.timer" "$SERVICE_NAME.service" 2>/dev/null || true
rm -f "/etc/systemd/system/$SERVICE_NAME.service" "/etc/systemd/system/$SERVICE_NAME-backup.service" "/etc/systemd/system/$SERVICE_NAME-backup.timer"
systemctl daemon-reload
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
if [ "$(readlink -f "$REPO_DIR")" != "$(readlink -f "$INSTALL_DIR")" ]; then rm -rf "$INSTALL_DIR"; fi
echo "Service removed. Kept: /var/lib/slm-informes (database + backups) and /etc/$SERVICE_NAME (settings)."
