#!/usr/bin/env bash
# SLM Weekly Informes — restore the database from a backup file.
#
#   sudo bash deploy/restore.sh /var/lib/slm-informes/backups/slm-informes-2026-10-15-02-15-00.sqlite
#
# The current database is saved first (backups/before-restore-<date>.sqlite), so a restore can be undone.
set -euo pipefail
SERVICE_NAME="${SERVICE_NAME:-slm-informes}"
ENV_FILE="/etc/$SERVICE_NAME/$SERVICE_NAME.env"
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
[ "$(id -u)" -eq 0 ] || { echo "Run it with sudo." >&2; exit 1; }
SRC="${1:-}"
[ -n "$SRC" ] && [ -f "$SRC" ] || { echo "Usage: sudo bash deploy/restore.sh <backup-file.sqlite>"; ls -1t /var/lib/slm-informes/backups 2>/dev/null | head -10 | sed 's#^#  /var/lib/slm-informes/backups/#'; exit 64; }
SRC="$(readlink -f "$SRC")"
DB_FILE="$(sed -n 's/^SLM_DB=//p' "$ENV_FILE" 2>/dev/null | tail -1)"
[ -n "$DB_FILE" ] || { echo "Service settings $ENV_FILE not found — is the service installed?" >&2; exit 1; }
SERVICE_USER="$(systemctl show -p User --value "$SERVICE_NAME" 2>/dev/null || true)"; SERVICE_USER="${SERVICE_USER:-slm}"
NODE_BIN="${NODE_BIN:-$(command -v node)}"
STAMP="$(date +%Y%m%d-%H%M%S)"

echo "Backup to restore: $(SLM_DB="$SRC" "$NODE_BIN" "$REPO_DIR/server/backup.js" --check)"
[ -f "$DB_FILE" ] && echo "Current database:  $(SLM_DB="$DB_FILE" "$NODE_BIN" "$REPO_DIR/server/backup.js" --check || echo 'unreadable')"
read -r -p "Replace the current database with this backup? [y/N] " a
[ "$a" = "y" ] || [ "$a" = "Y" ] || { echo "Cancelled."; exit 0; }

systemctl stop "$SERVICE_NAME"
if [ -f "$DB_FILE" ]; then
  SLM_DB="$DB_FILE" "$NODE_BIN" "$REPO_DIR/server/backup.js" --copy "$(dirname "$DB_FILE")/backups/before-restore-$STAMP.sqlite" || echo "(could not save the current database — continuing)"
fi
SLM_DB="$SRC" "$NODE_BIN" "$REPO_DIR/server/backup.js" --copy "$DB_FILE.restoring"
rm -f "$DB_FILE-wal" "$DB_FILE-shm"
mv "$DB_FILE.restoring" "$DB_FILE"
chown "$SERVICE_USER:$SERVICE_USER" "$DB_FILE"; chmod 0640 "$DB_FILE"
systemctl start "$SERVICE_NAME"
echo "Restored. $(SLM_DB="$DB_FILE" "$NODE_BIN" "$REPO_DIR/server/backup.js" --check)"
