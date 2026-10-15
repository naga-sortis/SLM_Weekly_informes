#!/usr/bin/env bash
# SLM Weekly Informes — install or update the systemd service.
#
#   sudo bash deploy/install-service.sh          (asks before stopping a manually started copy)
#   sudo bash deploy/install-service.sh --yes    (no questions)
#
# What it does (safe to run again for every update — the data is never deleted):
#   1. runs deploy/preflight.sh and stops if any check fails
#   2. creates the service user (no login) and the data folder  /var/lib/slm-informes
#   3. stops a manually started copy and MOVES its database to the data folder (verified copy; the old file is renamed, not deleted)
#   4. copies the code to /opt/slm-weekly-informes (the git folder stays yours; data is never inside the code folder)
#   5. writes /etc/slm-informes/slm-informes.env (only the first time) and the systemd units
#   6. enables + starts the service and a daily backup timer, checks that the page answers, makes a first backup
#
# Settings (environment variables): PORT HOST DATA_DIR INSTALL_DIR SERVICE_NAME SERVICE_USER NODE_BIN OLD_DB
#   FORCE_MIGRATE=1  replace an existing service database with OLD_DB (a backup of it is made first)
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SERVICE_NAME="${SERVICE_NAME:-slm-informes}"
ENV_FILE="/etc/$SERVICE_NAME/$SERVICE_NAME.env"
# On an update, keep the port/host already configured unless they are given explicitly.
if [ -f "$ENV_FILE" ]; then
  [ -n "${PORT:-}" ] || PORT="$(sed -n 's/^PORT=//p' "$ENV_FILE" | tail -1)"
  [ -n "${HOST:-}" ] || HOST="$(sed -n 's/^HOST=//p' "$ENV_FILE" | tail -1)"
  CONFIGURED_DB="$(sed -n 's/^SLM_DB=//p' "$ENV_FILE" | tail -1)"
  if [ -z "${DATA_DIR:-}" ] && [ -n "$CONFIGURED_DB" ]; then DATA_DIR="$(dirname "$CONFIGURED_DB")"; fi
fi
export PORT="${PORT:-8080}" HOST="${HOST:-0.0.0.0}"
export DATA_DIR="${DATA_DIR:-/var/lib/slm-informes}" INSTALL_DIR="${INSTALL_DIR:-/opt/slm-weekly-informes}"
export SERVICE_NAME SERVICE_USER="${SERVICE_USER:-slm}"
export NODE_BIN="${NODE_BIN:-$(command -v node 2>/dev/null || true)}"
YES=0
if [ "${1:-}" = "--yes" ] || [ "${1:-}" = "-y" ]; then YES=1; fi
DB_FILE="$DATA_DIR/slm-informes.sqlite"
STAMP="$(date +%Y%m%d-%H%M%S)"

say() { printf '\n\e[1m==> %s\e[0m\n' "$*"; }
die() { printf '\n\e[31mERROR:\e[0m %s\n' "$*" >&2; exit 1; }
confirm() {
  [ "$YES" = "1" ] && return 0
  read -r -p "$1 [y/N] " a; [ "$a" = "y" ] || [ "$a" = "Y" ] || [ "$a" = "yes" ]
}

[ "$(id -u)" -eq 0 ] || die "Run it with sudo:  sudo bash deploy/install-service.sh"
[ -n "$NODE_BIN" ] || die "Node.js not found (set NODE_BIN=/path/to/node)."
NODE_BIN="$(readlink -f "$NODE_BIN")"; export NODE_BIN

# ------------------------------------------------------------------ 1. pre-flight
say "1/6 Pre-flight checks"
if ! bash "$REPO_DIR/deploy/preflight.sh"; then die "Pre-flight checks failed — nothing was changed."; fi

# ------------------------------------------------------------------ 2. user + folders
say "2/6 Service user and data folder"
if ! getent passwd "$SERVICE_USER" >/dev/null; then
  useradd --system --user-group --no-create-home --home-dir "$DATA_DIR" --shell /usr/sbin/nologin "$SERVICE_USER"
  echo "Created system user '$SERVICE_USER'."
fi
install -d -m 0750 -o "$SERVICE_USER" -g "$SERVICE_USER" "$DATA_DIR" "$DATA_DIR/backups"
echo "Data folder: $DATA_DIR (owner $SERVICE_USER, mode 750)"

# ------------------------------------------------------------------ 3. stop manual copies, move the database
say "3/6 Existing server and database"
manual=""
for d in /proc/[0-9]*; do
  cmd="$(tr '\0' ' ' <"$d/cmdline" 2>/dev/null)" || continue
  case "$cmd" in *server/server.js*) ;; *) continue ;; esac
  case "$(basename "${cmd%% *}")" in node*) ;; *) continue ;; esac   # the node process itself, not npm/sh wrappers
  grep -q "$SERVICE_NAME.service" "$d/cgroup" 2>/dev/null && continue
  manual="$manual ${d#/proc/}"
done
OLD_DB="${OLD_DB:-}"
cands=("$REPO_DIR/data/slm-informes.sqlite")
for p in $manual; do
  cwd="$(readlink "/proc/$p/cwd" 2>/dev/null || true)"
  envdb="$(tr '\0' '\n' <"/proc/$p/environ" 2>/dev/null | sed -n 's/^SLM_DB=//p' || true)"
  if [ -n "$envdb" ]; then case "$envdb" in /*) cands+=("$envdb") ;; *) cands+=("$cwd/$envdb") ;; esac; fi
  cands+=("$cwd/data/slm-informes.sqlite")
done
if [ -z "$OLD_DB" ]; then for c in "${cands[@]}"; do [ -f "$c" ] && { OLD_DB="$c"; break; }; done; fi
if [ -n "$OLD_DB" ] && [ "$(readlink -f "$OLD_DB")" = "$(readlink -f "$DB_FILE")" ]; then OLD_DB=""; fi

if [ -n "$manual" ]; then
  echo "A manually started server is running (PID$manual)."
  confirm "Stop it now so the service can take over? (its data is kept and moved)" || die "Cancelled — nothing was changed except the user/data folder."
  for p in $manual; do kill -TERM "$p" 2>/dev/null || true; done
  for _ in $(seq 1 40); do alive=""; for p in $manual; do kill -0 "$p" 2>/dev/null && alive=1; done; [ -z "$alive" ] && break; sleep 0.25; done
  for p in $manual; do kill -0 "$p" 2>/dev/null && { echo "PID $p did not stop, forcing."; kill -KILL "$p" 2>/dev/null || true; }; done
  echo "Stopped."
fi

if [ -n "$OLD_DB" ] && [ -f "$OLD_DB" ]; then
  if [ -f "$DB_FILE" ] && [ "${FORCE_MIGRATE:-0}" != "1" ]; then
    echo "Keeping the existing service database $DB_FILE (not replaced by $OLD_DB; use FORCE_MIGRATE=1 to replace it)."
  else
    if [ -f "$DB_FILE" ]; then
      SLM_DB="$DB_FILE" "$NODE_BIN" "$REPO_DIR/server/backup.js" --copy "$DATA_DIR/backups/before-migration-$STAMP.sqlite"
    fi
    systemctl stop "$SERVICE_NAME" 2>/dev/null || true
    SLM_DB="$OLD_DB" "$NODE_BIN" "$REPO_DIR/server/backup.js" --copy "$DB_FILE"
    rm -f "$DB_FILE-wal" "$DB_FILE-shm"
    chown "$SERVICE_USER:$SERVICE_USER" "$DB_FILE"; chmod 0640 "$DB_FILE"
    # Rename (not delete) the old file so nobody keeps writing to it by mistake.
    for ext in "" -wal -shm; do [ -f "$OLD_DB$ext" ] && mv "$OLD_DB$ext" "$OLD_DB.moved-$STAMP$ext"; done
    echo "Database moved to $DB_FILE. The old file was renamed to $OLD_DB.moved-$STAMP (delete it once you are happy)."
  fi
elif [ -f "$DB_FILE" ]; then
  echo "Using the existing service database $DB_FILE."
else
  echo "No existing database: a new, empty one will be created at $DB_FILE."
fi

# ------------------------------------------------------------------ 4. code
say "4/6 Application code -> $INSTALL_DIR"
if [ "$(readlink -f "$REPO_DIR")" = "$(readlink -f "$INSTALL_DIR")" ]; then
  echo "The code folder is already $INSTALL_DIR — using it in place."
  [ ! -d "$INSTALL_DIR/data" ] || echo "Note: $INSTALL_DIR/data is not used by the service (data is in $DATA_DIR)."
else
  NEW="$INSTALL_DIR.new-$STAMP"
  rm -rf "$NEW"; mkdir -p "$NEW"
  cp -a "$REPO_DIR/index.html" "$REPO_DIR/package.json" "$REPO_DIR/README.md" "$REPO_DIR/assets" "$REPO_DIR/vendor" "$REPO_DIR/server" "$NEW/"
  if command -v git >/dev/null 2>&1 && [ -d "$REPO_DIR/.git" ]; then
    git -C "$REPO_DIR" log -1 --format='%H %cd' --date=iso >"$NEW/DEPLOYED_VERSION" 2>/dev/null || true
  fi
  chown -R root:root "$NEW"
  find "$NEW" -type d -exec chmod 0755 {} +
  find "$NEW" -type f -exec chmod 0644 {} +
  # Swap in one step, so a failed copy never leaves a half-updated site.
  if [ -d "$INSTALL_DIR" ]; then mv "$INSTALL_DIR" "$INSTALL_DIR.old-$STAMP"; fi
  mv "$NEW" "$INSTALL_DIR"
  rm -rf "$INSTALL_DIR.old-$STAMP"
  echo "Installed $(cat "$INSTALL_DIR/DEPLOYED_VERSION" 2>/dev/null || echo 'current code')."
fi

# ------------------------------------------------------------------ 5. configuration + units
say "5/6 Configuration and systemd units"
install -d -m 0750 -o root -g "$SERVICE_USER" "/etc/$SERVICE_NAME"
if [ ! -f "$ENV_FILE" ]; then
  cat >"$ENV_FILE" <<EOF
# SLM Weekly Informes — service settings. After editing:  sudo systemctl restart $SERVICE_NAME
PORT=$PORT
HOST=$HOST
SLM_DB=$DB_FILE
SLM_MAX_UPLOAD_MB=300
# Number of daily backups kept in $DATA_DIR/backups
SLM_BACKUP_KEEP=30
# Optional login for everybody (user:password). Leave empty to rely on the network/VPN.
SLM_BASIC_AUTH=
EOF
  echo "Wrote $ENV_FILE"
else
  echo "Keeping the existing settings in $ENV_FILE"
fi
chown root:"$SERVICE_USER" "$ENV_FILE"; chmod 0640 "$ENV_FILE"

if [ "$PORT" -lt 1024 ]; then CAPS=$'AmbientCapabilities=CAP_NET_BIND_SERVICE\nCapabilityBoundingSet=CAP_NET_BIND_SERVICE'
else CAPS='CapabilityBoundingSet='; fi
render() {
  local out="$1" tpl="$2" content
  content="$(sed -e "s#@INSTALL_DIR@#$INSTALL_DIR#g" -e "s#@DATA_DIR@#$DATA_DIR#g" -e "s#@NODE_BIN@#$NODE_BIN#g" \
    -e "s#@SERVICE_USER@#$SERVICE_USER#g" -e "s#@SERVICE_NAME@#$SERVICE_NAME#g" "$tpl")"
  printf '%s\n' "${content//@CAPABILITIES@/$CAPS}" >"$out"
  chmod 0644 "$out"
  echo "Wrote $out"
}
render "/etc/systemd/system/$SERVICE_NAME.service" "$REPO_DIR/deploy/systemd/slm-informes.service.in"
render "/etc/systemd/system/$SERVICE_NAME-backup.service" "$REPO_DIR/deploy/systemd/slm-informes-backup.service.in"
render "/etc/systemd/system/$SERVICE_NAME-backup.timer" "$REPO_DIR/deploy/systemd/slm-informes-backup.timer.in"
if command -v systemd-analyze >/dev/null 2>&1; then
  systemd-analyze verify "/etc/systemd/system/$SERVICE_NAME.service" "/etc/systemd/system/$SERVICE_NAME-backup.service" "/etc/systemd/system/$SERVICE_NAME-backup.timer" \
    || echo "(systemd-analyze reported the notes above — usually harmless)"
fi

# ------------------------------------------------------------------ 6. start + verify
say "6/6 Start and verify"
systemctl daemon-reload
systemctl enable "$SERVICE_NAME.service" "$SERVICE_NAME-backup.timer" >/dev/null
systemctl restart "$SERVICE_NAME.service"
systemctl start "$SERVICE_NAME-backup.timer"
CHECK_HOST="$HOST"
if [ "$HOST" = "0.0.0.0" ] || [ "$HOST" = "::" ]; then CHECK_HOST=127.0.0.1; fi
if HEALTH="$("$NODE_BIN" -e "
  (async()=>{for(let i=0;i<60;i++){try{const r=await fetch('http://$CHECK_HOST:$PORT/api/health');if(r.status===200||r.status===401){
    const j=r.status===200?await r.json():null;console.log(j?'version '+j.version+', '+j.stats.tickets+' tickets, '+j.stats.imports+' imports, '+j.stats.classifications+' classifications':'login required (SLM_BASIC_AUTH)');process.exit(0)}}catch(e){}
    await new Promise(r=>setTimeout(r,500))}process.exit(1)})()")"; then
  echo "Service is up: $HEALTH"
else
  systemctl --no-pager status "$SERVICE_NAME" | tail -15 || true
  journalctl -u "$SERVICE_NAME" -n 30 --no-pager || true
  die "The service did not answer on $CHECK_HOST:$PORT — see the log above."
fi
if systemctl start "$SERVICE_NAME-backup.service"; then
  echo "First backup: $(ls -1t "$DATA_DIR/backups"/slm-informes-*.sqlite 2>/dev/null | head -1)"
else
  echo "WARNING: the first backup failed — see: journalctl -u $SERVICE_NAME-backup"
fi

IP="$(hostname -I 2>/dev/null | awk '{print $1}')"
cat <<EOF

$(printf '\e[32m\e[1m')Installed.$(printf '\e[0m') The tool now starts by itself at boot and restarts after crashes and updates.

  Open:      http://${IP:-<server-ip>}:$PORT/
  Status:    systemctl status $SERVICE_NAME
  Logs:      journalctl -u $SERVICE_NAME -f
  Restart:   sudo systemctl restart $SERVICE_NAME
  Settings:  $ENV_FILE   (then restart)
  Data:      $DB_FILE
  Backups:   $DATA_DIR/backups  (daily at 02:15, last ${SLM_BACKUP_KEEP:-30} kept — copy them off this server too)
  Update:    cd $REPO_DIR && git pull && sudo bash deploy/install-service.sh --yes
  Restore:   sudo bash deploy/restore.sh <backup-file>
EOF
NEXT="$(systemctl list-timers "$SERVICE_NAME-backup.timer" --no-pager --no-legend 2>/dev/null | awk '{print $1, $2, $3}' || true)"
if [ -n "$NEXT" ]; then echo "  Next backup: $NEXT"; fi
