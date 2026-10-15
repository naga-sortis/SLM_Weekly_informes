#!/usr/bin/env bash
# SLM Weekly Informes — pre-flight checks before installing the systemd service.
# Read-only: it changes nothing on the server (the test server uses a temporary folder that is deleted).
#
#   sudo bash deploy/preflight.sh
#
# Settings (environment variables, same as install-service.sh):
#   PORT=8080  HOST=0.0.0.0  DATA_DIR=/var/lib/slm-informes  INSTALL_DIR=/opt/slm-weekly-informes
#   SERVICE_NAME=slm-informes  SERVICE_USER=slm  NODE_BIN=<auto>  OLD_DB=<auto>  SKIP_TESTS=1 (skip unit tests)
set -u

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PORT="${PORT:-8080}"
HOST="${HOST:-0.0.0.0}"
DATA_DIR="${DATA_DIR:-/var/lib/slm-informes}"
INSTALL_DIR="${INSTALL_DIR:-/opt/slm-weekly-informes}"
SERVICE_NAME="${SERVICE_NAME:-slm-informes}"
SERVICE_USER="${SERVICE_USER:-slm}"
NODE_BIN="${NODE_BIN:-$(command -v node 2>/dev/null || true)}"
MIN_NODE="22.13"

FAILS=0; WARNS=0
if [ -t 1 ]; then G=$'\e[32m'; Y=$'\e[33m'; R=$'\e[31m'; B=$'\e[1m'; N=$'\e[0m'; else G=; Y=; R=; B=; N=; fi
pass() { printf '  %sPASS%s  %s\n' "$G" "$N" "$*"; }
warn() { printf '  %sWARN%s  %s\n' "$Y" "$N" "$*"; WARNS=$((WARNS+1)); }
fail() { printf '  %sFAIL%s  %s\n' "$R" "$N" "$*"; FAILS=$((FAILS+1)); }
info() { printf '  INFO  %s\n' "$*"; }
section() { printf '\n%s%s%s\n' "$B" "$*" "$N"; }

echo "${B}SLM Weekly Informes — pre-flight checks${N}"
echo "Code: $REPO_DIR  ->  installs to $INSTALL_DIR"
echo "Data: $DATA_DIR   Port: $HOST:$PORT   Service: $SERVICE_NAME.service (user $SERVICE_USER)"

# ------------------------------------------------------------------ system
section "1. System"
if [ "$(uname -s)" = "Linux" ]; then
  pass "Linux: $( (. /etc/os-release 2>/dev/null && echo "$PRETTY_NAME") || uname -r)"
else
  fail "This installer is for Linux with systemd (found $(uname -s)). On Windows use NSSM — see README."
fi
if [ -d /run/systemd/system ] && command -v systemctl >/dev/null 2>&1; then
  pass "systemd is running ($(systemctl --version | head -1))"
else
  fail "systemd is not running on this machine, so a .service cannot be used."
fi
if [ "$(id -u)" -eq 0 ]; then pass "Running as root (needed by the installer)"
else warn "Not running as root: some checks are limited and install-service.sh must be run with sudo."; fi

MEM_MB=$(awk '/MemTotal/ {print int($2/1024)}' /proc/meminfo 2>/dev/null || echo 0)
if [ "$MEM_MB" -ge 1024 ]; then pass "Memory: ${MEM_MB} MB"
else warn "Memory: ${MEM_MB} MB — 1 GB or more is recommended for large Excel imports."; fi

if command -v timedatectl >/dev/null 2>&1; then
  if timedatectl show -p NTPSynchronized --value 2>/dev/null | grep -q yes; then pass "Clock synchronised (NTP) — import and review dates will be correct"
  else warn "Clock not reported as NTP-synchronised — the dates in the audit log come from this clock."; fi
fi

# ------------------------------------------------------------------ node
section "2. Node.js"
if [ -z "$NODE_BIN" ] || [ ! -x "$NODE_BIN" ]; then
  fail "Node.js not found. Install Node.js 22 LTS system-wide (e.g. NodeSource: https://github.com/nodesource/distributions) or set NODE_BIN=/path/to/node."
else
  NODE_BIN="$(readlink -f "$NODE_BIN")"
  NODE_VER="$("$NODE_BIN" -p 'process.versions.node' 2>/dev/null || echo 0.0.0)"
  if "$NODE_BIN" -e "const [a,b]=process.versions.node.split('.').map(Number);process.exit(a>22||(a===22&&b>=13)?0:1)" 2>/dev/null; then
    pass "Node.js $NODE_VER at $NODE_BIN (needs >= $MIN_NODE)"
  else
    fail "Node.js $NODE_VER at $NODE_BIN is too old (needs >= $MIN_NODE). Install the current LTS."
  fi
  if "$NODE_BIN" -e "process.emitWarning=()=>{};require('node:sqlite')" 2>/dev/null; then pass "Built-in SQLite (node:sqlite) available"
  else fail "node:sqlite is not available in this Node.js build."; fi
  # The service user must be able to reach the node binary (nvm installs under /root or /home are private).
  blocked=""
  p="$NODE_BIN"
  while [ "$p" != "/" ]; do
    p="$(dirname "$p")"
    perm="$(stat -c %A "$p" 2>/dev/null || echo ----------)"
    [ "${perm:9:1}" = "x" ] || [ "${perm:9:1}" = "t" ] || blocked="$p"
  done
  if [ -n "$blocked" ]; then
    fail "The service user cannot reach $NODE_BIN ($blocked is private — typical of nvm). Install Node.js system-wide (/usr/bin/node) or set NODE_BIN to a system path."
  else
    pass "Node.js binary is reachable by the service user"
  fi
fi

# ------------------------------------------------------------------ application
section "3. Application"
missing=""
for f in index.html package.json server/server.js server/db.js server/backup.js assets/js/core.js assets/js/app.js assets/js/docx-report.js assets/js/charts.js assets/css/styles.css; do
  [ -f "$REPO_DIR/$f" ] || missing="$missing $f"
done
[ -n "$(ls "$REPO_DIR"/vendor/*.js 2>/dev/null)" ] || missing="$missing vendor/*.js"
if [ -z "$missing" ]; then pass "All application files present"; else fail "Missing files:$missing (run git pull / check the copy)"; fi
if command -v git >/dev/null 2>&1 && [ -d "$REPO_DIR/.git" ]; then
  info "Version: $(git -C "$REPO_DIR" log -1 --format='%h %cd' --date=short 2>/dev/null) on branch $(git -C "$REPO_DIR" rev-parse --abbrev-ref HEAD 2>/dev/null)"
  if [ -n "$(git -C "$REPO_DIR" status --porcelain --untracked-files=no 2>/dev/null)" ]; then warn "The code folder has local, uncommitted changes — they will be installed as they are."; fi
fi

TMP="$(mktemp -d)"
cleanup() { [ -n "${SRV_PID:-}" ] && kill "$SRV_PID" 2>/dev/null; rm -rf "$TMP"; }
trap cleanup EXIT

if [ -n "$NODE_BIN" ] && [ -x "$NODE_BIN" ] && [ -z "$missing" ]; then
  if [ "${SKIP_TESTS:-0}" = "1" ]; then info "Unit tests skipped (SKIP_TESTS=1)"
  elif (cd "$REPO_DIR" && "$NODE_BIN" --test tests/core.test.js tests/server.test.js >"$TMP/tests.log" 2>&1); then
    pass "Unit tests pass on this server ($(grep -E '^# pass' "$TMP/tests.log" | awk '{print $3}') tests)"
  else
    fail "Unit tests fail on this server — see the output below:"; tail -20 "$TMP/tests.log" | sed 's/^/        /'
  fi

  # Smoke test: start the real server on a free port with a throw-away database.
  TPORT="$("$NODE_BIN" -e "const s=require('net').createServer().listen(0,'127.0.0.1',()=>{console.log(s.address().port);s.close()})")"
  PORT="$TPORT" HOST=127.0.0.1 SLM_DB="$TMP/smoke.sqlite" "$NODE_BIN" "$REPO_DIR/server/server.js" >"$TMP/smoke.log" 2>&1 &
  SRV_PID=$!
  if "$NODE_BIN" -e "
    const url='http://127.0.0.1:$TPORT/';
    (async()=>{for(let i=0;i<50;i++){try{const h=await (await fetch(url+'api/health')).json();const p=await fetch(url);
      if(h.ok&&p.ok){console.log(h.version);process.exit(0)}}catch(e){} await new Promise(r=>setTimeout(r,200))}process.exit(1)})()" >"$TMP/smoke.out" 2>&1; then
    pass "Test start OK: page and database API answer (version $(cat "$TMP/smoke.out"))"
  else
    fail "The server did not start in the test run:"; sed 's/^/        /' "$TMP/smoke.log" | tail -10
  fi
  kill "$SRV_PID" 2>/dev/null; wait "$SRV_PID" 2>/dev/null; SRV_PID=
fi

# ------------------------------------------------------------------ port & running copies
section "4. Port and running copies"
SERVICE_ACTIVE=0
if command -v systemctl >/dev/null 2>&1 && systemctl is-active --quiet "$SERVICE_NAME" 2>/dev/null; then SERVICE_ACTIVE=1; fi
manual=""
for d in /proc/[0-9]*; do
  cmd="$(tr '\0' ' ' <"$d/cmdline" 2>/dev/null)" || continue
  case "$cmd" in *server/server.js*) ;; *) continue ;; esac
  case "$(basename "${cmd%% *}")" in node*) ;; *) continue ;; esac   # the node process itself, not npm/sh wrappers
  pid="${d#/proc/}"
  if grep -q "$SERVICE_NAME.service" "$d/cgroup" 2>/dev/null; then continue; fi
  manual="$manual $pid"
  info "Manually started copy running: PID $pid ($cmd) in $(readlink "$d/cwd" 2>/dev/null)"
done
if [ -n "$manual" ]; then warn "A manually started server is running (PID$manual). The installer will stop it (after saving its database) so the service can take over."; fi

if [ -n "$NODE_BIN" ] && [ -x "$NODE_BIN" ]; then
  if "$NODE_BIN" -e "require('net').createServer().once('error',()=>process.exit(1)).listen($PORT,'$HOST',function(){this.close()})" 2>/dev/null; then
    pass "Port $PORT is free"
  elif [ "$SERVICE_ACTIVE" = "1" ]; then
    pass "Port $PORT is used by $SERVICE_NAME.service (already installed — the installer will update it)"
  elif [ -n "$manual" ]; then
    warn "Port $PORT is used — probably by the manual copy above, which the installer stops."
  elif [ "$PORT" -lt 1024 ] && [ "$(id -u)" -ne 0 ]; then
    info "Port $PORT is below 1024: re-run as root to test it (the service gets the permission to bind it)."
  else
    fail "Port $PORT is used by another program. Choose another one: PORT=8081 sudo -E bash deploy/install-service.sh"
  fi
fi

# ------------------------------------------------------------------ data
section "5. Data"
OLD_DB="${OLD_DB:-}"
if [ -z "$OLD_DB" ]; then
  cands=("$REPO_DIR/data/slm-informes.sqlite")
  for p in $manual; do
    cwd="$(readlink "/proc/$p/cwd" 2>/dev/null)"
    envdb="$(tr '\0' '\n' <"/proc/$p/environ" 2>/dev/null | sed -n 's/^SLM_DB=//p')"
    if [ -n "$envdb" ]; then case "$envdb" in /*) cands+=("$envdb") ;; *) cands+=("$cwd/$envdb") ;; esac; fi
    cands+=("$cwd/data/slm-informes.sqlite")
  done
  for c in "${cands[@]}"; do [ -f "$c" ] && { OLD_DB="$c"; break; }; done
fi
NEW_DB="$DATA_DIR/slm-informes.sqlite"
if [ -n "$OLD_DB" ] && [ -f "$OLD_DB" ]; then
  if out="$(SLM_DB="$OLD_DB" "$NODE_BIN" "$REPO_DIR/server/backup.js" --check 2>&1)"; then
    pass "Existing database found — it will be moved to $DATA_DIR: $out"
  else
    fail "Existing database $OLD_DB could not be read: $out"
  fi
else
  info "No existing database in the code folder (set OLD_DB=/path/file.sqlite if it is elsewhere)."
fi
if [ -f "$NEW_DB" ]; then
  if out="$(SLM_DB="$NEW_DB" "$NODE_BIN" "$REPO_DIR/server/backup.js" --check 2>&1)"; then pass "Service database already exists and will be kept: $out"
  else fail "Service database $NEW_DB is damaged: $out"; fi
  if [ -n "$OLD_DB" ] && [ -f "$OLD_DB" ] && [ "$(readlink -f "$OLD_DB")" != "$(readlink -f "$NEW_DB")" ]; then
    warn "Both $OLD_DB and $NEW_DB exist: the installer keeps $NEW_DB and does NOT overwrite it (set FORCE_MIGRATE=1 to replace it; a backup is made first)."
  fi
fi

probe="$DATA_DIR"; while [ ! -d "$probe" ]; do probe="$(dirname "$probe")"; done
FREE_MB=$(df -Pm "$probe" | awk 'NR==2 {print $4}')
DB_MB=0; [ -n "$OLD_DB" ] && [ -f "$OLD_DB" ] && DB_MB=$(( $(stat -c %s "$OLD_DB") / 1048576 ))
if [ "$FREE_MB" -lt 200 ] || [ "$FREE_MB" -lt $((DB_MB * 3)) ]; then fail "Only ${FREE_MB} MB free on $(df -P "$probe" | awk 'NR==2 {print $6}') for the database and its backups."
elif [ "$FREE_MB" -lt 1024 ]; then warn "${FREE_MB} MB free for data and backups — 1 GB or more is recommended."
else pass "${FREE_MB} MB free for data and daily backups ($(df -P "$probe" | awk 'NR==2 {print $6}'))"; fi
case "$DATA_DIR" in
  "$REPO_DIR"*|"$INSTALL_DIR"*) fail "DATA_DIR must be outside the code folders, otherwise updates can delete it." ;;
  /tmp*|/var/tmp*) fail "DATA_DIR must not be a temporary folder." ;;
  *) pass "Data folder is separate from the code (updates and re-installs never touch it)" ;;
esac

# ------------------------------------------------------------------ restarts, updates, network
section "6. Restarts, automatic updates and network"
if [ -f "/etc/systemd/system/$SERVICE_NAME.service" ]; then info "$SERVICE_NAME.service already installed — the installer will update it."
else info "$SERVICE_NAME.service not installed yet."; fi
if getent passwd "$SERVICE_USER" >/dev/null 2>&1; then info "Service user '$SERVICE_USER' exists."; else info "Service user '$SERVICE_USER' will be created (no login, no home)."; fi
if [ -f /etc/apt/apt.conf.d/20auto-upgrades ] && grep -q 'Unattended-Upgrade "1"' /etc/apt/apt.conf.d/20auto-upgrades 2>/dev/null; then
  reboot="$(grep -hs 'Automatic-Reboot ' /etc/apt/apt.conf.d/* | grep -o '"[a-z]*"' | tail -1)"
  info "Automatic updates are ON (automatic reboot: ${reboot:-not set}). With the service enabled the tool starts again by itself after every update/reboot."
elif command -v dnf >/dev/null 2>&1 && systemctl is-enabled --quiet dnf-automatic.timer 2>/dev/null; then
  info "Automatic updates are ON (dnf-automatic). With the service enabled the tool starts again by itself after every update/reboot."
fi
if [ -f /etc/needrestart/needrestart.conf ]; then info "needrestart is installed: after library updates it may restart the service — that is safe (it comes back in seconds, data is on disk)."; fi
if command -v ufw >/dev/null 2>&1 && ufw status 2>/dev/null | grep -q 'Status: active'; then
  if ufw status | grep -qE "^$PORT(/tcp)?\b"; then pass "Firewall (ufw) allows port $PORT"
  else warn "Firewall (ufw) is active and port $PORT is not open yet: sudo ufw allow from <office-network> to any port $PORT proto tcp"; fi
elif command -v firewall-cmd >/dev/null 2>&1 && firewall-cmd --state >/dev/null 2>&1; then
  if firewall-cmd --list-ports 2>/dev/null | grep -qw "$PORT/tcp"; then pass "Firewall (firewalld) allows port $PORT"
  else warn "Firewall (firewalld) is active and port $PORT is not open yet: sudo firewall-cmd --permanent --add-port=$PORT/tcp && sudo firewall-cmd --reload"; fi
else
  info "No local firewall detected — make sure only your network/VPN can reach port $PORT."
fi
if command -v getenforce >/dev/null 2>&1 && [ "$(getenforce 2>/dev/null)" = "Enforcing" ]; then
  warn "SELinux is enforcing: if the service cannot bind the port, run: sudo semanage port -a -t http_port_t -p tcp $PORT"
fi
if [ -z "${SLM_BASIC_AUTH:-}" ] && ! grep -qs '^SLM_BASIC_AUTH=.\+' "/etc/$SERVICE_NAME/$SERVICE_NAME.env"; then
  info "No login configured (SLM_BASIC_AUTH). Fine on a protected network; it can be added later in /etc/$SERVICE_NAME/$SERVICE_NAME.env"
fi

# ------------------------------------------------------------------ summary
echo
if [ "$FAILS" -gt 0 ]; then
  echo "${R}${B}$FAILS check(s) FAILED${N}, $WARNS warning(s). Fix the FAIL items before installing."
  exit 1
fi
echo "${G}${B}Ready to install${N} ($WARNS warning(s)). Next: sudo bash deploy/install-service.sh"
exit 0
