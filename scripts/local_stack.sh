#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

usage() {
  cat <<'USAGE'
Usage:
  scripts/local_stack.sh start  prod|dev
  scripts/local_stack.sh stop   prod|dev
  scripts/local_stack.sh status prod|dev

Fixed local environments:
  prod: frontend 3001, backend 8001, jobs runtime/prod/jobs
  dev : frontend 3002, backend 8002, jobs runtime/dev/jobs

prod は backend を reload なし、frontend を generate + preview で起動します。
dev は backend reload と Nuxt dev server で起動します。
USAGE
}

if [[ $# -ne 2 ]]; then
  usage
  exit 2
fi

ACTION="$1"
ENV_NAME="$2"

case "$ENV_NAME" in
  prod)
    FRONTEND_PORT=3001
    BACKEND_PORT=8001
    BACKEND_RELOAD=0
    FRONTEND_MODE="preview"
    ;;
  dev)
    FRONTEND_PORT=3002
    BACKEND_PORT=8002
    BACKEND_RELOAD=1
    FRONTEND_MODE="dev"
    ;;
  *)
    usage
    exit 2
    ;;
esac

RUNTIME_DIR="$ROOT_DIR/runtime/$ENV_NAME"
JOBS_DIR="$RUNTIME_DIR/jobs"
LOG_DIR="$RUNTIME_DIR/logs"
PID_DIR="$RUNTIME_DIR/pids"
BACKEND_PID="$PID_DIR/backend.pid"
FRONTEND_PID="$PID_DIR/frontend.pid"
BACKEND_LOG="$LOG_DIR/backend.log"
FRONTEND_LOG="$LOG_DIR/frontend.log"
BACKEND_URL="http://127.0.0.1:$BACKEND_PORT"
CORS_ORIGINS="http://localhost:$FRONTEND_PORT,http://127.0.0.1:$FRONTEND_PORT"
FRONTEND_BASE_URL="/"

ensure_dirs() {
  mkdir -p "$JOBS_DIR" "$LOG_DIR" "$PID_DIR"
}

pid_alive() {
  local pid_file="$1"
  [[ -f "$pid_file" ]] || return 1
  local pid
  pid="$(cat "$pid_file")"
  [[ -n "$pid" ]] || return 1
  kill -0 "$pid" >/dev/null 2>&1
}

port_pid() {
  local port="$1"
  lsof -nP -t -iTCP:"$port" -sTCP:LISTEN 2>/dev/null | head -1 || true
}

assert_port_free() {
  local port="$1"
  local label="$2"
  local pid
  pid="$(port_pid "$port")"
  if [[ -n "$pid" ]]; then
    echo "$label port $port is already in use by pid $pid" >&2
    exit 1
  fi
}

start_backend() {
  if pid_alive "$BACKEND_PID"; then
    echo "$ENV_NAME backend already running (pid $(cat "$BACKEND_PID"))"
    return
  fi
  assert_port_free "$BACKEND_PORT" "$ENV_NAME backend"
  : > "$BACKEND_LOG"
  if [[ "$BACKEND_RELOAD" == "1" ]]; then
    nohup bash -c \
      'cd "$1" && JOBS_DIR="$2" CORS_ALLOW_ORIGINS="$3" uv run uvicorn backend.main:app --reload --host 127.0.0.1 --port "$4"' \
      _ "$ROOT_DIR" "$JOBS_DIR" "$CORS_ORIGINS" "$BACKEND_PORT" >> "$BACKEND_LOG" 2>&1 &
  else
    nohup bash -c \
      'cd "$1" && JOBS_DIR="$2" CORS_ALLOW_ORIGINS="$3" uv run uvicorn backend.main:app --host 127.0.0.1 --port "$4"' \
      _ "$ROOT_DIR" "$JOBS_DIR" "$CORS_ORIGINS" "$BACKEND_PORT" >> "$BACKEND_LOG" 2>&1 &
  fi
  echo $! > "$BACKEND_PID"
  echo "$ENV_NAME backend started: $BACKEND_URL (pid $(cat "$BACKEND_PID"))"
}

start_frontend() {
  if pid_alive "$FRONTEND_PID"; then
    echo "$ENV_NAME frontend already running (pid $(cat "$FRONTEND_PID"))"
    return
  fi
  assert_port_free "$FRONTEND_PORT" "$ENV_NAME frontend"
  : > "$FRONTEND_LOG"
  if [[ "$FRONTEND_MODE" == "preview" ]]; then
    nohup bash -c \
      'cd "$1/frontend" && NUXT_PUBLIC_BACKEND_URL="$2" NUXT_PORT="$3" NUXT_APP_BASE_URL="$4" pnpm generate && NUXT_PUBLIC_BACKEND_URL="$2" NUXT_APP_BASE_URL="$4" pnpm preview --host 127.0.0.1 --port "$3"' \
      _ "$ROOT_DIR" "$BACKEND_URL" "$FRONTEND_PORT" "$FRONTEND_BASE_URL" >> "$FRONTEND_LOG" 2>&1 &
  else
    nohup bash -c \
      'cd "$1/frontend" && NUXT_PUBLIC_BACKEND_URL="$2" NUXT_PORT="$3" NUXT_APP_BASE_URL="$4" pnpm dev --host 127.0.0.1 --port "$3"' \
      _ "$ROOT_DIR" "$BACKEND_URL" "$FRONTEND_PORT" "$FRONTEND_BASE_URL" >> "$FRONTEND_LOG" 2>&1 &
  fi
  echo $! > "$FRONTEND_PID"
  echo "$ENV_NAME frontend starting: http://127.0.0.1:$FRONTEND_PORT (pid $(cat "$FRONTEND_PID"))"
}

stop_one() {
  local label="$1"
  local pid_file="$2"
  local port="$3"
  if ! [[ -f "$pid_file" ]]; then
    echo "$ENV_NAME $label not running (no pid file)"
    return
  fi
  local pid
  pid="$(cat "$pid_file")"
  if [[ -n "$pid" ]] && kill -0 "$pid" >/dev/null 2>&1; then
    pkill -TERM -P "$pid" >/dev/null 2>&1 || true
    kill "$pid" >/dev/null 2>&1 || true
    sleep 1
    if kill -0 "$pid" >/dev/null 2>&1; then
      pkill -KILL -P "$pid" >/dev/null 2>&1 || true
      kill -KILL "$pid" >/dev/null 2>&1 || true
    fi
    echo "$ENV_NAME $label stopped (pid $pid)"
  else
    echo "$ENV_NAME $label not running (stale pid $pid)"
  fi
  rm -f "$pid_file"
  local remaining_pid
  remaining_pid="$(port_pid "$port")"
  if [[ -n "$remaining_pid" ]]; then
    echo "$ENV_NAME $label warning: port $port is still used by pid $remaining_pid" >&2
  fi
}

status_one() {
  local label="$1"
  local pid_file="$2"
  local port="$3"
  if pid_alive "$pid_file"; then
    echo "$ENV_NAME $label: running pid $(cat "$pid_file"), port $port"
  else
    local pid
    pid="$(port_pid "$port")"
    if [[ -n "$pid" ]]; then
      echo "$ENV_NAME $label: port $port is used by unmanaged pid $pid"
    else
      echo "$ENV_NAME $label: stopped"
    fi
  fi
}

case "$ACTION" in
  start)
    ensure_dirs
    start_backend
    start_frontend
    echo "jobs: $JOBS_DIR"
    echo "logs: $LOG_DIR"
    ;;
  stop)
    stop_one frontend "$FRONTEND_PID" "$FRONTEND_PORT"
    stop_one backend "$BACKEND_PID" "$BACKEND_PORT"
    ;;
  status)
    status_one backend "$BACKEND_PID" "$BACKEND_PORT"
    status_one frontend "$FRONTEND_PID" "$FRONTEND_PORT"
    echo "jobs: $JOBS_DIR"
    echo "logs: $LOG_DIR"
    ;;
  *)
    usage
    exit 2
    ;;
esac
