#!/usr/bin/env bash
# Lightweight runner for the Python seeding workflow.
# Usage:
#   SUPABASE_URL="https://<project>.supabase.co" \
#   SUPABASE_SERVICE_ROLE_KEY="<service-role-key>" \
#   bash career-pathway-navigator-215219/career_navigator_frontend/scripts/run_python_seed.sh [--only roles,competencies,...] [--log-level INFO]
#
# Notes:
# - This script does not read .env files directly. Ensure environment variables are exported in your shell.
# - Falls back to REACT_APP_* variable names if SUPABASE_* are not set.

set -euo pipefail

# Compute repo root and seed script path robustly
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
SEED_PY="${ROOT_DIR}/career_navigator_frontend/scripts/seed.py"
REQ_TXT="${ROOT_DIR}/requirements.txt"

# Environment fallbacks (do not read .env; just map existing React-style vars if present)
if [[ -z "${SUPABASE_URL:-}" && -n "${REACT_APP_SUPABASE_URL:-}" ]]; then
  export SUPABASE_URL="${REACT_APP_SUPABASE_URL}"
fi
if [[ -z "${SUPABASE_SERVICE_ROLE_KEY:-}" ]]; then
  if [[ -n "${REACT_APP_SUPABASE_SERVICE_ROLE_KEY:-}" ]]; then
    export SUPABASE_SERVICE_ROLE_KEY="${REACT_APP_SUPABASE_SERVICE_ROLE_KEY}"
  elif [[ -n "${REACT_APP_SUPABASE_KEY:-}" ]]; then
    export SUPABASE_SERVICE_ROLE_KEY="${REACT_APP_SUPABASE_KEY}"
  fi
fi

# Validate required env
if [[ -z "${SUPABASE_URL:-}" ]]; then
  echo "[seed:runner] ERROR: SUPABASE_URL is not set. Please export SUPABASE_URL (or REACT_APP_SUPABASE_URL)." >&2
  exit 1
fi
if [[ -z "${SUPABASE_SERVICE_ROLE_KEY:-}" ]]; then
  echo "[seed:runner] ERROR: SUPABASE_SERVICE_ROLE_KEY is not set. Please export SUPABASE_SERVICE_ROLE_KEY (or REACT_APP_SUPABASE_SERVICE_ROLE_KEY / REACT_APP_SUPABASE_KEY)." >&2
  exit 1
fi

echo "[seed:runner] Installing Python requirements from ${REQ_TXT} ..."
python3 -m pip install --no-input -r "${REQ_TXT}"

echo "[seed:runner] Running Python seeding utility ..."
python3 "${SEED_PY}" --log-level INFO "$@"
