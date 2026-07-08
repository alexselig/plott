#!/usr/bin/env bash
# Deploy ChartForge to VibeHub (static export).
#   scripts/deploy-vibehub.sh "<basePath>" [extra -F fields...]
# Examples:
#   scripts/deploy-vibehub.sh "" "name=ChartForge" "description=..." "tags=..."
#   scripts/deploy-vibehub.sh "/app/<projectId>" "projectId=<projectId>"
set -euo pipefail
cd "$(dirname "$0")/.."

# shellcheck disable=SC1090
source ~/.env.vibehub
: "${VIBEHUB_API_KEY:?VIBEHUB_API_KEY not set in ~/.env.vibehub}"

BASE_PATH="${1-}"
shift || true

# Static export can't include the server-only API route — move it aside.
API_MOVED=0
if [ -d src/app/api ]; then mv src/app/api /tmp/cf_api_bak && API_MOVED=1; fi
restore() { if [ "$API_MOVED" = 1 ] && [ -d /tmp/cf_api_bak ]; then mv /tmp/cf_api_bak src/app/api; fi; }
trap restore EXIT

rm -rf out deploy.zip
echo "Building static export (basePath='${BASE_PATH}')..." >&2
if ! VIBEHUB_EXPORT=1 VIBEHUB_BASE_PATH="$BASE_PATH" npx next build >/tmp/cf_build.log 2>&1; then
  echo "BUILD FAILED:" >&2
  tail -25 /tmp/cf_build.log >&2
  exit 1
fi
test -f out/index.html || { echo "ERROR: no out/index.html" >&2; exit 1; }
( cd out && zip -qr ../deploy.zip . -x '.*' '__MACOSX/*' )
echo "Uploading $(du -h deploy.zip | cut -f1 | tr -d ' ') zip to VibeHub..." >&2

ARGS=(-s -X POST -H "X-API-Key: ${VIBEHUB_API_KEY}" -F "file=@deploy.zip")
for f in "$@"; do ARGS+=(-F "$f"); done
curl "${ARGS[@]}" https://vibehub.microsoft.com/api/external/push
echo
rm -f deploy.zip
