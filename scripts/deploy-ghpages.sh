#!/usr/bin/env bash
# Build Plott's static export and publish it to the `gh-pages` branch, which
# GitHub Pages serves at https://alexselig.github.io/plott/ (install page +
# add-in manifest included).
#
#   scripts/deploy-ghpages.sh
#
# Requires the personal `alexselig` gh token (repo scope). Re-run after any change
# you want live. (An Actions workflow would automate this, but pushing workflow
# files needs a `workflow`-scoped token.)
set -euo pipefail
cd "$(dirname "$0")/.."

ORIGIN="https://alexselig.github.io/plott"
REPO="https://github.com/alexselig/plott.git"

# Static export can't include the server-only AI API route — move it aside.
API_MOVED=0
if [ -d src/app/api ]; then mv src/app/api /tmp/plott_api_bak && API_MOVED=1; fi
restore() { if [ "$API_MOVED" = 1 ] && [ -d /tmp/plott_api_bak ]; then mv /tmp/plott_api_bak src/app/api; fi; }
trap restore EXIT

echo "Building static export (basePath=/plott)..." >&2
rm -rf out
VIBEHUB_EXPORT=1 VIBEHUB_BASE_PATH=/plott npx next build >/tmp/plott_ghpages_build.log 2>&1 || {
  echo "BUILD FAILED:" >&2; tail -25 /tmp/plott_ghpages_build.log >&2; exit 1;
}
test -f out/index.html || { echo "ERROR: no out/index.html" >&2; exit 1; }

# Point the hosted manifest at the Pages origin so "Upload My Add-in" works.
node scripts/make-manifest.mjs "$ORIGIN" out/manifest.xml
cp out/manifest.xml out/manifest.ghpages.xml
touch out/.nojekyll   # let Pages serve the _next/ assets

echo "Publishing ./out to gh-pages..." >&2
TK="$(gh auth token --user alexselig)"
HELPER='!f(){ echo username=alexselig; echo "password='"$TK"'"; }; f'
(
  cd out
  rm -rf .git
  git init -q
  git checkout -q -b gh-pages
  git add -A
  git -c user.name=alexselig -c user.email=32137968+alexselig@users.noreply.github.com \
    commit -qm "Deploy $(date -u +%Y-%m-%dT%H:%M:%SZ)"
  git -c credential.helper= -c credential.helper="$HELPER" push -f "$REPO" gh-pages
)
echo "Deployed. Live at ${ORIGIN}/install/" >&2
