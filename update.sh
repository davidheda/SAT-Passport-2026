#!/bin/bash
# One-command passport update: rebuild pages + QR codes, then push to GitHub Pages.
# First time: create the repo (private is fine, Pages stays public at unguessable URLs),
#   create the PRIVATE repo SAT-Passport-2026 on github.com (empty, no README), then
#   git remote add origin https://github.com/davidheda/SAT-Passport-2026.git && git push -u origin main
#   Settings > Pages > Deploy from a branch > main  /docs
set -e
cd "$(dirname "$0")"
python3 build_passport.py
if [ -d .git ]; then
  git add -A
  git commit -m "Passport update $(date +%Y-%m-%d)" || true
  git push
else
  echo "No git repository here yet — pages built locally in site/ (see comments in update.sh)."
fi
