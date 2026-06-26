#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

if ! command -v clasp >/dev/null 2>&1; then
  echo "clasp is not installed. Install it with: npm install -g @google/clasp"
  exit 1
fi

if [[ ! -f .clasp.json ]]; then
  echo "No .clasp.json found. Run clasp create first."
  exit 1
fi

PARENT_ID="$(node -e "const fs=require('fs'); const cfg=JSON.parse(fs.readFileSync('.clasp.json','utf8')); if(!cfg.parentId){process.exit(1)} process.stdout.write(cfg.parentId)")" || {
  echo "No parentId in .clasp.json. This project must be bound to a spreadsheet."
  exit 1
}

SCRIPT_ID="$(node -e "const fs=require('fs'); const cfg=JSON.parse(fs.readFileSync('.clasp.json','utf8')); process.stdout.write(cfg.scriptId)")"

echo "Pushing latest code..."
clasp push --force

echo "Running setupSheet for spreadsheet $PARENT_ID..."
if clasp run setupSheet --params "[\"$PARENT_ID\"]" 2>/dev/null; then
  echo "Done."
else
  echo ""
  echo "Remote setup via clasp run is not available yet."
  echo "Run setup manually using either option:"
  echo ""
  echo "Option A - Spreadsheet menu:"
  echo "  1. Open https://drive.google.com/open?id=$PARENT_ID"
  echo "  2. Refresh the page if needed"
  echo "  3. Click Quiz > Set up sheets"
  echo ""
  echo "Option B - Script editor:"
  echo "  1. Open https://script.google.com/d/$SCRIPT_ID/edit"
  echo "  2. Select setupSheet from the function dropdown"
  echo "  3. Click Run"
  echo ""
  echo "To enable clasp run later:"
  echo "  - Turn on the Google Apps Script API at https://script.google.com/home/usersettings"
  echo "  - Deploy the script as an API executable from the script editor"
  exit 0
fi

echo "Open your spreadsheet:"
echo "https://drive.google.com/open?id=$PARENT_ID"
