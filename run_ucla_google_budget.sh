#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

while [[ -z "${GOOGLE_MAPS_API_KEY:-}" ]]; do
  printf "Paste your Google Maps API key, then press Enter: "
  IFS= read -r GOOGLE_MAPS_API_KEY
  GOOGLE_MAPS_API_KEY="${GOOGLE_MAPS_API_KEY//[[:space:]]/}"
  if [[ ! "$GOOGLE_MAPS_API_KEY" =~ ^AIza[0-9A-Za-z_-]{30,}$ ]]; then
    printf "That does not look like a Google Maps API key. Please paste only the key, not an error message or command output.\n"
    GOOGLE_MAPS_API_KEY=""
  fi
done
export GOOGLE_MAPS_API_KEY

python3 find_apartments_google_budget.py \
  --school "UCLA" \
  --center-lat 34.0703 \
  --center-lon -118.4448 \
  --radius-meters 2200 \
  --max-walk-minutes 20 \
  --budget-usd 5 \
  --osm-markdown apartments_osm.md \
  --seed "Lindbrook Manor" \
  --area "Westwood" \
  --area "Westwood Village" \
  --area "Wilshire Corridor" \
  --skip-routes \
  --output apartments_google_budget.md

printf "\nDone. Output files:\n"
printf "  apartments_google_budget.md\n"
printf "  apartments_google_budget.csv\n"
