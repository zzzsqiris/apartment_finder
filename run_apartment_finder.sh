#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

prompt_default() {
  local prompt="$1"
  local default="$2"
  local value
  printf "%s [%s]: " "$prompt" "$default" >&2
  IFS= read -r value
  if [[ -z "${value//[[:space:]]/}" ]]; then
    printf "%s" "$default"
  else
    printf "%s" "$value"
  fi
}

prompt_optional() {
  local prompt="$1"
  local value
  printf "%s: " "$prompt" >&2
  IFS= read -r value
  printf "%s" "$value"
}

sanitize_slug() {
  printf "%s" "$1" \
    | tr '[:upper:]' '[:lower:]' \
    | sed -E 's/[^a-z0-9]+/_/g; s/^_+//; s/_+$//' \
    | cut -c 1-60
}

while [[ -z "${GOOGLE_MAPS_API_KEY:-}" ]]; do
  printf "Paste your Google Maps API key, then press Enter: "
  IFS= read -r GOOGLE_MAPS_API_KEY
  GOOGLE_MAPS_API_KEY="${GOOGLE_MAPS_API_KEY//[[:space:]]/}"
  if [[ ! "$GOOGLE_MAPS_API_KEY" =~ ^AIza[0-9A-Za-z_-]{30,}$ ]]; then
    printf "That does not look like a Google Maps API key. Please paste only the key.\n"
    GOOGLE_MAPS_API_KEY=""
  fi
done
export GOOGLE_MAPS_API_KEY

search_place="$(prompt_default "School, address, or place to search around" "UCLA")"
center_lat="$(prompt_optional "Optional center latitude; press Enter to let Google find the place")"
center_lon=""
if [[ -n "${center_lat//[[:space:]]/}" ]]; then
  center_lon="$(prompt_optional "Center longitude")"
fi
radius_meters="$(prompt_default "Search radius in meters" "2200")"
max_walk_minutes="$(prompt_default "Max walking minutes; used only if routes are enabled" "20")"
rent_budget="$(prompt_default "Rent budget for screening" "2500")"
target_units="$(prompt_default "Acceptable unit types, comma-separated" "studio,1b1b")"
private_bath_mode="$(prompt_default "Private bathroom preference: any/prefer/require" "any")"
api_budget="$(prompt_default "Google API budget cap for this run" "5")"
skip_routes_answer="$(prompt_default "Skip route calculation to save money? yes/no" "yes")"
areas_csv="$(prompt_optional "Optional neighborhood/commercial areas, comma-separated; press Enter to skip")"
seeds_csv="$(prompt_optional "Optional must-check apartment names, comma-separated; press Enter to skip")"

slug="$(sanitize_slug "$search_place")"
if [[ -z "$slug" ]]; then
  slug="apartment_search"
fi

raw_md="${slug}_apartments_google.md"
clean_md="${slug}_apartments_clean.md"
clean_csv="${slug}_apartments_clean.csv"
screen_csv="${slug}_screening.csv"
screen_xlsx="${slug}_screening.xlsx"
screen_docx="${slug}_screening_summary.docx"

google_args=(
  --school "$search_place"
  --radius-meters "$radius_meters"
  --max-walk-minutes "$max_walk_minutes"
  --budget-usd "$api_budget"
  --output "$raw_md"
)

if [[ -n "${center_lat//[[:space:]]/}" && -n "${center_lon//[[:space:]]/}" ]]; then
  google_args+=(--center-lat "$center_lat" --center-lon "$center_lon")
fi

if [[ "$skip_routes_answer" =~ ^[Yy] ]]; then
  google_args+=(--skip-routes)
fi

IFS=',' read -ra areas <<< "$areas_csv"
for area in "${areas[@]}"; do
  area="${area#"${area%%[![:space:]]*}"}"
  area="${area%"${area##*[![:space:]]}"}"
  if [[ -n "$area" ]]; then
    google_args+=(--area "$area")
  fi
done

IFS=',' read -ra seeds <<< "$seeds_csv"
for seed in "${seeds[@]}"; do
  seed="${seed#"${seed%%[![:space:]]*}"}"
  seed="${seed%"${seed##*[![:space:]]}"}"
  if [[ -n "$seed" ]]; then
    google_args+=(--seed "$seed")
  fi
done

printf "\nStep 1/3: searching apartments around %s...\n" "$search_place"
python3 find_apartments_google_budget.py "${google_args[@]}"

printf "\nStep 2/3: cleaning search results...\n"
python3 clean_google_budget_results.py \
  --input "${raw_md%.md}.csv" \
  --output-md "$clean_md" \
  --output-csv "$clean_csv"

printf "\nStep 3/3: screening apartment websites with rent budget $%s...\n" "$rent_budget"
python3 screen_apartments_from_web.py \
  --input "$clean_csv" \
  --budget "$rent_budget" \
  --target-units "$target_units" \
  --private-bath "$private_bath_mode" \
  --output-csv "$screen_csv" \
  --output-xlsx "$screen_xlsx" \
  --output-docx "$screen_docx"

printf "\nDone. Output files:\n"
printf "  %s\n" "$raw_md"
printf "  %s\n" "$clean_md"
printf "  %s\n" "$clean_csv"
printf "  %s\n" "$screen_csv"
printf "  %s\n" "$screen_xlsx"
printf "  %s\n" "$screen_docx"
