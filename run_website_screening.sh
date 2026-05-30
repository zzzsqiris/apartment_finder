#!/usr/bin/env bash
set -euo pipefail

python3 screen_apartments_from_web.py \
  --input apartments_google_budget_clean.csv \
  --target-units "${TARGET_UNITS:-studio,1b1b}" \
  --private-bath "${PRIVATE_BATH:-any}" \
  --budget "${RENT_BUDGET:-2500}" \
  --output-csv apartment_screening_auto.csv \
  --output-xlsx apartment_screening_auto.xlsx \
  --output-docx apartment_screening_auto_summary.docx
