#!/usr/bin/env bash
set -euo pipefail

python3 screen_apartments_from_web.py \
  --input apartments_google_budget_clean.csv \
  --output-csv apartment_screening_auto.csv \
  --output-xlsx apartment_screening_auto.xlsx \
  --output-docx apartment_screening_auto_summary.docx
