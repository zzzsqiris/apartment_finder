#!/usr/bin/env python3
"""Cross-platform interactive runner for the apartment finder pipeline."""

from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path


API_KEY_RE = re.compile(r"^AIza[0-9A-Za-z_-]{30,}$")


def main() -> int:
    os.chdir(Path(__file__).resolve().parent)

    api_key = get_api_key()
    search_place = prompt_default("School, address, or place to search around", "UCLA")
    center_lat = prompt_optional("Optional center latitude; press Enter to let Google find the place")
    center_lon = ""
    if center_lat.strip():
        center_lon = prompt_required("Center longitude")
    radius_meters = prompt_default("Search radius in meters", "2200")
    max_walk_minutes = prompt_default("Max walking minutes; used only if routes are enabled", "20")
    rent_budget = prompt_default("Rent budget for screening", "2500")
    target_units = prompt_default("Acceptable unit types, comma-separated", "studio,1b1b")
    private_bath_mode = prompt_choice("Private bathroom preference", "any", {"any", "prefer", "require"})
    api_budget = prompt_default("Google API budget cap for this run", "5")
    skip_routes = prompt_yes_no("Skip route calculation to save money?", default=True)
    areas = split_csv(prompt_optional("Optional neighborhood/commercial areas, comma-separated; press Enter to skip"))
    seeds = split_csv(prompt_optional("Optional must-check apartment names, comma-separated; press Enter to skip"))

    slug = sanitize_slug(search_place) or "apartment_search"
    raw_md = f"{slug}_apartments_google.md"
    clean_md = f"{slug}_apartments_clean.md"
    clean_csv = f"{slug}_apartments_clean.csv"
    screen_csv = f"{slug}_screening.csv"
    screen_xlsx = f"{slug}_screening.xlsx"
    screen_docx = f"{slug}_screening_summary.docx"

    google_args = [
        "--school",
        search_place,
        "--radius-meters",
        radius_meters,
        "--max-walk-minutes",
        max_walk_minutes,
        "--budget-usd",
        api_budget,
        "--output",
        raw_md,
    ]
    if center_lat.strip() and center_lon.strip():
        google_args.extend(["--center-lat", center_lat, "--center-lon", center_lon])
    if skip_routes:
        google_args.append("--skip-routes")
    for area in areas:
        google_args.extend(["--area", area])
    for seed in seeds:
        google_args.extend(["--seed", seed])

    env = os.environ.copy()
    env["GOOGLE_MAPS_API_KEY"] = api_key

    print(f"\nStep 1/3: searching apartments around {search_place}...")
    run_python("find_apartments_google_budget.py", google_args, env)

    print("\nStep 2/3: cleaning search results...")
    run_python(
        "clean_google_budget_results.py",
        [
            "--input",
            str(Path(raw_md).with_suffix(".csv")),
            "--output-md",
            clean_md,
            "--output-csv",
            clean_csv,
        ],
        env,
    )

    print(f"\nStep 3/3: screening apartment websites with rent budget ${rent_budget}...")
    run_python(
        "screen_apartments_from_web.py",
        [
            "--input",
            clean_csv,
            "--budget",
            rent_budget,
            "--target-units",
            target_units,
            "--private-bath",
            private_bath_mode,
            "--output-csv",
            screen_csv,
            "--output-xlsx",
            screen_xlsx,
            "--output-docx",
            screen_docx,
        ],
        env,
    )

    print("\nDone. Output files:")
    for path in [raw_md, clean_md, clean_csv, screen_csv, screen_xlsx, screen_docx]:
        print(f"  {path}")
    return 0


def get_api_key() -> str:
    api_key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    while not API_KEY_RE.match(api_key):
        api_key = input("Paste your Google Maps API key, then press Enter: ").strip()
        if not API_KEY_RE.match(api_key):
            print("That does not look like a Google Maps API key. Please paste only the key.")
    return api_key


def prompt_default(prompt: str, default: str) -> str:
    value = input(f"{prompt} [{default}]: ").strip()
    return value or default


def prompt_optional(prompt: str) -> str:
    return input(f"{prompt}: ").strip()


def prompt_required(prompt: str) -> str:
    while True:
        value = prompt_optional(prompt)
        if value:
            return value
        print("Please enter a value.")


def prompt_choice(prompt: str, default: str, choices: set[str]) -> str:
    while True:
        value = prompt_default(f"{prompt} ({'/'.join(sorted(choices))})", default).lower()
        if value in choices:
            return value
        print(f"Please enter one of: {', '.join(sorted(choices))}.")


def prompt_yes_no(prompt: str, default: bool) -> bool:
    suffix = "Y/n" if default else "y/N"
    while True:
        value = input(f"{prompt} [{suffix}]: ").strip().lower()
        if not value:
            return default
        if value in {"y", "yes"}:
            return True
        if value in {"n", "no"}:
            return False
        print("Please enter yes or no.")


def split_csv(value: str) -> list[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def sanitize_slug(value: str) -> str:
    value = re.sub(r"[^a-z0-9]+", "_", value.lower())
    value = value.strip("_")
    return value[:60]


def run_python(script: str, args: list[str], env: dict[str, str]) -> None:
    command = [sys.executable, script, *args]
    subprocess.run(command, check=True, env=env)


if __name__ == "__main__":
    raise SystemExit(main())
