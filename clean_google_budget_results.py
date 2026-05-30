#!/usr/bin/env python3
"""Clean Google apartment search results without making any API calls."""

from __future__ import annotations

import argparse
import csv
from pathlib import Path


EXCLUDE_TERMS = (
    "ucla health",
    "ucla store",
    "rotary club",
    "birthplace",
    "primary care",
    "sunset village",
    "rieber",
    "de neve",
    "weyburn commons",
    "furnished corporate housing",
    "corporate housing",
    "nms furnished",
    "mysuite",
    "hotel",
    "real estate",
    "improvement association",
    "village square",
    "plaza at westwood",
    "westwood apartment homes",
)

KEEP_TERMS = (
    "apartment",
    "apartments",
    "condo",
    "condominiums",
    "court",
    "lofts",
    "manor",
    "terrace",
    "towers",
    "villa",
    "villas",
    "residence",
    "residences",
    "regency",
    "house",
    "heights",
    "lindbrook",
    "gayley",
    "kelton",
    "strathmore",
    "wilshire",
    "midvale",
    "veteran",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="apartments_google_budget.csv")
    parser.add_argument("--output-md", default="apartments_google_budget_clean.md")
    parser.add_argument("--output-csv", default="apartments_google_budget_clean.csv")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rows = read_rows(args.input)
    cleaned = [row for row in rows if should_keep(row)]
    cleaned = dedupe(cleaned)
    cleaned.sort(key=lambda row: (distance(row), row["name"].lower()))
    write_csv(args.output_csv, cleaned)
    write_markdown(args.output_md, cleaned)
    print(f"Wrote {args.output_md} and {args.output_csv} with {len(cleaned)} rows.")
    return 0


def read_rows(path: str) -> list[dict[str, str]]:
    with open(path, newline="", encoding="utf-8") as file:
        return list(csv.DictReader(file))


def should_keep(row: dict[str, str]) -> bool:
    haystack = f"{row.get('name', '')} {row.get('address', '')} {row.get('website_uri', '')}".lower()
    if any(term in haystack for term in EXCLUDE_TERMS):
        return False
    return any(term in haystack for term in KEEP_TERMS)


def dedupe(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    seen: set[tuple[str, str]] = set()
    result = []
    for row in rows:
        key = (row.get("name", "").lower(), row.get("address", "").lower())
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def distance(row: dict[str, str]) -> int:
    value = row.get("straight_line_meters") or row.get("walk_distance_meters") or "999999"
    try:
        return int(float(value))
    except ValueError:
        return 999999


def group_name(row: dict[str, str]) -> str:
    text = f"{row.get('name', '')} {row.get('address', '')}".lower()
    if "strathmore" in text or "gayley" in text or "landfair" in text:
        return "North Westwood / Strathmore-Gayley"
    if "weyburn" in text or "levering" in text:
        return "Westwood Village / Weyburn-Levering"
    if "lindbrook" in text or "glendon" in text or "hilgard" in text or "tiverton" in text:
        return "Westwood Village Core"
    if "wilshire" in text:
        return "Wilshire Corridor"
    if "kelton" in text or "midvale" in text or "veteran" in text or "rochester" in text or "wellworth" in text or "ohio" in text:
        return "Southwest Westwood"
    return "Other Westwood"


def write_csv(path: str, rows: list[dict[str, str]]) -> None:
    fields = [
        "name",
        "clean_group",
        "address",
        "straight_line_meters",
        "google_maps_uri",
        "website_uri",
        "sources",
    ]
    with open(path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "name": row.get("name", ""),
                    "clean_group": group_name(row),
                    "address": row.get("address", ""),
                    "straight_line_meters": row.get("straight_line_meters", ""),
                    "google_maps_uri": row.get("google_maps_uri", ""),
                    "website_uri": row.get("website_uri", ""),
                    "sources": row.get("sources", ""),
                }
            )


def write_markdown(path: str, rows: list[dict[str, str]]) -> None:
    groups: dict[str, list[dict[str, str]]] = {}
    for row in rows:
        groups.setdefault(group_name(row), []).append(row)
    for items in groups.values():
        items.sort(key=lambda row: (distance(row), row["name"].lower()))

    lines = [
        "# Cleaned Google-budget apartments near UCLA",
        "",
        f"- Places shown: {len(rows)}",
        "- Distance: straight-line from UCLA center, not walking route",
        "- Source: cleaned locally from apartments_google_budget.csv; no extra Google API calls",
        "",
    ]
    for group, items in sorted(groups.items(), key=lambda item: distance(item[1][0])):
        lines.extend(
            [
                f"## {escape(group)}",
                "",
                "| # | Apartment | Google Maps | Website | Address | Distance | Sources |",
                "|---:|---|---|---|---|---:|---|",
            ]
        )
        for index, row in enumerate(items, start=1):
            lines.append(
                f"| {index} | {escape(row.get('name', ''))} | "
                f"{link('map', row.get('google_maps_uri', ''))} | "
                f"{link('website', row.get('website_uri', ''))} | "
                f"{escape(row.get('address', ''))} | "
                f"{format_meters(distance(row))} | "
                f"{escape(row.get('sources', ''))} |"
            )
        lines.append("")
    Path(path).write_text("\n".join(lines), encoding="utf-8")


def link(label: str, url: str) -> str:
    return f"[{label}]({url})" if url else ""


def format_meters(meters: int) -> str:
    if meters >= 1000:
        return f"{meters / 1000:.1f} km"
    return f"{meters} m"


def escape(value: str) -> str:
    return value.replace("|", "\\|")


if __name__ == "__main__":
    raise SystemExit(main())
