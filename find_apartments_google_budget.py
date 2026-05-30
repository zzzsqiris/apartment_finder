#!/usr/bin/env python3
"""Budget-capped apartment finder using Google Maps APIs.

The goal is to spend as little as possible while improving recall and route
accuracy beyond the free OpenStreetMap pass. It combines:
- targeted Google Text Search queries
- optional names from an existing OSM Markdown output
- user-provided seed names such as "Lindbrook Manor"
- Google Routes route matrix for walking and biking
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any


PLACES_TEXT_SEARCH_URL = "https://places.googleapis.com/v1/places:searchText"
ROUTES_MATRIX_URL = "https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix"

DEFAULT_RADIUS_METERS = 2600
DEFAULT_MAX_WALK_MINUTES = 20
DEFAULT_BUDGET_USD = 5.0
DEFAULT_OUTPUT = "apartments_google_budget.md"
CACHE_PATH = ".google_budget_cache.json"

# Conservative pay-as-you-go estimates per 1,000 requests/elements. The script
# also reports usage counts so the user can compare against current Google SKU
# pricing and free monthly quotas.
TEXT_SEARCH_USD_PER_1000 = 32.0
ROUTE_MATRIX_USD_PER_1000 = 5.0

HOUSING_NAME_TERMS = (
    "apartment",
    "apartments",
    "condo",
    "condominium",
    "condominiums",
    "court",
    "lofts",
    "loft",
    "manor",
    "terrace",
    "towers",
    "tower",
    "residences",
    "residence",
    "villa",
    "villas",
    "regency",
    "house",
    "heights",
)
EXCLUDE_NAME_TERMS = (
    "fraternity",
    "sorority",
    "hotel",
    "inn",
    "dorm",
    "residential building",
    "commons building",
    "ucla store",
)
HOUSING_TYPES = {
    "apartment_complex",
    "apartment_building",
    "condominium_complex",
    "housing_complex",
}


@dataclass(frozen=True)
class Location:
    latitude: float
    longitude: float


@dataclass
class Candidate:
    place_id: str
    name: str
    address: str
    location: Location
    google_maps_uri: str
    website_uri: str
    primary_type: str
    types: list[str]
    sources: set[str] = field(default_factory=set)
    walk_distance_meters: int | None = None
    walk_duration_seconds: int | None = None
    bike_distance_meters: int | None = None
    bike_duration_seconds: int | None = None
    straight_line_meters: int | None = None
    area: str = ""


@dataclass
class AreaAnchor:
    name: str
    location: Location


@dataclass
class CostTracker:
    budget_usd: float
    text_search_calls: int = 0
    route_matrix_elements: int = 0

    @property
    def estimated_usd(self) -> float:
        return (
            self.text_search_calls * TEXT_SEARCH_USD_PER_1000 / 1000
            + self.route_matrix_elements * ROUTE_MATRIX_USD_PER_1000 / 1000
        )

    def ensure_can_add(self, text_search_calls: int = 0, route_elements: int = 0) -> None:
        projected = (
            (self.text_search_calls + text_search_calls) * TEXT_SEARCH_USD_PER_1000 / 1000
            + (self.route_matrix_elements + route_elements) * ROUTE_MATRIX_USD_PER_1000 / 1000
        )
        if projected > self.budget_usd:
            raise RuntimeError(
                f"Budget stop: projected ${projected:.2f} exceeds ${self.budget_usd:.2f}."
            )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find apartments with Google Maps APIs while enforcing a hard cost cap."
    )
    parser.add_argument("--school", required=True)
    parser.add_argument("--api-key", default=os.environ.get("GOOGLE_MAPS_API_KEY"))
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--radius-meters", type=int, default=DEFAULT_RADIUS_METERS)
    parser.add_argument("--max-walk-minutes", type=int, default=DEFAULT_MAX_WALK_MINUTES)
    parser.add_argument("--budget-usd", type=float, default=DEFAULT_BUDGET_USD)
    parser.add_argument("--max-results-per-query", type=int, default=20)
    parser.add_argument("--max-candidates", type=int, default=160)
    parser.add_argument(
        "--skip-routes",
        action="store_true",
        help="Skip Google Routes and sort/filter by straight-line distance only.",
    )
    parser.add_argument("--osm-markdown", default="apartments_osm.md")
    parser.add_argument("--seed", action="append", default=[])
    parser.add_argument("--area", action="append", default=[])
    parser.add_argument("--center-lat", type=float)
    parser.add_argument("--center-lon", type=float)
    parser.add_argument("--language-code", default="en")
    parser.add_argument("--region-code", default="US")
    parser.add_argument("--sleep-seconds", type=float, default=0.05)
    parser.add_argument("--no-cache", action="store_true")
    parser.add_argument(
        "--dry-run-cost",
        action="store_true",
        help="Show planned query counts and a conservative estimate without calling Google.",
    )
    return parser.parse_args()


def post_json_any(
    url: str,
    api_key: str,
    field_mask: str,
    payload: dict[str, Any],
) -> Any:
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "X-Goog-Api-Key": api_key,
            "X-Goog-FieldMask": field_mask,
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=40) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google API returned HTTP {exc.code}: {details}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach Google API: {exc}") from exc


def load_cache(enabled: bool) -> dict[str, Any]:
    if not enabled or not Path(CACHE_PATH).exists():
        return {"text_search": {}, "routes": {}}
    with open(CACHE_PATH, "r", encoding="utf-8") as file:
        return json.load(file)


def save_cache(enabled: bool, cache: dict[str, Any]) -> None:
    if not enabled:
        return
    with open(CACHE_PATH, "w", encoding="utf-8") as file:
        json.dump(cache, file, indent=2, sort_keys=True)


def text_search(
    api_key: str,
    query: str,
    center: Location | None,
    radius_meters: int,
    max_results: int,
    language_code: str,
    region_code: str,
    tracker: CostTracker,
    cache: dict[str, Any],
    cache_enabled: bool,
) -> list[dict[str, Any]]:
    cache_key = json.dumps(
        {
            "query": query,
            "center": None if center is None else [center.latitude, center.longitude],
            "radius": radius_meters,
            "max": max_results,
            "language": language_code,
            "region": region_code,
        },
        sort_keys=True,
    )
    if cache_enabled and cache_key in cache["text_search"]:
        return cache["text_search"][cache_key]

    tracker.ensure_can_add(text_search_calls=1)
    payload: dict[str, Any] = {
        "textQuery": query,
        "maxResultCount": max_results,
        "languageCode": language_code,
        "regionCode": region_code,
    }
    if center is not None:
        payload["locationBias"] = {
            "circle": {
                "center": {
                    "latitude": center.latitude,
                    "longitude": center.longitude,
                },
                "radius": radius_meters,
            }
        }
    data = post_json_any(
        PLACES_TEXT_SEARCH_URL,
        api_key,
        (
            "places.id,places.displayName,places.formattedAddress,places.location,"
            "places.googleMapsUri,places.websiteUri,places.primaryType,places.types"
        ),
        payload,
    )
    tracker.text_search_calls += 1
    places = list(data.get("places", []))
    if cache_enabled:
        cache["text_search"][cache_key] = places
    return places


def find_school(
    api_key: str,
    school: str,
    language_code: str,
    region_code: str,
    tracker: CostTracker,
    cache: dict[str, Any],
    cache_enabled: bool,
) -> Candidate:
    places = text_search(
        api_key,
        school,
        None,
        DEFAULT_RADIUS_METERS,
        1,
        language_code,
        region_code,
        tracker,
        cache,
        cache_enabled,
    )
    if not places:
        raise RuntimeError(f'No school found for "{school}".')
    return parse_candidate(places[0], "school")


def parse_candidate(raw: dict[str, Any], source: str) -> Candidate:
    display_name = raw.get("displayName") or {}
    location = raw.get("location") or {}
    if "latitude" not in location or "longitude" not in location:
        raise ValueError("place has no location")
    return Candidate(
        place_id=raw.get("id", ""),
        name=display_name.get("text", "Unknown place"),
        address=raw.get("formattedAddress", ""),
        location=Location(float(location["latitude"]), float(location["longitude"])),
        google_maps_uri=raw.get("googleMapsUri", ""),
        website_uri=raw.get("websiteUri", ""),
        primary_type=raw.get("primaryType", ""),
        types=list(raw.get("types", [])),
        sources={source},
    )


def load_osm_names(path: str) -> list[str]:
    markdown_path = Path(path)
    if not markdown_path.exists():
        return []
    names: list[str] = []
    for line in markdown_path.read_text(encoding="utf-8").splitlines():
        if not line.startswith("| ") or line.startswith("| #"):
            continue
        cells = [cell.strip() for cell in line.strip("|").split("|")]
        if len(cells) < 2 or not cells[0].isdigit():
            continue
        name = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", cells[1]).strip()
        if name:
            names.append(name)
    return dedupe_strings(names)


def build_queries(school: str, areas: list[str], seeds: list[str], osm_names: list[str]) -> list[tuple[str, str]]:
    queries: list[tuple[str, str]] = []
    broad_terms = [
        f"apartments near {school}",
        f"apartment buildings near {school}",
        f"student apartments near {school}",
    ]
    for area in areas:
        broad_terms.extend(
            [
                f"{area} apartments near {school}",
                f"{area} apartment buildings",
            ]
        )
    for query in broad_terms:
        queries.append(("broad", query))

    for name in dedupe_strings([*seeds, *osm_names]):
        queries.append(("targeted", f"{name} near {school}"))
    return queries


def dedupe_strings(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result = []
    for value in values:
        normalized = normalize(value)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        result.append(value)
    return result


def discover_candidates(
    api_key: str,
    queries: list[tuple[str, str]],
    center: Location,
    radius_meters: int,
    max_results_per_query: int,
    max_candidates: int,
    language_code: str,
    region_code: str,
    tracker: CostTracker,
    cache: dict[str, Any],
    cache_enabled: bool,
    sleep_seconds: float,
) -> list[Candidate]:
    by_id: dict[str, Candidate] = {}
    for source, query in queries:
        print(f"Google Text Search [{source}]: {query}", flush=True)
        raw_places = text_search(
            api_key,
            query,
            center,
            radius_meters,
            max_results_per_query,
            language_code,
            region_code,
            tracker,
            cache,
            cache_enabled,
        )
        for raw in raw_places:
            try:
                candidate = parse_candidate(raw, source)
            except ValueError:
                continue
            if haversine_meters(center, candidate.location) > radius_meters * 1.35:
                continue
            if not looks_like_housing(candidate):
                continue
            existing = by_id.get(candidate.place_id)
            if existing:
                existing.sources.add(source)
            else:
                by_id[candidate.place_id] = candidate
        time.sleep(sleep_seconds)

    candidates = sorted(
        by_id.values(),
        key=lambda item: (haversine_meters(center, item.location), item.name.lower()),
    )
    return candidates[:max_candidates]


def looks_like_housing(candidate: Candidate) -> bool:
    name = candidate.name.lower()
    if any(term in name for term in EXCLUDE_NAME_TERMS):
        return False
    if candidate.primary_type in HOUSING_TYPES or any(t in HOUSING_TYPES for t in candidate.types):
        return True
    return any(term in name for term in HOUSING_NAME_TERMS)


def add_route_metrics(
    api_key: str,
    origin: Location,
    candidates: list[Candidate],
    travel_mode: str,
    tracker: CostTracker,
    cache: dict[str, Any],
    cache_enabled: bool,
) -> None:
    for batch_start in range(0, len(candidates), 25):
        batch = candidates[batch_start : batch_start + 25]
        uncached: list[Candidate] = []
        for candidate in batch:
            route_key = route_cache_key(origin, candidate.location, travel_mode)
            cached = cache["routes"].get(route_key) if cache_enabled else None
            if cached:
                apply_route(candidate, travel_mode, cached)
            else:
                uncached.append(candidate)
        if not uncached:
            continue

        tracker.ensure_can_add(route_elements=len(uncached))
        payload = {
            "origins": [{"waypoint": {"location": {"latLng": lat_lng(origin)}}}],
            "destinations": [
                {"waypoint": {"location": {"latLng": lat_lng(candidate.location)}}}
                for candidate in uncached
            ],
            "travelMode": travel_mode,
        }
        elements = post_json_any(
            ROUTES_MATRIX_URL,
            api_key,
            "originIndex,destinationIndex,status,condition,distanceMeters,duration",
            payload,
        )
        tracker.route_matrix_elements += len(uncached)
        if not isinstance(elements, list):
            continue
        for element in elements:
            destination_index = element.get("destinationIndex")
            if destination_index is None or destination_index >= len(uncached):
                continue
            if element.get("condition") == "ROUTE_NOT_FOUND":
                continue
            distance = element.get("distanceMeters")
            duration = parse_duration_seconds(element.get("duration", ""))
            if distance is None or duration is None:
                continue
            route_data = {"distanceMeters": int(distance), "durationSeconds": duration}
            candidate = uncached[destination_index]
            apply_route(candidate, travel_mode, route_data)
            if cache_enabled:
                cache["routes"][route_cache_key(origin, candidate.location, travel_mode)] = route_data


def apply_route(candidate: Candidate, travel_mode: str, route_data: dict[str, int]) -> None:
    if travel_mode == "WALK":
        candidate.walk_distance_meters = route_data["distanceMeters"]
        candidate.walk_duration_seconds = route_data["durationSeconds"]
    elif travel_mode == "BICYCLE":
        candidate.bike_distance_meters = route_data["distanceMeters"]
        candidate.bike_duration_seconds = route_data["durationSeconds"]


def route_cache_key(origin: Location, destination: Location, travel_mode: str) -> str:
    return json.dumps(
        {
            "mode": travel_mode,
            "origin": [round(origin.latitude, 6), round(origin.longitude, 6)],
            "destination": [round(destination.latitude, 6), round(destination.longitude, 6)],
        },
        sort_keys=True,
    )


def parse_duration_seconds(duration: str) -> int | None:
    if not duration.endswith("s"):
        return None
    try:
        return int(float(duration[:-1]))
    except ValueError:
        return None


def lat_lng(location: Location) -> dict[str, float]:
    return {"latitude": location.latitude, "longitude": location.longitude}


def filter_by_walk_time(candidates: list[Candidate], max_walk_minutes: int) -> list[Candidate]:
    max_seconds = max_walk_minutes * 60
    return [
        candidate
        for candidate in candidates
        if candidate.walk_duration_seconds is not None
        and candidate.walk_duration_seconds <= max_seconds
    ]


def add_straight_line_metrics(center: Location, candidates: list[Candidate]) -> None:
    for candidate in candidates:
        candidate.straight_line_meters = round(haversine_meters(center, candidate.location))


def sort_distance_key(candidate: Candidate) -> tuple[float, str]:
    if candidate.walk_distance_meters is not None:
        distance = candidate.walk_distance_meters
    elif candidate.straight_line_meters is not None:
        distance = candidate.straight_line_meters
    else:
        distance = math.inf
    return distance, candidate.name.lower()


def find_area_anchors(
    api_key: str,
    school: str,
    areas: list[str],
    center: Location,
    radius_meters: int,
    language_code: str,
    region_code: str,
    tracker: CostTracker,
    cache: dict[str, Any],
    cache_enabled: bool,
) -> list[AreaAnchor]:
    anchor_queries = [f"{area} {school}" for area in areas]
    anchor_queries.extend([f"shopping district near {school}", f"commercial district near {school}"])
    anchors: dict[str, AreaAnchor] = {}
    for query in dedupe_strings(anchor_queries):
        raw_places = text_search(
            api_key,
            query,
            center,
            max(radius_meters, 3000),
            5,
            language_code,
            region_code,
            tracker,
            cache,
            cache_enabled,
        )
        for raw in raw_places:
            try:
                candidate = parse_candidate(raw, "area")
            except ValueError:
                continue
            anchors[candidate.place_id] = AreaAnchor(candidate.name, candidate.location)
    return sorted(
        anchors.values(),
        key=lambda anchor: (haversine_meters(center, anchor.location), anchor.name.lower()),
    )


def assign_areas(candidates: list[Candidate], anchors: list[AreaAnchor]) -> None:
    for candidate in candidates:
        if not anchors:
            candidate.area = "Nearby"
            continue
        closest = min(
            anchors,
            key=lambda anchor: haversine_meters(candidate.location, anchor.location),
        )
        candidate.area = closest.name


def grouped(candidates: list[Candidate]) -> list[tuple[str, list[Candidate]]]:
    groups: dict[str, list[Candidate]] = {}
    for candidate in candidates:
        groups.setdefault(candidate.area or "Nearby", []).append(candidate)
    for items in groups.values():
        items.sort(key=sort_distance_key)
    return sorted(
        groups.items(),
        key=lambda item: (*sort_distance_key(item[1][0]), item[0]),
    )


def write_markdown(
    path: str,
    school: str,
    school_center: Location,
    radius_meters: int,
    max_walk_minutes: int,
    candidates: list[Candidate],
    tracker: CostTracker,
    skip_routes: bool,
) -> None:
    distance_label = "Straight-line distance" if skip_routes else "Route distance"
    lines = [
        f"# Google-budget apartments near {school}",
        "",
        f"- Center: {school_center.latitude:.6f}, {school_center.longitude:.6f}",
        f"- Candidate search radius: {radius_meters} meters",
        f"- Walking-route filter: {max_walk_minutes} minutes or less",
        f"- Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"- Places shown: {len(candidates)}",
        f"- Distance mode: {'straight-line only, Google Routes skipped' if skip_routes else 'Google Routes walking/biking'}",
        f"- Google Text Search calls charged/estimated: {tracker.text_search_calls}",
        f"- Google Route Matrix elements charged/estimated: {tracker.route_matrix_elements}",
        f"- Conservative estimated API cost before free quotas: ${tracker.estimated_usd:.2f}",
        f"- Budget cap: ${tracker.budget_usd:.2f}",
        "",
    ]
    for area, items in grouped(candidates):
        lines.extend(
            [
                f"## {escape_md(area)}",
                "",
                f"| # | Apartment | Google Maps | Website | Address | Walk | Bike | {distance_label} | Sources |",
                "|---:|---|---|---|---|---:|---:|---:|---|",
            ]
        )
        for index, candidate in enumerate(items, start=1):
            lines.append(
                "| {index} | {name} | {maps} | {website} | {address} | {walk} | {bike} | {distance} | {sources} |".format(
                    index=index,
                    name=escape_md(candidate.name),
                    maps=markdown_link("map", candidate.google_maps_uri),
                    website=markdown_link("website", candidate.website_uri),
                    address=escape_md(candidate.address),
                    walk=format_minutes(candidate.walk_duration_seconds) if not skip_routes else "",
                    bike=format_minutes(candidate.bike_duration_seconds) if not skip_routes else "",
                    distance=format_meters(
                        candidate.straight_line_meters
                        if skip_routes
                        else candidate.walk_distance_meters
                    ),
                    sources=", ".join(sorted(candidate.sources)),
                )
            )
        lines.append("")

    lines.extend(
        [
            "## Notes",
            "",
            "- This file uses Google Places for candidate discovery.",
            "- If routes are skipped, distance is straight-line distance from the school center, not walking-route distance.",
            "- Cost estimate is conservative and does not subtract free monthly SKU allowances.",
            "- Results are still search-based; verify any favorite apartment on Google Maps before making decisions.",
        ]
    )
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_csv(path: str, candidates: list[Candidate]) -> None:
    csv_path = Path(path).with_suffix(".csv")
    with csv_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "name",
                "area",
                "address",
                "walk_minutes",
                "bike_minutes",
                "walk_distance_meters",
                "straight_line_meters",
                "google_maps_uri",
                "website_uri",
                "sources",
            ]
        )
        for candidate in candidates:
            writer.writerow(
                [
                    candidate.name,
                    candidate.area,
                    candidate.address,
                    minutes_number(candidate.walk_duration_seconds),
                    minutes_number(candidate.bike_duration_seconds),
                    candidate.walk_distance_meters,
                    candidate.straight_line_meters,
                    candidate.google_maps_uri,
                    candidate.website_uri,
                    ",".join(sorted(candidate.sources)),
                ]
            )


def markdown_link(label: str, url: str) -> str:
    return f"[{label}]({url})" if url else ""


def escape_md(value: str) -> str:
    return value.replace("|", "\\|")


def format_meters(meters: int | None) -> str:
    if meters is None:
        return ""
    if meters >= 1000:
        return f"{meters / 1000:.1f} km"
    return f"{meters:,} m"


def format_minutes(seconds: int | None) -> str:
    value = minutes_number(seconds)
    return "" if value is None else f"{value} min"


def minutes_number(seconds: int | None) -> int | None:
    if seconds is None:
        return None
    return max(1, round(seconds / 60))


def haversine_meters(a: Location, b: Location) -> float:
    radius = 6_371_000
    lat1, lat2 = math.radians(a.latitude), math.radians(b.latitude)
    d_lat = math.radians(b.latitude - a.latitude)
    d_lon = math.radians(b.longitude - a.longitude)
    h = (
        math.sin(d_lat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(d_lon / 2) ** 2
    )
    return 2 * radius * math.asin(math.sqrt(h))


def normalize(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def main() -> int:
    args = parse_args()
    if not args.api_key and not args.dry_run_cost:
        print("Missing API key. Set GOOGLE_MAPS_API_KEY or pass --api-key.", file=sys.stderr)
        return 2

    cache_enabled = not args.no_cache
    cache = load_cache(cache_enabled)
    tracker = CostTracker(args.budget_usd)

    osm_names = load_osm_names(args.osm_markdown)
    seeds = dedupe_strings(args.seed)
    areas = dedupe_strings(args.area)
    queries = build_queries(args.school, areas, seeds, osm_names)
    conservative_text_cost = (len(queries) + len(areas) + 3) * TEXT_SEARCH_USD_PER_1000 / 1000
    route_elements = 0 if args.skip_routes else args.max_candidates * 2
    conservative_route_cost = route_elements * ROUTE_MATRIX_USD_PER_1000 / 1000
    if args.dry_run_cost:
        print(f"Planned Text Search queries: {len(queries) + len(areas) + 3}")
        print(f"Planned route elements upper bound: {route_elements}")
        print(f"Conservative estimate before free quotas: ${conservative_text_cost + conservative_route_cost:.2f}")
        return 0

    school_candidate = find_school(
        args.api_key,
        args.school,
        args.language_code,
        args.region_code,
        tracker,
        cache,
        cache_enabled,
    )
    center = school_candidate.location
    if (args.center_lat is None) != (args.center_lon is None):
        raise RuntimeError("Pass both --center-lat and --center-lon, or neither.")
    if args.center_lat is not None and args.center_lon is not None:
        center = Location(args.center_lat, args.center_lon)

    print(f"Using center: {center.latitude:.6f}, {center.longitude:.6f}", flush=True)
    print(f"Loaded {len(osm_names)} OSM names and {len(seeds)} seed names.", flush=True)
    candidates = discover_candidates(
        args.api_key,
        queries,
        center,
        args.radius_meters,
        args.max_results_per_query,
        args.max_candidates,
        args.language_code,
        args.region_code,
        tracker,
        cache,
        cache_enabled,
        args.sleep_seconds,
    )
    print(f"Google candidates after filtering/deduping: {len(candidates)}", flush=True)
    add_straight_line_metrics(center, candidates)
    if args.skip_routes:
        candidates = [
            candidate
            for candidate in candidates
            if candidate.straight_line_meters is not None
            and candidate.straight_line_meters <= args.radius_meters
        ]
    else:
        add_route_metrics(args.api_key, center, candidates, "WALK", tracker, cache, cache_enabled)
        add_route_metrics(args.api_key, center, candidates, "BICYCLE", tracker, cache, cache_enabled)
        candidates = filter_by_walk_time(candidates, args.max_walk_minutes)
    candidates.sort(key=sort_distance_key)
    anchors = find_area_anchors(
        args.api_key,
        args.school,
        areas,
        center,
        args.radius_meters,
        args.language_code,
        args.region_code,
        tracker,
        cache,
        cache_enabled,
    )
    assign_areas(candidates, anchors)
    write_markdown(
        args.output,
        args.school,
        center,
        args.radius_meters,
        args.max_walk_minutes,
        candidates,
        tracker,
        args.skip_routes,
    )
    write_csv(args.output, candidates)
    save_cache(cache_enabled, cache)
    print(f"Wrote {args.output} and {Path(args.output).with_suffix('.csv')}.")
    print(f"Conservative estimated API cost before free quotas: ${tracker.estimated_usd:.2f}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RuntimeError as exc:
        message = str(exc)
        print("\nError:", file=sys.stderr)
        if "SERVICE_DISABLED" in message and "places.googleapis.com" in message:
            print(
                "Places API (New) is not enabled for this Google Cloud project.",
                file=sys.stderr,
            )
            print(
                "Open this link, click Enable, wait 1-3 minutes, then run again:",
                file=sys.stderr,
            )
            print(
                "https://console.developers.google.com/apis/api/places.googleapis.com/overview?project=952773000627",
                file=sys.stderr,
            )
        elif "API_KEY_INVALID" in message:
            print(
                "The API key was not valid. Paste only the key itself; it should start with AIza.",
                file=sys.stderr,
            )
        else:
            print(message, file=sys.stderr)
        raise SystemExit(1)
