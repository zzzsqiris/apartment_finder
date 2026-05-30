#!/usr/bin/env python3
"""Free apartment finder using OpenStreetMap ecosystem services.

This version does not need a Google Maps API key. It uses:
- Nominatim for school geocoding
- Overpass API for apartment and commercial-area data
- OSRM public demo server for walking and biking routes
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from typing import Any


NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OVERPASS_URL = "https://overpass-api.de/api/interpreter"
OSRM_URL = "https://router.project-osrm.org/route/v1"

DEFAULT_RADIUS_METERS = 2200
DEFAULT_MAX_WALK_MINUTES = 20
DEFAULT_MAX_CANDIDATES = 40
DEFAULT_USER_AGENT = "codex-apartment-finder/1.0"
WALK_METERS_PER_MINUTE = 80
BIKE_METERS_PER_MINUTE = 250
EXCLUDED_STUDENT_HOUSING_TERMS = (
    "fraternity",
    "sorority",
    "residential building",
    "commons building",
    "triangle",
    "kappa",
    "gamma",
    "alpha",
    "chi omega",
    "pi beta phi",
    "beta theta pi",
    "phi delta theta",
    "sigma nu",
    "theta delta chi",
    "lambda chi alpha",
    "phi kappa psi",
    "theta xi",
    "zeta beta theta",
    "alpha tau omega",
    "sigma phi epsilon",
    "phi kappa sigma",
    "sigma pi",
    "delta sigma phi",
    "delta tau delta",
    "alpha gamma omega",
    "sigma alpha epsilon",
    "theta chi",
    "sigma chi",
    "pi kappa phi",
    "alpha epsilon pi",
)


@dataclass(frozen=True)
class Location:
    latitude: float
    longitude: float


@dataclass
class Place:
    osm_type: str
    osm_id: int
    name: str
    location: Location
    address: str
    website: str
    osm_url: str
    walk_distance_meters: float | None = None
    walk_duration_seconds: float | None = None
    bike_distance_meters: float | None = None
    bike_duration_seconds: float | None = None


@dataclass
class AreaAnchor:
    name: str
    location: Location
    osm_url: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find apartments near a school with free OpenStreetMap services."
    )
    parser.add_argument("--school", required=True, help="School name or address.")
    parser.add_argument(
        "--center-lat",
        type=float,
        help="Optional manual school-center latitude.",
    )
    parser.add_argument(
        "--center-lon",
        type=float,
        help="Optional manual school-center longitude.",
    )
    parser.add_argument(
        "--radius-meters",
        type=int,
        default=DEFAULT_RADIUS_METERS,
        help="Candidate search radius around the school.",
    )
    parser.add_argument(
        "--max-walk-minutes",
        type=int,
        default=DEFAULT_MAX_WALK_MINUTES,
        help="Keep apartments within this routed walking duration.",
    )
    parser.add_argument(
        "--max-candidates",
        type=int,
        default=DEFAULT_MAX_CANDIDATES,
        help="Route only this many nearest apartment candidates to keep free services responsive.",
    )
    parser.add_argument("--output", default="apartments_osm.md")
    parser.add_argument(
        "--include-student-housing",
        action="store_true",
        help="Include dorms, fraternities, and sororities when OSM returns them.",
    )
    parser.add_argument(
        "--user-agent",
        default=os.environ.get("OSM_USER_AGENT", DEFAULT_USER_AGENT),
        help="User-Agent sent to OSM services. Set OSM_USER_AGENT for your own identifier.",
    )
    parser.add_argument("--sleep-seconds", type=float, default=1.0)
    return parser.parse_args()


def get_json(url: str, user_agent: str) -> Any:
    request = urllib.request.Request(url, headers={"User-Agent": user_agent})
    with urllib.request.urlopen(request, timeout=45) as response:
        return json.loads(response.read().decode("utf-8"))


def post_form(url: str, data: dict[str, str], user_agent: str) -> Any:
    body = urllib.parse.urlencode(data).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "User-Agent": user_agent,
        },
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8"))


def geocode_school(query: str, user_agent: str) -> tuple[str, str, Location]:
    params = urllib.parse.urlencode(
        {
            "q": query,
            "format": "jsonv2",
            "limit": 1,
            "addressdetails": 1,
        }
    )
    data = get_json(f"{NOMINATIM_URL}?{params}", user_agent)
    if not data:
        raise RuntimeError(f'No school found for "{query}"')
    result = data[0]
    return (
        result.get("display_name", query).split(",")[0],
        result.get("display_name", ""),
        Location(float(result["lat"]), float(result["lon"])),
    )


def overpass_query(query: str, user_agent: str) -> list[dict[str, Any]]:
    data = post_form(OVERPASS_URL, {"data": query}, user_agent)
    return list(data.get("elements", []))


def find_apartments(center: Location, radius_meters: int, user_agent: str) -> list[Place]:
    query = f"""
[out:json][timeout:45];
(
  nwr["building"="apartments"](around:{radius_meters},{center.latitude},{center.longitude});
  nwr["building"="residential"]["name"](around:{radius_meters},{center.latitude},{center.longitude});
  nwr["residential"="apartments"](around:{radius_meters},{center.latitude},{center.longitude});
  nwr["amenity"="student_accommodation"](around:{radius_meters},{center.latitude},{center.longitude});
);
out center tags;
"""
    places: dict[tuple[str, int], Place] = {}
    for element in overpass_query(query, user_agent):
        tags = element.get("tags", {})
        name = tags.get("name")
        if not name:
            continue
        location = element_location(element)
        if location is None:
            continue
        key = (element["type"], int(element["id"]))
        places[key] = Place(
            osm_type=element["type"],
            osm_id=int(element["id"]),
            name=name,
            location=location,
            address=format_address(tags),
            website=first_tag(tags, "website", "contact:website", "url"),
            osm_url=osm_url(element["type"], int(element["id"])),
        )
    return sorted(places.values(), key=lambda place: place.name.lower())


def find_area_anchors(center: Location, radius_meters: int, user_agent: str) -> list[AreaAnchor]:
    anchor_radius = max(radius_meters, 3000)
    query = f"""
[out:json][timeout:45];
(
  nwr["place"~"neighbourhood|quarter|suburb"](around:{anchor_radius},{center.latitude},{center.longitude});
  nwr["shop"="mall"](around:{anchor_radius},{center.latitude},{center.longitude});
  nwr["landuse"="commercial"](around:{anchor_radius},{center.latitude},{center.longitude});
  nwr["amenity"="marketplace"](around:{anchor_radius},{center.latitude},{center.longitude});
);
out center tags;
"""
    anchors: dict[tuple[str, int], AreaAnchor] = {}
    for element in overpass_query(query, user_agent):
        tags = element.get("tags", {})
        name = tags.get("name")
        location = element_location(element)
        if not name or location is None:
            continue
        anchors[(element["type"], int(element["id"]))] = AreaAnchor(
            name=name,
            location=location,
            osm_url=osm_url(element["type"], int(element["id"])),
        )
    return sorted(
        anchors.values(),
        key=lambda anchor: (haversine_meters(center, anchor.location), anchor.name.lower()),
    )


def add_routes(center: Location, places: list[Place], user_agent: str, sleep_seconds: float) -> None:
    total = len(places)
    for index, place in enumerate(places, start=1):
        print(f"Routing {index}/{total}: {place.name}", flush=True)
        walk = route(center, place.location, "foot", user_agent)
        if walk:
            place.walk_distance_meters = walk[0]
            place.walk_duration_seconds = walk[0] / WALK_METERS_PER_MINUTE * 60
        time.sleep(sleep_seconds)

        bike = route(center, place.location, "bike", user_agent)
        if bike:
            place.bike_distance_meters = bike[0]
            place.bike_duration_seconds = bike[0] / BIKE_METERS_PER_MINUTE * 60
        time.sleep(sleep_seconds)


def route(
    origin: Location,
    destination: Location,
    profile: str,
    user_agent: str,
) -> tuple[float, float] | None:
    coords = (
        f"{origin.longitude},{origin.latitude};"
        f"{destination.longitude},{destination.latitude}"
    )
    url = f"{OSRM_URL}/{profile}/{coords}?overview=false&alternatives=false&steps=false"
    try:
        data = get_json(url, user_agent)
    except Exception:
        return None
    routes = data.get("routes") or []
    if not routes:
        return None
    first = routes[0]
    return float(first["distance"]), float(first["duration"])


def filter_by_walk_time(places: list[Place], max_walk_minutes: int) -> list[Place]:
    max_seconds = max_walk_minutes * 60
    return [
        place
        for place in places
        if place.walk_duration_seconds is not None
        and place.walk_duration_seconds <= max_seconds
    ]


def nearest_candidates(center: Location, places: list[Place], max_candidates: int) -> list[Place]:
    return sorted(
        places,
        key=lambda place: (haversine_meters(center, place.location), place.name.lower()),
    )[:max_candidates]


def exclude_student_housing(places: list[Place]) -> list[Place]:
    filtered = []
    seen: set[tuple[str, str]] = set()
    for place in places:
        normalized_name = place.name.lower()
        if any(term in normalized_name for term in EXCLUDED_STUDENT_HOUSING_TERMS):
            continue
        dedupe_key = (normalized_name, place.address.lower())
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)
        filtered.append(place)
    return filtered


def element_location(element: dict[str, Any]) -> Location | None:
    if "lat" in element and "lon" in element:
        return Location(float(element["lat"]), float(element["lon"]))
    center = element.get("center")
    if center and "lat" in center and "lon" in center:
        return Location(float(center["lat"]), float(center["lon"]))
    return None


def first_tag(tags: dict[str, str], *keys: str) -> str:
    for key in keys:
        value = tags.get(key)
        if value:
            return value
    return ""


def format_address(tags: dict[str, str]) -> str:
    parts = [
        " ".join(
            part
            for part in [tags.get("addr:housenumber", ""), tags.get("addr:street", "")]
            if part
        ).strip(),
        tags.get("addr:city", ""),
        tags.get("addr:state", ""),
        tags.get("addr:postcode", ""),
    ]
    return ", ".join(part for part in parts if part)


def osm_url(osm_type: str, osm_id: int) -> str:
    type_name = {"node": "node", "way": "way", "relation": "relation"}[osm_type]
    return f"https://www.openstreetmap.org/{type_name}/{osm_id}"


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


def area_name(place: Place, anchors: list[AreaAnchor]) -> str:
    if not anchors:
        return "Nearby Area"
    return min(
        anchors,
        key=lambda anchor: haversine_meters(place.location, anchor.location),
    ).name


def grouped_by_area(
    places: list[Place],
    anchors: list[AreaAnchor],
) -> list[tuple[str, list[Place]]]:
    groups: dict[str, list[Place]] = {}
    for place in places:
        groups.setdefault(area_name(place, anchors), []).append(place)
    for group_places in groups.values():
        group_places.sort(
            key=lambda place: (
                place.walk_distance_meters
                if place.walk_distance_meters is not None
                else math.inf,
                place.name.lower(),
            )
        )
    return sorted(
        groups.items(),
        key=lambda item: (
            item[1][0].walk_distance_meters
            if item[1][0].walk_distance_meters is not None
            else math.inf,
            item[0],
        ),
    )


def markdown_link(label: str, url: str) -> str:
    if not url:
        return ""
    return f"[{label}]({url})"


def format_meters(meters: float | None) -> str:
    if meters is None:
        return ""
    if meters >= 1000:
        return f"{meters / 1000:.1f} km"
    return f"{round(meters):,} m"


def format_minutes(seconds: float | None) -> str:
    if seconds is None:
        return ""
    return f"{max(1, round(seconds / 60))} min"


def write_markdown(
    output_path: str,
    school_query: str,
    school_name: str,
    school_address: str,
    radius_meters: int,
    max_walk_minutes: int,
    places: list[Place],
    anchors: list[AreaAnchor],
) -> None:
    lines = [
        f"# Apartments near {school_name}",
        "",
        f"- Search input: `{school_query}`",
        f"- Center: {school_address}",
        f"- Candidate search radius: {radius_meters} meters",
        f"- Walking-route filter: {max_walk_minutes} minutes or less",
        f"- Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        f"- Places shown: {len(places)}",
        f"- Commercial groups found: {len(anchors)}",
        f"- Data source: OpenStreetMap / Overpass / OSRM",
        "",
    ]

    for group_name, group_places in grouped_by_area(places, anchors):
        safe_group = group_name.replace("|", "\\|")
        lines.extend(
            [
                f"## {safe_group}",
                "",
                "| # | Apartment | Website | OSM | Address | Walk | Bike | Route distance |",
                "|---:|---|---|---|---|---:|---:|---:|",
            ]
        )
        for index, place in enumerate(group_places, start=1):
            website = markdown_link("website", place.website) or ""
            osm = markdown_link("map", place.osm_url)
            name = place.name.replace("|", "\\|")
            address = place.address.replace("|", "\\|")
            lines.append(
                f"| {index} | {name} | {website} | {osm} | {address} | "
                f"{format_minutes(place.walk_duration_seconds)} | "
                f"{format_minutes(place.bike_duration_seconds)} | "
                f"{format_meters(place.walk_distance_meters)} |"
            )
        lines.append("")

    lines.extend(
        [
            "## Notes",
            "",
            "- This free version depends on public community services and should be used gently.",
            "- OSM website links are only available when contributors have added website/contact tags.",
            "- Commercial grouping is approximate: each apartment is assigned to the nearest named commercial/neighbourhood anchor in OSM.",
        ]
    )
    with open(output_path, "w", encoding="utf-8") as file:
        file.write("\n".join(lines) + "\n")


def main() -> int:
    args = parse_args()
    school_name, school_address, center = geocode_school(args.school, args.user_agent)
    if (args.center_lat is None) != (args.center_lon is None):
        raise RuntimeError("Pass both --center-lat and --center-lon, or neither.")
    if args.center_lat is not None and args.center_lon is not None:
        center = Location(args.center_lat, args.center_lon)
        school_address = f"{school_address} (manual center: {args.center_lat}, {args.center_lon})"
    print(f"Found school: {school_address}", flush=True)
    time.sleep(args.sleep_seconds)
    places = find_apartments(center, args.radius_meters, args.user_agent)
    print(f"Found {len(places)} apartment candidates before routing.", flush=True)
    if not args.include_student_housing:
        places = exclude_student_housing(places)
        print(f"Kept {len(places)} candidates after filtering student housing/fraternities.", flush=True)
    places = nearest_candidates(center, places, args.max_candidates)
    print(f"Routing nearest {len(places)} candidates.", flush=True)
    time.sleep(args.sleep_seconds)
    anchors = find_area_anchors(center, args.radius_meters, args.user_agent)
    print(f"Found {len(anchors)} commercial/neighbourhood anchors.", flush=True)
    add_routes(center, places, args.user_agent, args.sleep_seconds)
    visible_places = filter_by_walk_time(places, args.max_walk_minutes)
    visible_places.sort(
        key=lambda place: (
            place.walk_distance_meters
            if place.walk_distance_meters is not None
            else math.inf,
            place.name.lower(),
        )
    )
    write_markdown(
        args.output,
        args.school,
        school_name,
        school_address,
        args.radius_meters,
        args.max_walk_minutes,
        visible_places,
        anchors,
    )
    print(f"Wrote {args.output} with {len(visible_places)} places.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except urllib.error.HTTPError as exc:
        print(f"HTTP error {exc.code}: {exc.read().decode('utf-8', errors='replace')}", file=sys.stderr)
        raise SystemExit(1)
