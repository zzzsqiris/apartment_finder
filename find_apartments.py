#!/usr/bin/env python3
"""Find apartment communities near a university using Google Maps APIs.

The script writes a Markdown report with each apartment's name, website URL
when Google has one, Google Maps URL, address, and route distance/time from the
school.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime
from typing import Any


PLACES_TEXT_SEARCH_URL = "https://places.googleapis.com/v1/places:searchText"
PLACES_NEARBY_SEARCH_URL = "https://places.googleapis.com/v1/places:searchNearby"
ROUTES_MATRIX_URL = "https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix"

DEFAULT_RADIUS_METERS = 1600
DEFAULT_MAX_WALK_MINUTES = 20
DEFAULT_TYPES = (
    "apartment_complex",
    "apartment_building",
    "condominium_complex",
    "housing_complex",
)
AREA_QUERIES = (
    "shopping district",
    "commercial district",
    "business district",
    "downtown",
    "shopping center",
)


@dataclass(frozen=True)
class Location:
    latitude: float
    longitude: float


@dataclass
class Place:
    place_id: str
    name: str
    location: Location | None
    address: str
    website_uri: str
    maps_uri: str
    primary_type: str
    types: list[str]
    distance_meters: float | None = None
    north_meters: float | None = None
    east_meters: float | None = None
    walk_distance_meters: int | None = None
    walk_duration_seconds: int | None = None
    bike_distance_meters: int | None = None
    bike_duration_seconds: int | None = None


@dataclass
class AreaAnchor:
    place_id: str
    name: str
    location: Location
    address: str
    maps_uri: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find apartment communities within walking distance of a school."
    )
    parser.add_argument(
        "--school",
        required=True,
        help='School name or address, for example "University of Washington".',
    )
    parser.add_argument(
        "--radius-meters",
        type=int,
        default=DEFAULT_RADIUS_METERS,
        help="Initial candidate search radius. Final output is filtered by route walking time.",
    )
    parser.add_argument(
        "--max-walk-minutes",
        type=int,
        default=DEFAULT_MAX_WALK_MINUTES,
        help="Keep apartments within this Google walking-route duration.",
    )
    parser.add_argument(
        "--output",
        default="apartments.md",
        help="Markdown output path.",
    )
    parser.add_argument(
        "--language-code",
        default="en",
        help='Google Places language code, for example "en" or "zh-CN".',
    )
    parser.add_argument(
        "--region-code",
        default="US",
        help='Google Places region code, for example "US".',
    )
    parser.add_argument(
        "--include-no-website",
        action="store_true",
        help="Also include places that do not have a website URL on Google Maps.",
    )
    parser.add_argument(
        "--api-key",
        default=os.environ.get("GOOGLE_MAPS_API_KEY"),
        help="Google Maps API key. Defaults to GOOGLE_MAPS_API_KEY.",
    )
    parser.add_argument(
        "--sleep-seconds",
        type=float,
        default=0.05,
        help="Small pause between API calls.",
    )
    return parser.parse_args()


def post_json(url: str, api_key: str, field_mask: str, payload: dict[str, Any]) -> dict[str, Any]:
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
        with urllib.request.urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google API returned HTTP {exc.code}: {details}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach Google API: {exc}") from exc


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
        with urllib.request.urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Google API returned HTTP {exc.code}: {details}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach Google API: {exc}") from exc


def find_school(api_key: str, query: str, language_code: str, region_code: str) -> Place:
    payload = {
        "textQuery": query,
        "maxResultCount": 1,
        "languageCode": language_code,
        "regionCode": region_code,
    }
    data = post_json(
        PLACES_TEXT_SEARCH_URL,
        api_key,
        "places.id,places.displayName,places.formattedAddress,places.location,places.googleMapsUri",
        payload,
    )
    places = data.get("places", [])
    if not places:
        raise RuntimeError(f'No school found for query: "{query}"')
    return parse_place(places[0])


def nearby_apartments(
    api_key: str,
    center: Location,
    radius_meters: int,
    language_code: str,
    region_code: str,
    sleep_seconds: float,
) -> list[Place]:
    by_id: dict[str, Place] = {}
    centers = search_centers(center, radius_meters)
    field_mask = (
        "places.id,places.displayName,places.formattedAddress,places.location,"
        "places.websiteUri,places.googleMapsUri,places.primaryType,places.types"
    )

    for place_type in DEFAULT_TYPES:
        for tile_center, tile_radius in centers:
            payload = {
                "includedTypes": [place_type],
                "maxResultCount": 20,
                "rankPreference": "DISTANCE",
                "languageCode": language_code,
                "regionCode": region_code,
                "locationRestriction": {
                    "circle": {
                        "center": {
                            "latitude": tile_center.latitude,
                            "longitude": tile_center.longitude,
                        },
                        "radius": tile_radius,
                    }
                },
            }
            data = post_json(PLACES_NEARBY_SEARCH_URL, api_key, field_mask, payload)
            for raw_place in data.get("places", []):
                place = parse_place(raw_place)
                if not place.place_id or place.place_id in by_id:
                    continue
                if place.location is None:
                    continue
                distance = haversine_meters(center, place.location)
                if distance <= radius_meters:
                    place.distance_meters = distance
                    north_m, east_m = relative_offset_meters(center, place.location)
                    place.north_meters = north_m
                    place.east_meters = east_m
                    by_id[place.place_id] = place
            time.sleep(sleep_seconds)

    return sorted(
        by_id.values(),
        key=lambda place: (
            place.distance_meters if place.distance_meters is not None else math.inf,
            place.name.lower(),
        ),
    )


def find_area_anchors(
    api_key: str,
    school_name: str,
    center: Location,
    radius_meters: int,
    language_code: str,
    region_code: str,
    sleep_seconds: float,
) -> list[AreaAnchor]:
    by_id: dict[str, AreaAnchor] = {}
    field_mask = (
        "places.id,places.displayName,places.formattedAddress,places.location,"
        "places.googleMapsUri,places.primaryType,places.types"
    )
    search_radius = max(radius_meters, 2500)

    for query in AREA_QUERIES:
        payload = {
            "textQuery": f"{query} near {school_name}",
            "maxResultCount": 10,
            "languageCode": language_code,
            "regionCode": region_code,
            "locationBias": {
                "circle": {
                    "center": {
                        "latitude": center.latitude,
                        "longitude": center.longitude,
                    },
                    "radius": search_radius,
                }
            },
        }
        data = post_json(PLACES_TEXT_SEARCH_URL, api_key, field_mask, payload)
        for raw_place in data.get("places", []):
            place = parse_place(raw_place)
            if not place.place_id or place.place_id in by_id or place.location is None:
                continue
            if haversine_meters(center, place.location) > search_radius:
                continue
            by_id[place.place_id] = AreaAnchor(
                place_id=place.place_id,
                name=place.name,
                location=place.location,
                address=place.address,
                maps_uri=place.maps_uri,
            )
        time.sleep(sleep_seconds)

    return sorted(
        by_id.values(),
        key=lambda anchor: (haversine_meters(center, anchor.location), anchor.name.lower()),
    )


def add_route_metrics(
    api_key: str,
    origin: Location,
    places: list[Place],
    travel_mode: str,
) -> None:
    destinations = [place for place in places if place.location is not None]
    for batch_start in range(0, len(destinations), 25):
        batch = destinations[batch_start : batch_start + 25]
        payload = {
            "origins": [{"waypoint": {"location": {"latLng": lat_lng(origin)}}}],
            "destinations": [
                {"waypoint": {"location": {"latLng": lat_lng(place.location)}}}
                for place in batch
                if place.location is not None
            ],
            "travelMode": travel_mode,
        }
        elements = post_json_any(
            ROUTES_MATRIX_URL,
            api_key,
            "originIndex,destinationIndex,status,condition,distanceMeters,duration",
            payload,
        )
        if not isinstance(elements, list):
            continue

        for element in elements:
            destination_index = element.get("destinationIndex")
            if destination_index is None or destination_index >= len(batch):
                continue
            place = batch[destination_index]
            if element.get("condition") == "ROUTE_NOT_FOUND":
                continue
            distance = element.get("distanceMeters")
            duration = parse_duration_seconds(element.get("duration", ""))
            if distance is None or duration is None:
                continue
            if travel_mode == "WALK":
                place.walk_distance_meters = int(distance)
                place.walk_duration_seconds = duration
            elif travel_mode == "BICYCLE":
                place.bike_distance_meters = int(distance)
                place.bike_duration_seconds = duration


def lat_lng(location: Location | None) -> dict[str, float]:
    if location is None:
        raise ValueError("location is required")
    return {"latitude": location.latitude, "longitude": location.longitude}


def parse_duration_seconds(duration: str) -> int | None:
    if not duration.endswith("s"):
        return None
    try:
        return int(float(duration[:-1]))
    except ValueError:
        return None


def filter_by_walk_time(places: list[Place], max_walk_minutes: int) -> list[Place]:
    max_seconds = max_walk_minutes * 60
    return [
        place
        for place in places
        if place.walk_duration_seconds is not None
        and place.walk_duration_seconds <= max_seconds
    ]


def search_centers(center: Location, radius_meters: int) -> list[tuple[Location, int]]:
    """Return one broad search plus smaller circles across the target radius.

    Nearby Search returns at most 20 results per request, so tiling the area helps
    discover more apartments when a campus is surrounded by dense housing.
    """
    tile_radius = max(350, min(700, radius_meters // 2))
    step = tile_radius
    centers: list[tuple[Location, int]] = [(center, radius_meters)]

    for north_m in range(-radius_meters, radius_meters + 1, step):
        for east_m in range(-radius_meters, radius_meters + 1, step):
            distance = math.hypot(north_m, east_m)
            if distance > radius_meters:
                continue
            if north_m == 0 and east_m == 0:
                continue
            centers.append((offset_location(center, north_m, east_m), tile_radius))

    return centers


def parse_place(raw: dict[str, Any]) -> Place:
    display_name = raw.get("displayName") or {}
    location = raw.get("location") or {}
    parsed_location = None
    if "latitude" in location and "longitude" in location:
        parsed_location = Location(float(location["latitude"]), float(location["longitude"]))

    return Place(
        place_id=raw.get("id", ""),
        name=display_name.get("text", "Unknown place"),
        location=parsed_location,
        address=raw.get("formattedAddress", ""),
        website_uri=raw.get("websiteUri", ""),
        maps_uri=raw.get("googleMapsUri", ""),
        primary_type=raw.get("primaryType", ""),
        types=list(raw.get("types", [])),
    )


def offset_location(origin: Location, north_meters: float, east_meters: float) -> Location:
    meters_per_degree_lat = 111_320
    meters_per_degree_lon = 111_320 * math.cos(math.radians(origin.latitude))
    return Location(
        origin.latitude + north_meters / meters_per_degree_lat,
        origin.longitude + east_meters / meters_per_degree_lon,
    )


def relative_offset_meters(origin: Location, point: Location) -> tuple[float, float]:
    meters_per_degree_lat = 111_320
    meters_per_degree_lon = 111_320 * math.cos(math.radians(origin.latitude))
    north_meters = (point.latitude - origin.latitude) * meters_per_degree_lat
    east_meters = (point.longitude - origin.longitude) * meters_per_degree_lon
    return north_meters, east_meters


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


def markdown_link(label: str, url: str) -> str:
    if not url:
        return ""
    escaped = label.replace("[", "\\[").replace("]", "\\]")
    return f"[{escaped}]({url})"


def area_name(place: Place, anchors: list[AreaAnchor]) -> str:
    if place.location is None or not anchors:
        return "Campus Area"
    closest = min(
        anchors,
        key=lambda anchor: haversine_meters(place.location, anchor.location),
    )
    return closest.name


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


def format_meters(meters: int | None) -> str:
    if meters is None:
        return ""
    if meters >= 1000:
        return f"{meters / 1000:.1f} km"
    return f"{meters:,} m"


def format_minutes(seconds: int | None) -> str:
    if seconds is None:
        return ""
    minutes = max(1, round(seconds / 60))
    return f"{minutes} min"


def write_markdown(
    output_path: str,
    school_query: str,
    school: Place,
    radius_meters: int,
    places: list[Place],
    anchors: list[AreaAnchor],
    include_no_website: bool,
    max_walk_minutes: int,
) -> None:
    visible_places = [
        place for place in places if include_no_website or bool(place.website_uri)
    ]
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    lines = [
        f"# Apartments near {school.name}",
        "",
        f"- Search input: `{school_query}`",
        f"- Center: {school.address or school.name}",
        f"- Candidate search radius: {radius_meters} meters",
        f"- Walking-route filter: {max_walk_minutes} minutes or less",
        f"- Generated: {generated_at}",
        f"- Places found in radius: {len(places)}",
        f"- Places shown: {len(visible_places)}",
        f"- Commercial groups found: {len(anchors)}",
        "",
    ]

    for group_name, group_places in grouped_by_area(visible_places, anchors):
        lines.extend(
            [
                f"## {group_name}",
                "",
                "| # | Apartment | Website | Google Maps | Address | Walk | Bike | Route distance |",
                "|---:|---|---|---|---|---:|---:|---:|",
            ]
        )

        for index, place in enumerate(group_places, start=1):
            website = markdown_link("website", place.website_uri) or "No website listed"
            maps = markdown_link("map", place.maps_uri) or ""
            name = place.name.replace("|", "\\|")
            address = place.address.replace("|", "\\|")
            walk = format_minutes(place.walk_duration_seconds)
            bike = format_minutes(place.bike_duration_seconds)
            route_distance = format_meters(place.walk_distance_meters)
            lines.append(
                f"| {index} | {name} | {website} | {maps} | {address} | {walk} | {bike} | {route_distance} |"
            )
        lines.append("")

    if not visible_places:
        lines.append(
            "No places with website links were found. Run again with `--include-no-website` to show every result."
        )
        lines.append("")

    lines.extend(
        [
            "",
            "## Notes",
            "",
            "- Apartment candidates come from Google Places API Nearby Search.",
            "- Walking and biking times/distances come from Google Routes API route matrix results.",
            "- Groups use nearby commercial-area anchors found through Google Places Text Search, then each apartment is assigned to its nearest anchor.",
            "- Nearby Search returns a limited number of results per request; this script uses multiple housing types and tiled searches to reduce missed places.",
        ]
    )

    with open(output_path, "w", encoding="utf-8") as file:
        file.write("\n".join(lines) + "\n")


def main() -> int:
    args = parse_args()
    if not args.api_key:
        print(
            "Missing API key. Set GOOGLE_MAPS_API_KEY or pass --api-key.",
            file=sys.stderr,
        )
        return 2

    school = find_school(
        args.api_key,
        args.school,
        args.language_code,
        args.region_code,
    )
    if school.location is None:
        raise RuntimeError(f'Could not get a location for "{args.school}"')

    places = nearby_apartments(
        args.api_key,
        school.location,
        args.radius_meters,
        args.language_code,
        args.region_code,
        args.sleep_seconds,
    )
    add_route_metrics(args.api_key, school.location, places, "WALK")
    add_route_metrics(args.api_key, school.location, places, "BICYCLE")
    route_places = filter_by_walk_time(places, args.max_walk_minutes)
    anchors = find_area_anchors(
        args.api_key,
        school.name,
        school.location,
        args.radius_meters,
        args.language_code,
        args.region_code,
        args.sleep_seconds,
    )
    write_markdown(
        args.output,
        args.school,
        school,
        args.radius_meters,
        route_places,
        anchors,
        args.include_no_website,
        args.max_walk_minutes,
    )
    print(f"Wrote {args.output} with {len(route_places)} places found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
