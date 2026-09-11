"""Public App Store metadata for App Benchmark.

Mount ``router`` with the application's existing authentication dependency.
This module never accepts a provider URL and never estimates private metrics.
Provider reference: https://developer.apple.com/library/archive/documentation/
AudioVideo/Conceptual/iTuneSearchAPI/Searching.html
"""

from collections import OrderedDict
from copy import deepcopy
from datetime import datetime, timezone
import math
import threading
import time
from typing import Annotated, Any, Literal
from urllib.parse import urlparse

from fastapi import APIRouter, HTTPException, Query
from pydantic import BaseModel, ConfigDict, Field, field_validator
import requests


Country = Literal["tr", "us", "gb", "de", "fr", "it", "es", "nl", "ae", "sa", "ca", "au", "br", "in", "jp"]
AppId = Annotated[int, Field(strict=True, gt=0, le=9223372036854775807)]
SOURCE = "Apple App Store"
BASE_URL = "https://itunes.apple.com"
CACHE_TTL_SECONDS = 15 * 60
CACHE_MAX_ENTRIES = 128
PROVIDER_MIN_INTERVAL_SECONDS = 3.1
router = APIRouter(prefix="/api/app-benchmark", tags=["App Benchmark"])


class CompareRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    app_ids: Annotated[list[AppId], Field(min_length=2, max_length=5)]
    country: Country = "tr"

    @field_validator("app_ids")
    @classmethod
    def unique_app_ids(cls, value: list[int]) -> list[int]:
        if len(set(value)) != len(value):
            raise ValueError("Choose 2–5 different apps; the first app is the baseline.")
        return value


class ReviewsRequest(CompareRequest):
    pass


class AppleProvider:
    """Bounded per-process cache, serialized misses and gentle provider pacing."""

    def __init__(self) -> None:
        self._cache: OrderedDict[tuple, tuple[float, dict]] = OrderedDict()
        self._cache_lock = threading.Lock()
        self._fetch_lock = threading.Lock()
        self._pending = threading.BoundedSemaphore(4)
        self._next_request_at = 0.0
        self._cooldown_until = 0.0

    def _cached(self, key: tuple) -> dict | None:
        with self._cache_lock:
            item = self._cache.get(key)
            if item is None:
                return None
            expires_at, value = item
            if expires_at <= time.monotonic():
                del self._cache[key]
                return None
            self._cache.move_to_end(key)
            return deepcopy(value)

    def fetch(self, endpoint: str, params: dict) -> dict:
        if endpoint not in ("search", "lookup"):
            raise ValueError("Unsupported Apple endpoint")
        key = (endpoint, tuple(sorted(params.items())))
        cached = self._cached(key)
        if cached is not None:
            return cached

        # Permit a short, bounded queue without occupying the entire thread pool.
        if not self._pending.acquire(blocking=False):
            raise HTTPException(503, "App Store is busy. Please try again shortly.", headers={"Retry-After": "3"})
        if not self._fetch_lock.acquire(timeout=4):
            self._pending.release()
            raise HTTPException(503, "App Store is busy. Please try again shortly.", headers={"Retry-After": "3"})
        try:
            cached = self._cached(key)
            if cached is not None:
                return cached
            cooldown = math.ceil(self._cooldown_until - time.monotonic())
            if cooldown > 0:
                raise HTTPException(503, "App Store request limit reached. Please try again shortly.", headers={"Retry-After": str(cooldown)})
            delay = self._next_request_at - time.monotonic()
            if delay > 0:
                time.sleep(delay)
            self._next_request_at = time.monotonic() + PROVIDER_MIN_INTERVAL_SECONDS
            url = f"{BASE_URL}/{endpoint}"
            source_url = requests.Request("GET", url, params=params).prepare().url
            try:
                response = requests.get(
                    url,
                    params=params,
                    headers={"Accept": "application/json", "User-Agent": "DataProvido-AppBenchmark/1.0"},
                    timeout=(3.05, 10),
                    allow_redirects=False,
                )
            except requests.Timeout as exc:
                raise HTTPException(503, "App Store request timed out. Please try again.", headers={"Retry-After": "5"}) from exc
            except requests.RequestException as exc:
                raise HTTPException(502, "Could not connect to the App Store. Please try again.") from exc

            if response.status_code == 429:
                self._cooldown_until = time.monotonic() + 60
                raise HTTPException(503, "App Store request limit reached. Please retry in a minute.", headers={"Retry-After": "60"})
            if response.status_code >= 500:
                raise HTTPException(503, "App Store is temporarily unavailable. Please try again.", headers={"Retry-After": "10"})
            if response.status_code != 200:
                raise HTTPException(502, "App Store returned an unexpected response.")
            try:
                payload = response.json()
            except ValueError as exc:
                raise HTTPException(502, "App Store returned an unreadable response.") from exc
            if not isinstance(payload, dict) or not isinstance(payload.get("results"), list):
                raise HTTPException(502, "App Store returned an invalid response.")
            if any(not isinstance(item, dict) for item in payload["results"]):
                raise HTTPException(502, "App Store returned invalid app data.")

            value = {
                "results": payload["results"],
                "source_url": source_url,
                "fetched_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            }
            with self._cache_lock:
                self._cache[key] = (time.monotonic() + CACHE_TTL_SECONDS, deepcopy(value))
                self._cache.move_to_end(key)
                while len(self._cache) > CACHE_MAX_ENTRIES:
                    self._cache.popitem(last=False)
            return value
        finally:
            self._fetch_lock.release()
            self._pending.release()

    def fetch_reviews(self, app_id: int, country: str) -> dict:
        """Fetch Apple's public, recent customer-review feed for one storefront."""
        key = ("reviews", country, app_id)
        cached = self._cached(key)
        if cached is not None:
            return cached

        if not self._pending.acquire(blocking=False):
            raise HTTPException(503, "App Store is busy. Please try again shortly.", headers={"Retry-After": "3"})
        if not self._fetch_lock.acquire(timeout=4):
            self._pending.release()
            raise HTTPException(503, "App Store is busy. Please try again shortly.", headers={"Retry-After": "3"})
        try:
            cached = self._cached(key)
            if cached is not None:
                return cached
            cooldown = math.ceil(self._cooldown_until - time.monotonic())
            if cooldown > 0:
                raise HTTPException(503, "App Store request limit reached. Please try again shortly.", headers={"Retry-After": str(cooldown)})
            delay = self._next_request_at - time.monotonic()
            if delay > 0:
                time.sleep(delay)
            self._next_request_at = time.monotonic() + PROVIDER_MIN_INTERVAL_SECONDS
            url = f"{BASE_URL}/{country}/rss/customerreviews/page=1/id={app_id}/sortby=mostrecent/json"
            try:
                response = requests.get(
                    url,
                    headers={"Accept": "application/json", "User-Agent": "DataProvido-AppBenchmark/1.0"},
                    timeout=(3.05, 10),
                    allow_redirects=False,
                )
            except requests.Timeout as exc:
                raise HTTPException(503, "App Store review feed timed out. Please try again.", headers={"Retry-After": "5"}) from exc
            except requests.RequestException as exc:
                raise HTTPException(502, "Could not connect to the App Store review feed. Please try again.") from exc

            fetched_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
            if response.status_code == 404:
                value = {"entries": [], "source_url": url, "fetched_at": fetched_at}
            else:
                if response.status_code == 429:
                    self._cooldown_until = time.monotonic() + 60
                    raise HTTPException(503, "App Store request limit reached. Please retry in a minute.", headers={"Retry-After": "60"})
                if response.status_code >= 500:
                    raise HTTPException(503, "App Store review feed is temporarily unavailable. Please try again.", headers={"Retry-After": "10"})
                if response.status_code != 200:
                    raise HTTPException(502, "App Store review feed returned an unexpected response.")
                try:
                    payload = response.json()
                except ValueError as exc:
                    raise HTTPException(502, "App Store review feed returned an unreadable response.") from exc
                feed = payload.get("feed") if isinstance(payload, dict) else None
                if not isinstance(feed, dict):
                    raise HTTPException(502, "App Store review feed returned an invalid response.")
                entries = feed.get("entry", [])
                if isinstance(entries, dict):
                    entries = [entries]
                if not isinstance(entries, list) or any(not isinstance(item, dict) for item in entries):
                    raise HTTPException(502, "App Store review feed returned invalid review data.")
                value = {"entries": entries, "source_url": url, "fetched_at": fetched_at}

            with self._cache_lock:
                self._cache[key] = (time.monotonic() + CACHE_TTL_SECONDS, deepcopy(value))
                self._cache.move_to_end(key)
                while len(self._cache) > CACHE_MAX_ENTRIES:
                    self._cache.popitem(last=False)
            return value
        finally:
            self._fetch_lock.release()
            self._pending.release()


provider = AppleProvider()


def _text(value: Any) -> str | None:
    return value.strip() if isinstance(value, str) and value.strip() else None


def _number(value: Any, *, integer: bool = False, maximum: float | None = None) -> float | int | None:
    if value is None or isinstance(value, bool) or not isinstance(value, (int, float, str)):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    if not math.isfinite(number) or number < 0 or (maximum is not None and number > maximum):
        return None
    if integer:
        return int(number) if number.is_integer() else None
    return number


def _url(value: Any) -> str | None:
    text = _text(value)
    if text:
        try:
            parsed = urlparse(text)
            if parsed.scheme == "https" and parsed.hostname and not parsed.username and not parsed.password:
                return text
        except ValueError:
            pass
    return None


def _label(record: dict, key: str) -> Any:
    value = record.get(key)
    return value.get("label") if isinstance(value, dict) else None


def normalize_review(record: dict) -> dict | None:
    rating = _number(_label(record, "im:rating"), integer=True, maximum=5)
    body = _text(_label(record, "content"))
    if rating is None or rating < 1 or not body:
        return None
    title = _text(_label(record, "title"))
    created_at = _text(_label(record, "updated"))
    review_id = _text(_label(record, "id"))
    sentiment = "positive" if rating >= 4 else "neutral" if rating == 3 else "negative"
    return {
        "id": review_id,
        "title": title[:300] if title else None,
        "body": body[:5000],
        "rating": rating,
        "sentiment": sentiment,
        "created_at": created_at,
    }


def review_summary(entries: list[dict]) -> dict:
    reviews = [review for entry in entries if (review := normalize_review(entry)) is not None]
    counts = {"positive": 0, "neutral": 0, "negative": 0}
    for review in reviews:
        counts[review["sentiment"]] += 1
    sample_count = len(reviews)
    percentages = {key: round(value * 100 / sample_count, 1) if sample_count else 0.0 for key, value in counts.items()}
    average = round(sum(review["rating"] for review in reviews) / sample_count, 2) if sample_count else None
    return {
        "sample_count": sample_count,
        "average_rating": average,
        "sentiment": counts,
        "sentiment_percent": percentages,
        "reviews": reviews,
    }


def normalize_app(record: dict, *, country: str, fetched_at: str, source_url: str) -> dict | None:
    """Keep absent observations null, including a genuinely unknown language list."""
    app_id = _number(record.get("trackId"), integer=True)
    if not app_id or record.get("kind", "software") != "software":
        return None
    raw_languages = record.get("languageCodesISO2A")
    languages = None
    if isinstance(raw_languages, list):
        languages = list(dict.fromkeys(language.strip() for language in raw_languages if isinstance(language, str) and language.strip()))
    return {
        "id": app_id,
        "name": _text(record.get("trackName")),
        "developer": _text(record.get("artistName")) or _text(record.get("sellerName")),
        "icon_url": _url(record.get("artworkUrl512")) or _url(record.get("artworkUrl100")) or _url(record.get("artworkUrl60")),
        "store_url": _url(record.get("trackViewUrl")),
        "bundle_id": _text(record.get("bundleId")),
        "category": _text(record.get("primaryGenreName")),
        "rating": _number(record.get("averageUserRating"), maximum=5),
        "rating_count": _number(record.get("userRatingCount"), integer=True),
        "current_rating": _number(record.get("averageUserRatingForCurrentVersion"), maximum=5),
        "current_rating_count": _number(record.get("userRatingCountForCurrentVersion"), integer=True),
        "price": _number(record.get("price")),
        "currency": _text(record.get("currency")),
        "version": _text(record.get("version")),
        "updated_at": _text(record.get("currentVersionReleaseDate")),
        "released_at": _text(record.get("releaseDate")),
        "size_bytes": _number(record.get("fileSizeBytes"), integer=True),
        "min_os": _text(record.get("minimumOsVersion")),
        "languages": languages,
        "content_rating": _text(record.get("contentAdvisoryRating")) or _text(record.get("trackContentRating")),
        "description": _text(record.get("description")),
        "release_notes": _text(record.get("releaseNotes")),
        "retrieved_at": fetched_at,
        "source_url": source_url,
        "country": country,
    }


def _apps(payload: dict, country: str) -> list[dict]:
    result = []
    seen = set()
    for record in payload["results"]:
        app = normalize_app(record, country=country, fetched_at=payload["fetched_at"], source_url=payload["source_url"])
        if app is not None and app["id"] not in seen:
            result.append(app)
            seen.add(app["id"])
    return result


@router.get("/search")
def search_apps(q: Annotated[str, Query(min_length=2, max_length=100)], country: Country = "tr") -> dict:
    query = q.strip()
    if len(query) < 2:
        raise HTTPException(422, "Enter at least 2 characters to search for an app.")
    payload = provider.fetch("search", {"term": query, "country": country, "media": "software", "entity": "software", "limit": 20})
    return {
        "apps": _apps(payload, country),
        "country": country,
        "source": SOURCE,
        "source_url": payload["source_url"],
        "fetched_at": payload["fetched_at"],
    }


@router.post("/compare")
def compare_apps(body: CompareRequest) -> dict:
    payload = provider.fetch("lookup", {"id": ",".join(str(app_id) for app_id in body.app_ids), "entity": "software", "country": body.country})
    by_id = {app["id"]: app for app in _apps(payload, body.country)}
    return {
        "apps": [by_id[app_id] for app_id in body.app_ids if app_id in by_id],
        "country": body.country,
        "source": SOURCE,
        "source_url": payload["source_url"],
        "fetched_at": payload["fetched_at"],
        "missing_app_ids": [app_id for app_id in body.app_ids if app_id not in by_id],
    }


@router.post("/reviews")
def compare_reviews(body: ReviewsRequest) -> dict:
    results = []
    for app_id in body.app_ids:
        payload = provider.fetch_reviews(app_id, body.country)
        summary = review_summary(payload["entries"])
        results.append({
            "app_id": app_id,
            **summary,
            "source_url": payload["source_url"],
            "fetched_at": payload["fetched_at"],
        })
    return {
        "reviews": results,
        "country": body.country,
        "source": "Apple App Store public customer-review feed",
        "method": "Rating-based sentiment: 4–5 positive, 3 neutral, 1–2 negative.",
        "sample_limit_per_app": 50,
    }
