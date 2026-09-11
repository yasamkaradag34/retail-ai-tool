"""Focused offline API/normalization/provider tests: python -m unittest -v."""

from concurrent.futures import ThreadPoolExecutor
import threading
import time
import unittest
from unittest.mock import Mock, patch

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
import requests

from functions import app_benchmark as benchmark


FETCHED_AT = "2026-09-11T10:00:00Z"
SOURCE_URL = "https://itunes.apple.com/lookup?id=11,22&country=tr&entity=software"


def payload(records):
    return {"results": records, "source_url": SOURCE_URL, "fetched_at": FETCHED_AT}


def response(status=200, records=None):
    result = Mock(status_code=status)
    result.json.return_value = {"results": records or []}
    return result


def review_entry(rating="5", title="Great", body="Very useful", updated="2026-09-10T09:00:00Z"):
    return {
        "id": {"label": "review-1"},
        "im:rating": {"label": rating},
        "title": {"label": title},
        "content": {"label": body},
        "updated": {"label": updated},
    }


class NormalizationTests(unittest.TestCase):
    def normalize(self, record):
        return benchmark.normalize_app(record, country="tr", fetched_at=FETCHED_AT, source_url=SOURCE_URL)

    def test_missing_values_are_unknown_not_fabricated_zero(self):
        app = self.normalize({"trackId": 11, "trackName": "First app"})
        for field in ("rating", "rating_count", "current_rating", "current_rating_count", "price", "size_bytes", "languages"):
            self.assertIsNone(app[field], field)
        self.assertEqual(app["retrieved_at"], FETCHED_AT)
        self.assertEqual(app["source_url"], SOURCE_URL)

    def test_real_zeros_and_public_metadata_survive(self):
        app = self.normalize({"trackId": 11, "userRatingCount": 0, "price": 0, "fileSizeBytes": "32000000", "averageUserRating": 4.7, "languageCodesISO2A": ["TR", "EN", "TR"], "minimumOsVersion": "16.0"})
        self.assertEqual(app["rating_count"], 0)
        self.assertEqual(app["price"], 0)
        self.assertEqual(app["size_bytes"], 32000000)
        self.assertEqual(app["rating"], 4.7)
        self.assertEqual(app["languages"], ["TR", "EN"])
        self.assertEqual(app["min_os"], "16.0")

    def test_invalid_provider_metrics_are_null_and_nonapps_ignored(self):
        app = self.normalize({"trackId": 11, "userRatingCount": -3, "price": "NaN", "averageUserRating": 6, "fileSizeBytes": False, "trackViewUrl": "javascript:alert(1)"})
        for field in ("rating_count", "price", "rating", "size_bytes", "store_url"):
            self.assertIsNone(app[field])
        self.assertIsNone(self.normalize({"trackId": 11, "kind": "song"}))

    def test_review_summary_uses_written_review_stars_and_ignores_invalid_entries(self):
        entries = [
            review_entry("5", "Loved it", "Fast and simple"),
            review_entry("3", "Okay", "It works"),
            review_entry("1", "Problem", "Crashes often"),
            review_entry("9", "Invalid", "Invalid rating"),
            {"im:rating": {"label": "4"}, "content": {"label": ""}},
        ]
        summary = benchmark.review_summary(entries)
        self.assertEqual(summary["sample_count"], 3)
        self.assertEqual(summary["sentiment"], {"positive": 1, "neutral": 1, "negative": 1})
        self.assertEqual(summary["sentiment_percent"], {"positive": 33.3, "neutral": 33.3, "negative": 33.3})
        self.assertEqual(summary["average_rating"], 3.0)
        self.assertEqual([item["sentiment"] for item in summary["reviews"]], ["positive", "neutral", "negative"])


class ApiTests(unittest.TestCase):
    def setUp(self):
        app = FastAPI()
        app.include_router(benchmark.router)
        self.client = TestClient(app)
        self.provider_patch = patch.object(benchmark, "provider")
        self.provider = self.provider_patch.start()
        self.addCleanup(self.provider_patch.stop)

    def test_search_parameters_and_country_are_explicit(self):
        self.provider.fetch.return_value = payload([{"trackId": 11, "trackName": "First"}])
        result = self.client.get("/api/app-benchmark/search", params={"q": "  grocery  ", "country": "gb"})
        self.assertEqual(result.status_code, 200)
        self.provider.fetch.assert_called_once_with("search", {"term": "grocery", "country": "gb", "media": "software", "entity": "software", "limit": 20})
        self.assertEqual(result.json()["apps"][0]["country"], "gb")

    def test_lookup_preserves_order_and_identifies_country_missing_apps(self):
        self.provider.fetch.return_value = payload([{"trackId": 22}, {"trackId": 11}, {"trackId": 999}])
        result = self.client.post("/api/app-benchmark/compare", json={"app_ids": [11, 33, 22], "country": "tr"})
        self.assertEqual(result.status_code, 200)
        self.assertEqual([app["id"] for app in result.json()["apps"]], [11, 22])
        self.assertEqual(result.json()["missing_app_ids"], [33])
        self.provider.fetch.assert_called_once_with("lookup", {"id": "11,33,22", "entity": "software", "country": "tr"})

    def test_empty_country_lookup_keeps_all_missing_ids(self):
        self.provider.fetch.return_value = payload([])
        result = self.client.post("/api/app-benchmark/compare", json={"app_ids": [11, 22]})
        self.assertEqual(result.json()["apps"], [])
        self.assertEqual(result.json()["missing_app_ids"], [11, 22])

    def test_reviews_preserve_app_order_and_report_public_sample(self):
        self.provider.fetch_reviews.side_effect = [
            {"entries": [review_entry("5"), review_entry("4")], "source_url": "https://itunes.apple.com/tr/rss/first", "fetched_at": FETCHED_AT},
            {"entries": [review_entry("2")], "source_url": "https://itunes.apple.com/tr/rss/second", "fetched_at": FETCHED_AT},
        ]
        result = self.client.post("/api/app-benchmark/reviews", json={"app_ids": [11, 22], "country": "tr"})
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual([item["app_id"] for item in result.json()["reviews"]], [11, 22])
        self.assertEqual(result.json()["reviews"][0]["sentiment"]["positive"], 2)
        self.assertEqual(result.json()["reviews"][1]["sentiment"]["negative"], 1)
        self.assertEqual(result.json()["sample_limit_per_app"], 50)
        self.assertEqual(self.provider.fetch_reviews.call_count, 2)

    def test_invalid_input_never_reaches_provider(self):
        invalid_bodies = [
            {"app_ids": [11]}, {"app_ids": [11, 11]}, {"app_ids": [1, 2, 3, 4, 5, 6]},
            {"app_ids": [True, 22]}, {"app_ids": ["11", 22]}, {"app_ids": [11.0, 22]},
            {"app_ids": [0, 22]}, {"app_ids": [-11, 22]},
            {"app_ids": ["http://127.0.0.1", 22]}, {"app_ids": [11, 22], "country": "xx"},
            {"app_ids": [11, 22], "url": "https://example.com"},
        ]
        for body in invalid_bodies:
            with self.subTest(body=body):
                self.assertEqual(self.client.post("/api/app-benchmark/compare", json=body).status_code, 422)
                self.assertEqual(self.client.post("/api/app-benchmark/reviews", json=body).status_code, 422)
        for params in ({"q": "a"}, {"q": "  "}, {"q": "x" * 101}, {"q": "grocery", "country": "xx"}):
            self.assertEqual(self.client.get("/api/app-benchmark/search", params=params).status_code, 422)
        self.provider.fetch.assert_not_called()
        self.provider.fetch_reviews.assert_not_called()

    def test_provider_failure_is_visible_not_an_empty_success(self):
        self.provider.fetch.side_effect = HTTPException(503, "App Store unavailable", headers={"Retry-After": "10"})
        result = self.client.get("/api/app-benchmark/search", params={"q": "grocery"})
        self.assertEqual(result.status_code, 503)
        self.assertEqual(result.headers["retry-after"], "10")
        self.assertIn("unavailable", result.json()["detail"])


class ProviderTests(unittest.TestCase):
    def setUp(self):
        self.provider = benchmark.AppleProvider()

    @patch.object(benchmark, "PROVIDER_MIN_INTERVAL_SECONDS", 0)
    @patch.object(benchmark.requests, "get")
    def test_cache_reuses_result_and_does_not_leak_mutation(self, get):
        get.return_value = response(records=[{"trackId": 11}])
        first = self.provider.fetch("search", {"term": "First", "country": "tr"})
        first["results"][0]["trackId"] = 99
        second = self.provider.fetch("search", {"term": "First", "country": "tr"})
        self.assertEqual(second["results"][0]["trackId"], 11)
        self.assertEqual(get.call_count, 1)
        self.assertEqual(get.call_args.kwargs["timeout"], (3.05, 10))
        self.assertFalse(get.call_args.kwargs["allow_redirects"])

    @patch.object(benchmark, "PROVIDER_MIN_INTERVAL_SECONDS", 0)
    @patch.object(benchmark.requests, "get")
    def test_cache_is_bounded_and_expires(self, get):
        get.return_value = response()
        with patch.object(benchmark, "CACHE_MAX_ENTRIES", 2):
            self.provider.fetch("search", {"term": "first"})
            self.provider.fetch("search", {"term": "second"})
            self.provider.fetch("search", {"term": "third"})
        self.assertEqual(len(self.provider._cache), 2)
        key = ("search", (("term", "third"),))
        self.provider._cache[key] = (0, payload([]))
        self.provider.fetch("search", {"term": "third"})
        self.assertEqual(get.call_count, 4)

    @patch.object(benchmark, "PROVIDER_MIN_INTERVAL_SECONDS", 0)
    @patch.object(benchmark.requests, "get")
    def test_concurrent_identical_misses_only_fetch_once(self, get):
        gate = threading.Barrier(2)
        def slow_response(*args, **kwargs):
            time.sleep(0.08)
            return response(records=[{"trackId": 11}])
        get.side_effect = slow_response
        def fetch():
            gate.wait()
            return self.provider.fetch("lookup", {"id": "11,22", "country": "tr"})
        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = [executor.submit(fetch) for _ in range(2)]
            self.assertEqual(futures[0].result(), futures[1].result())
        self.assertEqual(get.call_count, 1)

    @patch.object(benchmark, "PROVIDER_MIN_INTERVAL_SECONDS", 0)
    @patch.object(benchmark.requests, "get")
    def test_http_and_parse_failures_are_not_cached(self, get):
        cases = [(response(429), 503), (response(500), 503), (response(404), 502)]
        invalid = response()
        invalid.json.return_value = {"no-results": []}
        cases.append((invalid, 502))
        malformed = response()
        malformed.json.side_effect = ValueError("not json")
        cases.append((malformed, 502))
        for remote, expected_status in cases:
            self.provider = benchmark.AppleProvider()
            get.return_value = remote
            with self.subTest(status=remote.status_code), self.assertRaises(HTTPException) as caught:
                self.provider.fetch("search", {"term": "First"})
            self.assertEqual(caught.exception.status_code, expected_status)
            self.assertEqual(len(self.provider._cache), 0)

    @patch.object(benchmark, "PROVIDER_MIN_INTERVAL_SECONDS", 0)
    @patch.object(benchmark.requests, "get")
    def test_timeout_and_connection_errors_are_visible(self, get):
        for error, status in [(requests.Timeout(), 503), (requests.ConnectionError(), 502)]:
            get.side_effect = error
            with self.assertRaises(HTTPException) as caught:
                self.provider.fetch("search", {"term": "First"})
            self.assertEqual(caught.exception.status_code, status)
        self.assertEqual(len(self.provider._cache), 0)

    @patch.object(benchmark.requests, "get")
    def test_throttled_provider_is_not_immediately_hammered_again(self, get):
        get.return_value = response(429)
        for term in ("First", "Second"):
            with self.assertRaises(HTTPException) as caught:
                self.provider.fetch("search", {"term": term})
            self.assertEqual(caught.exception.status_code, 503)
            self.assertIn("Retry-After", caught.exception.headers)
        self.assertEqual(get.call_count, 1)

    @patch.object(benchmark, "PROVIDER_MIN_INTERVAL_SECONDS", 0)
    @patch.object(benchmark.requests, "get")
    def test_public_review_feed_is_parsed_cached_and_bounded_to_a_fixed_url(self, get):
        remote = Mock(status_code=200)
        remote.json.return_value = {"feed": {"entry": [review_entry("5")]}}
        get.return_value = remote
        first = self.provider.fetch_reviews(12345, "tr")
        first["entries"].clear()
        second = self.provider.fetch_reviews(12345, "tr")
        self.assertEqual(len(second["entries"]), 1)
        self.assertEqual(get.call_count, 1)
        self.assertEqual(get.call_args.args[0], "https://itunes.apple.com/tr/rss/customerreviews/page=1/id=12345/sortby=mostrecent/json")
        self.assertFalse(get.call_args.kwargs["allow_redirects"])

    @patch.object(benchmark, "PROVIDER_MIN_INTERVAL_SECONDS", 0)
    @patch.object(benchmark.requests, "get")
    def test_missing_public_review_feed_becomes_an_empty_sample(self, get):
        get.return_value = Mock(status_code=404)
        result = self.provider.fetch_reviews(12345, "tr")
        self.assertEqual(result["entries"], [])


if __name__ == "__main__":
    unittest.main()
