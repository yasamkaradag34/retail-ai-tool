"""Authentication integration tests for the mounted App Benchmark router.

Run from the project root after integrating app_benchmark.py and its auth
dependency: python -m unittest discover -s tests -p test_benchmark_auth.py -v
No login request, password, .env file, or external provider request is used.
"""

import json
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

from functions import app_benchmark as benchmark
import main


class BenchmarkAuthenticationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        paths = {getattr(route, "path", None) for route in main.app.routes}
        expected = {"/api/app-benchmark/search", "/api/app-benchmark/compare", "/api/app-benchmark/reviews"}
        if not expected.issubset(paths):
            raise AssertionError("Integrate the authenticated App Benchmark router before running these tests.")
        if not main.ALLOWED_LOGIN_EMAILS:
            raise AssertionError("The application must define its allowed login emails.")
        cls.allowed_email = sorted(main.ALLOWED_LOGIN_EMAILS)[0]

    def setUp(self):
        self.client = TestClient(main.app, base_url="https://testserver")
        self.addCleanup(self.client.close)
        self.provider_patch = patch.object(benchmark, "provider")
        self.provider = self.provider_patch.start()
        self.addCleanup(self.provider_patch.stop)
        self.provider.fetch.return_value = {
            "results": [{"kind": "software", "trackId": 11, "trackName": "First"}, {"kind": "software", "trackId": 22, "trackName": "Second"}],
            "source_url": "https://itunes.apple.com/lookup?id=11,22&country=tr&entity=software",
            "fetched_at": "2026-09-11T10:00:00Z",
        }
        self.provider.fetch_reviews.return_value = {
            "entries": [],
            "source_url": "https://itunes.apple.com/tr/rss/customerreviews/page=1/id=11/sortby=mostrecent/json",
            "fetched_at": "2026-09-11T10:00:00Z",
        }

    def signed_cookie(self, payload):
        return main._encrypt_token(json.dumps(payload))

    def call_routes(self):
        return (
            self.client.get("/api/app-benchmark/search", params={"q": "First", "country": "tr"}),
            self.client.post("/api/app-benchmark/compare", json={"app_ids": [11, 22], "country": "tr"}),
            self.client.post("/api/app-benchmark/reviews", json={"app_ids": [11, 22], "country": "tr"}),
        )

    def assert_unauthorized_without_provider(self, cookie=None):
        self.client.cookies.clear()
        if cookie is not None:
            self.client.cookies.set("gauth", cookie)
        for response in self.call_routes():
            self.assertEqual(response.status_code, 401, response.text)
        self.provider.fetch.assert_not_called()

    def test_missing_cookie_returns_401_before_provider_access(self):
        self.assert_unauthorized_without_provider()

    def test_malformed_or_tampered_cookie_returns_401(self):
        valid = self.signed_cookie({"email": self.allowed_email})
        tampered = ("0" if valid[0] != "0" else "1") + valid[1:]
        cookies = ["not-a-cookie", "invalid.invalid", tampered, main._encrypt_token("not JSON")]
        for cookie in cookies:
            with self.subTest(cookie_type=cookies.index(cookie)):
                self.assert_unauthorized_without_provider(cookie)

    def test_valid_signature_with_wrong_payload_shape_returns_401(self):
        for payload in (None, [], "string", 7, {}, {"email": None}, {"email": 7}, {"email": []}, {"email": {"value": self.allowed_email}}):
            with self.subTest(payload_type=type(payload).__name__):
                self.assert_unauthorized_without_provider(self.signed_cookie(payload))

    def test_valid_disallowed_email_returns_401_before_provider_access(self):
        self.assert_unauthorized_without_provider(self.signed_cookie({"email": "not-allowed@example.invalid"}))

    def test_signed_allowed_email_can_search_and_compare(self):
        self.client.cookies.set("gauth", self.signed_cookie({"email": self.allowed_email}))
        search, compare, reviews = self.call_routes()
        self.assertEqual(search.status_code, 200, search.text)
        self.assertEqual(compare.status_code, 200, compare.text)
        self.assertEqual(reviews.status_code, 200, reviews.text)
        self.assertEqual([app["id"] for app in compare.json()["apps"]], [11, 22])
        self.assertEqual(compare.json()["missing_app_ids"], [])
        self.assertEqual(self.provider.fetch.call_count, 2)
        self.assertEqual(self.provider.fetch_reviews.call_count, 2)


if __name__ == "__main__":
    unittest.main()
