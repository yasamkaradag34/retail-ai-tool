import unittest
from datetime import date
from unittest.mock import Mock, patch

from fastapi import HTTPException

from functions.merchant_insights import (GoogleMerchant, PRICE_QUERY, PRODUCT_QUERY,
                                         account_id, date_ranges, merchant_report,
                                         performance_query, sample_report)


class FixtureMerchant:
    def __init__(self):
        self.calls = []
        self.fail_prices = False

    def account(self, value):
        self.calls.append(("account", value))
        return {"id": value, "name": "Store"}

    def search(self, value, query, limit=10000):
        self.calls.append(("search", value, query))
        if "FROM product_view" in query:
            return ([{"productView": {"id": "online~en~US~SKU1", "offerId": "SKU1", "title": "Shoes",
                                       "brand": "Run", "categoryL1": "Sport", "availability": "OUT_OF_STOCK",
                                       "price": {"amountMicros": "120000000", "currencyCode": "USD"},
                                       "aggregatedReportingContextStatus": "ELIGIBLE", "clickPotential": "HIGH"}}], True)
        if "FROM price_competitiveness_product_view" in query:
            if self.fail_prices:
                raise HTTPException(403, {"code": "permission", "message": "no"})
            return ([{"priceCompetitivenessProductView": {"id": "online~en~US~SKU1", "offerId": "SKU1", "title": "Shoes",
                                                            "brand": "Run", "categoryL1": "Sport", "reportCountryCode": "US",
                                                            "price": {"amountMicros": "120000000", "currencyCode": "USD"},
                                                            "benchmarkPrice": {"amountMicros": "100000000", "currencyCode": "USD"}}}], True)
        if "FROM price_insights_product_view" in query:
            return ([{"priceInsightsProductView": {"offerId": "SKU1", "suggestedPrice": {"amountMicros": "105000000", "currencyCode": "USD"},
                                                     "effectiveness": "HIGH", "predictedClicksChangeFraction": .1}}], True)
        current = "2026-08-01" in query
        clicks, conversions, value = (100, 5, 500) if current else (80, 4, 400)
        return ([{"productPerformanceView": {"offerId": "SKU1", "title": "Shoes", "brand": "Run", "categoryL1": "Sport",
                                               "marketingMethod": "ORGANIC", "clicks": str(clicks), "impressions": "1000", "conversions": str(conversions),
                                               "conversionValue": {"amountMicros": str(value * 1_000_000), "currencyCode": "USD"}}}], True)


class ValidationTests(unittest.TestCase):
    def test_account_id_rejects_non_numeric_values(self):
        for value in ("", "accounts/123", "../123", "0", "1:search"):
            with self.subTest(value=value), self.assertRaises(HTTPException):
                account_id(value)

    def test_equal_previous_period(self):
        result = date_ranges(start_date="2026-08-01", end_date="2026-08-30", today=date(2026, 9, 1))
        self.assertEqual((result["previous_start"], result["previous_end"]), ("2026-07-02", "2026-07-31"))

    def test_queries_use_current_merchant_api_names(self):
        self.assertIn("FROM product_view", PRODUCT_QUERY)
        self.assertIn("FROM price_competitiveness_product_view", PRICE_QUERY)
        query = performance_query("2026-08-01", "2026-08-30")
        self.assertIn("conversion_value", query)
        self.assertIn("WHERE date BETWEEN '2026-08-01' AND '2026-08-30'", query)


class ReportTests(unittest.TestCase):
    def test_catalog_price_performance_and_suggestion_are_joined(self):
        provider = FixtureMerchant()
        report = merchant_report(provider, "123", start_date="2026-08-01", end_date="2026-08-30")
        row = report["rows"][0]
        self.assertEqual(row["availability"], "out_of_stock")
        self.assertEqual(row["price_status"], "above")
        self.assertEqual(row["price_gap_percent"], 20)
        self.assertEqual(row["performance"]["clicks"], 100)
        self.assertEqual(row["previous"]["clicks"], 80)
        self.assertEqual(row["performance"]["conversion_value"], 500)
        self.assertEqual(row["performance"]["conversion_currency"], "USD")
        self.assertEqual(row["suggested_price"]["amount"], 105)
        self.assertEqual(report["summary"]["out_of_stock"], 1)
        self.assertEqual(report["summary"]["price_positions"]["above"]["count"], 1)
        self.assertEqual(report["opportunities"]["out_of_stock_demand"][0]["offer_id"], "SKU1")

    def test_optional_market_report_failure_never_substitutes_sample(self):
        provider = FixtureMerchant()
        provider.fail_prices = True
        report = merchant_report(provider, "123", start_date="2026-08-01", end_date="2026-08-30")
        self.assertEqual(report["source"], "merchant_api")
        self.assertEqual(report["summary"]["benchmark_products"], 0)
        self.assertEqual(report["rows"][0]["price_status"], "unavailable")
        self.assertTrue(any("Price competitiveness is unavailable" in warning for warning in report["warnings"]))

    def test_sample_is_explicit_and_labelled(self):
        report = sample_report(start_date="2026-08-01", end_date="2026-08-30")
        self.assertEqual(report["source"], "sample")
        self.assertTrue(report["rows"])
        self.assertTrue(any("illustrative" in warning for warning in report["warnings"]))


class ProviderTests(unittest.TestCase):
    @patch("functions.merchant_insights.requests.request")
    def test_accounts_paginate_without_demo_fallback(self, request):
        request.side_effect = [
            Mock(status_code=200, json=lambda: {"accounts": [], "nextPageToken": "next"}),
            Mock(status_code=200, json=lambda: {"accounts": [{"name": "accounts/123", "accountName": "Store"}]}),
        ]
        self.assertEqual(GoogleMerchant("token").accounts(), [{"id": "123", "name": "Store"}])
        self.assertEqual(request.call_args.kwargs["params"]["pageToken"], "next")

    @patch("functions.merchant_insights.requests.request")
    def test_provider_errors_are_safe(self, request):
        for upstream, expected in ((401, 401), (403, 403), (429, 429), (400, 422), (500, 502)):
            request.return_value = Mock(status_code=upstream, text="private provider response")
            with self.subTest(upstream=upstream), self.assertRaises(HTTPException) as raised:
                GoogleMerchant("secret").accounts()
            self.assertEqual(raised.exception.status_code, expected)
            self.assertNotIn("private provider response", str(raised.exception.detail))


if __name__ == "__main__":
    unittest.main()
