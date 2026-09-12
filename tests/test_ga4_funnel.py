import unittest
from unittest.mock import Mock, patch

from fastapi import HTTPException

from functions.ga4_commerce import GoogleAnalytics
from functions.ga4_funnel import PRESETS, funnel_report, sample_report


def funnel_payload(multiplier=1):
    totals = [1000, 600, 300, 180, 100]
    table_rows = []
    for index, users in enumerate(totals, 1):
        next_users = totals[index] if index < len(totals) else users
        abandoned = users - next_users if index < len(totals) else 0
        for group, share in (("RESERVED_TOTAL", 1), ("mobile", .6), ("desktop", .4)):
            current = int(users * share * multiplier)
            next_current = int(next_users * share * multiplier)
            table_rows.append({"dimensionValues": [{"value": f"{index}. {PRESETS['ecommerce'][index-1][0]}"}, {"value": group}],
                               "metricValues": [{"value": str(current)}, {"value": str(next_current/current if current and index < len(totals) else 0)},
                                                {"value": str(int(abandoned*share*multiplier))}, {"value": str((current-next_current)/current if current and index < len(totals) else 0)}]})
    return {
        "funnelTable": {"dimensionHeaders": [{"name": "funnelStepName"}, {"name": "deviceCategory"}],
                        "metricHeaders": [{"name": "activeUsers"}, {"name": "funnelStepCompletionRate"}, {"name": "funnelStepAbandonments"}, {"name": "funnelStepAbandonmentRate"}], "rows": table_rows},
        "funnelVisualization": {"dimensionHeaders": [{"name": "funnelStepName"}, {"name": "funnelStepNextAction"}],
                                "metricHeaders": [{"name": "activeUsers"}], "rows": [
                                    {"dimensionValues": [{"value": "1. Session started"}, {"value": "page_view"}], "metricValues": [{"value": str(500*multiplier)}]},
                                    {"dimensionValues": [{"value": "1. Session started"}, {"value": "scroll"}], "metricValues": [{"value": str(250*multiplier)}]},
                                ]},
    }


class FixtureAnalytics:
    def __init__(self):
        self.funnel_calls = []

    def property(self, pid):
        return {"id": pid, "name": "Store", "currency": "USD", "time_zone": "UTC"}

    def funnel(self, pid, body):
        self.funnel_calls.append(body)
        return funnel_payload(1 if body["dateRanges"][0]["startDate"] == "2026-08-01" else .8)

    def report(self, pid, body):
        dimension = body["dimensions"][0]["name"]
        if dimension == "newVsReturning":
            return {"dimensionHeaders": [{"name": "dateRange"}, {"name": "newVsReturning"}],
                    "metricHeaders": [{"name": name} for name in body_metric_names(body)], "rows": [
                        {"dimensionValues": [{"value": "current"}, {"value": "new"}], "metricValues": [{"value": value} for value in ("700", "800", "500", ".625", ".375", "91", "5.2", "60", "12000")]},
                        {"dimensionValues": [{"value": "previous"}, {"value": "new"}], "metricValues": [{"value": "1"} for _ in body_metric_names(body)]},
                    ]}
        return {"dimensionHeaders": [{"name": "landingPagePlusQueryString"}],
                "metricHeaders": [{"name": name} for name in body_metric_names(body)], "rows": [
                    {"dimensionValues": [{"value": "/products/one"}], "metricValues": [{"value": value} for value in ("300", "350", ".7", ".3", "120", "20")]}
                ]}


def body_metric_names(body):
    return [item["name"] for item in body["metrics"]]


class FunnelReportTests(unittest.TestCase):
    def test_ordered_funnel_paths_breakdown_and_users_are_parsed(self):
        provider = FixtureAnalytics()
        report = funnel_report(provider, "123", start_date="2026-08-01", end_date="2026-08-30")
        self.assertEqual(report["source"], "ga4")
        self.assertEqual(report["steps"][0]["users"], 1000)
        self.assertEqual(report["summary"]["conversion_rate"], .1)
        self.assertAlmostEqual(report["summary"]["previous_conversion_rate"], .1)
        self.assertEqual(report["breakdown_rows"][0]["name"], "mobile")
        self.assertEqual(report["next_actions"]["1"][0]["name"], "page_view")
        self.assertEqual(report["user_segments"][0]["name"], "new")
        self.assertEqual(report["landing_pages"][0]["path"], "/products/one")
        self.assertEqual(sum("funnelNextAction" in body for body in provider.funnel_calls), 1)

    def test_optional_explorations_fail_without_hiding_the_funnel(self):
        provider = FixtureAnalytics()
        provider.report = Mock(side_effect=HTTPException(status_code=422, detail={"message": "Unsupported metric."}))
        report = funnel_report(provider, "123", start_date="2026-08-01", end_date="2026-08-30")
        self.assertEqual(report["steps"][0]["users"], 1000)
        self.assertEqual(report["user_segments"], [])
        self.assertEqual(report["landing_pages"], [])
        self.assertTrue(any("Unsupported metric" in warning for warning in report["warnings"]))

    def test_sample_is_explicit_and_labelled(self):
        report = sample_report(start_date="2026-08-01", end_date="2026-08-30")
        self.assertEqual(report["source"], "sample")
        self.assertTrue(report["steps"])
        self.assertTrue(any("illustrative" in warning for warning in report["warnings"]))

    @patch("functions.ga4_commerce.requests.request")
    def test_provider_uses_funnel_reporting_endpoint(self, request):
        request.return_value = Mock(status_code=200, json=lambda: {"funnelTable": {}, "funnelVisualization": {}})
        GoogleAnalytics("token").funnel("123", {"funnel": {}})
        self.assertIn("/v1alpha/properties/123:runFunnelReport", request.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
