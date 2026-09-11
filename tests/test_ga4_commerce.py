import copy
from datetime import date
import unittest
from unittest.mock import Mock, patch

from fastapi import HTTPException
from functions.ga4_commerce import (GoogleAnalytics, ITEM_METRICS, QUALITY_METRICS, ROW_LIMIT,
                                    commerce_report, date_ranges, parsed_rows)


def report(dimensions, metrics, rows, **extra):
    return {"dimensionHeaders": [{"name": d} for d in dimensions], "metricHeaders": [{"name": m} for m in metrics],
            "rows": [{"dimensionValues": [{"value": v} for v in ds], "metricValues": [{"value": str(v)} for v in ms]} for ds, ms in rows], "rowCount": len(rows), **extra}


class FixtureAnalytics:
    def __init__(self):
        self.calls = []
        self.compatible = QUALITY_METRICS.copy()
        self.quality_failure = False
        self.empty = False
        self.restrict = False
        self.limited = False

    def property(self, pid):
        return {"id":pid,"name":"Test store","currency":"EUR","time_zone":"Europe/Istanbul"}

    def compatible_quality(self, pid, dimensions, dimension_filter):
        self.calls.append(("compatibility", dimensions, dimension_filter))
        return self.compatible

    def report(self, pid, body):
        self.calls.append(copy.deepcopy(body))
        dims = [d["name"] for d in body["dimensions"]]
        metrics = [m["name"] for m in body["metrics"]]
        if self.empty:
            return report(dims + ["dateRange"], metrics, [])
        if dims == ["eventName"]:
            return report(dims+["dateRange"],metrics,[(["view_item","current"],[120]),(["purchase","current"],[3])])
        if metrics[0] == "sessions":
            if self.quality_failure:
                raise HTTPException(422,{"code":"report_incompatible","message":"Unavailable."})
            values = [20,15,12,.4,.6,165]
        else:
            values = [100,30,8,400,.22,.05]
        values = values[:len(metrics)]
        dimension_values = ["shoes"] if dims == ["itemCategory"] else ["SKU-1","Running shoes","shoes"] if "itemId" in dims else []
        rows = [(dimension_values+["previous"],[v/2 for v in values]),(dimension_values+["current"],values)]
        result = report(dims+["dateRange"],metrics,rows)
        if self.restrict:
            result["metadata"]={"subjectToThresholding":True,"schemaRestrictionResponse":{"activeMetricRestrictions":[{"metricName":"itemRevenue"}]}}
        if self.limited:
            result["rowCount"] = ROW_LIMIT + 1
        return result


class DateTests(unittest.TestCase):
    def test_custom_dates_and_previous_equal_length(self):
        ranges = date_ranges(start_date="2024-03-01",end_date="2024-03-31",today=date(2024,4,2))
        self.assertEqual(ranges[0]["startDate"],"2024-03-01")
        self.assertEqual(ranges[1]["startDate"],"2024-01-30")
        self.assertEqual(ranges[1]["endDate"],"2024-02-29")

    def test_last_7_days_excludes_today(self):
        ranges = date_ranges(days=7,today=date(2026,9,11))
        self.assertEqual((ranges[0]["startDate"],ranges[0]["endDate"]),("2026-09-04","2026-09-10"))
        self.assertEqual((ranges[1]["startDate"],ranges[1]["endDate"]),("2026-08-28","2026-09-03"))

    def test_invalid_ranges_rejected(self):
        for kwargs in [{"days":0},{"days":367},{"start_date":"2026-09-01"},{"start_date":"2026-09-10","end_date":"2026-09-01"},{"start_date":"2026-02-30","end_date":"2026-03-01"},{"start_date":"2026-09-01","end_date":"2099-01-01"}]:
            with self.subTest(kwargs=kwargs), self.assertRaises(HTTPException):
                date_ranges(**kwargs)


class CommerceTests(unittest.TestCase):
    def setUp(self):
        self.provider = FixtureAnalytics()

    def get_report(self, **kwargs):
        return commerce_report(self.provider,"123",start_date="2026-08-01",end_date="2026-08-30",**kwargs)

    def test_current_previous_and_ga4_rates_not_quantity_ratios(self):
        data = self.get_report()
        row=data["rows"][0]
        self.assertEqual(row["current"]["itemRevenue"],400)
        self.assertEqual(row["previous"]["itemRevenue"],200)
        self.assertEqual(row["current"]["cartToViewRate"],.22)  # not 30/100
        self.assertEqual(row["current"]["averageSessionDuration"],165)
        self.assertEqual(data["property_quality"]["current"]["sessions"],20)
        self.assertEqual(data["source"],"ga4")
        for call in self.provider.calls:
            if isinstance(call,dict):
                self.assertEqual(call["dateRanges"][0]["startDate"],"2026-08-01")
                self.assertEqual(call["dateRanges"][1]["endDate"],"2026-07-31")

    def test_product_filter_is_exact_and_never_applied_to_property_quality(self):
        data=self.get_report(view="products",category="shoes")
        self.assertEqual(data["rows"][0]["item_id"],"SKU-1")
        for call in self.provider.calls:
            if isinstance(call,dict):
                names=[m["name"] for m in call["metrics"]]
                if names[0] in ITEM_METRICS or call["dimensions"] == [{"name":"itemId"},{"name":"itemName"},{"name":"itemCategory"}]:
                    self.assertEqual(call["dimensionFilter"]["filter"]["stringFilter"],{"matchType":"EXACT","value":"shoes","caseSensitive":True})
                elif names[0] == "sessions":
                    self.assertNotIn("dimensionFilter",call)

    def test_incompatible_quality_does_not_fabricate_rows(self):
        self.provider.compatible=[]
        data=self.get_report()
        self.assertEqual(data["quality_metrics"],[])
        self.assertIsNone(data["rows"][0]["current"].get("bounceRate"))
        self.assertEqual(data["property_quality"]["current"]["bounceRate"],.4)
        self.assertTrue(data["warnings"])

    def test_quality_failure_preserves_item_report_with_warning(self):
        self.provider.quality_failure=True
        data=self.get_report()
        self.assertEqual(data["rows"][0]["current"]["itemsViewed"],100)
        self.assertFalse(data["quality_metrics"])
        self.assertIsNone(data["property_quality"]["current"].get("bounceRate"))
        self.assertTrue(any("unavailable" in w for w in data["warnings"]))

    def test_empty_report_has_no_sample_or_false_zero_rates(self):
        self.provider.empty=True
        data=self.get_report()
        self.assertEqual(data["rows"],[])
        self.assertIsNone(data["summary"]["current"]["purchaseToViewRate"])
        self.assertTrue(any("No events recorded" in w for w in data["warnings"]))
        self.assertNotIn("demo",data)

    def test_invalid_property_never_reaches_google(self):
        provider=Mock()
        for pid in ["", "properties/123", "../456", "123:runReport", "0"]:
            with self.subTest(pid=pid),self.assertRaises(HTTPException):
                commerce_report(provider,pid)
        provider.property.assert_not_called()

    def test_invalid_view_or_category_filter_rejected(self):
        for kwargs in [{"view":"other"},{"category":"shoes"},{"view":"products","category":"x"*301}]:
            with self.subTest(kwargs=kwargs), self.assertRaises(HTTPException):
                self.get_report(**kwargs)

    def test_private_metrics_null_and_thresholds_disclosed(self):
        self.provider.restrict=True
        data=self.get_report()
        self.assertIsNone(data["rows"][0]["current"]["itemRevenue"])
        self.assertTrue(any("privacy thresholds" in w for w in data["warnings"]))
        self.assertTrue(any("restricts" in w for w in data["warnings"]))

    def test_truncated_reports_are_explicit(self):
        self.provider.limited=True
        data=self.get_report()
        self.assertFalse(data["list_complete"])
        self.assertTrue(any("limited to" in w for w in data["warnings"]))

    def test_parser_uses_headers_and_never_serializes_nonfinite_numbers(self):
        data=report(["dateRange","itemCategory"],["itemRevenue","itemsViewed"],[(["previous","x"],["NaN",12])])
        dims,values=next(parsed_rows(data))
        self.assertEqual(dims["dateRange"],"previous")
        self.assertEqual(values["itemsViewed"],12)
        self.assertIsNone(values["itemRevenue"])


class ProviderTests(unittest.TestCase):
    @patch('functions.ga4_commerce.requests.request')
    def test_paginated_properties_no_demo_fallback(self,request):
        request.side_effect=[Mock(status_code=200,json=lambda:{"accountSummaries":[],"nextPageToken":"next"}),Mock(status_code=200,json=lambda:{"accountSummaries":[{"displayName":"A","propertySummaries":[{"property":"properties/123","displayName":"Store"}]}]})]
        rows=GoogleAnalytics("token").properties()
        self.assertEqual(rows,[{"id":"123","name":"Store","account":"A"}])
        self.assertEqual(request.call_args.kwargs["params"]["pageToken"],"next")
        request.return_value=Mock(status_code=200,json=lambda:{})
        request.side_effect=None
        self.assertEqual(GoogleAnalytics("token").properties(),[])

    @patch('functions.ga4_commerce.requests.request')
    def test_errors_are_safe_and_do_not_return_sample(self,request):
        for status,expected in [(401,401),(403,403),(429,429),(400,422),(500,502)]:
            request.return_value=Mock(status_code=status,text="private upstream error")
            with self.subTest(status=status),self.assertRaises(HTTPException) as raised:
                GoogleAnalytics("secret").properties()
            self.assertEqual(raised.exception.status_code,expected)
            self.assertNotIn("private",str(raised.exception.detail))


if __name__ == '__main__':
    unittest.main()
