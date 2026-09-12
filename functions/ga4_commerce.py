"""GA4 ecommerce reporting. Never substitute synthetic data for provider errors.

Item quantities and GA4 user-based conversion rates are separate from session
quality. Request compatible quality metrics for the actual item dimensions,
and retain property-level session metrics separately (never sum item sessions).
"""
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta, timezone
import math
import re
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import requests
from fastapi import HTTPException

ITEM_METRICS = ["itemsViewed", "itemsAddedToCart", "itemsPurchased", "itemRevenue", "cartToViewRate", "purchaseToViewRate"]
QUALITY_METRICS = ["sessions", "activeUsers", "engagedSessions", "bounceRate", "engagementRate", "averageSessionDuration"]
EVENTS = ["view_item", "add_to_cart", "begin_checkout", "purchase"]
ROW_LIMIT = 10000


def problem(status, code, message):
    return HTTPException(status_code=status, detail={"code": code, "message": message})


def property_id(value):
    if not re.fullmatch(r"[1-9][0-9]{0,19}", str(value)):
        raise problem(422, "invalid_property", "Choose a valid GA4 property.")
    return str(value)


class GoogleAnalytics:
    def __init__(self, token):
        self.headers = {"Authorization": "Bearer " + token}

    def request(self, method, url, **kwargs):
        try:
            response = requests.request(method, url, headers=self.headers, timeout=(5, 25), **kwargs)
        except requests.RequestException:
            raise problem(503, "google_unavailable", "Google Analytics could not be reached. Please try again.")
        if response.status_code == 401:
            raise problem(401, "google_reconnect", "Your Google connection has expired. Reconnect Google Analytics.")
        if response.status_code == 403:
            raise problem(403, "google_permission", "Google denied access. Check this account’s property permissions, Analytics API access and consent, then reconnect.")
        if response.status_code == 429:
            raise problem(429, "google_quota", "Google Analytics’ request quota has been reached. Wait a few minutes and refresh.")
        if response.status_code == 400:
            raise problem(422, "report_incompatible", "Google Analytics does not support this report combination for the selected property.")
        if response.status_code != 200:
            raise problem(502, "google_report_failed", "Google Analytics could not return this report. Please refresh or choose another property.")
        try:
            payload = response.json()
            if not isinstance(payload, dict):
                raise ValueError()
            return payload
        except ValueError:
            raise problem(502, "google_response_invalid", "Google Analytics returned an unreadable response. Please try again.")

    def properties(self):
        result, page_token = [], None
        seen_tokens = set()
        while True:
            params = {"pageSize": 200}
            if page_token:
                params["pageToken"] = page_token
            payload = self.request("GET", "https://analyticsadmin.googleapis.com/v1beta/accountSummaries", params=params)
            for account in payload.get("accountSummaries", []):
                for prop in account.get("propertySummaries", []):
                    pid = prop.get("property", "").removeprefix("properties/")
                    if re.fullmatch(r"[1-9][0-9]{0,19}", pid):
                        result.append({"id": pid, "name": prop.get("displayName", pid), "account": account.get("displayName", "")})
            page_token = payload.get("nextPageToken")
            if not page_token:
                break
            if page_token in seen_tokens or len(seen_tokens) >= 49:
                raise problem(502, "property_list_incomplete", "Your property list could not be fully loaded. Please try again.")
            seen_tokens.add(page_token)
        return list({p["id"]: p for p in result}.values())

    def property(self, pid):
        data = self.request("GET", f"https://analyticsadmin.googleapis.com/v1beta/properties/{property_id(pid)}")
        return {"id": str(pid), "name": data.get("displayName", str(pid)), "currency": data.get("currencyCode", ""), "time_zone": data.get("timeZone", "UTC")}

    def report(self, pid, body):
        return self.request("POST", f"https://analyticsdata.googleapis.com/v1beta/properties/{property_id(pid)}:runReport", json=body)

    def funnel(self, pid, body):
        return self.request("POST", f"https://analyticsdata.googleapis.com/v1alpha/properties/{property_id(pid)}:runFunnelReport", json=body)

    def compatible_quality(self, pid, dimensions, dimension_filter=None):
        body = {"dimensions": [{"name": d} for d in dimensions], "compatibilityFilter": "COMPATIBLE"}
        if dimension_filter:
            body["dimensionFilter"] = dimension_filter
        data = self.request("POST", f"https://analyticsdata.googleapis.com/v1beta/properties/{property_id(pid)}:checkCompatibility", json=body)
        allowed = {m.get("metricMetadata", {}).get("apiName") for m in data.get("metricCompatibilities", []) if m.get("compatibility") == "COMPATIBLE"}
        return [m for m in QUALITY_METRICS if m in allowed]


def date_ranges(days=30, start_date=None, end_date=None, time_zone="UTC", today=None):
    try:
        tz = ZoneInfo(time_zone)
    except (ZoneInfoNotFoundError, ValueError):
        tz = timezone.utc
    today = today or datetime.now(tz).date()
    if bool(start_date) != bool(end_date):
        raise problem(422, "invalid_dates", "Provide both the start date and end date.")
    try:
        if start_date and end_date:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", start_date) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", end_date):
                raise ValueError()
            start, end = date.fromisoformat(start_date), date.fromisoformat(end_date)
        else:
            if not 1 <= days <= 366:
                raise ValueError()
            end = today - timedelta(days=1)
            start = end - timedelta(days=days - 1)
        length = (end - start).days + 1
        if not 1 <= length <= 366 or end > today or start < date(2005, 1, 1):
            raise ValueError()
    except (ValueError, TypeError, OverflowError):
        raise problem(422, "invalid_dates", "Choose valid dates covering 1–366 days, ending today or earlier.")
    return [{"startDate": start.isoformat(), "endDate": end.isoformat(), "name": "current"}, {"startDate": (start - timedelta(days=length)).isoformat(), "endDate": (start - timedelta(days=1)).isoformat(), "name": "previous"}]


def report_body(ranges, dimensions, metrics, dimension_filter=None):
    body = {"dateRanges": ranges, "dimensions": [{"name": d} for d in dimensions], "metrics": [{"name": m} for m in metrics], "limit": ROW_LIMIT, "keepEmptyRows": True}
    if dimension_filter:
        body["dimensionFilter"] = dimension_filter
    return body


def parsed_rows(report):
    """Use response headers; GA4 adds dateRange when multiple ranges are used."""
    dimensions = [h["name"] for h in report.get("dimensionHeaders", [])]
    metrics = [h["name"] for h in report.get("metricHeaders", [])]
    restricted = set()
    for item in report.get("metadata", {}).get("schemaRestrictionResponse", {}).get("activeMetricRestrictions", []):
        restricted.add(item.get("metricName"))
    for row in report.get("rows", []):
        dims = dict(zip(dimensions, (v.get("value", "") for v in row.get("dimensionValues", []))))
        values = {}
        for name, metric in zip(metrics, row.get("metricValues", [])):
            try:
                value = float(metric["value"])
                values[name] = value if math.isfinite(value) and name not in restricted else None
            except (ValueError, TypeError, KeyError):
                values[name] = None
        yield dims, values


def aggregate(report):
    result = {"current": {}, "previous": {}}
    for dims, values in parsed_rows(report):
        period = dims.get("dateRange", "current")
        if period in result:
            result[period] = values
    return result


def metadata_warnings(report, label):
    meta = report.get("metadata", {})
    result = []
    if meta.get("subjectToThresholding"):
        result.append(f"{label}: Google applied privacy thresholds. Some rows may be withheld.")
    if meta.get("dataLossFromOtherRow"):
        result.append(f"{label}: Google grouped high-cardinality data into an ‘(other)’ row.")
    if meta.get("samplingMetadatas"):
        result.append(f"{label}: this report contains sampled data from Google Analytics.")
    if meta.get("schemaRestrictionResponse", {}).get("activeMetricRestrictions"):
        result.append(f"{label}: your GA4 role restricts some metrics. Restricted values are shown as unavailable.")
    return result


def commerce_report(provider, pid, view="categories", category="", days=30, start_date=None, end_date=None):
    property_id(pid)
    if view not in ("categories", "products") or len(category) > 300 or (category and view != "products"):
        raise problem(422, "invalid_breakdown", "Choose categories or products and a valid category filter.")
    # Admin and Data API calls use this user's own access token. Google enforces
    # property permissions on every request, including manually supplied IDs.
    prop = provider.property(pid)
    ranges = date_ranges(days, start_date, end_date, prop["time_zone"])
    dimensions = ["itemCategory"] if view == "categories" else ["itemId", "itemName", "itemCategory"]
    filt = {"filter": {"fieldName": "itemCategory", "stringFilter": {"matchType": "EXACT", "value": category, "caseSensitive": True}}} if category else None
    item_body = report_body(ranges, dimensions, ITEM_METRICS, filt)
    item_body["orderBys"] = [{"metric": {"metricName": "itemRevenue"}, "desc": True}]
    warnings = []

    def optional_report(body, label):
        try:
            return provider.report(pid, body), []
        except HTTPException as error:
            if error.status_code == 401:
                raise
            return {}, [f"{label} is unavailable. {error.detail['message']}"]

    def item_quality():
        try:
            supported = provider.compatible_quality(pid, dimensions, filt)
            if not supported:
                return {}, [], ["Google does not support session-quality metrics with this item breakdown."]
            report, notices = optional_report(report_body(ranges, dimensions, supported, filt), "Item engagement")
            return report, supported if report else [], notices
        except HTTPException as error:
            if error.status_code == 401:
                raise
            return {}, [], [f"Item engagement compatibility could not be checked. {error.detail['message']}"]

    event_filter = {"filter": {"fieldName": "eventName", "inListFilter": {"values": EVENTS}}}
    with ThreadPoolExecutor(max_workers=3) as pool:
        item_future = pool.submit(provider.report, pid, item_body)
        summary_future = pool.submit(optional_report, report_body(ranges, [], ITEM_METRICS, filt), "Item totals")
        quality_future = pool.submit(optional_report, report_body(ranges, [], QUALITY_METRICS), "Property engagement")
        events_future = pool.submit(optional_report, report_body(ranges, ["eventName"], ["eventCount"], event_filter), "Event coverage")
        item_quality_future = pool.submit(item_quality)
        items = item_future.result()
        summary, notices = summary_future.result(); warnings.extend(notices)
        quality, notices = quality_future.result(); warnings.extend(notices)
        events, notices = events_future.result(); warnings.extend(notices)
        row_quality, compatible, notices = item_quality_future.result(); warnings.extend(notices)

    prop["currency"] = items.get("metadata", {}).get("currencyCode") or summary.get("metadata", {}).get("currencyCode") or prop["currency"]
    if not prop["currency"]:
        warnings.append("Google did not provide a report currency. Revenue amounts are shown without a currency symbol.")

    grouped = {}
    def key_for(dims):
        return tuple(dims.get(d, "") for d in dimensions)
    for dims, values in parsed_rows(items):
        key = key_for(dims)
        if key not in grouped:
            grouped[key] = {"key": list(key), "name": dims.get("itemName", "") or dims.get("itemId", "") or dims.get("itemCategory", "") or "(not set)", "category": dims.get("itemCategory", "") or "(not set)", "item_id": dims.get("itemId", ""), "current": {}, "previous": {}}
        period = dims.get("dateRange", "current")
        if period in ("current", "previous"):
            grouped[key][period].update(values)
    for dims, values in parsed_rows(row_quality):
        period, key = dims.get("dateRange", "current"), key_for(dims)
        if key in grouped and period in ("current", "previous"):
            grouped[key][period].update(values)
    rows = list(grouped.values())
    # Never call a missing or invalid denominator a measured 0% conversion rate.
    for row in rows:
        for period in ("current", "previous"):
            values = row[period]
            if not values.get("itemsViewed"):
                values["cartToViewRate"] = values["purchaseToViewRate"] = None
            if not values.get("sessions"):
                for metric in ("bounceRate", "engagementRate", "averageSessionDuration"):
                    values[metric] = None
    totals = aggregate(summary)
    for values in totals.values():
        if not values.get("itemsViewed"):
            values["cartToViewRate"] = values["purchaseToViewRate"] = None
    property_quality = aggregate(quality)
    for values in property_quality.values():
        if not values.get("sessions"):
            for metric in ("bounceRate", "engagementRate", "averageSessionDuration"):
                values[metric] = None
    event_counts = {"current": {}, "previous": {}}
    if events:
        event_counts = {period: {e: 0 for e in EVENTS} for period in ("current", "previous")}
        for dims, values in parsed_rows(events):
            period, name = dims.get("dateRange", "current"), dims.get("eventName")
            if period in event_counts and name in EVENTS:
                event_counts[period][name] = values.get("eventCount")
        missing = [e for e in EVENTS if not event_counts["current"].get(e)]
        if missing:
            warnings.append("No events recorded for " + ", ".join(missing) + ". Check your ecommerce instrumentation or use a longer date range.")
    complete = int(items.get("rowCount", len(items.get("rows", [])))) <= len(items.get("rows", []))
    if not complete:
        warnings.append(f"The item list is limited to {ROW_LIMIT:,} Google report rows across both periods. Totals cover the full scope; the table and export contain only loaded rows.")
    if row_quality and int(row_quality.get("rowCount", len(row_quality.get("rows", [])))) > len(row_quality.get("rows", [])):
        warnings.append("The engagement report reached its row limit. Some item-level quality values are unavailable.")
    if any(r["category"] in ("(not set)", "(unset)") for r in rows):
        warnings.append("Some items have no category. Populate item_category in the items array of your ecommerce events.")
    for report, label in [(items, "Item report"), (summary, "Item totals"), (quality, "Property engagement"), (row_quality, "Item engagement"), (events, "Events")]:
        warnings.extend(metadata_warnings(report, label))
    return {"source": "ga4", "view": view, "property": prop, "rows": rows, "summary": totals, "property_quality": property_quality, "quality_metrics": compatible, "events": event_counts, "warnings": warnings, "list_complete": complete, "start_date": ranges[0]["startDate"], "end_date": ranges[0]["endDate"], "previous_start": ranges[1]["startDate"], "previous_end": ranges[1]["endDate"], "fetched_at": datetime.now(timezone.utc).isoformat()}
