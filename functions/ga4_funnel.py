"""GA4 ordered funnel, next-action path and aggregate user exploration reports."""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
import math
import re

from fastapi import HTTPException

from functions.ga4_commerce import date_ranges, metadata_warnings, parsed_rows, problem, property_id, report_body


PRESETS = {
    "ecommerce": [
        ("Session started", "session_start"),
        ("Product viewed", "view_item"),
        ("Added to cart", "add_to_cart"),
        ("Checkout started", "begin_checkout"),
        ("Purchase", "purchase"),
    ],
    "checkout": [
        ("Cart viewed", "view_cart"),
        ("Checkout started", "begin_checkout"),
        ("Shipping added", "add_shipping_info"),
        ("Payment added", "add_payment_info"),
        ("Purchase", "purchase"),
    ],
    "lead": [
        ("Session started", "session_start"),
        ("Page viewed", "page_view"),
        ("Form started", "form_start"),
        ("Lead generated", "generate_lead"),
    ],
}
BREAKDOWNS = {
    "deviceCategory": "Device category",
    "sessionDefaultChannelGroup": "Channel group",
    "country": "Country",
    "platform": "Platform",
}
USER_METRICS = ["activeUsers", "sessions", "engagedSessions", "engagementRate", "bounceRate", "averageSessionDuration", "eventsPerSession", "ecommercePurchases", "totalRevenue"]
LANDING_METRICS = ["activeUsers", "sessions", "engagementRate", "bounceRate", "averageSessionDuration", "keyEvents"]


def _number(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def _subreport_rows(report):
    dimensions = [item.get("name", "") for item in report.get("dimensionHeaders", [])]
    metrics = [item.get("name", "") for item in report.get("metricHeaders", [])]
    for row in report.get("rows", []):
        dims = dict(zip(dimensions, (item.get("value", "") for item in row.get("dimensionValues", []))))
        vals = dict(zip(metrics, (_number(item.get("value")) for item in row.get("metricValues", []))))
        yield dims, vals


def _step_number(value):
    match = re.match(r"\s*(\d+)\.", str(value or ""))
    return int(match.group(1)) if match else 0


def _funnel_body(date_range, preset, breakdown, is_open=False, next_actions=False):
    steps = PRESETS.get(preset)
    if not steps:
        raise problem(422, "invalid_funnel_preset", "Choose a supported funnel preset.")
    if breakdown not in BREAKDOWNS:
        raise problem(422, "invalid_funnel_breakdown", "Choose a supported funnel breakdown.")
    body = {
        "dateRanges": [date_range],
        "funnel": {
            "isOpenFunnel": bool(is_open),
            "steps": [
                {"name": label, "filterExpression": {"funnelEventFilter": {"eventName": event}}}
                for label, event in steps
            ],
        },
        "funnelBreakdown": {"breakdownDimension": {"name": breakdown}, "limit": "10"},
        "limit": "10000",
    }
    if next_actions:
        body["funnelNextAction"] = {"nextActionDimension": {"name": "eventName"}, "limit": "5"}
    return body


def _parse_funnel(report, preset, breakdown):
    table = report.get("funnelTable", {})
    totals, groups = {}, {}
    for dims, metrics in _subreport_rows(table):
        index = _step_number(dims.get("funnelStepName"))
        if not index:
            continue
        group = dims.get(breakdown, "RESERVED_TOTAL")
        item = {
            "step": index,
            "users": int(metrics.get("activeUsers") or 0),
            "completion_rate": metrics.get("funnelStepCompletionRate"),
            "abandonments": int(metrics.get("funnelStepAbandonments") or 0),
            "abandonment_rate": metrics.get("funnelStepAbandonmentRate"),
        }
        if group == "RESERVED_TOTAL":
            totals[index] = item
        else:
            groups.setdefault(group or "(not set)", {})[index] = item

    steps = []
    for index, (label, event) in enumerate(PRESETS[preset], 1):
        item = totals.get(index, {"users": 0, "completion_rate": None, "abandonments": 0, "abandonment_rate": None})
        steps.append({"step": index, "name": label, "event": event, **item})

    breakdown_rows = []
    for name, values in groups.items():
        first, last = values.get(1, {}), values.get(len(steps), {})
        first_users, last_users = int(first.get("users") or 0), int(last.get("users") or 0)
        breakdown_rows.append({
            "name": name,
            "entry_users": first_users,
            "completion_users": last_users,
            "conversion_rate": last_users / first_users if first_users else None,
            "largest_drop_step": max(values, key=lambda key: values[key].get("abandonments") or 0) if values else None,
        })
    breakdown_rows.sort(key=lambda item: item["entry_users"], reverse=True)

    paths = {}
    for dims, metrics in _subreport_rows(report.get("funnelVisualization", {})):
        index = _step_number(dims.get("funnelStepName"))
        action = dims.get("funnelStepNextAction", "")
        if index and action and action != "RESERVED_TOTAL":
            paths.setdefault(index, []).append({"name": action, "users": int(metrics.get("activeUsers") or 0)})
    for index, rows in paths.items():
        total = sum(row["users"] for row in rows)
        rows.sort(key=lambda row: row["users"], reverse=True)
        for row in rows:
            row["share"] = row["users"] / total if total else None

    return steps, breakdown_rows, paths


def _segments(report):
    rows = []
    for dims, values in parsed_rows(report):
        if dims.get("dateRange") != "current":
            continue
        rows.append({"name": dims.get("newVsReturning") or "(not set)", **values})
    rows.sort(key=lambda item: item.get("activeUsers") or 0, reverse=True)
    return rows


def _landing_pages(report):
    rows = []
    for dims, values in parsed_rows(report):
        path = dims.get("landingPagePlusQueryString") or "(not set)"
        rows.append({"path": path, **values})
    rows.sort(key=lambda item: item.get("sessions") or 0, reverse=True)
    return rows[:12]


def _comparison(current, previous):
    first_current = current[0]["users"] if current else 0
    last_current = current[-1]["users"] if current else 0
    first_previous = previous[0]["users"] if previous else 0
    last_previous = previous[-1]["users"] if previous else 0
    current_rate = last_current / first_current if first_current else None
    previous_rate = last_previous / first_previous if first_previous else None
    return {
        "entry_users": first_current,
        "completion_users": last_current,
        "conversion_rate": current_rate,
        "previous_conversion_rate": previous_rate,
        "conversion_change_pp": (current_rate - previous_rate) * 100 if current_rate is not None and previous_rate is not None else None,
        "total_abandonments": sum(step["abandonments"] for step in current[:-1]),
        "largest_drop_step": max(current[:-1], key=lambda item: item["abandonments"])["step"] if len(current) > 1 else None,
    }


def funnel_report(provider, pid, preset="ecommerce", breakdown="deviceCategory", is_open=False, days=30, start_date=None, end_date=None):
    pid = property_id(pid)
    if preset not in PRESETS or breakdown not in BREAKDOWNS:
        raise problem(422, "invalid_funnel_configuration", "Choose a supported funnel preset and breakdown.")
    prop = provider.property(pid)
    ranges = date_ranges(days, start_date, end_date, prop.get("time_zone") or "UTC")
    current_range, previous_range = ranges
    current_range = {key: value for key, value in current_range.items() if key != "name"}
    previous_range = {key: value for key, value in previous_range.items() if key != "name"}
    user_body = report_body(ranges, ["newVsReturning"], USER_METRICS)
    user_body["orderBys"] = [{"metric": {"metricName": "activeUsers"}, "desc": True}]
    landing_body = report_body([ranges[0]], ["landingPagePlusQueryString"], LANDING_METRICS)
    landing_body["limit"] = 50
    landing_body["orderBys"] = [{"metric": {"metricName": "sessions"}, "desc": True}]
    warnings = []

    def optional_core(body, label):
        try:
            return provider.report(pid, body), None
        except HTTPException as error:
            if error.status_code == 401:
                raise
            detail = error.detail if isinstance(error.detail, dict) else {}
            message = detail.get("message") or "Google Analytics could not return this report."
            return {}, f"{label} is unavailable. {message}"

    with ThreadPoolExecutor(max_workers=4) as pool:
        current_future = pool.submit(provider.funnel, pid, _funnel_body(current_range, preset, breakdown, is_open, True))
        previous_future = pool.submit(provider.funnel, pid, _funnel_body(previous_range, preset, breakdown, is_open, False))
        users_future = pool.submit(optional_core, user_body, "User exploration")
        landing_future = pool.submit(optional_core, landing_body, "Landing-page exploration")
        current_report = current_future.result()
        previous_report = previous_future.result()
        users_report, users_warning = users_future.result()
        landing_report, landing_warning = landing_future.result()

    warnings.extend(item for item in (users_warning, landing_warning) if item)

    current_steps, breakdown_rows, paths = _parse_funnel(current_report, preset, breakdown)
    previous_steps, _, _ = _parse_funnel(previous_report, preset, breakdown)
    for report, label in ((current_report.get("funnelTable", {}), "Current funnel"),
                          (current_report.get("funnelVisualization", {}), "Next actions"),
                          (previous_report.get("funnelTable", {}), "Previous funnel"),
                          (users_report, "User exploration"), (landing_report, "Landing pages")):
        warnings.extend(metadata_warnings(report, label))
    if not any(step["users"] for step in current_steps):
        warnings.append("No users completed the selected event sequence. Check event collection, try an open funnel or use a longer date range.")
    return {
        "source": "ga4",
        "property": prop,
        "preset": preset,
        "is_open": bool(is_open),
        "breakdown": {"dimension": breakdown, "label": BREAKDOWNS[breakdown]},
        "start_date": ranges[0]["startDate"],
        "end_date": ranges[0]["endDate"],
        "previous_start": ranges[1]["startDate"],
        "previous_end": ranges[1]["endDate"],
        "summary": _comparison(current_steps, previous_steps),
        "steps": current_steps,
        "breakdown_rows": breakdown_rows,
        "next_actions": {str(index): rows for index, rows in paths.items()},
        "user_segments": _segments(users_report),
        "landing_pages": _landing_pages(landing_report),
        "warnings": warnings,
        "fetched_at": datetime.now(timezone.utc).isoformat(),
    }


def sample_report(preset="ecommerce", breakdown="deviceCategory", is_open=False, days=30, start_date=None, end_date=None):
    if preset not in PRESETS or breakdown not in BREAKDOWNS:
        raise problem(422, "invalid_funnel_configuration", "Choose a supported funnel preset and breakdown.")
    ranges = date_ranges(days, start_date, end_date, "UTC")
    base = [84200, 51400, 18760, 9420, 4180] if preset != "lead" else [84200, 73100, 11400, 3650]
    scale = max(0.08, (datetime.fromisoformat(ranges[0]["endDate"]).date() - datetime.fromisoformat(ranges[0]["startDate"]).date()).days / 29)
    counts = [int(value * scale) for value in base]
    steps = []
    for index, ((label, event), users) in enumerate(zip(PRESETS[preset], counts), 1):
        next_users = counts[index] if index < len(counts) else users
        abandonments = max(0, users - next_users) if index < len(counts) else 0
        steps.append({"step": index, "name": label, "event": event, "users": users,
                      "completion_rate": next_users / users if users and index < len(counts) else None,
                      "abandonments": abandonments,
                      "abandonment_rate": abandonments / users if users else None})
    breakdown_names = {"deviceCategory": ["mobile", "desktop", "tablet"], "sessionDefaultChannelGroup": ["Organic Search", "Paid Search", "Direct", "Email"], "country": ["Türkiye", "Germany", "United Kingdom"], "platform": ["web", "Android", "iOS"]}[breakdown]
    shares = [.61, .31, .08, .05]
    breakdown_rows = []
    for name, share in zip(breakdown_names, shares):
        entry, completed = int(counts[0] * share), int(counts[-1] * share * (1.08 if name in ("desktop", "Email") else .96))
        breakdown_rows.append({"name": name, "entry_users": entry, "completion_users": completed, "conversion_rate": completed / entry if entry else None, "largest_drop_step": 2})
    next_actions = {}
    options = [[("page_view", 24600), ("scroll", 13900), ("view_search_results", 6700)], [("select_item", 9800), ("view_item_list", 7100), ("scroll", 5900)], [("view_cart", 6100), ("view_promotion", 3300), ("remove_from_cart", 2100)], [("add_shipping_info", 3800), ("add_payment_info", 2600), ("view_cart", 1400)]]
    for index, rows in enumerate(options[:len(steps)-1], 1):
        total = sum(value for _, value in rows)
        next_actions[str(index)] = [{"name": name, "users": int(value * scale), "share": value / total} for name, value in rows]
    return {
        "source": "sample", "property": {"id": "sample", "name": "Illustrative GA4 store", "currency": "USD", "time_zone": "UTC"},
        "preset": preset, "is_open": bool(is_open), "breakdown": {"dimension": breakdown, "label": BREAKDOWNS[breakdown]},
        "start_date": ranges[0]["startDate"], "end_date": ranges[0]["endDate"], "previous_start": ranges[1]["startDate"], "previous_end": ranges[1]["endDate"],
        "summary": {**_comparison(steps, [{**step, "users": int(step["users"] * 1.08)} for step in steps]), "conversion_change_pp": .34},
        "steps": steps, "breakdown_rows": breakdown_rows, "next_actions": next_actions,
        "user_segments": [
            {"name": "new", "activeUsers": int(53600*scale), "sessions": int(58200*scale), "engagedSessions": int(36100*scale), "engagementRate": .620, "bounceRate": .380, "averageSessionDuration": 96.4, "eventsPerSession": 5.8, "ecommercePurchases": int(1870*scale), "totalRevenue": 274000*scale},
            {"name": "returning", "activeUsers": int(30600*scale), "sessions": int(41900*scale), "engagedSessions": int(31200*scale), "engagementRate": .745, "bounceRate": .255, "averageSessionDuration": 174.8, "eventsPerSession": 8.7, "ecommercePurchases": int(2310*scale), "totalRevenue": 438000*scale},
        ],
        "landing_pages": [
            {"path": "/", "activeUsers": int(18900*scale), "sessions": int(22600*scale), "engagementRate": .71, "bounceRate": .29, "averageSessionDuration": 142.0, "keyEvents": int(1410*scale)},
            {"path": "/collections/new", "activeUsers": int(12400*scale), "sessions": int(14800*scale), "engagementRate": .66, "bounceRate": .34, "averageSessionDuration": 119.0, "keyEvents": int(820*scale)},
            {"path": "/products/signature", "activeUsers": int(9800*scale), "sessions": int(11200*scale), "engagementRate": .78, "bounceRate": .22, "averageSessionDuration": 206.0, "keyEvents": int(960*scale)},
            {"path": "/sale", "activeUsers": int(7200*scale), "sessions": int(8900*scale), "engagementRate": .59, "bounceRate": .41, "averageSessionDuration": 88.0, "keyEvents": int(430*scale)},
        ],
        "warnings": ["Sample data is illustrative and does not belong to your GA4 property."], "fetched_at": datetime.now(timezone.utc).isoformat(),
    }
