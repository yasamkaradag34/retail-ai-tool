"""Google Merchant Center catalog, performance and pricing insights.

The Merchant Reports API exposes current catalog/price snapshots and dated
performance metrics. Provider errors stay errors: sample data is generated only
when the caller explicitly asks to explore the sample workspace.
"""
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, timedelta, timezone, datetime
import math
import re

import requests
from fastapi import HTTPException


ROW_LIMIT = 10000
AT_MARKET_PERCENT = 3.0


def problem(status, code, message):
    return HTTPException(status_code=status, detail={"code": code, "message": message})


def account_id(value):
    value = str(value or "")
    if not re.fullmatch(r"[1-9][0-9]{0,19}", value):
        raise problem(422, "invalid_merchant_account", "Choose a valid Merchant Center account.")
    return value


def date_ranges(days=30, start_date=None, end_date=None, today=None):
    today = today or datetime.now(timezone.utc).date()
    if bool(start_date) != bool(end_date):
        raise problem(422, "invalid_dates", "Provide both the start date and end date.")
    try:
        if start_date and end_date:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", start_date) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", end_date):
                raise ValueError()
            start, end = date.fromisoformat(start_date), date.fromisoformat(end_date)
        else:
            days = int(days)
            if not 1 <= days <= 366:
                raise ValueError()
            end = today - timedelta(days=1)
            start = end - timedelta(days=days - 1)
        length = (end - start).days + 1
        if not 1 <= length <= 366 or end > today or start < date(2005, 1, 1):
            raise ValueError()
    except (ValueError, TypeError, OverflowError):
        raise problem(422, "invalid_dates", "Choose valid dates covering 1–366 days, ending today or earlier.")
    return {
        "start": start.isoformat(), "end": end.isoformat(),
        "previous_start": (start - timedelta(days=length)).isoformat(),
        "previous_end": (start - timedelta(days=1)).isoformat(),
    }


def _number(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


def _integer(value):
    number = _number(value)
    return int(number) if number is not None else None


def _money(value):
    if not isinstance(value, dict):
        return {"amount": None, "currency": ""}
    micros = _number(value.get("amountMicros"))
    return {"amount": micros / 1_000_000 if micros is not None else None,
            "currency": str(value.get("currencyCode") or "")}


def _availability(value):
    normalized = re.sub(r"[^a-z]+", "_", str(value or "").lower()).strip("_")
    return normalized or "unknown"


def _google_error_markers(response):
    """Return normalized provider error markers without exposing its payload."""
    try:
        payload = response.json()
    except (ValueError, TypeError, AttributeError):
        return ""
    if not isinstance(payload, dict):
        return ""
    error = payload.get("error", payload)
    if not isinstance(error, dict):
        return ""
    values = [error.get("status"), error.get("message")]
    for detail in error.get("details", []):
        if isinstance(detail, dict):
            values.extend((detail.get("reason"), detail.get("metadata", {}).get("reason") if isinstance(detail.get("metadata"), dict) else None))
    return " ".join(str(value) for value in values if value).upper()


class GoogleMerchant:
    def __init__(self, token):
        self.headers = {"Authorization": "Bearer " + token, "Content-Type": "application/json"}

    def request(self, method, url, **kwargs):
        try:
            response = requests.request(method, url, headers=self.headers, timeout=(5, 30), **kwargs)
        except requests.RequestException:
            raise problem(503, "merchant_unavailable", "Google Merchant Center could not be reached. Please try again.")
        markers = _google_error_markers(response)
        if response.status_code == 401:
            raise problem(401, "google_reconnect", "Google did not accept this authorization. Reconnect Merchant Center and choose a Google account that has Merchant Center access.")
        if response.status_code == 403:
            if "ACCESS_TOKEN_SCOPE_INSUFFICIENT" in markers or "INSUFFICIENT AUTHENTICATION SCOPES" in markers:
                raise problem(403, "merchant_scope_required", "Your current Google connection does not include Merchant Center permission. Reconnect Merchant Center and approve the requested access.")
            if "GCP_NOT_REGISTERED" in markers or "REGISTER_GCP" in markers:
                raise problem(403, "merchant_developer_registration", "DataProvido's Merchant API developer registration is incomplete. Register the Google Cloud project with the primary Merchant Center account, then retry.")
            if "SERVICE_DISABLED" in markers or "API HAS NOT BEEN USED" in markers:
                raise problem(403, "merchant_api_disabled", "Merchant API is not enabled for DataProvido's Google Cloud project. Enable it in Google Cloud, then retry.")
            if "REPORTING_MANAGER" in markers:
                raise problem(403, "merchant_reporting_permission", "This Google account needs Performance reporting access in Merchant Center.")
            raise problem(403, "merchant_permission", "Google denied Merchant Center access. Reconnect with Merchant Center permission and check your account role.")
        if response.status_code == 429:
            raise problem(429, "merchant_quota", "The Merchant API request quota has been reached. Wait a few minutes and refresh.")
        if response.status_code == 400:
            raise problem(422, "merchant_report_incompatible", "This Merchant Center account does not support the requested report or report fields.")
        if response.status_code != 200:
            raise problem(502, "merchant_report_failed", "Google Merchant Center could not return this report. Please refresh or choose another account.")
        try:
            payload = response.json()
            if not isinstance(payload, dict):
                raise ValueError()
            return payload
        except ValueError:
            raise problem(502, "merchant_response_invalid", "Google Merchant Center returned an unreadable response. Please try again.")

    def accounts(self):
        rows, page_token, seen = [], None, set()
        while True:
            params = {"pageSize": 500}
            if page_token:
                params["pageToken"] = page_token
            data = self.request("GET", "https://merchantapi.googleapis.com/accounts/v1/accounts", params=params)
            for item in data.get("accounts", []):
                aid = str(item.get("name", "")).removeprefix("accounts/")
                if re.fullmatch(r"[1-9][0-9]{0,19}", aid):
                    rows.append({"id": aid, "name": item.get("accountName") or item.get("displayName") or f"Merchant Center {aid}"})
            page_token = data.get("nextPageToken")
            if not page_token:
                break
            if page_token in seen or len(seen) >= 49:
                raise problem(502, "merchant_accounts_incomplete", "Your Merchant Center account list could not be fully loaded.")
            seen.add(page_token)
        return list({row["id"]: row for row in rows}.values())

    def account(self, value):
        aid = account_id(value)
        for item in self.accounts():
            if item["id"] == aid:
                return item
        raise problem(403, "merchant_account_denied", "This Google account cannot access the selected Merchant Center account.")

    def search(self, value, query, limit=ROW_LIMIT):
        aid = account_id(value)
        rows, page_token, seen = [], None, set()
        while len(rows) < limit:
            body = {"query": query, "pageSize": min(5000, limit - len(rows))}
            if page_token:
                body["pageToken"] = page_token
            data = self.request("POST", f"https://merchantapi.googleapis.com/reports/v1/accounts/{aid}/reports:search", json=body)
            page = data.get("results", [])
            if not isinstance(page, list):
                raise problem(502, "merchant_response_invalid", "Google Merchant Center returned an unreadable report.")
            rows.extend(page)
            page_token = data.get("nextPageToken")
            if not page_token:
                return rows, True
            if page_token in seen or len(seen) >= 49:
                break
            seen.add(page_token)
        return rows[:limit], not bool(page_token)


PRODUCT_QUERY = """SELECT offer_id, id, title, brand, category_l1, product_type_l1,
availability, price, aggregated_reporting_context_status, click_potential
FROM product_view"""

PRICE_QUERY = """SELECT offer_id, id, title, brand, category_l1, product_type_l1,
price, report_country_code, benchmark_price
FROM price_competitiveness_product_view"""

PRICE_INSIGHTS_QUERY = """SELECT offer_id, id, title, brand, category_l1, product_type_l1,
price, suggested_price, effectiveness, predicted_impressions_change_fraction,
predicted_clicks_change_fraction, predicted_conversions_change_fraction
FROM price_insights_product_view"""


def performance_query(start, end):
    return f"""SELECT offer_id, title, brand, category_l1, product_type_l1, marketing_method,
clicks, impressions, click_through_rate, conversions, conversion_rate, conversion_value
FROM product_performance_view
WHERE date BETWEEN '{start}' AND '{end}'"""


def _product_rows(raw):
    result = {}
    for row in raw:
        item = row.get("productView", {})
        offer = str(item.get("offerId") or "")
        if not offer:
            continue
        result[offer] = {
            "id": item.get("id") or "", "offer_id": offer,
            "title": item.get("title") or offer, "brand": item.get("brand") or "Unbranded",
            "category": item.get("categoryL1") or item.get("productTypeL1") or "Uncategorised",
            "availability": _availability(item.get("availability")),
            "listing_status": str(item.get("aggregatedReportingContextStatus") or "UNKNOWN"),
            "click_potential": str(item.get("clickPotential") or "UNKNOWN"),
            "price": _money(item.get("price")),
        }
    return result


def _performance_rows(raw):
    result = {}
    for row in raw:
        item = row.get("productPerformanceView", {})
        offer = str(item.get("offerId") or "")
        if not offer:
            continue
        values = result.setdefault(offer, {"clicks": 0, "impressions": 0, "organic_clicks": 0, "conversions": 0.0,
                                           "conversion_values": {}, "title": item.get("title") or offer,
                                           "brand": item.get("brand") or "Unbranded",
                                           "category": item.get("categoryL1") or item.get("productTypeL1") or "Uncategorised"})
        for key, api_key in (("clicks", "clicks"), ("impressions", "impressions"), ("conversions", "conversions")):
            number = _number(item.get(api_key))
            if number is not None:
                values[key] += number
        if item.get("marketingMethod") == "ORGANIC":
            values["organic_clicks"] += _number(item.get("clicks")) or 0
        conversion_value = _money(item.get("conversionValue"))
        if conversion_value["amount"] is not None and conversion_value["currency"]:
            currency = conversion_value["currency"]
            values["conversion_values"][currency] = values["conversion_values"].get(currency, 0) + conversion_value["amount"]
    for values in result.values():
        values["click_through_rate"] = values["clicks"] / values["impressions"] if values["impressions"] else None
        values["conversion_rate"] = values["conversions"] / values["organic_clicks"] if values["organic_clicks"] else None
        if len(values["conversion_values"]) == 1:
            values["conversion_currency"], values["conversion_value"] = next(iter(values["conversion_values"].items()))
        else:
            values["conversion_currency"], values["conversion_value"] = "", None
    return result


def _price_rows(raw):
    result = []
    for row in raw:
        item = row.get("priceCompetitivenessProductView", {})
        offer = str(item.get("offerId") or "")
        current, benchmark = _money(item.get("price")), _money(item.get("benchmarkPrice"))
        gap = None
        if current["amount"] is not None and benchmark["amount"] and current["currency"] == benchmark["currency"]:
            gap = (current["amount"] - benchmark["amount"]) / benchmark["amount"] * 100
        status = "unavailable" if gap is None else "below" if gap < -AT_MARKET_PERCENT else "above" if gap > AT_MARKET_PERCENT else "at_market"
        if offer:
            result.append({"id": item.get("id") or "", "offer_id": offer,
                           "title": item.get("title") or offer, "brand": item.get("brand") or "Unbranded",
                           "category": item.get("categoryL1") or item.get("productTypeL1") or "Uncategorised",
                           "country": item.get("reportCountryCode") or "", "price": current,
                           "benchmark_price": benchmark, "price_gap_percent": gap, "price_status": status})
    return result


def _suggestions(raw):
    result = {}
    for row in raw:
        item = row.get("priceInsightsProductView", {})
        offer = str(item.get("offerId") or "")
        if offer:
            result[offer] = {"suggested_price": _money(item.get("suggestedPrice")),
                             "effectiveness": str(item.get("effectiveness") or "UNKNOWN"),
                             "predicted_impressions_change": _number(item.get("predictedImpressionsChangeFraction")),
                             "predicted_clicks_change": _number(item.get("predictedClicksChangeFraction")),
                             "predicted_conversions_change": _number(item.get("predictedConversionsChangeFraction"))}
    return result


def _sum(rows, key):
    return sum((_number(row.get(key)) or 0) for row in rows.values())


def _performance_summary(rows):
    currencies = {row.get("conversion_currency") for row in rows.values() if numeric_or_none(row.get("conversion_value")) is not None and row.get("conversion_currency")}
    return {"clicks": _sum(rows, "clicks"), "impressions": _sum(rows, "impressions"),
            "conversions": _sum(rows, "conversions"),
            "conversion_value": _sum(rows, "conversion_value") if len(currencies) == 1 else None,
            "conversion_currency": next(iter(currencies)) if len(currencies) == 1 else ""}


def numeric_or_none(value):
    return _number(value)


def merchant_report(provider, merchant_account_id, days=30, start_date=None, end_date=None):
    account = provider.account(merchant_account_id)
    ranges = date_ranges(days, start_date, end_date)
    queries = {
        "products": PRODUCT_QUERY,
        "prices": PRICE_QUERY,
        "suggestions": PRICE_INSIGHTS_QUERY,
        "performance": performance_query(ranges["start"], ranges["end"]),
        "previous": performance_query(ranges["previous_start"], ranges["previous_end"]),
    }
    fetched, completeness, warnings = {}, {}, []
    required = {"products"}
    with ThreadPoolExecutor(max_workers=5) as pool:
        futures = {pool.submit(provider.search, merchant_account_id, query): name for name, query in queries.items()}
        for future in as_completed(futures):
            name = futures[future]
            try:
                fetched[name], completeness[name] = future.result()
            except HTTPException:
                if name in required:
                    raise
                fetched[name], completeness[name] = [], True
                warnings.append({
                    "prices": "Price competitiveness is unavailable. Market Insights eligibility, sufficient data and account access are required.",
                    "suggestions": "Google did not return price suggestions for this account.",
                    "performance": "Current-period Shopping performance is unavailable for this account or date range.",
                    "previous": "Previous-period Shopping performance is unavailable, so changes cannot be calculated.",
                }[name])

    products = _product_rows(fetched["products"])
    current = _performance_rows(fetched["performance"])
    previous = _performance_rows(fetched["previous"])
    suggestions = _suggestions(fetched["suggestions"])
    prices = _price_rows(fetched["prices"])

    price_by_offer = {}
    for item in prices:
        price_by_offer.setdefault(item["offer_id"], item)
    offers = sorted(set(products) | set(current) | set(price_by_offer))
    rows = []
    for offer in offers:
        catalog = products.get(offer, {})
        price = price_by_offer.get(offer, {})
        perf = current.get(offer, {})
        prior = previous.get(offer, {})
        base = catalog or price or perf
        rows.append({
            "id": base.get("id", ""), "offer_id": offer,
            "title": base.get("title") or price.get("title") or perf.get("title") or offer,
            "brand": base.get("brand") or price.get("brand") or perf.get("brand") or "Unbranded",
            "category": base.get("category") or price.get("category") or perf.get("category") or "Uncategorised",
            "availability": catalog.get("availability", "unknown"),
            "listing_status": catalog.get("listing_status", "UNKNOWN"),
            "click_potential": catalog.get("click_potential", "UNKNOWN"),
            "price": price.get("price", catalog.get("price", {"amount": None, "currency": ""})),
            "benchmark_price": price.get("benchmark_price", {"amount": None, "currency": ""}),
            "price_gap_percent": price.get("price_gap_percent"), "price_status": price.get("price_status", "unavailable"),
            "benchmark_country": price.get("country", ""),
            "performance": {key: perf.get(key) for key in ("clicks", "impressions", "organic_clicks", "click_through_rate", "conversions", "conversion_rate", "conversion_value", "conversion_currency")},
            "previous": {key: prior.get(key) for key in ("clicks", "impressions", "organic_clicks", "click_through_rate", "conversions", "conversion_rate", "conversion_value", "conversion_currency")},
            **suggestions.get(offer, {"suggested_price": {"amount": None, "currency": ""}, "effectiveness": "UNKNOWN",
                                      "predicted_impressions_change": None, "predicted_clicks_change": None, "predicted_conversions_change": None}),
        })

    status_counts = {key: 0 for key in ("below", "at_market", "above", "unavailable")}
    gaps = {key: [] for key in status_counts}
    for row in rows:
        status_counts[row["price_status"]] += 1
        if row["price_gap_percent"] is not None:
            gaps[row["price_status"]].append(row["price_gap_percent"])
    availability = {}
    for row in products.values():
        availability[row["availability"]] = availability.get(row["availability"], 0) + 1

    brands = {}
    for row in rows:
        group = brands.setdefault(row["brand"], {"brand": row["brand"], "products": 0, "benchmarked": 0,
                                                  "below": 0, "at_market": 0, "above": 0, "clicks": 0, "impressions": 0})
        group["products"] += 1
        group["clicks"] += row["performance"].get("clicks") or 0
        group["impressions"] += row["performance"].get("impressions") or 0
        if row["price_status"] != "unavailable":
            group["benchmarked"] += 1
            group[row["price_status"]] += 1

    oos_risk = sorted([row for row in rows if row["availability"] == "out_of_stock"], key=lambda row: row["performance"].get("clicks") or 0, reverse=True)
    overpriced = sorted([row for row in rows if row["price_status"] == "above"], key=lambda row: ((row["performance"].get("clicks") or 0), row["price_gap_percent"] or 0), reverse=True)
    margin = sorted([row for row in rows if row["price_status"] == "below"], key=lambda row: ((row["performance"].get("clicks") or 0), -(row["price_gap_percent"] or 0)), reverse=True)
    suggestion_rows = sorted([row for row in rows if row["suggested_price"]["amount"] is not None], key=lambda row: ({"HIGH": 3, "MEDIUM": 2, "LOW": 1}.get(row["effectiveness"], 0), row["performance"].get("clicks") or 0), reverse=True)

    if any(not complete for complete in completeness.values()):
        warnings.append(f"At least one Merchant report exceeded {ROW_LIMIT:,} rows. Tables and exports contain the loaded rows only.")
    warnings.append("Catalog availability and benchmark prices are current snapshots; the selected dates apply to Shopping performance metrics.")
    warnings.append("Merchant Center reports availability states, not physical stock quantity or days of cover. Connect ERP inventory data for low-stock and replenishment analysis.")
    if not prices:
        warnings.append("No price benchmark rows were returned. Google only provides this report for eligible accounts and products with sufficient market data.")
    if any(len(row.get("conversion_values", {})) > 1 for row in list(current.values()) + list(previous.values())):
        warnings.append("Conversion value was returned in multiple currencies for at least one product and is not combined across currencies.")

    return {
        "source": "merchant_api", "account": account, "fetched_at": datetime.now(timezone.utc).isoformat(),
        "start_date": ranges["start"], "end_date": ranges["end"],
        "previous_start": ranges["previous_start"], "previous_end": ranges["previous_end"],
        "summary": {
            "catalog_products": len(products), "in_stock": availability.get("in_stock", 0),
            "out_of_stock": availability.get("out_of_stock", 0),
            "preorder": availability.get("preorder", 0), "backorder": availability.get("backorder", 0),
            "benchmark_products": len(prices), "benchmark_coverage_rate": len(prices) / len(products) if products else None,
            "price_positions": {key: {"count": count, "average_gap_percent": sum(gaps[key]) / len(gaps[key]) if gaps[key] else None} for key, count in status_counts.items()},
            "performance": _performance_summary(current),
            "previous_performance": _performance_summary(previous),
        },
        "rows": rows, "brands": sorted(brands.values(), key=lambda row: row["clicks"], reverse=True),
        "opportunities": {"out_of_stock_demand": oos_risk[:20], "above_benchmark": overpriced[:20],
                          "below_benchmark": margin[:20], "price_suggestions": suggestion_rows[:20]},
        "warnings": warnings, "list_complete": all(completeness.values()),
        "capabilities": {"catalog": bool(products), "price_competitiveness": bool(prices),
                         "performance": bool(current), "previous_performance": bool(previous),
                         "price_suggestions": bool(suggestion_rows)},
        "definitions": {"at_market_threshold_percent": AT_MARKET_PERCENT},
    }


def sample_report(days=30, start_date=None, end_date=None):
    ranges = date_ranges(days, start_date, end_date)
    fixtures = [
        ("DP-1001", "Wireless ANC Headphones", "Nova", "Electronics", "in_stock", 129.0, 119.0, 1840, 48200, 62, 11740, 114.0, "HIGH"),
        ("DP-1002", "Smart Fitness Watch", "Nova", "Electronics", "out_of_stock", 89.0, 92.0, 1230, 31800, 41, 6970, 84.0, "MEDIUM"),
        ("DP-2001", "Linen Duvet Set", "Casa", "Home & Living", "in_stock", 74.0, 82.0, 910, 22400, 29, 4120, None, "UNKNOWN"),
        ("DP-3001", "Everyday Running Shoes", "Motion", "Sports", "in_stock", 96.0, 95.0, 770, 19800, 23, 3550, None, "UNKNOWN"),
        ("DP-4001", "Daily Face Serum", "Luma", "Beauty", "preorder", 48.0, 43.0, 610, 15200, 18, 2130, 42.0, "HIGH"),
        ("DP-5001", "Recycled Canvas Tote", "Earthline", "Accessories", "in_stock", 34.0, 41.0, 460, 12800, 16, 1350, None, "UNKNOWN"),
    ]
    rows = []
    for index, item in enumerate(fixtures):
        offer, title, brand, category, availability, price, benchmark, clicks, impressions, conversions, value, suggested, effectiveness = item
        gap = (price - benchmark) / benchmark * 100
        status = "below" if gap < -AT_MARKET_PERCENT else "above" if gap > AT_MARKET_PERCENT else "at_market"
        prior_factor = .82 if index % 2 == 0 else 1.12
        rows.append({"id": f"online~en~US~{offer}", "offer_id": offer, "title": title, "brand": brand, "category": category,
                     "availability": availability, "listing_status": "ELIGIBLE", "click_potential": "HIGH" if clicks > 800 else "MEDIUM",
                     "price": {"amount": price, "currency": "USD"}, "benchmark_price": {"amount": benchmark, "currency": "USD"},
                     "price_gap_percent": gap, "price_status": status, "benchmark_country": "US",
                     "performance": {"clicks": clicks, "impressions": impressions, "click_through_rate": clicks/impressions,
                                     "organic_clicks": clicks, "conversions": conversions, "conversion_rate": conversions/clicks, "conversion_value": value, "conversion_currency": "USD"},
                     "previous": {"clicks": round(clicks*prior_factor), "impressions": round(impressions*prior_factor),
                                  "organic_clicks": round(clicks*prior_factor), "click_through_rate": clicks/impressions, "conversions": round(conversions*prior_factor),
                                  "conversion_rate": conversions/clicks, "conversion_value": round(value*prior_factor), "conversion_currency": "USD"},
                     "suggested_price": {"amount": suggested, "currency": "USD"} if suggested else {"amount": None, "currency": ""},
                     "effectiveness": effectiveness, "predicted_impressions_change": .12 if suggested else None,
                     "predicted_clicks_change": .09 if suggested else None, "predicted_conversions_change": .07 if suggested else None})
    positions = {key: [row for row in rows if row["price_status"] == key] for key in ("below", "at_market", "above")}
    brands = []
    for brand in sorted({row["brand"] for row in rows}):
        group = [row for row in rows if row["brand"] == brand]
        brands.append({"brand": brand, "products": len(group), "benchmarked": len(group),
                       "below": sum(row["price_status"] == "below" for row in group),
                       "at_market": sum(row["price_status"] == "at_market" for row in group),
                       "above": sum(row["price_status"] == "above" for row in group),
                       "clicks": sum(row["performance"]["clicks"] for row in group),
                       "impressions": sum(row["performance"]["impressions"] for row in group)})
    summary_current = {key: sum(row["performance"][key] for row in rows) for key in ("clicks", "impressions", "conversions", "conversion_value")} | {"conversion_currency": "USD"}
    summary_previous = {key: sum(row["previous"][key] for row in rows) for key in ("clicks", "impressions", "conversions", "conversion_value")} | {"conversion_currency": "USD"}
    return {"source": "sample", "account": {"id": "sample", "name": "Sample Merchant Center store"},
            "fetched_at": datetime.now(timezone.utc).isoformat(), "start_date": ranges["start"], "end_date": ranges["end"],
            "previous_start": ranges["previous_start"], "previous_end": ranges["previous_end"],
            "summary": {"catalog_products": len(rows), "in_stock": 4, "out_of_stock": 1, "preorder": 1, "backorder": 0,
                        "benchmark_products": len(rows), "benchmark_coverage_rate": 1.0,
                        "price_positions": {key: {"count": len(group), "average_gap_percent": sum(row["price_gap_percent"] for row in group)/len(group) if group else None} for key, group in positions.items()} | {"unavailable": {"count": 0, "average_gap_percent": None}},
                        "performance": summary_current, "previous_performance": summary_previous},
            "rows": rows, "brands": sorted(brands, key=lambda row: row["clicks"], reverse=True),
            "opportunities": {"out_of_stock_demand": [row for row in rows if row["availability"] == "out_of_stock"],
                              "above_benchmark": sorted(positions["above"], key=lambda row: row["performance"]["clicks"], reverse=True),
                              "below_benchmark": sorted(positions["below"], key=lambda row: row["performance"]["clicks"], reverse=True),
                              "price_suggestions": [row for row in rows if row["suggested_price"]["amount"] is not None]},
            "warnings": ["This is illustrative sample data. It does not belong to a Merchant Center account.",
                         "Catalog availability and benchmark prices are snapshots; dates apply to Shopping performance."],
            "list_complete": True, "capabilities": {"catalog": True, "price_competitiveness": True, "performance": True,
                                                         "previous_performance": True, "price_suggestions": True},
            "definitions": {"at_market_threshold_percent": AT_MARKET_PERCENT}}
