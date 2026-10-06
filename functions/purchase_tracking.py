"""Verified Stripe Checkout purchases formatted for GA4 ecommerce tracking."""

import os
import re

import requests


_CHECKOUT_SESSION_ID = re.compile(r"^cs_(?:test_|live_)?[A-Za-z0-9]{8,}$")
_PLANS = {
    "standard": {
        "item_id": "dataprovido_standard",
        "item_name": "DataProvido Standard",
    },
    "pro": {
        "item_id": "dataprovido_pro",
        "item_name": "DataProvido Pro",
    },
}


def _minor_amount(value):
    if isinstance(value, bool):
        return 0
    try:
        return max(0, int(value or 0))
    except (TypeError, ValueError):
        return 0


def _price_ids(session):
    line_items = session.get("line_items", {}).get("data", [])
    if not isinstance(line_items, list):
        return set()
    result = set()
    for line_item in line_items:
        if not isinstance(line_item, dict):
            continue
        price = line_item.get("price")
        price_id = price.get("id") if isinstance(price, dict) else price
        if isinstance(price_id, str) and price_id:
            result.add(price_id)
    return result


def _verified_plan(session):
    metadata = session.get("metadata")
    metadata = metadata if isinstance(metadata, dict) else {}
    metadata_plan = metadata.get("plan") if metadata.get("app") == "dataprovido" else None
    if metadata_plan not in _PLANS:
        metadata_plan = None

    expected_prices = {
        "standard": os.getenv("STRIPE_STANDARD_PRICE_ID", "").strip(),
        "pro": os.getenv("STRIPE_PRO_PRICE_ID", "").strip(),
    }
    actual_prices = _price_ids(session)
    matched_plans = {plan for plan, price_id in expected_prices.items() if price_id and price_id in actual_prices}
    if len(matched_plans) > 1:
        return None
    price_plan = next(iter(matched_plans), None)
    if metadata_plan and price_plan and metadata_plan != price_plan:
        return None
    return metadata_plan or price_plan


def purchase_from_session(session, expected_session_id):
    """Return a PII-free GA4 purchase only for a completed DataProvido payment."""
    if not isinstance(session, dict) or session.get("id") != expected_session_id:
        return None
    if session.get("status") != "complete" or session.get("payment_status") not in {"paid", "no_payment_required"}:
        return None
    if session.get("mode") != "subscription":
        return None

    plan = _verified_plan(session)
    if not plan:
        return None

    currency = str(session.get("currency") or "").upper()
    if len(currency) != 3 or not currency.isalpha():
        return None

    total = _minor_amount(session.get("amount_total"))
    details = session.get("total_details")
    details = details if isinstance(details, dict) else {}
    tax = _minor_amount(details.get("amount_tax"))
    shipping = _minor_amount(details.get("amount_shipping"))
    value = max(0, total - tax - shipping)
    if value <= 0:
        return None

    plan_data = _PLANS[plan]
    value_major = round(value / 100, 2)
    return {
        "plan": plan,
        "ecommerce": {
            "transaction_id": expected_session_id,
            "affiliation": "DataProvido",
            "value": value_major,
            "tax": round(tax / 100, 2),
            "shipping": round(shipping / 100, 2),
            "currency": currency,
            "items": [{
                "item_id": plan_data["item_id"],
                "item_name": plan_data["item_name"],
                "affiliation": "DataProvido",
                "item_brand": "DataProvido",
                "item_category": "SaaS Subscription",
                "item_variant": "Monthly",
                "price": value_major,
                "quantity": 1,
            }],
        },
    }


def get_verified_purchase(session_id):
    """Retrieve a Checkout Session from Stripe and return verified GA4 data."""
    if not isinstance(session_id, str) or not _CHECKOUT_SESSION_ID.fullmatch(session_id):
        return None
    secret = os.getenv("STRIPE_SECRET_KEY", "").strip()
    if not secret:
        return None
    try:
        response = requests.get(
            "https://api.stripe.com/v1/checkout/sessions/" + session_id,
            auth=(secret, ""),
            params={"expand[]": "line_items.data.price.product"},
            timeout=(5, 15),
            allow_redirects=False,
        )
        if response.status_code != 200:
            return None
        return purchase_from_session(response.json(), session_id)
    except (requests.RequestException, ValueError, TypeError):
        return None
