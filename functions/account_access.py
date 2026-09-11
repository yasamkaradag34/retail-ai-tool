"""Read-only Stripe subscription verification for Google-verified customers.

Stripe is the source of truth, so no new database or unsigned webhook event is
used to grant access. A bounded, process-local cache limits repeated lookups.
"""
from collections import OrderedDict
import hashlib
import os
import threading
import time

import requests
from fastapi import HTTPException

_cache = OrderedDict()
_lock = threading.Lock()


def _stripe_get(path, secret, params):
    try:
        response = requests.get("https://api.stripe.com/v1/" + path, auth=(secret, ""), params=params, timeout=(5, 15))
        if response.status_code != 200:
            raise ValueError()
        data = response.json()
        if not isinstance(data, dict) or not isinstance(data.get("data"), list):
            raise ValueError()
        return data
    except (requests.RequestException, ValueError):
        raise HTTPException(503, detail={"code": "subscription_unavailable", "message": "Your subscription could not be verified. Please try signing in again shortly."})


def has_paid_subscription(email):
    secret = os.getenv("STRIPE_SECRET_KEY", "").strip()
    prices = {p for p in (os.getenv("STRIPE_STANDARD_PRICE_ID", "").strip(), os.getenv("STRIPE_PRO_PRICE_ID", "").strip()) if p}
    if not secret:
        return False
    fingerprint = hashlib.sha256((secret + "|" + ",".join(sorted(prices))).encode()).hexdigest()
    key = (email.strip().lower(), fingerprint)
    with _lock:
        cached = _cache.get(key)
        if cached and cached[0] > time.monotonic():
            _cache.move_to_end(key)
            return cached[1]
    allowed = False
    params = {"email": key[0], "limit": 100}
    # Paginate duplicate Stripe customers sharing the verified email.
    for _ in range(10):
        customers = _stripe_get("customers", secret, params)
        for customer in customers["data"]:
            if customer.get("email", "").strip().lower() != key[0] or customer.get("deleted"):
                continue
            subscription_params = {"customer": customer["id"], "status": "active", "limit": 100}
            for _ in range(10):
                subscriptions = _stripe_get("subscriptions", secret, subscription_params)
                for sub in subscriptions["data"]:
                    item_prices = {i.get("price", {}).get("id") for i in sub.get("items", {}).get("data", [])}
                    metadata = sub.get("metadata", {})
                    matches = bool(item_prices & prices) or (metadata.get("app") == "dataprovido" and metadata.get("plan") in ("standard", "pro"))
                    if sub.get("status") == "active" and not sub.get("pause_collection") and matches:
                        allowed = True
                        break
                if allowed or not subscriptions.get("has_more") or not subscriptions["data"]:
                    break
                subscription_params["starting_after"] = subscriptions["data"][-1]["id"]
            if allowed:
                break
        if allowed or not customers.get("has_more") or not customers["data"]:
            break
        params["starting_after"] = customers["data"][-1]["id"]
    with _lock:
        _cache[key] = (time.monotonic() + (60 if allowed else 10), allowed)
        _cache.move_to_end(key)
        while len(_cache) > 512:
            _cache.popitem(last=False)
    return allowed
