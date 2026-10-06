import os
import unittest
from unittest.mock import Mock, patch

from functions import purchase_tracking


class PurchaseTrackingTests(unittest.TestCase):
    session_id = "cs_test_1234567890"

    def session(self, **updates):
        value = {
            "id": self.session_id,
            "status": "complete",
            "payment_status": "paid",
            "mode": "subscription",
            "currency": "eur",
            "amount_total": 23900,
            "total_details": {"amount_tax": 4000, "amount_shipping": 0},
            "metadata": {"app": "dataprovido", "plan": "standard"},
            "customer_details": {"email": "must-not-enter-datalayer@example.com"},
            "line_items": {"data": [{"price": {"id": "price_standard"}}]},
        }
        value.update(updates)
        return value

    def test_verified_paid_session_becomes_ga4_purchase_without_pii(self):
        purchase = purchase_tracking.purchase_from_session(self.session(), self.session_id)
        self.assertEqual(purchase["plan"], "standard")
        ecommerce = purchase["ecommerce"]
        self.assertEqual(ecommerce["transaction_id"], self.session_id)
        self.assertEqual(ecommerce["currency"], "EUR")
        self.assertEqual(ecommerce["value"], 199.0)
        self.assertEqual(ecommerce["tax"], 40.0)
        self.assertEqual(ecommerce["items"][0]["item_id"], "dataprovido_standard")
        self.assertNotIn("email", str(purchase).lower())

    def test_unpaid_incomplete_or_foreign_sessions_are_rejected(self):
        self.assertIsNone(purchase_tracking.purchase_from_session(self.session(payment_status="unpaid"), self.session_id))
        self.assertIsNone(purchase_tracking.purchase_from_session(self.session(status="open"), self.session_id))
        self.assertIsNone(purchase_tracking.purchase_from_session(self.session(metadata={}), self.session_id))
        self.assertIsNone(purchase_tracking.purchase_from_session(self.session(id="cs_test_other1234"), self.session_id))

    @patch.dict(os.environ, {
        "STRIPE_STANDARD_PRICE_ID": "price_standard",
        "STRIPE_PRO_PRICE_ID": "price_pro",
    }, clear=False)
    def test_known_price_identifies_payment_link_session_without_metadata(self):
        purchase = purchase_tracking.purchase_from_session(self.session(metadata={}), self.session_id)
        self.assertEqual(purchase["plan"], "standard")

    @patch.dict(os.environ, {"STRIPE_SECRET_KEY": "sk_test_value"}, clear=False)
    @patch.object(purchase_tracking.requests, "get")
    def test_session_is_retrieved_from_fixed_stripe_endpoint(self, get):
        get.return_value = Mock(status_code=200, json=lambda: self.session())
        purchase = purchase_tracking.get_verified_purchase(self.session_id)
        self.assertEqual(purchase["ecommerce"]["transaction_id"], self.session_id)
        get.assert_called_once_with(
            "https://api.stripe.com/v1/checkout/sessions/" + self.session_id,
            auth=("sk_test_value", ""),
            params={"expand[]": "line_items.data.price.product"},
            timeout=(5, 15),
            allow_redirects=False,
        )

    @patch.object(purchase_tracking.requests, "get")
    def test_invalid_session_id_never_reaches_stripe(self, get):
        self.assertIsNone(purchase_tracking.get_verified_purchase("../../customers"))
        get.assert_not_called()


if __name__ == "__main__":
    unittest.main()
