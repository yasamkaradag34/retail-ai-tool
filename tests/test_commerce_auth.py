import json
import os
import time
import unittest
from unittest.mock import Mock, patch
from urllib.parse import parse_qs, urlparse
from http.cookies import SimpleCookie

from fastapi.testclient import TestClient
import main
from functions import account_access


def decoded_cookie(value):
    # Browsers send quoted Set-Cookie values; Starlette removes cookie quoting.
    return json.loads(main._decrypt_token(SimpleCookie('value=' + value)['value'].value))


class CommerceAuthTests(unittest.TestCase):
    def setUp(self):
        self.client=TestClient(main.app,base_url="https://testserver")
        self.addCleanup(self.client.close)

    def sign_in(self,**extra):
        user={"email":"dataprovido@gmail.com",**extra}
        self.client.cookies.set("gauth",main._encrypt_token(json.dumps(user)))

    def test_unauthenticated_cannot_list_or_report(self):
        with patch.object(main,"GoogleAnalytics") as provider:
            for url in ['/api/ga4/properties','/api/ga4/commerce-report?property_id=123','/api/ga4/category-report?property_id=123']:
                self.assertEqual(self.client.get(url).status_code,401)
            provider.assert_not_called()

    def test_email_login_has_empty_connection_instead_of_demo(self):
        self.sign_in()
        data=self.client.get('/api/ga4/properties').json()
        self.assertFalse(data['connected'])
        self.assertEqual(data['properties'],[])
        self.assertEqual(self.client.get('/api/ga4/commerce-report?property_id=123').status_code,401)

    @patch.object(main,'GoogleAnalytics')
    def test_selected_property_belongs_to_returned_account(self,provider):
        self.sign_in(access_token="test",expires_at=time.time()+5000,selected_ga4="999")
        provider.return_value.properties.return_value=[{"id":"123","name":"Store"}]
        response=self.client.get('/api/ga4/properties')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.json()['selected_property'],'')
        self.assertEqual(response.headers['cache-control'],'private, no-store')

    @patch.object(main,'has_paid_subscription')
    def test_verified_paid_customer_can_access_journey_and_ga4(self,paid):
        paid.return_value=True
        self.sign_in(email="customer@example.com",google_verified=True)
        self.assertEqual(self.client.get('/journey',follow_redirects=False).status_code,200)
        self.assertEqual(self.client.get('/api/ga4/properties').status_code,200)
        paid.assert_called_with("customer@example.com")

    @patch.object(main,'has_paid_subscription')
    def test_unverified_email_cannot_claim_paid_access(self,paid):
        paid.return_value=True
        self.sign_in(email="customer@example.com")
        self.assertEqual(self.client.get('/api/ga4/properties').status_code,403)
        paid.assert_not_called()

    @patch.object(main,'has_paid_subscription',return_value=False)
    def test_cancelled_or_unpaid_customer_denied(self,paid):
        self.sign_in(email="customer@example.com",google_verified=True)
        response=self.client.get('/journey',follow_redirects=False)
        self.assertIn('subscription_required',response.headers['location'])
        self.assertEqual(self.client.get('/api/ga4/properties').status_code,403)

    @patch.object(main.requests,'post')
    def test_invalid_state_never_exchanges_oauth_code(self,post):
        for cookie in [None,main._encrypt_token(json.dumps({'nonce':'one','expires_at':time.time()-1})),main._encrypt_token(json.dumps({'nonce':'one','expires_at':time.time()+600}))]:
            self.client.cookies.clear()
            if cookie: self.client.cookies.set('google_oauth_state',cookie)
            response=self.client.get('/api/auth/google/callback?code=code&state=wrong',follow_redirects=False)
            self.assertIn('oauth_state_expired',response.headers['location'])
        post.assert_not_called()

    @patch.object(main,'GOOGLE_CLIENT_ID','test-client')
    def test_analytics_consent_is_browser_bound_and_minimal(self):
        response=self.client.get('/api/auth/google?integration=analytics',follow_redirects=False)
        params=parse_qs(urlparse(response.headers['location']).query)
        cookie=next(c for c in self.client.cookies.jar if c.name=='google_oauth_state')
        payload=decoded_cookie(cookie.value)
        self.assertEqual(params['state'][0],payload['nonce'])
        self.assertIn('analytics.readonly',params['scope'][0])
        self.assertNotIn('adwords',params['scope'][0])
        self.assertNotIn('/content',params['scope'][0])

    @patch.object(main,'GOOGLE_CLIENT_ID','test-client')
    def test_merchant_consent_requests_only_content_and_identity(self):
        response=self.client.get('/api/auth/google?integration=merchant',follow_redirects=False)
        params=parse_qs(urlparse(response.headers['location']).query)
        cookie=next(c for c in self.client.cookies.jar if c.name=='google_oauth_state')
        payload=decoded_cookie(cookie.value)
        self.assertEqual(payload['integration'],'merchant')
        self.assertIn('/auth/content',params['scope'][0])
        self.assertNotIn('analytics.readonly',params['scope'][0])
        self.assertNotIn('adwords',params['scope'][0])
        self.assertEqual(params['include_granted_scopes'],['true'])

    @patch.object(main,'has_paid_subscription',return_value=True)
    @patch.object(main.requests,'get')
    @patch.object(main.requests,'post')
    def test_paid_oauth_callback_goes_directly_to_category(self,post,get,paid):
        self.client.cookies.set('google_oauth_state',main._encrypt_token(json.dumps({'nonce':'state','expires_at':time.time()+600})))
        post.return_value=Mock(status_code=200,json=lambda:{'access_token':'new','refresh_token':'refresh','expires_in':3600})
        get.return_value=Mock(status_code=200,json=lambda:{'email':'customer@example.com','verified_email':True,'name':'Customer'})
        response=self.client.get('/api/auth/google/callback?code=code&state=state',follow_redirects=False)
        self.assertEqual(response.status_code,303)
        self.assertIn('module=category_insights',response.headers['location'])
        cookie=next(c for c in self.client.cookies.jar if c.name=='gauth')
        user=decoded_cookie(cookie.value)
        self.assertTrue(user['google_verified'])
        self.assertEqual(user['email'],'customer@example.com')
        self.assertIn('HttpOnly',response.headers['set-cookie'])
        self.assertEqual(self.client.get('/journey',follow_redirects=False).status_code,200)

    @patch.object(main,'has_paid_subscription',return_value=True)
    @patch.object(main.requests,'get')
    @patch.object(main.requests,'post')
    def test_merchant_oauth_callback_returns_to_stock_workspace(self,post,get,paid):
        state={'nonce':'merchant-state','integration':'merchant','expires_at':time.time()+600}
        self.client.cookies.set('google_oauth_state',main._encrypt_token(json.dumps(state)))
        post.return_value=Mock(status_code=200,json=lambda:{'access_token':'new','refresh_token':'refresh','expires_in':3600,'scope':'https://www.googleapis.com/auth/content'})
        get.return_value=Mock(status_code=200,json=lambda:{'email':'customer@example.com','verified_email':True,'name':'Customer'})
        response=self.client.get('/api/auth/google/callback?code=code&state=merchant-state',follow_redirects=False)
        self.assertIn('module=stock_price_comp',response.headers['location'])
        cookie=next(c for c in self.client.cookies.jar if c.name=='gauth')
        self.assertIn('/auth/content',decoded_cookie(cookie.value)['scope'])

    def test_merchant_routes_require_console_and_connection(self):
        self.assertEqual(self.client.get('/api/merchant/accounts').status_code,401)
        self.sign_in()
        self.assertEqual(self.client.get('/api/merchant/accounts').json()['accounts'],[])
        self.assertEqual(self.client.get('/api/merchant/insights?account_id=123').status_code,401)

    @patch.object(main,'GoogleMerchant')
    def test_merchant_route_requires_content_scope_before_provider_call(self,provider):
        self.sign_in(access_token='test',expires_at=time.time()+3600,scope='https://www.googleapis.com/auth/analytics.readonly')
        response=self.client.get('/api/merchant/accounts')
        self.assertEqual(response.status_code,403)
        self.assertEqual(response.json()['detail']['code'],'merchant_scope_required')
        provider.assert_not_called()

    @patch.object(main,'GoogleMerchant')
    def test_merchant_selection_rechecks_access_and_origin(self,provider):
        self.sign_in(access_token='test',expires_at=time.time()+3600,scope=main.MERCHANT_SCOPE)
        denied=self.client.post('/api/merchant/selection',json={'account_id':'123'},headers={'Origin':'https://evil.example'})
        self.assertEqual(denied.status_code,403)
        provider.return_value.account.assert_not_called()
        saved=self.client.post('/api/merchant/selection',json={'account_id':'123'},headers={'Origin':'https://testserver'})
        self.assertEqual(saved.status_code,200)
        provider.return_value.account.assert_called_once_with('123')

    @patch.object(main,'has_paid_subscription')
    @patch.object(main.requests,'get')
    @patch.object(main.requests,'post')
    def test_google_unverified_email_rejected_before_stripe(self,post,get,paid):
        self.client.cookies.set('google_oauth_state',main._encrypt_token(json.dumps({'nonce':'s','expires_at':time.time()+600})))
        post.return_value=Mock(status_code=200,json=lambda:{'access_token':'new'})
        get.return_value=Mock(status_code=200,json=lambda:{'email':'customer@example.com','verified_email':False})
        response=self.client.get('/api/auth/google/callback?code=code&state=s',follow_redirects=False)
        self.assertIn('google_email_unverified',response.headers['location'])
        paid.assert_not_called()

    @patch.object(main,'GoogleAnalytics')
    @patch.object(main.requests,'post')
    def test_expired_token_refresh_is_persisted(self,post,provider):
        self.sign_in(access_token='old',refresh_token='refresh',expires_at=1,selected_ga4='123')
        post.return_value=Mock(status_code=200,json=lambda:{'access_token':'new','expires_in':3600})
        provider.return_value.properties.return_value=[{'id':'123','name':'Store'}]
        response=self.client.get('/api/ga4/properties')
        self.assertEqual(response.status_code,200)
        provider.assert_called_with('new')
        self.assertIn('gauth=',response.headers['set-cookie'])
        cookie=next(c for c in self.client.cookies.jar if c.name=='gauth' and c.domain=='testserver.local')
        self.assertEqual(decoded_cookie(cookie.value)['access_token'],'new')

    @patch.object(main.requests,'post')
    def test_failed_refresh_does_not_reuse_expired_token(self,post):
        self.sign_in(access_token='old',refresh_token='refresh',expires_at=1)
        post.return_value=Mock(status_code=400)
        with patch.object(main,'GoogleAnalytics') as provider:
            response=self.client.get('/api/ga4/properties')
            self.assertFalse(response.json()['connected'])
            provider.assert_not_called()

    @patch.object(main,'GoogleAnalytics')
    def test_save_rechecks_property_access_and_origin(self,provider):
        self.sign_in(access_token='test',expires_at=time.time()+3600)
        response=self.client.post('/api/ga4/selection',json={'property_id':'123'},headers={'Origin':'https://evil.example'})
        self.assertEqual(response.status_code,403)
        provider.return_value.property.assert_not_called()
        response=self.client.post('/api/ga4/selection',json={'property_id':'123'},headers={'Origin':'https://testserver'})
        self.assertEqual(response.status_code,200)
        provider.return_value.property.assert_called_once_with('123')
        self.assertIn('gauth=',response.headers['set-cookie'])

    def test_checkout_return_does_not_claim_payment_verification(self):
        response=self.client.get('/checkout/success?plan=pro&session_id=anything')
        self.assertNotIn('Payment Successful!',response.text)
        self.assertIn('Continue with Google',response.text)
        self.assertNotIn('anything',response.text)


class SubscriptionTests(unittest.TestCase):
    def setUp(self):
        account_access._cache.clear()
        self.env=patch.dict(os.environ,{'STRIPE_SECRET_KEY':'test-key','STRIPE_STANDARD_PRICE_ID':'price_standard','STRIPE_PRO_PRICE_ID':'price_pro'})
        self.env.start(); self.addCleanup(self.env.stop)

    def verify(self,subscription):
        with patch.object(account_access,'_stripe_get',side_effect=[{'data':[{'id':'cus_1','email':'customer@example.com'}]}, {'data':[subscription]}]):
            return account_access.has_paid_subscription('customer@example.com')

    def test_only_active_dataprovido_subscription_grants_access(self):
        self.assertTrue(self.verify({'status':'active','items':{'data':[{'price':{'id':'price_standard'}}]}}))

    def test_other_product_cancelled_trial_and_past_due_denied(self):
        for status,price in [('active','price_unrelated'),('canceled','price_standard'),('trialing','price_standard'),('past_due','price_standard'),('incomplete','price_standard')]:
            account_access._cache.clear()
            with self.subTest(status=status,price=price):
                self.assertFalse(self.verify({'status':status,'items':{'data':[{'price':{'id':price}}]}}))

    def test_app_created_metadata_subscription_is_recognized(self):
        self.assertTrue(self.verify({'status':'active','metadata':{'app':'dataprovido','plan':'pro'},'items':{'data':[]}}))

    def test_paused_collection_denied(self):
        self.assertFalse(self.verify({'status':'active','pause_collection':{'behavior':'void'},'items':{'data':[{'price':{'id':'price_standard'}}]}}))

    @patch.object(account_access,'_stripe_get')
    def test_mismatched_customer_email_denied(self,get):
        get.return_value={'data':[{'id':'cus_1','email':'someoneelse@example.com'}]}
        self.assertFalse(account_access.has_paid_subscription('customer@example.com'))
        self.assertEqual(get.call_count,1)


if __name__ == '__main__':
    unittest.main()
