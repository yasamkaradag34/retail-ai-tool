import json
import os
import time
import unittest
import base64
import hashlib
import hmac
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

    @patch.object(main, 'SUPABASE_ANON_KEY', 'anon-key')
    @patch.object(main.requests, 'post')
    def test_email_password_login_uses_supabase_without_local_password(self, post):
        post.return_value = Mock(status_code=200, json=lambda: {
            'access_token': 'provider-session-token',
            'user': {'email': 'dataprovido@gmail.com', 'user_metadata': {'full_name': 'DataProvido'}}
        })
        response = self.client.post('/api/auth/login', data={
            'email': 'dataprovido@gmail.com', 'password': 'provider-verified-value'
        }, follow_redirects=False)
        self.assertEqual(response.status_code, 303)
        self.assertIn('/journey', response.headers['location'])
        self.assertIn('gauth=', response.headers['set-cookie'])
        request = post.call_args
        self.assertIn('grant_type=password', request.args[0])
        self.assertEqual(request.kwargs['json']['password'], 'provider-verified-value')

    @patch.object(main, 'SUPABASE_ANON_KEY', 'anon-key')
    @patch.object(main.requests, 'post')
    def test_email_password_login_preserves_same_account_google_grant(self, post):
        post.return_value = Mock(status_code=200, json=lambda: {
            'access_token': 'provider-session-token',
            'user': {'email': 'dataprovido@gmail.com', 'user_metadata': {}},
        })
        self.sign_in(
            access_token='google-access', refresh_token='google-refresh',
            expires_at=time.time() + 3600,
            scope='https://www.googleapis.com/auth/analytics.readonly',
            selected_ga4='123', google_verified=True,
        )
        response = self.client.post('/api/auth/login', data={
            'email': 'dataprovido@gmail.com', 'password': 'provider-verified-value'
        }, follow_redirects=False)
        set_cookie = next(value for value in response.headers.get_list('set-cookie') if value.startswith('gauth='))
        response_cookie = SimpleCookie(set_cookie)['gauth'].value
        session = json.loads(main._decrypt_token(response_cookie))
        self.assertEqual(session['access_token'], 'google-access')
        self.assertEqual(session['refresh_token'], 'google-refresh')
        self.assertEqual(session['selected_ga4'], '123')

    def test_category_workspace_requires_real_ga4_instead_of_auto_sample(self):
        self.sign_in()
        response = self.client.get('/journey?module=category_insights')
        self.assertEqual(response.status_code, 200)
        self.assertIn('CHECKING GA4', response.text)
        self.assertIn('/static/category-analysis.css?v=2', response.text)
        self.assertIn('/static/category-analysis.js?v=2', response.text)
        self.assertIn('class="ca-command-bar"', response.text)
        self.assertIn('class="ca-date-actions"', response.text)
        script = (os.path.join(os.path.dirname(main.__file__), 'static', 'category-analysis.js'))
        script = open(script, encoding='utf-8').read()
        self.assertIn('Real GA4 data is required.', script)
        self.assertNotIn('if (!data.connected && testMode)', script)

    @patch.object(main, 'SUPABASE_ANON_KEY', 'anon-key')
    @patch.object(main.requests, 'post')
    def test_rejected_supabase_login_sets_no_session(self, post):
        post.return_value = Mock(status_code=400, json=lambda: {})
        response = self.client.post('/api/auth/login', data={
            'email': 'dataprovido@gmail.com', 'password': 'wrong'
        }, follow_redirects=False)
        self.assertIn('invalid_credentials', response.headers['location'])
        self.assertNotIn('gauth=', response.headers.get('set-cookie', ''))

    @patch.object(main, 'SUPABASE_ANON_KEY', 'anon-key')
    @patch.object(main.requests, 'put')
    def test_password_recovery_updates_supabase_user(self, put):
        put.return_value = Mock(status_code=200, json=lambda: {
            'id': 'user-id', 'email': 'dataprovido@gmail.com'
        })
        response = self.client.post('/api/auth/reset-password', data={
            'access_token': 'provider-recovery-token',
            'password': 'secure-value-1',
            'confirm_password': 'secure-value-1',
        }, follow_redirects=False)
        self.assertEqual(response.status_code, 303)
        self.assertIn('password_reset_success', response.headers['location'])
        self.assertEqual(put.call_args.kwargs['json'], {'password': 'secure-value-1'})
        self.assertEqual(
            put.call_args.kwargs['headers']['Authorization'],
            'Bearer provider-recovery-token',
        )

    @patch.object(main, 'SUPABASE_ANON_KEY', 'anon-key')
    @patch.object(main.requests, 'put')
    def test_password_recovery_rejects_mismatch_without_provider_call(self, put):
        response = self.client.post('/api/auth/reset-password', data={
            'access_token': 'provider-recovery-token',
            'password': 'secure-value-1',
            'confirm_password': 'different-value-2',
        }, follow_redirects=False)
        self.assertIn('password_mismatch', response.headers['location'])
        put.assert_not_called()

    def test_login_page_handles_recovery_fragment_without_embedded_secret(self):
        response = self.client.get('/login')
        self.assertEqual(response.status_code, 200)
        self.assertIn('id="resetPasswordForm"', response.text)
        self.assertIn("recoveryParams.get('access_token')", response.text)
        self.assertNotIn('provider-recovery-token', response.text)

    @patch.object(main, 'SUPABASE_ANON_KEY', 'anon-key')
    @patch.object(main.requests, 'post')
    def test_forgot_password_uses_production_recovery_redirect(self, post):
        post.return_value = Mock(status_code=200)
        response = self.client.post('/api/auth/forgot-password', data={
            'email': 'dataprovido@gmail.com'
        }, follow_redirects=False)
        self.assertEqual(response.status_code, 303)
        self.assertEqual(
            post.call_args.kwargs['params']['redirect_to'],
            'https://www.dataprovido.com/login',
        )
        self.assertEqual(post.call_args.kwargs['json'], {'email': 'dataprovido@gmail.com'})

    @patch.object(main,'GoogleAnalytics')
    def test_test_account_falls_back_to_accessible_property(self,provider):
        self.sign_in(access_token="test",expires_at=time.time()+5000,selected_ga4="999")
        provider.return_value.properties.return_value=[{"id":"123","name":"Store"}]
        response=self.client.get('/api/ga4/properties')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.json()['selected_property'],'123')
        self.assertEqual(response.headers['cache-control'],'private, no-store')

    @patch.object(main,'has_paid_subscription')
    def test_verified_paid_customer_can_access_journey_and_ga4(self,paid):
        paid.return_value=True
        self.sign_in(email="customer@example.com",google_verified=True,onboarding_complete=True)
        self.assertEqual(self.client.get('/journey',follow_redirects=False).status_code,200)
        self.assertEqual(self.client.get('/api/ga4/properties').status_code,200)
        paid.assert_called_with("customer@example.com")

    @patch.object(main,'has_paid_subscription',return_value=True)
    def test_paid_customer_completes_data_setup_before_journey(self,paid):
        self.sign_in(email="customer@example.com",google_verified=True)
        response=self.client.get('/journey',follow_redirects=False)
        self.assertEqual(response.status_code,303)
        self.assertEqual(response.headers['location'],'/connect-data')
        setup=self.client.get('/connect-data')
        self.assertEqual(setup.status_code,200)
        self.assertIn('Connect your data.',setup.text)

    @patch.object(main,'GoogleAnalytics')
    @patch.object(main,'has_paid_subscription',return_value=True)
    def test_data_setup_saves_property_and_unlocks_journey(self,paid,analytics):
        self.sign_in(
            email="customer@example.com",
            google_verified=True,
            access_token="access",
            expires_at=time.time()+3600,
            scope="https://www.googleapis.com/auth/analytics.readonly",
        )
        response=self.client.post('/connect-data/complete',data={'ga4_property_id':'123'},follow_redirects=False)
        self.assertEqual(response.status_code,303)
        self.assertEqual(response.headers['location'],'/journey?activated=true')
        analytics.return_value.property.assert_called_once_with('123')
        set_cookie=next(value for value in response.headers.get_list('set-cookie') if value.startswith('gauth='))
        response_cookie=SimpleCookie(set_cookie)['gauth'].value
        session=json.loads(main._decrypt_token(response_cookie))
        self.assertTrue(session['onboarding_complete'])
        self.assertEqual(session['selected_ga4'],'123')

    def test_test_account_skips_customer_data_setup(self):
        self.sign_in()
        response=self.client.get('/connect-data',follow_redirects=False)
        self.assertEqual(response.status_code,303)
        self.assertIn('/journey',response.headers['location'])

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

    @patch.object(main,'GOOGLE_CLIENT_ID','test-client')
    def test_unknown_integration_falls_back_to_analytics_without_ads(self):
        response=self.client.get('/api/auth/google?integration=ads',follow_redirects=False)
        params=parse_qs(urlparse(response.headers['location']).query)
        self.assertIn('analytics.readonly',params['scope'][0])
        self.assertNotIn('adwords',params['scope'][0])
        self.assertNotIn('/content',params['scope'][0])

    def test_session_cookie_is_encrypted_and_legacy_plaintext_is_rejected(self):
        payload=json.dumps({'email':'dataprovido@gmail.com','access_token':'secret-google-token'})
        sealed=main._encrypt_token(payload)
        self.assertTrue(sealed.startswith('v2.'))
        self.assertNotIn('dataprovido',sealed)
        self.assertNotIn('secret-google-token',sealed)
        self.assertEqual(main._decrypt_token(sealed),payload)
        legacy_body=base64.urlsafe_b64encode(payload.encode()).decode()
        legacy_signature=hmac.new(main.COOKIE_SECRET.encode(),legacy_body.encode(),hashlib.sha256).hexdigest()[:16]
        self.assertIsNone(main._decrypt_token(legacy_signature+'.'+legacy_body))

    @patch.object(main,'has_paid_subscription',return_value=True)
    @patch.object(main.requests,'get')
    @patch.object(main.requests,'post')
    def test_paid_oauth_callback_enters_data_setup(self,post,get,paid):
        self.client.cookies.set('google_oauth_state',main._encrypt_token(json.dumps({'nonce':'state','expires_at':time.time()+600})))
        post.return_value=Mock(status_code=200,json=lambda:{'access_token':'new','refresh_token':'refresh','expires_in':3600})
        get.return_value=Mock(status_code=200,json=lambda:{'email':'customer@example.com','verified_email':True,'name':'Customer'})
        response=self.client.get('/api/auth/google/callback?code=code&state=state',follow_redirects=False)
        self.assertEqual(response.status_code,303)
        self.assertIn('/connect-data?connected=analytics',response.headers['location'])
        cookie=next(c for c in self.client.cookies.jar if c.name=='gauth')
        user=decoded_cookie(cookie.value)
        self.assertTrue(user['google_verified'])
        self.assertEqual(user['email'],'customer@example.com')
        self.assertIn('HttpOnly',response.headers['set-cookie'])
        self.assertEqual(self.client.get('/journey',follow_redirects=False).status_code,303)

    @patch.object(main,'has_paid_subscription',return_value=True)
    @patch.object(main.requests,'get')
    @patch.object(main.requests,'post')
    def test_merchant_oauth_callback_returns_to_data_setup(self,post,get,paid):
        self.sign_in(
            email='customer@example.com',
            google_verified=True,
            scope='https://www.googleapis.com/auth/analytics.readonly',
        )
        state={'nonce':'merchant-state','integration':'merchant','expires_at':time.time()+600}
        self.client.cookies.set('google_oauth_state',main._encrypt_token(json.dumps(state)))
        post.return_value=Mock(status_code=200,json=lambda:{'access_token':'new','refresh_token':'refresh','expires_in':3600,'scope':'https://www.googleapis.com/auth/content'})
        get.return_value=Mock(status_code=200,json=lambda:{'email':'customer@example.com','verified_email':True,'name':'Customer'})
        response=self.client.get('/api/auth/google/callback?code=code&state=merchant-state',follow_redirects=False)
        self.assertIn('/connect-data?connected=merchant',response.headers['location'])
        set_cookie=next(value for value in response.headers.get_list('set-cookie') if value.startswith('gauth='))
        response_cookie=SimpleCookie(set_cookie)['gauth'].value
        scopes=json.loads(main._decrypt_token(response_cookie))['scope']
        self.assertIn('/auth/content',scopes)
        self.assertIn('analytics.readonly',scopes)

    @patch.object(main,'GoogleAnalytics')
    @patch.object(main.requests,'get')
    @patch.object(main.requests,'post')
    def test_test_account_oauth_selects_injector_marketing_property(self,post,get,analytics):
        state={'nonce':'test-state','integration':'analytics','expires_at':time.time()+600}
        self.client.cookies.set('google_oauth_state',main._encrypt_token(json.dumps(state)))
        post.return_value=Mock(status_code=200,json=lambda:{
            'access_token':'new',
            'refresh_token':'refresh',
            'expires_in':3600,
            'scope':'https://www.googleapis.com/auth/analytics.readonly',
        })
        get.return_value=Mock(status_code=200,json=lambda:{
            'email':'dataprovido@gmail.com',
            'verified_email':True,
            'name':'DataProvido',
        })
        analytics.return_value.properties.return_value=[
            {'id':'111','name':'Other Store'},
            {'id':'222','name':'Injector Marketing - GA4'},
        ]

        response=self.client.get('/api/auth/google/callback?code=code&state=test-state',follow_redirects=False)

        self.assertEqual(response.headers['location'],'/journey?module=category_insights&google_connected=true')
        set_cookie=next(value for value in response.headers.get_list('set-cookie') if value.startswith('gauth='))
        response_cookie=SimpleCookie(set_cookie)['gauth'].value
        self.assertEqual(json.loads(main._decrypt_token(response_cookie))['selected_ga4'],'222')

    def test_merchant_routes_require_console_and_connection(self):
        self.assertEqual(self.client.get('/api/merchant/accounts').status_code,401)
        self.sign_in()
        self.assertEqual(self.client.get('/api/merchant/accounts').json()['accounts'],[])
        self.assertEqual(self.client.get('/api/merchant/insights?account_id=123').status_code,401)

    def test_funnel_sample_is_explicit_and_requires_console_user(self):
        self.assertEqual(self.client.get('/api/ga4/funnel-report?sample=true').status_code,401)
        self.sign_in()
        response=self.client.get('/api/ga4/funnel-report?sample=true')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.json()['source'],'sample')
        self.assertEqual(self.client.get('/api/ga4/funnel-report?property_id=123').status_code,401)

    def test_journey_renders_new_funnel_workspace_and_assets(self):
        self.sign_in()
        response=self.client.get('/journey?module=funnel_analysis')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.text.count('id="funnelWorkspaceContainer"'),1)
        self.assertIn('id="funnelWorkspaceLegacyContainer"',response.text)
        self.assertIn('/static/funnel-analysis.css?v=1',response.text)
        self.assertIn('/static/funnel-analysis.js?v=1',response.text)
        self.assertIn('/static/heatmap-analysis.css?v=1',response.text)
        self.assertIn('Event funnel',response.text)
        self.assertIn('Path exploration',response.text)
        self.assertIn('User exploration',response.text)
        self.assertNotIn('data-key="digital_marketing"',response.text)
        self.assertIn('data-account-mode="test"',response.text)
        self.assertIn('Injector Marketing',response.text)
        self.assertIn('id="hmPathInput"',response.text)
        self.assertIn('Live heatmaps',response.text)

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

    @patch.object(main.requests,'post')
    def test_disconnect_revokes_google_grant_and_clears_session(self,post):
        self.sign_in(access_token='access-token',refresh_token='refresh-token',expires_at=time.time()+3600)
        response=self.client.get('/api/auth/google/disconnect',follow_redirects=False)
        self.assertEqual(response.status_code,303)
        self.assertIn('google_disconnected',response.headers['location'])
        self.assertEqual(post.call_args.args[0],'https://oauth2.googleapis.com/revoke')
        self.assertEqual(post.call_args.kwargs['data'],{'token':'refresh-token'})
        self.assertIn('gauth=',response.headers['set-cookie'])
        self.assertIn('Max-Age=0',response.headers['set-cookie'])

    def test_public_oauth_policy_pages_are_complete(self):
        terms=self.client.get('/terms')
        disclosure=self.client.get('/google-data')
        privacy=self.client.get('/privacy')
        self.assertEqual((terms.status_code,disclosure.status_code,privacy.status_code),(200,200,200))
        self.assertIn('Terms of Service',terms.text)
        self.assertIn('analytics.readonly',disclosure.text)
        self.assertIn('/auth/content',disclosure.text)
        self.assertIn('Disconnect Google',disclosure.text)
        self.assertIn('Google API Data Use',privacy.text)

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
        self.assertIn('Connect Your Data',response.text)
        self.assertNotIn('anything',response.text)

    def test_landing_uses_generic_connection_message_and_crm_card(self):
        response=self.client.get('/')
        self.assertIn('Connect Your Data',response.text)
        self.assertIn('Own CRM User Footprint',response.text)
        self.assertNotIn('DataProvido connects your GA4, Merchant Center and spreadsheet data',response.text)


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
