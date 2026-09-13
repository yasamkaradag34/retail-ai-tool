# DataProvido Google OAuth verification runbook

Last reviewed: 13 September 2026

## Production app

- Google Cloud project: `elevated-nature-406612` (`DataProvido`)
- Audience: External
- Current publishing status: In production
- Branding: verified and published
- Data access verification: pending demo video and final submission
- Homepage: `https://www.dataprovido.com/`
- Privacy Policy: `https://www.dataprovido.com/privacy`
- Terms of Service: `https://www.dataprovido.com/terms`
- Google data disclosure: `https://www.dataprovido.com/google-data`
- Redirect URI configured in production: `https://dataprovido.com/api/auth/google/callback`
- Authorized domain: `dataprovido.com`
- Domain ownership: verified in Google Search Console

## Requested scopes and reviewer justification

### Google Analytics read-only

Scope: `https://www.googleapis.com/auth/analytics.readonly`

Justification:

> DataProvido uses read-only Google Analytics access to list the GA4 accounts and properties available to the connected user and to request reporting dimensions and metrics for user-facing Category & Product Analysis and Funnel Analysis workspaces. Reports include aggregate sessions, active users, engagement, bounce rate, average session duration, ecommerce events, item performance and revenue. DataProvido does not edit the user's Analytics account or property. The scope is requested only when the user selects a GA4 feature.

### Google Merchant Center

Scope: `https://www.googleapis.com/auth/content`

Justification:

> DataProvido uses the Merchant content scope to list Merchant Center accounts available to the connected user and read catalog, availability, Shopping performance, price competitiveness and eligible price insight reports for the user-facing Stock & Price Comparison workspace. Google provides this permission as a read/write scope, but DataProvido's current implementation calls read-only account and report endpoints and does not create, update or delete listings. The scope is requested only when the user selects Connect Merchant Center.

### Identity scopes

- `openid`
- `https://www.googleapis.com/auth/userinfo.email`
- `https://www.googleapis.com/auth/userinfo.profile`

Justification:

> DataProvido uses basic identity information to identify the connected Google account, display the active account to the user, and associate the connection with the correct DataProvido subscription.

Google Ads access is intentionally excluded from this verification request. It requires a separate restricted-scope review and must not be added until the production feature, developer token and review materials are ready.

## Demonstration prerequisites

- A Google test user with access to a populated GA4 property.
- A Google test user with access to an eligible, non-test Merchant Center account and report data.
- Merchant API enabled and the Google Cloud project registered to the primary Merchant Center account.
- Production OAuth callback and stable `COOKIE_SECRET` configured.
- Reviewer can reach the product flow without an unexplained paywall or broken link.

## Cloud Console status

Completed on 13 September 2026:

- Published the OAuth audience to production.
- Verified and published the DataProvido brand.
- Saved the live Terms of Service URL.
- Declared the three non-sensitive identity scopes and the two sensitive product scopes.
- Opened the verification request and replaced the generic scope explanation with the exact production behavior and least-privilege rationale.

The Verification Center currently reports one missing field: a YouTube demo video that shows both sensitive scopes working with real data. The Analytics portion can use a populated GA4 property. The Merchant portion requires a real Merchant Center account available to the recording user; sample UI data is not sufficient for the verification recording.

## Video script

Record the actual production application. Keep Google and DataProvido UI in English. Do not simulate API responses.

1. Open the public DataProvido homepage and show the Google integrations section.
2. Open the Privacy Policy, Terms of Service and Google Data Use pages.
3. Sign in to DataProvido and open Category & Product Analysis.
4. Select Connect Google Analytics.
5. Show the entire Google consent screen, including the exact requested scopes.
6. Approve access and return to DataProvido.
7. Select a real GA4 property and show category, product and funnel reports populated from it.
8. Open Stock & Price Comparison and select Connect Merchant Center.
9. Show the entire incremental Google consent screen and the Merchant permission.
10. Approve access, select a real Merchant Center account and show catalog, availability, performance and pricing reports.
11. Open Google Integrations and select Disconnect Google.
12. Show that the grant was revoked and the local session was removed.

Upload the video to YouTube as Unlisted and confirm it is viewable without requesting access.

## Cloud Console submission order

1. In Branding, confirm the app name, logo, support email, homepage, Privacy Policy, Terms of Service, authorized domain and developer contacts.
2. In Data Access, declare all five scopes listed above and no others.
3. Confirm the production OAuth client has only owned HTTPS origins and the production callback URI.
4. In Audience, select Publish app.
5. In Verification Center, select Prepare for verification.
6. Paste the scope justifications, add the unlisted demo video and provide the three public documentation links.
7. Submit once and monitor both developer contact inboxes. Do not modify branding or scopes while review is in progress unless Google asks for a change.
