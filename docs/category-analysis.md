# Category & Product Analysis

The workspace at `/journey?module=category_insights` loads GA4 ecommerce reports through the signed-in customer's Google access token. It has Performance, Engagement quality and Measurement tabs, category-to-product navigation, a category filter, search, sortable tables, pagination and CSV export.

## Customer flow

1. Complete a DataProvido Stripe subscription using the same email as the Google account.
2. Continue with Google. The callback validates browser-bound OAuth state and Google's verified email, then verifies an active DataProvido subscription with Stripe.
3. Redirect directly to Category & Product Analysis. A saved accessible property, or the only available property, loads automatically. Multiple properties with no saved choice require one selection.
4. Changing the property or date preset loads a new report. Custom dates use Refresh data. Property selection is saved in the signed session cookie.

Existing administrator accounts remain supported. Customer access is checked against Stripe on protected console, GA4 and App Benchmark requests, with a maximum 60-second positive cache. Only active subscriptions matching configured DataProvido Price IDs or app-created DataProvido subscription metadata qualify. Trialing, incomplete, past-due, canceled and paused-collection subscriptions do not qualify. Checkout query parameters and incoming webhook JSON do not grant access.

## Report semantics

- Item dimensions: `itemCategory`, or `itemId`, `itemName`, `itemCategory` for products. Category drilldown uses an exact, case-sensitive item-category filter.
- Item metrics: `itemsViewed`, `itemsAddedToCart`, `itemsPurchased`, `itemRevenue`, `cartToViewRate`, `purchaseToViewRate`. Quantities are not transactions; rates come from GA4 rather than dividing item counts.
- Quality metrics: `sessions`, `activeUsers`, `engagedSessions`, `bounceRate`, `engagementRate`, `averageSessionDuration`. `checkCompatibility` selects supported metrics for each item breakdown. An unavailable quality report does not replace successful item data.
- Whole-property bounce rate and session duration are requested separately and explicitly labelled. Category filtering does not silently change these cards. Session/user rows are non-additive.
- Both periods are real GA4 queries. The comparison is the immediately preceding equal-length range; presets exclude today in the property's timezone. Rate differences are percentage points.
- Measurement shows property-wide event counts for `view_item`, `add_to_cart`, `begin_checkout`, `purchase`; this is event coverage, not an ordered, unique-user journey.
- Missing category values remain visible. Missing metrics and restricted revenue appear as unavailable. GA4 thresholds, sampling and high-cardinality grouping are disclosed.
- Each report has a 10,000-provider-row cap across both periods. Partial reports are labelled; separately requested totals cover the full scope. CSV contains all loaded rows matching the search, with property, dates, currency, source and completeness metadata. Rate values in CSV are raw GA4 fractions, and session durations are seconds.
- GA4 errors never trigger a demo fallback. Sample mode is explicit, uses only illustrative browser data and marks the export filename and source as sample.

Definitions: [GA4 schema](https://developers.google.com/analytics/devguides/reporting/data/v1/api-schema), [compatibility checks](https://developers.google.com/analytics/devguides/reporting/data/v1/rest/v1beta/properties/checkCompatibility), [ecommerce events](https://developers.google.com/analytics/devguides/collection/ga4/ecommerce), [Google OAuth](https://developers.google.com/identity/protocols/oauth2/web-server), [Stripe subscriptions](https://docs.stripe.com/api/subscriptions/list).

## Environment and release

Use the existing Railway hosting configuration. **Pushing main triggers production and requires explicit user authorization.** The user approved this release for testing on the live site on 2026-09-11.

The server reads environment variables directly; it does not load `.env` files automatically. `.env.example` lists the needed names. Configure Google OAuth credentials, an exact authorized callback URI, Analytics Admin API and Analytics Data API in that Google Cloud project. Google consent must permit the intended customers; a testing-mode OAuth application is limited to configured test users. Grant the connecting Google account Viewer access to its property and collect GA4 ecommerce item events.

Configure `STRIPE_SECRET_KEY` and the Standard/Pro Price IDs, particularly when using existing Stripe Payment Links. New app-created Checkout subscriptions also carry `app=dataprovido` and plan metadata. Configure a stable random `COOKIE_SECRET` shared by workers. The historical public example secrets are rejected in favor of a random process-local fallback, which changes on restart and is unsuitable for multiple workers. Changing the signing secret invalidates old sessions and requires sign-in again.

Google OAuth requests from this workspace and login ask only for Analytics read access plus identity. Other connector consent entry points retain their original scopes. Refresh tokens remain in the HttpOnly signed cookie, following the existing architecture; cookie contents are signed, not encrypted. Successfully refreshed access tokens are persisted.

## Validation

`python -m unittest discover -s tests -p 'test_*.py'` covers reporting, custom/prior dates, product filters, provider failures, restrictions, pagination metadata, paid access, OAuth state, verified email, token refresh and the existing App Benchmark endpoints. `node --check static/category-analysis.js` validates script syntax.

Browser checks use both explicit sample mode and a separate loopback-only provider fixture outside the repository: automatic loading, saved-property restore, denied properties, empty reports, unsupported engagement, category drilldown, search, pagination, custom dates, invalid dates and CSV download. The downloaded CSV was parsed and matched against displayed current/prior amounts, property, currency and dates. Desktop and 390px mobile layouts were inspected.

Local Google/Stripe credentials were not available during implementation, so no real customer purchase or real Google Analytics consent/report was performed. The production OAuth entry point was checked read-only: it currently redirects to Google with the bare-domain callback, and that callback redirects to the www host. Complete an authorized real-account smoke test with the configured credentials before declaring the paid customer onboarding validated in production.
