# Stock & Price Comp.

The workspace reads the signed-in user's Google Merchant Center data without storing report rows. Connect with `integration=merchant`; the OAuth request asks for identity scopes plus `https://www.googleapis.com/auth/content`. Previously granted Google scopes are retained through incremental authorization.

## Live reports

- `product_view`: current offer details, availability, price, listing status and click potential.
- `price_competitiveness_product_view`: current offer price and Google's benchmark price. DataProvido labels a product below the benchmark when the price gap is less than -3%, above the benchmark when it is greater than 3%, and competitively priced inside that range.
- `product_performance_view`: dated Shopping clicks and impressions. Conversions, conversion rate and conversion value are available only for free listings under Google's current definition.
- `price_insights_product_view`: eligible suggested prices, effectiveness buckets and predicted changes.

Catalog, availability, benchmark and suggestion values are current snapshots. The UI date range applies to product performance and compares it with the immediately preceding period of the same length. Conversion values in different currencies are never added together.

Merchant Center does not expose warehouse quantity, reorder points or days of cover through these reports. Add an ERP or inventory feed before presenting low-stock quantities, stock coverage or replenishment recommendations.

Market Insights reports require an eligible Merchant Center account, sufficient data, suitable account permissions and compliance with Google's reporting terms. Provider errors never trigger synthetic live-looking data. The sample workspace is available only through an explicit user action and is always labelled.

Official references:

- [Merchant Reports overview](https://developers.google.com/merchant/api/reference)
- [Merchant Center Query Language](https://developers.google.com/merchant/api/guides/reports/query-language)
- [Performance reports](https://developers.google.com/merchant/api/guides/reports/performance-reports)
- [Market Insights](https://developers.google.com/merchant/api/guides/reports/understand-the-market)

Use the existing Railway hosting configuration. Pushing `main` triggers production and requires explicit user authorization.
