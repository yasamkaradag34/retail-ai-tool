# Dataprovido.com — Privacy & Cookie Documentation

*Template draft — review with a qualified GDPR/KVKK lawyer before publishing. Placeholders in [brackets] must be completed.*

---

## 1. Privacy Policy

### 1.1 Who we are

Dataprovido ("we", "us", "the Platform") is operated by [Legal Company Name], registered at [Address], Türkiye. Contact: [privacy@dataprovido.com].

If you are located in the European Economic Area, your personal data is processed under the General Data Protection Regulation (GDPR). If you are located in Türkiye, your personal data is processed under Law No. 6698 on the Protection of Personal Data (KVKK).

### 1.1a Cross-border / extraterritorial scope

Dataprovido is established in Türkiye. However, because we run advertising campaigns (Google Ads, Meta) targeting individuals located in the European Union/EEA, GDPR applies to us extraterritorially under **Article 3(2)** — this applies to organizations outside the EU that offer goods/services to, or monitor the behavior of, individuals in the EU, which covers systematic ad targeting and retargeting of an EU audience. As a result:

- **EU Representative (GDPR Art. 27):** As our EU advertising activity is systematic rather than occasional, we [have appointed / are required to appoint] a written EU representative established in a Member State where our targeted data subjects are located, to serve as the local point of contact for supervisory authorities and data subjects. Representative details: [EU Representative Name, Address, Email — to be completed].
- **Dual regime:** Turkish visitors and customers are governed primarily by KVKK; EU-based individuals reached through our EU advertising activity are governed primarily by GDPR. Both frameworks apply concurrently to this website since it serves both audiences.
- **Lead authority:** Absent an EU main establishment, complaints from EU data subjects may be lodged with the supervisory authority of the Member State where they reside.

### 1.2 Two roles: Website visitor data vs. Customer business data

Dataprovido has two distinct data relationships, and this policy addresses them separately because the legal basis and our role differ:

- **As a Data Controller** — for visitors to dataprovido.com (this website) and for account/billing data of subscribing brands, we determine the purpose and means of processing. This section applies to that data.
- **As a Data Processor** — brands using the Dataprovido platform connect their own CRM, GA, and other business data into their self-serve workspace. We do not access, view, or process the underlying content of that data; the subscribing brand remains the controller of their own customer data, and Dataprovido acts solely as a processor under a separate Data Processing Agreement (see Section 3).

### 1.3 What we collect (website & account data)

| Category | Examples | Purpose |
|---|---|---|
| Identity & contact | Name, work email, company name | Account creation, billing, support |
| Usage data | Pages visited, session duration, click events | Product analytics, service improvement |
| Technical data | IP address, browser/device type, approximate location | Security, fraud prevention, analytics |
| Marketing interaction data | Ad click IDs, campaign source, conversion events | Measuring ad performance, retargeting |

### 1.4 Legal basis for processing

- **Contract** (Art. 6(1)(b) GDPR / KVKK Art. 5/2-c) — account provisioning, billing, support.
- **Legitimate interest** (Art. 6(1)(f) GDPR / KVKK Art. 5/2-f) — product analytics, security, service improvement.
- **Consent** (Art. 6(1)(a) GDPR / KVKK Art. 5/1) — marketing cookies, Google Ads and Meta advertising/retargeting, non-essential analytics for EEA and Turkish visitors, collected via the cookie consent banner before any non-essential tag fires.

### 1.5 Third-party tools we use on this website

| Tool | Purpose | Data shared | Data location |
|---|---|---|---|
| Google Analytics (GA4) | Website usage analytics | IP (anonymized/truncated), device, behavioral events | Google servers, may include US transfer under SCCs |
| Google Ads | Conversion tracking, retargeting | Click ID (GCLID), conversion events | Google servers |
| Meta Pixel / Conversions API | Ad performance measurement, retargeting | Hashed email/phone (if applicable), device data, event data | Meta servers, may include US transfer under SCCs |

Each of these is a **sub-processor** for website-analytics purposes. We only send data to them after consent is captured through our cookie banner, in line with Google Consent Mode v2 and Meta's data-sharing requirements.

### 1.6 International data transfers

Google and Meta may process data outside the EEA/Türkiye. Transfers rely on Standard Contractual Clauses (SCCs) and, where applicable, the EU-U.S. Data Privacy Framework. For Turkish users, cross-border transfer is carried out in accordance with KVKK Art. 9, relying on the relevant Board-approved transfer mechanism or explicit consent captured in the banner. Because we are established in Türkiye but run advertising to EU audiences (see Section 1.1a), data collected from EU visitors that we transfer to our own systems in Türkiye is itself a transfer out of the EEA and also requires an appropriate mechanism (SCCs) between us and any EU-based ad account or agency partner involved in the campaigns.

### 1.7 Data retention

- Account and billing data: retained for the duration of the subscription plus [X years] for legal/tax obligations.
- Website analytics data: retained per Google Analytics default retention settings ([14 months] recommended) unless configured otherwise.
- Marketing/ad platform data: governed by Google's and Meta's own retention policies.

### 1.8 Your rights

Under GDPR (Art. 15–22) and KVKK (Art. 11), you have the right to: request access to your data, request correction or deletion, object to processing, withdraw consent at any time, and request data portability. To exercise these rights, contact [privacy@dataprovido.com]. Turkish data subjects may also submit a written application per the KVKK application procedure; EEA data subjects have the right to lodge a complaint with their local supervisory authority.

---

## 2. Cookie Policy

### 2.1 What cookies we use

| Category | Cookie/Tag examples | Purpose | Default state |
|---|---|---|---|
| Strictly necessary | Session, load balancing, CSRF token | Site functionality, security | Always on |
| Analytics | `_ga`, `_ga_*` (GA4) | Understand site usage | Off until consent |
| Advertising | Google Ads conversion cookies, Meta Pixel `_fbp`/`_fbc` | Measure and optimize ad campaigns, retargeting | Off until consent |

### 2.2 Consent management

On first visit, users see a consent banner (recommended: Google Consent Mode v2 implementation) allowing them to Accept All, Reject All, or Manage Preferences by category. No analytics or advertising tag fires before explicit consent is given. Users can withdraw or change consent at any time via a persistent "Cookie Preferences" link in the site footer.

### 2.3 How to control cookies

Besides the on-site consent tool, users can manage or block cookies through their browser settings, and can opt out of Google's use of cookies for advertising via [Google Ads Settings](https://adssettings.google.com) and of Meta's via their [Meta Ad Preferences](https://www.facebook.com/adpreferences).

---

## 3. Data Processing Agreement (for subscribing brands) — summary

Because brands connect their own CRM/GA data into their private Dataprovido workspace, this relationship is governed separately from the public-facing privacy policy above:

- Dataprovido acts as **Processor**; the subscribing brand is **Controller** of their connected business data.
- Dataprovido processes that data **only on the documented instructions** of the brand, and only to run the platform's local/on-device AI functions.
- No cross-tenant visibility: one brand's connected data is never accessible to another brand or to Dataprovido staff in the ordinary course of operation.
- Sub-processor list, security measures, breach notification timelines, and audit rights should be detailed in a full DPA — a standard SCC-based template is recommended given the "local/on-premise AI" positioning, since this substantially limits the sub-processor footprint compared to cloud-AI competitors and is a strong compliance selling point worth stating explicitly in the DPA.

---

## 4. Recommended page structure for the website

1. `/privacy-policy` — Section 1 above
2. `/cookie-policy` — Section 2 above, linked from the consent banner's "Manage Preferences"
3. `/dpa` or `/legal/data-processing-agreement` — Section 3, offered to prospective B2B customers during sales/onboarding
4. Footer links: Privacy Policy · Cookie Policy · Cookie Preferences (re-opens the consent modal)

---

*This document is a structural template based on standard GDPR/KVKK requirements and is not legal advice. Have it reviewed by a lawyer qualified in both jurisdictions before publishing, particularly the retention periods, legal-basis wording, and cross-border transfer mechanism, which depend on your final infrastructure and corporate details.*
