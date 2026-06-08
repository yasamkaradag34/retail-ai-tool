import json
import os
from functools import lru_cache

import pandas as pd

from functions.sector_profiles import SECTOR_PROFILES


DATA_PATH = "data/ecommerce_ai_sample_data_200_rows.xlsx"
SHEET_NAME = "sample_data_200"


@lru_cache(maxsize=1)
def load_insight_data():
    if not os.path.exists(DATA_PATH):
        raise FileNotFoundError(f"Insight data bulunamadı: {DATA_PATH}")

    df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME)

    df.columns = [
        str(c)
        .strip()
        .replace(" ", "_")
        .replace("%", "pct")
        .replace("/", "_")
        .replace("-", "_")
        .replace("(", "")
        .replace(")", "")
        .replace(".", "")
        .lower()
        for c in df.columns
    ]

    return df


def generate_category_insight(
    category: str = "genel",
    sector: str = "consumer_electronics",
    period_name: str = "selected_period",
):
    """
    Kategori / sektör / dönem bazlı e-ticaret insight üretir.
    Stok, fiyat, funnel, kanal, traffic ve kategori etkisini birlikte yorumlar.
    """

    try:
        df = load_insight_data()
        print(
            "✅ INSIGHT DATA OKUNDU:",
            DATA_PATH,
            "SATIR:",
            len(df),
            "KOLON:",
            len(df.columns),
            flush=True,
        )
    except Exception as e:
        return json.dumps(
            {
                "error": "Insight data okunamadı.",
                "detail": str(e),
                "expected_path": DATA_PATH,
            },
            ensure_ascii=False,
        )

    sector = sector or "consumer_electronics"
    profile = SECTOR_PROFILES.get(sector, SECTOR_PROFILES["marketplace_general"])

    filtered_df = filter_by_category(df, category)

    if filtered_df.empty:
        return json.dumps(
            {
                "error": "Bu kategori için veri bulunamadı.",
                "category": category,
                "available_cat1": sorted(df["cat1"].dropna().unique().tolist()) if "cat1" in df.columns else [],
                "available_cat2": sorted(df["cat2"].dropna().unique().tolist()) if "cat2" in df.columns else [],
            },
            ensure_ascii=False,
        )

    summary = calculate_period_summary(filtered_df)
    channel_insights = calculate_group_insights(filtered_df, "sales_channel")
    traffic_insights = calculate_group_insights(filtered_df, "traffic_channel")

    breakdown_col = "cat2" if "cat2" in filtered_df.columns else "cat1"
    category_performance = calculate_group_insights(filtered_df, breakdown_col)

    winners = category_performance.sort_values("performance_score", ascending=False).head(5)
    losers = category_performance.sort_values("performance_score", ascending=True).head(5)

    root_causes = detect_root_causes(
        df=filtered_df,
        summary=summary,
        channel_insights=channel_insights,
        traffic_insights=traffic_insights,
        sector=sector,
    )

    actions = build_recommendations(
        category=category,
        sector=sector,
        summary=summary,
        root_causes=root_causes,
        winners=winners,
        losers=losers,
    )

    main_diagnosis = build_main_diagnosis(summary, root_causes)

    result = {
        "analysis_type": "category_sector_insight",
        "category": category,
        "sector": sector,
        "sector_name": profile["name"],
        "period_name": period_name,
        "data_scope": {
            "rows": int(len(filtered_df)),
            "brands": safe_unique_count(filtered_df, "brand"),
            "cat1_count": safe_unique_count(filtered_df, "cat1"),
            "cat2_count": safe_unique_count(filtered_df, "cat2"),
            "sku_count": safe_unique_count(filtered_df, "sku"),
        },
        "executive_summary": summary,
        "main_diagnosis": main_diagnosis,
        "channel_insights": df_to_records(channel_insights.head(10)),
        "traffic_insights": df_to_records(traffic_insights.head(10)),
        "winning_categories": df_to_records(winners),
        "losing_categories": df_to_records(losers),
        "root_causes": root_causes,
        "recommended_actions": actions,
        "caveat": (
            "Bu analiz mevcut sample Excel üzerindeki kolonlara göre üretilmiştir. "
            "Gerçek şirket datasında session, spend, margin, return, rating, delivery time gibi ek kolonlar eklenirse insight kalitesi artar."
        ),
    }

    return json.dumps(result, ensure_ascii=False, default=str)


def filter_by_category(df, category):
    if category is None:
        category = "genel"

    q = str(category).lower().strip()

    if q in ["", "genel", "all", "tüm", "tum", "overall", "site"]:
        return df.copy()

    search_cols = [
        c for c in ["cat1", "cat2", "brand", "product", "sku", "sales_channel", "traffic_channel"]
        if c in df.columns
    ]

    if not search_cols:
        return df.copy()

    mask = pd.Series(False, index=df.index)

    for col in search_cols:
        mask = mask | df[col].astype(str).str.lower().str.contains(q, na=False)

    filtered = df[mask].copy()

    if filtered.empty:
        return filtered

    return filtered


def calculate_period_summary(df):
    revenue = safe_sum(df, "revenue")
    transactions = safe_sum(df, "total_transactions_sum")
    pdp_views = safe_sum(df, "total_unique_pdp_views_sum")
    a2c = safe_sum(df, "total_unique_add_to_carts_sum")
    stock_qty = safe_sum(df, "stock_qty")
    lost_revenue = safe_sum(df, "estimated_lost_revenue")

    summary = {
        "revenue": round(revenue, 2),
        "transactions": round(transactions, 2),
        "pdp_views": round(pdp_views, 2),
        "add_to_carts": round(a2c, 2),
        "stock_qty": round(stock_qty, 2),
        "estimated_lost_revenue": round(lost_revenue, 2),

        "revenue_delta_pct": round(weighted_avg(df, "revenue_delta_pct", "revenue"), 2),
        "transactions_delta_pct": round(weighted_avg(df, "transactions_delta_pct", "total_transactions_sum"), 2),
        "pdp_delta_pct": round(weighted_avg(df, "pdp_delta_pct", "total_unique_pdp_views_sum"), 2),
        "a2c_delta_pct": round(weighted_avg(df, "a2c_delta_pct", "total_unique_add_to_carts_sum"), 2),
        "c2d_pct": round(weighted_avg(df, "c2d_pct", "total_unique_pdp_views_sum"), 2),
        "c2d_delta_pct": round(weighted_avg(df, "c2d_delta_pct", "total_unique_pdp_views_sum"), 2),
        "b2d_pct": round(weighted_avg(df, "b2d_pct", "total_unique_pdp_views_sum"), 2),
        "b2d_delta_pct": round(weighted_avg(df, "b2d_delta_pct", "total_unique_pdp_views_sum"), 2),
        "product_price_delta_pct": round(weighted_avg(df, "product_price_delta_pct", "revenue"), 2),
        "bounce_rate_pct": round(weighted_avg(df, "bounce_rate_pct", "total_unique_pdp_views_sum"), 2),
        "bounce_rate_delta_pct": round(weighted_avg(df, "bounce_rate_delta_pct", "total_unique_pdp_views_sum"), 2),
        "checkout_submit_delta_pct": round(weighted_avg(df, "checkout_submit_delta_pct", "checkout_submits"), 2),
    }

    if "stock_coverage_days" in df.columns:
        summary["avg_stock_coverage_days"] = round(weighted_avg(df, "stock_coverage_days", "revenue"), 2)

    summary["critical_stock_sku_count"] = count_critical_stock(df)
    summary["oos_sku_count"] = count_oos(df)

    return summary


def calculate_group_insights(df, group_col):
    if group_col not in df.columns:
        return pd.DataFrame()

    rows = []

    for key, g in df.groupby(group_col, dropna=False):
        revenue = safe_sum(g, "revenue")
        transactions = safe_sum(g, "total_transactions_sum")
        pdp_views = safe_sum(g, "total_unique_pdp_views_sum")
        a2c = safe_sum(g, "total_unique_add_to_carts_sum")

        revenue_delta = weighted_avg(g, "revenue_delta_pct", "revenue")
        transactions_delta = weighted_avg(g, "transactions_delta_pct", "total_transactions_sum")
        pdp_delta = weighted_avg(g, "pdp_delta_pct", "total_unique_pdp_views_sum")
        a2c_delta = weighted_avg(g, "a2c_delta_pct", "total_unique_add_to_carts_sum")
        c2d_delta = weighted_avg(g, "c2d_delta_pct", "total_unique_pdp_views_sum")
        b2d_delta = weighted_avg(g, "b2d_delta_pct", "total_unique_pdp_views_sum")

        performance_score = (
            revenue_delta * 0.35
            + transactions_delta * 0.25
            + pdp_delta * 0.15
            + a2c_delta * 0.15
            + b2d_delta * 0.10
        )

        rows.append(
            {
                group_col: str(key),
                "revenue": round(revenue, 2),
                "transactions": round(transactions, 2),
                "pdp_views": round(pdp_views, 2),
                "add_to_carts": round(a2c, 2),
                "revenue_delta_pct": round(revenue_delta, 2),
                "transactions_delta_pct": round(transactions_delta, 2),
                "pdp_delta_pct": round(pdp_delta, 2),
                "a2c_delta_pct": round(a2c_delta, 2),
                "c2d_delta_pct": round(c2d_delta, 2),
                "b2d_delta_pct": round(b2d_delta, 2),
                "stock_risk_sku_count": count_critical_stock(g),
                "oos_sku_count": count_oos(g),
                "estimated_lost_revenue": round(safe_sum(g, "estimated_lost_revenue"), 2),
                "performance_score": round(performance_score, 2),
            }
        )

    result = pd.DataFrame(rows)

    if result.empty:
        return result

    return result.sort_values("revenue", ascending=False)


def detect_root_causes(df, summary, channel_insights, traffic_insights, sector):
    causes = []

    if summary.get("pdp_delta_pct", 0) < -10:
        causes.append(
            {
                "cause": "Talep / trafik daralması",
                "evidence": f"PDP değişimi %{summary.get('pdp_delta_pct')} seviyesinde.",
                "confidence": "medium",
            }
        )

    if summary.get("revenue_delta_pct", 0) < -10 and summary.get("transactions_delta_pct", 0) < -10:
        causes.append(
            {
                "cause": "Satış ve ciro birlikte düşüyor",
                "evidence": f"Revenue delta %{summary.get('revenue_delta_pct')}, transaction delta %{summary.get('transactions_delta_pct')}.",
                "confidence": "high",
            }
        )

    if summary.get("c2d_delta_pct", 0) > 0 and summary.get("b2d_delta_pct", 0) < 0:
        causes.append(
            {
                "cause": "Sepete ekleme niyeti var ama satın alma bariyeri oluşuyor",
                "evidence": f"C2D delta %{summary.get('c2d_delta_pct')} iken B2D delta %{summary.get('b2d_delta_pct')}.",
                "confidence": "high",
            }
        )

    if summary.get("product_price_delta_pct", 0) > 5 and summary.get("b2d_delta_pct", 0) < 0:
        causes.append(
            {
                "cause": "Fiyat hassasiyeti",
                "evidence": f"Product price delta %{summary.get('product_price_delta_pct')} ve B2D delta %{summary.get('b2d_delta_pct')}.",
                "confidence": "medium",
            }
        )

    if summary.get("critical_stock_sku_count", 0) > 0 or summary.get("oos_sku_count", 0) > 0:
        causes.append(
            {
                "cause": "Stok / bulunabilirlik riski",
                "evidence": f"Kritik stok SKU sayısı {summary.get('critical_stock_sku_count')}, OOS SKU sayısı {summary.get('oos_sku_count')}.",
                "confidence": "medium",
            }
        )

    if summary.get("checkout_submit_delta_pct", 0) < -10 and summary.get("transactions_delta_pct", 0) < 0:
        causes.append(
            {
                "cause": "Checkout / ödeme adımı sürtünmesi",
                "evidence": f"Checkout submit delta %{summary.get('checkout_submit_delta_pct')}, transaction delta %{summary.get('transactions_delta_pct')}.",
                "confidence": "medium",
            }
        )

    paid_vs_unpaid = detect_paid_vs_unpaid_effect(traffic_insights)
    if paid_vs_unpaid:
        causes.append(paid_vs_unpaid)

    channel_shift = detect_channel_shift(channel_insights)
    if channel_shift:
        causes.append(channel_shift)

    if not causes:
        causes.append(
            {
                "cause": "Net tekil problem bulunamadı",
                "evidence": "Metrikler karışık sinyal veriyor. Kategori, kanal ve SKU bazında daha detaylı kırılım önerilir.",
                "confidence": "low",
            }
        )

    return causes


def detect_paid_vs_unpaid_effect(traffic_insights):
    if traffic_insights is None or traffic_insights.empty or "traffic_channel" not in traffic_insights.columns:
        return None

    df = traffic_insights.copy()
    df["traffic_channel_lower"] = df["traffic_channel"].astype(str).str.lower()

    paid_mask = df["traffic_channel_lower"].str.contains("paid|cpc|ppc|ads|search paid|social paid", na=False)
    unpaid_mask = df["traffic_channel_lower"].str.contains("organic|direct|unpaid|seo|crm", na=False)

    paid = df[paid_mask]
    unpaid = df[unpaid_mask]

    if paid.empty or unpaid.empty:
        return None

    paid_delta = paid["transactions_delta_pct"].mean()
    unpaid_delta = unpaid["transactions_delta_pct"].mean()

    if paid_delta < unpaid_delta - 10:
        return {
            "cause": "Paid traffic unpaid kanallara göre daha sert etkileniyor",
            "evidence": f"Paid transaction delta ortalama %{round(paid_delta, 2)}, unpaid transaction delta ortalama %{round(unpaid_delta, 2)}.",
            "confidence": "medium",
        }

    return None


def detect_channel_shift(channel_insights):
    if channel_insights is None or channel_insights.empty or "sales_channel" not in channel_insights.columns:
        return None

    df = channel_insights.copy()
    df["sales_channel_lower"] = df["sales_channel"].astype(str).str.lower()

    app_mobile = df[df["sales_channel_lower"].str.contains("app|mobile", na=False)]
    desktop = df[df["sales_channel_lower"].str.contains("desktop", na=False)]

    if app_mobile.empty or desktop.empty:
        return None

    app_delta = app_mobile["transactions_delta_pct"].mean()
    desktop_delta = desktop["transactions_delta_pct"].mean()

    if app_delta > desktop_delta + 10:
        return {
            "cause": "Mobil/App kanalı Desktop'a göre daha dirençli",
            "evidence": f"App/Mobile transaction delta %{round(app_delta, 2)}, Desktop transaction delta %{round(desktop_delta, 2)}.",
            "confidence": "medium",
        }

    return None


def build_main_diagnosis(summary, root_causes):
    if not root_causes:
        return "Genel performans karma sinyal veriyor."

    top_causes = [c["cause"] for c in root_causes[:3]]

    if "Talep / trafik daralması" in top_causes and "Fiyat hassasiyeti" in top_causes:
        return "Performans düşüşü büyük ölçüde talep daralması ve fiyat hassasiyetiyle açıklanabilir."

    if "Sepete ekleme niyeti var ama satın alma bariyeri oluşuyor" in top_causes:
        return "Kullanıcı ilgisi tamamen kaybolmamış; asıl problem sepete eklemeden satın almaya geçişte görünüyor."

    if "Stok / bulunabilirlik riski" in top_causes:
        return "Kategori performansında stok ve bulunabilirlik riski önemli bir rol oynuyor."

    return "Ana problem: " + ", ".join(top_causes)


def build_recommendations(category, sector, summary, root_causes, winners, losers):
    actions = []

    cause_names = [c["cause"] for c in root_causes]

    if "Talep / trafik daralması" in cause_names:
        actions.append("Talep daralması olan dönemde bütçeyi daha dirençli kanal ve kategorilere kaydır.")

    if "Paid traffic unpaid kanallara göre daha sert etkileniyor" in cause_names:
        actions.append("Paid budget’ı düşük dönüşüm alan kategorilerden, organik talebi güçlü kategorilere taşı.")

    if "Mobil/App kanalı Desktop'a göre daha dirençli" in cause_names:
        actions.append("Mobile/App görünürlüklerini ve kampanya kurgularını güçlendir.")

    if "Sepete ekleme niyeti var ama satın alma bariyeri oluşuyor" in cause_names:
        actions.append("C2D artıp B2D düşen ürünlerde fiyat, kargo, ödeme ve checkout hatalarını kontrol et.")

    if "Fiyat hassasiyeti" in cause_names:
        actions.append("Fiyat artışı yaşayan ve B2D düşen SKU’larda promo veya price test planı çıkar.")

    if "Stok / bulunabilirlik riski" in cause_names:
        actions.append("C2D/B2D güçlü ama stok riski olan SKU’lar için acil replenishment planı yap.")

    if "Checkout / ödeme adımı sürtünmesi" in cause_names:
        actions.append("Checkout submit ve payment adımlarında teknik hata, ödeme başarısızlığı ve UX sorunlarını incele.")

    if winners is not None and not winners.empty:
        top_winner_col = winners.columns[0]
        top_winner = winners.iloc[0][top_winner_col]
        actions.append(f"Kazanan kategori/segment olan '{top_winner}' için görünürlük ve kampanya desteğini artır.")

    if losers is not None and not losers.empty:
        top_loser_col = losers.columns[0]
        top_loser = losers.iloc[0][top_loser_col]
        actions.append(f"En zayıf kategori/segment olan '{top_loser}' için fiyat, stok ve funnel kırılımı detaylandır.")

    if not actions:
        actions.append("Kategori, kanal ve SKU bazında daha detaylı kırılım alarak problem kaynağını daralt.")

    return actions[:8]


def safe_sum(df, col):
    if col not in df.columns:
        return 0.0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def weighted_avg(df, value_col, weight_col=None):
    if value_col not in df.columns:
        return 0.0

    values = pd.to_numeric(df[value_col], errors="coerce")

    if weight_col is None or weight_col not in df.columns:
        return values.dropna().mean() if not values.dropna().empty else 0.0

    weights = pd.to_numeric(df[weight_col], errors="coerce").fillna(0)

    valid = values.notna() & weights.notna()

    if valid.sum() == 0 or weights[valid].sum() == 0:
        return values.dropna().mean() if not values.dropna().empty else 0.0

    return (values[valid] * weights[valid]).sum() / weights[valid].sum()


def count_critical_stock(df):
    if "stock_qty" not in df.columns:
        return 0

    if "reorder_point_qty" in df.columns:
        return int((pd.to_numeric(df["stock_qty"], errors="coerce") <= pd.to_numeric(df["reorder_point_qty"], errors="coerce")).sum())

    if "stock_risk_level" in df.columns:
        return int(df["stock_risk_level"].astype(str).str.lower().str.contains("risk|critical|oos", na=False).sum())

    return 0


def count_oos(df):
    if "stock_qty" in df.columns:
        return int((pd.to_numeric(df["stock_qty"], errors="coerce") == 0).sum())

    if "availability_status" in df.columns:
        return int(df["availability_status"].astype(str).str.lower().eq("out_of_stock").sum())

    return 0


def safe_unique_count(df, col):
    if col not in df.columns:
        return 0
    return int(df[col].nunique(dropna=True))


def df_to_records(df):
    if df is None or df.empty:
        return []

    clean_df = df.copy()
    clean_df = clean_df.where(pd.notnull(clean_df), None)

    return clean_df.to_dict(orient="records")