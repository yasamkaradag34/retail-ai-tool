import json
import os
import re
from functools import lru_cache

import pandas as pd


COMPANY_INPUT_PATH = "data/company_product_input.xlsx"
MERCHANT_BENCHMARK_SAMPLE_PATH = "data/merchant_price_benchmark_sample.xlsx"


COLUMN_ALIASES = {
    "gtin": ["gtin", "ean", "barcode", "barcodeno", "product_gtin"],
    "sku": ["sku", "item_id", "offer_id", "product_id", "id"],
    "product_title": ["product_title", "title", "name", "product_name", "product"],
    "brand": ["brand", "manufacturer", "marka"],
    "cat1": ["cat1", "category1", "category", "category_l1", "main_category"],
    "cat2": ["cat2", "category2", "subcategory", "category_l2", "sub_category"],
    "sales_channel": ["sales_channel", "channel", "platform"],
    "main_traffic_channel": ["main_traffic_channel", "maintrafficchannel", "traffic_channel"],
    "ref_channel": ["ref_channel", "refchannel", "source_medium"],

    "price": ["price", "your_price", "sale_price", "product_price", "productprice"],
    "benchmark_price": ["benchmark_price", "merchant_benchmark_price", "market_price"],
    "stock_qty": ["stock_qty", "stock", "inventory"],
    "reorder_point_qty": ["reorder_point_qty", "reorder_point", "critical_stock"],
    "availability_status": ["availability_status", "availability", "stock_status"],
    "pdp_views": ["pdp_views", "pdp", "pdp_view", "total_unique_pdp_views_sum"],
    "list_clicks": ["list_clicks", "listclicks", "list_click"],
    "add_to_carts": ["add_to_carts", "a2c", "total_unique_add_to_carts_sum"],
    "transactions": ["transactions", "trans", "orders", "total_transactions_sum"],
    "revenue": ["revenue", "gmv", "sales_amount", "ciro"],
    "c2d_pct": ["c2d_pct", "c2d"],
    "b2d_pct": ["b2d_pct", "b2d"],
    "bounce_rate_pct": ["bounce_rate_pct", "br", "bounce_rate"],
    "stock_coverage_days": ["stock_coverage_days", "stock_coverage"],
    "estimated_lost_revenue": ["estimated_lost_revenue", "lost_revenue"],
}


def normalize_text(value):
    value = str(value or "").lower().strip()
    tr_map = str.maketrans("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC")
    value = value.translate(tr_map).lower()
    value = re.sub(r"\s+", " ", value)
    return value


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df.columns = [
        str(c)
        .strip()
        .lower()
        .replace(" ", "_")
        .replace("-", "_")
        .replace(".", "_")
        .replace("/", "_")
        for c in df.columns
    ]

    normalized = pd.DataFrame()

    for target, aliases in COLUMN_ALIASES.items():
        found = None
        for alias in aliases:
            if alias in df.columns:
                found = alias
                break

        normalized[target] = df[found] if found else None

    return normalized


@lru_cache(maxsize=1)
def load_action_data():
    if not os.path.exists(COMPANY_INPUT_PATH):
        raise FileNotFoundError(
            f"Şirket ürün/funnel input dosyası bulunamadı: {COMPANY_INPUT_PATH}"
        )

    company_df = pd.read_excel(COMPANY_INPUT_PATH)
    company_df = normalize_columns(company_df)

    if os.path.exists(MERCHANT_BENCHMARK_SAMPLE_PATH):
        benchmark_df = pd.read_excel(MERCHANT_BENCHMARK_SAMPLE_PATH)
        benchmark_df = normalize_columns(benchmark_df)

        if "gtin" in company_df.columns and "gtin" in benchmark_df.columns:
            company_df["gtin"] = company_df["gtin"].astype(str).str.strip()
            benchmark_df["gtin"] = benchmark_df["gtin"].astype(str).str.strip()

            benchmark_cols = [
                c for c in ["gtin", "benchmark_price"]
                if c in benchmark_df.columns
            ]

            if "benchmark_price" in benchmark_cols:
                company_df = company_df.drop(columns=["benchmark_price"], errors="ignore")
                company_df = company_df.merge(
                    benchmark_df[benchmark_cols],
                    on="gtin",
                    how="left",
                )

    numeric_cols = [
        "price",
        "benchmark_price",
        "stock_qty",
        "reorder_point_qty",
        "pdp_views",
        "list_clicks",
        "add_to_carts",
        "transactions",
        "revenue",
        "c2d_pct",
        "b2d_pct",
        "bounce_rate_pct",
        "stock_coverage_days",
        "estimated_lost_revenue",
    ]

    for col in numeric_cols:
        if col in company_df.columns:
            company_df[col] = pd.to_numeric(company_df[col], errors="coerce")

    if "price" in company_df.columns and "benchmark_price" in company_df.columns:
        company_df["price_gap"] = company_df["price"] - company_df["benchmark_price"]
        company_df["price_gap_pct"] = (
            company_df["price_gap"] / company_df["benchmark_price"]
        ) * 100

    return company_df


def extract_category_filter(df, question):
    q = normalize_text(question)

    category_synonyms = {
        "mobile": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "mobil": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "gsm": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "telefon": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "tablet": ["tablet", "tabletler", "ipad"],
        "kulaklik": ["kulaklik", "kulaklık", "airpods", "headphone", "headphones"],
        "aksesuar": ["aksesuar", "telefon aksesuarlari", "oyuncu aksesuarlari"],
        "tv": ["tv", "televizyon"],
        "giyilebilir": ["giyilebilir", "watch", "saat"],
        "süpürge": ["supurge", "süpürge", "roborock"],
    }

    for user_word, values in category_synonyms.items():
        if user_word in q:
            return {
                "type": "contains_any",
                "values": values,
            }

    for col in ["cat2", "cat1", "brand"]:
        if col in df.columns:
            unique_values = [
                x for x in df[col].dropna().astype(str).unique().tolist()
                if x and x.lower() != "nan"
            ]

            for value in sorted(unique_values, key=len, reverse=True):
                if normalize_text(value) in q:
                    return {
                        "type": "exact",
                        "column": col,
                        "value": value,
                    }

    return None


def apply_category_filter(df, category_filter):
    if not category_filter:
        return df.copy()

    filtered = df.copy()

    if category_filter.get("type") == "exact":
        col = category_filter.get("column")
        value = category_filter.get("value")

        if col in filtered.columns:
            return filtered[
                filtered[col].astype(str).apply(normalize_text) == normalize_text(value)
            ].copy()

    if category_filter.get("type") == "contains_any":
        values = category_filter.get("values", [])
        mask = pd.Series(False, index=filtered.index)

        for col in ["cat1", "cat2", "product_title", "brand"]:
            if col in filtered.columns:
                for value in values:
                    mask = mask | filtered[col].astype(str).apply(normalize_text).str.contains(
                        normalize_text(value),
                        na=False,
                    )

        return filtered[mask].copy()

    return filtered.copy()


def build_replenishment_plan(df, question):
    work = df.copy()

    if "stock_qty" not in work.columns:
        return error_result(
            question,
            "Stok kolonu bulunamadı. Replenishment planı için stock_qty gerekli."
        )

    if "reorder_point_qty" in work.columns:
        stock_risk_mask = work["stock_qty"] <= work["reorder_point_qty"]
    else:
        stock_threshold = work["stock_qty"].quantile(0.25)
        stock_risk_mask = work["stock_qty"] <= stock_threshold

    if "availability_status" in work.columns:
        stock_risk_mask = stock_risk_mask | work["availability_status"].astype(str).apply(
            normalize_text
        ).str.contains("out|oos|stok yok", na=False)

    c2d_threshold = work["c2d_pct"].median() if "c2d_pct" in work.columns else None
    b2d_threshold = work["b2d_pct"].median() if "b2d_pct" in work.columns else None

    strong_mask = pd.Series(True, index=work.index)

    if c2d_threshold is not None:
        strong_mask = strong_mask & (work["c2d_pct"] >= c2d_threshold)

    if b2d_threshold is not None:
        strong_mask = strong_mask & (work["b2d_pct"] >= b2d_threshold)

    selected = work[stock_risk_mask & strong_mask].copy()

    # Çok kısıtlı olursa sadece stok riskiyle relax et
    if selected.empty:
        selected = work[stock_risk_mask].copy()

    if selected.empty:
        return error_result(
            question,
            "C2D/B2D güçlü ve stok riski taşıyan ürün bulunamadı."
        )

    if "reorder_point_qty" in selected.columns:
        selected["stock_gap_qty"] = (
            selected["reorder_point_qty"].fillna(0) - selected["stock_qty"].fillna(0)
        ).clip(lower=0)
    else:
        selected["stock_gap_qty"] = 0

    if "transactions" in selected.columns:
        selected["daily_sales_proxy"] = selected["transactions"].fillna(0) / 7
    else:
        selected["daily_sales_proxy"] = 0

    selected["suggested_replenishment_qty"] = (
        selected["stock_gap_qty"] + (selected["daily_sales_proxy"] * 14)
    ).round().astype(int)

    selected["priority_score"] = 0

    if "c2d_pct" in selected.columns:
        selected["priority_score"] += selected["c2d_pct"].rank(pct=True) * 30

    if "b2d_pct" in selected.columns:
        selected["priority_score"] += selected["b2d_pct"].rank(pct=True) * 30

    if "revenue" in selected.columns:
        selected["priority_score"] += selected["revenue"].rank(pct=True) * 25

    selected["priority_score"] += selected["stock_gap_qty"].rank(pct=True) * 15

    selected["priority"] = pd.cut(
        selected["priority_score"],
        bins=[-1, 45, 70, 100],
        labels=["Medium", "High", "Critical"],
    ).astype(str)

    selected["recommended_action"] = selected.apply(
        lambda r: (
            "Acil replenishment planla; C2D/B2D güçlü ve stok riski var."
            if r.get("priority") in ["High", "Critical"]
            else "Stok takibi ve replenishment kontrolü yap."
        ),
        axis=1,
    )

    selected["reason"] = selected.apply(
        lambda r: (
            f"Stok: {r.get('stock_qty')}, Reorder point: {r.get('reorder_point_qty')}, "
            f"C2D: %{round(r.get('c2d_pct'), 2) if pd.notna(r.get('c2d_pct')) else 'N/A'}, "
            f"B2D: %{round(r.get('b2d_pct'), 2) if pd.notna(r.get('b2d_pct')) else 'N/A'}"
        ),
        axis=1,
    )

    selected = selected.sort_values(
        ["priority_score", "suggested_replenishment_qty"],
        ascending=[False, False],
    )

    rows = records(selected)

    return json.dumps(
        {
            "analysis_type": "recommended_action_execution",
            "action_type": "replenishment_plan",
            "question": question,
            "summary": {
                "selected_sku_count": int(len(selected)),
                "critical_count": int((selected["priority"] == "Critical").sum()),
                "high_count": int((selected["priority"] == "High").sum()),
                "total_suggested_replenishment_qty": int(selected["suggested_replenishment_qty"].sum()),
            },
            "main_result": "C2D/B2D güçlü ve stok riski taşıyan SKU’lar için replenishment planı oluşturuldu.",
            "rows": rows,
            "recommended_actions": [
                "Critical öncelikli SKU’lar için satın alma / tedarik ekibine replenishment talebi aç.",
                "Yüksek C2D/B2D alan ama stok riski taşıyan ürünlerde kampanya artırmadan önce stok garantisi al.",
                "OOS veya kritik stok ürünlerinde paid traffic ve onsite visibility geçici olarak kontrol edilmeli.",
                "Excel çıktısını tedarik, kategori ve performans pazarlama ekipleriyle paylaş.",
            ],
        },
        ensure_ascii=False,
        default=str,
    )


def build_campaign_boost_plan(df, question):
    work = df.copy()

    category_filter = extract_category_filter(work, question)
    work = apply_category_filter(work, category_filter)

    if work.empty:
        return error_result(question, "Kampanya/görünürlük planı için ilgili segment bulunamadı.")

    sort_col = "revenue" if "revenue" in work.columns else "pdp_views"
    if sort_col in work.columns:
        work = work.sort_values(sort_col, ascending=False)

    work["recommended_action"] = "Görünürlük ve kampanya desteğini artır."
    work["priority"] = "High"
    work["reason"] = "Kazanan segment; revenue/PDP/funnel sinyalleri güçlü."

    rows = records(work.head(20))

    return json.dumps(
        {
            "analysis_type": "recommended_action_execution",
            "action_type": "campaign_boost_plan",
            "question": question,
            "summary": {
                "selected_sku_count": int(len(work)),
                "focus": "winning_segment_visibility",
            },
            "main_result": "Kazanan segment için kampanya ve görünürlük planı oluşturuldu.",
            "rows": rows,
            "recommended_actions": [
                "Kazanan segmentte onsite görünürlüğü artır.",
                "Paid/CRM bütçesini yüksek dönüşüm potansiyeli olan SKU’lara yönlendir.",
                "Stok yeterliliğini kontrol ettikten sonra kampanya desteğini artır.",
            ],
        },
        ensure_ascii=False,
        default=str,
    )


def build_weak_segment_plan(df, question):
    work = df.copy()

    category_filter = extract_category_filter(work, question)
    work = apply_category_filter(work, category_filter)

    if work.empty:
        return error_result(question, "Riskli segment analizi için ilgili kategori/segment bulunamadı.")

    if "b2d_pct" in work.columns:
        work = work.sort_values("b2d_pct", ascending=True)

    work["recommended_action"] = "Fiyat, stok ve funnel kırılımı detaylandır."
    work["priority"] = "High"
    work["reason"] = work.apply(
        lambda r: (
            f"B2D: %{round(r.get('b2d_pct'), 2) if pd.notna(r.get('b2d_pct')) else 'N/A'}, "
            f"Stok: {r.get('stock_qty')}, "
            f"Fiyat: {r.get('price')}, "
            f"Benchmark: {r.get('benchmark_price')}"
        ),
        axis=1,
    )

    rows = records(work.head(20))

    return json.dumps(
        {
            "analysis_type": "recommended_action_execution",
            "action_type": "weak_segment_diagnosis",
            "question": question,
            "summary": {
                "selected_sku_count": int(len(work)),
                "focus": "weak_segment_root_cause",
            },
            "main_result": "Riskli segment için fiyat, stok ve funnel kırılımı hazırlandı.",
            "rows": rows,
            "recommended_actions": [
                "B2D düşük SKU’larda fiyat gap ve benchmark farkını kontrol et.",
                "Stok riski olan ürünlerde replenishment planı çıkar.",
                "PDP yüksek ama B2D düşük ürünlerde ürün sayfası, fiyat, kargo ve ödeme bariyerlerini incele.",
            ],
        },
        ensure_ascii=False,
        default=str,
    )


def execute_recommended_action(question: str, last_result_json: str = None):
    try:
        df = load_action_data()
    except Exception as e:
        return json.dumps(
            {
                "analysis_type": "recommended_action_error",
                "error": "Action executor datası okunamadı.",
                "detail": str(e),
            },
            ensure_ascii=False,
            default=str,
        )

    q = normalize_text(question)

    if any(x in q for x in [
        "replenishment",
        "stok riski",
        "stok",
        "c2d/b2d",
        "c2d b2d",
        "tedarik",
        "bu urunler",
        "bu ürünler",
    ]):
        return build_replenishment_plan(df, question)

    if any(x in q for x in [
        "gorunurluk",
        "görünürlük",
        "kampanya",
        "kazanan",
        "destegini artir",
        "desteğini artır",
    ]):
        return build_campaign_boost_plan(df, question)

    if any(x in q for x in [
        "zayif",
        "zayıf",
        "riskli",
        "kaybeden",
        "detaylandir",
        "detaylandır",
        "fiyat stok funnel",
    ]):
        return build_weak_segment_plan(df, question)

    # Default: aksiyon gibi görünüyor ama sınıflanamadıysa stok/funnel action planına düş
    return build_replenishment_plan(df, question)


def records(df):
    cols = [
        "gtin",
        "sku",
        "product_title",
        "brand",
        "cat1",
        "cat2",
        "price",
        "benchmark_price",
        "price_gap_pct",
        "stock_qty",
        "reorder_point_qty",
        "stock_gap_qty",
        "suggested_replenishment_qty",
        "pdp_views",
        "add_to_carts",
        "transactions",
        "revenue",
        "c2d_pct",
        "b2d_pct",
        "priority",
        "priority_score",
        "recommended_action",
        "reason",
    ]

    existing = [c for c in cols if c in df.columns]
    clean = df[existing].copy()
    clean = clean.where(pd.notnull(clean), None)

    return clean.to_dict(orient="records")


def error_result(question, error):
    return json.dumps(
        {
            "analysis_type": "recommended_action_error",
            "question": question,
            "error": error,
        },
        ensure_ascii=False,
        default=str,
    )