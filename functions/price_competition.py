import json
import os
from functools import lru_cache

import pandas as pd


COMPANY_INPUT_PATH = "data/company_product_input.xlsx"
MERCHANT_BENCHMARK_SAMPLE_PATH = "data/merchant_price_benchmark_sample.xlsx"


COLUMN_ALIASES = {
    "gtin": ["gtin", "ean", "barcode", "barcodeno", "product_gtin"],
    "sku": ["sku", "item_id", "offer_id", "product_id", "id"],
    "product_title": ["product_title", "title", "name", "product_name", "product"],
    "brand": ["brand", "manufacturer"],
    "cat1": ["cat1", "category1", "category", "category_l1", "main_category"],
    "cat2": ["cat2", "category2", "subcategory", "category_l2", "sub_category"],
    "price": ["price", "your_price", "sale_price", "product_price", "productprice"],
    "currency": ["currency", "currency_code"],
    "stock_qty": ["stock_qty", "stock", "inventory"],
    "reorder_point_qty": ["reorder_point_qty", "reorder_point", "critical_stock"],
    "availability_status": ["availability_status", "availability", "stock_status"],
    "pdp_views": ["pdp_views", "pdp", "pdp_view", "total_unique_pdp_views_sum"],
    "add_to_carts": ["add_to_carts", "a2c", "total_unique_add_to_carts_sum"],
    "transactions": ["transactions", "trans", "orders", "total_transactions_sum"],
    "revenue": ["revenue", "gmv", "sales_amount"],
    "c2d_pct": ["c2d_pct", "c2d"],
    "b2d_pct": ["b2d_pct", "b2d"],

    "benchmark_price": ["benchmark_price", "merchant_benchmark_price", "market_price"],
    "benchmark_currency": ["benchmark_currency", "benchmark_price_currency_code"],
    "country_code": ["country_code", "country_of_sale"],
}


def normalize_columns(df: pd.DataFrame, include_benchmark: bool = True) -> pd.DataFrame:
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
        if not include_benchmark and target in [
            "benchmark_price",
            "benchmark_currency",
            "country_code",
        ]:
            continue

        found = None
        for alias in aliases:
            if alias in df.columns:
                found = alias
                break

        normalized[target] = df[found] if found else None

    return normalized


@lru_cache(maxsize=1)
def load_company_input():
    if not os.path.exists(COMPANY_INPUT_PATH):
        raise FileNotFoundError(
            f"Şirket ürün/funnel input dosyası bulunamadı: {COMPANY_INPUT_PATH}"
        )

    df = pd.read_excel(COMPANY_INPUT_PATH)

    # Şirket datasında benchmark kolonları kullanılmaz.
    # Benchmark sadece Merchant dosyasından gelecek.
    df = normalize_columns(df, include_benchmark=False)

    df["gtin"] = df["gtin"].astype(str).str.strip()
    df["sku"] = df["sku"].astype(str).str.strip()

    numeric_cols = [
        "price",
        "stock_qty",
        "reorder_point_qty",
        "pdp_views",
        "add_to_carts",
        "transactions",
        "revenue",
        "c2d_pct",
        "b2d_pct",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


@lru_cache(maxsize=1)
def load_merchant_benchmark_sample():
    if not os.path.exists(MERCHANT_BENCHMARK_SAMPLE_PATH):
        raise FileNotFoundError(
            f"Merchant benchmark sample dosyası bulunamadı: {MERCHANT_BENCHMARK_SAMPLE_PATH}"
        )

    df = pd.read_excel(MERCHANT_BENCHMARK_SAMPLE_PATH)
    df = normalize_columns(df, include_benchmark=True)

    df["gtin"] = df["gtin"].astype(str).str.strip()
    df["benchmark_price"] = pd.to_numeric(df["benchmark_price"], errors="coerce")

    return df


def classify_price_position(price_gap_pct):
    if pd.isna(price_gap_pct):
        return "benchmark_missing"

    if price_gap_pct >= 10:
        return "strongly_expensive"
    if price_gap_pct >= 5:
        return "expensive"
    if price_gap_pct <= -10:
        return "strongly_cheaper"
    if price_gap_pct <= -5:
        return "cheaper"

    return "parity"


def generate_price_competition_from_uploaded_inputs(
    category: str = "genel",
    period_name: str = "selected_period",
):
    """
    Şirket GTIN'li ürün/funnel datasını Merchant benchmark datası ile join eder.
    Internal benchmark üretmez.
    """

    try:
        company_df = load_company_input()
        benchmark_df = load_merchant_benchmark_sample()
    except Exception as e:
        return json.dumps(
            {
                "analysis_type": "price_competition_error",
                "error": "Input dosyaları okunamadı.",
                "detail": str(e),
                "required_files": [
                    COMPANY_INPUT_PATH,
                    MERCHANT_BENCHMARK_SAMPLE_PATH,
                ],
                "note": "Bu feature internal benchmark üretmez. Merchant benchmark datası zorunludur.",
            },
            ensure_ascii=False,
            default=str,
        )

    if "gtin" not in company_df.columns or company_df["gtin"].isna().all():
        return json.dumps(
            {
                "analysis_type": "price_competition_error",
                "error": "Şirket inputunda GTIN kolonu bulunamadı veya tamamen boş.",
                "required_column": "gtin",
            },
            ensure_ascii=False,
        )

    if "benchmark_price" not in benchmark_df.columns or benchmark_df["benchmark_price"].isna().all():
        return json.dumps(
            {
                "analysis_type": "price_competition_error",
                "error": "Merchant benchmark_price bulunamadı.",
                "required_column": "benchmark_price",
                "note": "Internal benchmark üretilmedi.",
            },
            ensure_ascii=False,
        )

    benchmark_cols = [
        "gtin",
        "benchmark_price",
        "benchmark_currency",
        "country_code",
    ]

    benchmark_cols = [c for c in benchmark_cols if c in benchmark_df.columns]

    merged = company_df.merge(
        benchmark_df[benchmark_cols],
        on="gtin",
        how="left",
    )

    if category and str(category).lower() not in ["genel", "all", "tüm", "tum", "overall"]:
        q = str(category).lower()

        mask = pd.Series(False, index=merged.index)

        for col in ["cat1", "cat2", "brand", "product_title", "sku"]:
            if col in merged.columns:
                mask = mask | merged[col].astype(str).str.lower().str.contains(q, na=False)

        # Mobile sorulunca Telefon / Cep Telefonları datasını da yakalasın
        if q in ["mobile", "mobil", "telefon", "phone"]:
            for col in ["cat1", "cat2", "product_title"]:
                if col in merged.columns:
                    mask = mask | merged[col].astype(str).str.lower().str.contains(
                        "telefon|cep|iphone|samsung|xiaomi",
                        na=False,
                        regex=True,
                    )

        merged = merged[mask].copy()

    if "benchmark_price" not in merged.columns:
        return json.dumps(
            {
                "analysis_type": "price_competition_error",
                "error": "Merge sonrası benchmark_price kolonu oluşmadı.",
                "available_columns": list(merged.columns),
                "note": "Şirket datasında benchmark kullanılmaz; Merchant benchmark dosyası kontrol edilmeli.",
            },
            ensure_ascii=False,
            default=str,
        )

    matched = merged.dropna(subset=["price", "benchmark_price"]).copy()

    if matched.empty:
        return json.dumps(
            {
                "analysis_type": "price_competition_error",
                "error": "Bu filtre için price ve Merchant benchmark_price birlikte dolu ürün bulunamadı.",
                "category": category,
                "uploaded_product_count": int(len(company_df)),
                "filtered_product_count": int(len(merged)),
                "matched_gtin_count": 0,
                "note": "GTIN eşleşmesi yoksa Merchant benchmark üretilemez.",
            },
            ensure_ascii=False,
            default=str,
        )

    matched["price_gap"] = matched["price"] - matched["benchmark_price"]
    matched["price_gap_pct"] = (matched["price_gap"] / matched["benchmark_price"]) * 100
    matched["price_position"] = matched["price_gap_pct"].apply(classify_price_position)

    expensive = matched[
        matched["price_position"].isin(["expensive", "strongly_expensive"])
    ].copy()

    cheaper = matched[
        matched["price_position"].isin(["cheaper", "strongly_cheaper"])
    ].copy()

    parity = matched[matched["price_position"] == "parity"].copy()

    if "b2d_pct" in expensive.columns:
        risky_expensive = expensive.sort_values(
            ["price_gap_pct", "b2d_pct"],
            ascending=[False, True],
        )
    else:
        risky_expensive = expensive.sort_values("price_gap_pct", ascending=False)

    cheap_stock_risk = cheaper.copy()

    if "stock_qty" in cheap_stock_risk.columns:
        cheap_stock_risk = cheap_stock_risk.sort_values(
            ["stock_qty", "price_gap_pct"],
            ascending=[True, True],
        )

    avg_gap = matched["price_gap_pct"].mean()
    median_gap = matched["price_gap_pct"].median()

    if avg_gap >= 5:
        diagnosis = "Genel fiyat pozisyonu Merchant benchmark üstünde; fiyat rekabeti riski var."
    elif avg_gap <= -5:
        diagnosis = "Genel fiyat pozisyonu Merchant benchmark altında; fiyat avantajı var."
    else:
        diagnosis = "Genel fiyat pozisyonu Merchant benchmark ile pariteye yakın."

    actions = []

    if len(expensive) > 0:
        actions.append("Benchmark üstü SKU’larda fiyat gap ve B2D birlikte kontrol edilmeli.")

    if len(cheaper) > 0:
        actions.append("Benchmark altı SKU’larda stok ve margin kontrol edilmeli.")

    if len(cheap_stock_risk) > 0:
        actions.append("Fiyat avantajı olan ürünlerde stok düşükse kampanya artırmadan önce replenishment planlanmalı.")

    actions.append("Bu analiz GTIN üzerinden Merchant benchmark ile eşleşen ürünlere dayanır; internal benchmark kullanılmamıştır.")

    result = {
        "analysis_type": "price_competition_uploaded_inputs",
        "benchmark_mode": "merchant_benchmark_sample_join",
        "benchmark_source": "Merchant Center Price Competitiveness sample / later BigQuery API",
        "category": category,
        "period_name": period_name,
        "summary": {
            "uploaded_product_count": int(len(company_df)),
            "filtered_product_count": int(len(merged)),
            "matched_product_count": int(len(matched)),
            "benchmark_match_rate_pct": round((len(matched) / len(company_df)) * 100, 2) if len(company_df) else 0,
            "avg_price_gap_pct": round(float(avg_gap), 2),
            "median_price_gap_pct": round(float(median_gap), 2),
            "benchmark_above_sku_count": int(len(expensive)),
            "benchmark_below_sku_count": int(len(cheaper)),
            "parity_sku_count": int(len(parity)),
        },
        "main_diagnosis": diagnosis,
        "top_expensive_products": records(risky_expensive.head(10)),
        "top_cheaper_products": records(cheap_stock_risk.head(10)),
        "recommended_actions": actions,
        "caveat": "Internal benchmark kullanılmamıştır. Benchmark yalnızca Merchant benchmark inputundan gelir.",
    }

    return json.dumps(result, ensure_ascii=False, default=str)


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
        "price_gap",
        "price_gap_pct",
        "price_position",
        "stock_qty",
        "pdp_views",
        "add_to_carts",
        "transactions",
        "revenue",
        "c2d_pct",
        "b2d_pct",
    ]

    existing = [c for c in cols if c in df.columns]
    clean = df[existing].copy()
    clean = clean.where(pd.notnull(clean), None)

    return clean.to_dict(orient="records")