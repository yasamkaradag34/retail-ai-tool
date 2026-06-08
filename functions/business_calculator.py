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
def load_business_data():
    if not os.path.exists(COMPANY_INPUT_PATH):
        raise FileNotFoundError(
            f"Şirket input dosyası bulunamadı: {COMPANY_INPUT_PATH}"
        )

    company_df = pd.read_excel(COMPANY_INPUT_PATH)
    company_df = normalize_columns(company_df)

    if os.path.exists(MERCHANT_BENCHMARK_SAMPLE_PATH):
        benchmark_df = pd.read_excel(MERCHANT_BENCHMARK_SAMPLE_PATH)
        benchmark_df = normalize_columns(benchmark_df)

        benchmark_cols = ["gtin", "benchmark_price"]
        benchmark_cols = [c for c in benchmark_cols if c in benchmark_df.columns]

        if "gtin" in benchmark_cols and "benchmark_price" in benchmark_cols:
            company_df["gtin"] = company_df["gtin"].astype(str).str.strip()
            benchmark_df["gtin"] = benchmark_df["gtin"].astype(str).str.strip()

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


def detect_aggregation(question):
    q = normalize_text(question)

    if any(x in q for x in ["ortalama", "average", "avg", "mean"]):
        return "avg"

    if any(x in q for x in ["toplam", "total", "sum", "ciro toplam"]):
        return "sum"

    if any(x in q for x in ["kac", "kaç", "adet", "sayisi", "sayısı", "count"]):
        return "count"

    if any(x in q for x in ["en yuksek", "en yüksek", "max", "maksimum"]):
        return "max"

    if any(x in q for x in ["en dusuk", "en düşük", "min", "minimum"]):
        return "min"

    if any(x in q for x in ["medyan", "median"]):
        return "median"

    return "avg"


def detect_metric(question):
    q = normalize_text(question)

    metric_rules = [
        ("benchmark_price", ["benchmark fiyat", "benchmark_price", "merchant fiyat", "rakip fiyat"]),
        ("price_gap_pct", ["price gap", "fiyat farki", "fiyat farkı", "gap", "benchmark fark"]),
        ("price", ["fiyat", "price", "ortalama fiyat"]),
        ("revenue", ["revenue", "ciro", "gmv", "satis tutari", "satış tutarı"]),
        ("transactions", ["transaction", "transactions", "trans", "siparis", "sipariş", "satis adedi", "satış adedi"]),
        ("pdp_views", ["pdp", "pdp view", "goruntulenme", "görüntülenme"]),
        ("list_clicks", ["list click", "listclick"]),
        ("add_to_carts", ["a2c", "add to cart", "sepete ekleme"]),
        ("c2d_pct", ["c2d", "cart to detail"]),
        ("b2d_pct", ["b2d", "buy to detail"]),
        ("bounce_rate_pct", ["bounce", "br", "bounce rate"]),
        ("stock_qty", ["stok", "stock", "envanter"]),
        ("stock_coverage_days", ["stock coverage", "stok coverage", "stok gun", "stok gün"]),
        ("estimated_lost_revenue", ["lost revenue", "kayip ciro", "kayıp ciro"]),
    ]

    for metric, keywords in metric_rules:
        if any(k in q for k in keywords):
            return metric

    return "price"


def detect_group_by(question):
    q = normalize_text(question)

    if any(x in q for x in ["marka bazinda", "marka kırılım", "brand bazinda", "brand kırılım"]):
        return "brand"

    if any(x in q for x in ["kategori bazinda", "kategori kırılım", "cat1"]):
        return "cat1"

    if any(x in q for x in ["alt kategori", "cat2", "subcategory"]):
        return "cat2"

    if any(x in q for x in ["kanal bazinda", "traffic", "trafik kanali"]):
        return "main_traffic_channel"

    return None


def detect_filters(df, question):
    q = normalize_text(question)
    filters = []

    # Marka filtresi
    if "brand" in df.columns:
        brands = [
            x for x in df["brand"].dropna().astype(str).unique().tolist()
            if x and x.lower() != "nan"
        ]

        for brand in sorted(brands, key=len, reverse=True):
            if normalize_text(brand) in q:
                filters.append({
                    "column": "brand",
                    "value": brand,
                    "match_type": "exact"
                })
                return filters

    # Kategori filtresi
    category_synonyms = {
        "mobile": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "mobil": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "gsm": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "telefon": ["gsm", "telefon", "cep", "cep telefonlari", "iphone", "galaxy"],
        "tablet": ["tablet", "tabletler", "ipad"],
        "kulaklik": ["kulaklik", "kulaklık", "airpods", "headphone", "headphones"],
        "aksesuar": ["aksesuar", "telefon aksesuarlari", "oyuncu aksesuarlari"],
        "tv": ["tv", "televizyon"],
    }

    for user_word, possible_values in category_synonyms.items():
        if user_word in q:
            filters.append({
                "column": "multi_category",
                "value": possible_values,
                "match_type": "contains_any"
            })
            return filters

    # Cat1 / Cat2 doğrudan eşleşme
    for col in ["cat1", "cat2"]:
        if col in df.columns:
            values = [
                x for x in df[col].dropna().astype(str).unique().tolist()
                if x and x.lower() != "nan"
            ]

            for value in sorted(values, key=len, reverse=True):
                if normalize_text(value) in q:
                    filters.append({
                        "column": col,
                        "value": value,
                        "match_type": "exact"
                    })
                    return filters

    return filters


def apply_filters(df, filters):
    filtered = df.copy()

    for f in filters:
        col = f.get("column")
        value = f.get("value")
        match_type = f.get("match_type")

        if col == "multi_category":
            mask = pd.Series(False, index=filtered.index)
            values = value if isinstance(value, list) else [value]

            for category_col in ["cat1", "cat2", "product_title"]:
                if category_col in filtered.columns:
                    for v in values:
                        mask = mask | filtered[category_col].astype(str).apply(
                            normalize_text
                        ).str.contains(normalize_text(v), na=False)

            filtered = filtered[mask].copy()

        elif col in filtered.columns:
            if match_type == "exact":
                filtered = filtered[
                    filtered[col].astype(str).apply(normalize_text)
                    == normalize_text(value)
                ].copy()
            else:
                filtered = filtered[
                    filtered[col].astype(str).apply(normalize_text)
                    .str.contains(normalize_text(value), na=False)
                ].copy()

    return filtered


def calculate_scalar(series, aggregation):
    series = pd.to_numeric(series, errors="coerce").dropna()

    if aggregation == "avg":
        return series.mean()

    if aggregation == "sum":
        return series.sum()

    if aggregation == "count":
        return series.count()

    if aggregation == "max":
        return series.max()

    if aggregation == "min":
        return series.min()

    if aggregation == "median":
        return series.median()

    return series.mean()


def format_number(value):
    if pd.isna(value):
        return None

    try:
        value = float(value)
    except Exception:
        return value

    return round(value, 2)


def calculate_business_metric(question: str):
    try:
        df = load_business_data()
    except Exception as e:
        return json.dumps(
            {
                "analysis_type": "business_metric_error",
                "error": "Business calculator datası okunamadı.",
                "detail": str(e),
            },
            ensure_ascii=False,
            default=str,
        )

    metric = detect_metric(question)
    aggregation = detect_aggregation(question)
    group_by = detect_group_by(question)
    filters = detect_filters(df, question)

    filtered = apply_filters(df, filters)

    if filtered.empty:
        return json.dumps(
            {
                "analysis_type": "business_metric_error",
                "error": "Filtre sonrası veri bulunamadı.",
                "question": question,
                "detected_filters": filters,
                "hint": "Marka, kategori veya ürün adı Excel'deki değerlerle eşleşmemiş olabilir.",
                "available_brands": sorted(df["brand"].dropna().astype(str).unique().tolist())[:30] if "brand" in df.columns else [],
                "available_cat1": sorted(df["cat1"].dropna().astype(str).unique().tolist())[:30] if "cat1" in df.columns else [],
                "available_cat2": sorted(df["cat2"].dropna().astype(str).unique().tolist())[:30] if "cat2" in df.columns else [],
            },
            ensure_ascii=False,
            default=str,
        )

    if metric not in filtered.columns:
        return json.dumps(
            {
                "analysis_type": "business_metric_error",
                "error": f"'{metric}' metriği datada bulunamadı.",
                "question": question,
                "available_columns": list(filtered.columns),
            },
            ensure_ascii=False,
            default=str,
        )

    # Top / bottom istekleri
    q = normalize_text(question)
    if any(x in q for x in ["en yuksek", "en yüksek", "top 10", "ilk 10", "en fazla"]):
        result_df = filtered.sort_values(metric, ascending=False).head(10)
        rows = result_df[
            [c for c in [
                "sku", "product_title", "brand", "cat1", "cat2", metric,
                "price", "benchmark_price", "stock_qty", "revenue",
                "pdp_views", "transactions", "c2d_pct", "b2d_pct"
            ] if c in result_df.columns]
        ].to_dict(orient="records")

        return json.dumps(
            {
                "analysis_type": "business_metric_calculation",
                "question": question,
                "calculation_type": "top_n",
                "metric": metric,
                "aggregation": "top_10",
                "filters": filters,
                "row_count": int(len(filtered)),
                "result": rows,
                "rows": rows,
            },
            ensure_ascii=False,
            default=str,
        )

    if any(x in q for x in ["en dusuk", "en düşük", "bottom 10", "son 10", "en az"]):
        result_df = filtered.sort_values(metric, ascending=True).head(10)
        rows = result_df[
            [c for c in [
                "sku", "product_title", "brand", "cat1", "cat2", metric,
                "price", "benchmark_price", "stock_qty", "revenue",
                "pdp_views", "transactions", "c2d_pct", "b2d_pct"
            ] if c in result_df.columns]
        ].to_dict(orient="records")

        return json.dumps(
            {
                "analysis_type": "business_metric_calculation",
                "question": question,
                "calculation_type": "bottom_n",
                "metric": metric,
                "aggregation": "bottom_10",
                "filters": filters,
                "row_count": int(len(filtered)),
                "result": rows,
                "rows": rows,
            },
            ensure_ascii=False,
            default=str,
        )

    # Group by
    if group_by and group_by in filtered.columns:
        if aggregation == "count":
            grouped = (
                filtered.groupby(group_by)["sku"]
                .nunique()
                .reset_index(name="value")
                .sort_values("value", ascending=False)
            )
        else:
            agg_map = {
                "avg": "mean",
                "sum": "sum",
                "max": "max",
                "min": "min",
                "median": "median",
            }
            pandas_agg = agg_map.get(aggregation, "mean")

            grouped = (
                filtered.groupby(group_by)[metric]
                .agg(pandas_agg)
                .reset_index(name="value")
                .sort_values("value", ascending=False)
            )

        grouped["value"] = grouped["value"].apply(format_number)
        rows = grouped.to_dict(orient="records")

        return json.dumps(
            {
                "analysis_type": "business_metric_calculation",
                "question": question,
                "calculation_type": "group_by",
                "metric": metric,
                "aggregation": aggregation,
                "group_by": group_by,
                "filters": filters,
                "row_count": int(len(filtered)),
                "result": rows,
                "rows": rows,
            },
            ensure_ascii=False,
            default=str,
        )

    # Count özel
    if aggregation == "count":
        if "sku" in filtered.columns:
            value = filtered["sku"].nunique()
            metric_used = "sku_unique_count"
        else:
            value = len(filtered)
            metric_used = "row_count"

        rows = [{
            "metric": metric_used,
            "value": int(value),
            "filtered_row_count": int(len(filtered)),
        }]

        return json.dumps(
            {
                "analysis_type": "business_metric_calculation",
                "question": question,
                "calculation_type": "scalar",
                "metric": metric_used,
                "aggregation": aggregation,
                "filters": filters,
                "row_count": int(len(filtered)),
                "result": int(value),
                "rows": rows,
            },
            ensure_ascii=False,
            default=str,
        )

    # Scalar hesap
    value = calculate_scalar(filtered[metric], aggregation)
    value = format_number(value)

    rows = [{
        "metric": metric,
        "aggregation": aggregation,
        "value": value,
        "filtered_row_count": int(len(filtered)),
    }]

    return json.dumps(
        {
            "analysis_type": "business_metric_calculation",
            "question": question,
            "calculation_type": "scalar",
            "metric": metric,
            "aggregation": aggregation,
            "filters": filters,
            "row_count": int(len(filtered)),
            "result": value,
            "rows": rows,
            "calculation_detail": {
                "dataset": "company_product_input.xlsx + merchant_price_benchmark_sample.xlsx",
                "metric": metric,
                "aggregation": aggregation,
                "filters": filters,
                "included_rows": int(len(filtered)),
            },
        },
        ensure_ascii=False,
        default=str,
    )