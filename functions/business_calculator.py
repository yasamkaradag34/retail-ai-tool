import json
import os
import re
from functools import lru_cache
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests


COMPANY_INPUT_PATH = "data/company_product_input.xlsx"
MERCHANT_BENCHMARK_SAMPLE_PATH = "data/merchant_price_benchmark_sample.xlsx"

OLLAMA_URL = "http://localhost:11434/api/chat"
MODEL = "llama3.1"


COLUMN_ALIASES = {
    "gtin": ["gtin", "ean", "barcode", "barcodeno", "product_gtin"],
    "sku": ["sku", "item_id", "offer_id", "product_id", "id"],
    "product_title": ["product_title", "title", "name", "product_name", "product"],
    "brand": ["brand", "manufacturer", "marka"],
    "cat1": ["cat1", "category1", "category", "category_l1", "main_category"],
    "cat2": ["cat2", "category2", "subcategory", "category_l2", "sub_category"],
    "sales_channel": ["sales_channel", "channel", "platform"],
    "main_traffic_channel": ["main_traffic_channel", "maintrafficchannel", "traffic_channel"],
    "ref_channel": ["ref_channel", "refchannel", "source_medium", "ref_channel_fy26"],

    "price": ["price", "your_price", "sale_price", "product_price", "productprice"],
    "benchmark_price": ["benchmark_price", "merchant_benchmark_price", "market_price"],
    "stock_qty": ["stock_qty", "stock", "inventory"],
    "reorder_point_qty": ["reorder_point_qty", "reorder_point", "critical_stock"],
    "availability_status": ["availability_status", "availability", "stock_status"],
    "pdp_views": ["pdp_views", "pdp", "pdp_view", "total_unique_pdp_views_sum"],
    "list_clicks": ["list_clicks", "listclicks", "list_click"],
    "add_to_carts": ["add_to_carts", "a2c", "total_unique_add_to_carts_sum"],
    "transactions": ["transactions", "trans", "orders", "order_count", "total_transactions_sum"],
    "revenue": ["revenue", "gmv", "sales_amount", "ciro"],
    "c2d_pct": ["c2d_pct", "c2d"],
    "b2d_pct": ["b2d_pct", "b2d"],
    "bounce_rate_pct": ["bounce_rate_pct", "br", "bounce_rate"],
    "stock_coverage_days": ["stock_coverage_days", "stock_coverage"],
    "estimated_lost_revenue": ["estimated_lost_revenue", "lost_revenue"],
}

DIMENSION_COLUMNS = [
    "gtin", "sku", "product_title", "brand", "cat1", "cat2",
    "sales_channel", "main_traffic_channel", "ref_channel", "availability_status"
]

NUMERIC_COLUMNS = [
    "price", "benchmark_price", "price_gap", "price_gap_pct",
    "stock_qty", "reorder_point_qty", "pdp_views", "list_clicks",
    "add_to_carts", "transactions", "revenue", "c2d_pct", "b2d_pct",
    "bounce_rate_pct", "stock_coverage_days", "estimated_lost_revenue", "aov"
]

ALLOWED_OPERATIONS = {
    "scalar",              # tek metrik hesabı: avg/sum/count/min/max/median
    "group_by",            # marka/kategori/kanal kırılımı
    "top_n",               # en yüksek ilk N
    "bottom_n",            # en düşük ilk N
    "share_of_total",      # bir filtrenin toplam içindeki payı
    "ratio",               # metric_a / metric_b
    "comparison",          # iki veya daha fazla segment karşılaştırması
    "filter_list",         # filtrelenmiş ürün listesi
}

ALLOWED_AGGREGATIONS = {
    "avg", "sum", "count", "unique_count", "min", "max", "median"
}


# -----------------------------
# Genel yardımcılar
# -----------------------------

def normalize_text(value: Any) -> str:
    value = str(value or "").strip()
    tr_map = str.maketrans("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC")
    value = value.translate(tr_map).lower()
    value = re.sub(r"\s+", " ", value)
    return value


def normalize_col_name(value: Any) -> str:
    value = str(value or "").strip().lower()
    value = value.replace("%", "pct")
    value = value.replace(" ", "_").replace("-", "_").replace(".", "_").replace("/", "_")
    value = value.replace("(", "").replace(")", "")
    value = re.sub(r"_+", "_", value)
    return value.strip("_")


def clean_json_text(text: str) -> str:
    text = (text or "").strip()
    text = text.replace("```json", "").replace("```", "").strip()
    match = re.search(r"(\{.*\})", text, flags=re.DOTALL)
    if match:
        text = match.group(1)
    return text


def safe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    try:
        return float(value)
    except Exception:
        return None


def format_number(value: Any, decimals: int = 2) -> Any:
    value = safe_float(value)
    if value is None:
        return None
    return round(value, decimals)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    raw = df.copy()
    raw.columns = [normalize_col_name(c) for c in raw.columns]

    normalized = pd.DataFrame()

    for target, aliases in COLUMN_ALIASES.items():
        found = None
        aliases_norm = [normalize_col_name(a) for a in aliases]
        for alias in aliases_norm:
            if alias in raw.columns:
                found = alias
                break
        normalized[target] = raw[found] if found else None

    return normalized


@lru_cache(maxsize=1)
def load_business_data() -> pd.DataFrame:
    if not os.path.exists(COMPANY_INPUT_PATH):
        raise FileNotFoundError(f"Şirket input dosyası bulunamadı: {COMPANY_INPUT_PATH}")

    company_df = pd.read_excel(COMPANY_INPUT_PATH)
    company_df = normalize_columns(company_df)

    # Merchant benchmark varsa GTIN üzerinden zenginleştir.
    if os.path.exists(MERCHANT_BENCHMARK_SAMPLE_PATH):
        benchmark_df = pd.read_excel(MERCHANT_BENCHMARK_SAMPLE_PATH)
        benchmark_df = normalize_columns(benchmark_df)

        if "gtin" in company_df.columns and "gtin" in benchmark_df.columns:
            company_df["gtin"] = company_df["gtin"].astype(str).str.strip()
            benchmark_df["gtin"] = benchmark_df["gtin"].astype(str).str.strip()
            benchmark_cols = ["gtin", "benchmark_price"]
            benchmark_cols = [c for c in benchmark_cols if c in benchmark_df.columns]

            if "benchmark_price" in benchmark_cols:
                company_df = company_df.drop(columns=["benchmark_price"], errors="ignore")
                company_df = company_df.merge(benchmark_df[benchmark_cols], on="gtin", how="left")

    for col in NUMERIC_COLUMNS:
        if col in company_df.columns:
            company_df[col] = pd.to_numeric(company_df[col], errors="coerce")

    # Derived metrics: eksikse hesapla veya güncelle.
    if "add_to_carts" in company_df.columns and "pdp_views" in company_df.columns:
        if "c2d_pct" not in company_df.columns or company_df["c2d_pct"].isna().all():
            company_df["c2d_pct"] = (company_df["add_to_carts"] / company_df["pdp_views"]) * 100

    if "transactions" in company_df.columns and "pdp_views" in company_df.columns:
        if "b2d_pct" not in company_df.columns or company_df["b2d_pct"].isna().all():
            company_df["b2d_pct"] = (company_df["transactions"] / company_df["pdp_views"]) * 100

    if "revenue" in company_df.columns and "transactions" in company_df.columns:
        company_df["aov"] = company_df["revenue"] / company_df["transactions"].replace(0, pd.NA)

    if "price" in company_df.columns and "benchmark_price" in company_df.columns:
        company_df["price_gap"] = company_df["price"] - company_df["benchmark_price"]
        company_df["price_gap_pct"] = (company_df["price_gap"] / company_df["benchmark_price"]) * 100

    return company_df


def get_schema_summary(df: pd.DataFrame) -> Dict[str, Any]:
    dimensions = [c for c in DIMENSION_COLUMNS if c in df.columns]
    metrics = [c for c in df.columns if c not in dimensions and pd.api.types.is_numeric_dtype(df[c])]

    samples = {}
    for col in dimensions:
        values = df[col].dropna().astype(str).unique().tolist()[:20]
        samples[col] = values

    return {
        "columns": list(df.columns),
        "dimensions": dimensions,
        "metrics": metrics,
        "dimension_samples": samples,
        "row_count": int(len(df)),
    }


# -----------------------------
# LLM: soru -> hesaplama spec
# -----------------------------

def parse_business_question_with_llm(question: str, df: pd.DataFrame) -> Dict[str, Any]:
    schema = get_schema_summary(df)

    prompt = f"""
Sen bir retail/e-commerce business calculation planner'sın.
Görevin kullanıcının Türkçe sorusunu SADECE JSON hesaplama planına çevirmek.
Hesabı sen yapma. Sadece JSON döndür.

Tablo şeması:
{json.dumps(schema, ensure_ascii=False, indent=2, default=str)}

Desteklenen operation değerleri:
- scalar: tek bir metrik için avg/sum/count/unique_count/min/max/median
- group_by: bir kırılıma göre metrik hesaplama
- top_n: metriğe göre en yüksek N satır veya grup
- bottom_n: metriğe göre en düşük N satır veya grup
- share_of_total: filtreli metrik toplamının genel toplam içindeki yüzdesi
- ratio: iki metriğin oranı
- comparison: iki veya daha fazla segmenti karşılaştırma
- filter_list: filtreye uyan satırları listeleme

Desteklenen aggregation değerleri:
avg, sum, count, unique_count, min, max, median

Filtre formatı:
{{"column":"brand", "operator":"equals", "value":"APPLE"}}
operator değerleri:
equals, contains, in, gt, gte, lt, lte, not_null, is_null,
above_benchmark, below_benchmark, stock_risk, oos

Önemli iş kuralları:
- APPLE, SAMSUNG, XIAOMI gibi değerler genelde brand filtresidir.
- GSM, mobile, mobil, telefon, cep telefonu sorulursa cat1/cat2/product_title içinde Telefon/Cep Telefonları/iPhone/Galaxy gibi değerleri yakalamak için contains filtresi üret.
- "ortalama fiyat" => metric price, aggregation avg.
- "toplam revenue/ciro" => metric revenue, aggregation sum.
- "payı", "toplam içindeki oranı", "% kaçı" => operation share_of_total.
- "en yüksek", "ilk 10", "top 10" => operation top_n.
- "en düşük", "bottom 10", "son 10" => operation bottom_n.
- "marka bazında" => group_by brand.
- "kategori bazında" => group_by cat1 veya cat2.
- "C2D" => metric c2d_pct.
- "B2D" => metric b2d_pct.
- "benchmark üstünde/pahalı" => filter operator above_benchmark.
- "benchmark altında/ucuz" => filter operator below_benchmark.
- "stok riski" => filter operator stock_risk.
- "OOS/out of stock" => filter operator oos.

JSON şeması:
{{
  "operation": "scalar | group_by | top_n | bottom_n | share_of_total | ratio | comparison | filter_list",
  "metric": "price",
  "aggregation": "avg",
  "filters": [{{"column":"brand", "operator":"equals", "value":"APPLE"}}],
  "group_by": null,
  "sort": {{"by":"price", "direction":"desc"}},
  "limit": 10,
  "numerator_metric": null,
  "denominator_metric": null,
  "comparison_groups": [],
  "formula": null,
  "format": "number | percentage | currency",
  "explanation": "kısa plan açıklaması"
}}

Kullanıcı sorusu:
{question}

Sadece JSON döndür. Markdown yok. Açıklama yok.
"""

    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": prompt},
            {"role": "user", "content": question},
        ],
        "stream": False,
    }

    response = requests.post(OLLAMA_URL, json=payload, timeout=120)
    response.raise_for_status()
    content = response.json().get("message", {}).get("content", "")
    content = clean_json_text(content)
    return json.loads(content)


# -----------------------------
# Fallback parser: LLM hata alırsa
# -----------------------------

def fallback_parse_business_question(question: str, df: pd.DataFrame) -> Dict[str, Any]:
    q = normalize_text(question)

    operation = "scalar"
    aggregation = "avg"
    metric = "price"
    group_by = None
    limit = 10
    filters: List[Dict[str, Any]] = []

    if any(x in q for x in ["top 10", "ilk 10", "en yuksek", "en yüksek", "en fazla"]):
        operation = "top_n"
        aggregation = "sum"
    elif any(x in q for x in ["bottom 10", "son 10", "en dusuk", "en düşük", "en az"]):
        operation = "bottom_n"
        aggregation = "sum"
    elif any(x in q for x in ["pay", "payi", "payı", "% kaci", "% kaçı", "toplam icindeki", "toplam içindeki"]):
        operation = "share_of_total"
        aggregation = "sum"
    elif any(x in q for x in ["marka bazinda", "marka bazında", "brand bazinda", "brand bazında"]):
        operation = "group_by"
        group_by = "brand"
    elif any(x in q for x in ["kategori bazinda", "kategori bazında", "cat1"]):
        operation = "group_by"
        group_by = "cat1"
    elif any(x in q for x in ["alt kategori", "cat2"]):
        operation = "group_by"
        group_by = "cat2"

    if any(x in q for x in ["toplam", "sum", "ciro toplam"]):
        aggregation = "sum"
    elif any(x in q for x in ["kac", "kaç", "adet", "sayisi", "sayısı", "count"]):
        aggregation = "count"
    elif any(x in q for x in ["min", "minimum", "en dusuk", "en düşük"]):
        aggregation = "min"
    elif any(x in q for x in ["max", "maksimum", "en yuksek", "en yüksek"]):
        aggregation = "max"
    elif any(x in q for x in ["median", "medyan"]):
        aggregation = "median"

    metric_rules = [
        ("price_gap_pct", ["price gap", "fiyat fark", "gap", "benchmark fark"]),
        ("benchmark_price", ["benchmark fiyat", "merchant fiyat", "rakip fiyat"]),
        ("revenue", ["revenue", "ciro", "gmv", "sales amount"]),
        ("transactions", ["transaction", "transactions", "trans", "siparis", "sipariş", "satis adedi", "satış adedi"]),
        ("pdp_views", ["pdp", "goruntulenme", "görüntülenme"]),
        ("add_to_carts", ["a2c", "add to cart", "sepete"]),
        ("c2d_pct", ["c2d"]),
        ("b2d_pct", ["b2d"]),
        ("stock_qty", ["stok", "stock", "envanter"]),
        ("price", ["fiyat", "price", "ortalama fiyat"]),
    ]
    for m, keywords in metric_rules:
        if any(k in q for k in keywords):
            metric = m
            break

    # Brand filter
    if "brand" in df.columns:
        brands = [x for x in df["brand"].dropna().astype(str).unique().tolist() if x]
        for brand in sorted(brands, key=len, reverse=True):
            if normalize_text(brand) in q:
                filters.append({"column": "brand", "operator": "equals", "value": brand})
                break

    # Common category synonyms
    if any(x in q for x in ["gsm", "mobile", "mobil", "telefon", "cep telefonu"]):
        filters.append({"column": "cat2", "operator": "contains", "value": "Cep Telefonları"})
    elif any(x in q for x in ["tablet", "ipad"]):
        filters.append({"column": "cat2", "operator": "contains", "value": "Tablet"})
    elif any(x in q for x in ["kulaklik", "kulaklık", "headphone", "airpods"]):
        filters.append({"column": "cat2", "operator": "contains", "value": "Kulaklık"})

    if any(x in q for x in ["benchmark ustu", "benchmark üstü", "pahali", "pahalı", "rakibe gore pahali", "rakibe göre pahalı"]):
        filters.append({"column": "price_gap_pct", "operator": "above_benchmark", "value": 0})
    if any(x in q for x in ["benchmark alti", "benchmark altı", "ucuz", "rakibe gore ucuz", "rakibe göre ucuz"]):
        filters.append({"column": "price_gap_pct", "operator": "below_benchmark", "value": 0})
    if "stok riski" in q:
        filters.append({"column": "stock_qty", "operator": "stock_risk", "value": None})

    return {
        "operation": operation,
        "metric": metric,
        "aggregation": aggregation,
        "filters": filters,
        "group_by": group_by,
        "sort": {"by": metric, "direction": "desc"},
        "limit": limit,
        "numerator_metric": None,
        "denominator_metric": None,
        "comparison_groups": [],
        "formula": None,
        "format": "percentage" if metric.endswith("pct") or operation == "share_of_total" else "number",
        "explanation": "Fallback parser ile oluşturuldu."
    }


# -----------------------------
# Spec validate / normalize
# -----------------------------

def map_metric(metric: Any, df: pd.DataFrame) -> Optional[str]:
    if metric is None:
        return None
    metric_norm = normalize_col_name(metric)
    if metric_norm in df.columns:
        return metric_norm

    # Alias üzerinden de dene
    for target, aliases in COLUMN_ALIASES.items():
        if metric_norm == normalize_col_name(target) or metric_norm in [normalize_col_name(a) for a in aliases]:
            if target in df.columns:
                return target
    return metric_norm if metric_norm in df.columns else None


def normalize_filter(f: Dict[str, Any], df: pd.DataFrame) -> Dict[str, Any]:
    f = dict(f or {})
    col = f.get("column")
    operator = f.get("operator", "contains")

    mapped_col = map_metric(col, df) if col else None
    f["column"] = mapped_col or col
    f["operator"] = normalize_text(operator).replace(" ", "_")
    return f


def validate_calculation_spec(spec: Dict[str, Any], df: pd.DataFrame) -> Dict[str, Any]:
    spec = dict(spec or {})

    operation = normalize_text(spec.get("operation", "scalar"))
    if operation not in ALLOWED_OPERATIONS:
        operation = "scalar"
    spec["operation"] = operation

    aggregation = normalize_text(spec.get("aggregation", "avg"))
    if aggregation not in ALLOWED_AGGREGATIONS:
        aggregation = "avg"
    spec["aggregation"] = aggregation

    metric = map_metric(spec.get("metric") or "price", df) or "price"
    if metric not in df.columns:
        metric = "price" if "price" in df.columns else next((c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])), None)
    spec["metric"] = metric

    group_by = spec.get("group_by")
    group_by = map_metric(group_by, df) if group_by else None
    spec["group_by"] = group_by

    filters = spec.get("filters") or []
    if not isinstance(filters, list):
        filters = []
    spec["filters"] = [normalize_filter(f, df) for f in filters if isinstance(f, dict)]

    sort = spec.get("sort") or {}
    if not isinstance(sort, dict):
        sort = {}
    sort_by = map_metric(sort.get("by") or metric, df) or metric
    direction = normalize_text(sort.get("direction", "desc"))
    if direction not in ["asc", "desc"]:
        direction = "desc"
    spec["sort"] = {"by": sort_by, "direction": direction}

    try:
        limit = int(spec.get("limit") or 10)
    except Exception:
        limit = 10
    spec["limit"] = max(1, min(limit, 100))

    spec["numerator_metric"] = map_metric(spec.get("numerator_metric"), df)
    spec["denominator_metric"] = map_metric(spec.get("denominator_metric"), df)

    comparison_groups = spec.get("comparison_groups") or []
    if not isinstance(comparison_groups, list):
        comparison_groups = []
    clean_groups = []
    for group in comparison_groups:
        if isinstance(group, dict):
            group_filters = group.get("filters") or []
            clean_groups.append({
                "name": group.get("name", "Segment"),
                "filters": [normalize_filter(f, df) for f in group_filters if isinstance(f, dict)]
            })
    spec["comparison_groups"] = clean_groups

    return spec


# -----------------------------
# Filter execution
# -----------------------------

def apply_single_filter(df: pd.DataFrame, f: Dict[str, Any]) -> pd.DataFrame:
    col = f.get("column")
    op = f.get("operator", "contains")
    value = f.get("value")

    filtered = df.copy()

    if op == "above_benchmark":
        if "price_gap_pct" in filtered.columns:
            return filtered[filtered["price_gap_pct"] > 0].copy()
        if "price" in filtered.columns and "benchmark_price" in filtered.columns:
            return filtered[filtered["price"] > filtered["benchmark_price"]].copy()
        return filtered.iloc[0:0].copy()

    if op == "below_benchmark":
        if "price_gap_pct" in filtered.columns:
            return filtered[filtered["price_gap_pct"] < 0].copy()
        if "price" in filtered.columns and "benchmark_price" in filtered.columns:
            return filtered[filtered["price"] < filtered["benchmark_price"]].copy()
        return filtered.iloc[0:0].copy()

    if op == "stock_risk":
        if "stock_qty" not in filtered.columns:
            return filtered.iloc[0:0].copy()
        if "reorder_point_qty" in filtered.columns:
            return filtered[filtered["stock_qty"] <= filtered["reorder_point_qty"]].copy()
        threshold = filtered["stock_qty"].quantile(0.25)
        return filtered[filtered["stock_qty"] <= threshold].copy()

    if op == "oos":
        mask = pd.Series(False, index=filtered.index)
        if "stock_qty" in filtered.columns:
            mask = mask | (filtered["stock_qty"].fillna(0) == 0)
        if "availability_status" in filtered.columns:
            mask = mask | filtered["availability_status"].astype(str).apply(normalize_text).str.contains("out|oos|stok yok", na=False, regex=True)
        return filtered[mask].copy()

    if not col or col not in filtered.columns:
        return filtered

    series = filtered[col]

    if op == "equals":
        return filtered[series.astype(str).apply(normalize_text) == normalize_text(value)].copy()
    if op == "contains":
        return filtered[series.astype(str).apply(normalize_text).str.contains(normalize_text(value), na=False, regex=False)].copy()
    if op == "in":
        values = value if isinstance(value, list) else [value]
        values_norm = [normalize_text(v) for v in values]
        return filtered[series.astype(str).apply(normalize_text).isin(values_norm)].copy()
    if op in ["gt", "gte", "lt", "lte"]:
        comp_value = safe_float(value)
        numeric = pd.to_numeric(series, errors="coerce")
        if comp_value is None:
            return filtered
        if op == "gt":
            return filtered[numeric > comp_value].copy()
        if op == "gte":
            return filtered[numeric >= comp_value].copy()
        if op == "lt":
            return filtered[numeric < comp_value].copy()
        if op == "lte":
            return filtered[numeric <= comp_value].copy()
    if op == "not_null":
        return filtered[series.notna()].copy()
    if op == "is_null":
        return filtered[series.isna()].copy()

    return filtered


def apply_filters(df: pd.DataFrame, filters: List[Dict[str, Any]]) -> pd.DataFrame:
    filtered = df.copy()
    for f in filters or []:
        filtered = apply_single_filter(filtered, f)
    return filtered


# -----------------------------
# Calculation execution
# -----------------------------

def aggregate_series(series: pd.Series, aggregation: str) -> Any:
    if aggregation == "count":
        return int(series.count())
    if aggregation == "unique_count":
        return int(series.nunique())

    numeric = pd.to_numeric(series, errors="coerce").dropna()
    if numeric.empty:
        return None

    if aggregation == "sum":
        return numeric.sum()
    if aggregation == "avg":
        return numeric.mean()
    if aggregation == "min":
        return numeric.min()
    if aggregation == "max":
        return numeric.max()
    if aggregation == "median":
        return numeric.median()
    return numeric.mean()


def rows_for_product_list(df: pd.DataFrame, limit: int) -> List[Dict[str, Any]]:
    cols = [
        "sku", "gtin", "product_title", "brand", "cat1", "cat2",
        "price", "benchmark_price", "price_gap_pct", "stock_qty",
        "pdp_views", "add_to_carts", "transactions", "revenue",
        "c2d_pct", "b2d_pct", "aov"
    ]
    existing = [c for c in cols if c in df.columns]
    clean = df[existing].head(limit).copy()
    clean = clean.where(pd.notnull(clean), None)
    return clean.to_dict(orient="records")


def execute_scalar(df: pd.DataFrame, spec: Dict[str, Any], question: str) -> Dict[str, Any]:
    metric = spec["metric"]
    aggregation = spec["aggregation"]
    filtered = apply_filters(df, spec.get("filters", []))

    value = aggregate_series(filtered[metric], aggregation) if metric in filtered.columns else None
    value = format_number(value)

    rows = [{
        "metric": metric,
        "aggregation": aggregation,
        "value": value,
        "filtered_row_count": int(len(filtered)),
    }]

    return result_payload(question, spec, "scalar", value, rows, len(filtered))


def execute_group_by(df: pd.DataFrame, spec: Dict[str, Any], question: str) -> Dict[str, Any]:
    metric = spec["metric"]
    aggregation = spec["aggregation"]
    group_by = spec.get("group_by") or "brand"
    limit = spec.get("limit", 20)
    filtered = apply_filters(df, spec.get("filters", []))

    if group_by not in filtered.columns:
        return error_payload(question, f"Group by kolonu bulunamadı: {group_by}", spec)

    if aggregation == "count":
        grouped = filtered.groupby(group_by).size().reset_index(name="value")
    elif aggregation == "unique_count":
        unique_col = "sku" if "sku" in filtered.columns else metric
        grouped = filtered.groupby(group_by)[unique_col].nunique().reset_index(name="value")
    else:
        agg_map = {"avg": "mean", "sum": "sum", "min": "min", "max": "max", "median": "median"}
        grouped = filtered.groupby(group_by)[metric].agg(agg_map.get(aggregation, "mean")).reset_index(name="value")

    grouped = grouped.sort_values("value", ascending=False).head(limit)
    grouped["value"] = grouped["value"].apply(format_number)
    rows = grouped.to_dict(orient="records")

    return result_payload(question, spec, "group_by", rows, rows, len(filtered))


def execute_top_bottom(df: pd.DataFrame, spec: Dict[str, Any], question: str, ascending: bool) -> Dict[str, Any]:
    metric = spec["metric"]
    group_by = spec.get("group_by")
    limit = spec.get("limit", 10)
    filtered = apply_filters(df, spec.get("filters", []))

    if group_by and group_by in filtered.columns:
        aggregation = spec.get("aggregation", "sum")
        if aggregation in ["count", "unique_count"]:
            grouped = filtered.groupby(group_by).size().reset_index(name=metric)
        else:
            agg_map = {"avg": "mean", "sum": "sum", "min": "min", "max": "max", "median": "median"}
            grouped = filtered.groupby(group_by)[metric].agg(agg_map.get(aggregation, "sum")).reset_index()
        result_df = grouped.sort_values(metric, ascending=ascending).head(limit)
        result_df[metric] = result_df[metric].apply(format_number)
        rows = result_df.to_dict(orient="records")
    else:
        result_df = filtered.sort_values(metric, ascending=ascending).head(limit)
        rows = rows_for_product_list(result_df, limit)

    calc_type = "bottom_n" if ascending else "top_n"
    return result_payload(question, spec, calc_type, rows, rows, len(filtered))


def execute_share_of_total(df: pd.DataFrame, spec: Dict[str, Any], question: str) -> Dict[str, Any]:
    metric = spec["metric"]
    filtered = apply_filters(df, spec.get("filters", []))

    numerator = pd.to_numeric(filtered[metric], errors="coerce").sum()
    denominator = pd.to_numeric(df[metric], errors="coerce").sum()
    share = (numerator / denominator * 100) if denominator else None

    rows = [{
        "metric": metric,
        "numerator": format_number(numerator),
        "denominator": format_number(denominator),
        "share_pct": format_number(share),
        "filtered_row_count": int(len(filtered)),
        "total_row_count": int(len(df)),
    }]

    return result_payload(question, spec, "share_of_total", format_number(share), rows, len(filtered))


def execute_ratio(df: pd.DataFrame, spec: Dict[str, Any], question: str) -> Dict[str, Any]:
    filtered = apply_filters(df, spec.get("filters", []))
    numerator_metric = spec.get("numerator_metric") or "transactions"
    denominator_metric = spec.get("denominator_metric") or "pdp_views"

    if numerator_metric not in filtered.columns or denominator_metric not in filtered.columns:
        return error_payload(question, "Ratio için numerator/denominator metrikleri bulunamadı.", spec)

    numerator = pd.to_numeric(filtered[numerator_metric], errors="coerce").sum()
    denominator = pd.to_numeric(filtered[denominator_metric], errors="coerce").sum()
    ratio_pct = (numerator / denominator * 100) if denominator else None

    rows = [{
        "numerator_metric": numerator_metric,
        "denominator_metric": denominator_metric,
        "numerator": format_number(numerator),
        "denominator": format_number(denominator),
        "ratio_pct": format_number(ratio_pct),
        "filtered_row_count": int(len(filtered)),
    }]

    return result_payload(question, spec, "ratio", format_number(ratio_pct), rows, len(filtered))


def execute_comparison(df: pd.DataFrame, spec: Dict[str, Any], question: str) -> Dict[str, Any]:
    metric = spec["metric"]
    aggregation = spec["aggregation"]
    rows = []

    groups = spec.get("comparison_groups") or []
    if not groups:
        return execute_group_by(df, {**spec, "group_by": spec.get("group_by") or "brand"}, question)

    for group in groups:
        name = group.get("name", "Segment")
        filtered = apply_filters(df, group.get("filters", []))
        value = aggregate_series(filtered[metric], aggregation) if metric in filtered.columns else None
        rows.append({
            "segment": name,
            "metric": metric,
            "aggregation": aggregation,
            "value": format_number(value),
            "row_count": int(len(filtered)),
        })

    return result_payload(question, spec, "comparison", rows, rows, sum(r.get("row_count", 0) for r in rows))


def execute_filter_list(df: pd.DataFrame, spec: Dict[str, Any], question: str) -> Dict[str, Any]:
    filtered = apply_filters(df, spec.get("filters", []))
    sort = spec.get("sort") or {}
    sort_by = sort.get("by") or spec.get("metric")
    ascending = sort.get("direction") == "asc"
    limit = spec.get("limit", 20)

    if sort_by in filtered.columns:
        filtered = filtered.sort_values(sort_by, ascending=ascending)

    rows = rows_for_product_list(filtered, limit)
    return result_payload(question, spec, "filter_list", rows, rows, len(filtered))


def result_payload(question: str, spec: Dict[str, Any], calculation_type: str, result: Any, rows: List[Dict[str, Any]], row_count: int) -> Dict[str, Any]:
    return {
        "analysis_type": "business_metric_calculation",
        "question": question,
        "calculation_type": calculation_type,
        "metric": spec.get("metric"),
        "aggregation": spec.get("aggregation"),
        "group_by": spec.get("group_by"),
        "filters": spec.get("filters", []),
        "operation": spec.get("operation"),
        "format": spec.get("format", "number"),
        "row_count": int(row_count),
        "result": result,
        "rows": rows,
        "calculation_detail": {
            "dataset": "company_product_input.xlsx + merchant_price_benchmark_sample.xlsx",
            "spec": spec,
            "formula": spec.get("formula"),
            "explanation": spec.get("explanation"),
        },
    }


def error_payload(question: str, error: str, spec: Optional[Dict[str, Any]] = None, detail: Optional[str] = None) -> Dict[str, Any]:
    return {
        "analysis_type": "business_metric_error",
        "question": question,
        "error": error,
        "detail": detail,
        "spec": spec or {},
    }


def execute_calculation_spec(df: pd.DataFrame, spec: Dict[str, Any], question: str) -> Dict[str, Any]:
    operation = spec.get("operation", "scalar")

    if operation == "scalar":
        return execute_scalar(df, spec, question)
    if operation == "group_by":
        return execute_group_by(df, spec, question)
    if operation == "top_n":
        return execute_top_bottom(df, spec, question, ascending=False)
    if operation == "bottom_n":
        return execute_top_bottom(df, spec, question, ascending=True)
    if operation == "share_of_total":
        return execute_share_of_total(df, spec, question)
    if operation == "ratio":
        return execute_ratio(df, spec, question)
    if operation == "comparison":
        return execute_comparison(df, spec, question)
    if operation == "filter_list":
        return execute_filter_list(df, spec, question)

    return execute_scalar(df, spec, question)


# -----------------------------
# Public function
# -----------------------------

def calculate_business_metric(question: str) -> str:
    try:
        df = load_business_data()
    except Exception as e:
        return json.dumps(
            error_payload(question, "Business calculator datası okunamadı.", detail=str(e)),
            ensure_ascii=False,
            default=str,
        )

    try:
        spec = parse_business_question_with_llm(question, df)
        spec_source = "llm_json_planner"
    except Exception as e:
        print("⚠️ Business Calculator LLM planner hata aldı, fallback parser kullanılacak:", str(e), flush=True)
        spec = fallback_parse_business_question(question, df)
        spec_source = "fallback_parser"

    try:
        spec = validate_calculation_spec(spec, df)
        result = execute_calculation_spec(df, spec, question)
        result["calculation_detail"]["spec_source"] = spec_source
        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as e:
        return json.dumps(
            error_payload(
                question,
                "Business calculation çalıştırılırken hata oluştu.",
                spec=spec,
                detail=str(e),
            ),
            ensure_ascii=False,
            default=str,
        )
