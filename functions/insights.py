import json
import os
import re
from functools import lru_cache
from typing import Dict, List, Optional, Tuple

import pandas as pd

try:
    from functions.sector_profiles import SECTOR_PROFILES
except Exception:
    SECTOR_PROFILES = {
        "consumer_electronics": {"name": "Consumer Electronics"},
        "fashion": {"name": "Fashion"},
        "fmcg": {"name": "FMCG"},
        "marketplace_general": {"name": "Marketplace General"},
    }

try:
    from functions.sector_norms import get_behavior_for_category
except Exception as e:
    print("Warning: could not import sector_norms:", e)
    get_behavior_for_category = lambda x: None


COMPANY_INPUT_PATH = "data/company_product_input.xlsx"
MERCHANT_BENCHMARK_SAMPLE_PATH = "data/merchant_price_benchmark_sample.xlsx"
FALLBACK_SAMPLE_PATH = "data/ecommerce_ai_sample_data_200_rows.xlsx"
FALLBACK_SAMPLE_SHEET = "sample_data_200"


CATEGORY_SYNONYMS = {
    "mobile": ["gsm", "telefon", "cep", "cep telefonlari", "cep telefonları", "iphone", "galaxy"],
    "mobil": ["gsm", "telefon", "cep", "cep telefonlari", "cep telefonları", "iphone", "galaxy"],
    "gsm": ["gsm", "telefon", "cep", "cep telefonlari", "cep telefonları", "iphone", "galaxy"],
    "telefon": ["gsm", "telefon", "cep", "cep telefonlari", "cep telefonları", "iphone", "galaxy"],
    "cep telefonlari": ["gsm", "telefon", "cep", "cep telefonlari", "cep telefonları", "iphone", "galaxy"],
    "cep telefonları": ["gsm", "telefon", "cep", "cep telefonlari", "cep telefonları", "iphone", "galaxy"],
    "tablet": ["tablet", "tabletler", "ipad"],
    "tabletler": ["tablet", "tabletler", "ipad"],
    "kulaklik": ["kulaklik", "kulaklık", "airpods", "headphone", "headphones", "jbl"],
    "kulaklık": ["kulaklik", "kulaklık", "airpods", "headphone", "headphones", "jbl"],
    "giyilebilir": ["giyilebilir", "watch", "saat", "akilli saat", "akıllı saat"],
    "giyilebilir teknoloji": ["giyilebilir", "watch", "saat", "akilli saat", "akıllı saat"],
    "aksesuar": ["aksesuar", "telefon aksesuarlari", "telefon aksesuarları", "oyuncu aksesuarlari", "oyuncu aksesuarları"],
    "tv": ["tv", "televizyon"],
    "televizyon": ["tv", "televizyon"],
    "supurge": ["supurge", "süpürge", "roborock"],
    "süpürge": ["supurge", "süpürge", "roborock"],
}

COLUMN_ALIASES = {
    "gtin": ["gtin", "ean", "barcode", "barcodeno", "product_gtin"],
    "sku": ["sku", "item_id", "offer_id", "product_id", "id"],
    "product_title": ["product_title", "title", "name", "product_name", "product"],
    "brand": ["brand", "manufacturer", "marka"],
    "cat1": ["cat1", "category1", "category", "category_l1", "main_category"],
    "cat2": ["cat2", "category2", "subcategory", "category_l2", "sub_category"],
    "sales_channel": ["sales_channel", "channel", "platform", "device"],
    "traffic_channel": ["traffic_channel", "main_traffic_channel", "maintrafficchannel", "ref_channel", "source_medium"],
    "main_traffic_channel": ["main_traffic_channel", "maintrafficchannel", "traffic_channel"],
    "ref_channel": ["ref_channel", "refchannel", "source_medium"],
    "price": ["price", "your_price", "sale_price", "product_price", "productprice"],
    "product_price_delta_pct": ["product_price_delta_pct", "price_delta_pct", "productprice_delta_pct"],
    "benchmark_price": ["benchmark_price", "merchant_benchmark_price", "market_price"],
    "stock_qty": ["stock_qty", "stock", "inventory"],
    "reorder_point_qty": ["reorder_point_qty", "reorder_point", "critical_stock"],
    "availability_status": ["availability_status", "availability", "stock_status"],
    "stock_coverage_days": ["stock_coverage_days", "stock_coverage"],
    "estimated_lost_revenue": ["estimated_lost_revenue", "lost_revenue"],
    "pdp_views": ["pdp_views", "pdp", "pdp_view", "total_unique_pdp_views_sum"],
    "pdp_delta_pct": ["pdp_delta_pct", "pdp_delta"],
    "list_clicks": ["list_clicks", "listclicks", "list_click"],
    "add_to_carts": ["add_to_carts", "a2c", "total_unique_add_to_carts_sum"],
    "a2c_delta_pct": ["a2c_delta_pct", "add_to_cart_delta_pct"],
    "cart_views": ["cart_views"],
    "shipping_views": ["shipping_views"],
    "payment_views": ["payment_views"],
    "summary_views": ["summary_views"],
    "checkout_submits": ["checkout_submits"],
    "checkout_submit_delta_pct": ["checkout_submit_delta_pct"],
    "transactions": ["transactions", "trans", "orders", "total_transactions_sum"],
    "transactions_delta_pct": ["transactions_delta_pct", "transaction_delta_pct"],
    "revenue": ["revenue", "gmv", "sales_amount", "ciro"],
    "revenue_delta_pct": ["revenue_delta_pct", "ciro_delta_pct"],
    "c2d_pct": ["c2d_pct", "c2d"],
    "c2d_delta_pct": ["c2d_delta_pct"],
    "b2d_pct": ["b2d_pct", "b2d"],
    "b2d_delta_pct": ["b2d_delta_pct"],
    "bounce_rate_pct": ["bounce_rate_pct", "br", "bounce_rate"],
    "bounce_rate_delta_pct": ["bounce_rate_delta_pct", "br_delta_pct", "br_delta"],
}

NUMERIC_COLUMNS = [
    "price", "product_price_delta_pct", "benchmark_price", "stock_qty", "reorder_point_qty",
    "stock_coverage_days", "estimated_lost_revenue", "pdp_views", "pdp_delta_pct",
    "list_clicks", "add_to_carts", "a2c_delta_pct", "cart_views", "shipping_views",
    "payment_views", "summary_views", "checkout_submits", "checkout_submit_delta_pct",
    "transactions", "transactions_delta_pct", "revenue", "revenue_delta_pct",
    "c2d_pct", "c2d_delta_pct", "b2d_pct", "b2d_delta_pct", "bounce_rate_pct",
    "bounce_rate_delta_pct", "price_gap", "price_gap_pct",
]


def normalize_text(value) -> str:
    value = str(value or "").strip().lower()
    tr_map = str.maketrans("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC")
    value = value.translate(tr_map).lower()
    value = re.sub(r"\s+", " ", value)
    return value


def normalize_column_name(col) -> str:
    return (
        str(col)
        .strip()
        .lower()
        .replace("%", "pct")
        .replace(" ", "_")
        .replace("-", "_")
        .replace(".", "_")
        .replace("/", "_")
        .replace("(", "")
        .replace(")", "")
    )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_column_name(c) for c in df.columns]

    normalized = pd.DataFrame(index=df.index)

    for target, aliases in COLUMN_ALIASES.items():
        found = None
        for alias in aliases:
            alias_norm = normalize_column_name(alias)
            if alias_norm in df.columns:
                found = alias_norm
                break
        normalized[target] = df[found] if found else None

    for col in NUMERIC_COLUMNS:
        if col in normalized.columns:
            normalized[col] = pd.to_numeric(normalized[col], errors="coerce")

    for col in ["gtin", "sku", "product_title", "brand", "cat1", "cat2", "sales_channel", "traffic_channel", "availability_status"]:
        if col in normalized.columns:
            normalized[col] = normalized[col].astype(str).replace("nan", "").str.strip()

    if "c2d_pct" in normalized.columns and normalized["c2d_pct"].isna().all():
        if "add_to_carts" in normalized.columns and "pdp_views" in normalized.columns:
            normalized["c2d_pct"] = safe_div(normalized["add_to_carts"], normalized["pdp_views"]) * 100

    if "b2d_pct" in normalized.columns and normalized["b2d_pct"].isna().all():
        if "transactions" in normalized.columns and "pdp_views" in normalized.columns:
            normalized["b2d_pct"] = safe_div(normalized["transactions"], normalized["pdp_views"]) * 100

    return normalized


def safe_div(numerator, denominator):
    return numerator / denominator.replace({0: pd.NA})


def read_excel_smart(path: str, sheet_name=0) -> pd.DataFrame:
    """
    Demo Excel'lerinde bazen ilk 2-3 satır açıklama, gerçek kolonlar daha aşağıda oluyor.
    Bu fonksiyon gerçek header satırını otomatik bulur.
    """
    normal = pd.read_excel(path, sheet_name=sheet_name)
    normal_cols = [normalize_column_name(c) for c in normal.columns]

    # Normal okumada kolonlar zaten doğruysa direkt dön.
    known_hits = sum(
        1 for c in normal_cols
        if c in {"gtin", "sku", "brand", "cat1", "cat2", "product_title", "price", "revenue", "benchmark_price"}
    )
    if known_hits >= 3:
        return normal

    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    best_idx = None
    best_score = -1
    known = {"gtin", "sku", "brand", "cat1", "cat2", "product_title", "price", "revenue", "benchmark_price", "pdp_views", "transactions"}

    for idx in range(min(len(raw), 15)):
        row = raw.iloc[idx].fillna("").astype(str).map(normalize_column_name).tolist()
        score = sum(1 for v in row if v in known)
        if score > best_score:
            best_score = score
            best_idx = idx

    if best_idx is not None and best_score >= 3:
        df = raw.iloc[best_idx + 1:].copy()
        df.columns = raw.iloc[best_idx].fillna("").astype(str).tolist()
        df = df.dropna(how="all").reset_index(drop=True)
        return df

    return normal


@lru_cache(maxsize=1)
def load_insight_data():
    source_path = None

    if os.path.exists(COMPANY_INPUT_PATH):
        source_path = COMPANY_INPUT_PATH
        df = read_excel_smart(COMPANY_INPUT_PATH)
    elif os.path.exists(FALLBACK_SAMPLE_PATH):
        source_path = FALLBACK_SAMPLE_PATH
        df = read_excel_smart(FALLBACK_SAMPLE_PATH, sheet_name=FALLBACK_SAMPLE_SHEET)
    else:
        raise FileNotFoundError(
            f"Insight data bulunamadı. Beklenen dosyalar: {COMPANY_INPUT_PATH} veya {FALLBACK_SAMPLE_PATH}"
        )

    df = normalize_columns(df)
    df["_source_path"] = source_path

    if os.path.exists(MERCHANT_BENCHMARK_SAMPLE_PATH):
        try:
            benchmark_df = read_excel_smart(MERCHANT_BENCHMARK_SAMPLE_PATH)
            benchmark_df = normalize_columns(benchmark_df)

            if "gtin" in df.columns and "gtin" in benchmark_df.columns:
                benchmark_cols = [c for c in ["gtin", "benchmark_price"] if c in benchmark_df.columns]
                if "benchmark_price" in benchmark_cols:
                    df = df.drop(columns=["benchmark_price"], errors="ignore")
                    df["gtin"] = df["gtin"].astype(str).str.strip()
                    benchmark_df["gtin"] = benchmark_df["gtin"].astype(str).str.strip()
                    df = df.merge(benchmark_df[benchmark_cols], on="gtin", how="left")
        except Exception as e:
            print("⚠️ Merchant benchmark join atlandı:", str(e), flush=True)

    if "price" in df.columns and "benchmark_price" in df.columns:
        df["price_gap"] = df["price"] - df["benchmark_price"]
        df["price_gap_pct"] = safe_div(df["price_gap"], df["benchmark_price"]) * 100

    return df


def detect_question_type(question: Optional[str]) -> str:
    q = normalize_text(question)

    if any(x in q for x in ["kazanan", "kaybeden", "winner", "loser", "segmentleri cikar", "segmentleri çıkar"]):
        return "winner_loser_analysis"
    if any(x in q for x in ["stoktan mi", "stoktan mı", "fiyattan mi", "fiyattan mı", "talepten mi", "talepten mı", "neden etkilen", "sebep", "root cause"]):
        return "demand_price_stock_analysis"
    if any(x in q for x in ["firsat", "fırsat", "opportunity", "buyume", "büyüme"]):
        return "opportunity_analysis"
    if any(x in q for x in ["risk", "riskli", "tehdit", "dusuyor", "düşüyor", "zayif", "zayıf"]):
        return "risk_analysis"
    if any(x in q for x in ["funnel", "c2d", "b2d", "a2c", "checkout", "sepete", "satin alma", "satın alma"]):
        return "funnel_analysis"
    if any(x in q for x in ["ceo", "yonetici", "yönetici", "executive", "ozet", "özet"]):
        return "ceo_summary"
    if any(x in q for x in ["aksiyon", "oner", "öner", "ne yap", "plan"]):
        return "action_recommendation"
    return "general_performance"


def detect_category_from_question(df: pd.DataFrame, question: Optional[str]) -> Optional[str]:
    q = normalize_text(question)
    if not q:
        return None

    for user_word, values in CATEGORY_SYNONYMS.items():
        if normalize_text(user_word) in q:
            return user_word
        for value in values:
            if normalize_text(value) in q:
                return user_word

    for col in ["cat2", "cat1", "brand"]:
        if col not in df.columns:
            continue
        values = [x for x in df[col].dropna().astype(str).unique().tolist() if x and x.lower() != "nan"]
        for value in sorted(values, key=len, reverse=True):
            if normalize_text(value) and normalize_text(value) in q:
                return value

    return None


def filter_by_category(df: pd.DataFrame, category: str, question: Optional[str] = None) -> Tuple[pd.DataFrame, str]:
    selected = category or "genel"

    if normalize_text(selected) in ["", "genel", "all", "tum", "tüm", "overall", "site"]:
        detected = detect_category_from_question(df, question)
        if detected:
            selected = detected

    q = normalize_text(selected)

    if q in ["", "genel", "all", "tum", "tüm", "overall", "site"]:
        return df.copy(), "Genel"

    search_values = []
    if q in CATEGORY_SYNONYMS:
        search_values = [q] + [normalize_text(v) for v in CATEGORY_SYNONYMS[q]]
    else:
        search_values = [q]

    search_cols = [c for c in ["cat1", "cat2", "brand", "product_title", "sku", "sales_channel", "traffic_channel"] if c in df.columns]
    mask = pd.Series(False, index=df.index)

    for col in search_cols:
        normalized_col = df[col].astype(str).apply(normalize_text)
        for value in search_values:
            if value:
                mask = mask | normalized_col.str.contains(value, na=False, regex=False)

    filtered = df[mask].copy()
    return filtered, selected


def generate_category_insight(
    category: str = "genel",
    sector: str = "consumer_electronics",
    period_name: str = "selected_period",
    question: Optional[str] = None,
):
    """
    Category Insights v2.
    Kategori / sektör / dönem bazlı retail insight üretir.
    Doğal CEO özeti, sinyal yorumu, root cause ve aksiyon önerisi döndürür.
    """

    try:
        df = load_insight_data()
        print(
            "✅ CATEGORY INSIGHT DATA OKUNDU:",
            df.get("_source_path", pd.Series(["unknown"])).iloc[0] if len(df) else "unknown",
            "SATIR:",
            len(df),
            "KOLON:",
            len(df.columns),
            flush=True,
        )
    except Exception as e:
        return json.dumps(
            {
                "analysis_type": "category_sector_insight_error",
                "error": "Insight data okunamadı.",
                "detail": str(e),
                "expected_paths": [COMPANY_INPUT_PATH, FALLBACK_SAMPLE_PATH],
            },
            ensure_ascii=False,
            default=str,
        )

    sector = sector or "consumer_electronics"
    profile = SECTOR_PROFILES.get(sector, SECTOR_PROFILES.get("marketplace_general", {"name": sector}))
    question_type = detect_question_type(question)

    filtered_df, resolved_category = filter_by_category(df, category, question)

    if filtered_df.empty:
        return json.dumps(
            {
                "analysis_type": "category_sector_insight_error",
                "error": "Bu kategori için veri bulunamadı.",
                "category": category,
                "resolved_category": resolved_category,
                "question": question,
                "available_cat1": sorted(df["cat1"].dropna().astype(str).unique().tolist())[:50] if "cat1" in df.columns else [],
                "available_cat2": sorted(df["cat2"].dropna().astype(str).unique().tolist())[:50] if "cat2" in df.columns else [],
            },
            ensure_ascii=False,
            default=str,
        )

    summary = calculate_period_summary(filtered_df)
    baseline = calculate_period_summary(df)
    metric_snapshot = build_metric_snapshot(summary, baseline)

    channel_insights = calculate_group_insights(filtered_df, "sales_channel")
    traffic_group_col = "main_traffic_channel" if "main_traffic_channel" in filtered_df.columns else "traffic_channel"
    traffic_insights = calculate_group_insights(filtered_df, traffic_group_col)

    breakdown_col = "cat2" if "cat2" in filtered_df.columns and safe_unique_count(filtered_df, "cat2") > 1 else "cat1"
    category_performance = calculate_group_insights(filtered_df if normalize_text(resolved_category) == "genel" else df, breakdown_col)

    winners = category_performance.sort_values("performance_score", ascending=False).head(5) if not category_performance.empty else pd.DataFrame()
    losers = category_performance.sort_values("performance_score", ascending=True).head(5) if not category_performance.empty else pd.DataFrame()

    signal_interpretation = classify_signals(summary, baseline, question_type)
    root_causes = build_root_causes(summary, baseline, signal_interpretation, channel_insights, traffic_insights)
    main_diagnosis = build_main_diagnosis(resolved_category, question_type, summary, baseline, root_causes, signal_interpretation)
    natural_summary = build_natural_summary(resolved_category, question_type, summary, baseline, root_causes, signal_interpretation)
    actions = build_recommendations(resolved_category, question_type, summary, root_causes, signal_interpretation, winners, losers)

    rows = build_category_rows(filtered_df)

    behavior_data = get_behavior_for_category(resolved_category)
    consumer_behavior_analysis = {
        "display_name": behavior_data.get("display_name", resolved_category) if behavior_data else resolved_category,
        "general_behavior": behavior_data.get("general_behavior", "") if behavior_data else "",
        "triggered_insights": []
    }
    if behavior_data:
        for rule in behavior_data.get("rules", []):
            try:
                if rule["condition"](summary):
                    consumer_behavior_analysis["triggered_insights"].append(rule["insight"])
            except Exception as e:
                print(f"Warning: error evaluating rule {rule.get('id')} in insights.py: {e}", flush=True)

    price_scenarios = analyze_price_competition_scenarios(filtered_df)
    seasonal_trends = get_seasonal_trends_insights(resolved_category)

    result = {
        "analysis_type": "category_sector_insight",
        "question": question,
        "question_type": question_type,
        "category": resolved_category,
        "sector": sector,
        "sector_name": profile.get("name", sector),
        "period_name": period_name,
        "data_scope": {
            "rows": int(len(filtered_df)),
            "brands": safe_unique_count(filtered_df, "brand"),
            "cat1_count": safe_unique_count(filtered_df, "cat1"),
            "cat2_count": safe_unique_count(filtered_df, "cat2"),
            "sku_count": safe_unique_count(filtered_df, "sku"),
        },
        "executive_summary": summary,
        "baseline_summary": baseline,
        "metric_snapshot": metric_snapshot,
        "natural_summary": natural_summary,
        "main_diagnosis": main_diagnosis,
        "signal_interpretation": signal_interpretation,
        "root_causes": root_causes,
        "consumer_behavior_analysis": consumer_behavior_analysis,
        "channel_insights": df_to_records(channel_insights.head(10)),
        "traffic_insights": df_to_records(traffic_insights.head(10)),
        "winning_categories": df_to_records(winners),
        "losing_categories": df_to_records(losers),
        "recommended_actions": actions,
        "price_scenarios": price_scenarios,
        "seasonal_trends": seasonal_trends,
        "rows": rows,
        "caveat": (
            "Bu analiz mevcut Excel datasındaki stok, fiyat, funnel ve satış sinyallerine göre üretilmiştir. "
            "Gerçek zamanlı BigQuery, Merchant Center, spend, margin, rating ve iade verileri bağlandığında neden analizi daha da güçlenir."
        ),
    }

    return json.dumps(result, ensure_ascii=False, default=str)


def calculate_period_summary(df: pd.DataFrame) -> Dict:
    revenue = safe_sum(df, "revenue")
    transactions = safe_sum(df, "transactions")
    pdp_views = safe_sum(df, "pdp_views")
    a2c = safe_sum(df, "add_to_carts")
    stock_qty = safe_sum(df, "stock_qty")
    lost_revenue = safe_sum(df, "estimated_lost_revenue")
    avg_price = weighted_avg(df, "price", "revenue")

    summary = {
        "revenue": round(revenue, 2),
        "transactions": round(transactions, 2),
        "pdp_views": round(pdp_views, 2),
        "add_to_carts": round(a2c, 2),
        "stock_qty": round(stock_qty, 2),
        "avg_price": round(avg_price, 2),
        "estimated_lost_revenue": round(lost_revenue, 2),
        "revenue_delta_pct": round(weighted_avg(df, "revenue_delta_pct", "revenue"), 2),
        "transactions_delta_pct": round(weighted_avg(df, "transactions_delta_pct", "transactions"), 2),
        "pdp_delta_pct": round(weighted_avg(df, "pdp_delta_pct", "pdp_views"), 2),
        "a2c_delta_pct": round(weighted_avg(df, "a2c_delta_pct", "add_to_carts"), 2),
        "c2d_pct": round(weighted_avg(df, "c2d_pct", "pdp_views"), 2),
        "c2d_delta_pct": round(weighted_avg(df, "c2d_delta_pct", "pdp_views"), 2),
        "b2d_pct": round(weighted_avg(df, "b2d_pct", "pdp_views"), 2),
        "b2d_delta_pct": round(weighted_avg(df, "b2d_delta_pct", "pdp_views"), 2),
        "product_price_delta_pct": round(weighted_avg(df, "product_price_delta_pct", "revenue"), 2),
        "bounce_rate_pct": round(weighted_avg(df, "bounce_rate_pct", "pdp_views"), 2),
        "bounce_rate_delta_pct": round(weighted_avg(df, "bounce_rate_delta_pct", "pdp_views"), 2),
        "checkout_submit_delta_pct": round(weighted_avg(df, "checkout_submit_delta_pct", "checkout_submits"), 2),
        "avg_stock_coverage_days": round(weighted_avg(df, "stock_coverage_days", "revenue"), 2),
        "avg_price_gap_pct": round(weighted_avg(df, "price_gap_pct", "revenue"), 2),
        "benchmark_above_sku_count": int((df.get("price_gap_pct", pd.Series(dtype=float)) > 5).sum()) if "price_gap_pct" in df.columns else 0,
        "critical_stock_sku_count": count_critical_stock(df),
        "oos_sku_count": count_oos(df),
        "sku_count": safe_unique_count(df, "sku"),
    }

    return summary


def build_metric_snapshot(summary: Dict, baseline: Dict) -> Dict:
    return {
        "category_revenue": summary.get("revenue"),
        "category_transactions": summary.get("transactions"),
        "category_pdp_views": summary.get("pdp_views"),
        "category_c2d_pct": summary.get("c2d_pct"),
        "category_b2d_pct": summary.get("b2d_pct"),
        "category_revenue_delta_pct": summary.get("revenue_delta_pct"),
        "category_transactions_delta_pct": summary.get("transactions_delta_pct"),
        "category_pdp_delta_pct": summary.get("pdp_delta_pct"),
        "category_a2c_delta_pct": summary.get("a2c_delta_pct"),
        "category_b2d_delta_pct": summary.get("b2d_delta_pct"),
        "category_stock_qty": summary.get("stock_qty"),
        "category_critical_stock_sku_count": summary.get("critical_stock_sku_count"),
        "category_oos_sku_count": summary.get("oos_sku_count"),
        "category_avg_price_gap_pct": summary.get("avg_price_gap_pct"),
        "overall_c2d_pct": baseline.get("c2d_pct"),
        "overall_b2d_pct": baseline.get("b2d_pct"),
        "overall_revenue_delta_pct": baseline.get("revenue_delta_pct"),
        "overall_transactions_delta_pct": baseline.get("transactions_delta_pct"),
        "overall_pdp_delta_pct": baseline.get("pdp_delta_pct"),
        "overall_b2d_delta_pct": baseline.get("b2d_delta_pct"),
    }


def calculate_group_insights(df: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if group_col not in df.columns:
        return pd.DataFrame()

    rows = []
    for key, g in df.groupby(group_col, dropna=False):
        summary = calculate_period_summary(g)
        stock_penalty = min(summary.get("critical_stock_sku_count", 0) * 1.5 + summary.get("oos_sku_count", 0) * 3, 25)
        price_penalty = 0
        if summary.get("avg_price_gap_pct", 0) > 5:
            price_penalty = min(summary.get("avg_price_gap_pct", 0), 15)

        performance_score = (
            summary.get("revenue_delta_pct", 0) * 0.30
            + summary.get("transactions_delta_pct", 0) * 0.25
            + summary.get("pdp_delta_pct", 0) * 0.15
            + summary.get("a2c_delta_pct", 0) * 0.15
            + summary.get("b2d_delta_pct", 0) * 0.15
            - stock_penalty
            - price_penalty
        )

        row = {
            group_col: str(key),
            **summary,
            "performance_score": round(performance_score, 2),
        }
        rows.append(row)

    result = pd.DataFrame(rows)
    if result.empty:
        return result
    return result.sort_values("revenue", ascending=False)


def classify_signals(summary: Dict, baseline: Dict, question_type: str) -> List[Dict]:
    signals = []

    def add(name, interpretation, evidence, polarity="neutral", confidence="medium"):
        signals.append(
            {
                "signal": name,
                "interpretation": interpretation,
                "evidence": evidence,
                "polarity": polarity,
                "confidence": confidence,
            }
        )

    pdp_delta = summary.get("pdp_delta_pct", 0)
    a2c_delta = summary.get("a2c_delta_pct", 0)
    b2d_delta = summary.get("b2d_delta_pct", 0)
    c2d_delta = summary.get("c2d_delta_pct", 0)
    revenue_delta = summary.get("revenue_delta_pct", 0)
    trans_delta = summary.get("transactions_delta_pct", 0)
    price_delta = summary.get("product_price_delta_pct", 0)
    price_gap = summary.get("avg_price_gap_pct", 0)
    stock_risk = summary.get("critical_stock_sku_count", 0) + summary.get("oos_sku_count", 0)

    if pdp_delta > 5 and a2c_delta > 5:
        add("Talep ilgisi", "Kullanıcı ilgisi güçlü; PDP ve sepete ekleme sinyalleri birlikte yukarı gidiyor.", f"PDP delta %{pdp_delta}, A2C delta %{a2c_delta}.", "positive", "high")
    elif pdp_delta < -8:
        add("Talep daralması", "Kategoriye gelen ilgi zayıflamış görünüyor.", f"PDP delta %{pdp_delta}.", "negative", "medium")

    if c2d_delta > 0 and b2d_delta < 0:
        add("Funnel bariyeri", "Sepete ekleme niyeti var; ancak satın alma tarafında bariyer oluşuyor.", f"C2D delta %{c2d_delta}, B2D delta %{b2d_delta}.", "negative", "high")
    elif b2d_delta > 0:
        add("Satın alma niyeti", "Ürüne bakan kullanıcıların satın alma eğilimi güçleniyor.", f"B2D delta %{b2d_delta}.", "positive", "medium")

    if revenue_delta > 5 and trans_delta > 5:
        add("Satış momentumu", "Revenue ve transaction birlikte büyüdüğü için kategori sağlıklı momentum gösteriyor.", f"Revenue delta %{revenue_delta}, transaction delta %{trans_delta}.", "positive", "high")
    elif revenue_delta < -8 and trans_delta < -8:
        add("Satış zayıflığı", "Revenue ve transaction birlikte düştüğü için kategori satış tarafında baskı altında.", f"Revenue delta %{revenue_delta}, transaction delta %{trans_delta}.", "negative", "high")

    if price_delta > 5 and b2d_delta < 0:
        add("Fiyat hassasiyeti", "Fiyat artışı satın alma dönüşümünü baskılıyor olabilir.", f"Fiyat delta %{price_delta}, B2D delta %{b2d_delta}.", "negative", "medium")

    if price_gap > 5 and b2d_delta < 0:
        add("Fiyat rekabeti", "Benchmark üstü fiyat pozisyonu conversion üzerinde baskı yaratıyor olabilir.", f"Ortalama price gap %{price_gap}, B2D delta %{b2d_delta}.", "negative", "medium")
    elif price_gap < -5:
        add("Fiyat avantajı", "Kategori benchmark altında fiyat avantajına sahip görünüyor.", f"Ortalama price gap %{price_gap}.", "positive", "medium")

    if stock_risk > 0:
        add("Stok riski", "Talep satışa dönüşmeden kaybedilebilir; stok ve bulunabilirlik kontrol edilmeli.", f"Kritik stok + OOS SKU sayısı {stock_risk}.", "negative", "medium")

    if not signals:
        add("Karışık sinyal", "Kategori tek bir net nedene işaret etmiyor; SKU, kanal ve fiyat kırılımı birlikte incelenmeli.", "Metrikler belirgin tek yönde hareket etmiyor.", "neutral", "low")

    return signals


def build_root_causes(summary: Dict, baseline: Dict, signals: List[Dict], channel_insights: pd.DataFrame, traffic_insights: pd.DataFrame) -> List[Dict]:
    causes = []

    signal_map = {s.get("signal"): s for s in signals}
    for key in ["Talep daralması", "Funnel bariyeri", "Fiyat hassasiyeti", "Fiyat rekabeti", "Stok riski", "Satış zayıflığı"]:
        if key in signal_map:
            s = signal_map[key]
            causes.append({"cause": key, "evidence": s.get("evidence"), "confidence": s.get("confidence", "medium")})

    paid_vs_unpaid = detect_paid_vs_unpaid_effect(traffic_insights)
    if paid_vs_unpaid:
        causes.append(paid_vs_unpaid)

    channel_shift = detect_channel_shift(channel_insights)
    if channel_shift:
        causes.append(channel_shift)

    if not causes:
        strongest = signals[0] if signals else {}
        causes.append(
            {
                "cause": strongest.get("signal", "Net tekil problem bulunamadı"),
                "evidence": strongest.get("evidence", "Metrikler karışık sinyal veriyor."),
                "confidence": strongest.get("confidence", "low"),
            }
        )

    return causes[:5]


def build_main_diagnosis(category: str, question_type: str, summary: Dict, baseline: Dict, root_causes: List[Dict], signals: List[Dict]) -> str:
    cat = category_display_name(category)
    positives = [s for s in signals if s.get("polarity") == "positive"]
    negatives = [s for s in signals if s.get("polarity") == "negative"]

    if question_type == "opportunity_analysis":
        if positives:
            return f"{cat}, mevcut sinyallere göre fırsat potansiyeli taşıyor; özellikle {positives[0]['signal'].lower()} tarafı güçlü görünüyor."
        return f"{cat} için fırsat sinyali sınırlı; büyüme kararı öncesinde stok, fiyat ve funnel kırılımları kontrol edilmeli."

    if question_type == "risk_analysis":
        if negatives:
            return f"{cat} kategorisindeki ana risk {negatives[0]['signal'].lower()} alanında yoğunlaşıyor."
        return f"{cat} kategorisinde belirgin kritik risk görünmüyor; ancak SKU bazında kontrol önerilir."

    if question_type == "winner_loser_analysis":
        return "Kategori bazlı performansta kazanan ve kaybeden segmentler, revenue/transaction/funnel/stok sinyallerine göre ayrıştırıldı."

    if question_type == "demand_price_stock_analysis":
        if root_causes:
            return f"Mevcut sinyallere göre performansı en çok {root_causes[0]['cause'].lower()} açıklıyor."
        return "Performans tek bir nedene bağlı görünmüyor; talep, fiyat ve stok sinyalleri birlikte değerlendirilmelidir."

    if negatives and not positives:
        return f"{cat} baskı altında; ana negatif sinyal {negatives[0]['signal'].lower()} olarak görünüyor."
    if positives and not negatives:
        return f"{cat} sağlıklı performans gösteriyor; ana pozitif sinyal {positives[0]['signal'].lower()} tarafında."
    if positives and negatives:
        return f"{cat} karışık sinyal veriyor; fırsat var ancak {negatives[0]['signal'].lower()} yönetilmeli."

    return f"{cat} için net bir ana yön oluşmuyor; daha detaylı SKU ve kanal kırılımı önerilir."


def build_natural_summary(category: str, question_type: str, summary: Dict, baseline: Dict, root_causes: List[Dict], signals: List[Dict]) -> str:
    cat = category_display_name(category)
    positive = [s for s in signals if s.get("polarity") == "positive"]
    negative = [s for s in signals if s.get("polarity") == "negative"]

    if question_type == "opportunity_analysis":
        if positive:
            base = f"{cat}, bu dönem fırsat kategorisi olarak değerlendirilebilir. {positive[0]['interpretation']}"
        else:
            base = f"{cat} için fırsat potansiyeli tamamen net değil; mevcut veri daha temkinli bir büyüme yaklaşımına işaret ediyor."
        if negative:
            base += f" Ancak {negative[0]['interpretation'].lower()} Bu nedenle fırsatı büyütmeden önce ilgili risk yönetilmeli."
        return base

    if question_type == "risk_analysis":
        if negative:
            return f"{cat} kategorisi riskli görünüyor çünkü {negative[0]['interpretation'].lower()} {negative[0]['evidence']}"
        return f"{cat} kategorisinde belirgin bir kırmızı alarm görünmüyor. Buna rağmen stok, fiyat ve funnel kırılımları düzenli izlenmeli."

    if question_type == "winner_loser_analysis":
        return "Kazanan ve kaybeden segmentler, revenue değişimi, transaction değişimi, PDP/A2C hareketi, B2D ve stok riski birlikte değerlendirilerek ayrıştırıldı."

    if question_type == "demand_price_stock_analysis":
        main = root_causes[0] if root_causes else None
        if main:
            return f"Satış performansındaki değişimi tek başına bir metrikle açıklamak doğru olmaz; ancak mevcut veri en güçlü sinyal olarak {main['cause'].lower()} alanını gösteriyor. {main['evidence']}"
        return "Satış performansı stok, fiyat ve talep sinyallerinin birleşimiyle şekilleniyor; mevcut veri tek bir ana neden göstermiyor."

    if question_type == "funnel_analysis":
        funnel_signal = next((s for s in signals if s.get("signal") in ["Funnel bariyeri", "Satın alma niyeti", "Talep ilgisi"]), None)
        if funnel_signal:
            return f"{cat} funnel tarafında {funnel_signal['interpretation'].lower()} {funnel_signal['evidence']}"
        return f"{cat} funnel performansı için PDP, A2C, C2D ve B2D kırılımlarının birlikte izlenmesi gerekiyor."

    if positive and negative:
        return f"{cat} kategorisi karışık ama aksiyon alınabilir sinyaller veriyor. {positive[0]['interpretation']} Bununla birlikte {negative[0]['interpretation'].lower()}"
    if positive:
        return f"{cat} kategorisi pozitif momentum gösteriyor. {positive[0]['interpretation']}"
    if negative:
        return f"{cat} kategorisinde dikkat edilmesi gereken ana konu {negative[0]['signal'].lower()}. {negative[0]['interpretation']}"

    return f"{cat} kategorisi için mevcut sinyaller dengeli görünüyor; daha net karar için SKU, kanal ve fiyat kırılımları birlikte incelenmeli."


def build_recommendations(category: str, question_type: str, summary: Dict, root_causes: List[Dict], signals: List[Dict], winners: pd.DataFrame, losers: pd.DataFrame) -> List[str]:
    cat = category_display_name(category)
    actions = []
    signal_names = {s.get("signal") for s in signals}

    if "Stok riski" in signal_names:
        actions.append(f"{cat} içinde stok riski taşıyan SKU’lar için replenishment planı çıkar ve kampanya görünürlüğünü stok güvenceye alınana kadar kontrollü artır.")
    if "Fiyat rekabeti" in signal_names or "Fiyat hassasiyeti" in signal_names:
        actions.append(f"{cat} için benchmark üstü fiyatlanan SKU’larda fiyat gap, B2D ve margin birlikte kontrol edilmeli.")
    if "Funnel bariyeri" in signal_names:
        actions.append(f"{cat} ürünlerinde sepete ekleme sonrası satın alma bariyerini bulmak için kargo, ödeme, stok ve fiyat kırılımı incelenmeli.")
    if "Talep ilgisi" in signal_names or "Satış momentumu" in signal_names:
        actions.append(f"{cat} için yüksek talep sinyali veren SKU’larda onsite görünürlük ve CRM/paid destek artırılabilir.")
    if "Talep daralması" in signal_names:
        actions.append(f"{cat} için talep yaratacak kampanya, landing visibility ve kanal bazlı trafik planı gözden geçirilmeli.")

    if question_type == "winner_loser_analysis":
        actions.append("Kazanan segmentlerde bütçe/görünürlük artırılırken kaybeden segmentlerde fiyat, stok ve funnel nedeni ayrı ayrı kontrol edilmeli.")

    if not actions:
        actions.append(f"{cat} için kategori, SKU ve kanal bazında daha detaylı kırılım yapılarak büyüme ve risk alanları ayrıştırılmalı.")

    actions.append("Bu insight ekran özeti olarak verildi; detaylı SKU ve metrik kırılımı Excel çıktısından incelenebilir.")
    return actions[:6]


def build_category_rows(df: pd.DataFrame) -> List[Dict]:
    cols = [
        "sku", "product_title", "brand", "cat1", "cat2", "sales_channel", "traffic_channel",
        "price", "benchmark_price", "price_gap_pct", "stock_qty", "reorder_point_qty",
        "pdp_views", "add_to_carts", "transactions", "revenue", "c2d_pct", "b2d_pct",
        "revenue_delta_pct", "transactions_delta_pct", "pdp_delta_pct", "a2c_delta_pct",
        "c2d_delta_pct", "b2d_delta_pct", "estimated_lost_revenue",
    ]
    existing = [c for c in cols if c in df.columns]
    out = df[existing].copy()
    sort_col = "revenue" if "revenue" in out.columns else existing[0]
    out = out.sort_values(sort_col, ascending=False).head(50)
    return df_to_records(out)


def detect_paid_vs_unpaid_effect(traffic_insights: pd.DataFrame):
    if traffic_insights is None or traffic_insights.empty:
        return None

    channel_col = "traffic_channel" if "traffic_channel" in traffic_insights.columns else "main_traffic_channel"
    if channel_col not in traffic_insights.columns:
        return None

    df = traffic_insights.copy()
    df["channel_lower"] = df[channel_col].astype(str).apply(normalize_text)
    paid = df[df["channel_lower"].str.contains("paid|cpc|ppc|ads|search paid|social paid", na=False, regex=True)]
    unpaid = df[df["channel_lower"].str.contains("organic|direct|unpaid|seo|crm", na=False, regex=True)]

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


def detect_channel_shift(channel_insights: pd.DataFrame):
    if channel_insights is None or channel_insights.empty or "sales_channel" not in channel_insights.columns:
        return None

    df = channel_insights.copy()
    df["sales_channel_lower"] = df["sales_channel"].astype(str).apply(normalize_text)
    app_mobile = df[df["sales_channel_lower"].str.contains("app|mobile|mobil", na=False, regex=True)]
    desktop = df[df["sales_channel_lower"].str.contains("desktop", na=False, regex=True)]

    if app_mobile.empty or desktop.empty:
        return None

    app_delta = app_mobile["transactions_delta_pct"].mean()
    desktop_delta = desktop["transactions_delta_pct"].mean()

    if app_delta > desktop_delta + 10:
        return {
            "cause": "Mobil/App kanalı Desktop’a göre daha dirençli",
            "evidence": f"App/Mobile transaction delta %{round(app_delta, 2)}, Desktop transaction delta %{round(desktop_delta, 2)}.",
            "confidence": "medium",
        }
    return None


def safe_sum(df: pd.DataFrame, col: str) -> float:
    if col not in df.columns:
        return 0.0
    return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())


def weighted_avg(df: pd.DataFrame, value_col: str, weight_col: str = None) -> float:
    if value_col not in df.columns:
        return 0.0

    values = pd.to_numeric(df[value_col], errors="coerce")

    if weight_col and weight_col in df.columns:
        weights = pd.to_numeric(df[weight_col], errors="coerce").fillna(0)
        valid = values.notna() & weights.notna() & (weights > 0)
        if valid.any() and weights[valid].sum() != 0:
            return float((values[valid] * weights[valid]).sum() / weights[valid].sum())

    return float(values.dropna().mean()) if values.dropna().shape[0] else 0.0


def count_critical_stock(df: pd.DataFrame) -> int:
    if "stock_qty" not in df.columns:
        return 0

    stock = pd.to_numeric(df["stock_qty"], errors="coerce")
    mask = pd.Series(False, index=df.index)

    if "reorder_point_qty" in df.columns:
        reorder = pd.to_numeric(df["reorder_point_qty"], errors="coerce")
        mask = mask | (stock <= reorder)
    else:
        q25 = stock.quantile(0.25) if stock.notna().any() else 0
        mask = mask | (stock <= q25)

    if "stock_coverage_days" in df.columns:
        coverage = pd.to_numeric(df["stock_coverage_days"], errors="coerce")
        mask = mask | (coverage <= 7)

    if "availability_status" in df.columns:
        availability = df["availability_status"].astype(str).apply(normalize_text)
        mask = mask | availability.str.contains("out|oos|stok yok", na=False, regex=True)

    return int(mask.fillna(False).sum())


def count_oos(df: pd.DataFrame) -> int:
    mask = pd.Series(False, index=df.index)

    if "stock_qty" in df.columns:
        stock = pd.to_numeric(df["stock_qty"], errors="coerce")
        mask = mask | (stock == 0)

    if "availability_status" in df.columns:
        availability = df["availability_status"].astype(str).apply(normalize_text)
        mask = mask | availability.str.contains("out_of_stock|oos|stok yok", na=False, regex=True)

    return int(mask.fillna(False).sum())


def safe_unique_count(df: pd.DataFrame, col: str) -> int:
    if col not in df.columns:
        return 0
    return int(df[col].replace("", pd.NA).nunique(dropna=True))


def df_to_records(df: pd.DataFrame) -> List[Dict]:
    if df is None or df.empty:
        return []
    clean_df = df.copy()
    clean_df = clean_df.where(pd.notnull(clean_df), None)
    return clean_df.to_dict(orient="records")


def category_display_name(category: str) -> str:
    if not category or normalize_text(category) in ["genel", "all", "overall", "site"]:
        return "Genel kategori performansı"

    mapping = {
        "mobile": "Cep Telefonları",
        "mobil": "Cep Telefonları",
        "gsm": "Cep Telefonları",
        "telefon": "Cep Telefonları",
        "cep telefonlari": "Cep Telefonları",
        "cep telefonları": "Cep Telefonları",
        "tablet": "Tabletler",
        "tabletler": "Tabletler",
        "kulaklik": "Kulaklıklar",
        "kulaklık": "Kulaklıklar",
        "giyilebilir": "Giyilebilir Teknoloji",
        "giyilebilir teknoloji": "Giyilebilir Teknoloji",
        "aksesuar": "Aksesuar",
        "tv": "TV",
        "televizyon": "TV",
        "supurge": "Süpürge",
        "süpürge": "Süpürge",
    }
    return mapping.get(normalize_text(category), str(category))


def analyze_price_competition_scenarios(df: pd.DataFrame) -> Dict:
    scenarios = {
        "expensive_falling_sales": [], # Pahalıyız ve satış düşüyor -> Price action öner
        "expensive_good_sales": [],    # Pahalıyız ama satış iyi -> Premium/güçlü ürün
        "cheap_no_sales": [],          # Ucuzuz ama satış yok -> Visibility/content/stok sorunu
        "cheap_good_sales": [],        # Ucuzuz ve satış iyi -> Trafik artır / bid artır
        "losing_competitiveness": []   # Fiyat rekabetini kaybettiğimiz marka/kategoriler
    }
    
    if df.empty:
        return scenarios
        
    # Make copies and fillnas to avoid errors
    df_clean = df.copy()
    for col in ["price_gap_pct", "revenue_delta_pct", "transactions", "revenue", "stock_qty"]:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0)
        else:
            df_clean[col] = 0.0
            
    # 1. Expensive & Falling Sales (Scenario 1)
    s1_mask = (df_clean["price_gap_pct"] > 2) & (df_clean["revenue_delta_pct"] < -5)
    s1_df = df_clean[s1_mask].sort_values("revenue", ascending=False).head(5)
    for _, row in s1_df.iterrows():
        scenarios["expensive_falling_sales"].append({
            "sku": row.get("sku", ""),
            "product_title": row.get("product_title", ""),
            "brand": row.get("brand", ""),
            "price": float(row.get("price", 0)),
            "benchmark_price": float(row.get("benchmark_price", 0)),
            "price_gap_pct": float(row.get("price_gap_pct", 0)),
            "revenue_delta_pct": float(row.get("revenue_delta_pct", 0)),
            "action": "Price Action Önerilir: Fiyatı benchmark seviyesine çekerek kaybı durdurun."
        })
        
    # 2. Expensive but Good Sales (Scenario 2)
    s2_mask = (df_clean["price_gap_pct"] > 2) & (df_clean["revenue_delta_pct"] >= 0)
    s2_df = df_clean[s2_mask].sort_values("revenue", ascending=False).head(5)
    for _, row in s2_df.iterrows():
        scenarios["expensive_good_sales"].append({
            "sku": row.get("sku", ""),
            "product_title": row.get("product_title", ""),
            "brand": row.get("brand", ""),
            "price": float(row.get("price", 0)),
            "benchmark_price": float(row.get("benchmark_price", 0)),
            "price_gap_pct": float(row.get("price_gap_pct", 0)),
            "revenue_delta_pct": float(row.get("revenue_delta_pct", 0)),
            "action": "Premium/Güçlü Ürün: Fiyat esnekliği düşük, mevcut fiyatı ve pozisyonu koruyun."
        })
        
    # 3. Cheap but No Sales (Scenario 3)
    s3_mask = (df_clean["price_gap_pct"] < -2) & (df_clean["transactions"] == 0)
    s3_df = df_clean[s3_mask].sort_values("pdp_views", ascending=False).head(5)
    for _, row in s3_df.iterrows():
        scenarios["cheap_no_sales"].append({
            "sku": row.get("sku", ""),
            "product_title": row.get("product_title", ""),
            "brand": row.get("brand", ""),
            "price": float(row.get("price", 0)),
            "benchmark_price": float(row.get("benchmark_price", 0)),
            "price_gap_pct": float(row.get("price_gap_pct", 0)),
            "stock_qty": int(row.get("stock_qty", 0)),
            "action": "Visibility/Content/Stok Sorunu: Fiyat avantajına rağmen satış yok. Görünürlüğü artırın, içeriği optimize edin veya stok durumunu kontrol edin."
        })
        
    # 4. Cheap and Good Sales (Scenario 4)
    s4_mask = (df_clean["price_gap_pct"] < -2) & (df_clean["transactions"] > 0) & (df_clean["revenue_delta_pct"] >= 0)
    s4_df = df_clean[s4_mask].sort_values("revenue", ascending=False).head(5)
    for _, row in s4_df.iterrows():
        scenarios["cheap_good_sales"].append({
            "sku": row.get("sku", ""),
            "product_title": row.get("product_title", ""),
            "brand": row.get("brand", ""),
            "price": float(row.get("price", 0)),
            "benchmark_price": float(row.get("benchmark_price", 0)),
            "price_gap_pct": float(row.get("price_gap_pct", 0)),
            "revenue_delta_pct": float(row.get("revenue_delta_pct", 0)),
            "action": "Trafik/Bid Artır: Fiyat avantajı satış getiriyor. Satış hacmini katlamak için trafiği ve PPC reklam tekliflerini artırın."
        })
        
    # 5. Losing Competitiveness - Fiyat rekabetini kaybettiğimiz markalar/kategoriler
    # Group by brand, calculate how many SKUs are expensive (>2%) vs total
    if "brand" in df_clean.columns and df_clean["brand"].nunique() > 0:
        brand_groups = []
        for brand, g in df_clean.groupby("brand"):
            if len(g) < 2:  # skip brands with single items for better representation
                continue
            total = len(g)
            expensive_count = (g["price_gap_pct"] > 2).sum()
            ratio = (expensive_count / total) * 100
            brand_groups.append({
                "brand": brand,
                "total_skus": total,
                "expensive_skus": int(expensive_count),
                "ratio": float(ratio),
                "avg_gap": float(g["price_gap_pct"].mean())
            })
        scenarios["losing_competitiveness"] = sorted(brand_groups, key=lambda x: x["ratio"], reverse=True)[:5]
        
    return scenarios


def get_seasonal_trends_insights(category_name: str) -> Optional[Dict]:
    trends_path = "data/google_trends_seasonal_3y.xlsx"
    if not os.path.exists(trends_path):
        return None
        
    try:
        trends_df = pd.read_excel(trends_path)
        norm_cat = normalize_text(category_name)
        
        # Simple string matching to find the column name
        matched_col = None
        for col in trends_df.columns:
            if col == "date":
                continue
            if normalize_text(col) in norm_cat or norm_cat in normalize_text(col):
                matched_col = col
                break
                
        if not matched_col:
            # Fallback to general laptop or first column if not found
            non_date_cols = [c for c in trends_df.columns if c != "date"]
            if non_date_cols:
                matched_col = non_date_cols[0]
            else:
                return None
                
        # Parse date column
        date_col = 'date' if 'date' in trends_df.columns else trends_df.columns[0]
        trends_df[date_col] = pd.to_datetime(trends_df[date_col])
        
        # Calculate monthly average search interest (seasonal pattern)
        trends_df['month'] = trends_df[date_col].dt.month
        monthly_avg = trends_df.groupby('month')[matched_col].mean()
        
        # Find peak months and low months
        peak_month = int(monthly_avg.idxmax())
        low_month = int(monthly_avg.idxmin())
        
        month_names = {
            1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
            7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
        }
        
        # Trend direction
        import numpy as np
        y = trends_df[matched_col].values
        x = np.arange(len(y))
        slope = np.polyfit(x, y, 1)[0] if len(y) > 1 else 0
        trend_direction = "Artış Eğiliminde" if slope > 0.05 else ("Azalış Eğiliminde" if slope < -0.05 else "Dengeli/Stabil")
        
        return {
            "keyword": matched_col,
            "peak_month": month_names.get(peak_month, str(peak_month)),
            "low_month": month_names.get(low_month, str(low_month)),
            "trend_direction": trend_direction,
            "avg_interest_score": round(float(monthly_avg.mean()), 2),
            "historical_peak_value": float(trends_df[matched_col].max())
        }
    except Exception as e:
        print(f"Warning: Failed to parse seasonal trends: {e}", flush=True)
        return None

