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


def get_professor_diagnosis(r, stats):
    c2d = r.get("c2d_pct")
    b2d = r.get("b2d_pct")
    pdp = r.get("pdp_views")
    stock = r.get("stock_qty")
    price_gap = r.get("price_gap_pct")
    reorder = r.get("reorder_point_qty") if pd.notna(r.get("reorder_point_qty")) else 5
    
    c2d_high = c2d >= stats["c2d_med"] if pd.notna(c2d) else False
    b2d_high = b2d >= stats["b2d_med"] if pd.notna(b2d) else False
    pdp_high = pdp >= stats["pdp_med"] if pd.notna(pdp) else False
    
    is_expensive = pd.notna(price_gap) and price_gap > 5
    is_oos_risk = pd.notna(stock) and stock <= reorder
    is_overstock = pd.notna(stock) and stock > reorder * 10
    
    # Kural 1: Kritik Stok Kaybı (Critical OOS Risk)
    if b2d_high and is_oos_risk:
        return "Kritik Stok (OOS Risk)", "Satışa dönüşen bu üründe stok tükenmek üzere. Acil tedarik sürecini başlatın.", 95, "Critical"
    
    # Kural 2: Fiyat Bariyeri (Pricing Barrier)
    if pdp_high and not b2d_high and is_expensive:
        gap_val = round(price_gap, 1) if pd.notna(price_gap) else 0
        return "Fiyat Bariyeri", f"Trafik yüksek ama rakibe göre %{gap_val} pahalı olduğundan dönüşüm düşük. Fiyatı veya kampanyaları gözden geçirin.", 85, "High"
        
    # Kural 3: Sepeti Terk Etme (Cart Abandonment)
    if pdp_high and c2d_high and not b2d_high and not is_expensive:
        return "Sepeti Terk (Cart Abandonment)", "Ürün sepete atılıyor ancak ödemeye geçilmiyor. Kargo bedeli veya ödeme bariyerlerini kontrol edin.", 80, "High"
        
    # Kural 4: Gizli Cevher (Hidden Gem)
    if not pdp_high and b2d_high and not is_oos_risk:
        return "Gizli Cevher (Hidden Gem)", "Trafik düşük olsa da satın alma oranı yüksek. Acilen onsite görünürlüğü ve reklam bütçesini artırın.", 75, "Medium"
        
    # Kural 5: Ölü Yatırım (Overstock)
    if not pdp_high and not b2d_high and is_overstock:
        return "Ölü Yatırım (Overstock)", "Ürün trafik almıyor, satmıyor ve stok maliyeti yaratıyor. 1 Alana 1 Bedava veya agresif indirimlerle stoku eritin.", 60, "Low"
        
    # Default: Core Performer
    if b2d_high and pdp_high:
        return "Core Performer", "Amiral gemisi ürün. Stok seviyesini koruyun ve organik satışını destekleyin.", 50, "Low"
        
    return "Standart Seyir", "Sıradışı bir anomali yok. Rakip rekabeti ve kategori dinamiklerini izlemeye devam edin.", 30, "Low"


def build_comprehensive_action_plan(df, question):
    work = df.copy()

    category_filter = extract_category_filter(work, question)
    work = apply_category_filter(work, category_filter)

    if work.empty:
        return error_result(question, "Profesör analizi için ilgili kategoride ürün bulunamadı.")

    stats = {
        "c2d_med": work["c2d_pct"].median() if "c2d_pct" in work.columns else 0,
        "b2d_med": work["b2d_pct"].median() if "b2d_pct" in work.columns else 0,
        "pdp_med": work["pdp_views"].median() if "pdp_views" in work.columns else 0,
    }

    results = []
    for idx, row in work.iterrows():
        r = row.to_dict()
        insight_cat, action, score, prio = get_professor_diagnosis(r, stats)
        r["insight_category"] = insight_cat
        r["professor_action"] = action
        r["priority_score"] = score
        r["priority"] = prio
        
        stock = r.get("stock_qty") or 0
        reorder = r.get("reorder_point_qty") or 0
        gap = max(0, reorder - stock) if pd.notna(reorder) and pd.notna(stock) else 0
        daily = (r.get("transactions") or 0) / 7.0 if pd.notna(r.get("transactions")) else 0
        r["suggested_replenishment_qty"] = int(round(gap + (daily * 14)))
        r["stock_gap_qty"] = gap
        
        r["reason"] = f"Teşhis: {insight_cat}"
        r["recommended_action"] = action
        results.append(r)

    final_df = pd.DataFrame(results)
    
    sort_col = "revenue" if "revenue" in final_df.columns else "pdp_views"
    if sort_col not in final_df.columns:
        sort_col = "priority_score"
        
    final_df = final_df.sort_values(
        ["priority_score", sort_col],
        ascending=[False, False]
    )

    rows = records(final_df.head(30))

    return json.dumps(
        {
            "analysis_type": "recommended_action_execution",
            "action_type": "comprehensive_professor_plan",
            "question": question,
            "summary": {
                "selected_sku_count": int(len(final_df)),
                "critical_count": int((final_df["priority"] == "Critical").sum()),
                "high_count": int((final_df["priority"] == "High").sum()),
            },
            "main_result": "E-Ticaret Profesörü teşhis motoru çalıştırıldı. Her SKU için Fiyat, Stok ve Dönüşüm metrikleri çaprazlanarak nokta atışı aksiyonlar üretildi.",
            "rows": rows,
            "recommended_actions": [
                "Kritik Stok (OOS Risk) grubundaki ürünler için acil tedarik emri oluşturulmalı.",
                "Fiyat Bariyeri teşhisi konan SKU'larda merchant benchmark verisine göre fiyat optimizasyonu kurgulanmalı.",
                "Gizli Cevher ürünlerine hemen CRM veya performance marketing bütçesi kaydırılmalı.",
                "Excel'i indirip her ürün için özel üretilen profesör aksiyonlarını ekiplere dağıtın."
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

    # Use the comprehensive professor engine for all action inquiries
    return build_comprehensive_action_plan(df, question)


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
        "insight_category",
        "professor_action",
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