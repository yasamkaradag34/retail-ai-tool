import json
import os
import re
from functools import lru_cache

import duckdb
import pandas as pd
import requests


DATA_PATH = "data/ecommerce_ai_sample_data_200_rows.xlsx"
SHEET_NAME = "sample_data_200"

OLLAMA_URL = "http://localhost:11434/api/chat"
MODEL = "llama3.1"


@lru_cache(maxsize=1)
def load_sample_data():
    if not os.path.exists(DATA_PATH):
        raise FileNotFoundError(
            f"Sample data bulunamadı. Beklenen yol: {DATA_PATH}"
        )

    df = pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME)

    # Kolon isimlerini SQL için daha güvenli hale getiriyoruz
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


def analyze_ecommerce_sample(question: str):
    """
    200 satırlık e-commerce sample data üzerinde kullanıcının sorusunu cevaplar.
    Önce bazı kritik intent'leri deterministic olarak çözer.
    Diğer sorularda Ollama SQL üretir, DuckDB Excel datası üzerinde çalıştırır.
    """

    try:
        df = load_sample_data()
        print(
            "✅ SAMPLE DATA OKUNDU:",
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
                "error": "Sample data okunamadı.",
                "detail": str(e),
                "expected_path": DATA_PATH,
            },
            ensure_ascii=False,
        )

    q = question.lower()

    # ─────────────────────────────────────────────────────────────
    # 1) C2D artan fakat B2D azalan ürünler
    # Kullanıcı sorusu örneği:
    # "C2D artan fakat B2D azalan ürünlerin sadece SKU'larını listeler misin?"
    # ─────────────────────────────────────────────────────────────
    if (
        "c2d" in q
        and "b2d" in q
        and any(x in q for x in ["artan", "artmış", "artiyor", "artıyor", "yükselen", "yukselen", "pozitif"])
        and any(x in q for x in ["azalan", "azalmış", "azaliyor", "azalıyor", "düşen", "dusen", "düşmüş", "dusmus", "negatif"])
    ):
        required_cols = ["sku", "c2d_delta_pct", "b2d_delta_pct"]

        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return json.dumps(
                {
                    "error": "Gerekli kolonlar bulunamadı.",
                    "missing_columns": missing,
                    "available_columns": list(df.columns),
                },
                ensure_ascii=False,
            )

        result_df = df[
            (df["c2d_delta_pct"] > 0)
            & (df["b2d_delta_pct"] < 0)
        ].copy()

        result_df = result_df.sort_values(
            ["c2d_delta_pct", "b2d_delta_pct"],
            ascending=[False, True],
        )

        # Kullanıcı sadece SKU istediyse sadece SKU döndür
        if any(x in q for x in ["sadece sku", "sku'larını", "sku larını", "sku listesini", "sku listesi"]):
            output_df = result_df[["sku"]].drop_duplicates()
        else:
            wanted_cols = [
                "sku",
                "brand",
                "cat1",
                "cat2",
                "product",
                "c2d_pct",
                "c2d_delta_pct",
                "b2d_pct",
                "b2d_delta_pct",
                "stock_qty",
                "reorder_point_qty",
                "stock_risk_level",
                "product_price",
                "product_price_delta_pct",
                "revenue",
                "revenue_delta_pct",
            ]
            existing_cols = [c for c in wanted_cols if c in result_df.columns]
            output_df = result_df[existing_cols]

        output_df = output_df.where(pd.notnull(output_df), None)

        return json.dumps(
            {
                "question": question,
                "analysis_type": "c2d_up_b2d_down",
                "logic": "c2d_delta_pct > 0 AND b2d_delta_pct < 0",
                "row_count": len(output_df),
                "rows": output_df.head(30).to_dict(orient="records"),
                "action_recommendation": (
                    "Bu ürünlerde sepete ekleme eğilimi artmış ama satın alma dönüşümü düşmüş. "
                    "Öncelik fiyat rekabeti, stok uygunluğu, kargo/ödeme adımı, ürün sayfası güven unsurları "
                    "ve checkout hataları kontrolü olmalı."
                ),
            },
            ensure_ascii=False,
            default=str,
        )

    # ─────────────────────────────────────────────────────────────
    # 2) C2D yüksek ama stoğu az SKU'lar
    # Kullanıcı sorusu örneği:
    # "C2D yüksek ama stoğu az SKU'lar hangileri?"
    # ─────────────────────────────────────────────────────────────
    if (
        "c2d" in q
        and any(x in q for x in ["stok", "stoğu", "stogu"])
        and any(x in q for x in ["az", "düşük", "dusuk", "kritik", "risk"])
    ):
        required_cols = ["sku", "c2d_pct", "stock_qty", "reorder_point_qty"]

        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return json.dumps(
                {
                    "error": "Gerekli kolonlar bulunamadı.",
                    "missing_columns": missing,
                    "available_columns": list(df.columns),
                },
                ensure_ascii=False,
            )

        result_df = df[
            df["stock_qty"] <= df["reorder_point_qty"]
        ].copy()

        result_df = result_df.sort_values(
            ["c2d_pct", "stock_qty"],
            ascending=[False, True],
        )

        wanted_cols = [
            "sku",
            "brand",
            "cat1",
            "cat2",
            "product",
            "c2d_pct",
            "c2d_delta_pct",
            "b2d_pct",
            "b2d_delta_pct",
            "stock_qty",
            "reorder_point_qty",
            "stock_risk_level",
            "availability_status",
            "revenue",
        ]
        existing_cols = [c for c in wanted_cols if c in result_df.columns]
        output_df = result_df[existing_cols]
        output_df = output_df.where(pd.notnull(output_df), None)

        return json.dumps(
            {
                "question": question,
                "analysis_type": "high_c2d_low_stock",
                "logic": "stock_qty <= reorder_point_qty ORDER BY c2d_pct DESC",
                "row_count": len(output_df),
                "rows": output_df.head(30).to_dict(orient="records"),
                "action_recommendation": (
                    "Bu SKU'larda kullanıcı ilgisi güçlü ama stok seviyesi kritik. "
                    "Öncelik replenishment, incoming stock kontrolü ve stok bitmeden kampanya/traffic optimizasyonu olmalı."
                ),
            },
            ensure_ascii=False,
            default=str,
        )

    # ─────────────────────────────────────────────────────────────
    # 3) B2D yüksek ama stoğu az SKU'lar
    # ─────────────────────────────────────────────────────────────
    if (
        "b2d" in q
        and any(x in q for x in ["stok", "stoğu", "stogu"])
        and any(x in q for x in ["az", "düşük", "dusuk", "kritik", "risk"])
    ):
        required_cols = ["sku", "b2d_pct", "stock_qty", "reorder_point_qty"]

        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return json.dumps(
                {
                    "error": "Gerekli kolonlar bulunamadı.",
                    "missing_columns": missing,
                    "available_columns": list(df.columns),
                },
                ensure_ascii=False,
            )

        result_df = df[
            df["stock_qty"] <= df["reorder_point_qty"]
        ].copy()

        result_df = result_df.sort_values(
            ["b2d_pct", "stock_qty"],
            ascending=[False, True],
        )

        wanted_cols = [
            "sku",
            "brand",
            "cat1",
            "cat2",
            "product",
            "b2d_pct",
            "b2d_delta_pct",
            "c2d_pct",
            "c2d_delta_pct",
            "stock_qty",
            "reorder_point_qty",
            "stock_risk_level",
            "availability_status",
            "revenue",
        ]
        existing_cols = [c for c in wanted_cols if c in result_df.columns]
        output_df = result_df[existing_cols]
        output_df = output_df.where(pd.notnull(output_df), None)

        return json.dumps(
            {
                "question": question,
                "analysis_type": "high_b2d_low_stock",
                "logic": "stock_qty <= reorder_point_qty ORDER BY b2d_pct DESC",
                "row_count": len(output_df),
                "rows": output_df.head(30).to_dict(orient="records"),
                "action_recommendation": (
                    "Bu SKU'larda satın alma dönüşümü güçlü ama stok seviyesi kritik. "
                    "Stok biterse direkt revenue kaybı oluşabilir; acil replenishment önerilir."
                ),
            },
            ensure_ascii=False,
            default=str,
        )

    # ─────────────────────────────────────────────────────────────
    # 4) OOS ama PDP View alan ürünler
    # ─────────────────────────────────────────────────────────────
    if (
        any(x in q for x in ["oos", "out of stock", "stokta olmayan", "stok yok", "stoğu biten", "stogu biten"])
        and any(x in q for x in ["pdp", "view", "görüntü", "goruntu"])
    ):
        required_cols = ["sku", "stock_qty", "total_unique_pdp_views_sum"]

        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return json.dumps(
                {
                    "error": "Gerekli kolonlar bulunamadı.",
                    "missing_columns": missing,
                    "available_columns": list(df.columns),
                },
                ensure_ascii=False,
            )

        if "availability_status" in df.columns:
            result_df = df[
                ((df["stock_qty"] == 0) | (df["availability_status"] == "out_of_stock"))
                & (df["total_unique_pdp_views_sum"] > 0)
            ].copy()
        else:
            result_df = df[
                (df["stock_qty"] == 0)
                & (df["total_unique_pdp_views_sum"] > 0)
            ].copy()

        result_df = result_df.sort_values(
            "total_unique_pdp_views_sum",
            ascending=False,
        )

        wanted_cols = [
            "sku",
            "brand",
            "cat1",
            "cat2",
            "product",
            "stock_qty",
            "availability_status",
            "total_unique_pdp_views_sum",
            "c2d_pct",
            "b2d_pct",
            "revenue",
            "estimated_lost_revenue",
        ]
        existing_cols = [c for c in wanted_cols if c in result_df.columns]
        output_df = result_df[existing_cols]
        output_df = output_df.where(pd.notnull(output_df), None)

        return json.dumps(
            {
                "question": question,
                "analysis_type": "oos_products_with_pdp_views",
                "logic": "stock_qty = 0 AND total_unique_pdp_views_sum > 0",
                "row_count": len(output_df),
                "rows": output_df.head(30).to_dict(orient="records"),
                "action_recommendation": (
                    "Bu ürünler stokta yokken kullanıcı ilgisi almaya devam ediyor. "
                    "Replenishment, alternatif ürün yönlendirmesi ve reklam trafiği kontrol edilmeli."
                ),
            },
            ensure_ascii=False,
            default=str,
        )

    # ─────────────────────────────────────────────────────────────
    # 5) Genel sorular için Ollama → SQL → DuckDB akışı
    # ─────────────────────────────────────────────────────────────
    schema_text = "\n".join(
        [f"- {col}: {str(df[col].dtype)}" for col in df.columns]
    )

    sql_prompt = f"""
Sen sadece DuckDB SQL üreten bir e-ticaret veri analistisin.

Tablo adı: ecommerce_sample

Kolonlar:
{schema_text}

İş kuralları:
- Sadece SELECT sorgusu üret.
- Açıklama yazma.
- Markdown kullanma.
- SQL dışında hiçbir şey döndürme.
- LIMIT kullan. Varsayılan LIMIT 10.
- C2D = total_unique_add_to_carts_sum / total_unique_pdp_views_sum * 100
- B2D = total_transactions_sum / total_unique_pdp_views_sum * 100
- Delta / pct_delta kolonları önceki dönem değişimidir.
- C2D artan demek: c2d_delta_pct > 0
- C2D azalan demek: c2d_delta_pct < 0
- B2D artan demek: b2d_delta_pct > 0
- B2D azalan demek: b2d_delta_pct < 0
- Stock coverage = stock_qty / daily_sales_qty_7d
- Eğer stock_coverage_days kolonu varsa onu kullan.
- OOS = stock_qty = 0 veya availability_status = 'out_of_stock'
- Kritik stok = stock_qty <= reorder_point_qty
- Overstock = stock_qty yüksek ama satış/dönüşüm düşük ürünler
- Funnel sırası:
  PDP View → A2C → Cart View → Shipping View → Payment View → Summary View → Checkout Submit → Transactions

Önemli:
- Kullanıcı "C2D artan" diyorsa mutlaka c2d_delta_pct kolonunu kullan.
- Kullanıcı "B2D azalan" diyorsa mutlaka b2d_delta_pct kolonunu kullan.
- Kullanıcı "sadece SKU" diyorsa sadece sku kolonunu SELECT et.
- Uydurma SKU yazma. Sadece tablodan gelen sku değerlerini döndür.

Kullanıcı sorusu:
{question}
"""

    try:
        response = requests.post(
            OLLAMA_URL,
            json={
                "model": MODEL,
                "messages": [
                    {"role": "system", "content": sql_prompt},
                    {"role": "user", "content": question},
                ],
                "stream": False,
            },
            timeout=120,
        )

        result = response.json()
        sql = result.get("message", {}).get("content", "").strip()
        sql = clean_sql(sql)

        print("🧠 GENERATED SQL:", sql, flush=True)

    except Exception as e:
        return json.dumps(
            {
                "error": "Ollama SQL üretirken hata aldı.",
                "detail": str(e),
            },
            ensure_ascii=False,
        )

    if not is_safe_select(sql):
        return json.dumps(
            {
                "error": "Güvenli SELECT sorgusu üretilemedi.",
                "generated_sql": sql,
                "hint": "Model SQL dışında metin üretmiş olabilir.",
            },
            ensure_ascii=False,
        )

    try:
        con = duckdb.connect()
        con.register("ecommerce_sample", df)
        result_df = con.execute(sql).df()
    except Exception as e:
        return json.dumps(
            {
                "error": "SQL çalıştırılırken hata oluştu.",
                "detail": str(e),
                "generated_sql": sql,
                "available_columns": list(df.columns),
            },
            ensure_ascii=False,
        )

    result_df = result_df.where(pd.notnull(result_df), None)

    return json.dumps(
        {
            "question": question,
            "generated_sql": sql,
            "row_count": len(result_df),
            "rows": result_df.head(30).to_dict(orient="records"),
        },
        ensure_ascii=False,
        default=str,
    )


def clean_sql(sql: str) -> str:
    sql = sql.replace("```sql", "").replace("```", "").strip()

    # Model bazen başına/sonuna açıklama koyarsa SELECT kısmını ayıkla
    match = re.search(r"(select\s+.*)", sql, flags=re.IGNORECASE | re.DOTALL)
    if match:
        sql = match.group(1).strip()

    # Noktalı virgülden sonrasını at
    if ";" in sql:
        sql = sql.split(";")[0] + ";"

    return sql


def is_safe_select(sql: str) -> bool:
    sql_lower = sql.lower().strip()

    if not sql_lower.startswith("select"):
        return False

    forbidden_words = [
        "insert",
        "update",
        "delete",
        "drop",
        "alter",
        "create",
        "truncate",
        "attach",
        "copy",
        "pragma",
    ]

    return not any(word in sql_lower for word in forbidden_words)