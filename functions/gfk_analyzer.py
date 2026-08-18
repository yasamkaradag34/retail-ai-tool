"""
GfK Analyzer Module
===================
GfK_Leaderpanel.xlsx ve gfk_sku.xlsx verilerini parse ederek pazar payı,
marka performansı, SKU sıralaması ve ecommerce metrikleriyle birleşik
analizler üretir.

Veri Kaynakları:
  - data/GfK_Leaderpanel.xlsx  → 15 sheet (Summary_value, Brand, PW vs. CW vb.)
  - data/gfk_sku.xlsx          → 32,885 satır SKU sıralaması
"""

import json
import os
from functools import lru_cache
from typing import Optional, List

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# Dosya Yolları
# ─────────────────────────────────────────────────────────────────────────────

GFK_LEADERPANEL_PATH = "data/GfK_Leaderpanel.xlsx"
GFK_SKU_PATH = "data/gfk_sku.xlsx"
ECOMMERCE_PATH = "data/ecommerce_ai_sample_data_200_rows.xlsx"

# ─────────────────────────────────────────────────────────────────────────────
# GfK Kategori ↔ İç Kategori Mapping (brand adı üzerinden desteklenir)
# ─────────────────────────────────────────────────────────────────────────────

GFK_CATEGORY_MAP = {
    "Smartphones": ["Cep Telefonları", "Telefon", "Akıllı Telefon", "smartphone", "gsm"],
    "COMPUTER HW": ["Bilgisayar", "Laptop", "Dizüstü", "Masaüstü", "PC", "computer hw"],
    "SDA": ["Küçük Ev Aletleri", "Kişisel Bakım", "sda"],
    "MDA": ["Büyük Ev Aletleri", "Çamaşır Makinesi", "Bulaşık Makinesi", "mda"],
    "CLIMATE SDA": ["İklim", "Klima", "Vantilatör", "climate"],
    "PTV/FLAT": ["Televizyon", "TV", "OLED", "QLED", "ptv"],
    "Headphones & Headsets": ["Kulaklık", "Headphone", "Headset", "kulaklık"],
    "COMPUTER ACCESSORIES": ["Bilgisayar Aksesuarları", "Mouse", "Klavye", "Monitör"],
    "CORE WEARABLES": ["Giyilebilir", "Akıllı Saat", "Smartwatch", "Wearable"],
    "VACUUM CLEANERS": ["Süpürge", "Robot Süpürge", "Dikey Süpürge"],
    "WASHING MACHINES": ["Çamaşır Makinesi"],
    "TUMBLE DRYERS": ["Çamaşır Kurutma"],
    "DISHWASHERS": ["Bulaşık Makinesi"],
    "COOLING": ["Buzdolabı", "Derin Dondurucu"],
    "MONITORS": ["Monitör"],
    "MOBILE COMPUTING": ["Tablet", "iPad", "Laptop", "Notebook"],
    "MEDIATABLETS": ["Tablet", "iPad"],
    "AIR CONDITIONERS": ["Klima", "Split Klima"],
    "HOT BEVER.MAKERS": ["Kahve Makinesi", "Kettle"],
    "HAIR DRYERS": ["Saç Kurutma Makinesi"],
    "HAIR STYLERS": ["Saç Şekillendirici"],
    "SHAVERS": ["Tıraş Makinesi", "Epilasyon"],
}

# Ters mapping: iç kategori → GfK kategori
INTERNAL_TO_GFK = {}
for gfk_cat, internal_list in GFK_CATEGORY_MAP.items():
    for internal in internal_list:
        INTERNAL_TO_GFK[internal.lower()] = gfk_cat


# ─────────────────────────────────────────────────────────────────────────────
# Veri Yükleme (Cached)
# ─────────────────────────────────────────────────────────────────────────────

@lru_cache(maxsize=1)
def load_gfk_summary_value():
    """Summary_value sheet'ini parse eder: kategori × hafta × (pazar + MM değer)"""
    if not os.path.exists(GFK_LEADERPANEL_PATH):
        return None, f"GfK Leaderpanel dosyası bulunamadı: {GFK_LEADERPANEL_PATH}"

    try:
        df = pd.read_excel(GFK_LEADERPANEL_PATH, sheet_name="Summary_value", header=3)
        df.columns = [str(c) for c in df.columns]
        df = df.rename(columns={
            "Unnamed: 0": "product_group",
            "Unnamed: 1": "product_group2",
            "Unnamed: 2": "metric"
        })
        # Boş satırları temizle
        df = df[df["metric"].notna() & df["metric"].str.startswith("Sum", na=False)]
        # product_group forward fill
        df["product_group"] = df["product_group"].ffill()
        # Metric tiplerini basitleştir
        df["metric_type"] = df["metric"].apply(
            lambda x: "mediamarkt_sales_try" if "MediaMarkt" in str(x) else "total_internet_sales_try"
        )
        return df, None
    except Exception as e:
        return None, str(e)


@lru_cache(maxsize=1)
def load_gfk_brand():
    """Brand sheet'ini parse eder: kategori × marka × haftalık MM satış payı (%)"""
    if not os.path.exists(GFK_LEADERPANEL_PATH):
        return None, f"GfK Leaderpanel dosyası bulunamadı: {GFK_LEADERPANEL_PATH}"

    try:
        df = pd.read_excel(GFK_LEADERPANEL_PATH, sheet_name="Brand", header=2)
        df.columns = [str(c) for c in df.columns]
        df = df.rename(columns={
            "Unnamed: 0": "product_group",
            "Unnamed: 1": "product_group2",
            "Unnamed: 2": "brand"
        })
        # Header satırını temizle
        df = df[df["brand"].notna() & (df["brand"] != "Brand")]
        # product_group forward fill
        df["product_group"] = df["product_group"].ffill()
        # Hafta kolon isimlerini bul (Week xx xxxx formatında)
        week_cols = [c for c in df.columns if "Week" in str(c) or "week" in str(c)]
        # Sayısal hafta kolon isimleri (17, 18, ... gibi)
        numeric_week_cols = [c for c in df.columns if str(c).isdigit()]
        all_week_cols = week_cols + numeric_week_cols
        return df, None
    except Exception as e:
        return None, str(e)


@lru_cache(maxsize=1)
def load_gfk_pw_cw():
    """PW vs. CW sheet'ini parse eder: geçen hafta vs bu hafta karşılaştırması"""
    if not os.path.exists(GFK_LEADERPANEL_PATH):
        return None, f"GfK Leaderpanel dosyası bulunamadı: {GFK_LEADERPANEL_PATH}"

    try:
        df = pd.read_excel(GFK_LEADERPANEL_PATH, sheet_name="PW vs. CW", header=2)
        df.columns = [str(c) for c in df.columns]
        df = df.rename(columns={"Unnamed: 0": "product_group"})
        # Header satırını temizle
        df = df[df["product_group"].notna() & (df["product_group"] != "ReportingProductgroup")]

        # Kolon isimlerini standartlaştır
        cols = list(df.columns)
        rename_map = {}
        for i, c in enumerate(cols):
            if c == "Unnamed: 3":
                rename_map[c] = "gap_placeholder"
            elif "Unnamed" in c:
                rename_map[c] = f"col_{i}"

        df = df.rename(columns=rename_map)

        # Sayısal değer kolonlarını dönüştür
        for col in df.columns:
            if col != "product_group":
                df[col] = pd.to_numeric(df[col], errors="coerce")

        return df, None
    except Exception as e:
        return None, str(e)


@lru_cache(maxsize=1)
def load_gfk_sku():
    """gfk_sku.xlsx'i yükler: ürün grubu × rank × marka × model"""
    if not os.path.exists(GFK_SKU_PATH):
        return None, f"GfK SKU dosyası bulunamadı: {GFK_SKU_PATH}"

    try:
        df = pd.read_excel(GFK_SKU_PATH, sheet_name="Sheet1")
        df.columns = [
            str(c).strip().lower().replace(" ", "_") for c in df.columns
        ]
        return df, None
    except Exception as e:
        return None, str(e)


@lru_cache(maxsize=1)
def load_ecommerce():
    """Ecommerce sample data'yı yükler"""
    if not os.path.exists(ECOMMERCE_PATH):
        return None, f"Ecommerce data bulunamadı: {ECOMMERCE_PATH}"
    try:
        df = pd.read_excel(ECOMMERCE_PATH, sheet_name="sample_data_200")
        df.columns = [
            str(c).strip().replace(" ", "_").replace("%", "pct").replace("/", "_")
            .replace("-", "_").replace("(", "").replace(")", "").replace(".", "").lower()
            for c in df.columns
        ]
        return df, None
    except Exception as e:
        return None, str(e)


# ─────────────────────────────────────────────────────────────────────────────
# Yardımcı Fonksiyonlar
# ─────────────────────────────────────────────────────────────────────────────

def fmt_try(value):
    """TRY değerini milyar/milyon cinsinden formatla"""
    try:
        v = float(value)
        if abs(v) >= 1_000_000_000:
            return f"{v/1_000_000_000:.2f} Milyar TRY"
        elif abs(v) >= 1_000_000:
            return f"{v/1_000_000:.1f} Milyon TRY"
        elif abs(v) >= 1_000:
            return f"{v/1_000:.1f} Bin TRY"
        return f"{v:,.0f} TRY"
    except Exception:
        return str(value)


def fmt_pct(value):
    try:
        return f"%{float(value)*100:.1f}"
    except Exception:
        return str(value)


def find_category_in_question(question: str, available_categories: list) -> Optional[str]:
    """Soru metninde GfK kategori adını veya eşlenik kategori adını arar"""
    q = question.lower()

    # Doğrudan GfK kategori eşleşmesi
    for cat in available_categories:
        if cat.lower() in q:
            return cat

    # İç kategori üzerinden GfK kategoriye eşle
    for internal, gfk_cat in INTERNAL_TO_GFK.items():
        if internal in q and gfk_cat in available_categories:
            return gfk_cat

    return None


def find_brand_in_question(question: str, available_brands: list) -> Optional[str]:
    """Soru metninde marka adını arar"""
    q = question.lower()
    for brand in available_brands:
        if brand.lower() in q:
            return brand
    return None


# ─────────────────────────────────────────────────────────────────────────────
# 1. Pazar Payı Analizi
# ─────────────────────────────────────────────────────────────────────────────

def analyze_gfk_market_share(question: str) -> str:
    """
    GfK Leaderpanel Summary_value ve PW vs. CW verilerinden pazar payı,
    kategori büyümesi ve MediaMarkt payı analizi yapar.

    Desteklenen soru tipleri:
    - "MediaMarkt'ın pazar payı nedir?"
    - "En çok büyüyen kategori hangisi?"
    - "Bu hafta vs geçen hafta satış değişimi nasıl?"
    - "AIR CONDITIONERS'da pazar durumu nedir?"
    - "Hangi kategoride en büyük satış değişimi var?"
    """
    q = question.lower()

    # PW vs CW — haftalık değişim
    df_pw, err = load_gfk_pw_cw()
    if err:
        return json.dumps({"error": err, "analysis_type": "gfk_market_share_error"}, ensure_ascii=False)

    df_summary, err2 = load_gfk_summary_value()
    if err2:
        return json.dumps({"error": err2, "analysis_type": "gfk_market_share_error"}, ensure_ascii=False)

    available_cats = list(df_pw["product_group"].dropna().unique())

    # Kolon tespiti — Orijinal Excel yapısı:
    # col 0: product_group | col 1: PW value | col 2: CW value | col 3: boşluk
    # col 4: WoW % | col 5: WoW abs val | col 6: Ihs | col 7: Rank
    # Yani col_4=WoW%, col_5=abs, col_6=Ihs, col_7=rank
    cols = list(df_pw.columns)
    numeric_cols = [c for c in cols if c not in ["product_group", "gap_placeholder"] and not str(c).startswith("col_") or str(c).startswith("col_")]
    
    # Sayısal değer kolonlarını pozisyona göre tespit et
    value_cols = [c for c in cols if str(c).replace(".", "").isdigit() and "." not in str(c)]  # "23", "24"
    pct_change_col = "col_4" if "col_4" in cols else None   # WoW %
    abs_change_col = "col_5" if "col_5" in cols else None   # WoW abs TRY
    ihs_col = "col_6" if "col_6" in cols else None          # MediaMarkt Ihs (pazar payı)
    rank_col = "col_7" if "col_7" in cols else None         # Rank

    # Kategori filtresi
    target_cat = find_category_in_question(question, available_cats)

    if target_cat:
        row = df_pw[df_pw["product_group"] == target_cat]
        if row.empty:
            return json.dumps({
                "analysis_type": "gfk_market_share",
                "error": f"{target_cat} kategorisi PW vs CW verisinde bulunamadı.",
                "available_categories": available_cats
            }, ensure_ascii=False)

        r = row.iloc[0]
        result = {
            "analysis_type": "gfk_market_share",
            "category": target_cat,
            "source": "GfK Leaderpanel — PW vs. CW",
        }

        if value_cols and len(value_cols) >= 2:
            pw = r.get(value_cols[0])
            cw = r.get(value_cols[1])
            result["previous_week_value_try"] = fmt_try(pw)
            result["current_week_value_try"] = fmt_try(cw)
            result["week_labels"] = [str(value_cols[0]), str(value_cols[1])]

        if pct_change_col:
            result["wow_change_pct"] = f"%{float(r.get(pct_change_col, 0)) * 100:.1f}" if pd.notna(r.get(pct_change_col)) else "N/A"
        if abs_change_col:
            result["wow_change_abs"] = fmt_try(r.get(abs_change_col)) if pd.notna(r.get(abs_change_col)) else "N/A"
        if ihs_col:
            result["mediamarkt_market_share_pct"] = fmt_pct(r.get(ihs_col)) if pd.notna(r.get(ihs_col)) else "N/A"
        if rank_col:
            result["mediamarkt_rank"] = int(r.get(rank_col)) if pd.notna(r.get(rank_col)) else "N/A"

        return json.dumps(result, ensure_ascii=False, default=str)

    # Genel — tüm kategorileri sırala
    rows = []
    for _, r in df_pw.iterrows():
        cat = r.get("product_group")
        if not cat or str(cat) in ["nan", "ReportingProductgroup"]:
            continue

        row_data = {"category": str(cat)}

        if value_cols and len(value_cols) >= 2:
            try:
                row_data["current_week_try"] = fmt_try(r.get(value_cols[1]))
            except Exception:
                pass

        if pct_change_col and pd.notna(r.get(pct_change_col)):
            try:
                pct = float(r.get(pct_change_col)) * 100
                row_data["wow_change_pct"] = round(pct, 1)
            except Exception:
                pass

        if ihs_col and pd.notna(r.get(ihs_col)):
            try:
                row_data["mediamarkt_share_pct"] = round(float(r.get(ihs_col)) * 100, 1)
            except Exception:
                pass

        if rank_col and pd.notna(r.get(rank_col)):
            try:
                row_data["mediamarkt_rank"] = int(r.get(rank_col))
            except Exception:
                pass

        rows.append(row_data)

    # Büyüme sıralaması veya genel analiz
    if any(x in q for x in ["büyüyen", "buyuyen", "artış", "artis", "artmış", "en çok"]):
        rows_sorted = sorted(
            [r for r in rows if "wow_change_pct" in r],
            key=lambda x: x.get("wow_change_pct", 0),
            reverse=True
        )
        return json.dumps({
            "analysis_type": "gfk_market_share",
            "view": "top_growing_categories",
            "source": "GfK Leaderpanel — PW vs. CW",
            "rows": rows_sorted[:10],
            "total_categories": len(rows)
        }, ensure_ascii=False, default=str)

    if any(x in q for x in ["düşen", "dusen", "azalan", "kaybeden"]):
        rows_sorted = sorted(
            [r for r in rows if "wow_change_pct" in r],
            key=lambda x: x.get("wow_change_pct", 0),
            reverse=False
        )
        return json.dumps({
            "analysis_type": "gfk_market_share",
            "view": "top_declining_categories",
            "source": "GfK Leaderpanel — PW vs. CW",
            "rows": rows_sorted[:10],
            "total_categories": len(rows)
        }, ensure_ascii=False, default=str)

    if any(x in q for x in ["pazar payı", "pazar payi", "market share", "ihs", "sıra", "rank"]):
        rows_sorted = sorted(
            [r for r in rows if "mediamarkt_rank" in r],
            key=lambda x: x.get("mediamarkt_rank", 999)
        )
        return json.dumps({
            "analysis_type": "gfk_market_share",
            "view": "mediamarkt_market_share_ranking",
            "source": "GfK Leaderpanel — PW vs. CW",
            "rows": rows_sorted[:15],
            "total_categories": len(rows)
        }, ensure_ascii=False, default=str)

    # Varsayılan — tüm genel özet
    return json.dumps({
        "analysis_type": "gfk_market_share",
        "view": "all_categories_overview",
        "source": "GfK Leaderpanel — PW vs. CW",
        "rows": rows[:20],
        "total_categories": len(rows)
    }, ensure_ascii=False, default=str)


# ─────────────────────────────────────────────────────────────────────────────
# 2. Marka Performansı Analizi
# ─────────────────────────────────────────────────────────────────────────────

def analyze_gfk_brand_performance(question: str) -> str:
    """
    GfK Brand sheet'inden marka bazında haftalık MediaMarkt internet satış payı (%)
    ve değişim analizleri üretir.

    Desteklenen soru tipleri:
    - "SAMSUNG bu hafta pazar payı ne kadar?"
    - "APPLE vs SAMSUNG karşılaştır"
    - "Smartphones'da hangi marka en yüksek payda?"
    - "MediaMarkt'ta en güçlü marka hangisi?"
    - "BOSCH çamaşır makinesinde kaçıncı sırada?"
    """
    q = question.lower()

    df, err = load_gfk_brand()
    if err:
        return json.dumps({"error": err, "analysis_type": "gfk_brand_error"}, ensure_ascii=False)

    available_cats = list(df["product_group"].dropna().unique())
    available_brands = list(df["brand"].dropna().unique())

    # Hafta kolonlarını tespit et
    week_cols = [c for c in df.columns if "Week" in str(c)]
    if not week_cols:
        # Sayısal kolonları dene
        week_cols = [c for c in df.columns if str(c).replace(".", "").isdigit()]

    latest_week = week_cols[-1] if week_cols else None
    prev_week = week_cols[-2] if len(week_cols) >= 2 else None

    target_cat = find_category_in_question(question, available_cats)
    target_brand = find_brand_in_question(question, available_brands)

    # Filtrele
    filtered = df.copy()
    if target_cat:
        filtered = filtered[filtered["product_group"] == target_cat]
    if target_brand:
        filtered = filtered[filtered["brand"] == target_brand]

    if filtered.empty:
        return json.dumps({
            "analysis_type": "gfk_brand_performance",
            "error": "Belirtilen kategori veya marka bulunamadı.",
            "available_categories": available_cats[:15],
            "available_brands": available_brands[:20]
        }, ensure_ascii=False)

    rows = []
    for _, r in filtered.iterrows():
        brand = r.get("brand", "N/A")
        cat = r.get("product_group", "N/A")
        row_data = {"brand": str(brand), "product_group": str(cat)}

        if latest_week and pd.notna(r.get(latest_week)):
            try:
                share = float(r.get(latest_week)) * 100
                row_data["latest_week_share_pct"] = round(share, 2)
                row_data["latest_week_label"] = str(latest_week)
            except Exception:
                pass

        if prev_week and pd.notna(r.get(prev_week)):
            try:
                prev_share = float(r.get(prev_week)) * 100
                row_data["prev_week_share_pct"] = round(prev_share, 2)
                if "latest_week_share_pct" in row_data:
                    row_data["wow_change_pp"] = round(
                        row_data["latest_week_share_pct"] - prev_share, 2
                    )
            except Exception:
                pass

        # Tüm haftalık seriyi ekle
        if week_cols:
            weekly_series = {}
            for wc in week_cols[-8:]:  # Son 8 hafta
                val = r.get(wc)
                if pd.notna(val):
                    try:
                        weekly_series[str(wc)] = round(float(val) * 100, 2)
                    except Exception:
                        pass
            if weekly_series:
                row_data["weekly_share_series"] = weekly_series

        rows.append(row_data)

    # Sıralama
    rows_sorted = sorted(
        [r for r in rows if "latest_week_share_pct" in r],
        key=lambda x: x.get("latest_week_share_pct", 0),
        reverse=True
    )

    summary = {
        "analysis_type": "gfk_brand_performance",
        "source": "GfK Leaderpanel — Brand Sheet",
        "filtered_by_category": target_cat,
        "filtered_by_brand": target_brand,
        "latest_week": str(latest_week) if latest_week else None,
        "prev_week": str(prev_week) if prev_week else None,
        "brand_count": len(rows_sorted),
        "rows": rows_sorted[:20]
    }

    if rows_sorted:
        top = rows_sorted[0]
        summary["top_brand"] = top.get("brand")
        summary["top_brand_share_pct"] = top.get("latest_week_share_pct")

    return json.dumps(summary, ensure_ascii=False, default=str)


# ─────────────────────────────────────────────────────────────────────────────
# 3. SKU Sıralaması Analizi
# ─────────────────────────────────────────────────────────────────────────────

def analyze_gfk_sku_ranking(question: str) -> str:
    """
    gfk_sku.xlsx'ten ürün grubu, marka ve rank bazında SKU listesi üretir.

    Desteklenen soru tipleri:
    - "Washing Machines'de top 10 SKU hangisi?"
    - "SAMSUNG'un en çok satan modelleri neler?"
    - "GfK'ya göre 1. sıradaki ürünler hangileri?"
    - "Smartphone'larda APPLE kaçıncı sırada?"
    - "LG bulaşık makinesinde hangi model 1. sırada?"
    """
    q = question.lower()

    df, err = load_gfk_sku()
    if err:
        return json.dumps({"error": err, "analysis_type": "gfk_sku_error"}, ensure_ascii=False)

    available_groups = list(df["reportingproductgroup"].dropna().unique())
    available_brands = list(df["brand"].dropna().unique())

    # Filtreler
    target_group = None
    for g in available_groups:
        if g.lower() in q:
            target_group = g
            break

    # Kategori mapping üzerinden de ara
    if not target_group:
        for gfk_cat, internals in GFK_CATEGORY_MAP.items():
            if any(i.lower() in q for i in internals):
                if gfk_cat in available_groups:
                    target_group = gfk_cat
                    break

    target_brand = find_brand_in_question(question, available_brands)

    # Rank filtresi
    rank_limit = 10  # varsayılan top 10
    for token in q.split():
        if token.isdigit():
            rank_limit = int(token)
            break
    if any(x in q for x in ["top 5", "ilk 5", "5 sku"]):
        rank_limit = 5
    elif any(x in q for x in ["top 20", "ilk 20", "20 sku"]):
        rank_limit = 20

    filtered = df.copy()
    if target_group:
        filtered = filtered[filtered["reportingproductgroup"] == target_group]
    if target_brand:
        filtered = filtered[filtered["brand"].str.upper() == target_brand.upper()]

    if filtered.empty:
        return json.dumps({
            "analysis_type": "gfk_sku_ranking",
            "error": "Belirtilen ürün grubu veya marka için SKU bulunamadı.",
            "query_product_group": target_group,
            "query_brand": target_brand,
            "available_product_groups": available_groups[:20],
            "available_brands": available_brands[:20]
        }, ensure_ascii=False)

    # Rank 1'den itibaren sırala, tekrarlı instore_code'ları deduplicate et
    filtered = filtered.sort_values(["rank", "brand", "item"])
    top_skus = (
        filtered
        .drop_duplicates(subset=["item", "brand"])
        .head(rank_limit)
    )

    rows = top_skus[[
        "reportingproductgroup", "rank", "brand", "item", "instore_code"
    ]].to_dict(orient="records")

    return json.dumps({
        "analysis_type": "gfk_sku_ranking",
        "source": "GfK SKU Leaderpanel",
        "filtered_by_product_group": target_group,
        "filtered_by_brand": target_brand,
        "rank_limit": rank_limit,
        "sku_count": len(rows),
        "rows": rows,
        "all_brands_in_group": (
            list(filtered["brand"].unique())[:15] if target_group and not target_brand else None
        )
    }, ensure_ascii=False, default=str)


# ─────────────────────────────────────────────────────────────────────────────
# 4. GfK + Ecommerce Birleşik Analiz
# ─────────────────────────────────────────────────────────────────────────────

def analyze_gfk_combined(question: str) -> str:
    """
    GfK pazar verilerini şirketin ecommerce metrikleriyle (C2D, B2D, revenue,
    stok) birleştirerek cross-analiz üretir.

    Desteklenen soru tipleri:
    - "GfK market share'imizi C2D ile kıyasla"
    - "Pazar payımız yüksek ama C2D düşük olan kategoriler"
    - "SAMSUNG'un pazar payı ile bizim satışlarımızı karşılaştır"
    - "GfK'da büyüyen ama bizim satışımız düşen kategoriler"
    - "Pazar genişliyor ama biz yararlanamıyor muyuz?"
    """
    q = question.lower()

    # GfK PW vs CW
    df_pw, err1 = load_gfk_pw_cw()
    # Ecommerce
    df_ec, err2 = load_ecommerce()
    # GfK Brand
    df_brand, err3 = load_gfk_brand()

    errors = []
    if err1:
        errors.append(f"GfK PW vs CW: {err1}")
    if err2:
        errors.append(f"Ecommerce: {err2}")

    if errors and df_pw is None and df_ec is None:
        return json.dumps({
            "analysis_type": "gfk_combined_error",
            "errors": errors
        }, ensure_ascii=False)

    result = {
        "analysis_type": "gfk_combined",
        "source": "GfK Leaderpanel + Ecommerce Sample Data",
    }

    # ─── GfK Pazar Özeti ────────────────────────────────────────────────
    if df_pw is not None:
        cols = list(df_pw.columns)
        value_cols = [c for c in cols if str(c).replace(".", "").isdigit() and "." not in str(c)]
        pct_col = "col_4" if "col_4" in cols else None    # WoW %
        ihs_col = "col_6" if "col_6" in cols else None    # Ihs

        gfk_rows = []
        for _, r in df_pw.iterrows():
            cat = r.get("product_group")
            if not cat or str(cat) in ["nan"]:
                continue
            cat_data = {"category": str(cat)}
            if value_cols and len(value_cols) >= 2:
                try:
                    cat_data["current_week_market_try"] = fmt_try(r.get(value_cols[1]))
                except Exception:
                    pass
            if pct_col and pd.notna(r.get(pct_col)):
                try:
                    cat_data["market_wow_growth_pct"] = round(float(r.get(pct_col)) * 100, 1)
                except Exception:
                    pass
            if ihs_col and pd.notna(r.get(ihs_col)):
                try:
                    cat_data["mediamarkt_market_share_pct"] = round(float(r.get(ihs_col)) * 100, 1)
                except Exception:
                    pass
            gfk_rows.append(cat_data)

        result["gfk_market_overview"] = gfk_rows[:15]


    # ─── Ecommerce Özeti ────────────────────────────────────────────────
    if df_ec is not None:
        ec_summary = []
        if "brand" in df_ec.columns:
            brand_grp = df_ec.groupby("brand").agg(
                revenue_sum=("revenue", "sum"),
                avg_c2d=("c2d_pct", "mean"),
                avg_b2d=("b2d_pct", "mean"),
                avg_revenue_delta=("revenue_delta_pct", "mean"),
                sku_count=("sku", "nunique")
            ).reset_index()
            brand_grp = brand_grp.sort_values("revenue_sum", ascending=False)
            for _, r in brand_grp.head(15).iterrows():
                ec_summary.append({
                    "brand": r["brand"],
                    "revenue_sum": fmt_try(r.get("revenue_sum")),
                    "avg_c2d_pct": round(float(r.get("avg_c2d", 0) or 0), 2),
                    "avg_b2d_pct": round(float(r.get("avg_b2d", 0) or 0), 2),
                    "avg_revenue_delta_pct": round(float(r.get("avg_revenue_delta", 0) or 0), 2),
                    "sku_count": int(r.get("sku_count", 0))
                })

        result["ecommerce_brand_performance"] = ec_summary

        # Kategori bazında ecommerce
        if "cat1" in df_ec.columns:
            cat_grp = df_ec.groupby("cat1").agg(
                revenue_sum=("revenue", "sum"),
                avg_c2d=("c2d_pct", "mean"),
                avg_b2d=("b2d_pct", "mean"),
            ).reset_index()
            cat_grp = cat_grp.sort_values("revenue_sum", ascending=False)
            cat_list = []
            for _, r in cat_grp.head(10).iterrows():
                cat_list.append({
                    "internal_category": r["cat1"],
                    "revenue_sum": fmt_try(r.get("revenue_sum")),
                    "avg_c2d_pct": round(float(r.get("avg_c2d", 0) or 0), 2),
                    "avg_b2d_pct": round(float(r.get("avg_b2d", 0) or 0), 2),
                })
            result["ecommerce_category_performance"] = cat_list

    # ─── Cross Insight ───────────────────────────────────────────────────
    cross_insights = []

    if "gfk_market_overview" in result and "ecommerce_brand_performance" in result:
        # Pazar büyürken iç satışı düşen/artan markalar
        gfk_growing_cats = {
            r["category"]: r.get("market_wow_growth_pct", 0)
            for r in result.get("gfk_market_overview", [])
            if r.get("market_wow_growth_pct", 0) > 0
        }

        for brand_row in result.get("ecommerce_brand_performance", []):
            brand = brand_row.get("brand", "")
            rev_delta = brand_row.get("avg_revenue_delta_pct", 0)
            c2d = brand_row.get("avg_c2d_pct", 0)
            b2d = brand_row.get("avg_b2d_pct", 0)

            if rev_delta < -5 and c2d < 5:
                cross_insights.append(
                    f"⚠️ {brand}: Pazar büyüme döneminde revenue %{abs(rev_delta):.1f} düşüyor, C2D {c2d:.1f}% — Fiyat veya stok riski incelenebilir."
                )
            elif rev_delta > 10 and b2d > 3:
                cross_insights.append(
                    f"✅ {brand}: Revenue +%{rev_delta:.1f} büyüyor, B2D {b2d:.1f}% — Momentum güçlü, trafik artırılabilir."
                )

    result["cross_insights"] = cross_insights[:5]
    result["category_mapping_note"] = (
        "GfK kategorileri (Smartphones, COMPUTER HW vb.) ile iç kategori isimleri "
        "(Cep Telefonları, Telefon vb.) marka adı üzerinden eşleştirilmiştir."
    )

    return json.dumps(result, ensure_ascii=False, default=str)
