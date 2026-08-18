# -*- coding: utf-8 -*-
import os
import re
import json
from typing import Optional, Dict, List
import numpy as np
import pandas as pd
from functools import lru_cache

COMPANY_INPUT_PATH = "data/company_product_input.xlsx"
MERCHANT_BENCHMARK_PATH = "data/merchant_price_benchmark_sample.xlsx"
TRENDS_PATH = "data/google_trends_seasonal_3y.xlsx"

COLUMN_ALIASES = {
    "gtin": ["gtin", "ean", "barcode", "product_gtin"],
    "sku": ["sku", "item_id", "product_id", "id"],
    "product_title": ["product_title", "title", "name", "product_name"],
    "brand": ["brand", "manufacturer", "marka"],
    "cat1": ["cat1", "category1", "category", "main_category"],
    "cat2": ["cat2", "category2", "subcategory"],
    "price": ["price", "your_price", "sale_price", "product_price"],
    "benchmark_price": ["benchmark_price", "market_price"],
    "stock_qty": ["stock_qty", "stock", "inventory"],
    "reorder_point_qty": ["reorder_point_qty", "reorder_point", "critical_stock"],
    "pdp_views": ["pdp_views", "pdp", "pdp_view"],
    "add_to_carts": ["add_to_carts", "a2c"],
    "transactions": ["transactions", "trans", "orders"],
    "revenue": ["revenue", "ciro"],
    "c2d_pct": ["c2d_pct", "c2d"],
    "b2d_pct": ["b2d_pct", "b2d"],
    "bounce_rate_pct": ["bounce_rate_pct", "bounce_rate"],
    "transactions_delta_pct": ["transactions_delta_pct", "transaction_delta_pct"],
    "revenue_delta_pct": ["revenue_delta_pct", "ciro_delta_pct"],
    "stock_coverage_days": ["stock_coverage_days", "stock_coverage"]
}

TRENDS_MAPPING = {
    "telefon": "akıllı telefon",
    "cep telefonları": "akıllı telefon",
    "cep telefonlari": "akıllı telefon",
    "iphone": "akıllı telefon",
    "bilgisayar": "laptop",
    "laptop": "laptop",
    "notebook": "laptop",
    "tablet": "tablet",
    "tabletler": "tablet",
    "ipad": "tablet",
    "süpürgeler": "robot süpürge",
    "süpürge": "robot süpürge",
    "robot süpürge": "robot süpürge",
    "dikey süpürge": "dikey süpürge",
    "kulaklıklar": "kulaklık",
    "kulaklık": "kulaklık",
    "kulaklik": "kulaklık",
    "akıllı saatler": "akıllı saat",
    "akilli saatler": "akıllı saat",
    "giyilebilir teknoloji": "akıllı saat",
    "televizyon": "televizyon",
    "tv": "televizyon",
    "kahve makinesi": "kahve makinesi",
    "ütü": "ütü",
    "monitör": "monitör",
    "klavye": "klavye"
}

def normalize_text(val) -> str:
    val = str(val or "").strip().lower()
    tr_map = str.maketrans("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC")
    val = val.translate(tr_map)
    return re.sub(r"\s+", " ", val)

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_").replace("-", "_").replace(".", "_") for c in df.columns]
    
    normalized = pd.DataFrame(index=df.index)
    for target, aliases in COLUMN_ALIASES.items():
        found = None
        for alias in aliases:
            if alias in df.columns:
                found = alias
                break
        normalized[target] = df[found] if found else None
        
    # Convert numerical columns
    num_cols = ["price", "benchmark_price", "stock_qty", "reorder_point_qty", "pdp_views", "add_to_carts", 
                "transactions", "revenue", "c2d_pct", "b2d_pct", "bounce_rate_pct", "transactions_delta_pct", 
                "revenue_delta_pct", "stock_coverage_days"]
    for col in num_cols:
        if col in normalized.columns:
            normalized[col] = pd.to_numeric(normalized[col], errors="coerce")
            
    # Strings
    str_cols = ["gtin", "sku", "product_title", "brand", "cat1", "cat2"]
    for col in str_cols:
        if col in normalized.columns:
            normalized[col] = normalized[col].astype(str).replace("nan", "").str.strip()
            
    return normalized

def load_cross_data() -> pd.DataFrame:
    if not os.path.exists(COMPANY_INPUT_PATH):
        raise FileNotFoundError(f"Şirket input dosyası bulunamadı: {COMPANY_INPUT_PATH}")
    
    comp_df = pd.read_excel(COMPANY_INPUT_PATH)
    comp = normalize_columns(comp_df)
    
    if os.path.exists(MERCHANT_BENCHMARK_PATH):
        bench_df = pd.read_excel(MERCHANT_BENCHMARK_PATH)
        bench = normalize_columns(bench_df)
        
        comp["gtin"] = comp["gtin"].astype(str).str.strip()
        bench["gtin"] = bench["gtin"].astype(str).str.strip()
        
        bench_cols = [c for c in ["gtin", "benchmark_price"] if c in bench.columns]
        if "benchmark_price" in bench_cols:
            comp = comp.drop(columns=["benchmark_price"], errors="ignore")
            comp = comp.merge(bench[bench_cols], on="gtin", how="left")
            
    if "price" in comp.columns and "benchmark_price" in comp.columns:
        comp["price_gap"] = comp["price"] - comp["benchmark_price"]
        comp["price_gap_pct"] = (comp["price_gap"] / comp["benchmark_price"]) * 100
        
    return comp

def get_trends_analysis(category_name: str) -> Optional[dict]:
    if not os.path.exists(TRENDS_PATH):
        return None
    try:
        trends_df = pd.read_excel(TRENDS_PATH)
        norm_cat = normalize_text(category_name)
        
        # Match trends column
        matched_col = None
        for key, trend_col in TRENDS_MAPPING.items():
            if normalize_text(key) in norm_cat or norm_cat in normalize_text(key):
                matched_col = trend_col
                break
                
        if not matched_col:
            # Try fuzzy match direct column names
            for col in trends_df.columns:
                if col == "date":
                    continue
                if normalize_text(col) in norm_cat or norm_cat in normalize_text(col):
                    matched_col = col
                    break
                    
        if not matched_col:
            # Default to first trend column
            non_date = [c for c in trends_df.columns if c != "date"]
            matched_col = non_date[0] if non_date else None
            
        if not matched_col or matched_col not in trends_df.columns:
            return None
            
        date_col = 'date' if 'date' in trends_df.columns else trends_df.columns[0]
        trends_df[date_col] = pd.to_datetime(trends_df[date_col])
        trends_df['month'] = trends_df[date_col].dt.month
        monthly_avg = trends_df.groupby('month')[matched_col].mean()
        
        peak_month = int(monthly_avg.idxmax())
        low_month = int(monthly_avg.idxmin())
        
        month_names = {
            1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
            7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
        }
        
        y = trends_df[matched_col].values
        x = np.arange(len(y))
        slope = np.polyfit(x, y, 1)[0] if len(y) > 1 else 0
        trend_direction = "Artış Eğiliminde" if slope > 0.05 else ("Azalış Eğiliminde" if slope < -0.05 else "Dengeli/Stabil")
        
        return {
            "keyword": matched_col,
            "peak_month": month_names.get(peak_month, str(peak_month)),
            "low_month": month_names.get(low_month, str(low_month)),
            "trend_direction": trend_direction,
            "avg_interest_score": round(float(monthly_avg.mean()), 2)
        }
    except Exception as e:
        print(f"Trends parsing warning: {e}")
        return None

def analyze_cross_performance(question: str) -> str:
    try:
        df = load_cross_data()
    except Exception as e:
        return json.dumps({
            "analysis_type": "cross_performance_error",
            "error": "Çapraz veri analizi dosyaları yüklenemedi.",
            "detail": str(e)
        }, ensure_ascii=False)
        
    # Standardize empty values
    for col in ["price_gap_pct", "revenue_delta_pct", "transactions", "revenue", "stock_qty", "pdp_views", "b2d_pct"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0.0
            
    # Resolve category filters
    qn = normalize_text(question)
    filtered = df.copy()
    category = "genel"
    
    # Try finding category in query
    for user_word, values in COLUMN_ALIASES.items():
        pass
        
    for cat_val in list(df["cat2"].dropna().unique()) + list(df["cat1"].dropna().unique()):
        if cat_val and normalize_text(str(cat_val)) in qn and len(str(cat_val)) > 2:
            category = str(cat_val)
            filtered = df[df["cat2"].astype(str).apply(normalize_text).str.contains(normalize_text(category), na=False) | 
                          df["cat1"].astype(str).apply(normalize_text).str.contains(normalize_text(category), na=False)].copy()
            break
            
    # Calculate scenarios
    scenarios = {
        "expensive_falling_sales": [],
        "expensive_good_sales": [],
        "cheap_no_sales": [],
        "cheap_good_sales": [],
        "losing_competitiveness": []
    }
    
    # 1. Scenario 1: Expensive & Falling
    s1_df = filtered[(filtered["price_gap_pct"] > 1) & (filtered["transactions_delta_pct"] < -5)].sort_values("revenue", ascending=False)
    for _, r in s1_df.head(5).iterrows():
        scenarios["expensive_falling_sales"].append({
            "sku": r.get("sku"),
            "product_title": r.get("product_title"),
            "brand": r.get("brand"),
            "price": float(r.get("price")),
            "benchmark_price": float(r.get("benchmark_price")),
            "price_gap_pct": float(r.get("price_gap_pct")),
            "revenue_delta_pct": float(r.get("revenue_delta_pct")),
            "transactions": int(r.get("transactions")),
            "action": "Price Action Önerilir: Fiyatı benchmark seviyesine çekerek kaybı durdurun."
        })
        
    # 2. Scenario 2: Expensive but Sales Good
    s2_df = filtered[(filtered["price_gap_pct"] > 1) & (filtered["transactions_delta_pct"] >= 0)].sort_values("revenue", ascending=False)
    for _, r in s2_df.head(5).iterrows():
        scenarios["expensive_good_sales"].append({
            "sku": r.get("sku"),
            "product_title": r.get("product_title"),
            "brand": r.get("brand"),
            "price": float(r.get("price")),
            "benchmark_price": float(r.get("benchmark_price")),
            "price_gap_pct": float(r.get("price_gap_pct")),
            "revenue_delta_pct": float(r.get("revenue_delta_pct")),
            "transactions": int(r.get("transactions")),
            "action": "Premium/Güçlü Ürün: Fiyat esnekliği düşük, mevcut fiyatı ve pozisyonu koruyun."
        })
        
    # 3. Scenario 3: Cheap but No Sales (Visibility/Content/Stock problems)
    s3_raw = filtered[(filtered["price_gap_pct"] < -1) & (filtered["transactions"] <= 18)]
    # Sort by lowest conversion rate (b2d_pct)
    s3_df = s3_raw.sort_values("pdp_views", ascending=False)
    for _, r in s3_df.head(5).iterrows():
        stock = int(r.get("stock_qty", 0))
        reorder = int(r.get("reorder_point_qty", 0)) if pd.notna(r.get("reorder_point_qty")) else 5
        
        if stock <= reorder:
            diag_category = "Stok Sorunu"
            action_desc = "Tedarik Sorunu: Fiyat avantajına rağmen stok yetersiz/tükenmiş (Stok: {}). Acil sipariş oluşturulmalı.".format(stock)
        else:
            diag_category = "Visibility/Content Sorunu"
            action_desc = "İçerik/Görünürlük Sorunu: Fiyat avantajı ve {} PDP görüntülenmeye rağmen satış yok. Ürün listelemesini, görselleri ve kargo bedelini optimize edin.".format(int(r.get("pdp_views", 0)))
            
        scenarios["cheap_no_sales"].append({
            "sku": r.get("sku"),
            "product_title": r.get("product_title"),
            "brand": r.get("brand"),
            "price": float(r.get("price")),
            "benchmark_price": float(r.get("benchmark_price")),
            "price_gap_pct": float(r.get("price_gap_pct")),
            "stock_qty": stock,
            "pdp_views": int(r.get("pdp_views")),
            "diag_category": diag_category,
            "action": action_desc
        })
        
    # 4. Scenario 4: Cheap & Good Sales
    s4_df = filtered[(filtered["price_gap_pct"] < -1) & (filtered["transactions"] > 18)].sort_values("revenue", ascending=False)
    for _, r in s4_df.head(5).iterrows():
        scenarios["cheap_good_sales"].append({
            "sku": r.get("sku"),
            "product_title": r.get("product_title"),
            "brand": r.get("brand"),
            "price": float(r.get("price")),
            "benchmark_price": float(r.get("benchmark_price")),
            "price_gap_pct": float(r.get("price_gap_pct")),
            "revenue_delta_pct": float(r.get("revenue_delta_pct")),
            "transactions": int(r.get("transactions")),
            "action": "Trafik/Bid Artır: Fiyat avantajı satış getiriyor. Satış hacmini katlamak için trafiği ve PPC reklam tekliflerini artırın."
        })
        
    # Losing Competitiveness
    if "brand" in filtered.columns and filtered["brand"].nunique() > 0:
        brand_groups = []
        for brand, g in filtered.groupby("brand"):
            total = len(g)
            exp_cnt = (g["price_gap_pct"] > 1).sum()
            ratio = (exp_cnt / total) * 100 if total > 0 else 0
            avg_gap = g["price_gap_pct"].mean()
            brand_groups.append({
                "brand": brand,
                "total_skus": total,
                "expensive_skus": int(exp_cnt),
                "ratio": float(ratio),
                "avg_gap": float(avg_gap)
            })
        scenarios["losing_competitiveness"] = sorted(brand_groups, key=lambda x: x["ratio"], reverse=True)[:5]
        
    # Google Trends Integration
    trends_info = get_trends_analysis(category)
    
    # Expensive SKUs list
    exp_skus_list = []
    exp_df = filtered[filtered["price_gap_pct"] > 1].sort_values("price_gap_pct", ascending=False)
    for _, r in exp_df.head(10).iterrows():
        exp_skus_list.append({
            "sku": r.get("sku"),
            "product_title": r.get("product_title"),
            "price": float(r.get("price")),
            "benchmark_price": float(r.get("benchmark_price")),
            "price_gap_pct": float(r.get("price_gap_pct")),
            "transactions": int(r.get("transactions"))
        })
        
    # Price Cut candidates (High PDP, High C2D, Expensive)
    price_cut_candidates = []
    pc_df = filtered[(filtered["price_gap_pct"] > 1) & (filtered["pdp_views"] > 1000)].sort_values("pdp_views", ascending=False)
    for _, r in pc_df.head(5).iterrows():
        price_cut_candidates.append({
            "sku": r.get("sku"),
            "product_title": r.get("product_title"),
            "price_gap_pct": float(r.get("price_gap_pct")),
            "pdp_views": int(r.get("pdp_views")),
            "c2d_pct": float(r.get("c2d_pct") if pd.notna(r.get("c2d_pct")) else 0),
            "b2d_pct": float(r.get("b2d_pct") if pd.notna(r.get("b2d_pct")) else 0),
            "transactions": int(r.get("transactions")),
            "action": "Fiyat indirimi ile satışları canlandırın. Ürün yüksek trafik alıyor ama fiyat bariyeri yüzünden satın almaya dönüşmüyor."
        })
        
    result = {
        "analysis_type": "cross_performance_analysis",
        "question": question,
        "category": category,
        "scenarios": scenarios,
        "trends": trends_info,
        "expensive_skus": exp_skus_list,
        "price_cut_candidates": price_cut_candidates,
        "summary": {
            "total_skus_analyzed": int(len(filtered)),
            "expensive_skus_count": int((filtered["price_gap_pct"] > 1).sum()),
            "cheap_skus_count": int((filtered["price_gap_pct"] < -1).sum()),
        }
    }
    
    return json.dumps(result, ensure_ascii=False, default=str)
