"""
Funnel Master: Uçtan uca kullanıcı yolculuğu (PDP → Transaction) analiz motoru.
Her adım arasındaki drop-off'u hesaplar, e-ticaret mantığıyla teşhis koyar,
kategori/kanal/cihaz bazında kırılım yapar ve nokta atışı aksiyonlar üretir.
"""

import json
import re
from functools import lru_cache
from typing import Optional

import pandas as pd

from functions.analytics import load_sample_data


# ─────────────────────────────────────────────────────────────────────────────
# Funnel adımları: sıralı tanım
# ─────────────────────────────────────────────────────────────────────────────
FUNNEL_STEPS = [
    {
        "key": "pdp",
        "label": "PDP View",
        "col": "total_unique_pdp_views_sum",
        "delta_col": "pdp_delta_pct",
        "rate_col": None,
    },
    {
        "key": "a2c",
        "label": "Add to Cart (A2C)",
        "col": "total_unique_add_to_carts_sum",
        "delta_col": "a2c_delta_pct",
        "rate_col": "c2d_pct",  # C2D = A2C / PDP
    },
    {
        "key": "cart_view",
        "label": "Cart View",
        "col": "cart_views",
        "delta_col": "cartview_delta_pct",
        "rate_col": "cart_view_rate_pct",
    },
    {
        "key": "shipping",
        "label": "Shipping View",
        "col": "shipping_views",
        "delta_col": "shipping_delta_pct",
        "rate_col": "shipping_rate_pct",
    },
    {
        "key": "payment",
        "label": "Payment View",
        "col": "payment_views",
        "delta_col": "payment_delta_pct",
        "rate_col": "payment_rate_pct",
    },
    {
        "key": "summary",
        "label": "Summary View",
        "col": "summary_views",
        "delta_col": "summary_delta_pct",
        "rate_col": "summary_rate_pct",
    },
    {
        "key": "checkout",
        "label": "Checkout Submit",
        "col": "checkout_submits",
        "delta_col": "checkout_submit_delta_pct",
        "rate_col": "checkout_submit_rate_pct",
    },
    {
        "key": "transaction",
        "label": "Transactions",
        "col": "total_transactions_sum",
        "delta_col": "transactions_delta_pct",
        "rate_col": "b2d_pct",  # B2D = Transactions / PDP
    },
]


# ─────────────────────────────────────────────────────────────────────────────
# Teşhis kütüphanesi — her adım çifti için e-ticaret yorumu
# ─────────────────────────────────────────────────────────────────────────────
STEP_DIAGNOSES = {
    ("pdp", "a2c"): {
        "high_drop": (
            "🛒 PDP → A2C'de yüksek kayıp: Ürün sayfasına gelen kullanıcıların büyük çoğunluğu sepete eklemiyor. "
            "Ürün görselleri, açıklaması veya fiyatlandırma caydırıcı olabilir. "
            "Fiyat benchmark'ını kontrol edin; rakibe göre pahalı ürünlerde bu oran tipik olarak düşer."
        ),
        "anomaly_up": (
            "⚡ A2C oranı PDP'yi geçiyor (anomali): Aynı kullanıcı birden fazla ürünü sepete ekliyor veya "
            "ölçüm tekniği (session başına değil toplam event) şişirilmiş veri üretiyor."
        ),
    },
    ("a2c", "cart_view"): {
        "high_drop": (
            "🛒 A2C → Cart View'da yüksek kayıp: Sepete ekleme oldu ama kullanıcı sepet sayfasına gitmiyor. "
            "Büyük ihtimalle 'Sepete Eklendi' pop-up'ı sonrası sepete yönlendirme yetersiz veya "
            "kullanıcı farklı bir ürün araştırmaya devam etti (browsing mode). "
            "Sepete yönlendirme CTA'sını (buton konumu, metni) güçlendirin."
        ),
        "low_drop": (
            "✅ A2C → Cart View geçişi sağlıklı: Sepete ekleme, sepeti görüntülemeye iyi dönüşüyor. "
            "Bir sonraki adım (Shipping) drop-off'unu inceleyin."
        ),
    },
    ("cart_view", "shipping"): {
        "high_drop": (
            "🚚 Cart → Shipping'de yüksek kayıp: En yaygın e-ticaret sorunlarından biri. "
            "Kullanıcı sepeti gördü ama kargo adımına geçmedi. Olası nedenler: "
            "Kargo ücreti sürprizi (ücretsiz kargo baremi yüksek), "
            "toplam tutar beklentiden yüksek geldi, kayıtlı üye olmak zorunda bırakıldı (guest checkout yok). "
            "Aksiyon: Kargo ücretini sepet sayfasında şeffaf gösterin, ücretsiz kargo baremini düşürün."
        ),
        "low_drop": (
            "✅ Cart → Shipping geçişi iyi. Kullanıcılar kargo adımına rahat geçiyor."
        ),
    },
    ("shipping", "payment"): {
        "high_drop": (
            "📦 Shipping → Payment'da yüksek kayıp: Kullanıcı kargo seçeneklerini gördü ama ödeme sayfasına geçmedi. "
            "Olası nedenler: Kargo seçenekleri yetersiz veya teslimat süreleri uzun, "
            "adres formu çok uzun/karmaşık. "
            "Aksiyon: Kargo seçeneklerini genişletin, adres formunu kısaltın, gün bazlı teslimat tahmini ekleyin."
        ),
        "anomaly_up": (
            "⚡ Shipping → Payment'da artış (anomali): Payment adımındaki hacim Shipping'den büyük. "
            "Kullanıcılar sayfalar arası gidip geliyor veya ölçüm tekrarlı pageview sayıyor. "
            "Veri kalitesini (unique session bazı) kontrol edin."
        ),
    },
    ("payment", "summary"): {
        "high_drop": (
            "💳 Payment → Summary'de yüksek kayıp: Kullanıcı ödeme sayfasında takılıyor. "
            "Olası nedenler: Taksit seçenekleri yetersiz, 3D Secure hataları, "
            "kabul edilen kart/cüzdan çeşitliliği az, güven unsurları (SSL, logo) görünmüyor. "
            "Aksiyon: Alternatif ödeme yöntemleri (BKM Express, havale, kapıda ödeme) ekleyin, "
            "3D Secure başarı oranını ödeme altyapısı ekibiyle inceleyin."
        ),
    },
    ("summary", "checkout"): {
        "high_drop": (
            "📋 Summary → Checkout Submit'te kayıp: Kullanıcı özet sayfasını gördü ama 'Siparişi Tamamla' butonuna basmadı. "
            "Son anda fiyat veya ek ücret sürprizi (sigorta, ambalaj vb.) caydırıcı olabilir. "
            "Aksiyon: Özet sayfasındaki gizli ücretleri kaldırın, 'Siparişi Tamamla' butonunu daha prominent yapın."
        ),
    },
    ("checkout", "transaction"): {
        "high_drop": (
            "❌ Checkout Submit → Transaction'da kayıp: Sipariş tamamlama tıklandı ama kayıt oluşmadı. "
            "Bu teknik bir problemin işareti: ödeme gateway timeout, 3D Secure redirect hatası, "
            "sunucu tarafı hata veya çift tıklama önleme eksikliği. "
            "Aksiyon: Checkout → Transaction başarı oranını (conversion rate of submit) gerçek zamanlı monitor edin, "
            "gateway hata loglarını inceleyin."
        ),
    },
}


def normalize_q(q: str) -> str:
    tr_map = str.maketrans("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC")
    return q.lower().strip().translate(tr_map)


def get_funnel_df(df: pd.DataFrame) -> pd.DataFrame:
    """Mevcut tüm funnel kolonlarını sayısal hale getirir."""
    funnel_df = df.copy()
    all_cols = [s["col"] for s in FUNNEL_STEPS] + [s["delta_col"] for s in FUNNEL_STEPS if s["delta_col"]]
    for col in all_cols:
        if col in funnel_df.columns:
            funnel_df[col] = pd.to_numeric(funnel_df[col], errors="coerce")
    return funnel_df


def compute_funnel_summary(df: pd.DataFrame) -> list:
    """
    Tüm veri seti üzerinde her adımın toplam hacmini ve
    bir önceki adıma göre drop-off yüzdesini hesaplar.
    """
    steps_data = []
    prev_vol = None
    for step in FUNNEL_STEPS:
        col = step["col"]
        if col not in df.columns:
            continue
        vol = df[col].sum()
        delta_col = step["delta_col"]
        avg_delta = df[delta_col].mean() if delta_col and delta_col in df.columns else None
        drop_from_prev = None
        if prev_vol and prev_vol > 0:
            drop_from_prev = round((1 - vol / prev_vol) * 100, 1)
        steps_data.append({
            "step": step["label"],
            "key": step["key"],
            "volume": int(vol),
            "drop_from_prev_pct": drop_from_prev,
            "avg_delta_pct": round(avg_delta, 1) if avg_delta is not None else None,
        })
        prev_vol = vol
    return steps_data


def find_biggest_bottleneck(steps_data: list) -> Optional[dict]:
    """En büyük drop-off yaşanan adımı bulur."""
    with_drop = [s for s in steps_data if s["drop_from_prev_pct"] is not None and s["drop_from_prev_pct"] > 0]
    if not with_drop:
        return None
    return max(with_drop, key=lambda x: x["drop_from_prev_pct"])


def get_step_diagnosis(prev_key: str, curr_key: str, drop_pct: float) -> str:
    pair = (prev_key, curr_key)
    diagnoses = STEP_DIAGNOSES.get(pair, {})
    if drop_pct > 60 and "anomaly_up" in diagnoses and drop_pct < 0:
        return diagnoses["anomaly_up"]
    if drop_pct < -5:  # Bir sonraki adım öncekinden büyük → anomali
        return diagnoses.get("anomaly_up", f"⚡ {curr_key} adımı önceki adımdan büyük — olası veri anomalisi (tekrarlı sayım).")
    if drop_pct >= 50:
        return diagnoses.get("high_drop", f"🚨 %{drop_pct:.0f} drop-off: {curr_key} adımında kritik kayıp yaşanıyor.")
    if 25 <= drop_pct < 50:
        return diagnoses.get("high_drop", f"⚠️ %{drop_pct:.0f} drop-off: {curr_key} adımında önemli kayıp.")
    return diagnoses.get("low_drop", f"✅ %{drop_pct:.0f} drop-off: Normal seviyede.")


def breakdown_by_dimension(df: pd.DataFrame, dimension_col: str) -> list:
    """Belirli bir boyuta göre (kategori, kanal, cihaz) funnel özetini döner."""
    if dimension_col not in df.columns:
        return []

    rows = []
    for dim_val, group in df.groupby(dimension_col):
        steps = compute_funnel_summary(group)
        bottleneck = find_biggest_bottleneck(steps)
        pdp = next((s["volume"] for s in steps if s["key"] == "pdp"), 0)
        txn = next((s["volume"] for s in steps if s["key"] == "transaction"), 0)
        overall_conv = round(txn / pdp * 100, 3) if pdp > 0 else 0
        rows.append({
            "dimension": str(dim_val),
            "pdp_views": pdp,
            "transactions": txn,
            "overall_conversion_pct": overall_conv,
            "biggest_bottleneck": bottleneck["step"] if bottleneck else "N/A",
            "bottleneck_drop_pct": bottleneck["drop_from_prev_pct"] if bottleneck else None,
        })

    rows.sort(key=lambda x: x["pdp_views"], reverse=True)
    return rows[:20]


def extract_dimension_from_question(q: str) -> Optional[str]:
    """Sorudan hangi kırılım boyutunun istendiğini çıkarır."""
    qn = normalize_q(q)
    if any(w in qn for w in ["kategori", "cat1", "cat2", "category"]):
        return "cat1"
    if any(w in qn for w in ["alt kategori", "subkategori", "cat2"]):
        return "cat2"
    if any(w in qn for w in ["kanal", "channel", "trafik", "traffic"]):
        return "traffic_channel"
    if any(w in qn for w in ["cihaz", "device", "mobil", "mobile", "desktop", "tablet"]):
        return "device"
    if any(w in qn for w in ["marka", "brand"]):
        return "brand"
    if any(w in qn for w in ["satis kanali", "sales channel", "platform"]):
        return "sales_channel"
    return None


def extract_segment_filter(df: pd.DataFrame, q: str) -> pd.DataFrame:
    """Sorudaki ürün/kategori/marka/cihaz filtrelerini uygular."""
    qn = normalize_q(q)
    filtered = df.copy()

    # Cihaz filtresi
    if "mobile" in qn or "mobil" in qn:
        if "device" in filtered.columns:
            filtered = filtered[filtered["device"].astype(str).str.lower().str.contains("mobile|mobil", na=False)]
    elif "desktop" in qn:
        if "device" in filtered.columns:
            filtered = filtered[filtered["device"].astype(str).str.lower().str.contains("desktop", na=False)]

    # Marka filtresi
    for brand_val in df.get("brand", pd.Series(dtype=str)).dropna().unique():
        if normalize_q(str(brand_val)) in qn and len(str(brand_val)) > 2:
            filtered = filtered[filtered["brand"].astype(str).apply(normalize_q) == normalize_q(str(brand_val))]
            break

    # Kategori filtresi
    for col in ["cat2", "cat1"]:
        if col not in df.columns:
            continue
        for val in sorted(df[col].dropna().unique(), key=lambda x: len(str(x)), reverse=True):
            if normalize_q(str(val)) in qn and len(str(val)) > 2:
                filtered = filtered[filtered[col].astype(str).apply(normalize_q) == normalize_q(str(val))]
                break

    return filtered


def analyze_funnel_master(question: str) -> str:
    """
    Ana giriş noktası. Kullanıcı sorusunu alır, funnel analizi yapar,
    e-ticaret kafasıyla teşhis koyar ve JSON döner.
    """
    try:
        df = load_sample_data()
    except Exception as e:
        return json.dumps({
            "analysis_type": "funnel_master_error",
            "error": "Sample data okunamadı.",
            "detail": str(e),
        }, ensure_ascii=False)

    funnel_df = get_funnel_df(df)

    # Segment filtresi uygula
    filtered = extract_segment_filter(funnel_df, question)
    if filtered.empty:
        filtered = funnel_df

    qn = normalize_q(question)

    # ─── Hangi soruya ne döneceğiz? ───────────────────────────────────────────

    # Soru: Kırılım analizi (kategori, cihaz, kanal bazında)
    dimension = extract_dimension_from_question(question)

    # Soru: Belirli bir adımı soruyor mu?
    step_focus = None
    for step in FUNNEL_STEPS:
        if step["key"] in qn or normalize_q(step["label"]) in qn:
            step_focus = step
            break
    if "kargo" in qn or "shipping" in qn:
        step_focus = next(s for s in FUNNEL_STEPS if s["key"] == "shipping")
    if "odeme" in qn or "payment" in qn or "ödeme" in qn:
        step_focus = next(s for s in FUNNEL_STEPS if s["key"] == "payment")
    if "sepet" in qn or "cart" in qn:
        step_focus = next(s for s in FUNNEL_STEPS if s["key"] == "cart_view")
    if "checkout" in qn:
        step_focus = next(s for s in FUNNEL_STEPS if s["key"] == "checkout")

    # ─── Genel funnel özeti (her zaman hesapla) ───────────────────────────────
    steps_data = compute_funnel_summary(filtered)
    bottleneck = find_biggest_bottleneck(steps_data)

    # Her adım için teşhis metni üret
    for i, step in enumerate(steps_data):
        if i == 0:
            step["diagnosis"] = "🔍 Funnel başlangıç noktası."
            continue
        prev_key = steps_data[i - 1]["key"]
        curr_key = step["key"]
        drop = step["drop_from_prev_pct"] or 0
        step["diagnosis"] = get_step_diagnosis(prev_key, curr_key, drop)

    # ─── Boyut kırılımı ───────────────────────────────────────────────────────
    breakdown = []
    if dimension:
        breakdown = breakdown_by_dimension(filtered, dimension)

    # ─── Genel dönüşüm oranı ─────────────────────────────────────────────────
    pdp_total = next((s["volume"] for s in steps_data if s["key"] == "pdp"), 0)
    txn_total = next((s["volume"] for s in steps_data if s["key"] == "transaction"), 0)
    overall_conv = round(txn_total / pdp_total * 100, 3) if pdp_total > 0 else 0

    # ─── Aksiyon önerileri ───────────────────────────────────────────────────
    recommended_actions = []
    if bottleneck:
        bk = bottleneck["key"]
        b_drop = bottleneck["drop_from_prev_pct"]
        bn_step = bottleneck["step"]
        recommended_actions.append(
            f"En kritik darboğaz: **{bn_step}** — %{b_drop} kullanıcı bu adımda ayrılıyor. "
            "Buraya odaklanmak en yüksek dönüşüm artışını sağlar (Pareto etkisi)."
        )
        if bk == "cart_view":
            recommended_actions.append("Sepet sayfasına geçişi artırmak için 'Sepete Eklendi' sonrası kullanıcıyı aktif sepete yönlendirin.")
        elif bk == "shipping":
            recommended_actions.append("Kargo maliyetini sepet sayfasında görünür yapın ve ücretsiz kargo baremini gözden geçirin.")
        elif bk == "payment":
            recommended_actions.append("Alternatif ödeme yöntemleri (BKM Express, Papara, havale) ve taksit seçenekleri ekleyin.")
        elif bk == "checkout":
            recommended_actions.append("Checkout gateway başarı oranını (submit → confirmed) teknik ekiple izleyin; 3D Secure hatalarını ölçün.")
        elif bk == "a2c":
            recommended_actions.append("Ürün sayfası fiyatını benchmark'la karşılaştırın; rakibe göre pahalı ürünlerde A2C düşer.")

    recommended_actions.append(
        f"Genel funnel dönüşüm oranınız PDP → Transaction: %{overall_conv}. "
        "Sektör benchmarkı %1-3 arasında; bu değerin altındaysanız öncelikli adımı optimize edin."
    )

    if any(s.get("avg_delta_pct", 0) and s["avg_delta_pct"] < -10 for s in steps_data):
        worsening = [s["step"] for s in steps_data if s.get("avg_delta_pct") and s["avg_delta_pct"] < -10]
        recommended_actions.append(
            f"Önceki dönemle kıyaslandığında {', '.join(worsening)} adımlarında %10+ gerileme var — "
            "bu adımlardaki değişiklikleri (UX güncellemesi, fiyat değişikliği, kampanya bitmesi) inceleyin."
        )

    return json.dumps({
        "analysis_type": "funnel_master_analysis",
        "question": question,
        "overall_conversion_pct": overall_conv,
        "pdp_total": pdp_total,
        "transaction_total": txn_total,
        "bottleneck": bottleneck,
        "funnel_steps": steps_data,
        "dimension_breakdown": breakdown,
        "dimension": dimension,
        "recommended_actions": recommended_actions,
    }, ensure_ascii=False, default=str)
