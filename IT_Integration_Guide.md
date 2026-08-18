# IT Integration Guide — Retail AI Tool
**Versiyon:** 2.0 — Haziran 2026
**Kapsam:** GfK Leaderpanel entegrasyonu, veri kaynakları, API endpoint'leri, format gereksinimleri

---

## İçindekiler

1. [Veri Kaynakları & Beklenen Dosya Formatları](#1-veri-kaynakları--beklenen-dosya-formatları)
2. [GfK Veri Entegrasyonu Detayları](#2-gfk-veri-entegrasyonu-detayları)
3. [Veri Yükleme Akışı](#3-veri-yükleme-akışı)
4. [API Endpoint'leri](#4-api-endpointleri)
5. [AI Model Araçları (Tools)](#5-ai-model-araçları-tools)
6. [Kategori Mapping Tablosu](#6-kategori-mapping-tablosu)
7. [Güvenlik & Deployment](#7-güvenlik--deployment)
8. [Sorun Giderme](#8-sorun-giderme)

---

## 1. Veri Kaynakları & Beklenen Dosya Formatları

Sistem `data/` dizinindeki aşağıdaki Excel dosyalarını kullanır. Tüm dosyalar `data/` klasörüne yerleştirilmelidir.

### 1.1 GfK Leaderpanel — `data/GfK_Leaderpanel.xlsx`

**Kaynak:** GfK Türkiye Leaderpanel raporu (haftalık)
**Boyut:** ~11 MB
**Güncelleme sıklığı:** Haftalık (her Perşembe)

| Sheet Adı | İçerik | Parse Header Row |
|---|---|---|
| `Summary_value` | Kategori × hafta × pazar değeri (TRY) + MM payı | 3 |
| `Summary_value%` | Kategori × hafta × yüzde paylar | 3 |
| `Summary_unit` | Kategori × hafta × birim satışlar (adet) | 3 |
| `Summary_valueYoY` | Yıllık bazlı değer karşılaştırması | 3 |
| `Brand` | Kategori × marka × haftalık MM internet payı (%) | 2 |
| `PW vs. CW` | Tüm kategoriler: geçen hafta vs bu hafta | 2 |
| `Hitlist_unit` | Kategori × haftalık hit listesi (adet) | — |
| `Hitlist_value` | Kategori × haftalık hit listesi (değer) | — |

**Zorunlu kolonlar (Summary_value sheet, header=3):**
- `Unnamed: 0` → `product_group` (kategori adı)
- `Unnamed: 2` → `metric` (`Sum of Internet Sales...` veya `Sum of MediaMarkt...`)
- Hafta kolonları: Sayısal (10, 11, 12...) veya `None` (YoY toplamlar için)

**Zorunlu kolonlar (Brand sheet, header=2):**
- `Unnamed: 0` → `product_group` (kategori adı)
- `Unnamed: 2` → `brand` (marka adı)
- `Week 17 2026`, `Week 18 2026`, ..., `Week 24 2026` → MM internet satış payı (0-1 arası ondalık)

**Zorunlu kolonlar (PW vs. CW sheet, header=2):**
- `Unnamed: 0` → `product_group` (kategori adı)
- Hafta kolonları: `Week 23 2026`, `Week 24 2026`
- `24 vs. 23 in %` → WoW yüzde değişim (ondalık, örn: -0.15 = -%15)
- `24 vs. 23 in val` → Mutlak değer farkı (TRY)
- `24 Ihs` → MediaMarkt pazar payı (0-1 arası ondalık)
- `24 Rank for Ihs` → Sıralama

---

### 1.2 GfK SKU Leaderpanel — `data/gfk_sku.xlsx`

**Kaynak:** GfK Türkiye SKU bazlı satış sıralaması
**Boyut:** ~940 KB | **Satır sayısı:** ~32,000+
**Sheet:** `Sheet1`

| Kolon | Tip | Açıklama |
|---|---|---|
| `ReportingProductgroup` | string | Ürün grubu adı (39 farklı değer) |
| `Rank` | integer | Satış sıralaması (1 = en çok satan) |
| `Brand` | string | Marka adı (1,312 farklı değer) |
| `Item` | string | Ürün model kodu (ör: `WGK264Z0TR`) |
| `Instore code` | integer | Mağaza kodu — bir ürün birden fazla satırda olabilir |

> **ÖNEMLİ:** Aynı `Item` birden fazla `Instore code` ile listelenebilir. Deduplicate işlemi `(Item, Brand)` üzerinden yapılır.

---

### 1.3 E-Commerce Sample Data — `data/ecommerce_ai_sample_data_200_rows.xlsx`

**Sheet:** `sample_data_200`
**Satır:** 200 | **Kolon:** 58

Kritik kolonlar:

| Kolon | Tip | Açıklama |
|---|---|---|
| `sku` | string | Ürün SKU kodu |
| `brand` | string | Marka adı |
| `cat1` | string | Ana kategori (ör: Telefon) |
| `cat2` | string | Alt kategori (ör: Cep Telefonları) |
| `c2d_pct` | float | Cart-to-Detail oranı (%) |
| `c2d_delta_pct` | float | C2D önceki dönem değişimi |
| `b2d_pct` | float | Buy-to-Detail oranı (%) |
| `b2d_delta_pct` | float | B2D önceki dönem değişimi |
| `revenue` | float | Satış geliri (TRY) |
| `revenue_delta_pct` | float | Revenue önceki dönem değişimi |
| `stock_qty` | integer | Mevcut stok adedi |
| `reorder_point_qty` | integer | Yeniden sipariş eşiği |
| `availability_status` | string | `in_stock`, `critical_low_stock`, `out_of_stock` |
| `total_unique_pdp_views_sum` | integer | Benzersiz ürün sayfası görüntülemesi |

---

### 1.4 Şirket Ürün Girdisi — `data/company_product_input.xlsx`

**Satır:** 50 | **Kolon:** 38

Kritik kolonlar: `gtin`, `sku`, `product_title`, `brand`, `cat1`, `cat2`, `price`, `stock_qty`, `pdp_views`, `add_to_carts`, `transactions`, `c2d_pct`, `b2d_pct`, `revenue`

> **GTIN:** Merchant Center benchmark eşleştirmesinde anahtar alan. Eksik veya hatalı GTIN, fiyat rekabeti analizini devre dışı bırakır.

---

### 1.5 Merchant Benchmark — `data/merchant_price_benchmark_sample.xlsx`

**Satır:** 50 | **Kolon:** 17

Kritik kolonlar: `gtin`, `sku`, `benchmark_price`, `benchmark_currency`, `min_competitor_price`, `median_competitor_price`, `max_competitor_price`, `market_price_trend_pct`

---

### 1.6 Google Trends — `data/google_trends_seasonal_3y.xlsx`

**Satır:** 156 (haftalık, 3 yıl) | **Kolon:** 21

Kritik kolonlar: `date` + 20 kategori sütunu (ör: `laptop`, `televizyon`, `akıllı telefon`)
Her değer 0–100 arası trending skoru.

---

## 2. GfK Veri Entegrasyonu Detayları

### 2.1 Parse Mantığı

```python
# Summary_value sheet parse (header row 3)
df = pd.read_excel("data/GfK_Leaderpanel.xlsx", sheet_name="Summary_value", header=3)
df["product_group"] = df["Unnamed: 0"].ffill()
df["metric_type"] = df["Unnamed: 2"].apply(
    lambda x: "mediamarkt" if "MediaMarkt" in str(x) else "total_internet"
)

# Brand sheet parse (header row 2)
df = pd.read_excel("data/GfK_Leaderpanel.xlsx", sheet_name="Brand", header=2)
# Week kolonları "Week 17 2026" formatında

# PW vs. CW parse (header row 2)
df = pd.read_excel("data/GfK_Leaderpanel.xlsx", sheet_name="PW vs. CW", header=2)
# WoW % kolonu ondalık: -0.15 = -%15, Ihs ondalık: 0.087 = %8.7
```

### 2.2 Kategori Adı Standartlaştırma

GfK kategori adları büyük harf ve İngilizce (ör: `WASHING MACHINES`). İç sistem kategorileri Türkçe (ör: `Çamaşır Makinesi`). Eşleştirme `functions/gfk_analyzer.py` içindeki `GFK_CATEGORY_MAP` sözlüğü üzerinden yapılır.

```python
GFK_CATEGORY_MAP = {
    "Smartphones": ["Cep Telefonları", "Telefon", "Akıllı Telefon"],
    "WASHING MACHINES": ["Çamaşır Makinesi"],
    "PTV/FLAT": ["Televizyon", "TV"],
    # ... (tam liste gfk_analyzer.py içinde)
}
```

### 2.3 Haftalık Güncelleme Prosedürü

1. GfK'dan yeni `GfK_Leaderpanel.xlsx` dosyasını indirin
2. `data/GfK_Leaderpanel.xlsx` dosyasını yeni dosyayla değiştirin
3. `gfk_sku.xlsx` değiştiyse `data/gfk_sku.xlsx` dosyasını güncelleyin
4. Sunucuyu yeniden başlatın (lru_cache sıfırlanır): `uvicorn main:app --reload`

> **DİKKAT:** `load_gfk_leaderpanel()`, `load_gfk_brand()`, `load_gfk_sku()` fonksiyonları `@lru_cache` ile önbelleğe alınır. Dosya değişikliği sonrası sunucunun yeniden başlatılması gerekir.

---

## 3. Veri Yükleme Akışı

### 3.1 Upload Endpoint

```http
POST /upload-data
Content-Type: multipart/form-data

files: [<dosya1.xlsx>, <dosya2.xlsx>, ...]
```

**Desteklenen dosya uzantıları:** `.xlsx`, `.csv`

**Dosya kayıt konumu:** `data/<dosya_adı>`

**Yanıt (başarılı):**
```json
{
  "saved": ["GfK_Leaderpanel.xlsx", "gfk_sku.xlsx"],
  "errors": []
}
```

**Yanıt (hatalı dosya tipi):**
```json
{
  "saved": [],
  "errors": ["GfK_report.pdf desteklenmeyen dosya tipi."]
}
```

### 3.2 Zorunlu Dosya Listesi

| Dosya | Zorunlu | Hangi Feature'ı Etkiler |
|---|---|---|
| `ecommerce_ai_sample_data_200_rows.xlsx` | ✅ | Tüm ecommerce analizleri |
| `company_product_input.xlsx` | ✅ | Fiyat rekabeti, business calculator |
| `merchant_price_benchmark_sample.xlsx` | ✅ | Fiyat rekabeti analizi |
| `google_trends_seasonal_3y.xlsx` | ✅ | Cross analiz, mevsimsel talep |
| `GfK_Leaderpanel.xlsx` | ✅ | GfK pazar payı, marka, PW vs CW |
| `gfk_sku.xlsx` | ✅ | GfK SKU sıralaması |

---

## 4. API Endpoint'leri

### `GET /`
Ana uygulama arayüzü. HTML döner.

---

### `POST /chat`
AI chat endpoint'i.

**Request:**
```json
{
  "message": "GfK'ya göre en çok büyüyen kategori hangisi?"
}
```

**Response:**
```json
{
  "reply": "📈 GfK Pazar Analizi\n\nKaynak: GfK Leaderpanel — PW vs. CW\n🚀 En Çok Büyüyen Kategoriler (WoW)..."
}
```

**Tool Routing Mantığı:**

| Soru İçeriği | Çağrılan Tool |
|---|---|
| `gfk`, `leaderpanel`, `pazar payı`, `market share`, `ihs` | `analyze_gfk_market_share` |
| `gfk` + marka adı (samsung, apple, bosch...) | `analyze_gfk_brand_performance` |
| `gfk` + `top 10`, `rank`, `ilk 5`, `model listesi` | `analyze_gfk_sku_ranking` |
| `gfk`/`pazar` + `c2d`, `b2d`, `kıyasla`, `karşılaştır` | `analyze_gfk_combined` |
| `mevsimsellik`, `trends`, `cross`, `ppc` | `analyze_cross_performance` |
| `funnel`, `drop-off`, `checkout`, `sepet` | `analyze_funnel_master` |
| `fiyat rekabeti`, `benchmark`, `gtin` | `generate_price_competition_from_uploaded_inputs` |
| `aksiyon`, `replenishment`, `excel çıkar` | `execute_recommended_action` |

---

### `POST /upload-data`
Dosya yükleme endpoint'i.

**Request:** `multipart/form-data`, alan adı: `files`

---

### `GET /export`
Son analiz sonucunu Excel olarak indirir.

**Response:** `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

**Excel Sheet yapısı (GfK analiz tiplerine göre):**

| Analiz Tipi | Sheet'ler |
|---|---|
| `gfk_market_share` | GfK Market Share Summary, Category Rows, Raw JSON |
| `gfk_brand_performance` | GfK Brand Summary, Brand Weekly Share, Raw JSON |
| `gfk_sku_ranking` | GfK SKU Summary, SKU Ranking, Raw JSON |
| `gfk_combined` | GfK Market Overview, EC Brand Performance, EC Category Performance, Cross Insights, Raw JSON |

---

## 5. AI Model Araçları (Tools)

Sistem Ollama üzerinde `llama3.1` modelini kullanır. Aşağıdaki tool'lar tanımlıdır:

| Tool | Kaynak Dosya | Açıklama |
|---|---|---|
| `analyze_gfk_market_share` | `functions/gfk_analyzer.py` | Kategori bazında pazar payı, WoW büyüme, MM Ihs % |
| `analyze_gfk_brand_performance` | `functions/gfk_analyzer.py` | Marka bazında haftalık MM internet satış payı |
| `analyze_gfk_sku_ranking` | `functions/gfk_analyzer.py` | Ürün grubu bazında satış sıralaması |
| `analyze_gfk_combined` | `functions/gfk_analyzer.py` | GfK + ecommerce cross analizi |
| `analyze_ecommerce_sample` | `functions/analytics.py` | 200 satır ecommerce veri analizi (DuckDB + Ollama SQL) |
| `analyze_funnel_master` | `functions/funnel_master.py` | Funnel drop-off analizi |
| `analyze_cross_performance` | `functions/cross_analyzer.py` | Google Trends + fiyat + PPC cross analiz |
| `generate_category_insight` | `functions/insights.py` | Kategori insight, sektörel yorum |
| `generate_price_competition_from_uploaded_inputs` | `functions/price_competition.py` | GTIN bazlı Merchant benchmark analizi |
| `calculate_business_metric` | `functions/business_calculator.py` | Ortalama, toplam, top/bottom hesaplama |
| `execute_recommended_action` | `functions/action_executor.py` | Aksiyon planı & Excel çıktısı |

---

## 6. Kategori Mapping Tablosu

| GfK Kategori | İç Sistem Kategorisi |
|---|---|
| Smartphones | Cep Telefonları, Telefon, Akıllı Telefon |
| COMPUTER HW | Bilgisayar, Laptop, Dizüstü, PC |
| SDA | Küçük Ev Aletleri, Kişisel Bakım |
| MDA | Büyük Ev Aletleri, Çamaşır Makinesi, Bulaşık Makinesi |
| CLIMATE SDA | Klima, İklim, Vantilatör |
| PTV/FLAT | Televizyon, TV, OLED, QLED |
| Headphones & Headsets | Kulaklık, Headphone, Headset |
| COMPUTER ACCESSORIES | Bilgisayar Aksesuarları, Mouse, Klavye, Monitör |
| CORE WEARABLES | Giyilebilir, Akıllı Saat, Smartwatch |
| VACUUM CLEANERS | Süpürge, Robot Süpürge, Dikey Süpürge |
| WASHING MACHINES | Çamaşır Makinesi |
| DISHWASHERS | Bulaşık Makinesi |
| COOLING | Buzdolabı, Derin Dondurucu |
| MONITORS | Monitör |
| MOBILE COMPUTING | Tablet, Laptop, Notebook |
| AIR CONDITIONERS | Klima, Split Klima |
| HOT BEVER.MAKERS | Kahve Makinesi, Kettle |

> **NOT:** Eşleştirme hem GfK kategori adı üzerinden hem de marka adı (SAMSUNG, APPLE vb.) üzerinden yapılır. Yeni eşleştirmeler `functions/gfk_analyzer.py` içindeki `GFK_CATEGORY_MAP` sözlüğüne eklenir.

---

## 7. Güvenlik & Deployment

### 7.1 Sistem Gereksinimleri

| Bileşen | Versiyon | Açıklama |
|---|---|---|
| Python | 3.9+ | FastAPI backend |
| Ollama | Son stabil | Local LLM çalıştırıcı |
| llama3.1 | — | Ana AI modeli (`MODEL = "llama3.1"`) |
| FastAPI | — | Web framework |
| pandas | — | Veri işleme |
| openpyxl | — | Excel okuma/yazma |
| duckdb | — | In-memory SQL engine |

### 7.2 Kurulum

```bash
# Bağımlılıkları yükle
pip install -r requirements.txt

# Ollama model indir
ollama pull llama3.1

# Uygulamayı başlat (geliştirme modu)
uvicorn main:app --reload --host 0.0.0.0 --port 8000

# Uygulamayı başlat (üretim modu)
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 1
```

> **ÖNEMLİ:** `--workers 1` kullanılmalıdır. Çoklu worker durumunda `lru_cache` bellekte paylaşılmaz ve her worker GfK dosyalarını yeniden yükler.

### 7.3 Ortam Değişkenleri

Şu an zorunlu ortam değişkeni yoktur. Aşağıdaki sabitler `main.py` içinde tanımlıdır:

```python
OLLAMA_URL = "http://localhost:11434/api/chat"  # Ollama adresi
MODEL = "llama3.1"                               # Kullanılan model
LOG_TOOL_JSON_TO_TERMINAL = True                 # Terminal log modu
SHOW_RAW_JSON_IN_UI = False                      # UI'da ham JSON göster
```

### 7.4 Ollama Konfigürasyonu

Ollama varsayılan olarak `http://localhost:11434` adresinde çalışır. Farklı bir host için `main.py` ve `functions/analytics.py` içindeki `OLLAMA_URL` sabitini güncelleyin.

```bash
# Model listesini kontrol et
ollama list

# llama3.1 yüklü değilse
ollama pull llama3.1

# Ollama servis durumu
ollama ps
```

### 7.5 Güvenlik Notları

- Uygulama, kullanıcı kimlik doğrulaması içermez. İnternet'e açık deploy'da reverse proxy (nginx, Caddy) + authentication katmanı ekleyin.
- Upload endpoint'i yalnızca `.xlsx` ve `.csv` dosyalarını kabul eder.
- DuckDB SQL güvenliği: Yalnızca `SELECT` sorgularına izin verilir (`is_safe_select()` fonksiyonu).
- GfK dosyaları hassas pazar verisi içerir; `data/` dizinini `.gitignore`'a ekleyin.

---

## 8. Sorun Giderme

### GfK verisi yüklenemiyor
```
FileNotFoundError: data/GfK_Leaderpanel.xlsx bulunamadı
```
→ Dosyanın `data/` dizinine kopyalandığını ve dosya adının tam eşleştiğini kontrol edin.

### GfK kategorisi bulunamıyor
```json
{"error": "AIR CONDITIONERS kategorisi PW vs CW verisinde bulunamadı"}
```
→ GfK dosyasının `PW vs. CW` sheet'inde o kategorinin bulunduğunu kontrol edin. Farklı hafta verisi içeren yeni bir dosya ise kategoriler değişmiş olabilir.

### Marka/kategori soruları GfK'ya yönlendirilmiyor
→ `should_use_gfk_*` fonksiyonlarındaki keyword listesini kontrol edin. Türkçe karakter normalizasyonu `tr_map` üzerinden yapılır.

### lru_cache stale veri problemi
→ Dosya güncellemesi sonrası sunucuyu yeniden başlatın: `uvicorn main:app --reload`

### Ollama bağlantı hatası
```
Ollama bağlantı hatası: Connection refused
```
→ `ollama serve` komutunu çalıştırın. Farklı port kullanıyorsanız `OLLAMA_URL` sabitini güncelleyin.

### Excel export boş sheet
→ `LAST_TOOL_RESULT_JSON` değişkeni boşsa export çalışmaz. Önce bir analiz sorusu sorun, ardından export yapın.

---

*Son güncelleme: 27 Haziran 2026 — GfK Leaderpanel v2.0 entegrasyonu*
