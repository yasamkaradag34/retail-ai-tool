# 🏢 DataProvido — On-Premise & Local IT Kurulum Kılavuzu
> **Gizlilik Öncelikli (Privacy-First) Yerel Perakende ve E-Ticaret Yapay Zeka Analitiği**

Bu doküman, kurum içi **Bilgi Teknolojileri (IT / DevOps)** ekiplerinin DataProvido platformunu şirket içi sunuculara (On-Premise) veya yerel bilgisayarlara **%100 veri gizliliği** ile kurabilmesi için hazırlanmıştır.

---

## 🎯 Temel Mimari ve Çalışma Modları

DataProvido, kurumunuzun güvenlik ve altyapı politikalarına göre iki farklı modda çalışabilir:

```
                          ┌────────────────────────────────┐
                          │     DataProvido Core Engine    │
                          │   (FastAPI + Analytics Engine) │
                          └───────────────┬────────────────┘
                                          │
                  ┌───────────────────────┴───────────────────────┐
                  ▼                                               ▼
       [MOD A: %100 YEREL & OFFLINE]                   [MOD B: HIZLI BULUT API]
       (Önerilen Kurumsal Mod)                         (Test / Düşük Donanım)
 ──────────────────────────────────────────     ────────────────────────────────────
  • Motor: Ollama + LLaMA 3.1 8B                 • Motor: Groq Cloud API
  • Ağ: Tamamen Offline / Kapalı Devre           • Ağ: Dış API Bağlantısı
  • Veri Akışı: 0 Bayt dışarı çıkmaz             • Donanım İhtiyacı: Minimum
  • KVKK / GDPR / ISO 27001: %100 Uyumlu         • Hız: Ultra Hızlı
```

---

## 💻 Sistem Gereksinimleri

### Mod A (Yerel Ollama / %100 Offline)
* **İşletim Sistemi:** Linux (Ubuntu 20.04+, Debian, RHEL), macOS (Apple Silicon M1/M2/M3/M4) veya Windows 11 (WSL2).
* **RAM:** Minimum 16 GB (32 GB önerilir).
* **GPU (Opsiyonel ama Önerilen):** NVIDIA GPU (8 GB+ VRAM) veya Apple Silicon Unified Memory.
* **Disk Alanı:** Minimum 15 GB boş alan (LLaMA 3.1 modeli ~4.7 GB).

### Mod B (Groq API Hibrit Mod)
* **İşletim Sistemi:** Herhangi bir OS.
* **RAM:** 4 GB.
* **GPU:** Gerekli değil.

---

## 🚀 1. YÖNTEM: Docker Compose ile Tek Tıkla Kurulum (Önerilen)

Kurumunuzda **Docker** ve **Docker Compose** kurulu ise tüm sistemi tek bir komutla ayağa kaldırabilirsiniz.

### Adım 1: Projeyi Klonlayın
```bash
git clone https://github.com/yasamkaradag34/retail-ai-tool.git
cd retail-ai-tool
```

### Adım 2: Konfigürasyonu Belirleyin
`docker-compose.yml` dosyasını açıp çalışma modunuzu seçin:

* **Mod A (%100 Yerel - Varsayılan):**
  ```yaml
  environment:
    - LLM_BACKEND=ollama
    - OLLAMA_URL=http://ollama-engine:11434/api/chat
    - OLLAMA_MODEL=llama3.1
  ```
* **Mod B (Groq API):**
  ```yaml
  environment:
    - LLM_BACKEND=groq
    - GROQ_API_KEY=your_groq_api_key_here
    - GROQ_MODEL=llama-3.1-8b-instant
  ```

### Adım 3: Sistemi Başlatın
```bash
docker-compose up -d
```

### Adım 4: Mod A için Yerel Modeli İndirin (Tek Seferlik)
```bash
docker exec -it dataprovido-ollama ollama pull llama3.1
```

🎉 **Sistem Hazır!** Kurum içi tarayıcınızdan **`http://localhost:8000`** veya sunucu IP adresiniz üzerinden erişebilirsiniz.

---

## 🛠️ 2. YÖNTEM: Manuel / Doğrudan Sunucu Kurulumu

Docker kullanmak istemeyen IT ekipleri için Python ortamında doğrudan çalıştırma adımları:

### 1. Ollama Kurulumu (Mod A için):
* **Linux/macOS:** `curl -fsSL https://ollama.com/install.sh | sh`
* **Modeli İndirin:** `ollama pull llama3.1`
* **Servisi Başlatın:** `ollama serve` (varsayılan port: `11434`)

### 2. Python Ortamını Hazırlayın:
```bash
python3 -m venv venv
source venv/bin/activate   # Windows için: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Ortam Değişkenleri (.env) Tanımlayın:
```bash
export LLM_BACKEND=ollama
export OLLAMA_URL=http://localhost:11434/api/chat
export OLLAMA_MODEL=llama3.1
```

### 4. Servisi Başlatın:
```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

---

## 📊 Veri Kaynakları & Entegrasyon

DataProvido, şirketinizin verilerini `data/` dizini altından veya arayüz üzerinden dinamik olarak okur:

| Dosya Adı | Açıklama |
|---|---|
| `company_product_input.xlsx` | Ürün, stok, maliyet, kategori ve satış metrikleri |
| `stok.xlsx` / `orders.xlsx` | Günlük sipariş ve anlık stok hareketleri |
| `GfK_Leaderpanel.xlsx` | Pazar payı ve sektör kıyaslama verileri |
| `google_trends_seasonal_3y.xlsx` | Sezonsallık ve arama trend verileri |

> 🔒 **Veri Güvenliği Notu:** Tüm veri analizleri Python Pandas & DuckDB motorları ile yerel RAM üzerinde çalıştırılır. Hiçbir veri parçası disk dışına veya üçüncü taraf bulut sunucularına aktarılmaz.

---

## 🛡️ Güvenlik, KVKK ve GDPR Uyumluluğu

* **Sıfır Dış Veri Transferi (Zero Data Leak):** Mod A kullanımında internet bağlantısı kapatılsa dahi sistem tam işlevsellikle çalışır.
* **Role-Based Access:** Kurum içi Nginx / Traefik ters vekil (reverse proxy) arkasında LDAP/SSO ile entegre edilebilir.
* **Audit Logs:** Tüm sorgular ve analiz geçmişi yerel log dosyalarında tutulur.

---

## 📞 Destek & İletişim
Kurulum ve entegrasyon süreçlerinde özel destek için:
* **E-posta:** support@dataprovido.com
* **Web:** [https://dataprovido.com](https://dataprovido.com)
