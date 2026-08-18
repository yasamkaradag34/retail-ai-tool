# -*- coding: utf-8 -*-

TECH_RETAIL_NORMS = {
    "mobile": {
        "display_name": "Cep Telefonu",
        "general_behavior": (
            "Akıllı telefonlar, teknoloji perakendeciliğinde en yüksek fiyat karşılaştırma "
            "oranına sahip kategoridir. Tüketiciler, satın alma öncesinde Akakçe, Cimri veya "
            "farklı satıcılar üzerinden fiyatları agresif şekilde kontrol ederler. Ayrıca yüksek "
            "sepet tutarı ve taksit sınırlamaları nedeniyle ödeme adımında alışveriş kredileri "
            "veya alternatif finansal çözümler (taksit seçenekleri vb.) kritik öneme sahiptir. "
            "Tüketiciler genellikle en az 128GB/256GB depolama ve yüksek kamera kalitesini önceler."
        ),
        "rules": [
            {
                "id": "mobile_price_sensitivity",
                "condition": lambda m: m.get("avg_price_gap_pct", 0) > 3 and m.get("b2d_delta_pct", 0) < 0,
                "insight": (
                    "Fiyat rekabetinde pazar benchmark'ının %3 ve daha üzerine çıkılması, cep telefonlarındaki "
                    "yüksek fiyat duyarlılığı ve yoğun karşılaştırma davranışı nedeniyle satın alma dönüşümünü "
                    "(B2D) doğrudan ve negatif etkilemiştir. Tüketiciler sepeti terk ederek daha ucuz alternatiflere yönelmiş görünüyor."
                )
            },
            {
                "id": "mobile_payment_friction",
                "condition": lambda m: m.get("c2d_pct", 0) > 15 and m.get("b2d_pct", 0) < 3,
                "insight": (
                    "Sepete ekleme oranının (C2D) yüksek olmasına rağmen sipariş dönüşümünün (B2D) düşük olması, "
                    "taksit sınırları veya ödeme adımlarındaki finansman zorluklarını işaret eder. Kullanıcılar "
                    "satın alma niyetiyle ürünü sepete eklemekte, ancak ödeme aşamasındaki taksit kısıtlamaları "
                    "veya kredi seçeneklerinin yetersizliği nedeniyle siparişi tamamlayamamaktadır."
                )
            },
            {
                "id": "mobile_stock_loss",
                "condition": lambda m: m.get("critical_stock_sku_count", 0) > 0 and m.get("pdp_delta_pct", 0) > 5,
                "insight": (
                    "Cep telefonlarında marka sadakati ve belirli model talebi çok yüksektir. Görüntüleme (PDP) "
                    "artarken kritik stok seviyesindeki ürünlerin bulunması, tüketicilerin alternatif bir modele "
                    "geçmek yerine doğrudan başka bir satıcıya gitmesine yol açar; bu durum doğrudan ciro kaybına sebep olmaktadır."
                )
            }
        ]
    },
    
    "tablet": {
        "display_name": "Tablet",
        "general_behavior": (
            "Tablet alıcıları, cihazları genellikle eğitim, hafif üretkenlik (klavyeli/kalemli 2-in-1 modeller) "
            "veya medya tüketimi (video, oyun) için tercih etmektedir. Akıllı telefonlara kıyasla yenileme "
            "döngüsü oldukça uzundur (ortalama 3-4 yıl). Bu nedenle talep, okula dönüş (Back-to-School) "
            "veya büyük indirim dönemlerinde (Kasım vb.) yoğunlaşır. Tüketiciler klavye ve kalem gibi "
            "ekosistem tamamlayıcı aksesuarlarla birlikte satın almaya yatkındır."
        ),
        "rules": [
            {
                "id": "tablet_seasonal_slump",
                "condition": lambda m: m.get("pdp_delta_pct", 0) < -10 and m.get("revenue_delta_pct", 0) < 0,
                "insight": (
                    "Tablet kategorisindeki talep daralması, bu ürün grubunun uzun yenileme döngüsü (3-4 yıl) ve "
                    "okula dönüş gibi özel sezonlar dışındaki durağanlığı ile doğrudan ilişkilidir. Dönemsel kampanyalar "
                    "veya aksesuar/öğrenci paketleri (bundle) ile yapay talep yaratılması gerekebilir."
                )
            },
            {
                "id": "tablet_accessory_bundle_opportunity",
                "condition": lambda m: m.get("c2d_pct", 0) > 12 and m.get("b2d_delta_pct", 0) < 0,
                "insight": (
                    "Tabletlerde sepete atma niyetinin güçlü olmasına rağmen satın almaya dönüşün zayıf kalması, "
                    "klavye, kılıf veya kalem gibi tamamlayıcı aksesuarların eksikliğinden veya set halinde "
                    "fiyatlandırılmamasından kaynaklanabilir. Aksesuar hediyeli veya indirimli paket teklifleri (bundle) "
                    "satın alma kararını hızlandırabilir."
                )
            }
        ]
    },
    
    "laptop": {
        "display_name": "Bilgisayar / Laptop",
        "general_behavior": (
            "Bilgisayar alıcıları yüksek araştırmacı davranış sergiler; ürün detay sayfalarındaki (PDP) işlemci, "
            "RAM, ekran kartı ve depolama gibi teknik özellikleri derinlemesine incelerler. Yenileme döngüleri "
            "en uzun kategorilerden biridir (4-5 yıl). Hibrit çalışma modelleri, dijital tasarım ve yüksek performanslı "
            "oyun (gaming) ana satın alma motivasyonlarıdır. Güvenilirlik, servis ağı ve garanti süresi kararda etkilidir."
        ),
        "rules": [
            {
                "id": "laptop_bounce_friction",
                "condition": lambda m: m.get("bounce_rate_pct", 0) > 40 and m.get("pdp_delta_pct", 0) > 0,
                "insight": (
                    "Bilgisayar kategorisinde ürün detay sayfası (PDP) trafiği artmasına rağmen hemen çıkma oranının "
                    "(bounce rate) yüksek seyretmesi, tüketicilerin aradıkları detaylı teknik özellikleri (RAM, CPU, ekran kartı vb.) "
                    "PDP üzerinde net olarak bulamadıklarını veya fiyat/donanım dengesinden memnun kalmadıklarını gösterir."
                )
            },
            {
                "id": "laptop_high_value_decision",
                "condition": lambda m: m.get("c2d_pct", 0) > 10 and m.get("b2d_pct", 0) < 2,
                "insight": (
                    "Bilgisayarlar yüksek bütçeli (AOV) yatırım ürünleri olduğu için tüketiciler sepete ekledikten sonra "
                    "ortalama 5-7 gün boyunca diğer platformlardaki yorumları ve garanti seçeneklerini incelerler. Karar süresini "
                    "kısaltmak için 'ek garanti', 'koşulsuz iade' veya 'hızlı kargo' gibi güven verici unsurlar öne çıkarılmalıdır."
                )
            }
        ]
    },
    
    "kulaklik": {
        "display_name": "Kulaklık ve Ses",
        "general_behavior": (
            "Kulaklıklar, spor, seyahat ve ofis içi kullanım gibi yaşam tarzı ihtiyaçlarına göre impulsif (anlık) "
            "satın almaya en yatkın teknoloji kategorilerinden biridir. Aktif Gürültü Engelleme (ANC), ergonomi/tasarım "
            "ve pil ömrü en önemli karar kriterleridir. Marka sadakati akıllı telefonlara göre düşüktür; trendler ve "
            "sosyal medya etkileşimi satın almada büyük rol oynar. Hijyen kuralları nedeniyle iade süreçleri hassastır."
        ),
        "rules": [
            {
                "id": "audio_impulse_loss",
                "condition": lambda m: m.get("avg_stock_coverage_days", 0) < 5 and m.get("avg_stock_coverage_days", 0) > 0,
                "insight": (
                    "Kulaklıklar hızlı tüketilen ve anlık hediyeleşme amacıyla da sık tercih edilen bir gruptur. "
                    "Stok coverage gününün çok düşük olması, tüketicilerin satın alma anında anında kargo veya anında stok "
                    "beklentisini karşılayamayacağı için trafiğin hızla rakip satıcılara kaymasına yol açmaktadır."
                )
            },
            {
                "id": "audio_trend_decay",
                "condition": lambda m: m.get("pdp_delta_pct", 0) < -15 and m.get("a2c_delta_pct", 0) < -15,
                "insight": (
                    "Kulaklık ve ses kategorisinde hem PDP hem de sepete atma oranlarındaki sert düşüş, trend etkisinin "
                    "kaybolduğuna veya yeni rakip modellerin pazara giriş yaparak tüketici odağını üzerine çektiğine işaret eder. "
                    "Sosyal medya görünürlüğü veya influencer iş birlikleri ile ilgi yeniden canlandırılmalıdır."
                )
            }
        ]
    },
    
    "wearable": {
        "display_name": "Giyilebilir Teknoloji",
        "general_behavior": (
            "Akıllı saat ve bileklik tüketicileri, sağlık takibi (uyku analizi, nabız, kandaki oksijen), spor modları "
            "ve pil dayanıklılığını (en az 5-7 gün şarj süresi) ön planda tutar. Tasarım, kordon çeşitliliği ve giysi "
            "uyumu nedeniyle estetik faktörler de kararda etkilidir. Anneler Günü, Babalar Günü ve Yılbaşı gibi dönemlerde "
            "hediyelik eşya olarak talep patlaması yaşar."
        ),
        "rules": [
            {
                "id": "wearable_compatibility_concern",
                "condition": lambda m: m.get("bounce_rate_pct", 0) > 35 and m.get("c2d_pct", 0) < 8,
                "insight": (
                    "Giyilebilir teknoloji ürünlerinde hemen çıkma oranının yüksek olması, tüketicilerin akıllı telefonları "
                    "ile olan işletim sistemi uyumluluğundan (iOS/Android uyumu) veya şarj süresi detayından emin olamadıklarını "
                    "gösterir. PDP üzerinde uyumluluk bilgileri daha görünür yapılmalıdır."
                )
            }
        ]
    },
    
    "vacuum": {
        "display_name": "Ev Teknolojileri / Robot Süpürge",
        "general_behavior": (
            "Robot süpürge ve dikey süpürge alıcıları, satın alma kararı vermeden önce kullanıcı yorumlarına, "
            "çekiş gücüne (Pa), haritalama teknolojisine ve garanti/servis güvencesine çok yüksek önem verirler. "
            "Karar verme süreçleri görece uzundur. Ev içi kolaylık ve zaman tasarrufu ana motivasyondur."
        ),
        "rules": [
            {
                "id": "vacuum_review_seeking",
                "condition": lambda m: m.get("c2d_pct", 0) > 8 and m.get("b2d_pct", 0) < 1.5,
                "insight": (
                    "Süpürge kategorisinde sepete ekleme sonrası satın alma kararı uzamaktadır. Tüketiciler "
                    "kronik arıza, haritalama performansı veya servis kalitesi hakkında dış sitelerden ve "
                    "kullanıcı yorumlarından teyit almaya çalışmaktadır. PDP altında kullanıcı değerlendirmelerini "
                    "öne çıkarmak dönüşümü artıracaktır."
                )
            }
        ]
    },
    
    "accessory": {
        "display_name": "Aksesuar",
        "general_behavior": (
            "Şarj aletleri, kablolar, kılıflar ve koruyucu camlar gibi aksesuarlar, ana cihazların (telefon, tablet vb.) "
            "yanında çapraz satış (cross-sell) olarak alınır. Tüketiciler bu grupta düşük fiyat hassasiyeti sergilerler; "
            "ancak ürünün ana cihaza tam uyumu ve anında teslim edilmesi kararlarındaki en kritik etkendir."
        ),
        "rules": [
            {
                "id": "accessory_cross_sell_loss",
                "condition": lambda m: m.get("c2d_pct", 0) < 5,
                "insight": (
                    "Aksesuar kategorisinde sepete ekleme oranının düşük kalması, bu ürünlerin ana cihazlarla "
                    "(telefon, bilgisayar vb.) sepet aşamasında yeterince çapraz satış (cross-sell) olarak "
                    "önerilmediğini gösterir. Sepet veya ödeme sayfasında 'birlikte al' paketleri sunulmalıdır."
                )
            }
        ]
    }
}


def get_behavior_for_category(category_name: str) -> dict:
    """
    Kategori adına göre eşleşen teknoloji perakendeciliği davranış modelini döner.
    """
    from functions.insights import normalize_text
    
    norm_cat = normalize_text(category_name)
    
    # Eşleşme sözlüğü
    mapping = {
        "mobile": ["gsm", "telefon", "cep", "cep telefonlari", "cep telefonları", "iphone", "galaxy", "samsung", "huawei", "xiaomi"],
        "tablet": ["tablet", "tabletler", "ipad"],
        "laptop": ["laptop", "bilgisayar", "pc", "notebook", "dizustu", "dizüstü", "macbook", "asus", "lenovo", "dell", "hp"],
        "kulaklik": ["kulaklik", "kulaklık", "airpods", "headphone", "headphones", "jbl", "ses", "hoparlor", "hoparlör"],
        "wearable": ["wearable", "giyilebilir", "saat", "akilli saat", "akıllı saat", "bileklik", "smartwatch", "apple watch"],
        "vacuum": ["supurge", "süpürge", "roborock", "dyson", "robot supurge", "robot süpürge", "vacuum"],
        "accessory": ["aksesuar", "sarj", "şarj", "kablo", "kilif", "kılıf", "koruyucu", "cam", "adaptor", "adaptör"]
    }
    
    for key, aliases in mapping.items():
        if norm_cat == key or any(alias in norm_cat for alias in aliases):
            return TECH_RETAIL_NORMS.get(key)
            
    return None
