TOOLS = [
    # --- STOK ---
        {
        "type": "function",
        "function": {
            "name": "analyze_ecommerce_sample",
            "description": "200 satırlık sample e-ticaret datası üzerinden stok, kategori, SKU, satış, revenue, fiyat, C2D, B2D, funnel, PDP, A2C, checkout, OOS, overstock ve stock coverage sorularını cevaplar. Stok, satış, funnel, kategori, marka, SKU, revenue, C2D, B2D gibi geniş analiz sorularında varsayılan olarak bunu kullan.",
            "parameters": {
                "type": "object",
                "properties": {
                    "question": {
                        "type": "string",
                        "description": "Kullanıcının doğal dilde sorduğu e-ticaret analiz sorusu"
                    }
                },
                "required": ["question"]
            }
        }
    },
        {
        "type": "function",
        "function": {
            "name": "execute_recommended_action",
            "description": "Önerilen aksiyonları çalıştırır. Replenishment planı, stok riski olan SKU listesi, kampanya/görünürlük planı, riskli segment detaylandırma ve Excel'e indirilecek aksiyon çıktıları üretir.",
            "parameters": {
                "type": "object",
                "properties": {
                    "question": {
                        "type": "string",
                        "description": "Kullanıcının aksiyon isteği. Örn: C2D/B2D güçlü ama stok riski olan SKU'lar için replenishment planı yap."
                    }
                },
                "required": ["question"]
            }
        }
    },    
        {
        "type": "function",
        "function": {
            "name": "generate_price_competition_from_uploaded_inputs",
            "description": "Şirketin GTIN'li ürün/funnel input datasını Merchant Center benchmark datasıyla GTIN üzerinden join ederek fiyat rekabeti insight'ı üretir. Internal benchmark kullanmaz.",
            "parameters": {
                "type": "object",
                "properties": {
                    "category": {
                        "type": "string",
                        "description": "Analiz edilecek kategori, marka veya ürün grubu. Örn: Mobile, Tablet, Fashion, genel"
                    },
                    "period_name": {
                        "type": "string",
                        "description": "Analiz dönemi adı. Örn: Tatil dönemi, Son 7 gün, selected_period"
                    }
                },
                "required": ["category"]
            }
        }
    },
            {
        "type": "function",
        "function": {
            "name": "calculate_business_metric",
            "description": "Şirket Excel datası üzerinde ortalama, toplam, adet, minimum, maksimum, medyan, top/bottom, marka/kategori kırılımı gibi matematiksel business hesaplamaları yapar. APPLE ürünlerinin ortalama fiyatı, GSM kategorisinin toplam revenue'u, B2D ortalaması, marka bazında PDP gibi sorularda kullanılır.",
            "parameters": {
                "type": "object",
                "properties": {
                    "question": {
                        "type": "string",
                        "description": "Kullanıcının doğal dilde sorduğu matematiksel business sorusu"
                    }
                },
                "required": ["question"]
            }
        }
    },
        {
        "type": "function",
        "function": {
            "name": "generate_category_insight",
            "description": "Kategori, sektör ve dönem bazlı e-ticaret insight raporu üretir. Stok, fiyat, funnel, kanal, traffic, C2D, B2D, revenue ve aksiyon önerilerini birlikte analiz eder.",
            "parameters": {
                "type": "object",
                "properties": {
                    "category": {
                        "type": "string",
                        "description": "Analiz edilecek kategori, marka, kanal veya ürün grubu. Örn: Mobile, Tablet, Telefon, Headphones, IT Accessories, genel"
                    },
                    "sector": {
                        "type": "string",
                        "description": "Sektör tipi. Örn: consumer_electronics, fashion, fmcg, marketplace_general"
                    },
                    "period_name": {
                        "type": "string",
                        "description": "Analiz dönemi adı. Örn: Tatil dönemi, 21-31 Mayıs, selected_period"
                    }
                },
                "required": ["category"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_stock_level",
            "description": "Belirli bir ürünün stok miktarını döner",
            "parameters": {
                "type": "object",
                "properties": {
                    "product_id": {"type": "string", "description": "Ürün kodu, örn: SKU-001"}
                },
                "required": ["product_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_all_stock",
            "description": "Tüm ürünlerin stok listesini getirir. Karşılaştırma, en fazla/az, toplam stok gibi sorularda kullan.",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_daily_sales_report",
            "description": "Günlük satış raporunu getirir",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "check_low_stock",
            "description": "Stoku az olan ürünleri listeler. Eşik belirtilmezse varsayılan 10 kabul edilir.",
            "parameters": {
                "type": "object",
                "properties": {
                    "threshold": {"type": "integer", "description": "Minimum stok eşiği, örn: 10"}
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_out_of_stock",
            "description": "Stoku tamamen tükenmiş ürünleri listeler",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_product_by_name",
            "description": "Ürün adıyla arama yapar. SKU kodu bilinmiyorsa kullan.",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Aranacak ürün adı, örn: Elma"}
                },
                "required": ["name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_stock_value",
            "description": "Tüm ürünlerin toplam stok değerini hesaplar (adet x fiyat)",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_stock",
            "description": "Belirli bir ürünün stok miktarını günceller",
            "parameters": {
                "type": "object",
                "properties": {
                    "product_id": {"type": "string", "description": "Ürün kodu, örn: SKU-001"},
                    "quantity": {"type": "integer", "description": "Yeni stok miktarı"}
                },
                "required": ["product_id", "quantity"]
            }
        }
    },
    # --- SİPARİŞ ---
    {
        "type": "function",
        "function": {
            "name": "get_order_status",
            "description": "Belirli bir siparişin durumunu getirir",
            "parameters": {
                "type": "object",
                "properties": {
                    "order_id": {"type": "string", "description": "Sipariş numarası, örn: ORD-001"}
                },
                "required": ["order_id"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_all_orders",
            "description": "Tüm siparişleri listeler",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_pending_orders",
            "description": "Bekleyen siparişleri listeler",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_orders_by_customer",
            "description": "Belirli bir müşterinin siparişlerini getirir",
            "parameters": {
                "type": "object",
                "properties": {
                    "customer": {"type": "string", "description": "Müşteri adı, örn: Ali"}
                },
                "required": ["customer"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "update_order_status",
            "description": "Sipariş durumunu günceller",
            "parameters": {
                "type": "object",
                "properties": {
                    "order_id": {"type": "string", "description": "Sipariş numarası, örn: ORD-001"},
                    "status": {"type": "string", "description": "Yeni durum: Beklemede, Kargoda, Teslim Edildi, İptal"}
                },
                "required": ["order_id", "status"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_todays_orders",
            "description": "Bugünkü siparişleri listeler",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    # --- RAPORLAMA ---
    {
        "type": "function",
        "function": {
            "name": "get_total_revenue",
            "description": "Toplam geliri hesaplar (teslim edilen siparişler)",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_best_selling_product",
            "description": "En çok satan ürünü getirir",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_sales_summary",
            "description": "Genel satış özetini getirir. Toplam sipariş, teslim, iptal ve gelir bilgisi.",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_low_stock_report",
            "description": "Kritik stok seviyesindeki ürünlerin raporunu getirir",
            "parameters": {"type": "object", "properties": {}}
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_stock_turnover",
            "description": "Ürün bazında satış ve mevcut stok karşılaştırması yapar",
            "parameters": {"type": "object", "properties": {}}
        }
    }
]