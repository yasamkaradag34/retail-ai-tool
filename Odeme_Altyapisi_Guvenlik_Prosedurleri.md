1**Ödeme Altyapısı Güvenlik Prosedürleri**

*Sızma Önleme, Erişim Kontrolü ve Kapsamlı Güvenlik Test Rehberi*

Bölüm 1 --- Ödeme Altyapısı Güvenlik Prosedürleri

Ödeme sistemi içeren bir altyapıda güvenlik katmanlı (defense in depth)
olmalı. Aşağıda alan alan adım adım bir çerçeve verilmiştir.

1\. Ağ ve Sunucu Güvenliği

-   **Segmentasyon:** Ödeme işleme sunucularını (CDE - Cardholder Data
    Environment) diğer tüm sistemlerden ayrı bir ağ segmentinde tutun.
    Web sunucusu, uygulama sunucusu, veritabanı katmanlarını ayrı
    VLAN/subnet\'lere bölün.

-   **Firewall kuralları:** Varsayılan olarak her şeyi reddet
    (deny-all), sadece gerekli portları/IP\'leri açın. Giden trafiği de
    kısıtlayın (egress filtering) --- sızma sonrası veri dışarı çıkışını
    zorlaştırır.

-   **Bastion/Jump host:** Sunuculara doğrudan SSH/RDP erişimini
    kapatın, sadece bastion host üzerinden, MFA ile erişim verin.

-   **Port ve servis minimizasyonu:** Kullanılmayan servisleri, portları
    kapatın. Gereksiz yazılımları kaldırın.

-   **Patch yönetimi:** OS ve tüm servisler için düzenli,
    otomatikleştirilmiş güncelleme süreci kurun (kritik CVE\'ler için
    24-72 saat içinde yama).

-   **DDoS koruması:** CDN/WAF önünde DDoS mitigasyonu (Cloudflare, AWS
    Shield vb.)

2\. Kimlik ve Erişim Yönetimi (IAM)

-   **En az yetki ilkesi (least privilege):** Her kullanıcı/servis
    sadece işini yapmak için gerekli minimum yetkiye sahip olsun.

-   **MFA zorunlu:** Tüm yönetim panellerine, SSH\'a, cloud
    konsollarına, domain yönetim paneline erişimde MFA (tercihen donanım
    anahtarı/TOTP, SMS değil).

-   **Rol bazlı erişim kontrolü (RBAC):** Geliştirici, DevOps, finans,
    destek ekiplerinin erişim seviyeleri net ayrılsın.

-   **Ayrıcalıklı hesap yönetimi (PAM):** Root/admin şifreleri bir
    \"vault\" sisteminde (HashiCorp Vault, CyberArk) tutulsun,
    çıkış/giriş loglansın, gerektiğinde \"just-in-time\" erişim
    verilsin.

-   **Düzenli erişim gözden geçirme:** 3 ayda bir kim hangi sisteme
    erişebiliyor gözden geçirin, işten ayrılanların erişimini anında
    kapatın (offboarding checklist).

-   **Servis hesapları:** İnsan olmayan (API/servis) hesaplar için ayrı,
    sıkı kısıtlanmış izinler ve düzenli rotasyon.

3\. Domain ve DNS Güvenliği

-   **Registrar hesabı kilidi:** Domain registrar hesabında \"registrar
    lock\" / \"transfer lock\" aktif olsun.

-   **MFA + ayrı e-posta:** Registrar ve DNS yönetim hesapları için
    kurumsal, kritik e-posta adresi kullanın, bu e-postaya erişim de
    sıkı korunsun.

-   **DNSSEC:** DNS kayıtlarının bütünlüğünü korumak için DNSSEC aktif
    edin.

-   **CAA kayıtları:** Hangi CA\'ların SSL sertifikası verebileceğini
    kısıtlayın.

-   **DNS değişiklik logları ve onay süreci:** Kritik DNS değişiklikleri
    (MX, A, CNAME) çift onaylı (four-eyes) süreçle yapılsın, değişiklik
    anında alert gitsin.

4\. Veritabanı Güvenliği

-   **Şifreleme:** Verinin hem \"rest\" hem \"transit\" halinde
    şifrelenmesi (TDE, TLS 1.2+).

-   **Kart verisi:** Mümkünse kart verisini hiç tutmayın ---
    tokenization/PCI-DSS onaylı ödeme sağlayıcısına (iyzico, Stripe vb.)
    devredin. Zorunluysa PCI-DSS Seviye 1 gereksinimlerine tam uyum.

-   **Erişim izolasyonu:** Veritabanına doğrudan internet erişimi kapalı
    olsun, sadece uygulama sunucusu üzerinden erişilsin.

-   **Sorgu/erişim loglama:** Kim, ne zaman, hangi tabloya eriştiğini
    loglayın (özellikle hassas tablolar için).

-   **Düzenli, şifreli, test edilmiş yedekleme:** 3-2-1 kuralı (3 kopya,
    2 farklı ortam, 1 offsite), yedekleri düzenli restore testinden
    geçirin.

-   **Statik/dinamik veri maskeleme:** Test/staging ortamlarında gerçek
    müşteri/kart verisi kullanılmasın, maskelensin.

5\. Uygulama Güvenliği

-   **Secrets yönetimi:** API key, DB şifresi, sertifika gibi bilgiler
    kod içine ya da .env dosyasına gömülmesin; Vault/AWS Secrets Manager
    gibi bir sistemde tutulsun.

-   **Kod tarafında güvenlik:** SAST/DAST taramaları CI/CD pipeline\'ına
    entegre edilsin, bağımlılık (dependency) taraması (Snyk, Dependabot)
    düzenli çalıştırılsın.

-   **Girdi doğrulama:** SQL injection, XSS, CSRF gibi saldırılara karşı
    input validation ve WAF katmanı.

-   **Rate limiting:** API ve login endpoint\'lerinde brute-force\'a
    karşı hız sınırlama.

-   **Sertifika yönetimi:** TLS sertifikalarının otomatik yenilenmesi
    (Let\'s Encrypt/ACM), süresi dolmadan alert.

6\. İzleme, Loglama ve Olay Müdahalesi

-   **Merkezi loglama (SIEM):** Tüm sunucu, ağ, uygulama, veritabanı
    logları merkezi bir sisteme (ELK, Splunk, Datadog) toplanıp
    korelasyon kurallarıyla izlensin.

-   **Anomali tespiti:** Alışılmadık giriş saatleri, coğrafi konum, veri
    dışa aktarım hacmi gibi davranışlar için alert kurulsun.

-   **Dosya bütünlük izleme (FIM):** Kritik sistem dosyalarında yetkisiz
    değişiklik tespiti.

-   **Olay müdahale planı (IR plan):** Bir sızma olduğunda kimin ne
    yapacağı önceden yazılı olsun (izolasyon, adli inceleme, bildirim,
    iletişim). Düzenli tabletop egzersizleri yapın.

-   **Penetrasyon testi ve zafiyet taraması:** Yılda en az 1-2 kez
    bağımsız pentest, sürekli otomatik vulnerability scanning.

7\. Uyumluluk ve Süreç

-   **PCI-DSS uyumluluğu:** Kart verisiyle temas eden her sistem için
    PCI-DSS gereksinimlerini (SAQ veya tam denetim) takip edin.

-   **KVKK/GDPR:** Türkiye pazarı için KVKK uyumu, veri saklama/silme
    politikaları.

-   **Tedarikçi/üçüncü parti risk yönetimi:** Kullandığınız her
    SaaS/servis sağlayıcısının güvenlik sertifikalarını (SOC 2,
    ISO 27001) kontrol edin.

-   **Değişiklik yönetimi:** Prod ortamına her değişiklik
    change-management süreciyle, onaylı ve loglanarak yapılsın.

Bölüm 2 --- Kapsamlı Güvenlik Testi (Full-Force Saldırgan Profili)

Ödeme altyapısına ek olarak, güçlü bir saldırgan profiline karşı test
edilmesi gereken diğer alanlar aşağıda kategori kategori listelenmiştir.

1\. Web Uygulama Katmanı (OWASP Top 10 bazlı)

-   **Injection saldırıları:** SQL injection, NoSQL injection, command
    injection, LDAP injection

-   **Broken authentication:** Zayıf şifre politikaları, session
    fixation, şifre sıfırlama akışındaki zafiyetler, \"remember me\"
    token güvenliği

-   **XSS (Stored/Reflected/DOM-based):** Kullanıcı girdisi alınan her
    alan (yorum, profil, arama kutusu, dosya adı)

-   **IDOR (Insecure Direct Object Reference):** ?user_id=123 gibi
    parametreleri değiştirerek başkasının verisine erişim

-   **CSRF:** Form işlemlerinde token kontrolü var mı

-   **SSRF:** Sunucunun sizin adınıza dış/iç kaynaklara istek atmasını
    sağlayan zafiyetler (özellikle URL/webhook alan yerlerde)

-   **File upload zafiyetleri:** Zararlı dosya yükleme, path traversal,
    dosya tipi doğrulama bypass

-   **XXE (XML External Entity):** XML parse eden endpoint\'ler varsa

-   **Business logic hataları:** Örneğin sepette fiyat manipülasyonu,
    kupon kodu tekrar kullanımı, negatif miktar girme

2\. API Güvenliği

-   Rate limiting olmayan endpoint\'ler (brute force, enumeration)

-   API versiyonlama üzerinden eski/yamasız endpoint\'lere erişim

-   Aşırı veri ifşası (response\'ta gereksiz alanlar dönmesi - mass
    assignment)

-   GraphQL kullanıyorsanız: introspection açık mı, derinlik limiti var
    mı

-   API key/token\'ların URL\'de taşınması (loglara düşer)

3\. Kimlik Doğrulama ve Oturum Yönetimi

-   Şifre politikası (uzunluk, karmaşıklık, breach-check - Have I Been
    Pwned API entegrasyonu)

-   Hesap kilitleme mekanizması (brute force koruması)

-   Session token\'ların güvenliği (HttpOnly, Secure, SameSite cookie
    flag\'leri)

-   JWT kullanıyorsanız: algoritma downgrade saldırısı (alg:none), süre
    kontrolü, imza doğrulama

-   Çoklu cihaz oturum yönetimi, \"tüm cihazlardan çıkış yap\" özelliği

4\. Altyapı ve Konfigürasyon

-   **Subdomain enumeration:** Unutulmuş/terkedilmiş subdomain\'ler
    (subdomain takeover riski)

-   **Açık portlar/servisler:** Shodan/nmap ile dışarıdan görünen tüm
    yüzeyi tarayın

-   **HTTP güvenlik başlıkları:** CSP, HSTS, X-Frame-Options,
    X-Content-Type-Options eksik mi

-   **Hata mesajları:** Stack trace, sürüm bilgisi gibi hassas bilgi
    sızdıran hata sayfaları

-   **Dizin listeleme (directory listing):** açık mı

-   **Yedek/gizli dosyalar:** .git, .env, .bak, wp-config.php.bak gibi
    erişilebilir dosyalar

-   **Admin panel erişimi:** /admin, /wp-admin gibi paneller dışarıya
    açık mı, IP kısıtlaması var mı

5\. Üçüncü Parti ve Tedarik Zinciri

-   Kullanılan JS kütüphanelerinde bilinen zafiyet var mı (npm audit,
    Snyk)

-   Üçüncü parti script/widget\'lar (chat widget, analytics) tehlikeye
    girerse ne olur (supply chain saldırısı)

-   CDN üzerinden yüklenen dosyalarda Subresource Integrity (SRI)
    kullanılıyor mu

6\. Sosyal Mühendislik ve İnsan Faktörü

-   Phishing simülasyonu (çalışanlara test amaçlı sahte e-posta)

-   Destek ekibinin kimlik doğrulama süreçleri (telefonla \"şifremi
    sıfırla\" diyen birine ne kadar kolay yardım ediliyor)

-   Fiziksel güvenlik (ofis erişimi, USB politikası) --- büyük ölçekli
    değilse öncelik değil ama not edilmeli

7\. Mobil Uygulama (varsa)

-   API anahtarlarının APK/IPA içine gömülü olup olmadığı

-   Certificate pinning var mı

-   Root/jailbreak tespiti

-   Local storage\'da hassas veri (token, kart bilgisi) düz metin
    tutuluyor mu

Pratik Öneri

Bunların hepsini kendi başınıza test etmek yerine, OWASP ZAP veya Burp
Suite ile otomatik tarama + manuel pentest kombinasyonu yapmanız, ve
mümkünse bağımsız bir güvenlik firmasından pentest almanız en sağlıklısı
--- özellikle ödeme verisi işleyen bir sistemde bu neredeyse zorunlu
(PCI-DSS de zaten pentest istiyor).
