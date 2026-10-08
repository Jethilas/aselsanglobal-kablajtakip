# Kullanıcı Yönetimi (Kayıt + Rol Bazlı Görünüm)

Site, **Supabase Auth** ile gerçek üyelik kullanır (ücretsiz katman yeter).
Kayıt olan herkes `izleyici` başlar; rolleri bir `admin` yükseltir.

## Rol → ekran haritası

| Rol | Görebildikleri |
|---|---|
| `izleyici` (varsayılan) | 📊 Genel Bakış, 📈 Aylık Trend |
| `yonetici` | + 🏭 Üretim Yeri, 🏷️ Tedarikçiler, 🧰 Malzeme, ⚖️ Ay Karşılaştır |
| `admin` | + 📥 Veri Yükle |

## Kurulum (10 dk, bir kez)

1. **supabase.com → New project** (ücretsiz). Project URL + `anon public` key'i al.
   (`service_role` anahtarını kimseyle paylaşma, siteye koyma.)
2. **SQL Editor → New query** → `supabase/schema.sql` içeriğini yapıştır → Run.
3. **Authentication → Sign In / Up → Email provider** açık olsun.
   Kurumsal kullanımda **"Confirm email"** ayarını ekibe göre seç:
   - Kapalı = kayıt olan direkt girer (pratik, dahili ekip için).
   - Açık = e-posta doğrulaması ister (daha güvenli).
4. `assets/js/supabase-config.js` dosyasını açıp URL + anon key'i yaz.
   Bu dosya **commit'lenir** (anon key'in herkese görünmesi normaldir; Supabase'in
   modeli böyledir, korumayı RLS politikaları yapar). `service_role` anahtarını
   ASLA bu dosyaya yazma.
5. Siteyi aç → **Kayıt Ol** ile ilk hesabı oluştur → SQL'de yükselt:
   ```sql
   update public.profiles set rol = 'admin' where email = 'sen@ornek.com';
   ```
6. Diğer kayıtları **Table Editor → profiles** tablosundan `rol` değiştirerek yetkilendir.
   (Veya SQL ile toplu: `update ... set rol='yonetici' where email in (...)`.)

## Yerel test

Anahtarlar boşken site **demo modunda** açılır (giriş ekranı atlanır, üstte uyarı bandı).
Gerçek akışı test etmek için 4. adımdaki dosyayı doldurup `python -m http.server` ile aç.

## ⚠️ Önemli güvenlik notu (oku!)

Sekme gizleme **arayüz seviyesindedir**: herkese açık (public) bir repoda
`data/summary.json` ve `data/*.xlsx` dosyaları linki bilen herkesçe indirilebilir.
Gerçek veri koruması için **ikisini birlikte** yap:

- **A — Kabul edilebilir:** Repoyu **Private** yap. (Not: ücretsiz planda private
  repoda Pages yayını kapalıdır; ya ücretli plan ya da B seçeneği gerekir.)
- **B — Önerilen:** Excel/JSON'ları repodan çıkarıp **Supabase Storage**'a koy
  (RLS ile "sadece giriş yapan okur"), site veriyi oradan çeksin. İstersen bu adıma
  geçiririm — `app.js`'teki `baslat()` veri kaynağı tek fonksiyonda toplu, değişiklik küçük.
- Şifre kuralı: Supabase → Auth → Policies'te min. şifre uzunluğunu 8+ yap.

Eski `denetim_pano.html`'deki `KULL=[["Yonetici","7ad9…"]]` tipi gömülü-hash girişi
**bilerek kullanılmadı**: HTML'e gömülü her sır, sayfa kaynağını açan herkesçe görülür.
