# SRM Portal Otomasyonu — Durum ve Yol Haritası

Kullanıcının isteği: *"SRM portaldan otomatik çekebiliyorsa öyle olsun, ekstradan klasöre atılan Excel'ler de görünsün."*

## Kısa cevap (dürüst)

**GitHub Pages'teki statik site tek başına SRM portalına giriş yapıp veri ÇEKEMEZ.**
Nedenleri:

1. SRM portalı kurumsal ağda / VPN arkasında ve oturum (SSO/şifre) ister. Tarayıcıdan
   `fetch("https://srm...")` çağrısı **CORS + kimlik doğrulama** duvarına takılır.
2. Kullanıcı adı/şifreyi statik siteye gömmek **güvenlik ihlali** olur (herkes görür).
3. GitHub Pages'in zamanlanmış iş çalıştırma yetkisi yoktur (o iş Actions'ındır).

## Önerilen mimari (3 katman — hepsi bu repoda hazır)

```
[SRM portal] --(A: Actions, secret ile)--> data/*.xlsx --(build_data.py)--> data/summary.json --(Pages)--> site
     \--(B: manuel)--> data/ klasörüne dosya atma ----^
      \--(C: anlık)--> sitede sürükle-bırak (commit'siz önizleme)
```

- **B ve C şu an çalışıyor.** A için aşağıdaki şablon var, sadece SRM erişim detayınla doldurulacak.

## A) Otomatik çekme şablonu (yapılacak)

`.github/workflows/srm_cek.yml.example` dosyasını `.github/workflows/srm_cek.yml` olarak
kopyala ve şu secret'ları repo → Settings → Secrets → Actions'a ekle:

| Secret | Açıklama |
|---|---|
| `SRM_URL` | Rapor indirme uç noktası (örn. `https://srm.sirket/.../export`) |
| `SRM_USER` / `SRM_PASS` | Servis hesabı (kişisel hesabını KULLANMA) |
| `SRM_RAPOR_ID` | Kaydedilmiş SRM raporunun kimliği / parametreleri |

Akış: ayda bir (cron) veya manuel tetikleme → SRM'den xlsx indir → `data/` altına
`<Ay>.xlsx` olarak kaydet → `build_data.py` çalıştır → değiştiyse otomatik commit+push →
Pages yeniden yayınlanır.

> Not: SRM'in dışa aktarma biçimi (OData? CSV? SAP GUI script?) siteden siteye değişir.
> Senin portalının "dışa aktar / zamanlanmış rapor e-postası" özelliği varsa en pratik
> yol çoğu zaman **e-posta ile gelen eki Actions'ın alması** veya **şirket içi bir
> zamanlanmış görevin dosyayı bu repoya push'lamasıdır**. Portal ekran görüntüsünü /
> dışa aktar menüsünü paylaşırsan şablonu gerçek uçlara göre doldururum.

## B) Klasöre atma (şu an aktif)

`data/*.xlsx` + `python scripts/build_data.py` + push. `update-data.yml` workflow'u
push sonrası `summary.json`'u tazeler (veya yerelde üretip commit'lersin).

## C) Sürükle-bırak (şu an aktif)

Site → **Veri Yükle** sekmesi. Dosyalar yalnızca tarayıcı belleğinde işlenir,
hiçbir yere yüklenmez (kurumsal veri için güvenli).
