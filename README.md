# Kablaj Hata Oran Takip 🔌📊

Aylık SRM dökümlerine (`Kablaj Senelik Özetler`: Ocak–Eylül `.xlsx`) göre
**karşılaştırmalı hata oranı analiz panosu**. GitHub Pages uyumlu **tam statik** site —
sunucu yok, veri tarayıcıda işlenir.

Eski `denetim_pano.html` (masaüstü exe çıktısı, gömülü Plotly) mantığı buraya taşındı:
tedarikçi skor kartı, günlük/haftalık seyir yerine **aylık trend + üretim yeri ısı
tablosu + malzeme kırılımı + ay-karşılaştırma** eklendi; grafikler Chart.js ile
dinamik, veriler `data/summary.json`'dan veya sürükle-bırak ile okunur.

## Hızlı başlangıç (yerel önizleme)

```powershell
cd C:\Users\user\Desktop\HataOranTakip
# 1) Aylık dosyaları kopyala
Copy-Item "C:\Users\user\Desktop\Kablaj Senelik Özetler\Ocak.xlsx" data\
#   ... (diğer aylar; adlar manifest.json ile eşleşmeli)
# 2) Özeti üret
python scripts/build_data.py
# 3) Önizle (herhangi biri):
python -m http.server 8000
#   -> http://localhost:8000
```

> `file://` ile açma — `fetch("data/...")` çalışmaz, mutlaka http ile servis et.

## Yayınlama (GitHub Pages)

1. Bu klasörü yeni repoya push'la (`main` dalı).
2. Repo → **Settings → Pages** → Source: **GitHub Actions**.
3. Push'la; `.github/workflows/pages.yml` yayınlar. Adres:
   `https://<kullanıcı>.github.io/<repo>/`
4. `data/*.xlsx` atıp push'layınca `update-data.yml` `summary.json`'u tazeler.

## Veri akışı

```
data/Ocak.xlsx … Eylül.xlsx
   │  python scripts/build_data.py
   ▼
data/summary.json  (+ data/manifest.json)
   │  fetch (öncelikli)  ──▶  site (Chart.js)
   │  yoksa: manifest → *.xlsx → SheetJS (tarayıcıda parse)
   └── sürükle-bırak (commit'siz anlık analiz) + tedarikçi CSV dışa aktar
```

## SRM otomasyonu?

`docs/SRM_ENTEGRASYON.md`'yi oku. Özet: statik site tek başına portala giriş
yapamaz; zamanlanmış çekme için Actions + secret gerekir, şablon hazır
(`srm_cek.yml.example`). Portalının dışa-aktar biçimini paylaşırsan doldururum.

## Kullanıcı girişi (kayıt + rol bazlı)

Site açılışta giriş ister (Supabase Auth, ücretsiz). Roller:
`izleyici` → Genel Bakış + Trend • `yonetici` → + Üretim/Tedarikçi/Malzeme/Karşılaştır •
`admin` → + Veri Yükle. Yeni kayıt `izleyici` başlar.

Kurulum: `docs/KULLANICI_YONETIMI.md` (10 dk: proje aç → `supabase/schema.sql` çalıştır →
`supabase-config.js` doldur → ilk hesabı admin yap). Anahtar girilmeden site **demo
modunda** açılır. ⚠️ Public repoda `data/` dosyaları linkle indirilebilir; gerçek veri
koruması için dokümandaki A/B seçeneklerini oku.

## Proje yapısı

```
index.html  assets/css/style.css  assets/js/{app,auth,supabase-config}.js
data/  (Ocak.xlsx… + manifest.json + summary.json + sure_ozet.json)
scripts/{build_data,build_sure}.py  supabase/schema.sql  docs/{SRM_ENTEGRASYON,KULLANICI_YONETIMI}.md
.github/workflows/{pages,update-data,srm_cek.example}.yml
```

## Süre verimliliği (Denetim Süreleri)

⏱ Süreler sekmesi; 190 bin satırlık süre dosyasını kablaj talepleriyle birleştirir
(ortanca denetim süresi, bekleme ve kapanma süreleri — firma/ay kırılımı).
23 MB'lık ham dosya repoya girmez; özeti yerinden üret:

```powershell
python scripts/build_sure.py --sure "C:\Users\user\Desktop\Kablaj Senelik Özetler\Denetim Süreleri (33).xlsx"
git add data/sure_ozet.json; git commit -m "veri: sure ozeti"; git push
```
