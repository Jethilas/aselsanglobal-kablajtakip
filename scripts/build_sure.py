#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Denetim Süreleri xlsx -> data/sure_ozet.json (süre verimliliği özeti).

Kapsam: SADECE kablaj. Süre dosyasındaki satırlar, data/*.xlsx aylık
dökümlerdeki "Denetim Talep No" ile eşleşenler alınır (ölçülen eşleşme ~%98).

Metrikler (ay × firma):
  - ortanca (medyan) denetim süresi [dk] + p90
  - bekleme süresi: Yaratma -> Denetim Başlangıcı [gün, medyan]
  - kapanma süresi: Denetim Bitişi -> Denetçi işlemi [saat, medyan]

Kullanım (23 MB'lık dosya repoya KONMAZ, yerinde okunur):
    python scripts/build_sure.py --sure "C:/Users/user/Desktop/Kablaj Senelik Özetler/Denetim Süreleri (33).xlsx"

Çıktı: data/sure_ozet.json (küçük, commit'lenir).
"""
import argparse, json, statistics, sys
from collections import Counter, defaultdict
from datetime import date, datetime, time, timedelta
from pathlib import Path

KOK = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(KOK / "scripts"))
import build_data  # norm + ay_etiketi için (main guard'lı, güvenli)

SURE_MAX_DK = 1440      # 24 saat üstü süreler veri hatası sayılır
BEKLEME_MAX_GUN = 60
KAPANMA_MAX_SAAT = 24 * 14


def saat_parse(v):
    if v is None or v == "":
        return time(0, 0)
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    s = str(v).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%H.%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None


def tarih_parse(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def birlestir(tarih_v, saat_v):
    d = tarih_parse(tarih_v)
    if d is None:
        return None
    t = saat_parse(saat_v)
    if t is None:
        return None
    return datetime.combine(d, t)


def aylik_talep_haritasi(data_dir):
    """Talep No -> ay (Ocak..Eylül). İlk bulunan ay kazanır (ay sırasına göre)."""
    harita = {}
    dosyalar = sorted(Path(data_dir).glob("*.xlsx"))
    for yol in dosyalar:
        if yol.name.startswith("~") or yol.name.startswith("Denetim"):
            continue
        ay = build_data.ay_etiketi(yol.name)
        wb = build_data.openpyxl.load_workbook(str(yol), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            if r[1] is not None:
                tid = str(r[1]).strip()
                if tid and tid not in harita:
                    harita[tid] = ay
    return harita


def medyan(lst):
    return round(statistics.median(lst), 1) if lst else None


def p90(lst):
    if not lst:
        return None
    s = sorted(lst)
    return round(s[min(len(s) - 1, int(len(s) * 0.9))], 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sure", default="",
                    help="Denetim Süreleri xlsx yolu (repo dışından okunabilir)")
    ap.add_argument("--data", default=str(KOK / "data"))
    ap.add_argument("--cikti", default=str(KOK / "data" / "sure_ozet.json"))
    args = ap.parse_args()

    sure_yol = Path(args.sure) if args.sure else None
    if sure_yol is None:
        adaylar = list(Path(args.data).glob("Denetim*.xlsx"))
        sure_yol = adaylar[0] if adaylar else None
    if sure_yol is None or not sure_yol.exists():
        sys.exit("Süre dosyası bulunamadı. --sure ile yol ver. "
                 'Örn: python scripts/build_sure.py --sure "C:/.../Denetim Süreleri (33).xlsx"')

    print("1/3 Aylık talep haritası okunuyor…")
    talep_ay = aylik_talep_haritasi(args.data)
    print(f"    {len(talep_ay)} kablaj talebi")

    print("2/3 Süre dosyası taranıyor (190 bin satır, biraz sürer)…")
    wb = build_data.openpyxl.load_workbook(str(sure_yol), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]

    # kolonlar sabit: 0 TalepID, 4 FirmaAd, 8 BasTarih, 9 BasSaat, 10 BitTarih,
    # 11 BitSaat, 12 Sure(dk), 19 YaratTarih, 20 YaratSaat, 22 IslemTarih, 23 IslemSaat
    sureler, beklemeler, kapanmalar = [], [], []
    ay_sure = defaultdict(list)
    ay_bekle = defaultdict(list)
    ay_kapan = defaultdict(list)
    firma_sure = defaultdict(list)
    firma_bekle = defaultdict(list)
    firma_kapan = defaultdict(list)
    firma_n = Counter()
    dagilim = Counter()
    kablaj_satir = 0
    eslesen_talep = set()
    dislanan = Counter()

    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        tid = str(r[0]).strip() if r[0] is not None else ""
        ay = talep_ay.get(tid)
        if not ay:  # kablaj dışı sektör veya eşleşmeyen ID
            continue
        kablaj_satir += 1
        eslesen_talep.add(tid)
        firma = (str(r[4]).strip() if r[4] is not None else "") or "Bilinmiyor"
        firma_n[firma] += 1

        try:
            s = float(str(r[12]).replace(",", "."))
        except (ValueError, TypeError, AttributeError):
            dislanan["sure_hatali"] += 1
            continue
        if not (0 < s <= SURE_MAX_DK):
            dislanan["sure_disi"] += 1
            continue
        sureler.append(s)
        ay_sure[ay].append(s)
        firma_sure[firma].append(s)
        for sinir, etiket in ((15, "0-15 dk"), (30, "15-30 dk"), (60, "30-60 dk"),
                              (120, "1-2 sa"), (240, "2-4 sa"), (10**9, "4 sa+")):
            if s <= sinir:
                dagilim[etiket] += 1
                break

        bas = birlestir(r[8], r[9])
        yarat = birlestir(r[19], r[20])
        bit = birlestir(r[10], r[11])
        islem = birlestir(r[22], r[23])
        if bas and yarat:
            b = (bas - yarat).total_seconds() / 86400
            if 0 <= b <= BEKLEME_MAX_GUN:
                beklemeler.append(b)
                ay_bekle[ay].append(b)
                firma_bekle[firma].append(b)
            else:
                dislanan["bekleme_disi"] += 1
        else:
            dislanan["bekleme_eksik"] += 1
        if bit and islem:
            k = (islem - bit).total_seconds() / 3600
            if 0 <= k <= KAPANMA_MAX_SAAT:
                kapanmalar.append(k)
                ay_kapan[ay].append(k)
                firma_kapan[firma].append(k)
            else:
                dislanan["kapanma_disi"] += 1
        else:
            dislanan["kapanma_eksik"] += 1

        if kablaj_satir % 10000 == 0:
            print(f"    {kablaj_satir} kablaj satırı…")

    print("3/3 Özet yazılıyor…")
    aylar = sorted(set(ay_sure) | set(ay_bekle),
                   key=lambda a: build_data.AYLAR.index(a) if a in build_data.AYLAR else 99)
    aylik = [{"ay": a, "n": len(ay_sure[a]),
              "sure_medyan": medyan(ay_sure[a]), "sure_p90": p90(ay_sure[a]),
              "bekleme_medyan_gun": medyan(ay_bekle[a]),
              "kapanma_medyan_saat": medyan(ay_kapan[a])} for a in aylar]

    firmalar = [{"ad": f, "n": firma_n[f],
                 "sure_medyan": medyan(firma_sure[f]), "sure_p90": p90(firma_sure[f]),
                 "bekleme_medyan_gun": medyan(firma_bekle[f]),
                 "kapanma_medyan_saat": medyan(firma_kapan[f])}
                for f, _ in firma_n.most_common(30)]

    cikti = {
        "meta": {
            "uretilme": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "kablaj_satir": kablaj_satir,
            "kablaj_talep": len(talep_ay),
            "eslesen_talep": len(eslesen_talep),
            "eslesme_oran": round(100 * len(eslesen_talep) / len(talep_ay), 1) if talep_ay else 0,
            "sure_medyan_dk": medyan(sureler),
            "sure_p90_dk": p90(sureler),
            "bekleme_medyan_gun": medyan(beklemeler),
            "kapanma_medyan_saat": medyan(kapanmalar),
            "dislanan": dict(dislanan),
            "kaynak": "build_sure.py (Denetim Süreleri + aylık döküm birleşimi)",
        },
        "aylik": aylik,
        "firmalar": firmalar,
        "dagilim": [{"aralik": k, "adet": dagilim[k]}
                    for k in ["0-15 dk", "15-30 dk", "30-60 dk", "1-2 sa", "2-4 sa", "4 sa+"]],
    }
    out = Path(args.cikti)
    out.write_text(json.dumps(cikti, ensure_ascii=False), encoding="utf-8")
    print(f"OK: {out} ({out.stat().st_size/1024:.0f} KB)")
    print(f"    kablaj satır: {kablaj_satir}, medyan süre: {cikti['meta']['sure_medyan_dk']} dk, "
          f"bekleme: {cikti['meta']['bekleme_medyan_gun']} gün, "
          f"kapanma: {cikti['meta']['kapanma_medyan_saat']} saat")


if __name__ == "__main__":
    main()
