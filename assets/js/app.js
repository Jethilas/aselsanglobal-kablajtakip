/* Kablaj Hata Oran Takip — statik GitHub Pages uygulaması
   Veri kaynakları (öncelik sırası):
   1) data/summary.json  (build_data.py çıktısı — hızlı yol)
   2) data/manifest.json + data/*.xlsx (tarayıcıda SheetJS ile parse)
   3) Sürükle-bırak ile yüklenen xlsx (geçici, commit gerekmez)
*/
"use strict";

const AY_SIRASI = ["Ocak","Şubat","Mart","Nisan","Mayıs","Mayyıs","Haziran","Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"];
const ayRank = a => { const i = AY_SIRASI.indexOf(a); return i < 0 ? 99 : i; };
const ayAd = s => ({"Mayyıs":"Mayıs"}[s] || s);

const DURUM = document.getElementById("veri-durum");
const charts = {};
function grafik(id, cfg) {
  if (charts[id]) charts[id].destroy();
  const el = document.getElementById(id);
  if (!el) return;
  charts[id] = new Chart(el, cfg);
}
function fmt(n){ return (n==null||isNaN(n)) ? "–" : Number(n).toLocaleString("tr-TR"); }
function pct(a,b){ return b>0 ? (a/b*100) : 0; }
function rozet(o){ // hata oranına göre renk
  if (o >= 10) return "r-kirmizi";
  if (o >= 6) return "r-turuncu";
  if (o > 0) return "r-mavi";
  return "r-yesil";
}
function isiRenk(o, max){
  if (!max) max = 15;
  const t = Math.min(1, o / max);
  const r = Math.round(220*t + 22*(1-t)), g = Math.round(38*t + 163*(1-t)), b = Math.round(38*t + 74*(1-t));
  return `rgb(${r},${g},${b})`;
}
function setDurum(msj, cls){ DURUM.textContent = msj; DURUM.className = "veri-durum " + (cls||""); }

/* ---------- durum ---------- */
let OZET = null;          // summary.json veya hesaplanan özet
let HAM = [];             // satır bazında ham kayıtlar (filtreler için)
let MANIFEST = [];
let SURE = null;          // sure_ozet.json (Denetim Süreleri birleşimi, opsiyonel)

/* ---------- kolon eşleştirme (Türkçe toleranslı) ---------- */
function norm(s){
  return String(s||"").toLocaleLowerCase("tr")
    .replace(/ğ/g,"g").replace(/ü/g,"u").replace(/ş/g,"s")
    .replace(/ı/g,"i").replace(/ö/g,"o").replace(/ç/g,"c").replace(/i̇/g,"i")
    .replace(/[^a-z0-9 ]/g," ").replace(/\s+/g," ").trim();
}
function kolonIndex(basliklar){
  const n = basliklar.map(norm);
  const bul = (...aday) => {
    for (const a of aday){ const i = n.findIndex(h => h.includes(norm(a))); if (i>=0) return i; }
    return -1;
  };
  return {
    uretim:  bul("uretim yeri"),
    satici:  bul("satici adi"),
    malzNo:  bul("malzeme no"),
    malzAd:  bul("malzeme tanimi"),
    sonuc:   bul("denetim sonucu"),
    kontrol: bul("kontrol turu"),
    tarih:   bul("yaratilm"),
    denetimTuru: bul("denetim turu"),
    talep:   bul("denetim talep"),
  };
}
function sonucSinif(v){
  const t = norm(v);
  if (!t) return "bos";
  if (t.includes("degil") || t.includes("red") || t.includes("ret")) return "red";
  if (t.includes("uygun")) return "uygun";
  return "diger";
}

/* ---------- SheetJS ile xlsx parse ---------- */
async function xlsxOku(arrayBuffer, ayEtiket){
  const wb = XLSX.read(arrayBuffer, {type:"array"});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
  if (rows.length < 2) return [];
  const baslik = rows[0].map(String);
  const K = kolonIndex(baslik);
  const out = [];
  for (let i=1;i<rows.length;i++){
    const r = rows[i];
    const sonuc = K.sonuc>=0 ? String(r[K.sonuc]||"").trim() : "";
    out.push({
      ay: ayEtiket,
      talep: K.talep>=0 ? String(r[K.talep]||"").trim() : "",
      uretim: K.uretim>=0 ? String(r[K.uretim]||"").trim() || "Bilinmiyor" : "Bilinmiyor",
      satici: K.satici>=0 ? String(r[K.satici]||"").trim() || "Bilinmiyor" : "Bilinmiyor",
      malzNo: K.malzNo>=0 ? String(r[K.malzNo]||"").trim() : "",
      malzAd: K.malzAd>=0 ? String(r[K.malzAd]||"").trim() : "",
      kontrol: K.kontrol>=0 ? String(r[K.kontrol]||"").trim() || "—" : "—",
      tarih: K.tarih>=0 ? String(r[K.tarih]||"").trim() : "",
      sonuc: sonuc,
      sinif: sonucSinif(sonuc),
    });
  }
  return out;
}

/* ---------- özet hesapla ---------- */
function ozetHesapla(kayitlar){
  const aylar = [...new Set(kayitlar.map(k=>k.ay))].sort((a,b)=>ayRank(a)-ayRank(b));
  const aylik = aylar.map(ay => {
    const rs = kayitlar.filter(k=>k.ay===ay);
    const uygun = rs.filter(k=>k.sinif==="uygun").length;
    const red = rs.filter(k=>k.sinif==="red").length;
    return {ay, toplam:rs.length, uygun, red, bos:rs.length-uygun-red, hata_oran:+pct(red,rs.length).toFixed(2)};
  });
  const uretimYerleri = [...new Set(kayitlar.map(k=>k.uretim))].sort();
  const isi = {}; // uretim -> ay -> {toplam, red, oran}
  uretimYerleri.forEach(u => { isi[u] = {}; aylar.forEach(a => {
    const rs = kayitlar.filter(k=>k.uretim===u && k.ay===a);
    const red = rs.filter(k=>k.sinif==="red").length;
    isi[u][a] = {toplam:rs.length, red, oran:+pct(red,rs.length).toFixed(2)};
  });});
  const uretimTop = uretimYerleri.map(u => {
    const rs = kayitlar.filter(k=>k.uretim===u);
    const red = rs.filter(k=>k.sinif==="red").length;
    return {yer:u, toplam:rs.length, red, oran:+pct(red,rs.length).toFixed(2)};
  }).sort((a,b)=>b.toplam-a.toplam);
  const sMap = {};
  kayitlar.forEach(k => {
    const ad = k.satici || "Bilinmiyor";
    sMap[ad] = sMap[ad] || {ad, toplam:0, uygun:0, red:0, aylar:{}, uretim:{}};
    const s = sMap[ad]; s.toplam++;
    if (k.sinif==="uygun") s.uygun++; else if (k.sinif==="red") s.red++;
    s.aylar[k.ay] = (s.aylar[k.ay]||0)+1;
    s.uretim[k.uretim] = (s.uretim[k.uretim]||0)+1;
  });
  const saticilar = Object.values(sMap).map(s => ({...s, oran:+pct(s.red,s.toplam).toFixed(2)}))
    .sort((a,b)=>b.toplam-a.toplam);
  const kMap = {};
  kayitlar.forEach(k => { kMap[k.kontrol] = kMap[k.kontrol] || {ad:k.kontrol, toplam:0, red:0}; kMap[k.kontrol].toplam++; if(k.sinif==="red") kMap[k.kontrol].red++; });
  const kontroller = Object.values(kMap).sort((a,b)=>b.toplam-a.toplam);
  const mMap = {};
  kayitlar.forEach(k => {
    const key = (k.malzNo||"") + " | " + (k.malzAd||"").slice(0,60);
    if (!k.malzNo && !k.malzAd) return;
    mMap[key] = mMap[key] || {malz:key, toplam:0, red:0};
    mMap[key].toplam++; if(k.sinif==="red") mMap[key].red++;
  });
  const malzemeler = Object.values(mMap).sort((a,b)=>b.red-a.red || b.toplam-a.toplam);
  const toplam = kayitlar.length;
  const red = kayitlar.filter(k=>k.sinif==="red").length;
  const uygun = kayitlar.filter(k=>k.sinif==="uygun").length;
  return {aylar, aylik, isi, uretimYerleri, uretimTop, saticilar, kontroller, malzemeler,
    genel:{toplam, uygun, red, oran:+pct(red,toplam).toFixed(2), saticiSayi:saticilar.length}};
}

/* ---------- filtreli kayıtlar ---------- */
function filtreli(){
  const secAy = [...document.getElementById("filtre-ay").selectedOptions].map(o=>o.value);
  const uretim = document.getElementById("filtre-uretim").value;
  const sat = document.getElementById("filtre-satici").value.toLocaleLowerCase("tr");
  const snc = document.getElementById("filtre-sonuc").value;
  return HAM.filter(k =>
    (!secAy.length || secAy.includes(k.ay)) &&
    (!uretim || k.uretim===uretim) &&
    (!sat || (k.satici||"").toLocaleLowerCase("tr").includes(sat)) &&
    (!snc || k.sonuc.includes(snc))
  );
}

/* ---------- render ---------- */
function renderTumu(){
  const kayitlar = filtreli();
  const o = ozetHesapla(kayitlar.length ? kayitlar : HAM);
  renderKPI(o);
  renderTrend(o);
  renderUretim(o);
  renderSatici(o);
  renderMalzeme(o);
  renderKarsilastirSecenek(o);
  document.getElementById("uretilme-bilgi").textContent =
    (OZET && OZET.meta ? "Özet: "+OZET.meta.uretilme+" • " : "") + fmt(kayitlar.length)+" / "+fmt(HAM.length)+" kayıt (filtreli/toplam)";
}

function renderKPI(o){
  const g = o.genel;
  const aylik = o.aylik;
  const son = aylik[aylik.length-1], once = aylik[aylik.length-2];
  const fark = (son && once) ? (son.hata_oran-once.hata_oran) : null;
  document.getElementById("kpi-kartlar").innerHTML = `
    <div class="kpi kpi-tik" data-kpi="tum" title="Tıkla: tüm kayıtları listele"><div class="ad">Toplam Denetim</div><div class="deger">${fmt(g.toplam)}</div><div class="delta">${o.aylar.length} ay • ${g.saticiSayi} satıcı</div></div>
    <div class="kpi yesil kpi-tik" data-kpi="uygun" title="Tıkla: uygun kayıtları listele"><div class="ad">✅ Uygun</div><div class="deger">${fmt(g.uygun)}</div><div class="delta">%${pct(g.uygun,g.toplam).toFixed(1)}</div></div>
    <div class="kpi kirmizi kpi-tik" data-kpi="red" title="Tıkla: red kayıtlarını listele"><div class="ad">❌ Uygun Değil</div><div class="deger">${fmt(g.red)}</div><div class="delta">%${g.oran}</div></div>
    <div class="kpi turuncu"><div class="ad">⚠️ Hata Oranı</div><div class="deger">%${g.oran}</div>
      <div class="delta ${fark==null?"":fark<=0?"iyi":"kotu"}">${fark==null?"":(fark>0?"+":"")+fark.toFixed(2)+" puan (önceki aya göre)"}</div></div>
    <div class="kpi"><div class="ad">En riskli üretim yeri</div><div class="deger" style="font-size:18px">${o.uretimTop.slice().sort((a,b)=>b.oran-a.oran)[0]?.yer||"–"}</div>
      <div class="delta">%${o.uretimTop.slice().sort((a,b)=>b.oran-a.oran)[0]?.oran ?? "–"}</div></div>`;
  // mini grafikler
  grafik("gr-trend-mini", {type:"line",
    data:{labels:aylik.map(a=>ayAd(a.ay)), datasets:[{label:"Hata oranı %", data:aylik.map(a=>a.hata_oran), borderColor:"#dc2626", backgroundColor:"rgba(220,38,38,.12)", fill:true, tension:.3}]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}}});
  grafik("gr-uretim-mini", {type:"bar",
    data:{labels:o.uretimTop.map(u=>u.yer), datasets:[{label:"Hata oranı %", data:o.uretimTop.map(u=>u.oran), backgroundColor:o.uretimTop.map(u=>isiRenk(u.oran,12))}]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}}});
  grafik("gr-sonuc", {type:"doughnut",
    data:{labels:["Uygun","Uygun Değil","Diğer/Boş"], datasets:[{data:[g.uygun,g.red,g.toplam-g.uygun-g.red],
      backgroundColor:["#16a34a","#dc2626","#94a3b8"]}]},
    options:{responsive:true, maintainAspectRatio:false}});
  grafik("gr-uretim-pay", {type:"doughnut",
    data:{labels:o.uretimTop.map(u=>u.yer), datasets:[{data:o.uretimTop.map(u=>u.toplam),
      backgroundColor:["#1d4ed8","#16a34a","#ea580c","#9333ea","#64748b","#0d9488"]}]},
    options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:"right"}}}});
  document.getElementById("tbl-aylar").innerHTML =
    `<tr><th>Ay</th><th>Toplam</th><th>Uygun</th><th>Uygun Değil</th><th>Hata Oranı</th><th>Önceki Aya Fark</th></tr>` +
    aylik.map((a,i)=>{ const f = i>0 ? a.hata_oran-aylik[i-1].hata_oran : null;
      return `<tr class="satir" data-ay="${a.ay}" title="Tıkla: ${ayAd(a.ay)} kayıtlarını listele"><td class="sol"><b>${ayAd(a.ay)}</b></td><td>${fmt(a.toplam)}</td><td>${fmt(a.uygun)}</td>
      <td>${fmt(a.red)}</td><td><span class="rozet ${rozet(a.hata_oran)}">%${a.hata_oran}</span></td>
      <td>${f==null?"—":`<b style="color:${f<=0?"#16a34a":"#dc2626"}">${f>0?"+":""}${f.toFixed(2)} puan</b>`}</td></tr>`;}).join("");
}

let trendTip = "bar";
function renderTrend(o){
  const aylik = o.aylik;
  grafik("gr-trend-adet", {type:trendTip,
    data:{labels:aylik.map(a=>ayAd(a.ay)),
      datasets:[{label:"Toplam", data:aylik.map(a=>a.toplam), backgroundColor:"#93c5fd", borderColor:"#1d4ed8"},
                {label:"Uygun Değil", data:aylik.map(a=>a.red), backgroundColor:"#dc2626", borderColor:"#dc2626"}]},
    options:{responsive:true, maintainAspectRatio:false}});
  const topYer = o.uretimTop.slice(0,4).map(u=>u.yer);
  const renk = ["#dc2626","#1d4ed8","#16a34a","#ea580c"];
  grafik("gr-trend-oran", {type:"line",
    data:{labels:aylik.map(a=>ayAd(a.ay)),
      datasets:topYer.map((u,i)=>({label:u,
        data:aylik.map(a=>o.isi[u]?.[a.ay]?.oran ?? 0),
        borderColor:renk[i%renk.length], tension:.3}))},
    options:{responsive:true, maintainAspectRatio:false}});
  document.getElementById("tbl-degisim").innerHTML =
    `<tr><th>Ay</th><th>Hata Oranı</th><th>Fark (puan)</th><th>Yorum</th></tr>` +
    aylik.map((a,i)=>{ const f = i>0 ? a.hata_oran-aylik[i-1].hata_oran : null;
      const yorum = f==null ? "—" : f<=-1 ? "🟢 Belirgin iyileşme" : f<0 ? "🟢 Hafif iyileşme" : f===0 ? "⚪ Stabil" : f<1 ? "🟡 Hafif artış" : "🔴 Belirgin artış";
      return `<tr><td class="sol"><b>${ayAd(a.ay)}</b></td><td>%${a.hata_oran}</td><td>${f==null?"—":(f>0?"+":"")+f.toFixed(2)}</td><td class="sol">${yorum}</td></tr>`;}).join("");
}

function renderUretim(o){
  // Sektör vitrini: her üretim yerinin uygunsuzluk talep oranı kartı (tıklanabilir)
  const topRed = o.saticilar.reduce((t, s) => t + s.red, 0);
  const sira = o.uretimTop.slice().sort((a, b) => b.oran - a.oran);
  document.getElementById("sektor-kartlar").innerHTML = sira.map(u =>
    `<div class="kpi kpi-tik ${u.oran >= 10 ? "kirmizi" : u.oran >= 6 ? "turuncu" : "yesil"}" data-sektor="${u.yer}" title="Tıkla: ${u.yer} kayıtlarını listele">
      <div class="ad">🏭 ${u.yer}</div><div class="deger">%${u.oran}</div>
      <div class="delta">${fmt(u.red)} red / ${fmt(u.toplam)} talep • redlerin %${pct(u.red, topRed).toFixed(1)}'i</div>
    </div>`).join("");
  grafik("gr-uretim-adet", {type:"bar",
    data:{labels:o.uretimTop.map(u=>u.yer),
      datasets:[{label:"Toplam", data:o.uretimTop.map(u=>u.toplam), backgroundColor:"#1d4ed8"},
                {label:"Red", data:o.uretimTop.map(u=>u.red), backgroundColor:"#dc2626"}]},
    options:{responsive:true, maintainAspectRatio:false, indexAxis:"y"}});
  const max = Math.max(5, ...o.uretimTop.map(u=>u.oran));
  document.getElementById("tbl-isi").innerHTML =
    `<tr><th>Üretim Yeri</th>` + o.aylar.map(a=>`<th>${ayAd(a)}</th>`).join("") + `<th>Toplam</th></tr>` +
    o.uretimTop.map(u=>`<tr><td class="sol"><b>${u.yer}</b></td>` +
      o.aylar.map(a=>{ const h=o.isi[u.yer][a]; return `<td data-u="${u.yer}" data-a="${a}" title="Tıkla: ${u.yer} / ${ayAd(a)} kayıtları"><span class="isi" style="background:${isiRenk(h.oran,max)}" title="${fmt(h.red)}/${fmt(h.toplam)}">%${h.oran}</span></td>`;}).join("") +
      `<td><span class="rozet ${rozet(u.oran)}">%${u.oran}</span><br><small>${fmt(u.red)}/${fmt(u.toplam)}</small></td></tr>`).join("");
}

let saticiSort = {key:"toplam", dir:-1};
function renderSatici(o){
  const top = o.saticilar.slice(0,15);
  grafik("gr-satici", {type:"bar",
    data:{labels:top.map(s=>s.ad.slice(0,34)), datasets:[{label:"Hata oranı %", data:top.map(s=>s.oran), backgroundColor:top.map(s=>isiRenk(s.oran,12))}]},
    options:{responsive:true, maintainAspectRatio:false, indexAxis:"y", plugins:{legend:{display:false}}}});
  // Pareto: en çok red üreten 12 satıcı + kümülatif pay
  const pr = o.saticilar.filter(s=>s.red>0).sort((a,b)=>b.red-a.red).slice(0,12);
  const totRed = o.saticilar.reduce((t,s)=>t+s.red,0) || 1;
  let kum = 0;
  grafik("gr-pareto", {type:"bar",
    data:{labels:pr.map(s=>s.ad.slice(0,28)), datasets:[
      {label:"Red adedi", data:pr.map(s=>s.red), backgroundColor:"#dc2626", xAxisID:"x"},
      {label:"Kümülatif %", data:pr.map(s=>{kum+=s.red; return +(kum/totRed*100).toFixed(1);}),
       type:"line", borderColor:"#1d4ed8", backgroundColor:"#1d4ed8", xAxisID:"x1"}]},
    options:{responsive:true, maintainAspectRatio:false, indexAxis:"y",
      scales:{x:{title:{display:true,text:"red adedi"}}, x1:{position:"top", min:0, max:100, title:{display:true,text:"kümülatif %"}, grid:{drawOnChartArea:false}}}}});
  const q = (document.getElementById("satici-ara").value||"").toLocaleLowerCase("tr");
  let rows = o.saticilar.filter(s=>!q || s.ad.toLocaleLowerCase("tr").includes(q));
  rows = rows.sort((a,b)=> (a[saticiSort.key]>b[saticiSort.key]?1:-1)*saticiSort.dir).slice(0,200);
  const th = (et,k)=>`<th data-k="${k}">${et} ${saticiSort.key===k?(saticiSort.dir===1?"▲":"▼"):""}</th>`;
  document.getElementById("tbl-satici").innerHTML =
    `<tr>${th("Satıcı","ad")}${th("Toplam","toplam")}${th("Uygun","uygun")}${th("Red","red")}${th("Hata %","oran")}<th>Yoğun Ay</th></tr>` +
    rows.map((s,i)=>{ const yAy = Object.entries(s.aylar).sort((a,b)=>b[1]-a[1])[0];
      return `<tr class="satir" data-i="${o.saticilar.indexOf(s)}"><td class="sol">${s.ad}</td><td>${fmt(s.toplam)}</td>
      <td>${fmt(s.uygun)}</td><td>${fmt(s.red)}</td>
      <td><span class="rozet ${rozet(s.oran)}">%${s.oran}</span></td>
      <td>${yAy?ayAd(yAy[0])+" ("+yAy[1]+")":"—"}</td></tr>`;}).join("") ||
    `<tr><td colspan="6">Kayıt yok</td></tr>`;
  document.querySelectorAll("#tbl-satici th[data-k]").forEach(el=>el.onclick=()=>{
    const k=el.dataset.k;
    if(saticiSort.key===k) saticiSort.dir*=-1; else saticiSort={key:k,dir:k==="ad"?1:-1};
    renderSatici(o);
  });
  document.querySelectorAll("#tbl-satici tr.satir").forEach(tr=>tr.onclick=()=>{
    const s = o.saticilar[+tr.dataset.i];
    const uret = Object.entries(s.uretim).sort((a,b)=>b[1]-a[1]).map(([k,v])=>`${k}: ${v}`).join(" • ");
    document.getElementById("satici-detay").innerHTML =
      `<div class="kpi"><div class="ad">${s.ad}</div>
       <div>Toplam <b>${fmt(s.toplam)}</b> • Uygun <b>${fmt(s.uygun)}</b> • Red <b>${fmt(s.red)}</b> • Hata <b>%${s.oran}</b></div>
       <div class="delta">${uret}</div>
       <div style="margin-top:8px"><button class="btn btn-mavi btn-kucuk" id="satici-kayitlar">🔍 ${fmt(s.toplam)} kaydı listele</button></div></div>`;
    document.getElementById("satici-detay").scrollIntoView({behavior:"smooth", block:"nearest"});
    document.getElementById("satici-kayitlar").onclick = () =>
      detayGoster(s.ad + " — kayıtlar", HAM.filter(k => k.satici === s.ad));
  });
}

function renderMalzeme(o){
  grafik("gr-kontrol", {type:"doughnut",
    data:{labels:o.kontroller.map(k=>k.ad), datasets:[{data:o.kontroller.map(k=>k.toplam),
      backgroundColor:["#1d4ed8","#16a34a","#ea580c","#9333ea","#64748b"]}]},
    options:{responsive:true, maintainAspectRatio:false}});
  const top = o.malzemeler.slice(0,15);
  grafik("gr-malzeme", {type:"bar",
    data:{labels:top.map(m=>m.malz.slice(0,40)), datasets:[{label:"Red adedi", data:top.map(m=>m.red), backgroundColor:"#dc2626"}]},
    options:{responsive:true, maintainAspectRatio:false, indexAxis:"y", plugins:{legend:{display:false}}}});
  document.getElementById("tbl-malzeme").innerHTML =
    `<tr><th>Malzeme</th><th>Toplam</th><th>Red</th><th>Hata %</th></tr>` +
    o.malzemeler.slice(0,100).map(m=>`<tr><td class="sol">${m.malz}</td><td>${m.toplam}</td><td>${m.red}</td>
      <td><span class="rozet ${rozet(pct(m.red,m.toplam))}">%${pct(m.red,m.toplam).toFixed(1)}</span></td></tr>`).join("");
}

function renderKarsilastirSecenek(o){
  for (const id of ["cmp-a","cmp-b"]){
    const sel = document.getElementById(id);
    const cur = sel.value;
    sel.innerHTML = o.aylar.map(a=>`<option value="${a}">${ayAd(a)}</option>`).join("");
    if ([...sel.options].some(x=>x.value===cur)) sel.value = cur;
  }
  if (o.aylar.length>=2 && !document.getElementById("cmp-a").value){
    document.getElementById("cmp-a").value = o.aylar[o.aylar.length-2];
    document.getElementById("cmp-b").value = o.aylar[o.aylar.length-1];
  }
}
function karsilastir(){
  const a = document.getElementById("cmp-a").value, b = document.getElementById("cmp-b").value;
  if (!a || !b) return;
  const ka = HAM.filter(k=>k.ay===a), kb = HAM.filter(k=>k.ay===b);
  const oa = ozetHesapla(ka.length?ka:[{ay:a}]), ob = ozetHesapla(kb.length?kb:[{ay:b}]);
  const ga = oa.genel, gb = ob.genel;
  const kart = (ad,va,vb,ters)=>{ const f = (vb-va);
    const iyi = ters ? f<0 : f>0; // hata için ters=true
    return `<div class="kpi"><div class="ad">${ad}</div><div class="deger" style="font-size:19px">${va} → ${vb}</div>
      <div class="delta ${f===0?"":iyi?"iyi":"kotu"}">${f===0?"değişim yok":(f>0?"+":"")+ (typeof f==="number"?f.toFixed(2):f)}</div></div>`; };
  document.getElementById("cmp-kartlar").innerHTML =
    kart("Toplam denetim", ga.toplam, gb.toplam) +
    kart("Red adedi", ga.red, gb.red, true) +
    kart("Hata oranı %", ga.oran, gb.oran, true);
  const yerler = [...new Set([...oa.uretimTop.map(u=>u.yer), ...ob.uretimTop.map(u=>u.yer)])];
  grafik("gr-cmp", {type:"bar",
    data:{labels:yerler, datasets:[
      {label:ayAd(a)+" %", data:yerler.map(y=>oa.uretimTop.find(u=>u.yer===y)?.oran||0), backgroundColor:"#93c5fd"},
      {label:ayAd(b)+" %", data:yerler.map(y=>ob.uretimTop.find(u=>u.yer===y)?.oran||0), backgroundColor:"#dc2626"}]},
    options:{responsive:true, maintainAspectRatio:false}});
  document.getElementById("tbl-cmp").innerHTML =
    `<tr><th>Üretim Yeri</th><th>${ayAd(a)} (red/toplam • %)</th><th>${ayAd(b)} (red/toplam • %)</th><th>Fark (puan)</th></tr>` +
    yerler.map(y=>{ const xa=oa.uretimTop.find(u=>u.yer===y), xb=ob.uretimTop.find(u=>u.yer===y);
      const pa=xa?.oran||0, pb=xb?.oran||0, f=pb-pa;
      return `<tr><td class="sol"><b>${y}</b></td>
        <td>${xa?`${fmt(xa.red)}/${fmt(xa.toplam)} • %${pa}`:"—"}</td>
        <td>${xb?`${fmt(xb.red)}/${fmt(xb.toplam)} • %${pb}`:"—"}</td>
        <td><b style="color:${f<=0?"#16a34a":"#dc2626"}">${f>0?"+":""}${f.toFixed(2)}</b></td></tr>`;}).join("");
}

/* ---------- süre verimliliği (Denetim Süreleri birleşimi) ---------- */
async function sureBaslat() {
  try {
    const r = await fetch("data/sure_ozet.json", { cache: "no-store" });
    if (!r.ok) throw new Error("yok");
    SURE = await r.json();
  } catch (e) { SURE = null; }
  renderSure();
}
function renderSure() {
  const kpi = document.getElementById("sure-kpi");
  if (!SURE) {
    kpi.innerHTML = `<div class="kpi turuncu"><div class="ad">Süre verisi yok</div>
      <div class="delta">Henüz üretilmedi — <code>python scripts/build_sure.py --sure "&lt;Denetim Süreleri.xlsx yolu&gt;"</code> çalıştır.</div></div>`;
    return;
  }
  const m = SURE.meta;
  kpi.innerHTML = `
    <div class="kpi"><div class="ad">⏱ Ortanca Denetim Süresi</div><div class="deger">${m.sure_medyan_dk} dk</div><div class="delta">${fmt(m.kablaj_satir)} kablaj satırı</div></div>
    <div class="kpi"><div class="ad">p90 Süre</div><div class="deger">${m.sure_p90_dk} dk</div><div class="delta">%90'ı bunun altında</div></div>
    <div class="kpi yesil"><div class="ad">⏳ Ortanca Bekleme</div><div class="deger">${m.bekleme_medyan_gun} gün</div><div class="delta">Yaratma → denetim başı</div></div>
    <div class="kpi turuncu"><div class="ad">🔒 Ortanca Kapanma</div><div class="deger">${m.kapanma_medyan_saat} sa</div><div class="delta">Denetim sonu → kayıt kapanışı</div></div>`;
  const A = SURE.aylik;
  grafik("gr-sure-trend", { type: "line",
    data: { labels: A.map(a => ayAd(a.ay)), datasets: [
      { label: "Medyan (dk)", data: A.map(a => a.sure_medyan), borderColor: "#1d4ed8", tension: .3 },
      { label: "p90 (dk)", data: A.map(a => a.sure_p90), borderColor: "#dc2626", borderDash: [6, 4], tension: .3 }]},
    options: { responsive: true, maintainAspectRatio: false }});
  grafik("gr-sure-dagilim", { type: "bar",
    data: { labels: SURE.dagilim.map(d => d.aralik), datasets: [{ label: "Adet", data: SURE.dagilim.map(d => d.adet), backgroundColor: "#1d4ed8" }]},
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }}}});
  grafik("gr-sure-bekle", { type: "bar",
    data: { labels: A.map(a => ayAd(a.ay)), datasets: [
      { label: "Bekleme (gün)", data: A.map(a => a.bekleme_medyan_gun), backgroundColor: "#93c5fd", yAxisID: "y" },
      { label: "Kapanma (saat)", data: A.map(a => a.kapanma_medyan_saat), type: "line", borderColor: "#ea580c", tension: .3, yAxisID: "y1" }]},
    options: { responsive: true, maintainAspectRatio: false,
      scales: { y: { title: { display: true, text: "gün" }}, y1: { position: "right", title: { display: true, text: "saat" }, grid: { drawOnChartArea: false }}}}});
  const top = SURE.firmalar.slice(0, 15);
  grafik("gr-sure-firma", { type: "bar",
    data: { labels: top.map(f => f.ad.slice(0, 34)), datasets: [{ label: "Medyan (dk)", data: top.map(f => f.sure_medyan), backgroundColor: top.map(f => isiRenk(f.sure_medyan, 30)) }]},
    options: { responsive: true, maintainAspectRatio: false, indexAxis: "y", plugins: { legend: { display: false }}}});
  document.getElementById("tbl-sure").innerHTML =
    `<tr><th>Firma</th><th>Denetim</th><th>Medyan (dk)</th><th>p90 (dk)</th><th>Bekleme (gün)</th><th>Kapanma (sa)</th></tr>` +
    SURE.firmalar.map(f => `<tr><td class="sol">${f.ad}</td><td>${fmt(f.n)}</td>
      <td><span class="rozet ${f.sure_medyan <= 20 ? "r-yesil" : f.sure_medyan <= 40 ? "r-mavi" : "r-turuncu"}">${f.sure_medyan}</span></td>
      <td>${f.sure_p90}</td><td>${f.bekleme_medyan_gun ?? "–"}</td><td>${f.kapanma_medyan_saat ?? "–"}</td></tr>`).join("");
  const d = m.dislanan || {};
  document.getElementById("sure-not").innerHTML =
    `Kapsam: aylık kablaj dökümlerindeki ${fmt(m.kablaj_talep)} talebin %${m.eslesme_oran}'i süre dosyasıyla eşleşti ` +
    `(${fmt(m.eslesen_talep)} talep, ${fmt(m.kablaj_satir)} denetim satırı, ${m.uretilme}). ` +
    `Aralık dışı bırakılanlar: süre ${fmt(d.sure_disi || 0)}, bekleme ${fmt(d.bekleme_disi || 0)}, kapanma ${fmt(d.kapanma_disi || 0)}. ` +
    `Not: firmaların farklı yazımları (örn. CDK / CDK SAV.) ayrı satırdadır — isim standardizasyonu SRM'de düzeltilmeli.`;
}

/* ---------- veri yükleme ---------- */
async function baslat(){
  sekmeleriBagla();
  try {
    const r = await fetch("data/summary.json", {cache:"no-store"});
    if (r.ok){
      OZET = await r.json();
      HAM = OZET.satirlar || [];
      // summary.json satır içermiyorsa (hafif mod) manifest+xlsx'e düş
      if (!HAM.length) throw new Error("satir yok");
      setDurum(`✅ ${OZET.meta.aylar.length} ay • ${fmt(HAM.length)} kayıt (summary.json)`, "hazir");
      filtreleriDoldur(); renderTumu(); return;
    }
    throw new Error("summary yok");
  } catch(e){ /* fallback: manifest + xlsx */ }
  try {
    const r = await fetch("data/manifest.json", {cache:"no-store"});
    if (!r.ok) throw new Error("manifest yok");
    const m = await r.json();
    MANIFEST = m.dosyalar || [];
    if (!MANIFEST.length) throw new Error("liste boş");
    setDurum(`⏳ ${MANIFEST.length} dosya okunuyor…`, "");
    HAM = [];
    for (const d of MANIFEST){
      try {
        const f = await fetch("data/"+d.dosya, {cache:"no-store"});
        if (!f.ok) continue;
        const buf = await f.arrayBuffer();
        const rows = await xlsxOku(buf, d.ay);
        HAM = HAM.concat(rows);
        setDurum(`⏳ ${fmt(HAM.length)} kayıt… (${d.ay})`, "");
      } catch(err){ console.warn(d.dosya, err); }
    }
    if (!HAM.length) throw new Error("hiç satır okunamadı");
    setDurum(`✅ ${MANIFEST.length} dosya • ${fmt(HAM.length)} kayıt`, "hazir");
    filtreleriDoldur(); renderTumu();
  } catch(e2){
    setDurum("⚠️ Veri bulunamadı — Veri Yükle sekmesinden Excel bırak", "hata");
    sekmeyeGit("yukle");
  }
}

function filtreleriDoldur(){
  const aylar = [...new Set(HAM.map(k=>k.ay))].sort((a,b)=>ayRank(a)-ayRank(b));
  const fa = document.getElementById("filtre-ay");
  fa.innerHTML = aylar.map(a=>`<option value="${a}">${ayAd(a)}</option>`).join("");
  fa.size = Math.min(6, Math.max(1, aylar.length));
  const yerler = [...new Set(HAM.map(k=>k.uretim))].sort();
  const fu = document.getElementById("filtre-uretim");
  fu.innerHTML = `<option value="">Tümü</option>` + yerler.map(u=>`<option>${u}</option>`).join("");
  renderKarsilastirSecenek(ozetHesapla(HAM));
}

/* ---------- detay modalı (tıkla → kayıtları listele) ---------- */
function detayGoster(baslik, kayitlar) {
  const goster = kayitlar.slice(0, 500);
  document.getElementById("detay-baslik").textContent = baslik;
  document.getElementById("detay-bilgi").textContent =
    `${fmt(kayitlar.length)} kayıt bulundu` + (kayitlar.length > 500 ? " — ilk 500 gösteriliyor" : "") + ".";
  document.getElementById("detay-tablo").innerHTML =
    `<tr><th>Ay</th><th>Talep No</th><th>Üretim Yeri</th><th>Satıcı</th><th>Malzeme</th><th>Sonuç</th></tr>` +
    (goster.map(k => `<tr><td>${k.ay ? ayAd(k.ay) : "–"}</td><td>${k.talep || "–"}</td><td>${k.uretim}</td>` +
      `<td class="sol">${(k.satici || "").slice(0, 45)}</td>` +
      `<td class="sol">${((k.malzNo || "") + " " + (k.malzAd || "")).slice(0, 45)}</td>` +
      `<td>${k.sinif === "red" ? "❌ " : k.sinif === "uygun" ? "✅ " : ""}${k.sonuc || "–"}</td></tr>`).join("") ||
     `<tr><td colspan="6">Kayıt yok</td></tr>`);
  document.getElementById("detay-modal").classList.add("acik");
}
function detayKapat() { document.getElementById("detay-modal").classList.remove("acik"); }
function mevcutKayitlar() { const f = filtreli(); return f.length ? f : HAM; }

function sekmeleriBagla(){
  document.querySelectorAll(".sekme").forEach(b=>b.onclick=()=>{
    document.querySelectorAll(".sekme").forEach(x=>x.classList.remove("aktif"));
    b.classList.add("aktif");
    document.querySelectorAll(".panel").forEach(p=>p.classList.remove("aktif"));
    document.getElementById("sekme-"+b.dataset.sekme).classList.add("aktif");
  });
  ["filtre-ay","filtre-uretim","filtre-sonuc"].forEach(id=>document.getElementById(id).onchange=renderTumu);
  document.getElementById("filtre-satici").oninput=()=>{clearTimeout(window.__t); window.__t=setTimeout(renderTumu,300);};
  document.getElementById("satici-ara").oninput=()=>renderTumu();
  document.getElementById("btn-sifirla").onclick=()=>{
    [...document.getElementById("filtre-ay").options].forEach(o=>o.selected=false);
    document.getElementById("filtre-uretim").value="";
    document.getElementById("filtre-satici").value="";
    document.getElementById("filtre-sonuc").value="";
    document.getElementById("satici-ara").value="";
    renderTumu();
  };
  document.getElementById("btn-cmp").onclick=karsilastir;
  document.getElementById("btn-csv").onclick=csvIndir;
  // detay tıklamaları (delegasyon — tablolar yeniden çizilse de çalışır)
  document.getElementById("kpi-kartlar").onclick = e => {
    const k = e.target.closest("[data-kpi]"); if (!k) return;
    const f = mevcutKayitlar();
    const map = { tum: [f, "Tüm kayıtlar"], uygun: [f.filter(x => x.sinif === "uygun"), "Uygun kayıtlar"], red: [f.filter(x => x.sinif === "red"), "Uygun değil (red) kayıtlar"] };
    detayGoster(map[k.dataset.kpi][1], map[k.dataset.kpi][0]);
  };
  document.getElementById("tbl-aylar").onclick = e => {
    const tr = e.target.closest("tr[data-ay]"); if (!tr) return;
    detayGoster(ayAd(tr.dataset.ay) + " — kayıtlar", mevcutKayitlar().filter(k => k.ay === tr.dataset.ay));
  };
  document.getElementById("tbl-isi").onclick = e => {
    const td = e.target.closest("td[data-u]"); if (!td) return;
    detayGoster(`${td.dataset.u} / ${ayAd(td.dataset.a)} — kayıtlar`,
      mevcutKayitlar().filter(k => k.uretim === td.dataset.u && k.ay === td.dataset.a));
  };
  document.getElementById("sektor-kartlar").onclick = e => {
    const k = e.target.closest("[data-sektor]"); if (!k) return;
    detayGoster(k.dataset.sektor + " sektörü — kayıtlar", mevcutKayitlar().filter(x => x.uretim === k.dataset.sektor));
  };
  // trend grafik tipi düğmeleri
  const tipGuncelle = () => {
    document.getElementById("tip-bar").classList.toggle("aktif-tip", trendTip === "bar");
    document.getElementById("tip-cizgi").classList.toggle("aktif-tip", trendTip === "line");
  };
  document.getElementById("tip-bar").onclick = () => { trendTip = "bar"; tipGuncelle(); renderTrend(ozetHesapla(mevcutKayitlar())); };
  document.getElementById("tip-cizgi").onclick = () => { trendTip = "line"; tipGuncelle(); renderTrend(ozetHesapla(mevcutKayitlar())); };
  tipGuncelle();
  // detay modal kapatma
  document.getElementById("detay-kapat").onclick = detayKapat;
  document.getElementById("detay-modal").onclick = e => { if (e.target.id === "detay-modal") detayKapat(); };
  document.addEventListener("keydown", e => { if (e.key === "Escape") detayKapat(); });
  // dropzone
  const dz = document.getElementById("dropzone");
  const inp = document.getElementById("dosya-input");
  ["dragover","dragenter"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add("surukle");}));
  ["dragleave","drop"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove("surukle");}));
  dz.addEventListener("drop", e=>dosyalariAl(e.dataTransfer.files));
  inp.addEventListener("change", ()=>dosyalariAl(inp.files));
  document.getElementById("btn-demo-temizle").onclick=()=>baslat();
}
function sekmeyeGit(ad){
  document.querySelector(`.sekme[data-sekme="${ad}"]`)?.click();
}
async function dosyalariAl(files){
  if (!files.length) return;
  setDurum(`⏳ ${files.length} dosya okunuyor…`, "");
  const liste = document.getElementById("yuklenen-liste");
  for (const f of files){
    try {
      const buf = await f.arrayBuffer();
      const ay = f.name.replace(/\.xlsx?$/i,"");
      const rows = await xlsxOku(buf, ay);
      HAM = HAM.concat(rows);
      liste.innerHTML += `<div>✅ ${f.name} — ${fmt(rows.length)} satır (${ay})</div>`;
    } catch(err){ liste.innerHTML += `<div>❌ ${f.name} okunamadı</div>`; }
  }
  setDurum(`✅ Toplam ${fmt(HAM.length)} kayıt (önbellek + yüklenenler)`, "hazir");
  filtreleriDoldur(); renderTumu(); sekmeyeGit("genel");
}
function csvIndir(){
  const o = ozetHesapla(filtreli().length?filtreli():HAM);
  const sat = ["Satici;Toplam;Uygun;Red;HataOran"]
    .concat(o.saticilar.map(s=>`"${s.ad.replace(/"/g,'""')}";${s.toplam};${s.uygun};${s.red};${String(s.oran).replace(".",",")}`));
  const blob = new Blob(["\ufeff"+sat.join("\n")], {type:"text/csv;charset=utf-8"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = "tedarikci_skor_kart.csv"; a.click();
}

document.addEventListener("DOMContentLoaded", () => {
  // Giriş kapısı varsa (auth.js) veri yükleme girişten sonra başlar.
  const boot = () => { baslat(); sureBaslat(); };
  if (window.AuthKapisi && window.AuthKapisi.beklet) window.AuthKapisi.hazir(boot);
  else boot();
});
