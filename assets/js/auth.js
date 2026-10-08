/* Rol bazlı giriş kapısı (Supabase Auth + profiles tablosu).
   - app.js açılışta AuthKapisi.hazir(baslat) diye bekler; giriş olmadan veri yüklenmez.
   - Supabase yapılandırılmamışsa "demo modu": giriş ekranı atlanır, bant gösterilir.
   - Roller (supabase/schema.sql): izleyici / yonetici / admin. Yeni kayıt → izleyici. */
(function () {
"use strict";

var cfg = window.SUPABASE_CONFIG || {};
var KURULU = !!(cfg.URL && cfg.ANON_KEY && cfg.URL.indexOf("http") === 0);

var ROL_SEKMELER = {
  admin:    ["genel", "trend", "uretim", "tedarikci", "malzeme", "karsilastir", "yukle"],
  yonetici: ["genel", "trend", "uretim", "tedarikci", "malzeme", "karsilastir"],
  izleyici: ["genel", "trend"]
};
var ROL_AD = { admin: "Yönetici", yonetici: "Sorumlu", izleyici: "İzleyici" };

var sb = null, baslatCb = null;

function $(id) { return document.getElementById(id); }
function hata(msj) { var e = $("auth-hata"); if (e) e.textContent = msj || ""; }

function sekmeleriKisitla(izinliler) {
  document.querySelectorAll(".sekme").forEach(function (b) {
    b.style.display = izinliler.indexOf(b.dataset.sekme) >= 0 ? "" : "none";
  });
  var aktif = document.querySelector(".sekme.aktif");
  if (aktif && aktif.style.display === "none") {
    var g = document.querySelector('.sekme[data-sekme="genel"]');
    if (g) g.click();
  }
}

function girisTamam(email, rol) {
  var ekran = $("giris-ekrani");
  if (ekran) ekran.style.display = "none";
  var ka = $("kullanici-alani");
  if (ka) {
    ka.innerHTML = "";
    var roz = document.createElement("span");
    roz.className = "kullanici-rozet";
    roz.textContent = email + " • " + (ROL_AD[rol] || rol);
    var btn = document.createElement("button");
    btn.className = "btn btn-gri btn-kucuk";
    btn.textContent = "Çıkış";
    btn.onclick = cikis;
    ka.appendChild(roz);
    ka.appendChild(btn);
  }
  sekmeleriKisitla(ROL_SEKMELER[rol] || ROL_SEKMELER.izleyici);
  if (baslatCb) { var f = baslatCb; baslatCb = null; f(); }
}

async function rolGetir() {
  var u = (await sb.auth.getUser()).data.user;
  if (!u) return null;
  try {
    var r = await sb.from("profiles").select("rol").eq("id", u.id).single();
    if (r.error) throw r.error;
    return { email: u.email, rol: r.data.rol };
  } catch (e) {
    console.warn("profil okunamadı, izleyici sayıldı:", e);
    return { email: u.email, rol: "izleyici" };
  }
}

async function cikis() {
  if (sb) await sb.auth.signOut();
  location.reload();
}

function bagla() {
  var tg = $("tab-giris"), tk = $("tab-kayit"), bg = $("btn-giris"), bk = $("btn-kayit");
  tg.onclick = function () {
    tg.classList.add("aktif"); tk.classList.remove("aktif");
    bg.style.display = ""; bk.style.display = "none"; hata("");
  };
  tk.onclick = function () {
    tk.classList.add("aktif"); tg.classList.remove("aktif");
    bk.style.display = ""; bg.style.display = "none"; hata("");
  };
  bg.onclick = async function () {
    hata(""); bg.disabled = true;
    try {
      var r = await sb.auth.signInWithPassword({
        email: $("auth-email").value.trim(), password: $("auth-sifre").value });
      if (r.error) throw r.error;
      var p = await rolGetir();
      girisTamam(p.email, p.rol);
    } catch (e) { hata("Giriş başarısız: " + (e.message || e)); }
    bg.disabled = false;
  };
  bk.onclick = async function () {
    hata(""); bk.disabled = true;
    try {
      var email = $("auth-email").value.trim(), sifre = $("auth-sifre").value;
      if (sifre.length < 6) throw new Error("Şifre en az 6 karakter olmalı.");
      var r = await sb.auth.signUp({ email: email, password: sifre });
      if (r.error) throw r.error;
      if (!r.data.session) {
        hata("Kayıt alındı. E-postana gelen doğrulama linkine tıkla, sonra giriş yap.");
      } else {
        var p = await rolGetir();
        girisTamam(p.email, (p && p.rol) || "izleyici");
      }
    } catch (e) { hata("Kayıt başarısız: " + (e.message || e)); }
    bk.disabled = false;
  };
}

function demoModu(neden) {
  var ekran = $("giris-ekrani");
  if (ekran) ekran.style.display = "none";
  var bant = $("demo-bant");
  if (bant) {
    bant.style.display = "block";
    bant.textContent = "⚠️ Demo modu (giriş kapalı" + (neden ? ": " + neden : "") +
      ") — docs/KULLANICI_YONETIMI.md ile Supabase bağla.";
  }
  sekmeleriKisitla(ROL_SEKMELER.admin);
  if (baslatCb) { var f = baslatCb; baslatCb = null; f(); }
}

window.AuthKapisi = {
  beklet: true,
  mod: function () { return KURULU ? "supabase" : "demo"; },
  hazir: function (cb) {
    baslatCb = cb;
    if (!KURULU) { demoModu("yapılandırma yok"); return; }
    if (!window.supabase) { demoModu("kütüphane yüklenemedi"); return; }
    try {
      sb = window.supabase.createClient(cfg.URL, cfg.ANON_KEY);
    } catch (e) { demoModu("istemci hatası"); return; }
    bagla();
    sb.auth.getSession().then(function (s) {
      if (s.data && s.data.session) {
        rolGetir().then(function (p) { if (p) girisTamam(p.email, p.rol); });
      }
      // Oturum yoksa giriş ekranı görünür kalır.
    });
  }
};
})();
