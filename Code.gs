// ============================================================
//  KURUMSAL TEDARİK VE ZİMMET YÖNETİM SİSTEMİ — v4.0
//  GitHub Pages + Google Apps Script + Google Sheets
//  Özel Login (E-posta/Şifre) & RBAC
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: "1HI5BxkFKK23jH_yZJxxZD9XJ8hjdxmd5LY65Ca9Bnns",
  CRITICAL_AMOUNT: 100000,
  APP_URL: "", // Deploy sonrası otomatik alınır
  SHEETS: {
    PERSONEL:      "REF_Personel",
    STOK:          "REF_Stok",
    KATEGORI_ONAY: "REF_Kategori_Onay",
    TALEPLER:      "TRX_Talepler",
    SATINALMA:     "TRX_Satinalma",
    ZIMMET:        "TRX_Zimmet",
  },
};

// ─── UTILITY ────────────────────────────────────────────────
function getSpreadsheet() { return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function getSheet(name) { return getSpreadsheet().getSheetByName(name); }
function now() { return Utilities.formatDate(new Date(), "Europe/Istanbul", "dd.MM.yyyy HH:mm"); }
function uuid() { return Utilities.getUuid().replace(/-/g, "").substring(0, 20); }

function genId(prefix, sheet) {
  var n = sheet.getLastRow() < 2 ? 1 : sheet.getLastRow();
  return prefix + "-" + ("00000" + n).slice(-5);
}

function findRow(sheet, col, val) {
  var d = sheet.getRange(1, col, sheet.getLastRow(), 1).getValues();
  for (var i = 1; i < d.length; i++) { if (String(d[i][0]) === String(val)) return i + 1; }
  return null;
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAppUrl() {
  try { return ScriptApp.getService().getUrl(); }
  catch(e) { return CONFIG.APP_URL; }
}

// ══════════════════════════════════════════════════════════════
//  doGet — TÜM OKUMA İŞLEMLERİ + E-POSTA ONAY LİNKLERİ
// ══════════════════════════════════════════════════════════════
function doGet(e) {
  var action = (e.parameter.action || "").toLowerCase();
  var p = e.parameter;

  try {
    switch (action) {
      case "login":
        return json(handleLogin(p.email, p.sifre));

      case "getdashboard":
        return json(getDashboardData(p.email));

      case "getstok":
        return json({ success: true, rows: getAllStok() });

      case "getpersonel":
        return json({ success: true, rows: getAllPersonel() });

      case "approve":
        return handleEmailAction(p, "approve");

      case "reject":
        return handleEmailAction(p, "reject");

      default:
        return json({ success: false, error: "Tanımsız action: " + action });
    }
  } catch (err) {
    return json({ success: false, error: err.message });
  }
}

// ══════════════════════════════════════════════════════════════
//  doPost — TÜM YAZMA İŞLEMLERİ
// ══════════════════════════════════════════════════════════════
function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ success: false, error: "Geçersiz JSON: " + err.message });
  }

  var action = (data.action || "").toLowerCase();

  try {
    switch (action) {
      case "createrequest":
        return json(createRequest(data));

      case "approverequest":
        return json(processApproval(data.talepId, "onayla", data.email));

      case "rejectrequest":
        return json(processApproval(data.talepId, "reddet", data.email));

      case "createsatinalma":
        return json(createSatinalma(data));

      case "createzimmet":
        return json(createZimmet(data));

      case "returnzimmet":
        return json(processIade(data.zimmetId, data.email));

      case "markdelivered":
        return json(markDelivered(data.talepId, data.email));

      case "userteslimal":
        return json(userTeslimAl(data));

      default:
        return json({ success: false, error: "Tanımsız action: " + action });
    }
  } catch (err) {
    return json({ success: false, error: err.message });
  }
}

// ══════════════════════════════════════════════════════════════
//  LOGIN — E-posta + Şifre Doğrulama
// ══════════════════════════════════════════════════════════════
function handleLogin(email, sifre) {
  if (!email || !sifre) return { success: false, error: "E-posta ve şifre gereklidir." };

  var sheet = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!sheet) return { success: false, error: "Personel tablosu bulunamadı." };

  var data = sheet.getDataRange().getValues();
  // Sütunlar: [0]ID [1]Ad_Soyad [2]Email [3]Sifre [4]Departman [5]Yonetici_Email [6]Rol
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
      if (String(data[i][3]).trim() === String(sifre).trim()) {
        return {
          success: true,
          user: {
            id:            data[i][0],
            isim:          data[i][1],
            email:         data[i][2],
            departman:     data[i][4],
            yoneticiEmail: data[i][5],
            rol:           data[i][6],
          }
        };
      } else {
        return { success: false, error: "Şifre hatalı." };
      }
    }
  }
  return { success: false, error: "Bu e-posta kayıtlı değil." };
}

// ══════════════════════════════════════════════════════════════
//  DASHBOARD VERİSİ — ROL BAZLI FİLTRELEME (switch-case)
// ══════════════════════════════════════════════════════════════
function getDashboardData(email) {
  if (!email) return { success: false, error: "E-posta gerekli." };

  var userInfo = getPersonelByEmail(email);
  if (!userInfo) return { success: false, error: "Kullanıcı bulunamadı." };

  var result = {
    success: true,
    user: userInfo,
    taleplerim: [],
    onayBekleyenler: [],
    gecmisIslemler: [],
    stok: getAllStok(), // Herkesin talep ekranında ürün seçebilmesi için stok herkese açık
    satinalma: [],
    zimmetlerim: [],
    tumTalepler: [],
    stats: {},
  };

  // ── Tüm talepleri yükle
  var allTalepler = loadAllTalepler();

  // ── Personel haritası (isim eşleme)
  var persMap = buildPersonelMap();

  // 1. Direktör İsmini Tablodan Bul
  var directorName = "Direktör"; 
  for (var emailKey in persMap) {
    if (persMap[emailKey].rol === "Director") { directorName = persMap[emailKey].isim; break; }
  }

  // 2. Teknik Onaycı Haritası (Kategorilere göre isimler)
  var techApproverMap = {};
  var katSheet = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  if (katSheet && katSheet.getLastRow() >= 2) {
    var katData = katSheet.getRange(2, 1, katSheet.getLastRow() - 1, 3).getValues();
    katData.forEach(function(row) {
      var cat = String(row[0]).toLowerCase().trim();
      var mail = String(row[2]).toLowerCase().trim();
      var p = persMap[mail];
      techApproverMap[cat] = p ? p.isim : row[1];
    });
  }

  // 3. Stok Kategorileri (Ürün IDsine göre kategori bulmak için)
  var stokSheet = getSheet(CONFIG.SHEETS.STOK);
  var itemCats = {};
  if (stokSheet && stokSheet.getLastRow() >= 2) {
    stokSheet.getRange(2, 1, stokSheet.getLastRow() - 1, 3).getValues().forEach(function(r) { 
      itemCats[String(r[0])] = String(r[2]).toLowerCase().trim(); 
    });
  }

  // İsim bilgisi ekle ve BEKLEYEN İSMİ Belirle
  allTalepler.forEach(function(t) {
    var p = persMap[String(t.talepEdenEmail).toLowerCase().trim()];
    t.talepEdenIsim = p ? p.isim : t.talepEdenEmail;
    t.departman = p ? p.departman : "";
    t.yoneticiEmail = p ? p.yoneticiEmail : "";

    // Dinamik Bekleyen İsim Belirleme (İsim + Unvan/Departman eklendi)
    t.bekleyenIsim = "";
    if (t.durum === "Müdür Onayı Bekliyor") {
      var m = persMap[String(t.yoneticiEmail).toLowerCase().trim()];
      t.bekleyenIsim = m ? m.isim + " (" + m.departman + " Müdürü)" : "Birim Yöneticisi";
    } else if (t.durum === "Direktör Onayı Bekliyor") {
      t.bekleyenIsim = directorName + " (Direktör)";
    } else if (t.durum === "Teknik Onay Bekliyor") {
      var cat = itemCats[t.urunId] || "IT";
      t.bekleyenIsim = (techApproverMap[cat] || "Teknik B.") + " (" + cat.toUpperCase() + " Sorumlusu)";
    } else if (t.durum === "Satınalma Aşamasında" || t.durum === "Sipariş Edilebilir") {
      for (var e in persMap) { if (persMap[e].rol === "Procurement") { t.bekleyenIsim = persMap[e].isim + " (Satınalma Birimi)"; break; } }
    } else if (t.durum === "Sipariş Edildi") {
      t.bekleyenIsim = "Lojistik/Haberleşme (Teslimat Bekleniyor)";
    }
  });

  // ═══ ROL BAZLI switch-case ═══
  switch (userInfo.rol) {

    case "User":
      result.taleplerim = allTalepler.filter(function(t) {
        return t.talepEdenEmail.toLowerCase() === email.toLowerCase();
      });
      result.zimmetlerim = getZimmetByEmail(email);
      break;

    case "Manager":
      result.taleplerim = allTalepler.filter(function(t) {
        return t.talepEdenEmail.toLowerCase() === email.toLowerCase();
      });
      // ── KRİTİK: Sadece Yonetici_Email eşleşen + "Müdür Onayı Bekliyor" olanlar
      result.onayBekleyenler = allTalepler.filter(function(t) {
        return t.yoneticiEmail.toLowerCase() === email.toLowerCase() &&
               t.durum === "Müdür Onayı Bekliyor";
      });
      result.gecmisIslemler = allTalepler.filter(function(t) {
        return t.yoneticiEmail.toLowerCase() === email.toLowerCase() &&
               t.durum !== "Müdür Onayı Bekliyor";
      });
      result.zimmetlerim = getZimmetByEmail(email);
      break;

    case "IT":
      result.taleplerim = allTalepler.filter(function(t) {
        return t.talepEdenEmail.toLowerCase() === email.toLowerCase();
      });
      // ── Sadece kendi kategorisindeki "Teknik Onay Bekliyor" talepler
      result.onayBekleyenler = allTalepler.filter(function(t) {
        return t.durum === "Teknik Onay Bekliyor" && isTeknikOnayci(email, t.urunId);
      });
      result.gecmisIslemler = allTalepler.filter(function(t) {
        return t.durum !== "Teknik Onay Bekliyor" &&
               t.sonIslemYapan && t.sonIslemYapan.toLowerCase().includes(email.toLowerCase());
      });
      result.stok = getAllStok();
      result.zimmetlerim = getZimmetByEmail(email);
      break;

    case "Procurement":
      result.taleplerim = allTalepler.filter(function(t) {
        return t.talepEdenEmail.toLowerCase() === email.toLowerCase();
      });
      // ── Satınalma hem fiyat girmesi gerekenleri hem de sipariş geçmesi gerekenleri görür
      result.onayBekleyenler = allTalepler.filter(function(t) {
        return ["Satınalma Aşamasında", "Sipariş Edilebilir", "Sipariş Edildi"].indexOf(t.durum) >= 0;
      });
      result.gecmisIslemler = allTalepler.filter(function(t) {
        return ["Teslim Alındı","Teslim Edildi","Zimmetlendi"].indexOf(t.durum) >= 0;
      });
      result.satinalma = getAllSatinalma();
      result.stok = getAllStok();
      result.zimmetlerim = getZimmetByEmail(email);
      break;

    case "Director":
      result.taleplerim = allTalepler.filter(function(t) {
        return t.talepEdenEmail.toLowerCase() === email.toLowerCase();
      });
      // ── Yüksek bütçeli onay bekleyenler
      result.onayBekleyenler = allTalepler.filter(function(t) {
        return t.durum === "Direktör Onayı Bekliyor";
      });
      result.tumTalepler = allTalepler;
      result.gecmisIslemler = allTalepler.filter(function(t) {
        return ["Reddedildi","Sipariş Edilebilir","Sipariş Edildi","Teslim Alındı","Teslim Edildi","Zimmetlendi"].indexOf(t.durum) >= 0;
      });
      result.satinalma = getAllSatinalma();
      result.stok = getAllStok();
      result.zimmetlerim = getAllZimmet();
      break;
  }

  // ── İstatistikler
  var base = userInfo.rol === "Director" ? allTalepler : result.taleplerim;
  result.stats = {
    toplam:       base.length,
    onayBekleyen: result.onayBekleyenler.length,
    satinalma:    allTalepler.filter(function(t){return t.durum === "Satınalma Aşamasında"}).length,
    siparis:      allTalepler.filter(function(t){return t.durum && t.durum.indexOf("Sipariş") >= 0}).length,
    teslim:       allTalepler.filter(function(t){return t.durum && (t.durum.indexOf("Teslim") >= 0 || t.durum === "Zimmetlendi")}).length,
    red:          allTalepler.filter(function(t){return t.durum === "Reddedildi"}).length,
    kritikStok:   (result.stok || []).filter(function(s){return Number(s.mevcutStok) <= Number(s.kritikSeviye)}).length,
  };

  return result;
}

// ══════════════════════════════════════════════════════════════
//  YENİ TALEP OLUŞTURMA
// ══════════════════════════════════════════════════════════════
function createRequest(data) {
  var user = getPersonelByEmail(data.email);
  if (!user) throw new Error("Yetkisiz.");

  var sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  var talepId = genId("TLP", sheet);
  var route = checkStockRoute(data.urunId, data.miktar || 1);
  var mKey = uuid(), tKey = uuid(), dKey = uuid();

  sheet.appendRow([
    talepId, now(), data.email,
    data.urunId || "", data.urunAdi || "", data.miktar || 1,
    route.ilkDurum, route.surecTipi, data.teknikNot || "",
    mKey, tKey, dKey,
    data.email, now()
  ]);

  // Müdüre e-posta
  if (user.yoneticiEmail) {
    sendApprovalEmail(user.yoneticiEmail, {
      talepId: talepId, talepEden: user.isim, departman: user.departman,
      urunAdi: data.urunAdi || data.urunId || "—", miktar: data.miktar || 1,
      surecTipi: route.surecTipi, teknikNot: data.teknikNot || "",
    }, "mudur", mKey);
  }

  return { success: true, talepId: talepId, surecTipi: route.surecTipi, mesaj: route.mesaj };
}

// ══════════════════════════════════════════════════════════════
//  STOK KONTROL VE YÖNLENDİRME
// ══════════════════════════════════════════════════════════════
function checkStockRoute(urunId, miktar) {
  if (!urunId) return { surecTipi:"Satınalma", ilkDurum:"Müdür Onayı Bekliyor", mesaj:"Manuel ürün → Satınalma sürecine yönlendirildi." };
  var s = getSheet(CONFIG.SHEETS.STOK);
  if (!s) return { surecTipi:"Satınalma", ilkDurum:"Müdür Onayı Bekliyor", mesaj:"Stok tablosu yok." };
  var d = s.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(urunId)) {
      var mevcut = Number(d[i][3]) || 0;
      if (mevcut >= miktar) return { surecTipi:"Stok", ilkDurum:"Müdür Onayı Bekliyor", mesaj:"Stokta mevcut (" + mevcut + "). Müdür onayına gönderildi." };
      return { surecTipi:"Satınalma", ilkDurum:"Müdür Onayı Bekliyor", mesaj:"Stok yetersiz (" + mevcut + "/" + miktar + "). Satınalma sürecine yönlendirildi." };
    }
  }
  return { surecTipi:"Satınalma", ilkDurum:"Müdür Onayı Bekliyor", mesaj:"Ürün stokta tanımlı değil." };
}

// ══════════════════════════════════════════════════════════════
//  ONAY İŞLEMİ (Dashboard butonlarından)
// ══════════════════════════════════════════════════════════════
function processApproval(talepId, aksiyonTipi, email) {
  var user = getPersonelByEmail(email);
  if (!user) throw new Error("Yetkisiz.");

  var sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  var rowIdx = findRow(sheet, 1, talepId);
  if (!rowIdx) throw new Error("Talep bulunamadı: " + talepId);

  var row = sheet.getRange(rowIdx, 1, 1, 14).getValues()[0];
  var durum = row[6], surecTipi = row[7], talepEmail = row[2];
  var authorized = false, newDurum = "";

  switch (durum) {
    case "Müdür Onayı Bekliyor":
      if (user.rol === "Manager") {
        var talep = getPersonelByEmail(talepEmail);
        if (talep && talep.yoneticiEmail.toLowerCase() === email.toLowerCase()) authorized = true;
      }
      if (user.rol === "Director") authorized = true;
      newDurum = aksiyonTipi === "onayla" ? "Teknik Onay Bekliyor" : "Reddedildi";
      break;

    case "Teknik Onay Bekliyor":
      if (user.rol === "IT" && isTeknikOnayci(email, row[3])) authorized = true;
      if (user.rol === "Director") authorized = true;
      if (aksiyonTipi === "onayla") {
        newDurum = surecTipi === "Stok" ? "Teslim Edilebilir" : "Satınalma Aşamasında";
      } else { newDurum = "Reddedildi"; }
      break;

    case "Direktör Onayı Bekliyor":
      if (user.rol === "Director") authorized = true;
      newDurum = aksiyonTipi === "onayla" ? "Sipariş Edilebilir" : "Reddedildi";
      break;

    case "Satınalma Aşamasında":
      if (user.rol === "Procurement" || user.rol === "Director") authorized = true;
      newDurum = aksiyonTipi === "onayla" ? "Sipariş Edilebilir" : "Reddedildi";
      break;

    case "Sipariş Edilebilir":
      if (user.rol === "Procurement" || user.rol === "Director") authorized = true;
      newDurum = "Sipariş Edildi";
      break;

    case "Sipariş Edildi":
      if (user.rol === "Procurement" || user.rol === "IT" || user.rol === "Director") authorized = true;
      newDurum = "Teslim Alındı";
      break;

    case "Teslim Alındı":
    case "Teslim Edilebilir":
      if (user.rol === "IT" || user.rol === "Procurement" || user.rol === "Director") authorized = true;
      newDurum = "Zimmetlendi";
      break;

    default:
      throw new Error("Bu talep üzerinde işlem yapılamaz. Durum: " + durum);
  }

  if (!authorized) throw new Error("Bu işlem için yetkiniz yok. (" + user.rol + ")");

  sheet.getRange(rowIdx, 7).setValue(newDurum);
  sheet.getRange(rowIdx, 13).setValue(email);
  sheet.getRange(rowIdx, 14).setValue(now());

  // Sonraki aşamaya e-posta
  if (aksiyonTipi === "onayla" && newDurum === "Teknik Onay Bekliyor") {
    sendTeknikEmail(sheet, rowIdx);
  }

  // Stok düş
  if (newDurum === "Zimmetlendi" && row[3]) {
    updateStok(row[3], -(Number(row[5]) || 1));
  }

  return { success: true, talepId: talepId, yeniDurum: newDurum };
}

// ══════════════════════════════════════════════════════════════
//  TESLİMAT
// ══════════════════════════════════════════════════════════════
function markDelivered(talepId, email) {
  return processApproval(talepId, "onayla", email);
}

// ══════════════════════════════════════════════════════════════
//  SATINALMA KAYDI
// ══════════════════════════════════════════════════════════════
function createSatinalma(data) {
  var user = getPersonelByEmail(data.email);
  if (!user || (user.rol !== "Procurement" && user.rol !== "Director"))
    throw new Error("Satınalma yetkiniz yok.");

  var sheet = getSheet(CONFIG.SHEETS.SATINALMA);
  var satId = genId("SAT", sheet);
  var miktar = Number(data.miktar) || 1;
  var toplam = (Number(data.birimFiyat) || 0) * miktar;
  var kritik = toplam >= CONFIG.CRITICAL_AMOUNT;

  sheet.appendRow([
    satId, data.talepId || "", data.firma || "",
    data.birimFiyat || 0, data.doviz || "TRY", toplam,
    data.vade || "", data.odemeTipi || "",
    "Beklemede", data.butceKodu || "",
    now(), data.email
  ]);

  // Talep durumunu güncelle
  var ts = getSheet(CONFIG.SHEETS.TALEPLER);
  var tr = findRow(ts, 1, data.talepId);
  if (tr) {
    var nextDurum = kritik ? "Direktör Onayı Bekliyor" : "Sipariş Edilebilir";
    ts.getRange(tr, 7).setValue(nextDurum);
    ts.getRange(tr, 13).setValue(data.email);
    ts.getRange(tr, 14).setValue(now());

    if (kritik) {
      sendDirektorEmail(ts, tr, data, toplam);
    }
  }

  return { success: true, satinalmaId: satId, kritik: kritik, mesaj: kritik ? "KRİTİK ALIM — Direktör onayına gönderildi." : "Satınalma kaydı oluşturuldu." };
}

// ══════════════════════════════════════════════════════════════
//  ZİMMET KAYDI
// ══════════════════════════════════════════════════════════════
function createZimmet(data) {
  var user = getPersonelByEmail(data.email);
  if (!user || (user.rol !== "IT" && user.rol !== "Procurement" && user.rol !== "Director"))
    throw new Error("Zimmet yetkiniz yok.");

  var sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  var zId = genId("ZMT", sheet);
  sheet.appendRow([
    zId, data.personelEmail || "", data.urunId || "",
    data.urunAdi || "", data.seriNo || "", now(), data.email, "Aktif"
  ]);

  if (data.urunId) updateStok(data.urunId, -(Number(data.miktar) || 1));

  // Talep durumunu güncelle
  if (data.talepId) {
    var ts = getSheet(CONFIG.SHEETS.TALEPLER);
    var tr = findRow(ts, 1, data.talepId);
    if (tr) {
      ts.getRange(tr, 7).setValue("Zimmetlendi");
      ts.getRange(tr, 13).setValue(data.email);
      ts.getRange(tr, 14).setValue(now());
    }
  }

  return { success: true, zimmetId: zId };
}

// ══════════════════════════════════════════════════════════════
//  KULLANICININ ZİMMETİ/TALEP EDİLENİ TESLİM ALMASI (DİJİTAL İMZA)
// ══════════════════════════════════════════════════════════════
function userTeslimAl(data) {
  var sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  var tr = findRow(sheet, 1, data.talepId);
  if (!tr) throw new Error("Talep bulunamadı.");

  var row = sheet.getRange(tr, 1, 1, 14).getValues()[0];
  if (row[2].toLowerCase() !== data.email.toLowerCase()) throw new Error("Sadece talebi oluşturan kişi form imzalayarak teslim alabilir.");

  if (row[6] !== "Teslim Alındı" && row[6] !== "Teslim Edilebilir") {
    throw new Error("Bu talep henüz teslim alma aşamasına (stoka/kuruma giriş yapılmış haline) gelmemiş.");
  }

  // 1. TALEBİN DURUMUNU ZİMMETLENDİ YAP
  sheet.getRange(tr, 7).setValue("Zimmetlendi");
  sheet.getRange(tr, 13).setValue(data.email + " (Form İmzalı)");
  sheet.getRange(tr, 14).setValue(now());

  var not = sheet.getRange(tr, 9).getValue();
  sheet.getRange(tr, 9).setValue(not + (not?"\n":"") + "[Kullanıcı Zimmet ve Teslim Formunu İmzaladı]");

  // 2. OTOMATİK ZİMMET TABLOSUNA KAYIT
  // ["ZimmetID","Personel_Email","UrunID","Urun_Adi","SeriNo","Teslim_Tarihi","Teslim_Eden","Durum"]
  var zSheet = getSheet(CONFIG.SHEETS.ZIMMET);
  var zId = genId("ZMT", zSheet);
  zSheet.appendRow([
    zId, data.email, row[3] || "", row[4] || row[3] || "", "Sistem-" + data.talepId, now(), "Otomatik Sistem İşlemi", "Aktif"
  ]);

  // 3. STOKTAN DÜŞÜŞ
  if (row[3]) {
    updateStok(row[3], -(Number(row[5]) || 1));
  }

  return { success: true };
}

// ══════════════════════════════════════════════════════════════
//  İADE
// ══════════════════════════════════════════════════════════════
function processIade(zimmetId, email) {
  var user = getPersonelByEmail(email);
  if (!user || (user.rol !== "IT" && user.rol !== "Procurement" && user.rol !== "Director"))
    throw new Error("İade yetkiniz yok.");

  var sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  var r = findRow(sheet, 1, zimmetId);
  if (!r) throw new Error("Zimmet bulunamadı.");
  sheet.getRange(r, 8).setValue("İade Edildi");
  var uId = sheet.getRange(r, 3).getValue();
  if (uId) updateStok(uId, 1);
  return { success: true };
}

// ══════════════════════════════════════════════════════════════
//  VERİ OKUMA YARDIMCILARI
// ══════════════════════════════════════════════════════════════
function loadAllTalepler() {
  var s = getSheet(CONFIG.SHEETS.TALEPLER);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).map(function(r, i) {
    return {
      rowIndex: i+2, talepId: r[0], tarih: r[1], talepEdenEmail: r[2],
      urunId: r[3], urunAdi: r[4], miktar: r[5], durum: r[6],
      surecTipi: r[7], teknikNot: r[8], sonIslemYapan: r[12], sonIslemTarihi: r[13]
    };
  });
}

function getAllStok() {
  var s = getSheet(CONFIG.SHEETS.STOK);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).map(function(r) {
    return { urunId: r[0], urunAdi: r[1], kategori: r[2], mevcutStok: r[3], kritikSeviye: r[4], birim: r[5] };
  });
}

function getAllPersonel() {
  var s = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).map(function(r) {
    return { id: r[0], isim: r[1], email: r[2], departman: r[4], rol: r[6] };
  });
}

function getAllSatinalma() {
  var s = getSheet(CONFIG.SHEETS.SATINALMA);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).map(function(r) {
    return { satinalmaId: r[0], talepId: r[1], firma: r[2], birimFiyat: r[3], doviz: r[4], toplamTL: r[5], vade: r[6], odemeTipi: r[7], direktorDurum: r[8], butceKodu: r[9], kayitTarihi: r[10], islemYapan: r[11] };
  }).reverse();
}

function getZimmetByEmail(email) {
  var s = getSheet(CONFIG.SHEETS.ZIMMET);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).filter(function(r) {
    return String(r[1]).toLowerCase().trim() === email.toLowerCase().trim();
  }).map(function(r) {
    return { zimmetId: r[0], personelEmail: r[1], urunId: r[2], urunAdi: r[3], seriNo: r[4], teslimTarihi: r[5], teslimEden: r[6], durum: r[7] };
  }).reverse();
}

function getAllZimmet() {
  var s = getSheet(CONFIG.SHEETS.ZIMMET);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).map(function(r) {
    return { zimmetId: r[0], personelEmail: r[1], urunId: r[2], urunAdi: r[3], seriNo: r[4], teslimTarihi: r[5], teslimEden: r[6], durum: r[7] };
  }).reverse();
}

function getPersonelByEmail(email) {
  if (!email) return null;
  var s = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!s) return null;
  var d = s.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
      return { id: d[i][0], isim: d[i][1], email: d[i][2], departman: d[i][4], yoneticiEmail: d[i][5], rol: d[i][6] };
    }
  }
  return null;
}

function buildPersonelMap() {
  var s = getSheet(CONFIG.SHEETS.PERSONEL);
  var map = {};
  if (!s || s.getLastRow() < 2) return map;
  s.getDataRange().getValues().slice(1).forEach(function(r) {
    map[String(r[2]).toLowerCase().trim()] = { isim: r[1], departman: r[4], yoneticiEmail: r[5], rol: r[6] };
  });
  return map;
}

function isTeknikOnayci(email, urunId) {
  var kat = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  if (!kat) return true;
  var kategori = "";
  if (urunId) {
    var st = getSheet(CONFIG.SHEETS.STOK);
    if (st) { var sd = st.getDataRange().getValues(); for (var i = 1; i < sd.length; i++) { if (String(sd[i][0]) === String(urunId)) { kategori = sd[i][2] || ""; break; } } }
  }
  var kd = kat.getDataRange().getValues();
  for (var j = 1; j < kd.length; j++) {
    if (String(kd[j][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
      if (!kategori || String(kd[j][0]).toLowerCase() === kategori.toLowerCase()) return true;
    }
  }
  return false;
}

function updateStok(urunId, delta) {
  if (!urunId) return;
  var s = getSheet(CONFIG.SHEETS.STOK); if (!s) return;
  var d = s.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(urunId)) {
      s.getRange(i+1, 4).setValue(Math.max(0, (Number(d[i][3])||0) + delta));
      return;
    }
  }
}

// ══════════════════════════════════════════════════════════════
//  E-POSTA GÖNDERİM
// ══════════════════════════════════════════════════════════════
function sendApprovalEmail(toEmail, info, type, key) {
  var url = getAppUrl();
  var approveUrl = url + "?action=approve&id=" + info.talepId + "&key=" + key + "&type=" + type;
  var rejectUrl  = url + "?action=reject&id=" + info.talepId + "&key=" + key + "&type=" + type;

  var titles = { mudur: "Müdür Onayı Gerekiyor", teknik: "Teknik Birim Onayı", direktor: "Direktör Onayı" };
  var title = titles[type] || "Onay Gerekiyor";

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Segoe UI,Arial,sans-serif;background:#f4f6f9;padding:20px}' +
    '.card{background:#fff;border-radius:12px;max-width:560px;margin:0 auto;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08)}' +
    '.top{background:#0f172a;padding:24px 28px;color:#fff}.top h1{font-size:1.15rem;margin:6px 0 0}' +
    '.top .logo{font-size:.7rem;letter-spacing:.12em;color:#60a5fa}' +
    '.body{padding:24px 28px}' +
    'table{width:100%;border-collapse:collapse;margin:16px 0}td{padding:8px 0;border-bottom:1px solid #f0f0f0;font-size:.85rem}' +
    'td:first-child{color:#6b7280;width:140px}td:last-child{font-weight:500}' +
    '.actions{display:flex;gap:10px;margin-top:16px}' +
    '.btn{flex:1;text-decoration:none;text-align:center;padding:12px;border-radius:8px;font-weight:700;font-size:.85rem;color:#fff}' +
    '.btn-ok{background:#059669}.btn-no{background:#dc2626}' +
    '.foot{border-top:1px solid #f0f0f0;padding:14px 28px;font-size:.7rem;color:#9ca3af;text-align:center}' +
    '</style></head><body><div class="card"><div class="top"><div class="logo">◈ TEDARİK SİSTEMİ</div><h1>' + title + '</h1></div>' +
    '<div class="body"><table>' +
    '<tr><td>Talep ID</td><td><strong>' + info.talepId + '</strong></td></tr>' +
    '<tr><td>Talep Eden</td><td>' + (info.talepEden||'') + ' (' + (info.departman||'') + ')</td></tr>' +
    '<tr><td>Ürün</td><td>' + (info.urunAdi||'') + '</td></tr>' +
    '<tr><td>Miktar</td><td>' + (info.miktar||'') + '</td></tr>' +
    '<tr><td>Süreç</td><td>' + (info.surecTipi||'') + '</td></tr>' +
    (info.teknikNot ? '<tr><td>Not</td><td>' + info.teknikNot + '</td></tr>' : '') +
    (info.firma ? '<tr><td>Firma</td><td>' + info.firma + '</td></tr><tr><td>Toplam</td><td style="font-weight:700;color:#dc2626">' + info.toplamTutar + '</td></tr>' : '') +
    '</table><div class="actions"><a href="' + approveUrl + '" class="btn btn-ok">✓ ONAYLA</a><a href="' + rejectUrl + '" class="btn btn-no">✕ REDDET</a></div></div>' +
    '<div class="foot">Bu e-posta otomatik gönderilmiştir.</div></div></body></html>';

  try {
    GmailApp.sendEmail(toEmail, "[" + info.talepId + "] " + title, title + " — Onay bekleniyor.", { htmlBody: html });
  } catch(err) { Logger.log("E-posta gönderilemedi: " + err); }
}

function sendTeknikEmail(talepSheet, rowIdx) {
  var row = talepSheet.getRange(rowIdx, 1, 1, 14).getValues()[0];
  var kat = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  var teknikEmail = "";
  if (kat) {
    var stokSheet = getSheet(CONFIG.SHEETS.STOK);
    var category = "IT";
    if (stokSheet && row[3]) { var sd = stokSheet.getDataRange().getValues(); for (var i = 1; i < sd.length; i++) { if (String(sd[i][0]) === String(row[3])) { category = sd[i][2] || "IT"; break; } } }
    var kd = kat.getDataRange().getValues();
    for (var j = 1; j < kd.length; j++) { if (String(kd[j][0]).toLowerCase() === category.toLowerCase()) { teknikEmail = kd[j][2]; break; } }
  }
  if (!teknikEmail) teknikEmail = "caglar.acer@mcsistem.com.tr";
  var p = getPersonelByEmail(row[2]);
  sendApprovalEmail(teknikEmail, { talepId:row[0], talepEden:p?p.isim:row[2], departman:p?p.departman:"", urunAdi:row[4]||row[3], miktar:row[5], surecTipi:row[7], teknikNot:row[8] }, "teknik", row[10]);
}

function sendDirektorEmail(ts, tr, data, toplam) {
  var row = ts.getRange(tr, 1, 1, 14).getValues()[0];
  var ps = getSheet(CONFIG.SHEETS.PERSONEL);
  var dirEmail = "";
  if (ps) { var pd = ps.getDataRange().getValues(); for (var i = 1; i < pd.length; i++) { if (String(pd[i][6]).toLowerCase() === "director") { dirEmail = pd[i][2]; break; } } }
  if (!dirEmail) dirEmail = "caglar.acer@mcsistem.com.tr";
  var p = getPersonelByEmail(row[2]);
  var sym = {TRY:"₺",USD:"$",EUR:"€"}[data.doviz]||"₺";
  sendApprovalEmail(dirEmail, { talepId:row[0], talepEden:p?p.isim:row[2], departman:p?p.departman:"", urunAdi:row[4]||row[3], miktar:row[5], surecTipi:row[7], firma:data.firma, toplamTutar:sym+Number(toplam).toLocaleString("tr-TR",{minimumFractionDigits:2}), kritikAlim:true }, "direktor", row[11]);
}

// ══════════════════════════════════════════════════════════════
//  E-POSTA LİNK İŞLEME
// ══════════════════════════════════════════════════════════════
function handleEmailAction(p, act) {
  var sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  var rowIdx = findRow(sheet, 1, p.id);
  if (!rowIdx) return resultPage("⚠️", "Talep Bulunamadı", "", "#f59e0b");

  var keyMap = {mudur:10, teknik:11, direktor:12};
  var keyCol = keyMap[p.type] || 10;
  if (String(sheet.getRange(rowIdx, keyCol).getValue()) !== String(p.key))
    return resultPage("🔒", "Geçersiz Bağlantı", "Bu onay linki geçersiz veya kullanılmış.", "#ef4444");

  var durum = sheet.getRange(rowIdx, 7).getValue();
  var surec = sheet.getRange(rowIdx, 8).getValue();

  if (act === "reject") {
    sheet.getRange(rowIdx, 7).setValue("Reddedildi");
    sheet.getRange(rowIdx, 13).setValue("E-posta (" + p.type + ")");
    sheet.getRange(rowIdx, 14).setValue(now());
    return resultPage("❌", p.id + " — Reddedildi", "Talep reddedildi.", "#ef4444");
  }

  var newDurum = "";
  if (p.type === "mudur") newDurum = "Teknik Onay Bekliyor";
  else if (p.type === "teknik") newDurum = surec === "Stok" ? "Teslim Edilebilir" : "Satınalma Aşamasında";
  else if (p.type === "direktor") newDurum = "Sipariş Edilebilir";

  sheet.getRange(rowIdx, 7).setValue(newDurum);
  sheet.getRange(rowIdx, 13).setValue("E-posta (" + p.type + ")");
  sheet.getRange(rowIdx, 14).setValue(now());
  if (newDurum === "Teknik Onay Bekliyor") sendTeknikEmail(sheet, rowIdx);

  return resultPage("✅", p.id + " — Onaylandı", 'Yeni durum: "' + newDurum + '"', "#059669");
}

function resultPage(icon, title, msg, color) {
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet"><style>body{font-family:Inter,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f8fafc}.box{background:#fff;border-radius:16px;padding:56px 48px;text-align:center;max-width:440px;box-shadow:0 4px 24px rgba(0,0,0,.08);border:1px solid #e2e8f0}.icon{font-size:3.5rem;margin-bottom:16px;display:block}h2{color:' + color + ';margin-bottom:10px;font-size:1.2rem}p{color:#64748b;font-size:.88rem;line-height:1.6}</style></head><body><div class="box"><span class="icon">' + icon + '</span><h2>' + title + '</h2><p>' + msg + '</p><p style="margin-top:16px;font-size:.75rem;color:#94a3b8">Bu sekmeyi kapatabilirsiniz.</p></div></body></html>'
  ).setTitle("Tedarik Sistemi");
}

// ══════════════════════════════════════════════════════════════
//  KURULUM (bir kez çalıştırın)
// ══════════════════════════════════════════════════════════════
function setupAllSheets() {
  var ss = getSpreadsheet();
  makeSheet(ss, CONFIG.SHEETS.PERSONEL,    ["PersonelID","Ad_Soyad","Email","Sifre","Departman","Yonetici_Email","Rol"]);
  makeSheet(ss, CONFIG.SHEETS.STOK,        ["UrunID","Urun_Adi","Kategori","Mevcut_Stok","Kritik_Seviye","Birim"]);
  makeSheet(ss, CONFIG.SHEETS.KATEGORI_ONAY,["Kategori","Teknik_Birim","Teknik_Onayci_Email"]);
  makeSheet(ss, CONFIG.SHEETS.TALEPLER,    ["TalepID","Tarih","Talep_Eden_Email","UrunID","Urun_Adi","Miktar","Durum","Surec_Tipi","Teknik_Not","Mudur_Onay_Key","Teknik_Onay_Key","Direktor_Onay_Key","Son_Islem_Yapan","Son_Islem_Tarihi"]);
  makeSheet(ss, CONFIG.SHEETS.SATINALMA,   ["SatinalmaID","TalepID","Tedarikci","Birim_Fiyat","Doviz","Toplam_TL","Vade","Odeme_Tipi","Direktor_Durum","Butce_Kodu","Kayit_Tarihi","Islem_Yapan"]);
  makeSheet(ss, CONFIG.SHEETS.ZIMMET,      ["ZimmetID","Personel_Email","UrunID","Urun_Adi","SeriNo","Teslim_Tarihi","Teslim_Eden","Durum"]);
  Logger.log("✅ Tüm sayfalar oluşturuldu.");
}

function makeSheet(ss, name, headers) {
  var s = ss.getSheetByName(name) || ss.insertSheet(name);
  s.getRange(1,1,1,headers.length).setValues([headers]).setBackground("#0f172a").setFontColor("#60a5fa").setFontWeight("bold").setFontFamily("Courier New").setFontSize(10);
  s.setFrozenRows(1);
  headers.forEach(function(_,i){ s.setColumnWidth(i+1, Math.max(120, headers[i].length * 12)); });
  if (name === CONFIG.SHEETS.TALEPLER) { s.hideColumns(10); s.hideColumns(11); s.hideColumns(12); }
}

function seedSampleData() {
  var p = getSheet(CONFIG.SHEETS.PERSONEL);
  if (p.getLastRow() < 2) {
    p.appendRow([1,"Çağlar Acer","caglar.acer@mcsistem.com.tr","Feal3669","Bilgi Teknolojileri","","Director"]);
    p.appendRow([2,"Ahmet Yılmaz","ahmet.yilmaz@mcsistem.com.tr","123456","Üretim","ayse.demir@mcsistem.com.tr","User"]);
    p.appendRow([3,"Mehmet Kaya","mehmet.kaya@mcsistem.com.tr","123456","Lojistik","ayse.demir@mcsistem.com.tr","User"]);
    p.appendRow([4,"Ayşe Demir","ayse.demir@mcsistem.com.tr","123456","İnsan Kaynakları","caglar.acer@mcsistem.com.tr","Manager"]);
    p.appendRow([5,"Fatma Şahin","fatma.sahin@mcsistem.com.tr","123456","Satınalma","caglar.acer@mcsistem.com.tr","Procurement"]);
    p.appendRow([6,"Ali Öztürk","ali.ozturk@mcsistem.com.tr","123456","Bilgi Teknolojileri","caglar.acer@mcsistem.com.tr","IT"]);
  }
  var s = getSheet(CONFIG.SHEETS.STOK);
  if (s.getLastRow() < 2) {
    s.appendRow(["STK-001","Laptop (Dell Latitude 5540)","IT",5,2,"Adet"]);
    s.appendRow(["STK-002","Monitor (24\" LG IPS)","IT",8,3,"Adet"]);
    s.appendRow(["STK-003","Klavye + Mouse Set","IT",15,5,"Adet"]);
    s.appendRow(["STK-004","A4 Kağıt (500'lü)","İdari",120,30,"Paket"]);
    s.appendRow(["STK-005","Toner (HP LaserJet)","İdari",6,2,"Adet"]);
    s.appendRow(["STK-006","Ofis Koltuğu","İdari",4,2,"Adet"]);
  }
  var k = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  if (k.getLastRow() < 2) {
    k.appendRow(["IT","Bilgi Teknolojileri","ali.ozturk@mcsistem.com.tr"]);
    k.appendRow(["İdari","İdari İşler","ayse.demir@mcsistem.com.tr"]);
    k.appendRow(["Üretim","Üretim Müh.","caglar.acer@mcsistem.com.tr"]);
  }
  Logger.log("✅ Örnek veriler eklendi.");
}
