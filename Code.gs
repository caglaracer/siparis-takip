// ============================================================
//  KURUMSAL TEDARİK VE ZİMMET YÖNETİM SİSTEMİ — v3.0 RBAC
//  Google Apps Script Backend
//  Session.getActiveUser() ile Rol Bazlı Erişim Kontrolü
// ============================================================

// ──────────────────────────────────────────────────────────────
//  AYARLAR
// ──────────────────────────────────────────────────────────────
const CONFIG = {
  SPREADSHEET_ID: "1HI5BxkFKK23jH_yZJxxZD9XJ8hjdxmd5LY65Ca9Bnns",
  CRITICAL_AMOUNT: 100000,
  SHEETS: {
    PERSONEL:      "REF_Personel",
    STOK:          "REF_Stok",
    KATEGORI_ONAY: "REF_Kategori_Onay",
    TALEPLER:      "TRX_Talepler",
    SATINALMA:     "TRX_Satinalma",
    ZIMMET:        "TRX_Zimmet",
  },
};

// ──────────────────────────────────────────────────────────────
//  YARDIMCI FONKSİYONLAR
// ──────────────────────────────────────────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function nowFormatted() {
  return Utilities.formatDate(new Date(), "Europe/Istanbul", "dd.MM.yyyy HH:mm");
}

function newUuid() {
  return Utilities.getUuid().replace(/-/g, "").substring(0, 24);
}

function generateId(prefix, sheet) {
  const lastRow = sheet.getLastRow();
  const num = lastRow < 2 ? 1 : lastRow;
  return prefix + "-" + String(num).padStart(5, "0");
}

function findRowByColumnValue(sheet, colIndex, value) {
  const data = sheet.getRange(1, colIndex, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(value)) return i + 1;
  }
  return null;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
//  KİMLİK DOĞRULAMA VE YETKİLENDİRME (RBAC ÇEKİRDEĞİ)
// ══════════════════════════════════════════════════════════════

/**
 * Aktif Google oturumundan kullanıcı e-postasını alır
 * ve REF_Personel tablosunda arar.
 * Bulunamazsa null döner → erişim engellenir.
 */
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;

  const sheet = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
      return {
        personelId:    data[i][0],
        isim:          data[i][1],
        email:         data[i][2],
        departman:     data[i][3],
        yoneticiEmail: data[i][4],
        rol:           data[i][5],
        vekilEmail:    data[i][6],
      };
    }
  }
  return null;
}

/**
 * Rol yetki kontrolü.
 * requiredRoles: İzin verilen rollerin dizisi.
 */
function checkRole(user, requiredRoles) {
  if (!user) return false;
  return requiredRoles.includes(user.rol);
}

// ══════════════════════════════════════════════════════════════
//  doGet — ANA GİRİŞ NOKTASI
// ══════════════════════════════════════════════════════════════

function doGet(e) {
  const action = (e.parameter.action || "").toLowerCase();

  // ── E-posta onay linkleri (dış erişim, kullanıcı oturumu gerekmez)
  if (action === "approve" || action === "reject") {
    return handleEmailApproval(e);
  }

  // ── HTML Arayüzü Servis Et
  const user = getCurrentUser();

  if (!user) {
    // Kullanıcı tanınmıyor → Erişim engellendi sayfası
    return HtmlService.createHtmlOutput(buildAccessDeniedPage())
      .setTitle("Erişim Engellendi")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Ana uygulamayı renderla — kullanıcı bilgisi şablona enjekte edilir
  const template = HtmlService.createTemplateFromFile("Index");
  template.user = JSON.stringify(user);
  template.userRole = user.rol;
  template.userName = user.isim;
  template.userEmail = user.email;

  return template.evaluate()
    .setTitle("Tedarik ve Zimmet Yönetim Sistemi")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

/**
 * Styles.html veya JavaScript.html dosyalarını dahil etmek için.
 * Index.html içinde <?!= include('Styles') ?> şeklinde kullanılır.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ══════════════════════════════════════════════════════════════
//  ROL BAZLI VERİ ÇEKME — FRONTEND'DEN ÇAĞRILIR
//  (google.script.run ile)
// ══════════════════════════════════════════════════════════════

/**
 * Ana dashboard verisi — rol bazlı filtrelenmiş.
 * Her rol sadece kendi görmesi gereken verileri alır.
 */
function getDashboardData() {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz erişim.");

  const result = {
    user: user,
    taleplerim: [],
    onayBekleyenler: [],
    gecmisOnaylarim: [],
    stok: [],
    satinalma: [],
    zimmetlerim: [],
    tumTalepler: [],
    stats: {},
  };

  // ── Talepler
  const talepSheet = getSheet(CONFIG.SHEETS.TALEPLER);
  let allTalepler = [];
  if (talepSheet && talepSheet.getLastRow() > 1) {
    allTalepler = talepSheet.getDataRange().getValues().slice(1).map((r, i) => ({
      rowIndex:       i + 2,
      talepId:        r[0],
      tarih:          r[1],
      talepEdenEmail: r[2],
      urunId:         r[3],
      urunAdi:        r[4],
      miktar:         r[5],
      durum:          r[6],
      surecTipi:      r[7],
      teknikNot:      r[8],
      sonIslemYapan:  r[12],
      sonIslemTarihi: r[13],
    }));
  }

  // ── Personel listesi (isim eşleme için)
  const persSheet = getSheet(CONFIG.SHEETS.PERSONEL);
  let personelMap = {};
  if (persSheet && persSheet.getLastRow() > 1) {
    persSheet.getDataRange().getValues().slice(1).forEach(r => {
      personelMap[String(r[2]).toLowerCase().trim()] = {
        isim: r[1], departman: r[3], yoneticiEmail: r[4], rol: r[5],
      };
    });
  }

  // İsim bilgisi ekle
  allTalepler = allTalepler.map(t => {
    const p = personelMap[String(t.talepEdenEmail).toLowerCase().trim()];
    t.talepEdenIsim = p ? p.isim : t.talepEdenEmail;
    t.talepEdenDepartman = p ? p.departman : "";
    t.yoneticiEmail = p ? p.yoneticiEmail : "";
    return t;
  });

  // ═══ ROL BAZLI FİLTRELEME (switch-case) ═══
  switch (user.rol) {

    case "User":
      // Sadece kendi talepleri
      result.taleplerim = allTalepler.filter(t =>
        t.talepEdenEmail.toLowerCase() === user.email.toLowerCase()
      );
      // Kendi zimmetleri
      result.zimmetlerim = getZimmetByEmail(user.email);
      break;

    case "Manager":
      // Kendi talepleri
      result.taleplerim = allTalepler.filter(t =>
        t.talepEdenEmail.toLowerCase() === user.email.toLowerCase()
      );
      // Onay bekleyenler: Yönetici_Email kendisiyle eşleşen VE "Müdür Onayı Bekliyor" durumundakiler
      result.onayBekleyenler = allTalepler.filter(t =>
        t.yoneticiEmail.toLowerCase() === user.email.toLowerCase() &&
        t.durum === "Müdür Onayı Bekliyor"
      );
      // Geçmiş onayları: Daha önce onay/ret verdiği talepler
      result.gecmisOnaylarim = allTalepler.filter(t =>
        t.yoneticiEmail.toLowerCase() === user.email.toLowerCase() &&
        t.durum !== "Müdür Onayı Bekliyor"
      );
      // Kendi zimmetleri
      result.zimmetlerim = getZimmetByEmail(user.email);
      break;

    case "IT":
      // Kendi talepleri
      result.taleplerim = allTalepler.filter(t =>
        t.talepEdenEmail.toLowerCase() === user.email.toLowerCase()
      );
      // Teknik onay bekleyenler: IT kategorisine düşen ve "Teknik Onay Bekliyor"
      result.onayBekleyenler = allTalepler.filter(t =>
        t.durum === "Teknik Onay Bekliyor" &&
        isTeknikOnayci(user.email, t.urunId)
      );
      // Geçmiş
      result.gecmisOnaylarim = allTalepler.filter(t =>
        t.sonIslemYapan && t.sonIslemYapan.toLowerCase().includes(user.email.toLowerCase()) &&
        t.durum !== "Teknik Onay Bekliyor"
      );
      // Zimmet + Stok
      result.zimmetlerim = getZimmetByEmail(user.email);
      result.stok = getAllStokData();
      break;

    case "Procurement":
      // Kendi talepleri
      result.taleplerim = allTalepler.filter(t =>
        t.talepEdenEmail.toLowerCase() === user.email.toLowerCase()
      );
      // Satınalma aşamasındakiler
      result.onayBekleyenler = allTalepler.filter(t =>
        t.durum === "Satınalma Aşamasında"
      );
      // Sipariş edilenler
      result.gecmisOnaylarim = allTalepler.filter(t =>
        ["Sipariş Edilebilir", "Sipariş Edildi", "Teslim Alındı", "Teslim Edildi"].includes(t.durum)
      );
      // Satınalma kayıtları
      result.satinalma = getAllSatinalmaData();
      result.stok = getAllStokData();
      result.zimmetlerim = getZimmetByEmail(user.email);
      break;

    case "Director":
      // Direktör her şeyi görür
      result.taleplerim = allTalepler.filter(t =>
        t.talepEdenEmail.toLowerCase() === user.email.toLowerCase()
      );
      // Direktör onayı bekleyenler
      result.onayBekleyenler = allTalepler.filter(t =>
        t.durum === "Direktör Onayı Bekliyor"
      );
      // Tüm talepler (yönetim görünümü)
      result.tumTalepler = allTalepler;
      result.gecmisOnaylarim = allTalepler.filter(t =>
        !["Müdür Onayı Bekliyor", "Teknik Onay Bekliyor", "Direktör Onayı Bekliyor", "Satınalma Aşamasında"].includes(t.durum)
      );
      result.satinalma = getAllSatinalmaData();
      result.stok = getAllStokData();
      result.zimmetlerim = getAllZimmetData();
      break;

    default:
      throw new Error("Tanımsız rol: " + user.rol);
  }

  // ── İstatistikler
  const base = user.rol === "Director" ? allTalepler : result.taleplerim;
  result.stats = {
    toplam:       user.rol === "Director" ? allTalepler.length : result.taleplerim.length,
    onayBekleyen: result.onayBekleyenler.length,
    satinalma:    allTalepler.filter(t => t.durum === "Satınalma Aşamasında").length,
    siparis:      allTalepler.filter(t => t.durum && t.durum.includes("Sipariş")).length,
    teslim:       allTalepler.filter(t => t.durum && t.durum.includes("Teslim")).length,
    red:          allTalepler.filter(t => t.durum === "Reddedildi").length,
    kritikStok:   (result.stok || []).filter(s => Number(s.mevcutStok) <= Number(s.kritikSeviye)).length,
  };

  return result;
}

// ══════════════════════════════════════════════════════════════
//  TEKNİK ONAYCI KONTROLÜ
// ══════════════════════════════════════════════════════════════

function isTeknikOnayci(email, urunId) {
  const katSheet = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  if (!katSheet) return true; // Tablo yoksa herkes onaylayabilir

  // Ürünün kategorisini bul
  let kategori = "";
  if (urunId) {
    const stokSheet = getSheet(CONFIG.SHEETS.STOK);
    if (stokSheet) {
      const stokData = stokSheet.getDataRange().getValues();
      for (let i = 1; i < stokData.length; i++) {
        if (String(stokData[i][0]) === String(urunId)) {
          kategori = stokData[i][2] || "";
          break;
        }
      }
    }
  }

  // Kategori-Onay tablosunda bu e-postayı ara
  const katData = katSheet.getDataRange().getValues();
  for (let i = 1; i < katData.length; i++) {
    if (String(katData[i][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
      if (!kategori || String(katData[i][0]).toLowerCase() === kategori.toLowerCase()) {
        return true;
      }
    }
  }
  return false;
}

// ══════════════════════════════════════════════════════════════
//  VERİ ÇEKME YARDIMCILARI
// ══════════════════════════════════════════════════════════════

function getAllStokData() {
  const sheet = getSheet(CONFIG.SHEETS.STOK);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(r => ({
    urunId: r[0], urunAdi: r[1], kategori: r[2],
    mevcutStok: r[3], kritikSeviye: r[4], birim: r[5],
  }));
}

function getZimmetByEmail(email) {
  const sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => String(r[1]).toLowerCase().trim() === email.toLowerCase().trim())
    .map(r => ({
      zimmetId: r[0], personelEmail: r[1], urunId: r[2],
      urunAdi: r[3], seriNo: r[4], teslimTarihi: r[5],
      teslimEden: r[6], durum: r[7],
    })).reverse();
}

function getAllZimmetData() {
  const sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(r => ({
    zimmetId: r[0], personelEmail: r[1], urunId: r[2],
    urunAdi: r[3], seriNo: r[4], teslimTarihi: r[5],
    teslimEden: r[6], durum: r[7],
  })).reverse();
}

function getAllSatinalmaData() {
  const sheet = getSheet(CONFIG.SHEETS.SATINALMA);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(r => ({
    satinalmaId: r[0], talepId: r[1], firma: r[2],
    birimFiyat: r[3], doviz: r[4], tlKarsiligi: r[5],
    vade: r[6], odemeTipi: r[7], direktorOnayDurumu: r[8],
    butceKodu: r[9], kayitTarihi: r[10], islemYapan: r[11],
  })).reverse();
}

function getStokListForForm() {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz.");
  return getAllStokData();
}

function getPersonelListForForm() {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz.");
  // Sadece IT, Procurement ve Director personel listesini görebilir
  if (!checkRole(user, ["IT", "Procurement", "Director"])) return [];
  const sheet = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map(r => ({
    personelId: r[0], isim: r[1], email: r[2],
    departman: r[3], rol: r[5],
  }));
}

// ══════════════════════════════════════════════════════════════
//  YENİ TALEP OLUŞTURMA
// ══════════════════════════════════════════════════════════════

function createNewTalep(formData) {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz erişim.");

  const sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const talepId = generateId("TLP", sheet);
  const now = nowFormatted();

  // Stok kontrolü
  const route = checkStockAndRoute(formData.urunId, formData.miktar || 1);

  // Onay anahtarları
  const mudurKey    = newUuid();
  const teknikKey   = newUuid();
  const direktorKey = newUuid();

  sheet.appendRow([
    talepId,                             // A: TalepID
    now,                                 // B: Tarih
    user.email,                          // C: Talep_Eden_Email (oturumdaki kullanıcı)
    formData.urunId || "",               // D: UrunID
    formData.urunAdi || "",              // E: Urun_Adi
    formData.miktar || 1,                // F: Miktar
    route.ilkDurum,                      // G: Durum
    route.surecTipi,                     // H: Surec_Tipi
    formData.teknikNot || "",            // I: Teknik_Not
    mudurKey,                            // J: Mudur_Onay_Key
    teknikKey,                           // K: Teknik_Onay_Key
    direktorKey,                         // L: Direktor_Onay_Key
    user.email,                          // M: Son_Islem_Yapan
    now,                                 // N: Son_Islem_Tarihi
  ]);

  // Müdüre onay e-postası
  if (user.yoneticiEmail) {
    sendApprovalEmail(
      user.yoneticiEmail,
      {
        talepId: talepId,
        talepEden: user.isim,
        departman: user.departman,
        urunAdi: formData.urunAdi || formData.urunId || "—",
        miktar: formData.miktar || 1,
        surecTipi: route.surecTipi,
        teknikNot: formData.teknikNot || "",
      },
      "mudur", mudurKey
    );
  }

  return {
    success: true,
    talepId: talepId,
    surecTipi: route.surecTipi,
    mesaj: route.mesaj,
  };
}

// ══════════════════════════════════════════════════════════════
//  STOK KONTROLÜ VE YÖNLENDİRME
// ══════════════════════════════════════════════════════════════

function checkStockAndRoute(urunId, miktar) {
  if (!urunId) {
    return {
      surecTipi: "Satınalma",
      ilkDurum: "Müdür Onayı Bekliyor",
      mesaj: "Ürün stokta tanımlı değil. Satınalma sürecine yönlendirildi.",
    };
  }

  const stokSheet = getSheet(CONFIG.SHEETS.STOK);
  if (!stokSheet) {
    return {
      surecTipi: "Satınalma",
      ilkDurum: "Müdür Onayı Bekliyor",
      mesaj: "Stok tablosu bulunamadı.",
    };
  }

  const data = stokSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(urunId)) {
      const mevcutStok = Number(data[i][3]) || 0;
      if (mevcutStok >= miktar) {
        return {
          surecTipi: "Stok",
          ilkDurum: "Müdür Onayı Bekliyor",
          mesaj: "Stokta mevcut (" + mevcutStok + " adet). Müdür onayına gönderildi.",
        };
      } else {
        return {
          surecTipi: "Satınalma",
          ilkDurum: "Müdür Onayı Bekliyor",
          mesaj: "Stok yetersiz (" + mevcutStok + "/" + miktar + "). Satınalma sürecine yönlendirildi.",
        };
      }
    }
  }

  return {
    surecTipi: "Satınalma",
    ilkDurum: "Müdür Onayı Bekliyor",
    mesaj: "Ürün stok tanımında bulunamadı. Satınalma sürecine yönlendirildi.",
  };
}

// ══════════════════════════════════════════════════════════════
//  ONAY İŞLEMİ (Dashboard butonundan — google.script.run)
// ══════════════════════════════════════════════════════════════

/**
 * Frontend'den çağrılır. Kullanıcının rolüne göre yetki kontrolü yapılır.
 * @param {string} talepId - Talep numarası
 * @param {string} aksiyonTipi - "onayla" veya "reddet"
 */
function processApproval(talepId, aksiyonTipi) {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz erişim.");

  const sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const rowIndex = findRowByColumnValue(sheet, 1, talepId);
  if (!rowIndex) throw new Error("Talep bulunamadı: " + talepId);

  const row = sheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
  const currentDurum = row[6];
  const surecTipi = row[7];
  const talepEdenEmail = row[2];

  // ── YETKİ KONTROLÜ — Sadece yetkili rol bu işlemi yapabilir
  let authorized = false;
  let newDurum = "";

  switch (currentDurum) {
    case "Müdür Onayı Bekliyor":
      // Sadece Manager (ve talep edenin yöneticisi) veya Director onaylayabilir
      if (checkRole(user, ["Manager", "Director"])) {
        // Manager ise, sadece kendi ekibinin taleplerini onaylayabilir
        if (user.rol === "Manager") {
          const talepEdenInfo = getPersonelByEmail(talepEdenEmail);
          if (talepEdenInfo && talepEdenInfo.yoneticiEmail.toLowerCase() === user.email.toLowerCase()) {
            authorized = true;
          }
        } else {
          authorized = true; // Director her şeyi onaylayabilir
        }
      }
      if (aksiyonTipi === "onayla") {
        newDurum = "Teknik Onay Bekliyor";
      } else {
        newDurum = "Reddedildi";
      }
      break;

    case "Teknik Onay Bekliyor":
      // Sadece IT veya Director
      if (checkRole(user, ["IT", "Director"])) {
        if (user.rol === "IT") {
          authorized = isTeknikOnayci(user.email, row[3]);
        } else {
          authorized = true;
        }
      }
      if (aksiyonTipi === "onayla") {
        newDurum = surecTipi === "Stok" ? "Teslim Edilebilir" : "Satınalma Aşamasında";
      } else {
        newDurum = "Reddedildi";
      }
      break;

    case "Direktör Onayı Bekliyor":
      // Sadece Director
      if (checkRole(user, ["Director"])) {
        authorized = true;
      }
      if (aksiyonTipi === "onayla") {
        newDurum = "Sipariş Edilebilir";
      } else {
        newDurum = "Reddedildi";
      }
      break;

    case "Satınalma Aşamasında":
      // Sadece Procurement veya Director
      if (checkRole(user, ["Procurement", "Director"])) {
        authorized = true;
      }
      newDurum = aksiyonTipi === "onayla" ? "Sipariş Edilebilir" : "Reddedildi";
      break;

    case "Sipariş Edilebilir":
      if (checkRole(user, ["Procurement", "Director"])) {
        authorized = true;
      }
      newDurum = "Sipariş Edildi";
      break;

    case "Sipariş Edildi":
      if (checkRole(user, ["Procurement", "IT", "Director"])) {
        authorized = true;
      }
      newDurum = "Teslim Alındı";
      break;

    case "Teslim Alındı":
    case "Teslim Edilebilir":
      if (checkRole(user, ["IT", "Procurement", "Director"])) {
        authorized = true;
      }
      newDurum = "Teslim Edildi";
      break;

    default:
      throw new Error("Bu talep üzerinde işlem yapılamaz. Mevcut durum: " + currentDurum);
  }

  if (!authorized) {
    throw new Error("Bu işlem için yetkiniz bulunmamaktadır. Rol: " + user.rol);
  }

  // ── Durumu güncelle
  sheet.getRange(rowIndex, 7).setValue(newDurum);
  auditLog(sheet, rowIndex, user.email);

  // ── Sonraki aşamaya e-posta gönder
  if (aksiyonTipi === "onayla") {
    if (newDurum === "Teknik Onay Bekliyor") {
      sendTeknikOnayEmail(sheet, rowIndex);
    }
    // Stok güncelleme: Teslim edildiyse
    if (newDurum === "Teslim Alındı" && row[3]) {
      updateStokMiktar(row[3], Number(row[5]) || 1);
    }
  }

  return {
    success: true,
    talepId: talepId,
    yeniDurum: newDurum,
    mesaj: aksiyonTipi === "onayla"
      ? "Talep onaylandı: " + newDurum
      : "Talep reddedildi.",
  };
}

// ══════════════════════════════════════════════════════════════
//  SATINALMA KAYDI OLUŞTURMA
// ══════════════════════════════════════════════════════════════

function createSatinalmaKaydi(formData) {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz erişim.");
  if (!checkRole(user, ["Procurement", "Director"])) {
    throw new Error("Satınalma kaydı oluşturma yetkiniz yok.");
  }

  const sheet = getSheet(CONFIG.SHEETS.SATINALMA);
  const satId = generateId("SAT", sheet);
  const now = nowFormatted();
  const toplamTL = (formData.birimFiyat || 0) * (formData.miktar || 1);
  const kritikAlim = toplamTL >= CONFIG.CRITICAL_AMOUNT;

  sheet.appendRow([
    satId,
    formData.talepId || "",
    formData.firma || "",
    formData.birimFiyat || 0,
    formData.doviz || "TRY",
    "",  // TL formülü
    formData.vade || "",
    formData.odemeTipi || "",
    "Beklemede",
    formData.butceKodu || "",
    now,
    user.email,
  ]);

  // TL formülü
  const row = sheet.getLastRow();
  const formula = '=IF(E' + row + '="TRY", D' + row + '*' + (formData.miktar || 1) + ', D' + row + '*' + (formData.miktar || 1) + '*IFERROR(GOOGLEFINANCE("CURRENCY:"&E' + row + '&"TRY"),1))';
  sheet.getRange(row, 6).setFormula(formula);

  // Talep durumunu güncelle
  const talepSheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const talepRow = findRowByColumnValue(talepSheet, 1, formData.talepId);
  if (talepRow) {
    const nextDurum = kritikAlim ? "Direktör Onayı Bekliyor" : "Sipariş Edilebilir";
    talepSheet.getRange(talepRow, 7).setValue(nextDurum);
    auditLog(talepSheet, talepRow, user.email);

    // Kritik alımsa Direktöre e-posta
    if (kritikAlim) {
      sendDirektorOnay(talepSheet, talepRow, formData, toplamTL);
    }
  }

  return {
    success: true,
    satinalmaId: satId,
    kritikAlim: kritikAlim,
    mesaj: kritikAlim
      ? "KRİTİK ALIM — Direktör onayına gönderildi."
      : "Satınalma kaydı oluşturuldu.",
  };
}

// ══════════════════════════════════════════════════════════════
//  ZİMMET KAYDI OLUŞTURMA
// ══════════════════════════════════════════════════════════════

function createZimmetKaydi(formData) {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz erişim.");
  if (!checkRole(user, ["IT", "Procurement", "Director"])) {
    throw new Error("Zimmet oluşturma yetkiniz yok. Sadece IT veya İdari İşler onayıyla zimmet oluşturulabilir.");
  }

  const sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  const zimmetId = generateId("ZMT", sheet);
  const now = nowFormatted();

  sheet.appendRow([
    zimmetId,
    formData.personelEmail || "",
    formData.urunId || "",
    formData.urunAdi || "",
    formData.seriNo || "",
    now,
    user.email,
    "Aktif",
  ]);

  // Stok güncelle
  if (formData.urunId) {
    updateStokMiktar(formData.urunId, -(formData.miktar || 1));
  }

  // Talep durumunu güncelle
  if (formData.talepId) {
    const talepSheet = getSheet(CONFIG.SHEETS.TALEPLER);
    const talepRow = findRowByColumnValue(talepSheet, 1, formData.talepId);
    if (talepRow) {
      talepSheet.getRange(talepRow, 7).setValue("Teslim Edildi");
      auditLog(talepSheet, talepRow, user.email);
    }
  }

  return { success: true, zimmetId: zimmetId, mesaj: "Zimmet kaydı oluşturuldu." };
}

// ══════════════════════════════════════════════════════════════
//  İADE İŞLEMİ
// ══════════════════════════════════════════════════════════════

function processIade(zimmetId) {
  const user = getCurrentUser();
  if (!user) throw new Error("Yetkisiz erişim.");
  if (!checkRole(user, ["IT", "Procurement", "Director"])) {
    throw new Error("İade işlemi yetkiniz yok.");
  }

  const sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  const rowIndex = findRowByColumnValue(sheet, 1, zimmetId);
  if (!rowIndex) throw new Error("Zimmet bulunamadı.");

  sheet.getRange(rowIndex, 8).setValue("İade Edildi");
  const urunId = sheet.getRange(rowIndex, 3).getValue();
  if (urunId) updateStokMiktar(urunId, 1);

  return { success: true, mesaj: "Ürün iade edildi, stok güncellendi." };
}

// ══════════════════════════════════════════════════════════════
//  STOK MİKTAR GÜNCELLEME
// ══════════════════════════════════════════════════════════════

function updateStokMiktar(urunId, delta) {
  if (!urunId) return;
  const sheet = getSheet(CONFIG.SHEETS.STOK);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(urunId)) {
      const current = Number(data[i][3]) || 0;
      sheet.getRange(i + 1, 4).setValue(Math.max(0, current + delta));
      return;
    }
  }
}

// ══════════════════════════════════════════════════════════════
//  PERSONEL BİLGİSİ
// ══════════════════════════════════════════════════════════════

function getPersonelByEmail(email) {
  if (!email) return null;
  const sheet = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
      return {
        personelId: data[i][0], isim: data[i][1], email: data[i][2],
        departman: data[i][3], yoneticiEmail: data[i][4], rol: data[i][5],
        vekilEmail: data[i][6],
      };
    }
  }
  return null;
}

// ══════════════════════════════════════════════════════════════
//  E-POSTA GÖNDERİM FONKSİYONLARI
// ══════════════════════════════════════════════════════════════

function sendApprovalEmail(targetEmail, data, type, approvalKey) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const approveUrl = webAppUrl + "?action=approve&id=" + data.talepId + "&key=" + approvalKey + "&type=" + type;
  const rejectUrl  = webAppUrl + "?action=reject&id=" + data.talepId + "&key=" + approvalKey + "&type=" + type;

  let title = "";
  let badge = "";
  let extraInfo = "";

  switch (type) {
    case "mudur":
      title = "Departman Müdürü Onayı Gerekiyor";
      badge = "MÜDÜR ONAYI BEKLİYOR";
      break;
    case "teknik":
      title = "Teknik Birim Onayı Gerekiyor";
      badge = "TEKNİK ONAY BEKLİYOR";
      break;
    case "direktor":
      title = "Direktör Onayı Gerekiyor";
      badge = data.kritikAlim ? "KRİTİK ALIM — DİREKTÖR ONAYI" : "DİREKTÖR ONAYI BEKLİYOR";
      if (data.firma) {
        extraInfo = '<tr><td>Firma</td><td>' + data.firma + '</td></tr>' +
          '<tr><td>Toplam Tutar</td><td style="font-weight:700;color:' + (data.kritikAlim ? '#dc2626' : '#059669') + '">' + data.toplamTutar + '</td></tr>';
      }
      break;
  }

  const htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Segoe UI,Arial,sans-serif;background:#f4f6f9;padding:20px;color:#1a1a2e}' +
    '.card{background:#fff;border-radius:12px;max-width:580px;margin:0 auto;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08)}' +
    '.top{background:#0f172a;padding:28px 32px;color:#fff}.top .logo{font-size:.72rem;letter-spacing:.12em;color:#60a5fa;text-transform:uppercase;margin-bottom:6px}' +
    '.top h1{font-size:1.25rem;font-weight:700;margin:0}.body{padding:28px 32px}' +
    '.badge{display:inline-block;background:#fef3c7;color:#92400e;font-size:.72rem;font-weight:600;padding:5px 12px;border-radius:5px;margin-bottom:20px}' +
    'table.detail{width:100%;border-collapse:collapse;margin-bottom:24px}table.detail td{padding:10px 0;border-bottom:1px solid #f0f0f0;font-size:.875rem}' +
    'table.detail td:first-child{color:#6b7280;width:150px}table.detail td:last-child{font-weight:500}' +
    '.actions{display:flex;gap:12px;margin-top:8px}' +
    '.btn-approve{flex:1;background:#059669;color:#fff;text-decoration:none;text-align:center;padding:14px;border-radius:8px;font-weight:700;font-size:.9rem}' +
    '.btn-reject{flex:1;background:#dc2626;color:#fff;text-decoration:none;text-align:center;padding:14px;border-radius:8px;font-weight:700;font-size:.9rem}' +
    '.footer{border-top:1px solid #f0f0f0;padding:16px 32px;font-size:.72rem;color:#9ca3af;text-align:center}' +
    '</style></head><body><div class="card"><div class="top"><div class="logo">◈ Tedarik Yönetim Sistemi</div>' +
    '<h1>' + title + '</h1></div><div class="body"><div class="badge">' + badge + ' — ' + data.talepId + '</div>' +
    '<table class="detail"><tr><td>Talep Eden</td><td>' + (data.talepEden || '—') + ' (' + (data.departman || '') + ')</td></tr>' +
    '<tr><td>Ürün</td><td>' + (data.urunAdi || '—') + '</td></tr>' +
    '<tr><td>Miktar</td><td>' + (data.miktar || '—') + '</td></tr>' +
    '<tr><td>Süreç</td><td>' + (data.surecTipi || '—') + '</td></tr>' +
    (data.teknikNot ? '<tr><td>Not</td><td>' + data.teknikNot + '</td></tr>' : '') +
    extraInfo +
    '</table><div class="actions"><a href="' + approveUrl + '" class="btn-approve">✓ ONAYLA</a>' +
    '<a href="' + rejectUrl + '" class="btn-reject">✕ REDDET</a></div></div>' +
    '<div class="footer">Bu e-posta otomatik gönderilmiştir.</div></div></body></html>';

  try {
    GmailApp.sendEmail(targetEmail,
      "[" + data.talepId + "] " + title + ": " + (data.urunAdi || ""),
      title + " — " + (data.talepEden || "") + ". Onay için: " + approveUrl,
      { htmlBody: htmlBody }
    );
  } catch (err) {
    Logger.log("E-posta gönderilemedi: " + err.toString());
  }
}

function sendTeknikOnayEmail(talepSheet, rowIndex) {
  const row = talepSheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
  const urunAdi = row[4] || row[3];
  const talepId = row[0];
  const teknikKey = row[10];

  // Kategori-Onay tablosundan teknik onayıcı bul
  const katSheet = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  let teknikEmail = "";
  if (katSheet) {
    const stokSheet = getSheet(CONFIG.SHEETS.STOK);
    let kat = "IT";
    if (stokSheet && row[3]) {
      const stokData = stokSheet.getDataRange().getValues();
      for (let i = 1; i < stokData.length; i++) {
        if (String(stokData[i][0]) === String(row[3])) { kat = stokData[i][2] || "IT"; break; }
      }
    }
    const katData = katSheet.getDataRange().getValues();
    for (let i = 1; i < katData.length; i++) {
      if (String(katData[i][0]).toLowerCase() === kat.toLowerCase()) {
        teknikEmail = katData[i][2]; break;
      }
    }
  }
  if (!teknikEmail) teknikEmail = "caglar.acer@mcsistem.com.tr";

  const personelInfo = getPersonelByEmail(row[2]);
  sendApprovalEmail(teknikEmail, {
    talepId: talepId, talepEden: personelInfo ? personelInfo.isim : row[2],
    departman: personelInfo ? personelInfo.departman : "",
    urunAdi: urunAdi, miktar: row[5], surecTipi: row[7], teknikNot: row[8],
  }, "teknik", teknikKey);
}

function sendDirektorOnay(talepSheet, talepRow, formData, toplamTL) {
  const row = talepSheet.getRange(talepRow, 1, 1, 14).getValues()[0];
  const direktorKey = row[11];
  const persSheet = getSheet(CONFIG.SHEETS.PERSONEL);
  let direktorEmail = "";
  if (persSheet) {
    const pData = persSheet.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][5]).toLowerCase() === "director") { direktorEmail = pData[i][2]; break; }
    }
  }
  if (!direktorEmail) direktorEmail = "caglar.acer@mcsistem.com.tr";

  const sym = {TRY:"₺",USD:"$",EUR:"€"}[formData.doviz] || "₺";
  const personelInfo = getPersonelByEmail(row[2]);
  sendApprovalEmail(direktorEmail, {
    talepId: row[0], talepEden: personelInfo ? personelInfo.isim : row[2],
    departman: personelInfo ? personelInfo.departman : "",
    urunAdi: row[4] || row[3], miktar: row[5], surecTipi: row[7],
    firma: formData.firma, toplamTutar: "₺" + Number(toplamTL).toLocaleString("tr-TR", {minimumFractionDigits:2}),
    odemeTipi: formData.odemeTipi, vade: formData.vade, butceKodu: formData.butceKodu,
    kritikAlim: true,
  }, "direktor", direktorKey);
}

// ══════════════════════════════════════════════════════════════
//  E-POSTA ONAY LİNKİ İŞLEME
// ══════════════════════════════════════════════════════════════

function handleEmailApproval(e) {
  const talepId = e.parameter.id;
  const key = e.parameter.key;
  const action = e.parameter.action;
  const type = e.parameter.type || "mudur";

  const sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const rowIndex = findRowByColumnValue(sheet, 1, talepId);
  if (!rowIndex) return createResultPage("⚠️", "Talep Bulunamadı", "Bu talep ID geçerli değil.", "#f59e0b");

  const keyColMap = {mudur:10, teknik:11, direktor:12};
  const keyCol = keyColMap[type] || 10;
  const storedKey = sheet.getRange(rowIndex, keyCol).getValue();
  if (String(storedKey) !== String(key)) {
    return createResultPage("🔒", "Geçersiz Bağlantı", "Bu onay bağlantısı geçersiz.", "#ef4444");
  }

  const currentDurum = sheet.getRange(rowIndex, 7).getValue();
  const surecTipi = sheet.getRange(rowIndex, 8).getValue();

  if (action === "reject") {
    sheet.getRange(rowIndex, 7).setValue("Reddedildi");
    auditLog(sheet, rowIndex, "E-posta (" + type + ")");
    return createResultPage("❌", "Talep " + talepId + " — Reddedildi", "Talep reddedildi.", "#ef4444");
  }

  let newDurum = "";
  if (type === "mudur") {
    newDurum = "Teknik Onay Bekliyor";
  } else if (type === "teknik") {
    newDurum = surecTipi === "Stok" ? "Teslim Edilebilir" : "Satınalma Aşamasında";
  } else if (type === "direktor") {
    newDurum = "Sipariş Edilebilir";
  }

  sheet.getRange(rowIndex, 7).setValue(newDurum);
  auditLog(sheet, rowIndex, "E-posta (" + type + ")");

  if (newDurum === "Teknik Onay Bekliyor") {
    sendTeknikOnayEmail(sheet, rowIndex);
  }

  return createResultPage("✅", "Talep " + talepId + " — Onaylandı", 'Durum güncellendi: "' + newDurum + '"', "#059669");
}

// ══════════════════════════════════════════════════════════════
//  YARDIMCI: AUDIT LOG, SONUÇ SAYFASI
// ══════════════════════════════════════════════════════════════

function auditLog(sheet, rowIndex, user) {
  sheet.getRange(rowIndex, 13).setValue(user);
  sheet.getRange(rowIndex, 14).setValue(nowFormatted());
}

function createResultPage(icon, title, message, color) {
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">' +
    '<style>body{font-family:Inter,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f8fafc}' +
    '.box{background:#fff;border-radius:16px;padding:56px 48px;text-align:center;max-width:440px;box-shadow:0 4px 24px rgba(0,0,0,.08);border:1px solid #e2e8f0}' +
    '.icon{font-size:3.5rem;margin-bottom:20px;display:block}h2{color:' + color + ';font-size:1.3rem;margin-bottom:12px;font-weight:700}' +
    'p{color:#64748b;font-size:.9rem;line-height:1.6}</style></head>' +
    '<body><div class="box"><span class="icon">' + icon + '</span><h2>' + title + '</h2><p>' + message + '</p>' +
    '<p style="margin-top:20px;font-size:.78rem;color:#94a3b8">Bu sekmeyi kapatabilirsiniz.</p></div></body></html>'
  );
}

function buildAccessDeniedPage() {
  const email = Session.getActiveUser().getEmail() || "Bilinmiyor";
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">' +
    '<style>body{font-family:Inter,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f8fafc}' +
    '.box{background:#fff;border-radius:16px;padding:56px 48px;text-align:center;max-width:480px;box-shadow:0 4px 24px rgba(0,0,0,.08);border:1px solid #e2e8f0}' +
    '.icon{font-size:3.5rem;margin-bottom:20px;display:block}h2{color:#dc2626;font-size:1.3rem;margin-bottom:12px;font-weight:700}' +
    'p{color:#64748b;font-size:.9rem;line-height:1.6}.email{font-family:monospace;background:#f1f5f9;padding:4px 10px;border-radius:4px;font-size:.85rem}</style></head>' +
    '<body><div class="box"><span class="icon">🚫</span><h2>Erişim Engellendi</h2>' +
    '<p>Bu sisteme erişim yetkiniz bulunmamaktadır.</p>' +
    '<p style="margin-top:12px">Giriş yapılan hesap: <span class="email">' + email + '</span></p>' +
    '<p style="margin-top:16px;font-size:.78rem;color:#94a3b8">Bu bir hata ise sistem yöneticinize başvurun.</p></div></body></html>';
}

// ══════════════════════════════════════════════════════════════
//  İLK KURULUM — Bir kez çalıştırın
// ══════════════════════════════════════════════════════════════

function setupAllSheets() {
  const ss = getSpreadsheet();

  createOrResetSheet(ss, CONFIG.SHEETS.PERSONEL, [
    "PersonelID", "Isim", "Email", "Departman", "Yonetici_Email", "Rol", "Vekil_Email"
  ]);
  createOrResetSheet(ss, CONFIG.SHEETS.STOK, [
    "UrunID", "Urun_Adi", "Kategori", "Mevcut_Stok", "Kritik_Seviye", "Birim"
  ]);
  createOrResetSheet(ss, CONFIG.SHEETS.KATEGORI_ONAY, [
    "Kategori", "Teknik_Birim", "Teknik_Onayci_Email"
  ]);
  createOrResetSheet(ss, CONFIG.SHEETS.TALEPLER, [
    "TalepID", "Tarih", "Talep_Eden_Email", "UrunID", "Urun_Adi", "Miktar",
    "Durum", "Surec_Tipi", "Teknik_Not",
    "Mudur_Onay_Key", "Teknik_Onay_Key", "Direktor_Onay_Key",
    "Son_Islem_Yapan", "Son_Islem_Tarihi"
  ]);
  createOrResetSheet(ss, CONFIG.SHEETS.SATINALMA, [
    "SatinalmaID", "TalepID", "Firma", "Birim_Fiyat", "Doviz", "TL_Karsiligi",
    "Vade", "Odeme_Tipi", "Direktor_Onay_Durumu", "Butce_Kodu",
    "Kayit_Tarihi", "Islem_Yapan"
  ]);
  createOrResetSheet(ss, CONFIG.SHEETS.ZIMMET, [
    "ZimmetID", "Personel_Email", "UrunID", "Urun_Adi", "SeriNo",
    "Teslim_Tarihi", "Teslim_Eden", "Durum"
  ]);
  Logger.log("✅ Tüm sayfalar oluşturuldu.");
}

function createOrResetSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground("#0f172a");
  hRange.setFontColor("#60a5fa");
  hRange.setFontWeight("bold");
  hRange.setFontFamily("Courier New");
  hRange.setFontSize(10);
  sheet.setFrozenRows(1);
  for (let i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, Math.max(120, headers[i-1].length * 12));
  }
  if (name === CONFIG.SHEETS.TALEPLER) {
    sheet.hideColumns(10); sheet.hideColumns(11); sheet.hideColumns(12);
  }
}

function seedSampleData() {
  const p = getSheet(CONFIG.SHEETS.PERSONEL);
  if (p.getLastRow() < 2) {
    p.appendRow([1,"Çağlar Acer","caglar.acer@mcsistem.com.tr","Bilgi Teknolojileri","","Director",""]);
    p.appendRow([2,"Ahmet Yılmaz","ahmet.yilmaz@mcsistem.com.tr","Üretim","caglar.acer@mcsistem.com.tr","User",""]);
    p.appendRow([3,"Mehmet Kaya","mehmet.kaya@mcsistem.com.tr","Lojistik","caglar.acer@mcsistem.com.tr","User",""]);
    p.appendRow([4,"Ayşe Demir","ayse.demir@mcsistem.com.tr","İnsan Kaynakları","caglar.acer@mcsistem.com.tr","Manager",""]);
    p.appendRow([5,"Fatma Şahin","fatma.sahin@mcsistem.com.tr","Satınalma","caglar.acer@mcsistem.com.tr","Procurement",""]);
    p.appendRow([6,"Ali Öztürk","ali.ozturk@mcsistem.com.tr","Bilgi Teknolojileri","caglar.acer@mcsistem.com.tr","IT",""]);
  }
  const s = getSheet(CONFIG.SHEETS.STOK);
  if (s.getLastRow() < 2) {
    s.appendRow(["STK-001","Laptop (Dell Latitude 5540)","IT",5,2,"Adet"]);
    s.appendRow(["STK-002","Monitor (24\" LG)","IT",8,3,"Adet"]);
    s.appendRow(["STK-003","Klavye + Mouse Set","IT",15,5,"Adet"]);
    s.appendRow(["STK-004","A4 Kağıt (500'lü)","İdari",120,30,"Paket"]);
    s.appendRow(["STK-005","Toner (HP LaserJet)","İdari",6,2,"Adet"]);
    s.appendRow(["STK-006","Ofis Koltuğu","İdari",4,2,"Adet"]);
  }
  const k = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  if (k.getLastRow() < 2) {
    k.appendRow(["IT","Bilgi Teknolojileri","ali.ozturk@mcsistem.com.tr"]);
    k.appendRow(["İdari","İdari İşler","ayse.demir@mcsistem.com.tr"]);
    k.appendRow(["Üretim","Üretim Mühendisliği","caglar.acer@mcsistem.com.tr"]);
  }
  Logger.log("✅ Örnek veriler eklendi.");
}
