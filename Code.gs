// ============================================================
//  KURUMSAL TEDARİK VE ZİMMET YÖNETİM SİSTEMİ
//  Google Apps Script Backend — v2.0
// ============================================================

// ──────────────────────────────────────────────────────────────
//  AYARLAR
// ──────────────────────────────────────────────────────────────
const CONFIG = {
  SPREADSHEET_ID: "1HI5BxkFKK23jH_yZJxxZD9XJ8hjdxmd5LY65Ca9Bnns",
  APP_URL: "https://script.google.com/macros/s/AKfycbz9EpgfR15uESBCJMDaovg6KWJmFiw2p11kffWxKLwHxjuIClb9Y-J0bCGBV9lAv0u2/exec",
  ACCESS_PASSWORD: "Feal3669",
  CRITICAL_AMOUNT: 100000, // TL - Direktör onayı eşiği
  SHEETS: {
    PERSONEL:       "REF_Personel",
    STOK:           "REF_Stok",
    KATEGORI_ONAY:  "REF_Kategori_Onay",
    TALEPLER:       "TRX_Talepler",
    SATINALMA:      "TRX_Satinalma",
    ZIMMET:         "TRX_Zimmet",
  },
};

// ──────────────────────────────────────────────────────────────
//  YARDIMCI / UTILITY FONKSİYONLARI
// ──────────────────────────────────────────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateId(prefix, sheet) {
  const lastRow = sheet.getLastRow();
  const num = lastRow < 2 ? 1 : lastRow; // Başlık satırı hariç
  return `${prefix}-${String(num).padStart(5, "0")}`;
}

function nowFormatted() {
  return Utilities.formatDate(new Date(), "Europe/Istanbul", "dd.MM.yyyy HH:mm");
}

function newUuid() {
  return Utilities.getUuid().replace(/-/g, "").substring(0, 24);
}

function findRowByColumnValue(sheet, colIndex, value) {
  const data = sheet.getRange(1, colIndex, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(value)) return i + 1;
  }
  return null;
}

// ──────────────────────────────────────────────────────────────
//  doGet — Ana Router
// ──────────────────────────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter.action || "").toLowerCase();

  try {
    switch (action) {
      // Veri çekme
      case "getall":          return getAllTalepler();
      case "getstok":         return getAllStok();
      case "getpersonel":     return getAllPersonel();
      case "getzimmet":       return getAllZimmet();
      case "getsatinalma":    return getAllSatinalma();
      case "getkategorionay": return getAllKategoriOnay();
      case "getdashboard":    return getDashboardData();

      // Onay işlemleri (e-posta linkleri)
      case "approve":
      case "reject":
        return handleApproval(e);

      // Durum güncelleme
      case "updatestatus":    return updateTalepStatus(e);

      // Satınalma kaydı güncelleme
      case "updatesatinalma": return updateSatinalma(e);

      // Zimmet işlemleri
      case "teslimal":        return teslimAl(e);
      case "iadeet":          return iadeEt(e);

      default:
        return jsonResponse({ error: "Bilinmeyen işlem: " + action });
    }
  } catch (err) {
    Logger.log("doGet Error: " + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ──────────────────────────────────────────────────────────────
//  doPost — Form Gönderimi
// ──────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const postAction = payload.action || "yeniTalep";

    switch (postAction) {
      case "yeniTalep":
        return createNewTalep(payload);
      case "satinalmaKaydi":
        return createSatinalmaKaydi(payload);
      case "zimmetKaydi":
        return createZimmetKaydi(payload);
      default:
        return jsonResponse({ error: "Bilinmeyen POST işlemi." });
    }
  } catch (err) {
    Logger.log("doPost Error: " + err.toString());
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ──────────────────────────────────────────────────────────────
//  YENİ TALEP OLUŞTURMA + STOK KONTROL + YÖNLENDİRME
// ──────────────────────────────────────────────────────────────
function createNewTalep(payload) {
  const sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const talepId = generateId("TLP", sheet);
  const now = nowFormatted();

  // Stok kontrolü
  const routeResult = checkStockAndRoute(payload.urunId, payload.miktar || 1);

  // Onay anahtarları
  const mudurKey = newUuid();
  const teknikKey = newUuid();
  const direktorKey = newUuid();

  // Talep Eden bilgilerini al
  const personelInfo = getPersonelByEmail(payload.talepEdenEmail || "");

  sheet.appendRow([
    talepId,                                  // A: TalepID
    now,                                      // B: Tarih
    payload.talepEdenEmail || "",              // C: Talep_Eden_Email
    payload.urunId || "",                      // D: UrunID
    payload.urunAdi || "",                     // E: Urun_Adi
    payload.miktar || 1,                       // F: Miktar
    routeResult.ilkDurum,                      // G: Durum
    routeResult.surecTipi,                     // H: Surec_Tipi
    payload.teknikNot || "",                   // I: Teknik_Not
    mudurKey,                                  // J: Mudur_Onay_Key
    teknikKey,                                 // K: Teknik_Onay_Key
    direktorKey,                               // L: Direktor_Onay_Key
    payload.talepEdenEmail || "",              // M: Son_Islem_Yapan
    now,                                       // N: Son_Islem_Tarihi
  ]);

  // Müdüre onay e-postası gönder
  if (personelInfo && personelInfo.yoneticiEmail) {
    sendApprovalEmail(
      personelInfo.yoneticiEmail,
      {
        talepId: talepId,
        talepEden: personelInfo.isim,
        departman: personelInfo.departman,
        urunAdi: payload.urunAdi || payload.urunId,
        miktar: payload.miktar || 1,
        surecTipi: routeResult.surecTipi,
        teknikNot: payload.teknikNot || "",
      },
      "mudur",
      mudurKey
    );
  }

  return jsonResponse({
    success: true,
    talepId: talepId,
    surecTipi: routeResult.surecTipi,
    mesaj: routeResult.mesaj,
  });
}

// ──────────────────────────────────────────────────────────────
//  STOK KONTROLÜ VE YÖNLENDİRME
// ──────────────────────────────────────────────────────────────
function checkStockAndRoute(urunId, miktar) {
  if (!urunId) {
    return {
      surecTipi: "Satınalma",
      ilkDurum: "Müdür Onayı Bekliyor",
      mesaj: "Ürün stokta tanımlı değil. Satınalma sürecine yönlendirildi.",
    };
  }

  const stokSheet = getSheet(CONFIG.SHEETS.STOK);
  const data = stokSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(urunId)) {
      const mevcutStok = Number(data[i][3]) || 0;
      if (mevcutStok >= miktar) {
        return {
          surecTipi: "Stok",
          ilkDurum: "Müdür Onayı Bekliyor",
          mesaj: `Stokta mevcut (${mevcutStok} adet). Müdür onayına gönderildi.`,
        };
      } else {
        return {
          surecTipi: "Satınalma",
          ilkDurum: "Müdür Onayı Bekliyor",
          mesaj: `Stok yetersiz (${mevcutStok}/${miktar}). Satınalma sürecine yönlendirildi.`,
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

// ──────────────────────────────────────────────────────────────
//  PERSONEL BİLGİSİ
// ──────────────────────────────────────────────────────────────
function getPersonelByEmail(email) {
  if (!email) return null;
  const sheet = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase().trim() === email.toLowerCase().trim()) {
      return {
        personelId: data[i][0],
        isim: data[i][1],
        email: data[i][2],
        departman: data[i][3],
        yoneticiEmail: data[i][4],
        rol: data[i][5],
        vekilEmail: data[i][6],
      };
    }
  }
  return null;
}

// ──────────────────────────────────────────────────────────────
//  ONAY E-POSTASI GÖNDERİMİ
// ──────────────────────────────────────────────────────────────
function sendApprovalEmail(targetEmail, data, type, approvalKey) {
  const approveUrl = `${CONFIG.APP_URL}?action=approve&id=${data.talepId}&key=${approvalKey}&type=${type}`;
  const rejectUrl  = `${CONFIG.APP_URL}?action=reject&id=${data.talepId}&key=${approvalKey}&type=${type}`;

  let title = "";
  let badge = "";
  let extraInfo = "";

  switch (type) {
    case "mudur":
      title = "Departman Müdürü Onayı Gerekiyor";
      badge = "⏳ MÜDÜR ONAYI BEKLİYOR";
      break;
    case "teknik":
      title = "Teknik Birim Onayı Gerekiyor";
      badge = "🔧 TEKNİK ONAY BEKLİYOR";
      break;
    case "direktor":
      title = "Direktör Onayı Gerekiyor";
      badge = data.kritikAlim ? "🚨 KRİTİK ALIM — DİREKTÖR ONAYI" : "📋 DİREKTÖR ONAYI BEKLİYOR";
      if (data.firma) {
        extraInfo = `
          <tr><td>Tedarikçi Firma</td><td>${data.firma}</td></tr>
          <tr><td>Birim Fiyat</td><td>${data.birimFiyat}</td></tr>
          <tr><td>Toplam Tutar</td><td style="font-weight:700;color:${data.kritikAlim ? '#dc2626' : '#059669'}">${data.toplamTutar}</td></tr>
          <tr><td>Ödeme Tipi</td><td>${data.odemeTipi || "—"}</td></tr>
          <tr><td>Vade</td><td>${data.vade || "—"}</td></tr>
          <tr><td>Bütçe Kodu</td><td>${data.butceKodu || "—"}</td></tr>`;
      }
      break;
  }

  const htmlBody = `
<!DOCTYPE html>
<html lang="tr">
<head><meta charset="UTF-8">
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; color: #1a1a2e; }
  .card { background: white; border-radius: 12px; max-width: 580px; margin: 0 auto; overflow: hidden; box-shadow: 0 4px 24px rgba(0,0,0,0.08); }
  .top { background: #0f172a; padding: 28px 32px; color: white; }
  .top .logo { font-size: 0.72rem; letter-spacing: 0.12em; color: #60a5fa; text-transform: uppercase; margin-bottom: 6px; }
  .top h1 { font-size: 1.25rem; font-weight: 700; margin: 0; }
  .body { padding: 28px 32px; }
  .badge { display: inline-block; background: #fef3c7; color: #92400e; font-size: 0.72rem; font-weight: 600; padding: 5px 12px; border-radius: 5px; letter-spacing: 0.04em; margin-bottom: 20px; }
  .badge.critical { background: #fecaca; color: #991b1b; }
  table.detail { width: 100%; border-collapse: collapse; margin-bottom: 24px; }
  table.detail tr td { padding: 10px 0; border-bottom: 1px solid #f0f0f0; font-size: 0.875rem; }
  table.detail tr td:first-child { color: #6b7280; width: 150px; }
  table.detail tr td:last-child { font-weight: 500; }
  .actions { display: flex; gap: 12px; margin-top: 8px; }
  .btn-approve { flex: 1; background: #059669; color: white; text-decoration: none; text-align: center; padding: 14px; border-radius: 8px; font-weight: 700; font-size: 0.9rem; }
  .btn-reject { flex: 1; background: #dc2626; color: white; text-decoration: none; text-align: center; padding: 14px; border-radius: 8px; font-weight: 700; font-size: 0.9rem; }
  .footer { border-top: 1px solid #f0f0f0; padding: 16px 32px; font-size: 0.72rem; color: #9ca3af; text-align: center; }
  .process-badge { display: inline-block; background: #dbeafe; color: #1e40af; font-size: 0.7rem; font-weight: 600; padding: 3px 8px; border-radius: 4px; margin-left: 8px; }
</style></head>
<body>
<div class="card">
  <div class="top">
    <div class="logo">◈ Tedarik ve Zimmet Yönetim Sistemi</div>
    <h1>${title}</h1>
  </div>
  <div class="body">
    <div class="badge ${data.kritikAlim ? 'critical' : ''}">${badge} — ${data.talepId}</div>
    <table class="detail">
      <tr><td>Talep Eden</td><td>${data.talepEden || "—"} <span class="process-badge">${data.departman || ""}</span></td></tr>
      <tr><td>Ürün / Hizmet</td><td>${data.urunAdi || "—"}</td></tr>
      <tr><td>Miktar</td><td>${data.miktar || "—"}</td></tr>
      <tr><td>Süreç Tipi</td><td><span class="process-badge">${data.surecTipi || "—"}</span></td></tr>
      ${data.teknikNot ? `<tr><td>Teknik Not</td><td>${data.teknikNot}</td></tr>` : ""}
      ${extraInfo}
    </table>
    <div class="actions">
      <a href="${approveUrl}" class="btn-approve">✓ ONAYLA</a>
      <a href="${rejectUrl}" class="btn-reject">✕ REDDET</a>
    </div>
  </div>
  <div class="footer">Bu e-posta Tedarik Yönetim Sistemi tarafından otomatik gönderilmiştir.</div>
</div>
</body></html>`;

  try {
    GmailApp.sendEmail(
      targetEmail,
      `[${data.talepId}] ${title}: ${data.urunAdi}`,
      `${title} — ${data.talepEden}. Onay: ${approveUrl}`,
      { htmlBody }
    );
  } catch (err) {
    Logger.log("E-posta gönderilemedi: " + err.toString());
  }
}

// ──────────────────────────────────────────────────────────────
//  ONAY / RED İŞLEMİ (E-posta linkinden)
// ──────────────────────────────────────────────────────────────
function handleApproval(e) {
  const talepId = e.parameter.id;
  const key = e.parameter.key;
  const action = e.parameter.action; // "approve" veya "reject"
  const type = e.parameter.type || "mudur"; // "mudur", "teknik", "direktor"

  const sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const rowIndex = findRowByColumnValue(sheet, 1, talepId);

  if (!rowIndex) {
    return createResultPage("⚠️", "Talep Bulunamadı", "Bu talep ID'si geçerli değil.", "#f59e0b");
  }

  // Key doğrulama — hangi tip onay key'i kontrol edileceğini belirle
  const keyColMap = { mudur: 10, teknik: 11, direktor: 12 };
  const keyCol = keyColMap[type] || 10;
  const storedKey = sheet.getRange(rowIndex, keyCol).getValue();

  if (String(storedKey) !== String(key)) {
    return createResultPage("🔒", "Geçersiz Bağlantı", "Bu onay bağlantısı geçersiz veya süresi dolmuş.", "#ef4444");
  }

  const currentDurum = sheet.getRange(rowIndex, 7).getValue();
  const surecTipi = sheet.getRange(rowIndex, 8).getValue();

  // Reddedilme durumu
  if (action === "reject") {
    sheet.getRange(rowIndex, 7).setValue("Reddedildi");
    auditLog(sheet, rowIndex, e.parameter.approver || "E-posta Onayı");
    return createResultPage("❌", `Talep ${talepId} — Reddedildi`, "Talep reddedildi ve kayıt güncellendi.", "#ef4444");
  }

  // Onay işlemi — aşamaya göre sonraki adımı belirle
  let newDurum = "";
  let nextAction = null;

  if (type === "mudur") {
    if (surecTipi === "Stok") {
      // Stok süreci: Müdür onayından sonra Teknik Birime git
      newDurum = "Teknik Onay Bekliyor";
      nextAction = "teknik";
    } else {
      // Satınalma süreci: Müdür onayından sonra Teknik Birime git
      newDurum = "Teknik Onay Bekliyor";
      nextAction = "teknik";
    }
  } else if (type === "teknik") {
    if (surecTipi === "Stok") {
      // Stok: Teknik onaydan sonra direkt teslimata
      newDurum = "Teslim Edilebilir";
    } else {
      // Satınalma: Teknik onaydan sonra Satınalma birimine
      newDurum = "Satınalma Aşamasında";
    }
  } else if (type === "direktor") {
    newDurum = "Sipariş Edilebilir";
  }

  sheet.getRange(rowIndex, 7).setValue(newDurum);
  auditLog(sheet, rowIndex, "E-posta Onayı (" + type + ")");

  // Sonraki aşamaya e-posta gönder
  if (nextAction === "teknik") {
    sendTeknikOnayEmail(sheet, rowIndex);
  }

  const icon = "✅";
  return createResultPage(icon, `Talep ${talepId} — Onaylandı`, `Durum güncellendi: "${newDurum}"`, "#059669");
}

// ──────────────────────────────────────────────────────────────
//  TEKNİK BİRİM ONAY E-POSTASI
// ──────────────────────────────────────────────────────────────
function sendTeknikOnayEmail(talepSheet, rowIndex) {
  const row = talepSheet.getRange(rowIndex, 1, 1, 14).getValues()[0];
  const urunAdi = row[4] || row[3]; // E veya D sütunu
  const talepId = row[0];
  const teknikKey = row[10]; // K sütunu

  // Kategori-Onay tablosundan teknik onayıcıyı bul
  const katSheet = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  let teknikEmail = "";

  if (katSheet) {
    const katData = katSheet.getDataRange().getValues();
    // Ürün kategorisini stoktan bul
    const stokSheet = getSheet(CONFIG.SHEETS.STOK);
    let kategori = "IT"; // varsayılan
    if (stokSheet && row[3]) {
      const stokData = stokSheet.getDataRange().getValues();
      for (let i = 1; i < stokData.length; i++) {
        if (String(stokData[i][0]) === String(row[3])) {
          kategori = stokData[i][2] || "IT";
          break;
        }
      }
    }
    for (let i = 1; i < katData.length; i++) {
      if (String(katData[i][0]).toLowerCase() === kategori.toLowerCase()) {
        teknikEmail = katData[i][2];
        break;
      }
    }
  }

  if (!teknikEmail) {
    Logger.log("Teknik onayıcı bulunamadı, varsayılan kullanılıyor.");
    teknikEmail = CONFIG.APP_URL.includes("mcsistem") ? "it@mcsistem.com.tr" : "caglar.acer@mcsistem.com.tr";
  }

  const personelInfo = getPersonelByEmail(row[2]);

  sendApprovalEmail(
    teknikEmail,
    {
      talepId: talepId,
      talepEden: personelInfo ? personelInfo.isim : row[2],
      departman: personelInfo ? personelInfo.departman : "",
      urunAdi: urunAdi,
      miktar: row[5],
      surecTipi: row[7],
      teknikNot: row[8],
    },
    "teknik",
    teknikKey
  );
}

// ──────────────────────────────────────────────────────────────
//  SATINALMA KAYDI OLUŞTURMA
// ──────────────────────────────────────────────────────────────
function createSatinalmaKaydi(payload) {
  const sheet = getSheet(CONFIG.SHEETS.SATINALMA);
  const satId = generateId("SAT", sheet);
  const now = nowFormatted();

  const toplamTL = (payload.birimFiyat || 0) * (payload.miktar || 1);
  const kritikAlim = toplamTL >= CONFIG.CRITICAL_AMOUNT;

  sheet.appendRow([
    satId,                              // A: SatinalmaID
    payload.talepId || "",              // B: TalepID
    payload.firma || "",                // C: Firma
    payload.birimFiyat || 0,            // D: Birim_Fiyat
    payload.doviz || "TRY",             // E: Doviz
    "",                                 // F: TL_Karsiligi (formül)
    payload.vade || "",                 // G: Vade
    payload.odemeTipi || "",            // H: Odeme_Tipi
    "Beklemede",                        // I: Direktor_Onay_Durumu
    payload.butceKodu || "",            // J: Butce_Kodu
    now,                                // K: Kayit_Tarihi
    payload.islemYapan || "",           // L: Islem_Yapan
  ]);

  // TL Karşılığı formülü
  const row = sheet.getLastRow();
  const tlFormula = `=IF(E${row}="TRY", D${row}*${payload.miktar || 1}, D${row}*${payload.miktar || 1}*IFERROR(GOOGLEFINANCE("CURRENCY:"&E${row}&"TRY"),1))`;
  sheet.getRange(row, 6).setFormula(tlFormula);

  // Talep durumunu güncelle
  const talepSheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const talepRow = findRowByColumnValue(talepSheet, 1, payload.talepId);
  if (talepRow) {
    talepSheet.getRange(talepRow, 7).setValue(kritikAlim ? "Direktör Onayı Bekliyor" : "Sipariş Edilebilir");
    auditLog(talepSheet, talepRow, payload.islemYapan || "Satınalma Birimi");
  }

  // Kritik alımsa Direktöre onay gönder
  if (kritikAlim && talepRow) {
    sendDirektorOnay(talepSheet, talepRow, payload, toplamTL);
  }

  return jsonResponse({
    success: true,
    satinalmaId: satId,
    kritikAlim: kritikAlim,
    mesaj: kritikAlim
      ? "❗ Tutar 100.000 TL üzeri — Direktör onayına gönderildi."
      : "✅ Satınalma kaydı oluşturuldu. Sipariş verilebilir.",
  });
}

// ──────────────────────────────────────────────────────────────
//  DİREKTÖR ONAY E-POSTASI
// ──────────────────────────────────────────────────────────────
function sendDirektorOnay(talepSheet, talepRow, payload, toplamTL) {
  const row = talepSheet.getRange(talepRow, 1, 1, 14).getValues()[0];
  const direktorKey = row[11]; // L sütunu

  // Direktör e-postasını personel tablosundan bul
  const persSheet = getSheet(CONFIG.SHEETS.PERSONEL);
  let direktorEmail = "";
  if (persSheet) {
    const persData = persSheet.getDataRange().getValues();
    for (let i = 1; i < persData.length; i++) {
      if (String(persData[i][5]).toLowerCase() === "director") {
        direktorEmail = persData[i][2];
        break;
      }
    }
  }

  if (!direktorEmail) direktorEmail = "caglar.acer@mcsistem.com.tr";

  const sym = { TRY: "₺", USD: "$", EUR: "€" }[payload.doviz] || "₺";
  const personelInfo = getPersonelByEmail(row[2]);

  sendApprovalEmail(
    direktorEmail,
    {
      talepId: row[0],
      talepEden: personelInfo ? personelInfo.isim : row[2],
      departman: personelInfo ? personelInfo.departman : "",
      urunAdi: row[4] || row[3],
      miktar: row[5],
      surecTipi: row[7],
      teknikNot: row[8],
      firma: payload.firma,
      birimFiyat: `${sym}${Number(payload.birimFiyat).toLocaleString("tr-TR", { minimumFractionDigits: 2 })}`,
      toplamTutar: `₺${Number(toplamTL).toLocaleString("tr-TR", { minimumFractionDigits: 2 })}`,
      odemeTipi: payload.odemeTipi,
      vade: payload.vade,
      butceKodu: payload.butceKodu,
      kritikAlim: true,
    },
    "direktor",
    direktorKey
  );
}

// ──────────────────────────────────────────────────────────────
//  ZİMMET KAYDI OLUŞTURMA
// ──────────────────────────────────────────────────────────────
function createZimmetKaydi(payload) {
  const sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  const zimmetId = generateId("ZMT", sheet);
  const now = nowFormatted();

  sheet.appendRow([
    zimmetId,                           // A: ZimmetID
    payload.personelEmail || "",        // B: Personel_Email
    payload.urunId || "",               // C: UrunID
    payload.urunAdi || "",              // D: Urun_Adi
    payload.seriNo || "",               // E: SeriNo
    now,                                // F: Teslim_Tarihi
    payload.teslimEden || "",           // G: Teslim_Eden
    "Aktif",                            // H: Durum
  ]);

  // Stok güncelle — miktarı düşür
  updateStokMiktar(payload.urunId, -(payload.miktar || 1));

  // Talep durumunu güncelle
  if (payload.talepId) {
    const talepSheet = getSheet(CONFIG.SHEETS.TALEPLER);
    const talepRow = findRowByColumnValue(talepSheet, 1, payload.talepId);
    if (talepRow) {
      talepSheet.getRange(talepRow, 7).setValue("Teslim Edildi");
      auditLog(talepSheet, talepRow, payload.teslimEden || "Zimmet İşlemi");
    }
  }

  return jsonResponse({
    success: true,
    zimmetId: zimmetId,
    mesaj: "✅ Zimmet kaydı oluşturuldu ve stok güncellendi.",
  });
}

// ──────────────────────────────────────────────────────────────
//  TESLİM ALMA (Sipariş edilen ürün geldi)
// ──────────────────────────────────────────────────────────────
function teslimAl(e) {
  const talepId = e.parameter.id;
  const islemYapan = e.parameter.user || "Sistem";

  const talepSheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const talepRow = findRowByColumnValue(talepSheet, 1, talepId);

  if (!talepRow) return jsonResponse({ success: false, error: "Talep bulunamadı." });

  talepSheet.getRange(talepRow, 7).setValue("Teslim Alındı");
  auditLog(talepSheet, talepRow, islemYapan);

  // Stok güncelle — ürünü stoka ekle
  const urunId = talepSheet.getRange(talepRow, 4).getValue();
  const miktar = talepSheet.getRange(talepRow, 6).getValue();
  if (urunId) {
    updateStokMiktar(urunId, miktar);
  }

  return jsonResponse({ success: true, mesaj: "Ürün teslim alındı, stok güncellendi." });
}

// ──────────────────────────────────────────────────────────────
//  İADE İŞLEMİ
// ──────────────────────────────────────────────────────────────
function iadeEt(e) {
  const zimmetId = e.parameter.id;
  const islemYapan = e.parameter.user || "Sistem";

  const zimmetSheet = getSheet(CONFIG.SHEETS.ZIMMET);
  const zimmetRow = findRowByColumnValue(zimmetSheet, 1, zimmetId);

  if (!zimmetRow) return jsonResponse({ success: false, error: "Zimmet bulunamadı." });

  zimmetSheet.getRange(zimmetRow, 8).setValue("İade Edildi");

  // Stok güncelle — iade edilen ürünü stoka ekle
  const urunId = zimmetSheet.getRange(zimmetRow, 3).getValue();
  if (urunId) {
    updateStokMiktar(urunId, 1);
  }

  return jsonResponse({ success: true, mesaj: "Ürün iade edildi, stok güncellendi." });
}

// ──────────────────────────────────────────────────────────────
//  STOK MİKTAR GÜNCELLEME
// ──────────────────────────────────────────────────────────────
function updateStokMiktar(urunId, delta) {
  if (!urunId) return;
  const stokSheet = getSheet(CONFIG.SHEETS.STOK);
  if (!stokSheet) return;
  const data = stokSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(urunId)) {
      const current = Number(data[i][3]) || 0;
      const newVal = Math.max(0, current + delta);
      stokSheet.getRange(i + 1, 4).setValue(newVal);

      // Kritik seviye uyarısı
      const kritik = Number(data[i][4]) || 0;
      if (newVal <= kritik && newVal > 0) {
        Logger.log(`⚠️ Stok uyarısı: ${data[i][1]} — Mevcut: ${newVal}, Kritik: ${kritik}`);
      }
      return;
    }
  }
}

// ──────────────────────────────────────────────────────────────
//  DURUM GÜNCELLEME (Dashboard butonlarından)
// ──────────────────────────────────────────────────────────────
function updateTalepStatus(e) {
  const talepId = e.parameter.id;
  const status = e.parameter.status;
  const user = e.parameter.user || "Dashboard";

  const sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const rowIndex = findRowByColumnValue(sheet, 1, talepId);
  if (!rowIndex) return jsonResponse({ success: false, error: "Kayıt bulunamadı." });

  sheet.getRange(rowIndex, 7).setValue(status);
  auditLog(sheet, rowIndex, user);

  return jsonResponse({ success: true, mesaj: `Durum güncellendi: ${status}` });
}

// ──────────────────────────────────────────────────────────────
//  SATINALMA GÜNCELLEME
// ──────────────────────────────────────────────────────────────
function updateSatinalma(e) {
  const satId = e.parameter.id;
  const status = e.parameter.status || "";

  const sheet = getSheet(CONFIG.SHEETS.SATINALMA);
  const rowIndex = findRowByColumnValue(sheet, 1, satId);
  if (!rowIndex) return jsonResponse({ success: false, error: "Satınalma kaydı bulunamadı." });

  if (status) {
    sheet.getRange(rowIndex, 9).setValue(status); // I sütunu: Direktor_Onay_Durumu
  }

  return jsonResponse({ success: true });
}

// ──────────────────────────────────────────────────────────────
//  VERİ ÇEKME FONKSİYONLARI
// ──────────────────────────────────────────────────────────────
function getAllTalepler() {
  const sheet = getSheet(CONFIG.SHEETS.TALEPLER);
  if (!sheet) return jsonResponse({ rows: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ rows: [] });

  const rows = data.slice(1).map(r => ({
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
  })).reverse();

  return jsonResponse({ rows });
}

function getAllStok() {
  const sheet = getSheet(CONFIG.SHEETS.STOK);
  if (!sheet) return jsonResponse({ rows: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ rows: [] });

  const rows = data.slice(1).map(r => ({
    urunId:       r[0],
    urunAdi:      r[1],
    kategori:     r[2],
    mevcutStok:   r[3],
    kritikSeviye: r[4],
    birim:        r[5],
  }));

  return jsonResponse({ rows });
}

function getAllPersonel() {
  const sheet = getSheet(CONFIG.SHEETS.PERSONEL);
  if (!sheet) return jsonResponse({ rows: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ rows: [] });

  const rows = data.slice(1).map(r => ({
    personelId:    r[0],
    isim:          r[1],
    email:         r[2],
    departman:     r[3],
    yoneticiEmail: r[4],
    rol:           r[5],
    vekilEmail:    r[6],
  }));

  return jsonResponse({ rows });
}

function getAllZimmet() {
  const sheet = getSheet(CONFIG.SHEETS.ZIMMET);
  if (!sheet) return jsonResponse({ rows: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ rows: [] });

  const rows = data.slice(1).map(r => ({
    zimmetId:       r[0],
    personelEmail:  r[1],
    urunId:         r[2],
    urunAdi:        r[3],
    seriNo:         r[4],
    teslimTarihi:   r[5],
    teslimEden:     r[6],
    durum:          r[7],
  })).reverse();

  return jsonResponse({ rows });
}

function getAllSatinalma() {
  const sheet = getSheet(CONFIG.SHEETS.SATINALMA);
  if (!sheet) return jsonResponse({ rows: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ rows: [] });

  const rows = data.slice(1).map(r => ({
    satinalmaId:       r[0],
    talepId:           r[1],
    firma:             r[2],
    birimFiyat:        r[3],
    doviz:             r[4],
    tlKarsiligi:       r[5],
    vade:              r[6],
    odemeTipi:         r[7],
    direktorOnayDurumu: r[8],
    butceKodu:         r[9],
    kayitTarihi:       r[10],
    islemYapan:        r[11],
  })).reverse();

  return jsonResponse({ rows });
}

function getAllKategoriOnay() {
  const sheet = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  if (!sheet) return jsonResponse({ rows: [] });
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ rows: [] });

  const rows = data.slice(1).map(r => ({
    kategori:         r[0],
    teknikBirim:      r[1],
    teknikOnayciEmail: r[2],
  }));

  return jsonResponse({ rows });
}

// ──────────────────────────────────────────────────────────────
//  DASHBOARD VERİSİ (Özet istatistikler)
// ──────────────────────────────────────────────────────────────
function getDashboardData() {
  const talepSheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const stokSheet = getSheet(CONFIG.SHEETS.STOK);
  const zimmetSheet = getSheet(CONFIG.SHEETS.ZIMMET);

  let talepler = [];
  if (talepSheet && talepSheet.getLastRow() > 1) {
    talepler = talepSheet.getDataRange().getValues().slice(1).map(r => ({
      talepId: r[0], tarih: r[1], talepEdenEmail: r[2],
      urunId: r[3], urunAdi: r[4], miktar: r[5],
      durum: r[6], surecTipi: r[7], teknikNot: r[8],
      sonIslemYapan: r[12], sonIslemTarihi: r[13],
    })).reverse();
  }

  let stoklar = [];
  if (stokSheet && stokSheet.getLastRow() > 1) {
    stoklar = stokSheet.getDataRange().getValues().slice(1).map(r => ({
      urunId: r[0], urunAdi: r[1], kategori: r[2],
      mevcutStok: r[3], kritikSeviye: r[4], birim: r[5],
    }));
  }

  let zimmetler = [];
  if (zimmetSheet && zimmetSheet.getLastRow() > 1) {
    zimmetler = zimmetSheet.getDataRange().getValues().slice(1).map(r => ({
      zimmetId: r[0], personelEmail: r[1], urunId: r[2],
      urunAdi: r[3], seriNo: r[4], teslimTarihi: r[5],
      teslimEden: r[6], durum: r[7],
    })).reverse();
  }

  // İstatistik hesapla
  const stats = {
    toplam: talepler.length,
    mudurOnay: talepler.filter(t => t.durum === "Müdür Onayı Bekliyor").length,
    teknikOnay: talepler.filter(t => t.durum === "Teknik Onay Bekliyor").length,
    satinalma: talepler.filter(t => t.durum === "Satınalma Aşamasında").length,
    direktorOnay: talepler.filter(t => t.durum === "Direktör Onayı Bekliyor").length,
    siparisEdildi: talepler.filter(t => ["Sipariş Edildi", "Sipariş Edilebilir"].includes(t.durum)).length,
    teslimEdildi: talepler.filter(t => ["Teslim Edildi", "Teslim Alındı"].includes(t.durum)).length,
    reddedildi: talepler.filter(t => t.durum === "Reddedildi").length,
    kritikStok: stoklar.filter(s => Number(s.mevcutStok) <= Number(s.kritikSeviye)).length,
    aktifZimmet: zimmetler.filter(z => z.durum === "Aktif").length,
  };

  return jsonResponse({ talepler, stoklar, zimmetler, stats });
}

// ──────────────────────────────────────────────────────────────
//  AUDIT LOG
// ──────────────────────────────────────────────────────────────
function auditLog(sheet, rowIndex, user) {
  const now = nowFormatted();
  sheet.getRange(rowIndex, 13).setValue(user);     // M: Son_Islem_Yapan
  sheet.getRange(rowIndex, 14).setValue(now);       // N: Son_Islem_Tarihi
}

// ──────────────────────────────────────────────────────────────
//  SONUÇ SAYFASI (Onay/Red sonrası gösterilir)
// ──────────────────────────────────────────────────────────────
function createResultPage(icon, title, message, color) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html><head><meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
      body { font-family: 'Inter', sans-serif; display: flex; align-items: center; justify-content: center;
             min-height: 100vh; margin: 0; background: #f8fafc; }
      .box { background: white; border-radius: 16px; padding: 56px 48px; text-align: center;
             max-width: 440px; box-shadow: 0 4px 24px rgba(0,0,0,0.08); border: 1px solid #e2e8f0; }
      .icon { font-size: 3.5rem; margin-bottom: 20px; display: block; }
      h2 { color: ${color}; font-size: 1.3rem; margin-bottom: 12px; font-weight: 700; }
      p { color: #64748b; font-size: 0.9rem; line-height: 1.6; }
    </style></head>
    <body><div class="box">
      <span class="icon">${icon}</span>
      <h2>${title}</h2>
      <p>${message}</p>
      <p style="margin-top:20px;font-size:0.78rem;color:#94a3b8;">Bu sekmeyi kapatabilirsiniz.</p>
    </div></body></html>
  `);
}

// ══════════════════════════════════════════════════════════════
//  İLK KURULUM — Bir kez çalıştırın
// ══════════════════════════════════════════════════════════════
function setupAllSheets() {
  const ss = getSpreadsheet();

  // ── REF_Personel ──
  createOrResetSheet(ss, CONFIG.SHEETS.PERSONEL, [
    "PersonelID", "Isim", "Email", "Departman", "Yonetici_Email", "Rol", "Vekil_Email"
  ]);

  // ── REF_Stok ──
  createOrResetSheet(ss, CONFIG.SHEETS.STOK, [
    "UrunID", "Urun_Adi", "Kategori", "Mevcut_Stok", "Kritik_Seviye", "Birim"
  ]);

  // ── REF_Kategori_Onay ──
  createOrResetSheet(ss, CONFIG.SHEETS.KATEGORI_ONAY, [
    "Kategori", "Teknik_Birim", "Teknik_Onayci_Email"
  ]);

  // ── TRX_Talepler ──
  createOrResetSheet(ss, CONFIG.SHEETS.TALEPLER, [
    "TalepID", "Tarih", "Talep_Eden_Email", "UrunID", "Urun_Adi", "Miktar",
    "Durum", "Surec_Tipi", "Teknik_Not",
    "Mudur_Onay_Key", "Teknik_Onay_Key", "Direktor_Onay_Key",
    "Son_Islem_Yapan", "Son_Islem_Tarihi"
  ]);

  // ── TRX_Satinalma ──
  createOrResetSheet(ss, CONFIG.SHEETS.SATINALMA, [
    "SatinalmaID", "TalepID", "Firma", "Birim_Fiyat", "Doviz", "TL_Karsiligi",
    "Vade", "Odeme_Tipi", "Direktor_Onay_Durumu", "Butce_Kodu",
    "Kayit_Tarihi", "Islem_Yapan"
  ]);

  // ── TRX_Zimmet ──
  createOrResetSheet(ss, CONFIG.SHEETS.ZIMMET, [
    "ZimmetID", "Personel_Email", "UrunID", "Urun_Adi", "SeriNo",
    "Teslim_Tarihi", "Teslim_Eden", "Durum"
  ]);

  Logger.log("✅ Tüm sayfalar başarıyla oluşturuldu.");
}

function createOrResetSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    // Sadece başlık ekle, mevcut veriyi koru
    if (sheet.getLastRow() === 0) {
      // Boşsa başlık yaz
    }
  }

  // Başlıkları yaz/güncelle
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Başlık stili
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground("#0f172a");
  hRange.setFontColor("#60a5fa");
  hRange.setFontWeight("bold");
  hRange.setFontFamily("Courier New");
  hRange.setFontSize(10);

  // İlk satırı dondur
  sheet.setFrozenRows(1);

  // Sütun genişlikleri
  for (let i = 1; i <= headers.length; i++) {
    sheet.setColumnWidth(i, Math.max(120, headers[i - 1].length * 12));
  }

  // Onay key sütunlarını gizle (TRX_Talepler)
  if (name === CONFIG.SHEETS.TALEPLER) {
    sheet.hideColumns(10); // Mudur_Onay_Key
    sheet.hideColumns(11); // Teknik_Onay_Key
    sheet.hideColumns(12); // Direktor_Onay_Key
  }
}

// ══════════════════════════════════════════════════════════════
//  ÖRNEK VERİ EKLEME — İlk test için
// ══════════════════════════════════════════════════════════════
function seedSampleData() {
  // Personel
  const persSheet = getSheet(CONFIG.SHEETS.PERSONEL);
  if (persSheet.getLastRow() < 2) {
    persSheet.appendRow([1, "Çağlar Acer", "caglar.acer@mcsistem.com.tr", "Bilgi Teknolojileri", "", "Director", ""]);
    persSheet.appendRow([2, "Ahmet Yılmaz", "ahmet.yilmaz@mcsistem.com.tr", "Üretim", "caglar.acer@mcsistem.com.tr", "User", ""]);
    persSheet.appendRow([3, "Mehmet Kaya", "mehmet.kaya@mcsistem.com.tr", "Lojistik", "caglar.acer@mcsistem.com.tr", "User", ""]);
    persSheet.appendRow([4, "Ayşe Demir", "ayse.demir@mcsistem.com.tr", "İnsan Kaynakları", "caglar.acer@mcsistem.com.tr", "Manager", ""]);
    persSheet.appendRow([5, "Fatma Şahin", "fatma.sahin@mcsistem.com.tr", "Satınalma", "caglar.acer@mcsistem.com.tr", "Procurement", ""]);
    persSheet.appendRow([6, "Ali Öztürk", "ali.ozturk@mcsistem.com.tr", "Bilgi Teknolojileri", "caglar.acer@mcsistem.com.tr", "IT", ""]);
  }

  // Stok
  const stokSheet = getSheet(CONFIG.SHEETS.STOK);
  if (stokSheet.getLastRow() < 2) {
    stokSheet.appendRow(["STK-001", "Laptop (Dell Latitude 5540)", "IT", 5, 2, "Adet"]);
    stokSheet.appendRow(["STK-002", "Monitor (24\" LG)", "IT", 8, 3, "Adet"]);
    stokSheet.appendRow(["STK-003", "Klavye + Mouse Set", "IT", 15, 5, "Adet"]);
    stokSheet.appendRow(["STK-004", "A4 Kağıt (500'lü)", "İdari", 120, 30, "Paket"]);
    stokSheet.appendRow(["STK-005", "Toner (HP LaserJet)", "İdari", 6, 2, "Adet"]);
    stokSheet.appendRow(["STK-006", "Masa Üstü Telefon", "İdari", 3, 1, "Adet"]);
    stokSheet.appendRow(["STK-007", "Güvenlik Kamerası", "IT", 2, 1, "Adet"]);
    stokSheet.appendRow(["STK-008", "Ofis Koltuğu", "İdari", 4, 2, "Adet"]);
  }

  // Kategori Onay
  const katSheet = getSheet(CONFIG.SHEETS.KATEGORI_ONAY);
  if (katSheet.getLastRow() < 2) {
    katSheet.appendRow(["IT", "Bilgi Teknolojileri", "ali.ozturk@mcsistem.com.tr"]);
    katSheet.appendRow(["İdari", "İdari İşler", "ayse.demir@mcsistem.com.tr"]);
    katSheet.appendRow(["Üretim", "Üretim Mühendisliği", "caglar.acer@mcsistem.com.tr"]);
  }

  Logger.log("✅ Örnek veriler eklendi.");
}

// ══════════════════════════════════════════════════════════════
//  ESKİ VERİ MİGRASYONU — Mevcut "Talepler" sayfasından
// ══════════════════════════════════════════════════════════════
function migrateOldData() {
  const ss = getSpreadsheet();
  const oldSheet = ss.getSheetByName("Talepler");
  if (!oldSheet) {
    Logger.log("Eski 'Talepler' sayfası bulunamadı.");
    return;
  }

  const newSheet = getSheet(CONFIG.SHEETS.TALEPLER);
  const data = oldSheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log("Eski sayfada veri yok.");
    return;
  }

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const talepId = `TLP-${String(r[0]).padStart(5, "0")}`;

    // Eskiden zaten migrate edilmişse atla
    if (findRowByColumnValue(newSheet, 1, talepId)) continue;

    newSheet.appendRow([
      talepId,           // TalepID
      r[1],              // Tarih
      "",                // Talep_Eden_Email (eski sistemde yoktu)
      "",                // UrunID
      r[3],              // Urun_Adi
      r[5],              // Miktar
      r[9] || "Beklemede", // Durum
      "Satınalma",       // Surec_Tipi
      r[4] || "",        // Teknik_Not (Açıklama)
      r[10] || "",       // Mudur_Onay_Key
      "",                // Teknik_Onay_Key
      "",                // Direktor_Onay_Key
      "Migrasyon",       // Son_Islem_Yapan
      nowFormatted(),    // Son_Islem_Tarihi
    ]);
    count++;
  }

  Logger.log(`✅ ${count} kayıt migrate edildi.`);
}
