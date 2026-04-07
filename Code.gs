// ============================================================
//  SATIN ALMA ONAY SİSTEMİ — Google Apps Script Backend
//  Dosyayı Apps Script editörüne yapıştırın, kaydedin ve
//  "Web App olarak dağıt" adımlarını takip edin.
// ============================================================

// ──────────────────────────────────────────────────────────────
//  AYARLAR — Buraya kendi değerlerinizi girin
// ──────────────────────────────────────────────────────────────
const CONFIG = {
  SPREADSHEET_ID: "1HI5BxkFKK23jH_yZJxxZD9XJ8hjdxmd5LY65Ca9Bnns",
  SHEET_NAME: "Talepler",
  MANAGER_EMAIL: "caglar.acer@mcsistem.com.tr",
  APP_URL: "https://script.google.com/macros/s/AKfycbz9EpgfR15uESBCJMDaovg6KWJmFiw2p11kffWxKLwHxjuIClb9Y-J0bCGBV9lAv0u2/exec",
};

// ──────────────────────────────────────────────────────────────
//  CORS Headers (GitHub Pages erişimi için)
// ──────────────────────────────────────────────────────────────
function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function jsonResponse(data, code = 200) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ──────────────────────────────────────────────────────────────
//  doGet — Dashboard verisi çekme + Onay/Red işlemi
// ──────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || "";

  // 1) Onay mekanizması: e-postadaki bağlantıdan gelir
  if (action === "approve" || action === "reject") {
    return handleApproval(e);
  }

  // 2) Tüm satırları getir (Dashboard için)
  if (action === "getAll") {
    return getAllRows();
  }

  // 3) Durum güncelleme (Satınalma dept. aksiyonları)
  if (action === "updateStatus") {
    return updateStatus(e);
  }

  return jsonResponse({ error: "Bilinmeyen işlem." }, 400);
}

// ──────────────────────────────────────────────────────────────
//  doPost — Talep formu veri gönderimi
// ──────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    // Otomatik ID
    const lastRow = sheet.getLastRow();
    const newId = lastRow < 2 ? 1 : sheet.getRange(lastRow, 1).getValue() + 1;

    // Benzersiz onay anahtarı
    const approvalKey = Utilities.getUuid().replace(/-/g, "").substring(0, 24);

    // Tarih formatı
    const now = Utilities.formatDate(new Date(), "Europe/Istanbul", "dd.MM.yyyy HH:mm");

    // Satırı ekle
    sheet.appendRow([
      newId,                          // A: ID
      now,                            // B: Tarih
      payload.talep_eden || "",       // C: Talep_Eden
      payload.urun_adi || "",         // D: Urun_Adi
      payload.aciklama || "",         // E: Aciklama
      payload.miktar || 0,            // F: Miktar
      payload.birim_fiyat || 0,       // G: Birim_Fiyat
      payload.para_birimi || "TRY",   // H: Para_Birimi
      "",                             // I: TL_Karsiligi (formül ile doldurulacak)
      "Beklemede",                    // J: Durum
      approvalKey,                    // K: Yonetici_Onay_Key
    ]);

    // TL karşılığı formülü — I sütununa ekle
    const row = sheet.getLastRow();
    const tlFormula = `=IF(H${row}="TRY", F${row}*G${row}, F${row}*G${row}*IFERROR(GOOGLEFINANCE("CURRENCY:"&H${row}&"TRY"),1))`;
    sheet.getRange(row, 9).setFormula(tlFormula);

    // Yöneticiye bildirim e-postası gönder
    sendNotificationEmail(newId, payload, approvalKey);

    return jsonResponse({ success: true, id: newId });

  } catch (err) {
    Logger.log(err);
    return jsonResponse({ success: false, error: err.toString() }, 500);
  }
}

// ──────────────────────────────────────────────────────────────
//  Yönetici Bildirim E-postası
// ──────────────────────────────────────────────────────────────
function sendNotificationEmail(id, payload, key) {
  const approveUrl = `${CONFIG.APP_URL}?action=approve&id=${id}&key=${key}`;
  const rejectUrl  = `${CONFIG.APP_URL}?action=reject&id=${id}&key=${key}`;

  const toplam = (payload.miktar || 0) * (payload.birim_fiyat || 0);
  const sym = { TRY: "₺", USD: "$", EUR: "€" }[payload.para_birimi] || "";

  const htmlBody = `
<!DOCTYPE html>
<html lang="tr">
<head><meta charset="UTF-8">
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 20px; color: #1a1a2e; }
  .card { background: white; border-radius: 12px; max-width: 560px; margin: 0 auto; overflow: hidden; box-shadow: 0 4px 24px rgba(0,0,0,0.08); }
  .top { background: #111318; padding: 28px 32px; color: white; }
  .top .logo { font-size: 0.75rem; letter-spacing: 0.12em; color: #e8c547; text-transform: uppercase; margin-bottom: 6px; }
  .top h1 { font-size: 1.3rem; font-weight: 700; margin: 0; }
  .body { padding: 28px 32px; }
  .badge { display: inline-block; background: #fef3c7; color: #92400e; font-size: 0.72rem; font-weight: 600; padding: 4px 10px; border-radius: 5px; letter-spacing: 0.04em; margin-bottom: 20px; }
  table.detail { width: 100%; border-collapse: collapse; margin-bottom: 24px; }
  table.detail tr td { padding: 9px 0; border-bottom: 1px solid #f0f0f0; font-size: 0.875rem; }
  table.detail tr td:first-child { color: #6b7280; width: 140px; }
  table.detail tr td:last-child { font-weight: 500; }
  .total-row td { border-bottom: none !important; font-size: 1rem !important; }
  .total-row td:last-child { color: #e8c547; font-size: 1.1rem !important; font-weight: 700 !important; }
  .actions { display: flex; gap: 12px; margin-top: 8px; }
  .btn-approve { flex: 1; background: #059669; color: white; text-decoration: none; text-align: center; padding: 13px; border-radius: 8px; font-weight: 700; font-size: 0.9rem; }
  .btn-reject { flex: 1; background: #dc2626; color: white; text-decoration: none; text-align: center; padding: 13px; border-radius: 8px; font-weight: 700; font-size: 0.9rem; }
  .footer { border-top: 1px solid #f0f0f0; padding: 16px 32px; font-size: 0.75rem; color: #9ca3af; text-align: center; }
</style></head>
<body>
<div class="card">
  <div class="top">
    <div class="logo">◈ Tedarik Yönetim Sistemi</div>
    <h1>Yeni Satınalma Talebi</h1>
  </div>
  <div class="body">
    <div class="badge">⏳ ONAY BEKLİYOR — #${id}</div>
    <table class="detail">
      <tr><td>Talep Eden</td><td>${payload.talep_eden || "—"}</td></tr>
      <tr><td>Ürün / Hizmet</td><td>${payload.urun_adi || "—"}</td></tr>
      <tr><td>Açıklama</td><td>${payload.aciklama || "—"}</td></tr>
      <tr><td>Miktar</td><td>${payload.miktar}</td></tr>
      <tr><td>Birim Fiyat</td><td>${sym}${Number(payload.birim_fiyat).toLocaleString("tr-TR", {minimumFractionDigits:2})}</td></tr>
      <tr class="total-row"><td>Toplam Tutar</td><td>${sym}${Number(toplam).toLocaleString("tr-TR", {minimumFractionDigits:2})}</td></tr>
    </table>
    <div class="actions">
      <a href="${approveUrl}" class="btn-approve">✓ ONAYLA</a>
      <a href="${rejectUrl}" class="btn-reject">✕ REDDET</a>
    </div>
  </div>
  <div class="footer">Bu e-posta otomatik olarak gönderilmiştir. Dashboard: ${CONFIG.APP_URL}</div>
</div>
</body></html>`;

  GmailApp.sendEmail(
    CONFIG.MANAGER_EMAIL,
    `[Satınalma #${id}] Onay Bekleyen Talep: ${payload.urun_adi}`,
    `Yeni talep: ${payload.urun_adi} — ${payload.talep_eden}. Onay için: ${approveUrl}`,
    { htmlBody }
  );
}

// ──────────────────────────────────────────────────────────────
//  Onay / Red işlemi (e-posta linkinden)
// ──────────────────────────────────────────────────────────────
function handleApproval(e) {
  const id = parseInt(e.parameter.id);
  const key = e.parameter.key;
  const action = e.parameter.action; // "approve" veya "reject"
  const sheet = getSheet();

  const rowIndex = findRowById(sheet, id);
  if (!rowIndex) {
    return HtmlService.createHtmlOutput("<h2>⚠️ Talep bulunamadı.</h2>");
  }

  const storedKey = sheet.getRange(rowIndex, 11).getValue(); // K sütunu
  if (storedKey !== key) {
    return HtmlService.createHtmlOutput("<h2>🔒 Geçersiz ya da süresi dolmuş bağlantı.</h2>");
  }

  const currentStatus = sheet.getRange(rowIndex, 10).getValue();
  if (currentStatus !== "Beklemede") {
    return HtmlService.createHtmlOutput(`<h2>Bu talep zaten işleme alınmış. Mevcut durum: <b>${currentStatus}</b></h2>`);
  }

  const newStatus = action === "approve" ? "Onaylandı" : "Reddedildi";
  sheet.getRange(rowIndex, 10).setValue(newStatus);

  const color = action === "approve" ? "#059669" : "#dc2626";
  const icon  = action === "approve" ? "✅" : "❌";
  return HtmlService.createHtmlOutput(`
    <html><head><meta charset="UTF-8">
    <style>body{font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f4f6f9;}
    .box{background:white;border-radius:12px;padding:48px 40px;text-align:center;max-width:420px;box-shadow:0 4px 24px rgba(0,0,0,0.1);}
    h2{color:${color};font-size:1.4rem;margin-bottom:12px;} p{color:#6b7280;}</style></head>
    <body><div class="box"><div style="font-size:3rem;margin-bottom:16px;">${icon}</div>
    <h2>Talep #${id} — ${newStatus}</h2>
    <p>İşlem başarıyla kaydedildi. Bu sekmeyi kapatabilirsiniz.</p></div></body></html>`);
}

// ──────────────────────────────────────────────────────────────
//  Tüm satırları getir (Dashboard)
// ──────────────────────────────────────────────────────────────
function getAllRows() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ rows: [] });

  const rows = data.slice(1).map(r => ({
    id:            r[0],
    tarih:         r[1],
    talep_eden:    r[2],
    urun_adi:      r[3],
    aciklama:      r[4],
    miktar:        r[5],
    birim_fiyat:   r[6],
    para_birimi:   r[7],
    tl_karsiligi:  r[8],
    durum:         r[9],
  })).reverse(); // En yeniler üste

  return jsonResponse({ rows });
}

// ──────────────────────────────────────────────────────────────
//  Durum güncelle (Dashboard butonları)
// ──────────────────────────────────────────────────────────────
function updateStatus(e) {
  const id = parseInt(e.parameter.id);
  const status = e.parameter.status;
  const sheet = getSheet();
  const rowIndex = findRowById(sheet, id);
  if (!rowIndex) return jsonResponse({ success: false, error: "Kayıt bulunamadı." });
  sheet.getRange(rowIndex, 10).setValue(status);
  return jsonResponse({ success: true });
}

// ──────────────────────────────────────────────────────────────
//  Yardımcı fonksiyonlar
// ──────────────────────────────────────────────────────────────
function getSheet() {
  return SpreadsheetApp
    .openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);
}

function findRowById(sheet, id) {
  const idCol = sheet.getRange("A:A").getValues();
  for (let i = 1; i < idCol.length; i++) {
    if (idCol[i][0] == id) return i + 1; // 1-indexed
  }
  return null;
}

// ──────────────────────────────────────────────────────────────
//  Google Sheet İlk Kurulum (Manuel çalıştır)
//  Apps Script Editöründe bu fonksiyonu bir kez çalıştırın.
// ──────────────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

  const headers = [
    "ID", "Tarih", "Talep_Eden", "Urun_Adi", "Aciklama",
    "Miktar", "Birim_Fiyat", "Para_Birimi", "TL_Karsiligi",
    "Durum", "Yonetici_Onay_Key"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Başlık stili
  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground("#111318");
  hRange.setFontColor("#e8c547");
  hRange.setFontWeight("bold");
  hRange.setFontFamily("Courier New");

  // Kolon genişlikleri
  sheet.setColumnWidth(1, 50);   // ID
  sheet.setColumnWidth(2, 130);  // Tarih
  sheet.setColumnWidth(3, 200);  // Talep Eden
  sheet.setColumnWidth(4, 220);  // Ürün
  sheet.setColumnWidth(5, 200);  // Açıklama
  sheet.setColumnWidth(6, 70);   // Miktar
  sheet.setColumnWidth(7, 100);  // Birim Fiyat
  sheet.setColumnWidth(8, 100);  // Para Birimi
  sheet.setColumnWidth(9, 130);  // TL Karşılığı
  sheet.setColumnWidth(10, 120); // Durum
  sheet.setColumnWidth(11, 200); // Onay Key

  // Key sütununu gizle
  sheet.hideColumns(11);

  Logger.log("✅ Sheet kurulumu tamamlandı.");
}
