const INV_ACCESS_PIN = "";
const INV_SHEET_NAME = "Sheet1";

function invCheckPin_(pin) {
  if (!INV_ACCESS_PIN) return true;
  return String(pin || "") === String(INV_ACCESS_PIN);
}

function invJsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const pin = (e && e.parameter && e.parameter.pin) || "";
  if (!invCheckPin_(pin)) {
    return invJsonOut_({ ok: false, error: "Invalid PIN" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INV_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return invJsonOut_({ ok: true, products: [] });

  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const products = [];

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const code = String(r[1] || "").trim();
    if (!code) continue;

    const rawStatus = String(r[11] || "").trim();
    let normStatus = rawStatus;
    const low = rawStatus.toLowerCase();
    if (["not avilable", "not available", "unavailable", "no"].includes(low)) {
      normStatus = "Not Available";
    } else if (["available", "avilable", "yes"].includes(low)) {
      normStatus = "Available";
    }

    products.push({
      row: i + 2,
      sn: String(r[0] || "").trim(),
      code: code,
      cost: String(r[2] || "").trim(),
      daraz_sp: String(r[3] || "").trim(),
      net_profit: String(r[7] || "").trim(),
      group: String(r[8] || "").trim(),
      name: String(r[9] || "").trim(),
      price: String(r[10] || "").trim(),
      status: normStatus,
      in_drive: r[12] === true || String(r[12]).toUpperCase() === "TRUE",
    });
  }

  return invJsonOut_({ ok: true, products: products });
}

function doPost(e) {
  const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  if (!invCheckPin_(payload.pin)) {
    return invJsonOut_({ ok: false, error: "Invalid PIN" });
  }

  const row = Number(payload.row);
  if (!row || row < 2) {
    return invJsonOut_({ ok: false, error: "Invalid row" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INV_SHEET_NAME);
  if (payload.action === "update_status") {
    sheet.getRange(row, 12).setValue(payload.status || "");
    return invJsonOut_({ ok: true, row: row, status: payload.status || "" });
  }
  if (payload.action === "update_in_drive") {
    sheet.getRange(row, 13).setValue(Boolean(payload.in_drive));
    return invJsonOut_({ ok: true, row: row, in_drive: Boolean(payload.in_drive) });
  }

  return invJsonOut_({ ok: false, error: "Unknown action" });
}
