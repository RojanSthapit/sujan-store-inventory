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

  const values = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
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
      image_url: String(r[13] || "").trim(),
    });
  }

  return invJsonOut_({ ok: true, products: products });
}

function doPost(e) {
  const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  if (!invCheckPin_(payload.pin)) {
    return invJsonOut_({ ok: false, error: "Invalid PIN" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INV_SHEET_NAME);
  const action = payload.action;

  if (action === "update_status") {
    const row = Number(payload.row);
    if (!row || row < 2) return invJsonOut_({ ok: false, error: "Invalid row" });
    sheet.getRange(row, 12).setValue(payload.status || "");
    return invJsonOut_({ ok: true, row: row, status: payload.status || "" });
  }

  if (action === "update_in_drive") {
    const row = Number(payload.row);
    if (!row || row < 2) return invJsonOut_({ ok: false, error: "Invalid row" });
    sheet.getRange(row, 13).setValue(Boolean(payload.in_drive));
    return invJsonOut_({ ok: true, row: row, in_drive: Boolean(payload.in_drive) });
  }

  if (action === "add_product") {
    const lastRow = Math.max(1, sheet.getLastRow());
    const colB = lastRow > 1 ? sheet.getRange(2, 2, lastRow - 1, 1).getValues() : [];
    const colA = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
    let targetRow = 2;
    let maxSn = 0;
    let maxCodeNum = 0;
    let codePrefix = "SS25-";
    for (let i = 0; i < colB.length; i++) {
      const codeStr = String(colB[i][0] || "").trim();
      if (codeStr) {
        targetRow = i + 3;
        const m = codeStr.match(/^(.*?)(\d+)$/);
        if (m) {
          const n = parseInt(m[2], 10);
          if (n > maxCodeNum) {
            maxCodeNum = n;
            codePrefix = m[1];
          }
        }
      }
      const snVal = Number(colA[i][0]);
      if (!isNaN(snVal) && snVal > maxSn) maxSn = snVal;
    }

    const pdtCode = String(payload.code || "").trim() || (codePrefix + String((maxCodeNum > 0 ? maxCodeNum + 2 : 1)).padStart(3, "0"));
    const cost = Number(payload.cost) || 0;
    const darazSp = Number(payload.daraz_sp) || 0;
    const commRate = payload.daraz_comm !== undefined && payload.daraz_comm !== "" ? Number(payload.daraz_comm) : 0.3;

    sheet.getRange(targetRow, 1, 1, 14).setValues([[
      maxSn + 1,
      pdtCode,
      cost,
      darazSp,
      commRate,
      `=D${targetRow}*E${targetRow}`,
      `=D${targetRow}-F${targetRow}`,
      `=G${targetRow}-C${targetRow}`,
      String(payload.group || "").trim(),
      String(payload.name || "").trim(),
      "",
      String(payload.status || "").trim(),
      Boolean(payload.in_drive),
      String(payload.image_url || "").trim(),
    ]]);
    sheet.getRange(targetRow, 13).insertCheckboxes();
    return invJsonOut_({ ok: true, row: targetRow });
  }

  if (action === "edit_product") {
    const row = Number(payload.row);
    if (!row || row < 2) return invJsonOut_({ ok: false, error: "Invalid row" });
    sheet.getRange(row, 2).setValue(String(payload.code || "").trim());
    sheet.getRange(row, 3).setValue(Number(payload.cost) || 0);
    sheet.getRange(row, 4).setValue(Number(payload.daraz_sp) || 0);
    sheet.getRange(row, 9).setValue(String(payload.group || "").trim());
    sheet.getRange(row, 10).setValue(String(payload.name || "").trim());
    sheet.getRange(row, 12).setValue(String(payload.status || "").trim());
    sheet.getRange(row, 13).setValue(Boolean(payload.in_drive));
    if (payload.image_url !== undefined) {
      sheet.getRange(row, 14).setValue(String(payload.image_url || "").trim());
    }
    return invJsonOut_({ ok: true, row: row });
  }

  if (action === "delete_product") {
    const row = Number(payload.row);
    if (!row || row < 2) return invJsonOut_({ ok: false, error: "Invalid row" });
    sheet.deleteRow(row);
    return invJsonOut_({ ok: true, deleted_row: row });
  }

  return invJsonOut_({ ok: false, error: "Unknown action" });
}
