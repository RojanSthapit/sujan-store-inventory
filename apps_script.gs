const INV_ACCESS_PIN = "";
const INV_SHEET_NAME = "Sheet1";
const INV_SOLD_SHEET_NAME = "Sold";
const INV_DRIVE_FOLDER_ID = "1-4hk_FmRbfv6MCCA2LWbcdBhT7dkduq8";

function invCheckPin_(pin) {
  if (!INV_ACCESS_PIN) return true;
  return String(pin || "") === String(INV_ACCESS_PIN);
}

function invJsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function invFindDriveImageForCode_(pdtCode) {
  try {
    const root = DriveApp.getFolderById(INV_DRIVE_FOLDER_ID);
    const folders = root.getFoldersByName(pdtCode);
    if (!folders.hasNext()) return { inDrive: false, imageUrl: "" };
    const folder = folders.next();
    const files = folder.getFiles();
    const imgNames = [];
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName().toLowerCase();
      if (/\.(jpg|jpeg|png|webp|gif)$/.test(name)) {
        imgNames.push({ name: f.getName(), id: f.getId() });
      }
    }
    if (!imgNames.length) return { inDrive: true, imageUrl: "" };
    imgNames.sort((a, b) => a.name.localeCompare(b.name));
    return {
      inDrive: true,
      imageUrl: "https://drive.google.com/thumbnail?id=" + imgNames[0].id + "&sz=w800",
    };
  } catch (err) {
    return { inDrive: false, imageUrl: "" };
  }
}

function invNormalizeDriveUrl_(url) {
  const s = String(url || "").trim();
  if (!s) return "";
  const m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m && s.indexOf("drive.google.com") !== -1) {
    return "https://drive.google.com/thumbnail?id=" + m[1] + "&sz=w800";
  }
  return s;
}

function doGet(e) {
  const pin = (e && e.parameter && e.parameter.pin) || "";
  if (!invCheckPin_(pin)) {
    return invJsonOut_({ ok: false, error: "Invalid PIN" });
  }

  const mode = (e && e.parameter && e.parameter.mode) || "products";
  if (mode === "sales") {
    const soldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INV_SOLD_SHEET_NAME);
    const lastRow = soldSheet.getLastRow();
    if (lastRow < 2) return invJsonOut_({ ok: true, sales: [], next_bill_no: 1 });
    const vals = soldSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const sales = [];
    let maxBill = 0;
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const billNo = Number(r[0]) || 0;
      const code = String(r[1] || "").trim();
      const name = String(r[2] || "").trim();
      if (!code && !name) continue;
      if (billNo > maxBill) maxBill = billNo;
      let dateStr = "";
      if (r[9] instanceof Date) {
        dateStr = Utilities.formatDate(r[9], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dateStr = String(r[9] || "").trim();
      }
      sales.push({
        row: i + 2,
        bill_no: billNo || String(r[0] || ""),
        code: code,
        name: name,
        qty: Number(r[3]) || 1,
        daraz: Number(r[4]) || 0,
        cp: Number(r[5]) || 0,
        es_prof: Number(r[6]) || 0,
        sp: Number(r[7]) || 0,
        profit: Number(r[8]) || 0,
        sold_date: dateStr,
        platform: String(r[10] || "").trim(),
        daraz_code: String(r[11] || "").trim(),
      });
    }
    return invJsonOut_({ ok: true, sales: sales.reverse(), next_bill_no: maxBill + 1 });
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

    let imageUrl = invNormalizeDriveUrl_(payload.image_url);
    let inDrive = Boolean(payload.in_drive);
    if (!imageUrl) {
      const driveInfo = invFindDriveImageForCode_(pdtCode);
      if (driveInfo.inDrive) inDrive = true;
      if (driveInfo.imageUrl) imageUrl = driveInfo.imageUrl;
    }

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
      inDrive,
      imageUrl,
    ]]);
    sheet.getRange(targetRow, 13).insertCheckboxes();
    return invJsonOut_({ ok: true, row: targetRow, code: pdtCode, image_url: imageUrl, in_drive: inDrive });
  }

  if (action === "edit_product") {
    const row = Number(payload.row);
    if (!row || row < 2) return invJsonOut_({ ok: false, error: "Invalid row" });
    const pdtCode = String(payload.code || "").trim();
    let imageUrl = invNormalizeDriveUrl_(payload.image_url);
    let inDrive = Boolean(payload.in_drive);
    if (!imageUrl && pdtCode) {
      const driveInfo = invFindDriveImageForCode_(pdtCode);
      if (driveInfo.inDrive) inDrive = true;
      if (driveInfo.imageUrl) imageUrl = driveInfo.imageUrl;
    }

    sheet.getRange(row, 2).setValue(pdtCode);
    sheet.getRange(row, 3).setValue(Number(payload.cost) || 0);
    sheet.getRange(row, 4).setValue(Number(payload.daraz_sp) || 0);
    sheet.getRange(row, 9).setValue(String(payload.group || "").trim());
    sheet.getRange(row, 10).setValue(String(payload.name || "").trim());
    sheet.getRange(row, 12).setValue(String(payload.status || "").trim());
    sheet.getRange(row, 13).setValue(inDrive);
    sheet.getRange(row, 14).setValue(imageUrl);
    return invJsonOut_({ ok: true, row: row, image_url: imageUrl, in_drive: inDrive });
  }

  if (action === "sync_drive_images") {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return invJsonOut_({ ok: true, updated: 0 });
    const root = DriveApp.getFolderById(INV_DRIVE_FOLDER_ID);
    const sub = root.getFolders();
    const driveMap = {};
    while (sub.hasNext()) {
      const f = sub.next();
      driveMap[f.getName().trim()] = f;
    }
    const vals = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    let updated = 0;
    for (let i = 0; i < vals.length; i++) {
      const code = String(vals[i][1] || "").trim();
      if (!code || !driveMap[code]) continue;
      const hasImg = String(vals[i][13] || "").trim();
      const inDrive = vals[i][12] === true || String(vals[i][12]).toUpperCase() === "TRUE";
      if (!inDrive) {
        sheet.getRange(i + 2, 13).setValue(true);
      }
      if (!hasImg) {
        const files = driveMap[code].getFiles();
        const imgs = [];
        while (files.hasNext()) {
          const file = files.next();
          if (/\.(jpg|jpeg|png|webp|gif)$/i.test(file.getName())) {
            imgs.push({ name: file.getName(), id: file.getId() });
          }
        }
        if (imgs.length) {
          imgs.sort((a, b) => a.name.localeCompare(b.name));
          sheet.getRange(i + 2, 14).setValue("https://drive.google.com/thumbnail?id=" + imgs[0].id + "&sz=w800");
          updated++;
        }
      }
    }
    return invJsonOut_({ ok: true, updated: updated });
  }

  if (action === "delete_product") {
    const row = Number(payload.row);
    if (!row || row < 2) return invJsonOut_({ ok: false, error: "Invalid row" });
    sheet.deleteRow(row);
    return invJsonOut_({ ok: true, deleted_row: row });
  }

  if (action === "add_sale") {
    const soldSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INV_SOLD_SHEET_NAME);
    const lastRow = Math.max(1, soldSheet.getLastRow());
    const colB = lastRow > 1 ? soldSheet.getRange(2, 2, lastRow - 1, 1).getValues() : [];
    const colA = lastRow > 1 ? soldSheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
    let targetRow = 2;
    let maxBill = 0;
    for (let i = 0; i < colB.length; i++) {
      if (String(colB[i][0] || "").trim()) {
        targetRow = i + 3;
        const b = Number(colA[i][0]);
        if (!isNaN(b) && b > maxBill) maxBill = b;
      }
    }

    const items = Array.isArray(payload.items) ? payload.items : [];
    if (!items.length) return invJsonOut_({ ok: false, error: "No items" });

    const billNo = Number(payload.bill_no) || (maxBill + 1);
    const soldDate = payload.sold_date ? new Date(payload.sold_date + "T00:00:00") : new Date();
    const platform = String(payload.platform || "Office").trim();
    const darazCode = String(payload.daraz_code || "").trim();

    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const r = targetRow + i;
      const qty = Number(it.qty) || 1;
      const darazPrice = Number(it.daraz) || 0;
      const cp = Number(it.cp) || 0;
      const sp = Number(it.sp) || 0;
      soldSheet.getRange(r, 1, 1, 12).setValues([[
        billNo,
        String(it.code || "").trim(),
        String(it.name || "").trim(),
        qty,
        darazPrice,
        cp,
        `=E${r}-F${r}`,
        sp,
        `=H${r}-F${r}`,
        soldDate,
        platform,
        darazCode,
      ]]);
    }
    return invJsonOut_({ ok: true, bill_no: billNo, start_row: targetRow, count: items.length });
  }

  return invJsonOut_({ ok: false, error: "Unknown action" });
}
