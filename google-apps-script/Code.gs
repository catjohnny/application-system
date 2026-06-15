const SPREADSHEET_ID = "1RnrcGm6imGgE3J_MEwuvGv5U5tJfJ3LLCBaJWUZHlTM";
const SHEET_NAME = "";

function doGet(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || "lookup";

    if (action === "lookup") {
      return jsonResponse({
        success: true,
        items: findChildren(params.q || params.name || params.phone || "")
      });
    }

    if (action === "updateStatus") {
      return jsonResponse(updateStatus(params.q || params.name || params.phone || "", params.status || ""));
    }

    if (action === "phoneByName") {
      return jsonResponse({
        success: true,
        items: findPhoneByName(params.name || "")
      });
    }

    if (action === "statusLists") {
      return jsonResponse(getStatusLists());
    }

    return jsonResponse({ success: false, message: "未知的 action" });
  } catch (error) {
    return jsonResponse({ success: false, message: error.toString() });
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return SHEET_NAME
    ? spreadsheet.getSheetByName(SHEET_NAME)
    : spreadsheet.getSheets()[0];
}

function readRows() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  return values.map((row, index) => ({
    rowNumber: index + 2,
    name: String(row[0] || "").trim(),
    phone: String(row[1] || "").trim(),
    team: String(row[2] || "").trim(),
    status: String(row[3] || "").trim()
  })).filter((item) => item.name || item.phone);
}

function findChildren(query) {
  const keyword = String(query || "").trim();
  if (!keyword) return [];

  const normalizedKeywordPhone = normalizePhone(keyword);
  return readRows()
    .filter((item) => {
      const nameMatched = item.name.indexOf(keyword) !== -1;
      const phoneMatched = normalizedKeywordPhone &&
        normalizePhone(item.phone).indexOf(normalizedKeywordPhone) !== -1;
      return nameMatched || phoneMatched;
    })
    .map(toPublicItem);
}

function updateStatus(query, status) {
  const targetStatus = String(status || "").trim();
  if (targetStatus !== "已報到" && targetStatus !== "離場") {
    return { success: false, message: "狀態只能是已報到或離場" };
  }

  const keyword = String(query || "").trim();
  if (!keyword) {
    return { success: false, message: "請輸入姓名或電話" };
  }

  const sheet = getSheet();
  const normalizedKeywordPhone = normalizePhone(keyword);
  const matchedRows = readRows().filter((item) => {
    const nameMatched = item.name === keyword;
    const phoneMatched = normalizedKeywordPhone &&
      normalizePhone(item.phone) === normalizedKeywordPhone;
    return nameMatched || phoneMatched;
  });

  if (!matchedRows.length) {
    return { success: false, message: "查無資料" };
  }

  matchedRows.forEach((item) => {
    sheet.getRange(item.rowNumber, 4).setValue(targetStatus);
    item.status = targetStatus;
  });

  return {
    success: true,
    items: matchedRows.map(toPublicItem)
  };
}

function findPhoneByName(name) {
  const keyword = String(name || "").trim();
  if (!keyword) return [];

  return readRows()
    .filter((item) => item.name.indexOf(keyword) !== -1)
    .map((item) => ({
      name: item.name,
      phone: item.phone
    }));
}

function getStatusLists() {
  const result = {
    notCheckedIn: [],
    checkedIn: [],
    left: []
  };

  readRows().forEach((item) => {
    const publicItem = toPublicItem(item);
    if (item.status === "已報到") {
      result.checkedIn.push(publicItem);
    } else if (item.status === "離場") {
      result.left.push(publicItem);
    } else {
      result.notCheckedIn.push(publicItem);
    }
  });

  return result;
}

function toPublicItem(item) {
  return {
    name: item.name,
    phone: item.phone,
    team: item.team,
    status: item.status
  };
}

function normalizePhone(value) {
  return String(value || "").replace(/\D/g, "");
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
