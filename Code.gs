/**
 * TRAVEL LEDGER — Google Apps Script backend
 * ------------------------------------------------
 * Turns a Google Sheet into a free JSON "database" for the travel expense
 * tracker. Deploy as a Web App (Execute as: Me, Who has access: Anyone),
 * then paste the deployment URL into index.html as API_URL.
 *
 * Sheets (auto-created on first run):
 *
 *   "Expenses" tab:
 *   Date       | Item          | Category | Amount | Currency | Rate (1 SGD = x NTD) | SGD   | Paid By | Remarks
 *   2026-10-02 | Beef noodles  | Food     | 250    | NTD      | 23.80                | 10.50 | Wei     | Queue was 20 min
 *   2026-10-02 | Airport taxi  | Transport| 32     | SGD      | 1                    | 32.00 |
 *
 *   SGD = Amount / Rate for NTD entries. The rate is saved with each entry, so changing the rate in the admin
 *   panel later does not rewrite what past expenses cost.
 *
 *   "Settings" tab (key / value):
 *   Key              | Value
 *   defaultCurrency  | NTD
 *   sgdToNtd         | 23.8
 */

const EXPENSES_SHEET = "Expenses";
const SETTINGS_SHEET = "Settings";
const CURRENCIES = ["SGD", "NTD"];
const DEFAULT_SETTINGS = { defaultCurrency: "SGD", sgdToNtd: 23.8 };

function getExpensesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EXPENSES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EXPENSES_SHEET);
    sheet.appendRow(["Date", "Item", "Category", "Amount", "Currency", "Rate (1 SGD = x NTD)", "SGD", "Paid By", "Remarks"]);
    sheet.setFrozenRows(1);
  } else if (sheet.getRange(1, 9).getValue() === "") {
    // Adds the Remarks header to sheets created by the earlier version.
    sheet.getRange(1, 9).setValue("Remarks");
  }
  return sheet;
}

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.appendRow(["Key", "Value"]);
    sheet.appendRow(["defaultCurrency", DEFAULT_SETTINGS.defaultCurrency]);
    sheet.appendRow(["sgdToNtd", DEFAULT_SETTINGS.sgdToNtd]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Normalizes any cell value (Date object or string) to "yyyy-MM-dd".
function normalizeDate_(value) {
  if (value === "" || value === null || value === undefined) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = String(value).trim();
  const match = s.match(/^\d{4}-\d{2}-\d{2}/);
  if (match) return match[0];
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return s;
}

function normalizeCurrency_(value) {
  const c = String(value || "").trim().toUpperCase();
  if (c === "TWD" || c === "NT$" || c === "NT") return "NTD";
  return CURRENCIES.indexOf(c) !== -1 ? c : "SGD";
}

function round2_(n) {
  return Math.round(Number(n) * 100) / 100;
}

function readSettings_() {
  const values = getSettingsSheet_().getDataRange().getValues();
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  values.slice(1).forEach(r => {
    const key = String(r[0] || "").trim();
    if (key === "defaultCurrency") settings.defaultCurrency = normalizeCurrency_(r[1]);
    if (key === "sgdToNtd") {
      const n = Number(r[1]);
      if (!isNaN(n) && n > 0) settings.sgdToNtd = n;
    }
  });
  return settings;
}

function writeSetting_(sheet, key, value) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET -> returns:
// {
//   entries: [{date, item, category, amount, currency, rate, sgd, paidBy, remarks}, ...],
//   settings: { defaultCurrency: "NTD", sgdToNtd: 23.8 }
// }
function doGet(e) {
  const values = getExpensesSheet_().getDataRange().getValues();
  const entries = values.slice(1)
    .filter(r => r[0] !== "" && r[0] !== null)
    .map(r => {
      const currency = normalizeCurrency_(r[4]);
      const amount = Number(r[3]) || 0;
      const rate = currency === "NTD" ? (Number(r[5]) || 0) : 1;
      const storedSgd = Number(r[6]);
      const sgd = !isNaN(storedSgd) && r[6] !== "" ? storedSgd : (rate > 0 ? round2_(amount / rate) : 0);
      return {
        date: normalizeDate_(r[0]),
        item: String(r[1] || ""),
        category: String(r[2] || "Other"),
        amount: amount,
        currency: currency,
        rate: rate,
        sgd: sgd,
        paidBy: String(r[7] || ""),
        remarks: String(r[8] || "")
      };
    });

  return json_({ entries: entries, settings: readSettings_() });
}

// POST -> two shapes depending on body.type:
//
//   expense (default): {date, item, category, amount, currency, paidBy, remarks}
//     -> appends a row to Expenses. The SGD value is calculated here using
//        the current rate from Settings, then frozen on that row.
//
//   settings: {type:"settings", defaultCurrency:"NTD", sgdToNtd:23.8}
//     -> updates the Settings tab
//
// Sent as text/plain from the browser to dodge a CORS preflight.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.type === "settings") {
      const sheet = getSettingsSheet_();
      if (body.defaultCurrency !== undefined) {
        writeSetting_(sheet, "defaultCurrency", normalizeCurrency_(body.defaultCurrency));
      }
      if (body.sgdToNtd !== undefined) {
        const rate = Number(body.sgdToNtd);
        if (isNaN(rate) || rate <= 0) throw new Error("Exchange rate must be a number above zero");
        writeSetting_(sheet, "sgdToNtd", rate);
      }
      return json_({ status: "ok", settings: readSettings_() });
    }

    // default: expense entry
    if (!body.date || body.amount === undefined || !body.category) {
      throw new Error("Missing date, amount, or category");
    }
    const amount = Number(body.amount);
    if (isNaN(amount) || amount <= 0) throw new Error("Amount must be above zero");

    const currency = normalizeCurrency_(body.currency);
    const rate = currency === "NTD" ? readSettings_().sgdToNtd : 1;
    const sgd = round2_(amount / rate);

    getExpensesSheet_().appendRow([
      body.date,
      String(body.item || ""),
      String(body.category),
      amount,
      currency,
      rate,
      sgd,
      String(body.paidBy || "").trim(),
      String(body.remarks || "").trim()
    ]);
    return json_({ status: "ok" });

  } catch (err) {
    return json_({ status: "error", message: err.message });
  } finally {
    lock.releaseLock();
  }
}
