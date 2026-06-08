const SHEET_NAME = "valen_dashboard";
const GOALS_SHEET = "valen_goals";

function doGet(e) {
  if (e.parameter.action === 'load') return loadData();
  return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'save') {
      saveData(data.records || [], data.goals || []);
      return ContentService.createTextOutput(JSON.stringify({status:'saved'})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function loadData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = readSheet(ss, SHEET_NAME, ['id','month','week','part','brand','sa','ra','memo']);
  const goals = readSheet(ss, GOALS_SHEET, ['id','month','part','brand','sg','rg']);
  return ContentService.createTextOutput(JSON.stringify({records, goals})).setMimeType(ContentService.MimeType.JSON);
}

function readSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function saveData(records, goals) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeSheet(ss, SHEET_NAME, ['id','month','week','part','brand','sa','ra','memo'], records);
  writeSheet(ss, GOALS_SHEET, ['id','month','part','brand','sg','rg'], goals);
}

function writeSheet(ss, name, headers, data) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  sheet.appendRow(headers);
  const now = new Date().toISOString();
  data.forEach(r => sheet.appendRow(headers.map(h => r[h] !== undefined ? r[h] : '')));
}
