const SHEET_NAME = "valen_dashboard";
const GOALS_SHEET = "valen_goals";
const KPI_SHEET_ID = "1hdo2zSWKUJJ1UhZVJQYt815Gn82fOx_KOcodnCoRAms";

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'load') return loadData();
  if (action === 'load_kpi') return loadKPI();
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

// ── 대시보드 데이터 로드 ──
function loadData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = readSheet(ss, SHEET_NAME, ['id','month','week','part','brand','sa','ra','memo']);
  const goals = readSheet(ss, GOALS_SHEET, ['id','month','part','brand','sg','rg']);
  return ContentService.createTextOutput(JSON.stringify({records, goals})).setMimeType(ContentService.MimeType.JSON);
}

// ── KPI 시트 로드 ──
function loadKPI() {
  try {
    const ss = SpreadsheetApp.openById(KPI_SHEET_ID);

    // 2026 시트
    const sheet2026 = ss.getSheetByName('올영파트_대시보드_2026');
    // 월별 피벗
    const pivotSheet = ss.getSheetByName('월별 피벗');

    const result = {
      summary: {},
      monthly: [],
      brandFee: []
    };

    // ── 요약 (팀 총 매출, 목표 달성률) ──
    if (sheet2026) {
      const vals = sheet2026.getDataRange().getValues();
      for (let i = 0; i < vals.length; i++) {
        const row = vals[i];
        if (String(row[0]).includes('팀 총 매출')) result.summary.totalRevenue = row[1];
        if (String(row[0]).includes('팀 목표 달성률')) result.summary.achieveRate = row[1];
        if (String(row[0]).includes('팀 총 매출') && row[3] === '1Q') {
          // 분기별 데이터
          result.summary.quarters = [];
        }
      }

      // 분기별 데이터 파싱
      const quarters = [];
      for (let i = 0; i < vals.length; i++) {
        const row = vals[i];
        const q = String(row[3]);
        if (q === '1Q' || q === '2Q' || q === '3Q' || q === '4Q') {
          quarters.push({
            quarter: q,
            target: row[4],
            achieved: row[5] || 0,
            rate: row[6] || 0
          });
        }
      }
      result.summary.quarters = quarters;

      // 월별 목표/달성 파싱
      for (let i = 0; i < vals.length; i++) {
        const row = vals[i];
        if (String(row[0]) === '광고 수익' || String(row[1]) === '목표') {
          // 헤더 행 찾기
          const headerRow = vals[i - 1]; // 구분 행
          if (headerRow) {
            const months = [];
            for (let c = 2; c < headerRow.length; c++) {
              if (headerRow[c] && !isNaN(Number(headerRow[c]))) {
                months.push({ col: c, month: Number(headerRow[c]) });
              }
            }
            // 목표
            const targetRow = vals[i];
            // 달성
            const achieveRow = vals[i + 1] || [];
            // 달성율
            const rateRow = vals[i + 2] || [];

            months.forEach(m => {
              result.monthly.push({
                month: m.month,
                target: targetRow[m.col] || 0,
                achieved: achieveRow[m.col] || 0,
                rate: rateRow[m.col] || 0
              });
            });
            break;
          }
        }
      }
    }

    // ── 브랜드별 FEE (월별 피벗에서) ──
    if (pivotSheet) {
      const pvals = pivotSheet.getDataRange().getValues();
      // "브랜드 월별 광고비 및 수익" 섹션 파싱
      // 2026년 데이터만 추출
      let inSection = false;
      const brandMap = {};

      for (let i = 0; i < pvals.length; i++) {
        const row = pvals[i];
        const year = row[0];
        const month = row[1];
        const brand = row[2];
        const adCost = row[3];
        const fee = row[4];

        if (year === 2026 && brand && typeof brand === 'string' && brand.trim() !== '' &&
            !brand.includes('총계') && !brand.includes('브랜드') && fee) {
          const key = brand.trim();
          if (!brandMap[key]) brandMap[key] = { brand: key, totalFee: 0, months: {} };
          brandMap[key].totalFee += (typeof fee === 'number' ? fee : 0);
          if (month && typeof month === 'number') {
            brandMap[key].months[month] = (brandMap[key].months[month] || 0) + (typeof fee === 'number' ? fee : 0);
          }
        }
      }

      result.brandFee = Object.values(brandMap)
        .filter(b => b.totalFee > 0)
        .sort((a, b) => b.totalFee - a.totalFee)
        .slice(0, 20);
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 시트 읽기 ──
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

// ── 저장 ──
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
  data.forEach(r => sheet.appendRow(headers.map(h => r[h] !== undefined ? r[h] : '')));
}
