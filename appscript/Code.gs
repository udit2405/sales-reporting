// Sales Reporting — complete Apps Script app (form + sheet writer).
// Setup: open the sheet > Extensions > Apps Script > paste this as Code.gs, add Index.html (File > + > HTML, name "Index"),
// then Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone > Deploy. Share the web app URL.
const FIELDS = ['VIP Sale', 'SKY Sale', 'ARI Sale', 'VIP Total Sale', 'VIP MTD Sale',
                'AT', 'KT', 'AT TOTAL SALE', 'AT MTD', 'SAFARI SALE', 'SAFARI MTD'];
const FIRST_VALUE_COL = 3;                 // column C
const STAMP_COL = FIRST_VALUE_COL + FIELDS.length; // column N

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Daily Sales Report')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function sheet_() { return SpreadsheetApp.getActive().getSheets()[0]; }

// live store list straight from the sheet (col A = SAP code, col B = store name)
function getStores() {
  const v = sheet_().getDataRange().getValues();
  return v.slice(1).filter(r => r[0] !== '' && r[1] !== '')
          .map(r => ({ code: String(r[0]).trim(), name: String(r[1]).trim() }));
}

function submitSale(code, values) {
  const sh = sheet_();
  const v = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === String(code).trim()) { row = i + 1; break; }
  if (row < 0) throw new Error('Store not found');
  sh.getRange(row, FIRST_VALUE_COL, 1, FIELDS.length)
    .setValues([FIELDS.map(f => values[f] === '' || values[f] == null ? '' : Number(values[f]))]);
  sh.getRange(1, STAMP_COL).setValue('Last Updated');
  sh.getRange(row, STAMP_COL).setValue(new Date());
  return 'ok';
}
