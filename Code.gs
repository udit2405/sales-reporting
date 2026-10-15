// Google Apps Script — paste into Extensions > Apps Script of the Sales sheet, then Deploy > Web app
// (Execute as: Me, Who has access: Anyone). Copy the /exec URL into index.html (SCRIPT_URL).
const SHEET_GID = 0;
const FIELDS = ['VIP Sale', 'SKY Sale', 'ARI Sale', 'VIP Total Sale', 'VIP MTD Sale',
                'AT', 'KT', 'AT TOTAL SALE', 'AT MTD', 'SAFARI SALE', 'SAFARI MTD'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.getActive().getSheets().filter(s => s.getSheetId() === SHEET_GID)[0];
    const vals = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < vals.length; i++) if (String(vals[i][0]).trim() === String(d.code).trim()) { row = i + 1; break; }
    if (row < 0) return out({ ok: false, error: 'Store not found' });
    // columns C..M (3..13) = the 11 value columns, N = last updated
    sh.getRange(row, 3, 1, FIELDS.length).setValues([FIELDS.map(f => d.values[f] === '' || d.values[f] == null ? '' : Number(d.values[f]))]);
    sh.getRange(1, 14).setValue('Last Updated');
    sh.getRange(row, 14).setValue(new Date());
    return out({ ok: true });
  } catch (err) { return out({ ok: false, error: String(err) }); }
}
function doGet() { return out({ ok: true, msg: 'Sales reporting endpoint live' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
