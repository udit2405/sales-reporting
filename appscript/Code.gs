// BACKEND ONLY — the website (GitHub Pages) is the interface; this just saves data into the sheet.
// Paste over Code.gs > Deploy > Manage deployments > Edit > New version (access: Anyone).
var FIELDS = ['VIP Sale', 'SKY Sale', 'ARI Sale', 'VIP Total Sale', 'VIP MTD Sale',
              'AT', 'KT', 'AT TOTAL SALE', 'AT MTD', 'SAFARI SALE', 'SAFARI MTD'];
var FIRST_VALUE_COL = 3;   // column C
var STAMP_COL = 14;        // column N

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var sh = SpreadsheetApp.getActive().getSheets()[0];
    var v = sh.getDataRange().getValues(), row = -1;
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][0]).trim() === String(d.code).trim()) { row = i + 1; break; }
    }
    if (row < 0) return out_({ ok: false, error: 'Store not found' });
    var vals = FIELDS.map(function (f) { var x = d.values[f]; return (x === '' || x == null) ? '' : Number(x); });
    sh.getRange(row, FIRST_VALUE_COL, 1, FIELDS.length).setValues([vals]);
    sh.getRange(1, STAMP_COL).setValue('Last Updated');
    sh.getRange(row, STAMP_COL).setValue(new Date());
    return out_({ ok: true });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}
function doGet() { return out_({ ok: true, msg: 'Sales endpoint live' }); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
