// Ойын нәтижелерін Google Кестеге жинау.
// Кесте → Кеңейтімдер → Apps Script → осы кодты қойып, «Веб-қосымша» ретінде жариялаңыз.
const SHEET_NAME = 'Нәтижелер';
const HEAD = ['Уақыты', 'Ойын', 'Топ', 'Оқиға', 'Тапсырма', 'Жауап', 'Нәтиже', 'Ұпай / уақыт', 'Жалпы'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return sh;
}

// Ойыннан келген нәтижені жаңа жол етіп жазады
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_().appendRow([
      new Date(), d.game || '', d.team || '', d.event || '',
      d.task || '', d.answer || '', d.result || '',
      d.points === undefined ? '' : d.points, d.total === undefined ? '' : d.total
    ]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// Мұғалім панелі нәтижелерді осы арқылы оқиды
function doGet() {
  const rows = sheet_().getDataRange().getValues();
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}
