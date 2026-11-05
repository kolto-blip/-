// Ойын нәтижелерін Google Кестеге жинау.
// Кесте → Кеңейтімдер → Apps Script → осы кодты қойып, «Веб-қосымша» ретінде жариялаңыз.
const SHEET_NAME = 'Нәтижелер';
const HEAD = ['Уақыты', 'Ойын', 'Топ', 'Оқиға', 'Тапсырма', 'Жауап', 'Нәтиже', 'Ұпай / уақыт', 'Жалпы'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Скрипт кестеге байланбаған. Кодты кестенің ішінен (Кеңейтімдер → Apps Script) ашылған жобаға қойыңыз.');
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Ойыннан келген нәтижені жаңа жол етіп жазады
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    lock.waitLock(10000);
    sheet_().appendRow([
      new Date(), d.game || '', d.team || '', d.event || '',
      d.task || '', d.answer || '', d.result || '',
      d.points === undefined ? '' : d.points, d.total === undefined ? '' : d.total
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// Мұғалім панелі нәтижелерді осы арқылы оқиды. ?test=1 болса, кестеге тексеру жолын жазады.
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.test) {
      sheet_().appendRow([new Date(), 'Тексеру', 'Мұғалім', 'Тексеру', '', '', 'байланыс бар', '', '']);
      return json_({ ok: true, test: true });
    }
    return json_(sheet_().getDataRange().getValues());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
