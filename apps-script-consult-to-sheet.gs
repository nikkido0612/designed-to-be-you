function formatHeader() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = ['送出時間', '姓名', '電話', 'LINE帳號', 'Email', '困惑或決定'];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#C9A87C');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);

  var subId = body.id;
  if (subId) {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'sub_' + subId;
    if (cache.get(cacheKey)) {
      return ContentService.createTextOutput(JSON.stringify({status: 'duplicate_ignored'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    cache.put(cacheKey, '1', 300);
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    formatHeader();
  }

  var data = body.data || {};

  sheet.appendRow([
    new Date(),
    data['姓名'] || '',
    data['電話'] || '',
    data['LINE帳號'] || '',
    data['Email'] || '',
    data['困惑或決定'] || ''
  ]);

  if (data['Email']) {
    sendConfirmationEmail(data);
  }

  return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendConfirmationEmail(data) {
  var name = data['姓名'] || '你好';

  var body =
    name + '，你好：\n\n' +
    '感謝你預約「Designed to be YOU」單次深度解讀！\n\n' +
    '我已經收到你的預約資料，會主動透過 LINE 與你聯繫，確認時間：\n' +
    'LINE：nikkido0612\n\n' +
    '期待與你相見！\n' +
    '楊明怡 NIKKI';

  MailApp.sendEmail({
    to: data['Email'],
    subject: '🌱 預約確認｜Designed to be YOU 一對一深度解讀',
    body: body
  });
}
