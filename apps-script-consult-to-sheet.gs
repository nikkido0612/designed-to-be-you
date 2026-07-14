var NOTIFY_EMAIL = 'nikki.humandesign@gmail.com';

var HEADERS = [
  '送出時間', '預約項目', '姓名', '電話', 'LINE帳號', 'Email',
  '出生日期', '出生時間', '出生地點',
  '伴侶姓名', '伴侶出生日期', '伴侶出生時間', '伴侶出生地點',
  '方便時段', '其他方便時間備註',
  '付款方式', '轉帳後5碼',
  '想探索的面向'
];

function formatHeader() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground('#7289AD');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
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
    data['預約項目'] || '',
    data['姓名'] || '',
    data['電話'] || '',
    data['LINE帳號'] || '',
    data['Email'] || '',
    data['出生日期'] || '',
    data['出生時間'] || '',
    data['出生地點'] || '',
    data['伴侶姓名'] || '',
    data['伴侶出生日期'] || '',
    data['伴侶出生時間'] || '',
    data['伴侶出生地點'] || '',
    data['方便時段'] || '',
    data['其他方便時間備註'] || '',
    data['付款方式'] || '',
    data['轉帳後5碼'] || '',
    data['想探索的面向'] || ''
  ]);

  if (data['Email']) {
    sendConfirmationEmail(data);
  }

  sendNotificationEmail(data);

  return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendConfirmationEmail(data) {
  var name = data['姓名'] || '你好';

  var body =
    name + '，你好：\n\n' +
    '感謝你預約人類圖諮詢！\n\n' +
    '我已經收到你的預約資料（' + (data['預約項目'] || '') + '），會主動透過 LINE 與你聯繫，確認時間：\n' +
    'LINE：nikkido0612\n\n' +
    '期待與你相見！\n' +
    '楊明怡 NIKKI';

  MailApp.sendEmail({
    to: data['Email'],
    subject: '🌱 預約確認｜人類圖諮詢',
    body: body
  });
}

function sendNotificationEmail(data) {
  var body =
    '有新的預約送出：\n\n' +
    '預約項目：' + (data['預約項目'] || '') + '\n' +
    '姓名：' + (data['姓名'] || '') + '\n' +
    '電話：' + (data['電話'] || '') + '\n' +
    'LINE 帳號：' + (data['LINE帳號'] || '') + '\n' +
    'Email：' + (data['Email'] || '') + '\n' +
    '方便時段：' + (data['方便時段'] || '') + '\n' +
    (data['其他方便時間備註'] ? ('其他方便時間備註：' + data['其他方便時間備註'] + '\n') : '') +
    '付款方式：' + (data['付款方式'] || '') + '\n' +
    (data['轉帳後5碼'] ? ('轉帳後5碼：' + data['轉帳後5碼'] + '\n') : '') +
    '\n完整資料請查看 Google Sheet。';

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: '📋 新預約：' + (data['姓名'] || '') + '（' + (data['預約項目'] || '') + '）',
    body: body
  });
}
