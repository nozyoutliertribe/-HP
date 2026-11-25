const RECIPIENT = 'informationozy@gmail.com';
const DAILY_LIMIT = 50;

function doPost(e) {
  if (!e || !e.parameter) {
    return resultPage('送信できませんでした', 'フォームから送信してください。');
  }

  const fields = e.parameter;
  const name = getField(fields, 'name');
  const email = getField(fields, 'email');
  const phone = getField(fields, 'phone');
  const childName = getField(fields, 'childName');
  const age = getField(fields, 'age');
  const grade = getField(fields, 'grade');
  const tennisHistory = getField(fields, 'tennisHistory');
  const topic = getField(fields, 'topic');
  const preferredDate = getField(fields, 'preferredDate');
  const alternateDate = getField(fields, 'alternateDate');
  const message = getField(fields, 'message');
  const topicLabels = {
    trial: '体験レッスン',
    rental: 'レンタルコートの予約・空き状況',
    team: 'ジュニアチームについて',
    other: 'その他'
  };

  if (getField(fields, 'website')) {
    return resultPage('送信できませんでした', '入力内容を確認してください。');
  }

  if (
    !name ||
    !email ||
    !phone ||
    !Object.prototype.hasOwnProperty.call(topicLabels, topic) ||
    !message ||
    name.length > 80 ||
    email.length > 254 ||
    phone.length > 40 ||
    childName.length > 80 ||
    age.length > 20 ||
    grade.length > 30 ||
    tennisHistory.length > 120 ||
    preferredDate.length > 32 ||
    alternateDate.length > 32 ||
    message.length > 3000 ||
    (['trial', 'rental'].indexOf(topic) !== -1 && !preferredDate) ||
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
  ) {
    return resultPage('入力内容をご確認ください', 'お名前、メールアドレス、メッセージをご確認ください。');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return resultPage('送信できませんでした', '時間をおいて、もう一度お試しください。');
  }

  try {
    const properties = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    const savedDate = properties.getProperty('contactLimitDate');
    const savedCount = Number(properties.getProperty('contactLimitCount') || 0);
    const sentToday = savedDate === today && Number.isFinite(savedCount) ? savedCount : 0;

    let logSheet = null;
    let logRow = 0;
    let logRecorded = false;
    try {
      logSheet = getLogSheet();
      logRow = logSheet.getLastRow() + 1;
      logSheet.appendRow([
        new Date(),
        spreadsheetValue(name),
        spreadsheetValue(email),
        spreadsheetValue(phone),
        spreadsheetValue(message),
        'メール通知待ち',
        spreadsheetValue(childName),
        spreadsheetValue(age),
        spreadsheetValue(grade),
        spreadsheetValue(tennisHistory),
        spreadsheetValue(topicLabels[topic]),
        spreadsheetValue(preferredDate),
        spreadsheetValue(alternateDate)
      ]);
      logRecorded = true;
    } catch (error) {
      console.error('問い合わせの記録に失敗しました: ' + error);
    }

    if (sentToday >= DAILY_LIMIT) {
      if (!logRecorded) {
        return resultPage('受付できませんでした', '本日のメール通知上限に達し、記録にも失敗しました。時間をおいてお試しください。');
      }

      try {
        logSheet.getRange(logRow, 6).setValue('上限超過によりメール未通知');
      } catch (error) {
        console.error('メール通知状況の更新に失敗しました: ' + error);
      }

      return resultPage('受付完了', 'お問い合わせを記録しました。本日のメール通知上限に達したため、メール通知は行われていません。');
    }

    const body = [
      'NOT TCウェブサイトからお問い合わせがありました。',
      '',
      'お名前: ' + name,
      'メールアドレス: ' + email,
      '電話番号: ' + phone,
      '体験者のお子様のお名前: ' + (childName || '未入力'),
      '年齢: ' + (age || '未入力'),
      '学年: ' + (grade || '未入力'),
      'テニス歴: ' + (tennisHistory || '未入力'),
      'ご用件: ' + topicLabels[topic],
      '希望日時（第1希望）: ' + (preferredDate || '未入力'),
      '希望日時（第2希望）: ' + (alternateDate || '未入力'),
      '',
      'お問い合わせ内容:',
      message
    ].join('\n');

    try {
      GmailApp.sendEmail(RECIPIENT, 'NOT TC｜' + topicLabels[topic], body, {
        replyTo: email
      });
    } catch (error) {
      console.error(error);
      if (logRecorded) {
        try {
          logSheet.getRange(logRow, 6).setValue('メール通知失敗');
        } catch (logError) {
          console.error('メール通知状況の更新に失敗しました: ' + logError);
        }
        return resultPage('受付完了', 'お問い合わせはスプレッドシートに記録しましたが、メール通知に失敗しました。');
      }
      return resultPage('送信できませんでした', 'メール通知とスプレッドシートへの記録に失敗しました。時間をおいて、もう一度お試しください。');
    }

    try {
      properties.setProperties({
        contactLimitDate: today,
        contactLimitCount: String(sentToday + 1)
      });
    } catch (error) {
      console.error('メール通知上限の更新に失敗しました: ' + error);
    }

    if (logRecorded) {
      try {
        logSheet.getRange(logRow, 6).setValue('メール通知済み');
      } catch (error) {
        console.error('メール通知状況の更新に失敗しました: ' + error);
        return resultPage('送信完了', 'メールを送信しましたが、スプレッドシートの通知状況を更新できませんでした。');
      }
      if (topic === 'rental') {
        return resultPage('受付完了', 'レンタルコートの希望日時を記録し、メールで通知しました。予約確定は空き状況を確認後にご連絡します。');
      }
      return resultPage('送信完了', 'お問い合わせを記録し、メールを送信しました。ご連絡ありがとうございます。');
    }

    return resultPage('送信完了', 'メールを送信しましたが、スプレッドシートへの記録に失敗しました。');
  } finally {
    lock.releaseLock();
  }
}

function getLogSheet() {
  const headers = [
    '受付日時',
    'お名前',
    'メールアドレス',
    '電話番号',
    'お問い合わせ内容',
    'メール通知状況',
    '体験者のお子様のお名前',
    '年齢',
    '学年',
    'テニス歴',
    'ご用件',
    '希望日時（第1希望）',
    '希望日時（第2希望）'
  ];
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('contactSpreadsheetId');
  const spreadsheet = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.create('NOT TCお問い合わせ一覧');
  let sheet = spreadsheet.getSheetByName('お問い合わせ');

  if (!sheet) {
    if (spreadsheetId) {
      sheet = spreadsheet.insertSheet('お問い合わせ');
    } else {
      sheet = spreadsheet.getSheets()[0];
      sheet.setName('お問い合わせ');
    }
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else {
    const lastColumn = sheet.getLastColumn();
    const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const missingHeaders = headers.filter(function (header) {
      return existingHeaders.indexOf(header) === -1;
    });

    if (missingHeaders.length) {
      sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length).setValues([missingHeaders]);
      sheet.setFrozenRows(1);
    }
  }

  if (!spreadsheetId) {
    properties.setProperty('contactSpreadsheetId', spreadsheet.getId());
  }

  return sheet;
}

function spreadsheetValue(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function getField(fields, name) {
  return String(fields[name] || '').trim();
}

function resultPage(title, message) {
  const safeTitle = escapeHtml(title);
  const safeMessage = escapeHtml(message);
  const html = [
    '<!doctype html>',
    '<html lang="ja">',
    '<head>',
    '<meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>' + safeTitle + '</title>',
    '<style>body{font-family:sans-serif;line-height:1.7;margin:3rem auto;max-width:40rem;padding:0 1rem;color:#1d2b3a}h1{color:#102a43}</style>',
    '</head>',
    '<body>',
    '<h1>' + safeTitle + '</h1>',
    '<p>' + safeMessage + '</p>',
    '<p>この画面を閉じてください。</p>',
    '</body>',
    '</html>'
  ].join('');

  return HtmlService.createHtmlOutput(html).setTitle(title);
}

function escapeHtml(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
