/**
 * 予約フォーム受信用 Google Apps Script（雛形）
 * 1. スプレッドシートを開き「拡張機能 > Apps Script」にこのコードを貼り付け
 * 2. 「デプロイ > 新しいデプロイ > ウェブアプリ」
 *    実行ユーザー: 自分 / アクセス: 全員 でデプロイ
 * 3. 表示されたURLを index.html の GAS_URL に貼り付け
 */
const SHOP_NAME = 'いえピカ（仮）';
const SHOP_TEL  = '000-0000-0000';
const NOTIFY_TO = Session.getEffectiveUser().getEmail(); // 店舗側の通知先

const COLUMNS = [
  ['submittedAt','受付日時'],['name','お名前'],['tel','電話番号'],['email','メール'],
  ['area','エリア'],['zip','郵便番号'],['address','住所'],['house','住居タイプ'],
  ['units','台数'],['auto','お掃除機能付き'],['set3','3点セット'],['coat','防カビ抗菌'],
  ['outdoor','室外機'],['drain','ドレンホース'],['total','概算金額'],
  ['date1','第1希望日'],['time1','第1時間'],['date2','第2希望日'],['time2','第2時間'],
  ['date3','第3希望日'],['time3','第3時間'],['payment','支払方法'],['parking','駐車場'],
  ['other','その他掃除'],['memo','備考'],['referral','紹介コード'],
];

function doPost(e) {
  const d = e.parameter;
  const sheet = getSheet_();
  sheet.appendRow(COLUMNS.map(([k]) => d[k] || ''));

  const summary = COLUMNS.slice(1).map(([k, label]) => `${label}：${d[k] || ''}`).join('\n');

  if (d.email) {
    MailApp.sendEmail({
      to: d.email,
      subject: `【${SHOP_NAME}】ご予約を受け付けました`,
      body: `${d.name} 様\n\nこの度はご予約ありがとうございます。\n以下の内容で受け付けました。作業日は担当者よりご連絡のうえ確定いたします。\n\n${summary}\n\n${SHOP_NAME}\nTEL ${SHOP_TEL}`,
      name: SHOP_NAME,
    });
  }
  MailApp.sendEmail(NOTIFY_TO, `【新規予約】${d.name} 様 / ¥${Number(d.total || 0).toLocaleString()}`, summary);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('予約');
  if (!sh) {
    sh = ss.insertSheet('予約');
    sh.appendRow(COLUMNS.map(([, label]) => label));
    sh.setFrozenRows(1);
  }
  return sh;
}
