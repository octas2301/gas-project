/**
 * §9.7 数量確認メール送信（MailApp）
 * Python が「⏱数量確認メール下書き」へ subject/body/to/token を書いたあと、
 * メニュー / WebApp(doGet) から送信する。
 */
var TS_QTY_DRAFT_SHEET_ = '⏱数量確認メール下書き';
var TS_QTY_DEFAULT_TO_ = 'contact@octas2301.com';

/**
 * 広告スプシを取得（メニュー=Active／Web=ScriptProperties ads_spreadsheet_id）
 */
function tsQtyConfirmSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) { /* web / clasp */ }
  const id = String(PropertiesService.getScriptProperties().getProperty('ads_spreadsheet_id') || '').trim();
  if (!id) {
    throw new Error('ads_spreadsheet_id がありません（メニュー実行なら Active スプシを開いてください）');
  }
  return SpreadsheetApp.openById(id);
}

/**
 * 下書きシートから送信（token 必須時は B5 と一致）
 * @param {string=} requireToken
 */
function sendTimeSaleQtyConfirmMailCore_(requireToken) {
  const ss = tsQtyConfirmSpreadsheet_();
  const sh = ss.getSheetByName(TS_QTY_DRAFT_SHEET_);
  if (!sh) {
    throw new Error('シート「' + TS_QTY_DRAFT_SHEET_ + '」がありません。先に mail_qty_confirm.py を実行してください。');
  }
  if (requireToken !== undefined && requireToken !== null) {
    const expect = String(sh.getRange('B5').getValue() || '').trim();
    const got = String(requireToken || '').trim();
    if (!expect || expect !== got) {
      throw new Error('token mismatch');
    }
  }
  const to = String(sh.getRange('B1').getValue() || TS_QTY_DEFAULT_TO_).trim();
  const subject = String(sh.getRange('B2').getValue() || '').trim();
  const body = String(sh.getRange('B3').getValue() || '').trim();
  if (!to) throw new Error('宛先(B1)が空です');
  if (!subject) throw new Error('件名(B2)が空です');
  if (!body) throw new Error('本文(B3)が空です');

  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: body
  });

  sh.getRange('B4').setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') +
    ' 送信済 → ' + to
  );
  sh.getRange('B5').setValue('');
  Logger.log(JSON.stringify({
    stepName: 'sendTimeSaleQtyConfirmMail',
    state: 'DONE',
    to: to,
    subject: subject,
    bodyChars: body.length
  }));
  return { ok: true, to: to, subject: subject };
}

/**
 * メニュー／手動: 下書きシートの内容をメール送信
 */
function sendTimeSaleQtyConfirmMail() {
  const result = sendTimeSaleQtyConfirmMailCore_();
  try {
    SpreadsheetApp.getUi().alert('送信しました。\nTo: ' + result.to + '\n件名: ' + result.subject);
  } catch (e) {
    // UI 無し
  }
  return result;
}

/** メニュー用エイリアス */
function menuTimeSaleQtyConfirmSend() {
  sendTimeSaleQtyConfirmMail();
}

/**
 * ポイントリマインド下書き送信（mail_points_remind.py が書いたシート）
 */
var TS_POINTS_DRAFT_SHEET_ = '⏱ポイントリマインド下書き';

function sendTimeSalePointsRemindMailCore_(requireToken) {
  const ss = tsQtyConfirmSpreadsheet_();
  const sh = ss.getSheetByName(TS_POINTS_DRAFT_SHEET_);
  if (!sh) {
    throw new Error('シート「' + TS_POINTS_DRAFT_SHEET_ + '」がありません。先に mail_points_remind.py を実行してください。');
  }
  if (requireToken !== undefined && requireToken !== null) {
    const expect = String(sh.getRange('B5').getValue() || '').trim();
    const got = String(requireToken || '').trim();
    if (!expect || expect !== got) {
      throw new Error('token mismatch');
    }
  }
  const to = String(sh.getRange('B1').getValue() || TS_QTY_DEFAULT_TO_).trim();
  const subject = String(sh.getRange('B2').getValue() || '').trim();
  const body = String(sh.getRange('B3').getValue() || '').trim();
  if (!to) throw new Error('宛先(B1)が空です');
  if (!subject) throw new Error('件名(B2)が空です');
  if (!body) throw new Error('本文(B3)が空です');

  MailApp.sendEmail({ to: to, subject: subject, body: body });
  sh.getRange('B4').setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') +
    ' 送信済 → ' + to
  );
  sh.getRange('B5').setValue('');
  Logger.log(JSON.stringify({
    stepName: 'sendTimeSalePointsRemindMail',
    state: 'DONE',
    to: to,
    subject: subject,
    bodyChars: body.length
  }));
  return { ok: true, to: to, subject: subject };
}

function sendTimeSalePointsRemindMail() {
  const result = sendTimeSalePointsRemindMailCore_();
  try {
    SpreadsheetApp.getUi().alert('送信しました。\nTo: ' + result.to + '\n件名: ' + result.subject);
  } catch (e) { /* UI 無し */ }
  return result;
}

function menuTimeSalePointsRemindSend() {
  sendTimeSalePointsRemindMail();
}

/**
 * P0-G5: Points 送信は SP-API のため GAS 本体では走らせない。
 * P1b／施策同期と同型の Cursor 指示ダイアログのみ。
 */
function menuTimeSalePointsApplyCursorPrompt() {
  const prompt = [
    'gas-project で Amazon タイムセールのポイント apply（期間中%）を実行してください。',
    '',
    'ルール:',
    '- 既定 dry_run。本番は社長承認後のみ --prod --i-confirm-prod',
    '- セール前%が空なら先に fetch（G8）。既に期間中%なら差分なし→SC確認のみ可',
    '- 対象は施策B連動（解除は --all-master／--sku）',
    '',
    '1. cd tools/amazon_deals_bulk',
    '2. python points_fetch.py --write',
    '3. python points_send.py --mode apply --backup-before',
    '4. （承認後）python points_send.py --mode apply --backup-before --prod --i-confirm-prod --wait --update-sheet',
    '5. feedStatus／マスタのポイント状態を報告',
    '',
    'リマインド下書き生成のみなら:',
    '  python mail_points_remind.py --kind apply --days 1',
    '送信はメニュー「3. セール開始前日／終了翌日：ポイント作業の催促メールを送る（下書き送信）」または --send',
    '',
    '要件§10.10: docs/org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md',
    '手順: docs/org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md'
  ].join('\n');
  showTimeSaleCursorDialog_('4. 期間中ポイント%をAmazonに載せる — Agent依頼文', prompt, 'menuTimeSalePointsApplyCursorPrompt');
}

function menuTimeSalePointsRestoreCursorPrompt() {
  const prompt = [
    'gas-project で Amazon タイムセールのポイント restore（減衰中%へ戻す）を実行してください。',
    '',
    'ルール:',
    '- 既定 dry_run。本番は社長承認後のみ --prod --i-confirm-prod',
    '- 戻す先は 減衰中ポイント%（カレンダー位置）。最終終着%（セール前列）ではない',
    '- 先に python taper_send.py --poll --mail でカレンダー同期してから restore',
    '- B期間中の restore 禁止。終了後に店頭1% → その日の減衰中%',
    '- 対象は施策B連動（解除は --all-master／--sku）',
    '',
    '1. cd tools/amazon_deals_bulk',
    '2. python taper_send.py --poll --mail',
    '3. python points_send.py --mode restore',
    '4. （承認後）python points_send.py --mode restore --prod --i-confirm-prod --wait --update-sheet',
    '5. feedStatus／マスタのポイント状態を報告',
    '',
    'リマインド下書き生成のみなら:',
    '  python mail_points_remind.py --kind restore --days 1',
    '送信はメニュー「3. セール開始前日／終了翌日：ポイント作業の催促メールを送る（下書き送信）」または --send',
    '',
    '要件§10.10 / §10.13 / §10.14: docs/org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md',
    '手順: docs/org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md'
  ].join('\n');
  showTimeSaleCursorDialog_('5. 減衰中ポイント%をAmazonに戻す — Agent依頼文', prompt, 'menuTimeSalePointsRestoreCursorPrompt');
}

/**
 * 実質戻し: 減衰計画 dry_run／売価スナップ（--snap）
 */
function menuTimeSalePriceRecoverySendCursorPrompt() {
  const prompt = [
    'gas-project で Amazon タイムセールのポイント減衰（taper）を実行してください。',
    '',
    '方針: 最終売価固定。減衰はポイント%。細承認なし。結果メールを見て異常なら手動修正。',
    '',
    '1. cd tools/amazon_deals_bulk',
    '2. python taper_send.py --poll --mail',
    '3. （1SKU初回）python taper_send.py --start --sku \"…\" --mail',
    '4. （承認後本番）python taper_send.py --sku \"…\" --prod --i-confirm-prod --wait --update-sheet --mail',
    '5. 日次タスク推奨: python taper_send.py --poll --mail   ※最初1週間は dry_run',
    '6. 売価スナップが必要なら: python price_recovery_send.py --snap --sku \"…\"',
    '',
    '手動リカバリ: メールの商品URL確認 → マスタの減衰中ポイント%を直す →',
    '  python points_send.py --sku \"…\" --all --prod --i-confirm-prod --wait --update-sheet',
    '',
    '要件§10.14: docs/org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md',
    '手順: docs/org/D_MENU_AMAZON_DEALS_BULK_P1B_HUMAN_RUN.md'
  ].join('\n');
  showTimeSaleCursorDialog_('99-⑦ 減衰を手動で1回回す — Agent依頼文', prompt, 'menuTimeSalePriceRecoverySendCursorPrompt');
}

/**
 * WebApp: ?ssid=...&token=... （B5 の web_token と一致で1回送信）
 * ?kind=points でポイントリマインド下書きを送信
 */
function doGet(e) {
  e = e || { parameter: {} };
  const p = e.parameter || {};
  const ssid = String(p.ssid || '').trim();
  if (ssid) {
    PropertiesService.getScriptProperties().setProperty('ads_spreadsheet_id', ssid);
  }
  try {
    const kind = String(p.kind || '').trim().toLowerCase();
    const result = (kind === 'points')
      ? sendTimeSalePointsRemindMailCore_(String(p.token || ''))
      : sendTimeSaleQtyConfirmMailCore_(String(p.token || ''));
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
