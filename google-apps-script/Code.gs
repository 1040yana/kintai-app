/**
 * らくらく勤怠 - Google スプレッドシート連携用 Apps Script
 * このスクリプトは、空のGoogleスプレッドシートに「拡張機能 > Apps Script」から貼り付けます。
 */
const MEMBER_SHEET = 'メンバー';
const SITE_SHEET = '現場';
const HOLIDAY_SHEET = '休日';
const RECORD_SHEET = '打刻記録';
const HISTORY_SHEET = '変更履歴';
const TYPES = ['出勤', '直行', '退勤', '直帰'];

function setup() {
  const book = SpreadsheetApp.getActive();
  ensureSheet_(book, MEMBER_SHEET, ['氏名', '有効']);
  ensureSheet_(book, SITE_SHEET, ['現場名', '有効']);
  ensureSheet_(book, HOLIDAY_SHEET, ['日付', '休日名', '有効']);
  ensureSheet_(book, RECORD_SHEET, ['記録ID', '氏名', '日付', '打刻種別', '打刻日時', '削除済', '現場名']);
  ensureSheet_(book, HISTORY_SHEET, ['記録日時', '操作', '対象ID', '対象者', '変更前', '変更後']);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PIN')) props.setProperty('ADMIN_PIN', '1234');
}

/**
 * 管理者暗証番号を初期値 1234 に戻します。
 * Apps Script エディタから、この関数を選んで1回だけ実行してください。
 */
function resetAdminPin() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', '1234');
}

function doGet(e) {
  const callback = String(e.parameter.callback || '');
  const result = e.parameter.action === 'verifyAdmin'
    ? { ok: String(e.parameter.pin || '') === PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') }
    : { ok: true, data: readAll_() };
  const body = callback && /^[A-Za-z_$][\w$]*$/.test(callback)
    ? `${callback}(${JSON.stringify(result)});`
    : JSON.stringify(result);
  return ContentService.createTextOutput(body).setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doPost(e) {
  const raw = (e.parameter && e.parameter.payload) || (e.postData && e.postData.contents) || '{}';
  const input = JSON.parse(raw);
  let result;
  try { result = mutate_(input); } catch (error) { result = { ok: false, error: error.message }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function readAll_() {
  const book = SpreadsheetApp.getActive();
  const members = ensureSheet_(book, MEMBER_SHEET, ['氏名', '有効']).getDataRange().getValues().slice(1)
    .filter(row => row[0] && row[1] !== false).map(row => String(row[0]));
  const sites = ensureSheet_(book, SITE_SHEET, ['現場名', '有効']).getDataRange().getValues().slice(1)
    .filter(row => row[0] && row[1] !== false).map(row => String(row[0]));
  const holidays = ensureSheet_(book, HOLIDAY_SHEET, ['日付', '休日名', '有効']).getDataRange().getValues().slice(1)
    .filter(row => row[0] && row[2] !== false).map(row => ({ date: formatDate_(row[0]), name: String(row[1]) }));
  const records = ensureSheet_(book, RECORD_SHEET, ['記録ID', '氏名', '日付', '打刻種別', '打刻日時', '削除済', '現場名']).getDataRange().getValues().slice(1)
    .filter(row => row[0] && row[5] !== true).map(row => ({ id: String(row[0]), name: String(row[1]), date: formatDate_(row[2]), type: String(row[3]), at: new Date(row[4]).getTime(), sites: parseSites_(row[6]) }));
  return { employees: members, sites, holidays, records };
}

function mutate_(input) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const book = SpreadsheetApp.getActive();
    const members = ensureSheet_(book, MEMBER_SHEET, ['氏名', '有効']);
    const records = ensureSheet_(book, RECORD_SHEET, ['記録ID', '氏名', '日付', '打刻種別', '打刻日時', '削除済', '現場名']);
    if (input.action === 'punch') {
      if (!input.name || !input.date || !TYPES.includes(input.type)) throw new Error('打刻内容が不正です');
      const list = members.getDataRange().getValues().slice(1);
      if (!list.some(row => String(row[0]) === input.name && row[1] !== false)) throw new Error('登録されていないメンバーです');
      const group = ['出勤', '直行'].includes(input.type) ? ['出勤', '直行'] : ['退勤', '直帰'];
      const done = records.getDataRange().getValues().slice(1).some(row => String(row[1]) === input.name && formatDate_(row[2]) === input.date && group.includes(String(row[3])) && row[5] !== true);
      const sites = Array.isArray(input.sites) ? input.sites.map(String) : [];
      if (!done) { const id=Utilities.getUuid();records.appendRow([id, input.name, input.date, input.type, new Date(Number(input.at)), false, JSON.stringify(sites)]);audit_(book, '打刻', id, input.name, '', `${input.date} ${input.type}${sites.length ? ' / '+sites.join('、') : ''}`); }
      return { ok: true };
    }
    if (input.action === 'addSiteByStaff') {
      const name=String(input.name || '').trim();
      if (!name) throw new Error('現場名を入力してください');
      const sheet=ensureSheet_(book, SITE_SHEET, ['現場名', '有効']);
      const exists=sheet.getDataRange().getValues().slice(1).some(row=>String(row[0])===name && row[1]!==false);
      if (!exists) { sheet.appendRow([name, true]); audit_(book, 'スタッフによる現場追加', '', String(input.staff || ''), '', name); }
      return { ok: true };
    }
    if (input.action === 'updateRecordSites') {
      const sites = Array.isArray(input.sites) ? input.sites.map(String) : [];
      const rows=records.getDataRange().getValues(); let found=-1;
      for(let i=1;i<rows.length;i++) if(String(rows[i][1])===input.name && formatDate_(rows[i][2])===input.date && String(rows[i][3])===input.type && rows[i][5]!==true) found=i;
      if(found<0) throw new Error('対象の打刻が見つかりません');
      const before=parseSites_(rows[found][6]).join('、');
      records.getRange(found+1,7).setValue(JSON.stringify(sites));
      audit_(book, '現場選択変更', rows[found][0], rows[found][1], before, sites.join('、'));
      return { ok: true };
    }
    if (String(input.pin || '') !== PropertiesService.getScriptProperties().getProperty('ADMIN_PIN')) throw new Error('管理者暗証番号が違います');
    if (input.action === 'addMember') { if (!input.name) throw new Error('氏名を入力してください'); members.appendRow([input.name, true]);audit_(book, 'メンバー追加', '', input.name, '', input.name); }
    else if (input.action === 'removeMember') { const rows=members.getDataRange().getValues(); for(let i=1;i<rows.length;i++) if(String(rows[i][0])===input.name) members.getRange(i+1,2).setValue(false);audit_(book, 'メンバー無効化', '', input.name, input.name, ''); }
    else if (input.action === 'changePin') { if (String(input.newPin || '').length < 4) throw new Error('暗証番号は4文字以上です'); PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', input.newPin); }
    else if (input.action === 'addSite') { if (!input.name) throw new Error('現場名を入力してください'); ensureSheet_(book, SITE_SHEET, ['現場名', '有効']).appendRow([input.name, true]);audit_(book, '現場追加', '', input.name, '', input.name); }
    else if (input.action === 'removeSite') { const sheet=ensureSheet_(book, SITE_SHEET, ['現場名', '有効']), rows=sheet.getDataRange().getValues(); for(let i=1;i<rows.length;i++) if(String(rows[i][0])===input.name) sheet.getRange(i+1,2).setValue(false);audit_(book, '現場無効化', '', input.name, input.name, ''); }
    else if (input.action === 'addHoliday') { if (!input.date || !input.name) throw new Error('休日と名称を入力してください'); ensureSheet_(book, HOLIDAY_SHEET, ['日付', '休日名', '有効']).appendRow([input.date, input.name, true]);audit_(book, '休日追加', '', '', '', `${input.date} ${input.name}`); }
    else if (input.action === 'removeHoliday') { const sheet=ensureSheet_(book, HOLIDAY_SHEET, ['日付', '休日名', '有効']), rows=sheet.getDataRange().getValues(); for(let i=1;i<rows.length;i++) if(formatDate_(rows[i][0])===input.date) sheet.getRange(i+1,3).setValue(false);audit_(book, '休日無効化', '', '', input.date, ''); }
    else if (input.action === 'updateRecord') {
      if (!input.id || !input.date || !TYPES.includes(input.type)) throw new Error('修正内容が不正です');
      const rows=records.getDataRange().getValues(); let found=-1;
      for(let i=1;i<rows.length;i++) if(String(rows[i][0])===input.id && rows[i][5]!==true) found=i;
      if(found<0) throw new Error('対象の打刻が見つかりません');
      const before=`${formatDate_(rows[found][2])} ${rows[found][3]} ${Utilities.formatDate(new Date(rows[found][4]),'Asia/Tokyo','HH:mm')}`;
      records.getRange(found+1,3,1,3).setValues([[input.date,input.type,new Date(Number(input.at))]]);
      audit_(book, '打刻修正', input.id, rows[found][1], before, `${input.date} ${input.type} ${Utilities.formatDate(new Date(Number(input.at)),'Asia/Tokyo','HH:mm')}`);
    }
    else if (input.action === 'deleteRecord') {
      const rows=records.getDataRange().getValues(); let found=-1;
      for(let i=1;i<rows.length;i++) if(String(rows[i][0])===input.id && rows[i][5]!==true) found=i;
      if(found<0) throw new Error('対象の打刻が見つかりません');
      records.getRange(found+1,6).setValue(true);
      audit_(book, '打刻削除', input.id, rows[found][1], `${formatDate_(rows[found][2])} ${rows[found][3]}`, '削除済');
    }
    else throw new Error('未対応の操作です');
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function ensureSheet_(book, name, header) {
  const sheet = book.getSheetByName(name) || book.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(header);
  const current=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),header.length)).getValues()[0];
  header.forEach((title,index)=>{if(!current[index])sheet.getRange(1,index+1).setValue(title);});
  return sheet;
}

function audit_(book, action, id, name, before, after) {
  ensureSheet_(book, HISTORY_SHEET, ['記録日時', '操作', '対象ID', '対象者', '変更前', '変更後'])
    .appendRow([new Date(), action, id, name, before, after]);
}

function formatDate_(value) {
  return Utilities.formatDate(new Date(value), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function parseSites_(value) {
  try { const sites=JSON.parse(value || '[]'); return Array.isArray(sites) ? sites : []; } catch (_) { return []; }
}
