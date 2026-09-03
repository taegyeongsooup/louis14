/**
 * 짐이 곧 국가다 - Google Sheets 공용 랭킹 서버
 * Google 스프레드시트 > 확장 프로그램 > Apps Script 에 붙여넣어 사용합니다.
 *
 * 시트 구조
 * - Users: 학생별 누적 요약
 * - Runs : 플레이 1회당 1행
 */

const USERS_SHEET = 'Users';
const RUNS_SHEET = 'Runs';

const USER_HEADERS = [
  'studentNo','name','pinHash','totalRuns','completedRuns','bestScore','totalScore',
  'titlesJson','aOrBetterStylesJson','createdAt','updatedAt'
];

const RUN_HEADERS = [
  'runId','studentNo','name','score','grade','completed','turn','statsJson','unrest',
  'similarity','similarityReached','style','titleIdsJson','cause','createdAt'
];

function setup() {
  ensureSheets_();
}

function doGet() {
  ensureSheets_();
  return json_({ok:true, service:'짐이 곧 국가다 랭킹 서버'});
}

function doPost(e) {
  try {
    ensureSheets_();
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(data.action || '');

    if (action === 'login') return json_(handleLogin_(data));
    if (action === 'submitRun') return json_(handleSubmitRun_(data));
    if (action === 'rankings') return json_(handleRankings_());

    return json_({ok:false, error:'알 수 없는 요청입니다.'});
  } catch (err) {
    return json_({ok:false, error:String(err && err.message ? err.message : err)});
  }
}

function handleLogin_(data) {
  const studentNo = clean_(data.studentNo);
  const name = clean_(data.name);
  const pinHash = clean_(data.pinHash);
  if (!studentNo || !name || !pinHash) return {ok:false,error:'로그인 정보가 부족합니다.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(USERS_SHEET);
    const found = findUser_(sheet, studentNo, name);

    if (!found) {
      const now = new Date().toISOString();
      sheet.appendRow([studentNo,name,pinHash,0,0,0,0,'[]','[]',now,now]);
    } else if (String(found.values[2]) !== pinHash) {
      return {ok:false,error:'PIN이 일치하지 않습니다.'};
    }

    return {ok:true, profile:buildProfile_(studentNo,name)};
  } finally {
    lock.releaseLock();
  }
}

function handleSubmitRun_(data) {
  const studentNo = clean_(data.studentNo);
  const name = clean_(data.name);
  const pinHash = clean_(data.pinHash);
  const record = data.record || {};
  if (!studentNo || !name || !pinHash || !record.runId) return {ok:false,error:'저장할 기록이 올바르지 않습니다.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = getSheet_(USERS_SHEET);
    const found = findUser_(users, studentNo, name);
    if (!found) return {ok:false,error:'등록되지 않은 사용자입니다.'};
    if (String(found.values[2]) !== pinHash) return {ok:false,error:'PIN이 일치하지 않습니다.'};

    const runs = getSheet_(RUNS_SHEET);
    if (!runExists_(runs, String(record.runId))) {
      runs.appendRow([
        String(record.runId), studentNo, name,
        num_(record.score), String(record.grade || ''), !!record.completed, num_(record.turn),
        JSON.stringify(record.stats || {}), num_(record.unrest),
        record.similarity == null ? '' : num_(record.similarity),
        num_(record.similarityReached), String(record.style || ''),
        JSON.stringify(Array.isArray(record.titleIds) ? record.titleIds : []),
        String(record.cause || ''), String(record.createdAt || new Date().toISOString())
      ]);

      const v = found.values;
      const totalRuns = num_(v[3]) + 1;
      const completedRuns = num_(v[4]) + (record.completed ? 1 : 0);
      const bestScore = Math.max(num_(v[5]), num_(record.score));
      const totalScore = num_(v[6]) + num_(record.score);
      const titles = union_(parseArray_(v[7]), Array.isArray(data.earnedTitleIds) ? data.earnedTitleIds : []);
      const styles = union_(parseArray_(v[8]), Array.isArray(data.aOrBetterStyles) ? data.aOrBetterStyles : []);
      const createdAt = v[9] || new Date().toISOString();
      const updatedAt = new Date().toISOString();

      users.getRange(found.row, 1, 1, USER_HEADERS.length).setValues([[
        studentNo,name,pinHash,totalRuns,completedRuns,bestScore,totalScore,
        JSON.stringify(titles),JSON.stringify(styles),createdAt,updatedAt
      ]]);
    }

    return {ok:true, profile:buildProfile_(studentNo,name)};
  } finally {
    lock.releaseLock();
  }
}

function handleRankings_() {
  const sheet = getSheet_(USERS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {ok:true,profiles:[]};

  const rows = sheet.getRange(2,1,lastRow-1,USER_HEADERS.length).getValues();
  const profiles = rows
    .filter(r => clean_(r[0]) && clean_(r[1]))
    .map(r => ({
      studentNo:String(r[0]),
      name:String(r[1]),
      totalRuns:num_(r[3]),
      completedRuns:num_(r[4]),
      bestScore:num_(r[5]),
      totalScore:num_(r[6]),
      titles:parseArray_(r[7]),
      aOrBetterStyles:parseArray_(r[8])
    }));
  return {ok:true,profiles};
}

function buildProfile_(studentNo, name) {
  const users = getSheet_(USERS_SHEET);
  const found = findUser_(users, studentNo, name);
  if (!found) throw new Error('사용자 정보를 찾을 수 없습니다.');
  const r = found.values;

  const runsSheet = getSheet_(RUNS_SHEET);
  const runs = [];
  const lastRow = runsSheet.getLastRow();
  if (lastRow >= 2) {
    const rows = runsSheet.getRange(2,1,lastRow-1,RUN_HEADERS.length).getValues();
    rows.forEach(x => {
      if (String(x[1]) !== String(studentNo) || String(x[2]) !== String(name)) return;
      runs.push({
        runId:String(x[0]), score:num_(x[3]), grade:String(x[4] || ''), completed:bool_(x[5]),
        turn:num_(x[6]), stats:parseObject_(x[7]), unrest:num_(x[8]),
        similarity:x[9] === '' ? null : num_(x[9]), similarityReached:num_(x[10]),
        style:String(x[11] || ''), titleIds:parseArray_(x[12]), cause:String(x[13] || ''),
        createdAt:String(x[14] || '')
      });
    });
  }

  return {
    studentNo:String(r[0]), name:String(r[1]), pinHash:String(r[2]),
    totalRuns:num_(r[3]), completedRuns:num_(r[4]), bestScore:num_(r[5]), totalScore:num_(r[6]),
    titles:parseArray_(r[7]), aOrBetterStyles:parseArray_(r[8]), runs:runs
  };
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('연결된 스프레드시트를 찾지 못했습니다. 스프레드시트의 확장 프로그램 > Apps Script에서 실행해 주세요.');
  ensureSheet_(ss, USERS_SHEET, USER_HEADERS);
  ensureSheet_(ss, RUNS_SHEET, RUN_HEADERS);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  return sheet;
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss && ss.getSheetByName(name);
  if (!sheet) throw new Error(name + ' 시트를 찾지 못했습니다. setup()을 먼저 실행해 주세요.');
  return sheet;
}

function findUser_(sheet, studentNo, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2,1,lastRow-1,USER_HEADERS.length).getValues();
  for (let i=0;i<rows.length;i++) {
    if (String(rows[i][0]) === String(studentNo) && String(rows[i][1]) === String(name)) {
      return {row:i+2, values:rows[i]};
    }
  }
  return null;
}

function runExists_(sheet, runId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const ids = sheet.getRange(2,1,lastRow-1,1).getValues();
  return ids.some(r => String(r[0]) === runId);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(v) { return String(v == null ? '' : v).trim(); }
function num_(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; }
function bool_(v) { return v === true || String(v).toLowerCase() === 'true'; }
function parseArray_(v) { try { const x = JSON.parse(String(v || '[]')); return Array.isArray(x) ? x : []; } catch (_) { return []; } }
function parseObject_(v) { try { const x = JSON.parse(String(v || '{}')); return x && typeof x === 'object' && !Array.isArray(x) ? x : {}; } catch (_) { return {}; } }
function union_(a,b) { return Array.from(new Set([].concat(a || [], b || []).map(String))); }
