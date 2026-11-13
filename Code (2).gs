/**
 * HTP 구약프레임 - 창세기 읽기 기록 앱
 * 이 코드는 구글 시트 「창세기 읽기 모임 읽기표」에 붙여서 씁니다.
 *
 * 시트에 새로 생기는 탭
 *   앱기록   : 이름별 창세기 1~50장 읽은 횟수 (앱이 여기에 기록해요)
 *   앱로그   : 누가 언제 몇 장을 읽었는지 한 줄씩 쌓여요
 *   출석숙제 : 1~8주 출석(O/X는 총무님이 입력)과 숙제(앱에서 체크)
 *   설정     : B1 칸 = 이번 주 읽기 카드 번호 (1 창조 ~ 8 요셉)
 */

const NAMES = ['강다솜', '김미애', '김미현', '김미희', '김소순', '김은희', '김효은', '유경희',
  '윤희원', '이선미', '이승진', '이영미', '이한나', '정소영', '한재윤', '샤론맘'];
const TAB_REC = '앱기록';
const TAB_LOG = '앱로그';
const TAB_ATT = '출석숙제';
const TAB_SET = '설정';
const WEEKS = 8;

/** 웹 앱 주소로 들어오면 화면을 보여줘요. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HTP 구약프레임 - 창세기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    // 구약프레임 아이콘 입구 페이지(GitHub) 안에서도 열리게 허용해요
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ★ 처음 한 번만 실행하세요.
 * 기존 이름별 탭의 '읽은 횟수' 줄을 읽어서 '앱기록' 탭으로 옮겨요. 기존 탭은 바꾸지 않아요.
 */
function 처음설정() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(TAB_REC)) {
    throw new Error("이미 '앱기록' 탭이 있어요. 다시 옮기려면 '앱기록' 탭을 지우고 실행하세요.");
  }
  const rec = ss.insertSheet(TAB_REC);
  const head = ['이름'];
  for (let i = 1; i <= 50; i++) head.push(i + '장');
  head.push('마지막 기록');
  const rows = [head];
  const report = [];
  NAMES.forEach(function (name) {
    let counts = new Array(50).fill(0);
    const sh = ss.getSheetByName(name);
    if (sh) {
      const colA = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      const r = colA.findIndex(function (v) { return String(v[0]).trim() === '읽은 횟수'; });
      if (r >= 0) {
        counts = sh.getRange(r + 1, 2, 1, 50).getValues()[0].map(function (v) { return Number(v) || 0; });
      }
    }
    report.push(name + ' ' + counts.reduce(function (a, b) { return a + b; }, 0) + '장');
    rows.push([name].concat(counts, ['']));
  });
  rec.getRange(1, 1, rows.length, head.length).setValues(rows);
  rec.setFrozenRows(1);
  rec.setFrozenColumns(1);

  const log = ss.getSheetByName(TAB_LOG) || ss.insertSheet(TAB_LOG);
  if (log.getLastRow() === 0) log.appendRow(['시각', '이름', '장', '변화']);

  if (!ss.getSheetByName(TAB_ATT)) {
    const att = ss.insertSheet(TAB_ATT);
    const h = ['이름'];
    for (let w = 1; w <= WEEKS; w++) h.push(w + '주 출석');
    for (let w = 1; w <= WEEKS; w++) h.push(w + '주 숙제');
    const ar = [h].concat(NAMES.map(function (n) {
      return [n].concat(new Array(WEEKS).fill(''), new Array(WEEKS).fill(false));
    }));
    att.getRange(1, 1, ar.length, h.length).setValues(ar);
    att.getRange(2, 2 + WEEKS, NAMES.length, WEEKS).insertCheckboxes();
    att.setFrozenRows(1);
    att.setFrozenColumns(1);
  }

  if (!ss.getSheetByName(TAB_SET)) {
    const set = ss.insertSheet(TAB_SET);
    set.getRange('A1:B1').setValues([['이번 주 읽기 카드 번호 (1 창조, 2 타락, 3 홍수, 4 노아의 후손들, 5 아브라함, 6 이삭, 7 야곱, 8 요셉)', 3]]);
    set.setColumnWidth(1, 520);
  }
  Logger.log('옮긴 기록: ' + report.join(', '));
}

/** 화면이 처음 열릴 때와 1분마다 최신 기록을 가져가요. */
function getData() {
  const ss = SpreadsheetApp.getActive();
  const rec = ss.getSheetByName(TAB_REC);
  if (!rec) throw new Error("'처음설정'을 먼저 실행해 주세요.");
  const rows = rec.getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]).trim(); });

  const attMap = {};
  const attSh = ss.getSheetByName(TAB_ATT);
  if (attSh && attSh.getLastRow() > 1) {
    attSh.getDataRange().getValues().slice(1).forEach(function (r) {
      const name = String(r[0]).trim();
      if (!name) return;
      attMap[name] = {
        att: r.slice(1, 1 + WEEKS).map(toAtt),
        hw: r.slice(1 + WEEKS, 1 + 2 * WEEKS).map(function (v) { return v === true || /^(o|O|○|✓|v|V)$/.test(String(v).trim()) ? 1 : 0; })
      };
    });
  }
  const setSh = ss.getSheetByName(TAB_SET);
  const now = setSh ? Number(setSh.getRange('B1').getValue()) : 3;

  return {
    now: now >= 1 && now <= 8 ? now : 3,
    members: rows.map(function (r) {
      const name = String(r[0]).trim();
      const a = attMap[name] || { att: new Array(WEEKS).fill(null), hw: new Array(WEEKS).fill(0) };
      return { name: name, c: r.slice(1, 51).map(function (v) { return Math.max(0, Number(v) || 0); }), att: a.att, hw: a.hw };
    })
  };
}

function toAtt(v) {
  const t = String(v).trim();
  if (v === true || /^(o|O|○|출석)$/.test(t)) return 1;
  if (/^(x|X|×|결석)$/.test(t)) return 0;
  return null;
}

/** 읽은 장을 기록해요. delta 는 1(읽음) 또는 -1(되돌리기). */
function addReads(name, chapters, delta) {
  delta = Number(delta);
  if (delta !== 1 && delta !== -1) throw new Error('잘못된 요청이에요.');
  const chs = (chapters || []).map(Number).filter(function (c) { return c >= 1 && c <= 50 && Math.floor(c) === c; });
  if (!chs.length || chs.length > 50) throw new Error('잘못된 장 번호예요.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActive();
    const rec = ss.getSheetByName(TAB_REC);
    const row = findRow(rec, name);
    const range = rec.getRange(row, 2, 1, 50);
    const vals = range.getValues()[0].map(function (v) { return Number(v) || 0; });
    chs.forEach(function (c) { vals[c - 1] = Math.max(0, vals[c - 1] + delta); });
    range.setValues([vals]);
    const stamp = new Date();
    rec.getRange(row, 52).setValue(stamp);
    const log = ss.getSheetByName(TAB_LOG);
    if (log) {
      const lines = chs.map(function (c) { return [stamp, name, c, delta]; });
      log.getRange(log.getLastRow() + 1, 1, lines.length, 4).setValues(lines);
    }
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** 내 숙제를 체크하거나 취소해요. week 는 0~7. */
function setHw(name, week, value) {
  week = Number(week);
  if (!(week >= 0 && week < WEEKS)) throw new Error('잘못된 주차예요.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(TAB_ATT);
    const row = findRow(sh, name);
    sh.getRange(row, 2 + WEEKS + week).setValue(Number(value) === 1);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

function findRow(sheet, name) {
  if (!sheet) throw new Error('기록 탭이 없어요.');
  const names = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const i = names.findIndex(function (v) { return String(v[0]).trim() === String(name).trim(); });
  if (i < 1) throw new Error('등록되지 않은 이름이에요.');
  return i + 1;
}
