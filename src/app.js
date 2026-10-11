/**
 * 카드내역 → 가계부
 *
 * 카드사 이용내역 엑셀을 올리면 목록으로 보여주고, 적요를 고른 뒤 완료하면
 * 이용일 기준으로 해당 월 탭의 날짜 칸에 넣습니다.
 *  - 같은 날 같은 적요가 있으면 금액 칸에 더하기 (=a+b)
 *  - 없으면 그날 빈 줄에 새로 넣기
 *  - 같은 금액이 ±2일 안에 이미 있으면 "이미 있음": 체크하면 새로 넣지 않고 그 항목의 적요·메모를 수정
 *  - 메모는 금액 칸의 메모로 (더할 때는 기존 메모 뒤에 +메모)
 *  - 가맹점별로 고른 적요·메모를 기억해서 다음에 자동으로 채움
 *
 * GitHub에서 불러와 실행됩니다 (시트에는 Loader.gs 하나만 붙여넣음).
 * Host.ss() 는 Loader 가 넘겨주는 가계부 스프레드시트.
 */


var LEARNED_KEY = 'cardImport.learned';
var MEMO_KEY = 'cardImport.memo';

function ss_() { return Host.ss(); }

function getLearned_(key) {
  var raw = PropertiesService.getDocumentProperties().getProperty(key || LEARNED_KEY);
  try { return raw ? JSON.parse(raw) : {}; } catch (e) { return {}; }
}

function saveLearned_(map, key) {
  PropertiesService.getDocumentProperties().setProperty(key || LEARNED_KEY, JSON.stringify(map));
}

/** 가장 최근 월 탭의 적요 목록 (BS14:BS28). */
function getCategories_(ss) {
  var tabs = ss.getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return /^\d{4}\.\d{2}$/.test(n); }).sort();
  if (!tabs.length) return [];
  var sh = ss.getSheetByName(tabs[tabs.length - 1]);
  return sh.getRange(Layout.CAT_RANGE).getDisplayValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return v; });
}

function readGrid_(sh) {
  var nRows = Layout.EXP_LAST_ROW - Layout.EXP_FIRST_ROW + 1;
  var nCols = Layout.LAST_COL - Layout.FIRST_COL + 1;
  var range = sh.getRange(Layout.EXP_FIRST_ROW, Layout.FIRST_COL, nRows, nCols);
  return {
    cats: range.getDisplayValues(),
    vals: range.getValues(),
    forms: range.getFormulas(),
    notes: range.getNotes()
  };
}

/** 화면이 처음 열릴 때: 적요 목록 + 기억해 둔 가맹점 적요. */
function getSetup() {
  var ss = ss_();
  return { categories: getCategories_(ss), learned: getLearned_(), memos: getLearned_(MEMO_KEY), title: ss.getName() };
}

/**
 * 올린 내역에 대해 탭이 있는지, 이미 입력된 것 같은지 확인.
 * items: [{ key, y, m, d, amount }] → [{ key, tab, tabMissing, dup }]
 */
function checkItems(items) {
  var ss = ss_();
  var byTab = groupByTab_(items);
  var out = [];
  Object.keys(byTab).forEach(function (tab) {
    var sh = ss.getSheetByName(tab);
    var list = byTab[tab];
    if (!sh) {
      list.forEach(function (it) { out.push({ key: it.key, tab: tab, tabMissing: true, dup: false }); });
      return;
    }
    Plan.markDuplicates(readGrid_(sh), list);
    list.forEach(function (it) {
      out.push({ key: it.key, tab: tab, tabMissing: false, dup: it.dup, dupDay: it.dupDay,
                 dupCat: it.dupCat || '', dupNote: it.dupNote || '', dupCount: it.dupCount || 0 });
    });
  });
  return out;
}

/**
 * 완료: 고른 항목을 시트에 넣고 가맹점 적요를 기억.
 * items: [{ key, y, m, d, amount, category, merchant, memo, mode: 'add'|'update' }]
 */
function commitItems(items) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var ss = ss_();
    var byTab = groupByTab_(items);
    var summary = { written: 0, merged: 0, updated: 0, moved: 0, same: 0, full: [], missingTab: [] };
    Object.keys(byTab).sort().forEach(function (tab) {
      var sh = ss.getSheetByName(tab);
      var list = byTab[tab];
      if (!sh) {
        list.forEach(function (it) { summary.missingTab.push(it.key); });
        return;
      }
      var grid = readGrid_(sh);
      // "이미 있음" kept checked → fix the existing entry instead of adding it again.
      var upd = Plan.planUpdates(grid, list.filter(function (it) { return it.mode === 'update'; }));
      upd.results.forEach(function (r) {
        if (r.status === 'updated') summary.updated++;
        else if (r.status === 'moved') summary.moved++;
        else if (r.status === 'same') summary.same++;
      });
      var adds = list.filter(function (it) { return it.mode !== 'update'; })
        .concat(upd.notFound, upd.moved)
        .sort(function (a, b) { return a.d - b.d || a.order - b.order; });
      var plan = Plan.planWrites(grid, adds);
      upd.writes.concat(plan.writes).forEach(function (w) {
        var cell = sh.getRange(w.row, w.col);
        if ('note' in w) cell.setNote(w.note); else cell.setValue(w.value);
      });
      var movedKeys = {};
      upd.moved.forEach(function (it) { movedKeys[it.key] = true; });
      plan.results.forEach(function (r) {
        var it = adds[r.idx];
        if (r.status === 'full') summary.full.push(it.key);
        else if (movedKeys[it.key]) return;           // already counted as moved
        else if (r.status === 'merged') summary.merged++;
        else summary.written++;
      });
    });
    SpreadsheetApp.flush();

    var learned = getLearned_();
    var memos = getLearned_(MEMO_KEY);
    items.forEach(function (it) {
      var k = CardParse.normalizeMerchant(it.merchant || '');
      if (!k) return;
      if (it.category) learned[k] = it.category;
      if (it.memo) memos[k] = String(it.memo).trim();
    });
    saveLearned_(learned);
    saveLearned_(memos, MEMO_KEY);
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function groupByTab_(items) {
  var by = {};
  items.forEach(function (it) {
    var tab = Layout.tabName(it.y, it.m);
    (by[tab] = by[tab] || []).push(it);
  });
  // 날짜 → 원래 순서대로
  Object.keys(by).forEach(function (t) {
    by[t].sort(function (a, b) { return a.d - b.d || a.order - b.order; });
  });
  return by;
}


var App = { getSetup: getSetup, checkItems: checkItems, commitItems: commitItems };
