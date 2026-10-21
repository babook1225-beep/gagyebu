// Pure layout logic for the 가계부 monthly tabs. No SpreadsheetApp calls here, so it can be
// tested outside Apps Script.
//
// Monthly tab "YYYY.MM": day N uses two columns starting at F — [적요 chip, 금액].
// Expense slots are rows 5–14 (row 15 is the 합 row). Category list lives in BS14:BS28.
var Layout = {
  FIRST_COL: 6,        // F
  EXP_FIRST_ROW: 5,
  EXP_LAST_ROW: 14,
  INC_FIRST_ROW: 2,    // 수입 2–3행 (4행은 합)
  INC_LAST_ROW: 3,
  INC_CAT_RANGE: 'BS2:BS13',
  LAST_COL: 6 + 31 * 2 - 1,   // BO
  CAT_RANGE: 'BS14:BS28',
  tabName: function (y, m) { return y + '.' + (m < 10 ? '0' : '') + m; },
  catCol: function (d) { return 6 + (d - 1) * 2; },
  amtCol: function (d) { return 7 + (d - 1) * 2; },
  // 어디서 온 내역인지 칸 배경색으로 표시. 여러 곳에서 온 금액이 한 칸에 더해지면 MIXED.
  COLORS: {
    '현대카드': '#dbe8fd', '삼성카드': '#d9f2e1', '국민은행': '#fdf0c4',
    '국민카드': '#fde2cc', '신한카드': '#e3e0fb', '롯데카드': '#fbdde3', '하나카드': '#d4f0ee',
    '우리카드': '#d6ecfa', '농협카드': '#e9f3d2', 'BC카드': '#f8dcdc',
    '신한은행': '#e3e0fb', '우리은행': '#d6ecfa', '하나은행': '#d4f0ee', '농협': '#e9f3d2',
    '카카오뱅크': '#fff6b3', '토스뱅크': '#dde6f7'
  },
  MIXED: '#ececec',
  WHITE: '#ffffff',
  colorOf: function (origin) { return (origin && Layout.COLORS[origin]) || ''; }
};

var Plan = (function () {
  // Amount cell -> list of numbers it is made of ("=1000+2000" -> [1000, 2000]).
  function terms(value, formula) {
    if (formula && /^=\s*-?[\d,.]+(\s*[+\-]\s*[\d,.]+)*\s*$/.test(formula)) {
      var out = [];
      var re = /([+\-]?)\s*([\d,.]+)/g, m, first = true;
      var body = formula.replace(/^=\s*/, '');
      while ((m = re.exec(body))) {
        var n = Number(m[2].replace(/,/g, ''));
        out.push(m[1] === '-' ? -n : n);
        first = false;
      }
      return out;
    }
    if (typeof value === 'number' && value !== 0) return [value];
    var v = Number(String(value || '').replace(/,/g, ''));
    return v ? [v] : [];
  }

  function addTerm(value, formula, amount) {
    var part = amount < 0 ? '-' + Math.abs(amount) : '+' + amount;
    if (formula && formula.charAt(0) === '=') return formula + part;
    var cur = typeof value === 'number' ? value : Number(String(value || '').replace(/,/g, ''));
    if (!cur) return amount;
    return '=' + cur + part;
  }

  /**
   * grid: { cats[r][c], vals[r][c], forms[r][c] } for rows 5..14 × cols F..BO (0-based offsets).
   * items: [{ d, amount, category?, ... }]
   * Marks duplicates: same day, and the amount already appears in that day's amounts.
   */
  // Card "이용일" and the day people write things down often differ by a day or two,
  // so look on the same day first, then ±1, then ±2.
  var DAY_OFFSETS = [0, -1, 1, -2, 2];

  function markDuplicates(grid, items) {
    var used = {};
    items.forEach(function (it) {
      it.dup = false;
      it.dupDay = null;
      for (var o = 0; o < DAY_OFFSETS.length && !it.dup; o++) {
        var day = it.d + DAY_OFFSETS[o];
        if (day < 1 || day > 31) continue;
        var a = Layout.amtCol(day) - Layout.FIRST_COL;
        for (var r = 0; r < grid.vals.length && !it.dup; r++) {
          var ts = terms(grid.vals[r][a], grid.forms[r][a]);
          for (var i = 0; i < ts.length; i++) {
            var key = day + '/' + r + ':' + i;
            if (!used[key] && ts[i] === it.amount) {
              used[key] = true;
              it.dup = true;
              it.dupDay = day;
              it.dupRow = r;
              it.dupTerm = i;
              it.dupCount = ts.length;
              it.dupCat = String(grid.cats[r][a - 1] || '').trim();
              it.dupNote = grid.notes ? String(grid.notes[r][a] || '').trim() : '';
              break;
            }
          }
        }
      }
    });
    return items;
  }

  function isWhite(c) { return !c || /^#?f{6}$/i.test(String(c)); }

  // 칸 두 개(적요·금액)의 배경을 정함: 빈 칸이면 출처 색, 같은 색이면 그대로, 다르면 섞임 색.
  function paint(grid, writes, rowIdx, c, row, d, color, fresh) {
    if (!color || !grid.bgs) return;
    var a = c + 1;
    var cur = grid.bgs[rowIdx][a];
    var next;
    if (fresh) next = color;
    else if (!isWhite(cur) && String(cur).toLowerCase() === color.toLowerCase()) next = cur;
    else next = Layout.MIXED;
    if (String(next).toLowerCase() === String(cur || '').toLowerCase()) return;
    grid.bgs[rowIdx][a] = next; grid.bgs[rowIdx][c] = next;
    writes.push({ row: row, col: Layout.catCol(d), bg: next });
    writes.push({ row: row, col: Layout.amtCol(d), bg: next });
  }

  /**
   * Returns { writes: [{row, col, value}], results: [{idx, status, row}] }.
   * status: 'merged' (added to an existing same-category cell), 'new' (empty slot), 'full'.
   * grid is updated in place so later items see earlier writes.
   */
  function planWrites(grid, items) {
    var writes = [];
    var results = [];
    items.forEach(function (it, idx) {
      var c = Layout.catCol(it.d) - Layout.FIRST_COL;
      var a = c + 1;
      var target = -1, status = 'full';
      for (var r = 0; r < grid.cats.length; r++) {
        if (String(grid.cats[r][c]).trim() === it.category) { target = r; status = 'merged'; break; }
      }
      if (target < 0) {
        for (var r2 = 0; r2 < grid.cats.length; r2++) {
          var emptyCat = String(grid.cats[r2][c]).trim() === '';
          var emptyAmt = String(grid.vals[r2][a]).trim() === '' && !grid.forms[r2][a];
          if (emptyCat && emptyAmt) { target = r2; status = 'new'; break; }
        }
      }
      if (target < 0) { results.push({ idx: idx, status: 'full' }); return; }

      var row = (grid.firstRow || Layout.EXP_FIRST_ROW) + target;
      var memo = String(it.memo || '').trim();
      if (memo) {
        if (!grid.notes) grid.notes = grid.cats.map(function (r) { return r.map(function () { return ''; }); });
        var cur = String(grid.notes[target][a] || '').trim();
        // Same style as the sheet: "커피+점심+간식"
        var note = status === 'merged' && cur ? cur + '+' + memo : (cur && status === 'new' ? cur + '+' + memo : memo);
        grid.notes[target][a] = note;
        writes.push({ row: row, col: Layout.amtCol(it.d), note: note });
      }
      var color = Layout.colorOf(it.origin);
      // 새 칸이면 출처 색. 기존 칸에 더할 때: 손으로 적은 흰 칸이나 다른 출처면 섞임 색.
      paint(grid, writes, target, c, row, it.d, color, status === 'new' ? true : false);
      if (status === 'new') {
        writes.push({ row: row, col: Layout.catCol(it.d), value: it.category });
        writes.push({ row: row, col: Layout.amtCol(it.d), value: it.amount });
        grid.cats[target][c] = it.category;
        grid.vals[target][a] = it.amount;
        grid.forms[target][a] = '';
      } else {
        var nv = addTerm(grid.vals[target][a], grid.forms[target][a], it.amount);
        writes.push({ row: row, col: Layout.amtCol(it.d), value: nv });
        if (typeof nv === 'string') { grid.forms[target][a] = nv; }
        else { grid.vals[target][a] = nv; grid.forms[target][a] = ''; }
      }
      results.push({ idx: idx, status: status, row: row });
    });
    // Several items can land in the same cell; only the final value needs writing.
    var last = {};
    var kind = function (w) { return w.row + ':' + w.col + (('note' in w) ? ':n' : ('bg' in w) ? ':b' : ':v'); };
    writes.forEach(function (w, i) { last[kind(w)] = i; });
    writes = writes.filter(function (w, i) { return last[kind(w)] === i; });
    return { writes: writes, results: results };
  }

  // Rebuild an amount cell without one of its terms. Returns '' when nothing is left.
  function withoutTerm(ts, idx) {
    var rest = ts.filter(function (_, i) { return i !== idx; });
    if (!rest.length) return '';
    if (rest.length === 1) return rest[0];
    return '=' + rest.map(function (n, i) {
      return i === 0 ? String(n) : (n < 0 ? '-' + Math.abs(n) : '+' + n);
    }).join('');
  }

  /**
   * "이미 있음" items the user kept checked: fix the existing entry instead of adding again.
   *  - same 적요: amount stays; memo replaces the note when that cell holds only this amount
   *  - different 적요: take the amount out of the old cell and put it under the new 적요
   *    (same day as the existing entry)
   * Returns { writes, results: [{idx, status}], moved: [items to run through planWrites],
   *           notFound: [items no longer in the sheet → add normally] }.
   */
  function planUpdates(grid, items) {
    var writes = [], results = [], moved = [], notFound = [];
    var probe = items.map(function (it) { return { d: it.d, amount: it.amount }; });
    markDuplicates(grid, probe);
    items.forEach(function (it, idx) {
      var p = probe[idx];
      if (!p.dup) { notFound.push(it); results.push({ idx: idx, status: 'notfound' }); return; }
      var a = Layout.amtCol(p.dupDay) - Layout.FIRST_COL;
      var c = a - 1;
      var row = (grid.firstRow || Layout.EXP_FIRST_ROW) + p.dupRow;
      var memo = String(it.memo || '').trim();
      if (!it.category || it.category === p.dupCat) {
        // 손으로 적어 둔 걸 카드·은행 내역으로 확인했으면 그 칸에도 출처 색 (한 건짜리 칸만)
        if (p.dupCount === 1) paint(grid, writes, p.dupRow, c, row, p.dupDay, Layout.colorOf(it.origin), true);
        if (memo && p.dupCount === 1 && memo !== p.dupNote) {
          if (grid.notes) grid.notes[p.dupRow][a] = memo;
          writes.push({ row: row, col: Layout.amtCol(p.dupDay), note: memo });
          results.push({ idx: idx, status: 'updated' });
        } else {
          results.push({ idx: idx, status: 'same' });
        }
        return;
      }
      // Move to another 적요.
      var ts = terms(grid.vals[p.dupRow][a], grid.forms[p.dupRow][a]);
      var nv = withoutTerm(ts, p.dupTerm);
      writes.push({ row: row, col: Layout.amtCol(p.dupDay), value: nv });
      if (typeof nv === 'string' && nv) { grid.forms[p.dupRow][a] = nv; }
      else { grid.vals[p.dupRow][a] = nv; grid.forms[p.dupRow][a] = ''; }
      if (nv === '') {
        grid.cats[p.dupRow][c] = '';
        writes.push({ row: row, col: Layout.catCol(p.dupDay), value: '' });
        if (grid.notes) grid.notes[p.dupRow][a] = '';
        writes.push({ row: row, col: Layout.amtCol(p.dupDay), note: '' });
        if (grid.bgs && !isWhite(grid.bgs[p.dupRow][a])) {
          grid.bgs[p.dupRow][a] = grid.bgs[p.dupRow][c] = Layout.WHITE;
          writes.push({ row: row, col: Layout.catCol(p.dupDay), bg: Layout.WHITE });
          writes.push({ row: row, col: Layout.amtCol(p.dupDay), bg: Layout.WHITE });
        }
      }
      var copy = {};
      for (var k in it) copy[k] = it[k];
      copy.d = p.dupDay;
      moved.push(copy);
      results.push({ idx: idx, status: 'moved' });
    });
    return { writes: writes, results: results, moved: moved, notFound: notFound };
  }

  return { terms: terms, addTerm: addTerm, markDuplicates: markDuplicates, planWrites: planWrites,
           planUpdates: planUpdates, withoutTerm: withoutTerm };
})();
if (typeof module !== 'undefined') module.exports = { Plan: Plan, Layout: Layout };
