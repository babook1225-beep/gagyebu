// Card statement parser + category suggestions. Pure functions, used by the dialog
// (inlined into Index.html) and by the node tests.
var CardParse = (function () {
  var DATE_RE = /(\d{4})\D{1,3}(\d{1,2})\D{1,3}(\d{1,2})/;
  var SHORT_DATE_RE = /^(\d{2})[.\-\/](\d{1,2})[.\-\/](\d{1,2})$/;
  var COMPACT_DATE_RE = /^(20\d{2})(\d{2})(\d{2})/;     // 삼성카드: 20260901
  var SKIP_RE = /소계|합계|총\s*합|청\s*구\s*할\s*인|이월약정/;

  function cellText(v) {
    return v == null ? '' : String(v).replace(/\s+/g, ' ').trim();
  }

  function toNumber(v) {
    if (typeof v === 'number') return v;
    // 삼성카드는 "(10,780)", 취소는 "(-25,630)" 처럼 괄호로 감싸서 줌
    var s = cellText(v).replace(/[,원\s()]/g, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    return Number(s);
  }

  function parseDate(v) {
    var s = cellText(v);
    var m = s.match(COMPACT_DATE_RE);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    m = s.match(DATE_RE);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    m = s.match(SHORT_DATE_RE);
    if (m) return { y: 2000 + +m[1], m: +m[2], d: +m[3] };
    return null;
  }

  // Find the header row: needs a date column, a merchant column and an amount column.
  function findHeader(rows) {
    for (var r = 0; r < Math.min(rows.length, 40); r++) {
      var row = rows[r] || [];
      var col = { date: -1, card: -1, merchant: -1, amount: -1, out: -1, inn: -1, memo: -1 };
      for (var c = 0; c < row.length; c++) {
        var t = cellText(row[c]).replace(/\s/g, '');
        if (col.date < 0 && /^(이용일|이용일자|거래일|거래일자|거래일시|승인일|승인일자|사용일|사용일자|매출일)/.test(t)) col.date = c;
        else if (col.card < 0 && /(이용카드|카드명|카드구분|카드번호|이용구분|^적요$|^거래구분$)/.test(t)) col.card = c;
        else if (col.merchant < 0 && /(가맹점|이용처|사용처|이용하신곳|보낸분|받는분|기재내용|^내용$|거래내용)/.test(t)) col.merchant = c;
        else if (col.out < 0 && /^(출금액|출금금액|찾으신금액|출금)$/.test(t)) col.out = c;
        else if (col.inn < 0 && /^(입금액|입금금액|맡기신금액|입금)$/.test(t)) col.inn = c;
        else if (col.memo < 0 && /^(송금메모|메모)$/.test(t)) col.memo = c;
        else if (col.amount < 0 && /^(이용금액|승인금액|거래금액|사용금액|금액)/.test(t)) col.amount = c;
      }
      // 은행 거래내역: 금액이 출금액/입금액 두 칸으로 나뉨
      col.bank = col.out >= 0 && col.inn >= 0;
      if (col.date >= 0 && col.merchant >= 0 && (col.amount >= 0 || col.bank)) return { row: r, col: col };
    }
    return null;
  }

  /** rows: array of arrays (SheetJS sheet_to_json header:1). Returns transactions. */
  function parseRows(rows) {
    var h = findHeader(rows);
    if (!h) throw new Error('카드 내역 표를 찾지 못했어요 (이용일·가맹점·금액 칸이 필요해요)');
    if (h.col.bank) return parseBankRows(rows, h);
    var out = [];
    for (var r = h.row + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var date = parseDate(row[h.col.date]);
      var merchant = cellText(row[h.col.merchant]);
      var amount = toNumber(row[h.col.amount]);
      if (!date || !merchant || amount == null) continue;
      if (SKIP_RE.test(merchant)) continue;
      if (amount === 0) continue;               // 연회비 0원, 할인 표시용 줄 등
      out.push({
        // "08월티머니지하철 0003건", "고속도로 0002건": a month's rides billed as one line.
        // 삼성카드는 "M티머니통행료 0001건 (20260917~20260917)" 처럼 기간이 뒤에 붙음
        bundle: /\d+건(\s*\([^)]*\))?$/.test(merchant) && !/세/.test(merchant),
        y: date.y, m: date.m, d: date.d,
        card: h.col.card >= 0 ? cellText(row[h.col.card]) : '',
        merchant: merchant,
        amount: amount
      });
    }
    return out;
  }

  // 카드대금, 내 계좌·저축으로 옮긴 돈, 대출 상환은 지출이 아니거나 이미 카드내역으로 들어가니
  // 기본으로 빼 둠 (체크하면 넣을 수 있음).
  var CARD_BILL_RE = /카드|국민카드|비씨|BC|KB국민|삼성카드|현대카드|신한카드|롯데카드|하나카드|우리카드|NH카드/i;
  var SAVING_RE = /청약|적금|저축|도약|펀드|증권|투자|연금저축|IRP|ISA|MMF/i;
  var LOAN_RE = /기일출금|대출|원리금|이자출금/;
  var PERSON_RE = /^[가-힣]{2,4}$/;
  var TRANSFER_KIND_RE = /이체|출금|송금|입금|자동/;

  function bankFlag(kind, name, memo, dir) {
    if (dir === 'out' && CARD_BILL_RE.test(name) && !/카드깡/.test(name)) return 'cardbill';
    if (LOAN_RE.test(kind) || LOAN_RE.test(name) || /금리/.test(memo)) return 'loan';
    if (SAVING_RE.test(name)) return 'saving';
    if (PERSON_RE.test(name) && !/모임|회$|계$/.test(name) && TRANSFER_KIND_RE.test(kind)) return 'transfer';
    return '';
  }

  function parseBankRows(rows, h) {
    var out = [];
    for (var r = h.row + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var date = parseDate(row[h.col.date]);
      if (!date) continue;
      var name = cellText(row[h.col.merchant]);
      var kind = h.col.card >= 0 ? cellText(row[h.col.card]) : '';
      var memo = h.col.memo >= 0 ? cellText(row[h.col.memo]) : '';
      var o = toNumber(row[h.col.out]) || 0, i = toNumber(row[h.col.inn]) || 0;
      if (!o && !i) continue;
      var dir = o ? 'out' : 'in';
      if (!name) name = kind || memo;
      if (!name || SKIP_RE.test(name)) continue;
      out.push({
        bank: true, income: dir === 'in', flag: bankFlag(kind, name, memo, dir),
        y: date.y, m: date.m, d: date.d,
        card: kind + (memo && memo !== name ? ' · ' + memo : ''),
        merchant: name,
        amount: o || i
      });
    }
    return out;
  }

  // A payment and its cancellation in the same file ("쿠팡 25,630" + "쿠팡 -25,630") cancel
  // out: mark both so the screen can leave them unchecked.
  function markCancelled(tx) {
    tx.forEach(function (t, i) {
      if (t.amount >= 0 || t.cancelled) return;
      for (var j = 0; j < tx.length; j++) {
        var o = tx[j];
        if (j !== i && !o.cancelled && o.amount === -t.amount && o.merchant === t.merchant) {
          o.cancelled = true; t.cancelled = true; break;
        }
      }
    });
    return tx;
  }

  // 어느 카드사·은행 파일인지 (시트 칸 색깔용). 파일 이름 → 표 위쪽 안내 줄 → 전체 순서로 찾음.
  var ORIGINS = [
    [/현대카드|hyundai\s*card|hyundaicard/i, '현대카드'],
    [/삼성카드|samsung\s*card|samsungcard/i, '삼성카드'],
    [/KB국민카드|국민카드/i, '국민카드'],
    [/신한카드/, '신한카드'], [/롯데카드/, '롯데카드'], [/하나카드/, '하나카드'], [/우리카드/, '우리카드'],
    [/NH농협카드|농협카드/, '농협카드'], [/비씨카드|BC카드/i, 'BC카드']
  ];
  var BANK_ORIGINS = [
    [/국민은행|KB|ＫＢ|스타뱅킹/i, '국민은행'], [/신한은행|신한/, '신한은행'], [/우리은행|우리/, '우리은행'],
    [/하나은행|하나/, '하나은행'], [/농협|NH/i, '농협'], [/카카오뱅크|카카오/, '카카오뱅크'], [/토스뱅크|토스/, '토스뱅크']
  ];
  var ORIGIN_NAMES = ORIGINS.map(function (o) { return o[1]; }).concat(BANK_ORIGINS.map(function (o) { return o[1]; }));

  function detectOrigin(texts, isBank, compactDates) {
    var list = isBank ? BANK_ORIGINS : ORIGINS;
    for (var t = 0; t < texts.length; t++) {
      for (var i = 0; i < list.length; i++) if (list[i][0].test(texts[t])) return list[i][1];
    }
    if (!isBank && compactDates) return '삼성카드';   // 삼성카드만 20260901 형식 날짜를 씀
    return isBank ? '은행' : '카드';
  }

  /** All transactions in a card file (every sheet that has a transaction table). */
  function parseFile(XLSX, buffer, fileName) {
    var info = {};
    var sheets = readSheets(XLSX, buffer, info);
    var tx = [];
    var top = [], all = [];
    sheets.forEach(function (rows) {
      tx = tx.concat(parseRows(rows));
      var h = findHeader(rows);
      rows.forEach(function (r, i) { (i <= h.row ? top : all).push((r || []).map(cellText).join(' ')); });
    });
    var isBank = tx.some(function (t) { return t.bank; });
    var compact = sheets.some(function (rows) {
      var h = findHeader(rows);
      return rows.slice(h.row + 1, h.row + 6).some(function (r) { return COMPACT_DATE_RE.test(cellText((r || [])[h.col.date])); });
    });
    var origin = detectOrigin([fileName || '', (info.sheetNames || []).join(' ') + ' ' + top.join(' '), all.join(' ')], isBank, compact);
    tx.forEach(function (t) { t.origin = origin; });
    if (!sheets.length) throw new Error('내역 표를 찾지 못했어요 (날짜·가맹점(받는분)·금액 칸이 필요해요)');
    tx.sort(function (a, b) { return (a.y - b.y) || (a.m - b.m) || (a.d - b.d); });
    return markCancelled(tx.filter(function (t) { return !t.bank; })).concat(tx.filter(function (t) { return t.bank; }))
      .sort(function (a, b) { return (a.y - b.y) || (a.m - b.m) || (a.d - b.d); });
  }

  // Starting rules before anything has been learned. Only used when the category exists
  // in the sheet's list; learned merchant→category always wins.
  var RULES = [
    [/택시|지하철|버스|하이패스|도로|고속|대교|주유|충전소|철도|코레일|SRT|에어|항공|주차|파킹|카카오T|티머니/i, '교통비'],
    [/의원|병원|약국|의학과|치과|한의원|상담센터|클리닉|메디/, '병원비'],
    [/커피|카페|스타벅스|메가엠지씨|컴포즈|할리스|이디야|투썸|버거|맥도날드|롯데리아|치킨|피자|식당|마트|하나로|씨유|CU|GS25|세븐일레븐|이마트24|편의점|우아한형제들|배달|요기요|쿠팡이츠|아이스크림|베이커리|빵/i, '식비'],
    [/SKT|KT|LG\s*U|통신|지방세|세금|국세|관리비|가스|전기|수도|헤어|미용/i, '생활비'],
    [/메가박스|CGV|롯데시네마/i, '데이트'],
    [/손해|손보|화재|생명|라이나|AIA|메트라이프|^KB손|^DB손|^미래\d|보험/i, '생활비'],
    [/모임|계비|회비/, '계비'],
    // 입금
    [/급여|월급|상여/, '월급'],
    [/이자|결산/, '이자'],
    [/지원금|환급|장려금/, '지원금']
  ];

  function suggest(merchant, learned, categories) {
    var key = normalizeMerchant(merchant);
    if (learned && learned[key] && categories.indexOf(learned[key]) >= 0) {
      return { category: learned[key], source: 'learned' };
    }
    for (var i = 0; i < RULES.length; i++) {
      if (RULES[i][0].test(merchant) && categories.indexOf(RULES[i][1]) >= 0) {
        return { category: RULES[i][1], source: 'rule' };
      }
    }
    return { category: '', source: '' };
  }

  // "OO커피서울역점" and "OO커피부산역점" should share what was learned,
  // but "네이버파이낸셜(주) - 항공사" vs "- 영화관" should not.
  function normalizeMerchant(s) {
    return cellText(s)
      .replace(/\s*\([^)]*\)\s*$/, '')
      .replace(/\s*\d+건$/, '')
      .replace(/_\d+$/, '')
      .replace(/([^\d\s])\d{4,}$/, '$1')        // 은행 CMS: "보험사07509" → 매달 바뀌는 번호 떼기
      .replace(/\d{1,2}월/g, '')
      .replace(/[\s-]*\*\*[\d*\-]*$/, '')
      .replace(/[\s\-_]+$/, '')
      .replace(/(\S+?)(대전|서울|부산|대구|광주|인천|울산|세종|수원|오송)\S*점$/, '$1')
      .trim();
  }

  // Many Korean card sites hand out "xls" files that are really HTML (sometimes EUC-KR, and
  // with sloppy tags like "</td   >" that trip SheetJS). Decode + tidy them first.
  function readSheets(XLSX, buffer, info) {
    var bytes = new Uint8Array(buffer);
    var head = '';
    for (var i = 0; i < Math.min(bytes.length, 2048); i++) head += String.fromCharCode(bytes[i]);
    var isHtml = /<\s*(!doctype|html|table|meta)/i.test(head);
    var wb;
    if (isHtml) {
      var charset = (head.match(/charset\s*=\s*["']?([\w-]+)/i) || [])[1] || 'utf-8';
      var text;
      try { text = new TextDecoder(charset.toLowerCase()).decode(bytes); }
      catch (e) { text = new TextDecoder('utf-8').decode(bytes); }
      if (/�/.test(text.slice(0, 5000)) && !/euc-kr/i.test(charset)) {
        try { text = new TextDecoder('euc-kr').decode(bytes); } catch (e) { }
      }
      text = text.replace(/<\/(td|th|tr|table)\s+>/gi, '</$1>');
      wb = XLSX.read(text, { type: 'string' });
    } else {
      wb = XLSX.read(bytes, { type: 'array' });
    }
    // Every sheet that has a recognisable transaction table (일시불 / 할부 탭 등).
    var found = [];
    if (info) info.sheetNames = wb.SheetNames;
    for (var s = 0; s < wb.SheetNames.length; s++) {
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[s]],
        { header: 1, raw: true, defval: '' });
      if (findHeader(rows)) found.push(rows);
    }
    return found;
  }

  function readRows(XLSX, buffer) {
    var all = readSheets(XLSX, buffer);
    return all.length ? all[0] : [];
  }

  return { parseFile: parseFile, ORIGIN_NAMES: ORIGIN_NAMES, detectOrigin: detectOrigin, readRows: readRows, parseRows: parseRows, suggest: suggest, normalizeMerchant: normalizeMerchant,
           parseDate: parseDate, toNumber: toNumber };
})();
if (typeof module !== 'undefined') module.exports = CardParse;
