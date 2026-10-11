// Card statement parser + category suggestions. Pure functions, used by the dialog
// (inlined into Index.html) and by the node tests.
var CardParse = (function () {
  var DATE_RE = /(\d{4})\D{1,3}(\d{1,2})\D{1,3}(\d{1,2})/;
  var SHORT_DATE_RE = /^(\d{2})[.\-\/](\d{1,2})[.\-\/](\d{1,2})$/;
  var SKIP_RE = /소계|합계|총\s*합|청\s*구\s*할\s*인|이월약정/;

  function cellText(v) {
    return v == null ? '' : String(v).replace(/\s+/g, ' ').trim();
  }

  function toNumber(v) {
    if (typeof v === 'number') return v;
    var s = cellText(v).replace(/[,원\s]/g, '');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    return Number(s);
  }

  function parseDate(v) {
    var s = cellText(v);
    var m = s.match(DATE_RE);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    m = s.match(SHORT_DATE_RE);
    if (m) return { y: 2000 + +m[1], m: +m[2], d: +m[3] };
    return null;
  }

  // Find the header row: needs a date column, a merchant column and an amount column.
  function findHeader(rows) {
    for (var r = 0; r < Math.min(rows.length, 40); r++) {
      var row = rows[r] || [];
      var col = { date: -1, card: -1, merchant: -1, amount: -1 };
      for (var c = 0; c < row.length; c++) {
        var t = cellText(row[c]).replace(/\s/g, '');
        if (col.date < 0 && /^(이용일|이용일자|거래일|거래일자|승인일|승인일자|사용일|사용일자|매출일)/.test(t)) col.date = c;
        else if (col.card < 0 && /(이용카드|카드명|카드구분|카드번호)/.test(t)) col.card = c;
        else if (col.merchant < 0 && /(가맹점|이용처|사용처|이용하신곳)/.test(t)) col.merchant = c;
        else if (col.amount < 0 && /^(이용금액|승인금액|거래금액|사용금액|금액)/.test(t)) col.amount = c;
      }
      if (col.date >= 0 && col.merchant >= 0 && col.amount >= 0) return { row: r, col: col };
    }
    return null;
  }

  /** rows: array of arrays (SheetJS sheet_to_json header:1). Returns transactions. */
  function parseRows(rows) {
    var h = findHeader(rows);
    if (!h) throw new Error('카드 내역 표를 찾지 못했어요 (이용일·가맹점·금액 칸이 필요해요)');
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
        bundle: /\d+건$/.test(merchant) && !/세/.test(merchant),
        y: date.y, m: date.m, d: date.d,
        card: h.col.card >= 0 ? cellText(row[h.col.card]) : '',
        merchant: merchant,
        amount: amount
      });
    }
    return out;
  }

  // Starting rules before anything has been learned. Only used when the category exists
  // in the sheet's list; learned merchant→category always wins.
  var RULES = [
    [/택시|지하철|버스|하이패스|도로|고속|대교|주유|충전소|철도|코레일|SRT|에어|항공|주차|파킹|카카오T|티머니/i, '교통비'],
    [/의원|병원|약국|의학과|치과|한의원|상담센터|클리닉|메디/, '병원비'],
    [/커피|카페|스타벅스|메가엠지씨|컴포즈|할리스|이디야|투썸|버거|맥도날드|롯데리아|치킨|피자|식당|마트|하나로|씨유|CU|GS25|세븐일레븐|이마트24|편의점|우아한형제들|배달|요기요|쿠팡이츠|아이스크림|베이커리|빵/i, '식비'],
    [/SKT|KT|LG\s*U|통신|지방세|세금|국세|관리비|가스|전기|수도|헤어|미용/i, '생활비'],
    [/메가박스|CGV|롯데시네마/i, '데이트']
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
      .replace(/\s*\d+건$/, '')
      .replace(/_\d+$/, '')
      .replace(/\d{1,2}월/g, '')
      .replace(/[\s-]*\*\*[\d*\-]*$/, '')
      .replace(/[\s\-_]+$/, '')
      .replace(/(\S+?)(대전|서울|부산|대구|광주|인천|울산|세종|수원|오송)\S*점$/, '$1')
      .trim();
  }

  // Many Korean card sites hand out "xls" files that are really HTML (sometimes EUC-KR, and
  // with sloppy tags like "</td   >" that trip SheetJS). Decode + tidy them first.
  function readRows(XLSX, buffer) {
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
    // Use the first sheet that has a recognisable transaction table.
    for (var s = 0; s < wb.SheetNames.length; s++) {
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[s]],
        { header: 1, raw: true, defval: '' });
      if (findHeader(rows)) return rows;
    }
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
  }

  return { readRows: readRows, parseRows: parseRows, suggest: suggest, normalizeMerchant: normalizeMerchant,
           parseDate: parseDate, toNumber: toNumber };
})();
if (typeof module !== 'undefined') module.exports = CardParse;
