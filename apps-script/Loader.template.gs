/**
 * 가계부 · 카드내역 가져오기 — 불러오기 코드
 *
 * 이 파일 하나만 Apps Script에 붙여넣어 두면, 실제 프로그램은 매번 GitHub
 * (github.com/babook1225-beep/gagyebu)에서 최신 버전을 받아 실행해요.
 * 그래서 프로그램을 고쳐도 여기는 다시 붙여넣을 필요가 없어요.
 *
 * 화면은 google.script.run.api(이름, [인자]) 하나로만 서버를 불러요.
 */

var GAGYEBU_REPO = 'https://raw.githubusercontent.com/babook1225-beep/gagyebu/main/dist/';
var GAGYEBU_SHEET_ID = '여기에_가계부_시트_ID';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('카드내역')
    .addItem('카드내역 가져오기', 'openImporter')
    .addToUi();
}

function openImporter() {
  var html = HtmlService.createHtmlOutput(gagyebuFetch_('Index.html'))
    .setWidth(1100).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '카드내역 가져오기');
}

// 폰: 배포 > 새 배포 > 웹 앱 으로 배포하면 같은 화면을 주소로 열 수 있어요.
function doGet() {
  return HtmlService.createHtmlOutput(gagyebuFetch_('Index.html'))
    .setTitle('카드내역 가져오기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 화면에서 부르는 유일한 입구. */
function api(name, args) {
  var App = gagyebuLoad_();
  if (typeof App[name] !== 'function') throw new Error('알 수 없는 기능: ' + name);
  return App[name].apply(null, args || []);
}

function gagyebuSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(GAGYEBU_SHEET_ID);
}

function gagyebuLoad_() {
  var src = ['Parser.js', 'Plan.js', 'App.js'].map(gagyebuFetch_).join('\n;\n');
  var Host = { ss: gagyebuSheet_ };
  return new Function('Host', src + '\nreturn App;')(Host);
}

/** GitHub에서 파일 받기. 실패하면 마지막으로 받았던 것(최대 6시간)을 씀. */
function gagyebuFetch_(file) {
  var cache = CacheService.getScriptCache();
  var key = 'gagyebu:' + file;
  try {
    var res = UrlFetchApp.fetch(GAGYEBU_REPO + file + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var text = res.getContentText('UTF-8');
      try { cache.put(key, text, 21600); } catch (e) { /* too big for cache: fine */ }
      return text;
    }
  } catch (e) { /* network: fall through to cache */ }
  var cached = cache.get(key);
  if (cached) return cached;
  throw new Error(file + ' 을(를) GitHub에서 받지 못했어요. 잠시 뒤 다시 해 주세요.');
}
