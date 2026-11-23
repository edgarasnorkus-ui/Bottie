// Bottie: läser din budget (Excel-fil eller Google-kalkylblad i Drive) och skickar översikten till Bottie.
// Inställning (en gång):
//  1. Nytt Apps Script-projekt: script.google.com > Nytt projekt, klistra in hela den här filen.
//  2. Till vänster: Tjänster (+) > Drive API > välj version v3 > Lägg till.
//  3. Projektinställningar > Skriptegenskaper > Lägg till: BUDGET_KEY = ett eget långt lösenord.
//  4. Kör funktionen test() en gång och godkänn behörigheterna.
//  5. Distribuera > Ny distribution > Webbapp. Kör som: Jag. Åtkomst: Alla. Kopiera adressen som slutar på /exec.
var SOURCE_ID = '1pJQG3yg5pe0nWIfDOltlUCKjcRDPhlbb'; // din budgetfil i Drive (Excel eller Google-kalkylblad)
var MONTHS = ['Sausis', 'Vasaris', 'Kovas', 'Balandis', 'Gegužė', 'Birželis', 'Liepa', 'Rugpjūtis', 'Rugsėjis', 'Spalis', 'Lapkritis', 'Gruodis'];

function out(o) {
  return ContentService.createTextOutput(typeof o === 'string' ? o : JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var key = PropertiesService.getScriptProperties().getProperty('BUDGET_KEY');
  if (!key || p.key !== key) return out({ error: 'unauthorized' });
  var m = parseInt(p.month, 10);
  if (!(m >= 0 && m <= 11)) m = new Date().getMonth();
  var cache = CacheService.getScriptCache(), ck = 'budget' + m;
  if (!p.fresh) {
    var hit = cache.get(ck);
    if (hit) return out(hit);
  }
  try {
    var data = readMonth(m);
    var s = JSON.stringify(data);
    try { cache.put(ck, s, 300); } catch (x) {}
    return out(s);
  } catch (err) {
    return out({ error: String(err) });
  }
}

function openSource() {
  var f = DriveApp.getFileById(SOURCE_ID);
  if (f.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
    return { ss: SpreadsheetApp.openById(SOURCE_ID), tmp: null };
  }
  // Excel-fil: gör en tillfällig Google-kopia, läs den och släng den sedan.
  var copy = Drive.Files.create({ name: 'bottie-budget-tmp', mimeType: 'application/vnd.google-apps.spreadsheet' }, f.getBlob());
  return { ss: SpreadsheetApp.openById(copy.id), tmp: copy.id };
}

function readMonth(m) {
  var src = openSource();
  try {
    var sh = src.ss.getSheetByName(MONTHS[m]);
    if (!sh) return { error: 'Fliken "' + MONTHS[m] + '" finns inte', month: m, name: MONTHS[m] };
    var d = parseSheet(sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), 10).getValues());
    d.month = m;
    d.name = MONTHS[m];
    d.cur = 'SEK';
    d.updated = new Date().toISOString();
    return d;
  } finally {
    if (src.tmp) { try { DriveApp.getFileById(src.tmp).setTrashed(true); } catch (x) {} }
  }
}

function num(v) {
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined) return 0;
  var s = String(v).replace(/[−–]/g, '-').replace(/SEK|kr/gi, '').replace(/[\s ]/g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// vals: rader A:J. En kategori = rubrikrad (bara kolumn A) följd av en rad där F börjar med "Suplanuotos",
// sedan poster (A = namn, F = planerat, H = faktiskt) till raden "VISO:".
function parseSheet(vals) {
  var sections = [], cur = null, totals = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i], a = String(r[0] === undefined ? '' : r[0]).trim();
    var next = vals[i + 1];
    var onlyA = a !== '' && r.slice(1).every(function (c) { return c === '' || c === null; });
    if (!cur && onlyA && next && String(next[5]).indexOf('Suplanuotos') === 0) {
      cur = { n: a, items: [] };
      i++; // hoppa över kolumnrubrikerna
      continue;
    }
    if (a === 'VISO:') {
      if (cur) {
        cur.p = round2(cur.items.reduce(function (s, x) { return s + x.p; }, 0));
        cur.a = round2(cur.items.reduce(function (s, x) { return s + x.a; }, 0));
        sections.push(cur);
        cur = null;
      } else {
        totals.push(num(r[5]));
      }
      continue;
    }
    if (cur && a !== '' && a !== '-') {
      cur.items.push({ n: a, p: num(r[5]), a: num(r[7]) });
    }
  }
  var planned = round2(sections.reduce(function (s, x) { return s + x.p; }, 0));
  var actual = round2(sections.reduce(function (s, x) { return s + x.a; }, 0));
  var income = totals.length ? totals[0] : 0;   // första VISO: före kategorierna = pajamos
  var savings = totals.length > 1 ? totals[1] : 0; // andra = månadens sparande
  return { income: income, savings: savings, planned: planned, actual: actual, left: round2(income - savings - actual), sections: sections };
}

function round2(n) { return Math.round(n * 100) / 100; }

// Kör en gång i editorn för att godkänna behörigheter och se att det fungerar.
function test() {
  var d = readMonth(new Date().getMonth());
  Logger.log(JSON.stringify(d).slice(0, 1500));
}
