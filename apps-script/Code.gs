/**
 * Backs the Spanish flashcards deck with a Google Sheet.
 *
 * Setup:
 * 1. Create a Google Sheet. Rename its first tab to "Verbs" with header row:
 *    name | meaning | conj | examples
 * 2. Extensions > Apps Script, replace the default code with this file, save.
 * 3. Run seedInitialVerbs() once (select it in the toolbar dropdown, click Run)
 *    to authorize the script and populate the starter verbs.
 * 4. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the deployment's web app URL and paste it into APPS_SCRIPT_URL
 *    near the top of spanish-tense-flashcards.html.
 *
 * Re-deploy (Deploy > Manage deployments > edit > new version) any time you
 * change this file, or the live endpoint keeps running the old version.
 */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Verbs') || ss.getSheets()[0];
}

function safeParse_(str) {
  try { return JSON.parse(str); } catch (e) { return null; }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

var MAX_NAME_LENGTH = 60;
var MAX_MEANING_LENGTH = 200;
var MAX_CONJ_LENGTH = 4000;
var MAX_EXAMPLES_LENGTH = 4000;

/* Sheets treats a cell starting with =, +, -, or @ as a formula, even when
   set via the API — a leading apostrophe forces it to be stored as plain
   text instead (and the apostrophe itself isn't kept). Guards against a
   malicious verb name/meaning turning into a live formula when the sheet
   is opened. */
function sanitizeCell_(value) {
  var str = String(value == null ? '' : value);
  return /^[=+\-@]/.test(str) ? ("'" + str) : str;
}

function doGet(e) {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return jsonOutput_([]);
  var headers = data.shift();
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  var verbs = data
    .filter(function (row) { return row[idx.name]; })
    .map(function (row) {
      return {
        name: row[idx.name],
        meaning: row[idx.meaning] || '',
        conj: safeParse_(row[idx.conj]) || {},
        examples: safeParse_(row[idx.examples]) || {}
      };
    });

  return jsonOutput_(verbs);
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'Bad JSON body' });
  }

  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0] || ['name', 'meaning', 'conj', 'examples'];
  var nameCol = headers.indexOf('name');

  if (body.action === 'delete') {
    if (!body.name) return jsonOutput_({ ok: false, error: 'Missing name' });
    for (var r = data.length - 1; r >= 1; r--) {
      if (data[r][nameCol] === body.name) sheet.deleteRow(r + 1);
    }
    return jsonOutput_({ ok: true });
  }

  var verb = body.verb;
  if (!verb || !verb.name) return jsonOutput_({ ok: false, error: 'Missing verb' });

  var name = String(verb.name);
  var meaning = String(verb.meaning || '');
  var conjStr = JSON.stringify(verb.conj || {});
  var examplesStr = JSON.stringify(verb.examples || {});

  if (name.length > MAX_NAME_LENGTH || meaning.length > MAX_MEANING_LENGTH ||
      conjStr.length > MAX_CONJ_LENGTH || examplesStr.length > MAX_EXAMPLES_LENGTH) {
    return jsonOutput_({ ok: false, error: 'Field too long' });
  }

  var rowValues = [
    sanitizeCell_(name),
    sanitizeCell_(meaning),
    conjStr,
    examplesStr
  ];

  var existingRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][nameCol] === verb.name) { existingRow = i + 1; break; }
  }
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return jsonOutput_({ ok: true });
}

/** Run once from the Apps Script editor to set up headers + starter verbs. */
function seedInitialVerbs() {
  var sheet = getSheet_();
  sheet.clear();
  sheet.appendRow(['name', 'meaning', 'conj', 'examples']);

  var verbs = [
    {
      name: 'hablar', meaning: 'to speak',
      conj: {
        Present: { yo: 'hablo', tu: 'hablas', el: 'habla', nosotros: 'hablamos', vosotros: 'habláis', ellos: 'hablan' },
        Preterite: { yo: 'hablé', tu: 'hablaste', el: 'habló', nosotros: 'hablamos', vosotros: 'hablasteis', ellos: 'hablaron' },
        Imperfect: { yo: 'hablaba', tu: 'hablabas', el: 'hablaba', nosotros: 'hablábamos', vosotros: 'hablabais', ellos: 'hablaban' },
        Future: { yo: 'hablaré', tu: 'hablarás', el: 'hablará', nosotros: 'hablaremos', vosotros: 'hablaréis', ellos: 'hablarán' },
        Conditional: { yo: 'hablaría', tu: 'hablarías', el: 'hablaría', nosotros: 'hablaríamos', vosotros: 'hablaríais', ellos: 'hablarían' }
      },
      examples: {
        Present: { es: 'Yo hablo español.', en: 'I speak Spanish.' },
        Preterite: { es: 'Ayer hablé con mi amigo.', en: 'Yesterday I spoke with my friend.' },
        Imperfect: { es: 'Cuando era niño, hablaba mucho.', en: 'When I was a child, I used to talk a lot.' },
        Future: { es: 'Mañana hablaré contigo.', en: 'Tomorrow I will speak with you.' },
        Conditional: { es: 'Hablaría más si tuviera tiempo.', en: 'I would speak more if I had time.' }
      }
    },
    {
      name: 'comer', meaning: 'to eat',
      conj: {
        Present: { yo: 'como', tu: 'comes', el: 'come', nosotros: 'comemos', vosotros: 'coméis', ellos: 'comen' },
        Preterite: { yo: 'comí', tu: 'comiste', el: 'comió', nosotros: 'comimos', vosotros: 'comisteis', ellos: 'comieron' },
        Imperfect: { yo: 'comía', tu: 'comías', el: 'comía', nosotros: 'comíamos', vosotros: 'comíais', ellos: 'comían' },
        Future: { yo: 'comeré', tu: 'comerás', el: 'comerá', nosotros: 'comeremos', vosotros: 'comeréis', ellos: 'comerán' },
        Conditional: { yo: 'comería', tu: 'comerías', el: 'comería', nosotros: 'comeríamos', vosotros: 'comeríais', ellos: 'comerían' }
      }
    },
    {
      name: 'vivir', meaning: 'to live',
      conj: {
        Present: { yo: 'vivo', tu: 'vives', el: 'vive', nosotros: 'vivimos', vosotros: 'vivís', ellos: 'viven' },
        Preterite: { yo: 'viví', tu: 'viviste', el: 'vivió', nosotros: 'vivimos', vosotros: 'vivisteis', ellos: 'vivieron' },
        Imperfect: { yo: 'vivía', tu: 'vivías', el: 'vivía', nosotros: 'vivíamos', vosotros: 'vivíais', ellos: 'vivían' },
        Future: { yo: 'viviré', tu: 'vivirás', el: 'vivirá', nosotros: 'viviremos', vosotros: 'viviréis', ellos: 'vivirán' },
        Conditional: { yo: 'viviría', tu: 'vivirías', el: 'viviría', nosotros: 'viviríamos', vosotros: 'viviríais', ellos: 'vivirían' }
      }
    }
  ];

  verbs.forEach(function (v) {
    sheet.appendRow([v.name, v.meaning, JSON.stringify(v.conj), JSON.stringify(v.examples || {})]);
  });
}
