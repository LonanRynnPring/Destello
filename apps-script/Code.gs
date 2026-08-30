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
      },
      examples: {
        Present: { es: 'Yo como frutas cada día.', en: 'I eat fruit every day.' },
        Preterite: { es: 'Ayer comí pizza con mis amigos.', en: 'Yesterday I ate pizza with my friends.' },
        Imperfect: { es: 'Cuando era niño, comía mucho pan.', en: 'When I was a child, I used to eat a lot of bread.' },
        Future: { es: 'Mañana comeré en un restaurante nuevo.', en: 'Tomorrow I will eat at a new restaurant.' },
        Conditional: { es: 'Comería más verduras si tuvieran mejor sabor.', en: 'I would eat more vegetables if they tasted better.' }
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
      },
      examples: {
        Present: { es: 'Yo vivo en una ciudad pequeña.', en: 'I live in a small city.' },
        Preterite: { es: 'El año pasado viví en Madrid.', en: 'Last year I lived in Madrid.' },
        Imperfect: { es: 'De pequeño, vivía cerca de la playa.', en: 'As a child, I used to live near the beach.' },
        Future: { es: 'Algún día viviré en el extranjero.', en: 'Someday I will live abroad.' },
        Conditional: { es: 'Viviría en el campo si pudiera.', en: 'I would live in the countryside if I could.' }
      }
    },
    {
      name: 'estudiar', meaning: 'to study',
      conj: {
        Present: { yo: 'estudio', tu: 'estudias', el: 'estudia', nosotros: 'estudiamos', vosotros: 'estudiáis', ellos: 'estudian' },
        Preterite: { yo: 'estudié', tu: 'estudiaste', el: 'estudió', nosotros: 'estudiamos', vosotros: 'estudiasteis', ellos: 'estudiaron' },
        Imperfect: { yo: 'estudiaba', tu: 'estudiabas', el: 'estudiaba', nosotros: 'estudiábamos', vosotros: 'estudiabais', ellos: 'estudiaban' },
        Future: { yo: 'estudiaré', tu: 'estudiarás', el: 'estudiará', nosotros: 'estudiaremos', vosotros: 'estudiaréis', ellos: 'estudiarán' },
        Conditional: { yo: 'estudiaría', tu: 'estudiarías', el: 'estudiaría', nosotros: 'estudiaríamos', vosotros: 'estudiaríais', ellos: 'estudiarían' }
      },
      examples: {
        Present: { es: 'Yo estudio español todos los días.', en: 'I study Spanish every day.' },
        Preterite: { es: 'Anoche estudié para el examen.', en: 'Last night I studied for the exam.' },
        Imperfect: { es: 'Cuando estaba en la universidad, estudiaba mucho.', en: 'When I was in university, I used to study a lot.' },
        Future: { es: 'El próximo mes estudiaré en Barcelona.', en: 'Next month I will study in Barcelona.' },
        Conditional: { es: 'Estudiaría más si tuviera más tiempo libre.', en: 'I would study more if I had more free time.' }
      }
    }
  ];

  verbs.forEach(function (v) {
    sheet.appendRow([v.name, v.meaning, JSON.stringify(v.conj), JSON.stringify(v.examples || {})]);
  });
}
