/***** KONSTANTY (měň podle potřeby) *****/
const AA_URL_DE = 'https://www.higarden.de/export/products.csv?patternId=-6&partnerId=7&hash=41ab5d0b6071ae7d2d2a1a33f0d962b2a860fcbad25938e1e845806243ea41be';
const AA_URL_EN = 'https://www.higarden.eu/export/products.csv?patternId=-6&partnerId=7&hash=f7a687c4a7800a046a18345229a333d819f1981a6680d03b81d58a43fffbdf75';

const AA_SHEET_DE = 'skladDE';     // cílový list pro DE
const AA_SHEET_EN = 'skladEN';     // cílový list pro EN
const AA_START_COLUMN = 2;         // B = 2 (začít psát od sloupce B)
const AA_ENCODING = 'windows-1250';
const AA_WRITE_HEADERS = true;     // zapsat hlavičku do prvního řádku?
const AA_VARIANT_PREFIX = 'variant:'; // prefix pro variantní sloupce
const AA_POSSIBLE_DELIMITERS = ['\t', ';', ',']; // pokusně: tab, středník, čárka


/***** VSTUPNÍ FUNKCE *****/
function AA_importAll() {
  AA_importToSheet(AA_URL_DE, AA_SHEET_DE);
  AA_importToSheet(AA_URL_EN, AA_SHEET_EN);
}

function AA_importDE() {
  AA_importToSheet(AA_URL_DE, AA_SHEET_DE);
}

function AA_importEN() {
  AA_importToSheet(AA_URL_EN, AA_SHEET_EN);
}

/***** JÁDRO IMPORTU *****/
function AA_importToSheet(url, sheetName) {
  const rows = AA_fetchAndParseCsv(url);
  if (!rows || rows.length === 0) {
    throw new Error('CSV je prázdné nebo nešlo načíst.');
  }

  const output = AA_selectColumnsAndCompose(rows); // vybere code, pairCode, name, a vytvoří Name+Variant

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  // Vymazat starý obsah od sloupce B vpravo (ponechat sloupec A nedotčený)
  const lastCol = sheet.getMaxColumns();
  sheet.getRange(1, AA_START_COLUMN, sheet.getMaxRows(), Math.max(1, lastCol - AA_START_COLUMN + 1)).clearContent();

  // Zápis nových dat od B1
  if (output.length && output[0].length) {
    const rng = sheet.getRange(1, AA_START_COLUMN, output.length, output[0].length);
    rng.setValues(output);
    // všechno jako text (zachování počátečních nul apod.)
    rng.setNumberFormat('@');
    // pro jistotu zarovnat vlevo (vizuálně)
    rng.setHorizontalAlignment('left');
  }
}

/***** STAŽENÍ A PARSOVÁNÍ CSV *****/
function AA_fetchAndParseCsv(url) {
  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'Accept': 'text/csv,*/*;q=0.1' }
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' při stahování: ' + url);
  }

  // Dekódování Windows-1250
  const bytes = res.getContent();
  let text = Utilities.newBlob(bytes).getDataAsString(AA_ENCODING);
  if (!text) throw new Error('Nepodařilo se dekódovat CSV jako ' + AA_ENCODING);

  // odstranit případný BOM
  text = text.replace(/^\uFEFF/, '');

  // detekce oddělovače
  const delimiter = AA_detectDelimiter(text);

  // parse
  const rows = Utilities.parseCsv(text, delimiter) || [];

  // odstranit prázdné úplné řádky
  return rows.filter(r => r && r.some(c => String(c).trim() !== ''));
}

/***** DETEKCE ODDĚLOVAČE *****/
function AA_detectDelimiter(text) {
  const firstLine = (text.split(/\r\n|\n|\r/, 1)[0] || '');
  let best = AA_POSSIBLE_DELIMITERS[0], bestCount = -1;

  AA_POSSIBLE_DELIMITERS.forEach(d => {
    // pro tab musíme hledat \t
    const pattern = d === '\t' ? /\t/g : new RegExp(escapeRegExp(d), 'g');
    const count = (firstLine.match(pattern) || []).length;
    if (count > bestCount) { bestCount = count; best = d; }
  });

  return best;
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/***** VÝBĚR SLOUPCŮ A SESTAVENÍ Name+Variant *****/
function AA_selectColumnsAndCompose(rows) {
  if (rows.length === 0) return [];

  const header = rows[0].map(h => String(h).trim());
  const headerLower = header.map(h => h.toLowerCase());

  const idxCode     = headerLower.indexOf('code');
  const idxPairCode = headerLower.indexOf('paircode');
  const idxName     = headerLower.indexOf('name');

  if (idxCode < 0 || idxPairCode < 0 || idxName < 0) {
    throw new Error('V CSV chybí některý z povinných sloupců: code, pairCode, name.');
  }

  // všechny indexy sloupců, které začínají "variant:"
  const variantIdxs = headerLower
    .map((h, i) => h.startsWith(AA_VARIANT_PREFIX) ? i : -1)
    .filter(i => i >= 0);

  if (variantIdxs.length === 0) {
    // není to fatální — jen nebude Name+Variant rozšířen o variantu
    // (ponecháme jen name)
  }

  const out = [];

  if (AA_WRITE_HEADERS) {
    out.push(['code', 'pairCode', 'name', 'Name + Variant']);
  }

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];

    const code     = safeCell(row[idxCode]);
    const pairCode = safeCell(row[idxPairCode]);
    const name     = safeCell(row[idxName]);

    // posbírat neprázdné varianty a spojit je mezerou
    const variants = (variantIdxs.length ? variantIdxs.map(i => safeCell(row[i])) : [])
      .filter(v => v !== '');
    const variantStr = variants.join(' ').trim();

    const namePlusVariant = variantStr ? (name + ' ' + variantStr) : name;

    // výstupní pořadí: code, pairCode, name, Name+Variant
    out.push([code, pairCode, name, namePlusVariant]);
  }

  return out;
}

function safeCell(v) {
  // převod na text + ořezání mezer
  return (v === null || v === undefined) ? '' : String(v).trim();
}
