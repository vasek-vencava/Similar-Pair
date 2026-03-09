/*** === Nastavení === ***/
const URLS_FEED       = 'https://www.higarden.cz/url3.xml?hash=pBn37QiZhEJ6QyTgo2JqN0n';
const URLS_SHEET_NAME = 'URLs';

/*** === Hlavní funkce === ***/
function urlsimport() {
  const sheet = getOrCreateSheet_(URLS_SHEET_NAME);
  const xmlText = fetchXmlText_(URLS_FEED);
  const items = parseShopItems_(xmlText); // [{code, url}, ...]

  // připravíme data pro zápis
  const timestamp = new Date();
  const header = [['CODE', 'URL', 'Imported at']];
  const rows = items.map(row => [row.code || '', row.url || '', timestamp]);

  // vyčistit a zapsat
  sheet.clear({contentsOnly: true});
  sheet.getRange(1, 1, header.length, header[0].length).setValues(header);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header[0].length).setValues(rows);
  }

  // kosmetika
  sheet.setFrozenRows(1);

  SpreadsheetApp.getActive().toast(`Import hotový: ${rows.length} položek.`, 'URL Import', 5);
}


/*** === Volitelné: naplánování (spouštění každou hodinu) === ***/
function createHourlyTrigger() {
  ScriptApp.newTrigger('urlsimport').timeBased().everyHours(1).create();
}

/*** === Pomocné funkce === ***/

// Stáhne XML jako text. Cache použijeme jen pokud je payload malý (< ~90 kB)
function fetchXmlText_(url) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'feed:' + url;

  // 1) zkusit cache
  try {
    const cached = cache.get(cacheKey);
    if (cached) return cached;
  } catch (e) {
    // ignoruj chyby cache
  }

  // 2) stáhnout
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    followRedirects: true,
    muteHttpExceptions: true,
    validateHttpsCertificates: true,
    // volitelně: { 'Accept-Encoding': 'gzip' } ale GAS si řeší sám
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Stahování selhalo (HTTP ${code}) z ${url}`);
  }

  const text = res.getContentText(); // UTF-8 XML jako string

  // 3) uložit do cache jen pokud je hodnotově malé (CacheService limit ~100 kB)
  try {
    // odhad velikosti (UTF-16 JS string ≠ přesná velikost v bajtech; použijeme JSON)
    const approxBytes = Utilities.newBlob(text).getBytes().length;
    if (approxBytes < 90 * 1024) {
      cache.put(cacheKey, text, 180); // 3 min cache
    }
  } catch (e) {
    // ignoruj chyby cache (přetečení apod.)
  }

  return text;
}

// Vrátí/ vytvoří list
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// Bezpečně projde XML a vrátí pole objektů {code, url}
function parseShopItems_(xmlText) {
  const doc = XmlService.parse(xmlText);
  const root = doc.getRootElement();

  // Najdi všechny elementy SHOPITEM kdekoliv ve stromu:
  const shopItems = findElementsByNameRecursive_(root, 'SHOPITEM');

  // Z každého vytáhneme CODE a URL (ignorujeme namespaces, case-insensitive)
  return shopItems.map(itemEl => ({
    code: getFirstChildTextByName_(itemEl, 'CODE'),
    url:  getFirstChildTextByName_(itemEl, 'URL'),
  }));
}

/*** === Utility pro práci s XmlService (bez XPath) === ***/

// Rekurzivně najde všechny elementy s daným názvem (case-insensitive)
function findElementsByNameRecursive_(element, nameWanted) {
  const wanted = String(nameWanted).toUpperCase();
  const out = [];
  const stack = [element];

  while (stack.length) {
    const el = stack.pop();
    if (el.getName && el.getName().toUpperCase() === wanted) {
      out.push(el);
    }
    const children = el.getChildren();
    for (let i = 0; i < children.length; i++) {
      stack.push(children[i]);
    }
  }
  return out;
}

// Vrátí text prvního přímého potomka s daným názvem (case-insensitive, ignoruje namespace)
function getFirstChildTextByName_(element, childName) {
  const children = element.getChildren();
  const target = String(childName).toUpperCase();
  for (let i = 0; i < children.length; i++) {
      const ch = children[i];
      if (ch.getName && ch.getName().toUpperCase() === target) {
        return (ch.getText && ch.getText()) || '';
      }
  }
  return '';
}
