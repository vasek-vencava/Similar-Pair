/*******************************
 * SMART PAIRING V2 + SIZE LOGIC (KOMPLETNÍ)
 * 
 * Integruje logiku řazení velikostí (Script A) do párování (Script B).
 * Pokud oba produkty obsahují velikost (ml, g, kg, l...), script je porovná.
 * - Shoda velikosti = Velký BONUS ke skóre.
 * - Rozdíl velikosti = Velká PENALIZACE (i když je název stejný).
 *******************************/

/***** ====== KONFIGURACE ====== *****/
// Kandidáti (referenční názvy, proti kterým se páruje)
const SOURCE_SHEET = 'Nase_nazvy';

// Struktura listu Nase_nazvy (A ignorujeme):
// B = značka (výrobce), C–R = naše/alternativní názvy
const SOURCE_BRAND_COL     = 2;  // B = značka
const SOURCE_DISPLAY_COL   = 3;  // C = CO SE MÁ VŽDY ZAPSAT DO VÝSTUPU
const SOURCE_ALT_START_COL = 3;  // C
const SOURCE_ALT_END_COL   = 18; // R (včetně)

// Dotazy (naše názvy k párování) – A ignorujeme (kód dodavatele)
const QUERY_SHEET      = 'Nazvy_k_Parovani';
const QUERY_NAME_COL   = 2; // B = název dotazu
const QUERY_BRAND_COL  = 3; // C = značka v dotazu (volitelně)

// Výstup (TopK shod)
const OUTPUT_SHEET     = 'Nazvy_k_Parovani';
const OUTPUT_START_COL = 15; // O

// Počet návrhů k vypsání
const TOP_K = 3;

// Výkon
const BATCH_SIZE = 1000;
const MAX_CANDIDATES = 250;
const TOKEN_CANDIDATE_MULT = 80;
const NUMERIC_BONUS_CAND = 150;

// Váhy
const BRAND_WEIGHT = 0.10;

// Soubor indexu na Google Drive (gzip JSON)
// Změněno jméno souboru, aby se vynutilo přegenerování indexu s novou logikou spec-matching
const DRIVE_INDEX_FILE = 'PAIRING_INDEX_v11.json.gz';

// Spec matching: bonus za shodu, penalizace za neshodu
const SPEC_MATCH_BONUS   = 0.15;  // +15% za každý sedící spec
const SPEC_MISMATCH_MULT = 0.2;   // ×0.2 za každý nesedící spec (tvrdá penalizace)

// Auto-pokračování
const MAX_RUN_MS = 5.5 * 60 * 1000;
const CONTINUE_DELAY_SEC = 15;
const RUN_FLAG_KEY = 'PAIRING_RUNNING';


/***** ====== MENU ====== *****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Párování produktů')
    .addItem('Spustit vše (po 1000, auto)', 'RunAllBatches')
    .addItem('Spustit JEDNU dávku (1000)', 'RunNextBatch')
    .addSeparator()
    .addItem('Vymazat výstup', 'VymazVystup')
    .addItem('Smazat cache', 'SmazCache')
    .addItem('Smazat index (Drive)', 'SmazIndexDrive')
    .addSeparator()
    .addItem('Seřadit…', 'SizeSortV1OpenSortDialog')
    .addSeparator()
    .addItem('URLs Aktualizace', 'urlsimport')
    .addToUi();
}

/***** ====== VSTUPNÍ BODY (dávky) ====== *****/

function RunAllBatches() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  const props = PropertiesService.getScriptProperties();
  props.setProperty(RUN_FLAG_KEY, '1');

  const t0 = Date.now();
  let totalFilled = 0;

  try {
    while ((Date.now() - t0) < MAX_RUN_MS) {
      const filled = runOneBatchInternal(BATCH_SIZE, /*interactive=*/false);
      totalFilled += filled;
      if (filled === 0) {
        CancelAutoContinue();
        props.deleteProperty(RUN_FLAG_KEY);
        console.log(`RunAllBatches: hotovo. Celkem vyplněno ~${totalFilled} řádků.`);
        return;
      }
    }
    scheduleAutoContinue();
    console.log(`RunAllBatches: doběhl čas, naplánováno pokračování.`);
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

function RunNextBatch() {
  const filled = runOneBatchInternal(BATCH_SIZE, /*interactive=*/true);
  SpreadsheetApp.getUi().alert(`Jednorázová dávka hotová. Vyplněno ~${filled} řádků.`);
}


/***** ====== JÁDRO DÁVKY ====== *****/

function runOneBatchInternal(limit, interactive) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(QUERY_SHEET);
  const outSheet = ss.getSheetByName(OUTPUT_SHEET);
  const outCols = TOP_K * 6;

  const info = findNextStartRow_();
  if (!info || info.startRow === -1) {
    if (interactive) SpreadsheetApp.getUi().alert('Nenalezen žádný další řádek k párování.');
    return 0;
  }

  const before = outSheet.getRange(info.startRow, OUTPUT_START_COL, limit, outCols).getValues();
  NajdiShodyTop3({ startRow: info.startRow, limit: limit });
  const after  = outSheet.getRange(info.startRow, OUTPUT_START_COL, limit, outCols).getValues();

  let filled = 0;
  for (let r = 0; r < after.length; r++) {
    const a = after[r], b = before[r];
    if (!a) break;
    if (a.some((v, j) => (v || '') !== (b?.[j] || ''))) filled++;
  }
  return filled;
}

function findNextStartRow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(QUERY_SHEET);
  const outSheet = ss.getSheetByName(OUTPUT_SHEET);
  const outCols = TOP_K * 6;

  const qLast = qSheet.getLastRow();
  if (qLast < 2) return { startRow: -1 };

  const numRows = Math.max(0, qLast - 1);
  const qNamesRaw = qSheet.getRange(2, QUERY_NAME_COL, numRows, 1).getValues().map(r => safeStr(r[0]));
  const outs = outSheet.getRange(2, OUTPUT_START_COL, numRows, outCols).getValues();

  for (let i = 0; i < numRows; i++) {
    const hasQuery = !!cleanCell(qNamesRaw[i]);
    const outEmpty = outs[i].every(v => v === '' || v == null);
    if (hasQuery && outEmpty) return { startRow: i + 2 };
  }
  return { startRow: -1 };
}


/***** ====== AUTO-POKRAČOVÁNÍ ====== *****/

function scheduleAutoContinue() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'RunAllBatches') return;
  }
  ScriptApp.newTrigger('RunAllBatches').timeBased().after(CONTINUE_DELAY_SEC * 1000).create();
}

function CancelAutoContinue() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    try {
      if (t.getHandlerFunction && t.getHandlerFunction() === 'RunAllBatches') ScriptApp.deleteTrigger(t);
    } catch(_) {}
  }
  try { PropertiesService.getScriptProperties().deleteProperty(RUN_FLAG_KEY); } catch(_) {}
}


/***** ====== HLAVNÍ PÁROVÁNÍ ====== *****/

function NajdiShodyTop3(opts = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Načtení kandidátů ---
  const src = ss.getSheetByName(SOURCE_SHEET);
  if (!src) throw new Error(`Sheet "${SOURCE_SHEET}" nenalezen.`);
  const srcLast = src.getLastRow();
  if (srcLast < 2) throw new Error(`V "${SOURCE_SHEET}" chybí data.`);

  const nRows = srcLast - 1;
  const brandsRaw   = src.getRange(2, SOURCE_BRAND_COL,   nRows, 1).getValues().map(r => safeStr(r[0]));
  const displayRaw  = src.getRange(2, SOURCE_DISPLAY_COL, nRows, 1).getValues().map(r => safeStr(r[0]));

  const altColsCount = SOURCE_ALT_END_COL - SOURCE_ALT_START_COL + 1;
  const altsMatrixRaw = altColsCount > 0
    ? src.getRange(2, SOURCE_ALT_START_COL, nRows, altColsCount).getValues()
    : Array.from({ length: nRows }, () => []);

  const brands   = brandsRaw.map(cleanCell);
  const displays = displayRaw.map(cleanCell);
  const altsMatrix = altsMatrixRaw.map(row => row.map(v => cleanCell(v)).filter(Boolean));

  const refDisplay = [];
  const refIndexVals = [];
  const refBrandsNorm = [];
  const oldVals = [];

  for (let i = 0; i < nRows; i++) {
    const br   = brands[i];
    const disp = displays[i];
    const alts = altsMatrix[i] || [];

    const firstAlt = alts.length ? alts[0] : disp;
    const indexStr = joinClean([br, disp, ...alts]);

    if (!indexStr) continue;

    refDisplay.push(disp);
    refIndexVals.push(indexStr);
    refBrandsNorm.push(normalize(br));
    oldVals.push(joinClean([br, firstAlt]) || '');
  }

  // --- Index ---
  const idx = buildOrLoadIndex(refIndexVals, oldVals, refBrandsNorm);
  idx.refItems.forEach((it, j) => { it.display = refDisplay[j] || ''; });

  // --- Načtení dotazů ---
  const qSheet = ss.getSheetByName(QUERY_SHEET);
  if (!qSheet) throw new Error(`Sheet "${QUERY_SHEET}" nenalezen.`);
  const qLast = qSheet.getLastRow();
  if (qLast < 2) { console.log('Žádné dotazy.'); return; }

  const startRow = Math.max(2, opts.startRow || 2);
  const maxEnd = opts.limit ? (startRow + opts.limit - 1) : qLast;

  const qNamesAllRaw = qSheet.getRange(2, QUERY_NAME_COL, Math.max(0, qLast - 1), 1).getValues().map(r => safeStr(r[0]));
  let lastNonEmpty = qNamesAllRaw.length + 1;
  for (let i = qNamesAllRaw.length - 1; i >= 0; i--) {
    if (cleanCell(qNamesAllRaw[i])) { lastNonEmpty = i + 2; break; }
  }
  const endRow = Math.min(maxEnd, lastNonEmpty);
  if (endRow < startRow) { console.log(`Nic ke zpracování.`); return; }

  const qNames  = qSheet.getRange(startRow, QUERY_NAME_COL,  endRow - startRow + 1, 1).getValues().map(r => cleanCell(r[0]));
  const qBrands = qSheet.getRange(startRow, QUERY_BRAND_COL, endRow - startRow + 1, 1).getValues().map(r => cleanCell(r[0]));

  // --- Výstup ---
  const outSheet = ss.getSheetByName(OUTPUT_SHEET);
  const outCols = TOP_K * 6;
  const header = [];
  for (let k = 1; k <= TOP_K; k++) header.push('Shoda ' + k, 'Celkem ' + k, 'Nazev ' + k, 'Fuzzy ' + k, 'Param ' + k, 'Stav ' + k);
  outSheet.getRange(1, OUTPUT_START_COL, 1, outCols).setValues([header]);

  const out = Array.from({ length: qNames.length }, () => Array(outCols).fill(''));

  // --- Scoring ---
  qNames.forEach((qName, i) => {
    if (!qName) return;

    const qBrand = qBrands[i] || '';
    const qBrandNorm = normalize(qBrand);
    const qCombined = joinClean([qBrand, qName]);
    
    // Preprocess včetně nové size-logiky
    const qp = preprocessName(applySynonyms(qCombined, idx.synonymMap), -1);

    // 1. Výběr kandidátů
    const candSet = new Set();
    
    // Podle tokenů
    const weightedTokens = weightTokensByIDF(qp.toks, idx.idf)
      .slice(0, Math.max(3, Math.ceil(qp.toks.length * 0.6)));

    for (const t of weightedTokens) {
      const arr = idxGet(idx.tokenIndex, t);
      for (let k = 0; k < arr.length && k < TOKEN_CANDIDATE_MULT; k++) candSet.add(arr[k]);
    }

    // Podle čísel (zachováno pro ID kódy atd)
    const numKeys = qp.num.map(n => numKey(n));
    for (const k of numKeys) {
      const arr = idxGet(idx.numIndex, k);
      for (let j = 0; j < arr.length && j < NUMERIC_BONUS_CAND; j++) candSet.add(arr[j]);
    }

    // Fallback
    if (candSet.size < 50) {
      const longest = qp.toks.slice().sort((a, b) => b.length - a.length)[0];
      if (longest) {
        const arr = idxGet(idx.tokenIndex, longest);
        for (let k = 0; k < arr.length && k < TOKEN_CANDIDATE_MULT; k++) candSet.add(arr[k]);
      }
    }
    let cands = Array.from(candSet);
    if (cands.length === 0) {
      const sampleSize = Math.min(1000, idx.refItems.length);
      const step = Math.max(1, Math.floor(idx.refItems.length / sampleSize));
      cands = [];
      for (let s = 0; s < idx.refItems.length && cands.length < sampleSize; s += step) cands.push(s);
    }

    // 2. Hrubé skóre
    const coarse = cands.map(id => {
      const r = idx.refItems[id];
      // Hrubý Jaccard + lehký bonus, pokud sedí normalizovaná velikost (SizeSortV1)
      let bonus = 0;
      if (qp.sizeVal !== null && r.sizeVal !== null && Math.abs(qp.sizeVal - r.sizeVal) < 0.001) {
        bonus = 0.3; 
      }
      return { id, s: jaccard(qp.toks, r.toks) + bonus };
    }).sort((a, b) => b.s - a.s);

    cands = coarse.slice(0, Math.min(MAX_CANDIDATES, coarse.length)).map(x => x.id);

    // 3. Finální skóre (Detailní porovnání)
    const scored = cands.map(id => {
      const r = idx.refItems[id];
      var detail = similarityScoreDetailed(qp, r);
      var bScore = brandMatchScore(qBrandNorm, r.brand);
      var final = detail.total + BRAND_WEIGHT * bScore;
      return {
        id: id, name: r.display || '', score: final,
        nameP: detail.nameP, fuzzyP: detail.fuzzyP, paramP: detail.paramP
      };
    }).sort((a, b) => b.score - a.score).slice(0, TOP_K);

    scored.forEach((t, k) => {
      var base = k * 6;
      var celkem = Math.round((t.score || 0) * 100);
      var nazev = Math.round((t.nameP || 0) * 100);
      var fuzzy = Math.round((t.fuzzyP || 0) * 100);
      var param = t.paramP !== null ? Math.round(t.paramP * 100) : '';
      var stav = interpretMatch(nazev, fuzzy, t.paramP !== null ? Math.round(t.paramP * 100) : null, celkem);
      out[i][base] = t.name || '';
      out[i][base + 1] = celkem;
      out[i][base + 2] = nazev;
      out[i][base + 3] = fuzzy;
      out[i][base + 4] = param;
      out[i][base + 5] = stav;
    });
  });

  outSheet.getRange(startRow, OUTPUT_START_COL, out.length, outCols).setValues(out);
  console.log(`NajdiShodyTop3: zpracováno ${out.length} řádků.`);
}


/***** ====== INDEX (DRIVE) ====== *****/

function buildOrLoadIndex(refVals, oldVals, brandsNormArr = []) {
  const first = (refVals[0] || '').slice(0, 32);
  const N = refVals.length;
  // Zmena verze => vynuti pregenerovani indexu (normalize: carka, split cislo/pismeno)
  const version = `v11-DETAILED_SCORING|N=${N}|F=${first}`;

  const fromDrive = loadIndexFromDrive();
  if (fromDrive && fromDrive.version === version) return fromDrive;

  const synonymMap = buildSynonymMap(oldVals, refVals);
  const refItems = refVals.map((v, i) => {
    const item = preprocessName(applySynonyms(v, synonymMap), i);
    item.brand = normalize(brandsNormArr[i] || '');
    return item;
  });

  const df = new Map();
  refItems.forEach(r => {
    const seen = new Set(r.toks);
    seen.forEach(t => df.set(t, (df.get(t) || 0) + 1));
  });
  const idf = Object.create(null);
  df.forEach((d, t) => idf[t] = Math.log(1 + N / (1 + d)));

  const tokenIndex = {};
  refItems.forEach(r => {
    const seen = new Set(r.toks);
    seen.forEach(t => {
      if (!tokenIndex[t]) tokenIndex[t] = [];
      tokenIndex[t].push(r.idx);
    });
  });

  const numIndex = {};
  refItems.forEach(r => {
    const keys = r.num.map(n => numKey(n));
    new Set(keys).forEach(k => {
      if (!numIndex[k]) numIndex[k] = [];
      numIndex[k].push(r.idx);
    });
  });

  const payload = { refItems, tokenIndex, idf, numIndex, synonymMap, version };
  saveIndexToDrive(payload);
  return payload;
}

function saveIndexToDrive(obj) {
  const json = JSON.stringify(obj);
  const gzBlob = Utilities.newBlob(json, 'application/json', 'index.json').setName(DRIVE_INDEX_FILE);
  const gz = Utilities.gzip(gzBlob);
  
  const files = DriveApp.getFilesByName(DRIVE_INDEX_FILE);
  while (files.hasNext()) { try { files.next().setTrashed(true); } catch (_) {} }
  DriveApp.createFile(gz);
}

function loadIndexFromDrive() {
  const files = DriveApp.getFilesByName(DRIVE_INDEX_FILE);
  if (!files.hasNext()) return null;
  try {
    const gz = files.next().getBlob();
    const unzipped = Utilities.ungzip(gz);
    return JSON.parse(unzipped.getDataAsString());
  } catch (e) { return null; }
}

function deleteIndexFromDrive() {
  const files = DriveApp.getFilesByName(DRIVE_INDEX_FILE);
  while (files.hasNext()) { try { files.next().setTrashed(true); } catch (_) {} }
}


/***** ====== POMOCNÉ FUNKCE ====== *****/

function cleanCell(v) {
  const s = safeStr(v).trim();
  if (!s || /^#?N\/A|#REF|#VALUE|#DIV\/0|#NAME|#NULL|#ERROR/i.test(s.replace(/\s/g,''))) return '';
  return s;
}

function idxGet(idx, key) {
  if (idx instanceof Map) return idx.get(key) || [];
  return (idx && idx[key]) ? idx[key] : [];
}

function buildSynonymMap(oldArr, newArr) {
  if (!Array.isArray(oldArr) || !Array.isArray(newArr)) return {};
  const counts = {};
  for (let i = 0; i < Math.min(oldArr.length, newArr.length); i++) {
    const o = normalize(oldArr[i]), n = normalize(newArr[i]);
    if (!o || !n) continue;
    const ot = tokenSet(o), nt = tokenSet(n);
    const onlyO = fastDiff(ot, nt), onlyN = fastDiff(nt, ot);
    if (onlyO.length === 1 && onlyN.length === 1) {
      const key = `${onlyO[0]}=>${onlyN[0]}`;
      counts[key] = (counts[key] || 0) + 1;
    }
  }
  const map = {};
  for (const k in counts) if (counts[k] >= 2) {
    const [from, to] = k.split('=>'); map[from] = to;
  }
  const builtIns = { 'x': '×', 'pack': 'baleni', 'sacek': 'sacek', 'litru': 'l', 'litr': 'l' };
  return { ...builtIns, ...map };
}

function applySynonyms(str, synMap) {
  const toks = normalize(str).split(/\s+/).filter(Boolean);
  for (let i = 0; i < toks.length; i++) toks[i] = synMap[toks[i]] || toks[i];
  return toks.join(' ');
}

function preprocessName(str, idx) {
  var norm = normalize(str);
  var toks = tokens(norm);
  var tri = charTrigrams(norm);
  var num = extractNumbersWithUnits(norm);
  var specs = extractSpecs(norm);
  // nameToks: tokeny bez cisel a jednotek — cista shoda nazvu/kodu
  var nameNorm = norm
    .replace(/\d+(?:\.\d+)?\s*(?:ml|l|kg|g|mm|cm|m\s*3|m|kw|w|ks|pack|bal)\b/g, '')
    .replace(/\b\d+(?:\.\d+)?\b/g, '')
    .replace(/\s+/g, ' ').trim();
  var nameToks = tokens(nameNorm);
  // sizeVal zachovan pro zpetnou kompatibilitu (coarse scoring)
  var sizeVal = null;
  for (var i = 0; i < specs.length; i++) {
    if (specs[i].category === 'volume' || specs[i].category === 'weight') { sizeVal = specs[i].baseVal; break; }
  }
  return { idx: idx, norm: norm, toks: toks, tri: tri, num: num, specs: specs, sizeVal: sizeVal, nameToks: nameToks };
}

function weightTokensByIDF(toks, idf) {
  const uniq = Array.from(new Set(toks));
  return uniq.map(t => [t, idf[t] || 0]).sort((a,b)=>b[1]-a[1]).map(x=>x[0]);
}

/***** ====== EXTRAKCE SPECIFIKACÍ (ROZŠÍŘENÁ LOGIKA) ====== *****/

/**
 * Extrahuje všechny měřitelné specifikace z názvu produktu.
 * Vrací pole objektů { category, baseVal } seskupených podle kategorie:
 *   - "volume"  : ml, l           -> normalizováno na ml
 *   - "weight"  : g, kg           -> normalizováno na g
 *   - "length"  : mm, cm, m       -> normalizováno na mm
 *   - "flow"    : m3, m3/h        -> normalizováno na m3
 *   - "power"   : w, kw           -> normalizováno na W
 * Kazda kategorie se vyskytne max jednou (posledni vyskyt vyhraje).
 */
function extractSpecs(str) {
  if (!str) return [];
  var raw = String(str).toLowerCase().replace(/,/g, '.').trim();
  var specs = {};
  var m;

  // Prutok/vykon ventilatoru: 1300m3, 1300 m 3, 1300m3/h
  var flowRe = /(\d+(?:\.\d+)?)\s*m\s*3(?:\s*\/\s*h)?\b/g;
  while ((m = flowRe.exec(raw)) !== null) {
    specs['flow'] = parseFloat(m[1]);
  }

  // Vykon: 150w, 1.5kw
  var powerRe = /(\d+(?:\.\d+)?)\s*(kw|w)\b/g;
  while ((m = powerRe.exec(raw)) !== null) {
    var val = parseFloat(m[1]);
    specs['power'] = m[2] === 'kw' ? val * 1000 : val;
  }

  // Objem, vaha, delka
  var sizeRe = /(\d+(?:\.\d+)?)\s*(ml|l|kg|g|mm|cm|m)\b/g;
  while ((m = sizeRe.exec(raw)) !== null) {
    var sVal = parseFloat(m[1]);
    var unit = m[2];
    if (!Number.isFinite(sVal)) continue;

    switch (unit) {
      case 'ml': specs['volume'] = sVal; break;
      case 'l':  specs['volume'] = sVal * 1000; break;
      case 'g':  specs['weight'] = sVal; break;
      case 'kg': specs['weight'] = sVal * 1000; break;
      case 'mm': specs['length'] = sVal; break;
      case 'cm': specs['length'] = sVal * 10; break;
      case 'm':
        if (!specs['flow']) specs['length'] = sVal * 1000;
        break;
    }
  }

  var result = [];
  for (var cat in specs) {
    result.push({ category: cat, baseVal: specs[cat] });
  }
  return result;
}

/** Zpetna kompatibilita: vraci sizeVal jako drive (pro coarse scoring). */
function extractNormalizedSize(str) {
  var specs = extractSpecs(str);
  for (var i = 0; i < specs.length; i++) {
    if (specs[i].category === 'volume' || specs[i].category === 'weight') return specs[i].baseVal;
  }
  return null;
}

/***** ====== VÝPOČET SKÓRE PODOBNOSTI (V3 — DETAILNÍ) ====== *****/

/**
 * Vraci objekt s detailnimi slozkami:
 *   total    — celkove skore (bez capu, muze byt nad 1.0)
 *   nameP    — shoda nazvu/kodu 0-1 (jaccard na tokenech bez cisel)
 *   fuzzyP   — fuzzy shoda textu 0-1 (cosine na trigramech)
 *   paramP   — shoda parametru 0-1, nebo null pokud zadne specs k porovnani
 */
function similarityScoreDetailed(a, b) {
  // 1. Nazev: Jaccard na tokenech bez cisel/jednotek
  var nameP = jaccard(a.nameToks || [], b.nameToks || []);

  // 2. Fuzzy: Cosine na trigramech (plny text vcetne cisel)
  var fuzzyP = cosineSim(freq(a.tri), freq(b.tri));

  // 3. Parametry: podil sedících specs
  var specResult = compareSpecs(a.specs || [], b.specs || []);
  var totalSpecs = specResult.matches + specResult.mismatches;
  var paramP = totalSpecs > 0 ? (specResult.matches / totalSpecs) : null;

  // 4. Celkove skore: text zaklad + spec bonusy/penalty
  var textScore = 0.6 * fuzzyP + 0.4 * jaccard(a.toks, b.toks);
  var total = textScore;
  for (var p = 0; p < specResult.mismatches; p++) {
    total *= SPEC_MISMATCH_MULT;
  }
  total += specResult.matches * SPEC_MATCH_BONUS;

  // Stary ciselny bonus (ks, pack, x)
  var oldNumBonus = numberUnitBonus(a.num, b.num);
  total += 0.05 * oldNumBonus;

  return {
    total: Math.max(0, total),
    nameP: nameP,
    fuzzyP: fuzzyP,
    paramP: paramP
  };
}

/**
 * Interpretace shody na zaklade komponent.
 * nameP, fuzzyP, paramP: 0-100 (procenta), paramP muze byt null.
 * totalP: celkove procento.
 */
function interpretMatch(nameP, fuzzyP, paramP, totalP) {
  if (totalP < 15) return 'Nenalezeno';
  if (nameP < 40 && fuzzyP < 50) return 'Neshoda';

  // Parametry nejsou k dispozici
  if (paramP === null) {
    if (nameP >= 80 && fuzzyP >= 80) return 'Shoda bez parametru';
    if (nameP >= 60 && fuzzyP >= 70) return 'Pravdepodobna shoda';
    if (nameP >= 40 && fuzzyP >= 60) return 'Slaba shoda';
    return 'Neshoda';
  }

  // Vsechno sedi
  if (nameP >= 80 && fuzzyP >= 80 && paramP >= 80) return 'Jista shoda';
  if (nameP >= 60 && fuzzyP >= 70 && paramP >= 80) return 'Pravdepodobna shoda';

  // Nazev nesedi ale parametry ano
  if (nameP < 60 && fuzzyP >= 60 && paramP >= 80) return 'Zkontroluj nazev';

  // Nazev sedi ale parametry ne
  if (nameP >= 60 && fuzzyP >= 60 && paramP < 50) return 'Zkontroluj parametry';
  if (nameP >= 40 && fuzzyP >= 60 && paramP < 50) return 'Mozna varianta';

  // Nic jasne nesedi
  if (nameP < 60 && fuzzyP < 70) return 'Slaba shoda';

  return 'Slaba shoda';
}

/**
 * Porovna specs dvou produktu.
 * Vraci { matches, mismatches } — pocet kategorii kde se shodly/neshodly.
 * Kategorie ktera chybi u jednoho z produktu se ignoruje (neni match ani mismatch).
 */
function compareSpecs(specsA, specsB) {
  var mapA = {};
  for (var i = 0; i < specsA.length; i++) mapA[specsA[i].category] = specsA[i].baseVal;
  var mapB = {};
  for (var j = 0; j < specsB.length; j++) mapB[specsB[j].category] = specsB[j].baseVal;

  var matches = 0, mismatches = 0;
  for (var cat in mapA) {
    if (!(cat in mapB)) continue; // kategorie chybi u druheho — ignorujeme
    var ratio = mapA[cat] / mapB[cat];
    if (ratio > 0.95 && ratio < 1.05) {
      matches++;
    } else {
      mismatches++;
    }
  }
  return { matches: matches, mismatches: mismatches };
}


/***** TEXT UTILS *****/
function safeStr(v){ return (v==null)?'':String(v); }
function joinClean(parts){ return parts.filter(Boolean).join(' ').replace(/\s+/g,' ').trim(); }
function normalize(s){
  return safeStr(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    // carka jako desetinny oddelovac: "0,5" -> "0.5" (pred odstranenim nepovolených znaku)
    .replace(/(\d),(\d)/g, '$1.$2')
    // "+" je soucasti nazvu produktu (Q240+ vs Q240 = jiny produkt)
    .replace(/[^\p{L}\p{N}\s.%×x\-\/+]/gu,' ')
    // rozdelit cisla od pismen: "16pot" -> "16 pot", "easy2grow" -> "easy 2 grow"
    .replace(/(\d)([a-z])/gi, '$1 $2')
    .replace(/([a-z])(\d)/gi, '$1 $2')
    .replace(/\s+/g,' ').trim();
}
const STOPWORDS = new Set(['a','i','u','v','ve','s','se','z','ze','k','ke','o','do','na','za','pro','bez','od','pod','nad','po']);
function tokens(s){
  return s.split(' ').filter(t => t && !STOPWORDS.has(t));
}
function tokenSet(s){ return new Set(tokens(s)); }
function fastDiff(aSet, bSet){ const r=[]; aSet.forEach(x=>{ if(!bSet.has(x)) r.push(x); }); return r; }
function charTrigrams(s){
  const t = `  ${s}  `, res = [];
  for (let i=0;i<t.length-2;i++) res.push(t.slice(i,i+3));
  return res;
}
function freq(arr){ const m={}; for (let i=0;i<arr.length;i++) m[arr[i]]=(m[arr[i]]||0)+1; return m; }
function cosineSim(a,b){
  let dot=0,na=0,nb=0;
  for (const k in a){ const av=a[k]; na+=av*av; if(b[k]) dot+=av*b[k]; }
  for (const k in b) nb+=b[k]**2;
  return (!na||!nb) ? 0 : dot/Math.sqrt(na*nb);
}
function jaccard(aArr,bArr){
  const a=new Set(aArr), b=new Set(bArr);
  if (!a.size && !b.size) return 1;
  let inter=0; a.forEach(x=>{ if(b.has(x)) inter++; });
  return inter/(a.size+b.size-inter);
}

// Stará logika pro ostatní čísla (balení, kusy) - ponechána jako doplněk
function extractNumbersWithUnits(s){
  const re = /(\d+(?:[\.,]\d+)?)(\s?(ml|l|g|kg|ks|mm|cm|m|pack|bal|x|×))?/gi;
  const arr = []; let m;
  while ((m = re.exec(s)) !== null) {
    arr.push({ val: parseFloat(m[1].replace(',', '.')), unit: (m[3] || '').toLowerCase() });
  }
  return arr;
}
function numKey(n){ return `${Math.round(n.val*100)/100}|${n.unit||''}`; }
function brandMatchScore(qBrand, cBrand) {
  if (!qBrand || !cBrand) return 0;
  return jaccard(tokens(qBrand), tokens(cBrand));
}
function numberUnitBonus(aNums,bNums){
  if (!aNums.length || !bNums.length) return 0;
  let best=0;
  for (const a of aNums) {
    for (const b of bNums) {
      // Zde ignorujeme ml/g/l, protože to řeší nová logika nahoře. Řešíme jen ks, pack, x...
      const unitIgnored = ['ml','l','g','kg','mm','cm','m'];
      if (unitIgnored.includes(a.unit) || unitIgnored.includes(b.unit)) continue;

      const sameUnit=(!a.unit && !b.unit)||a.unit===b.unit||(a.unit==='x'&&b.unit==='×')||(a.unit==='×'&&b.unit==='x');
      if(!sameUnit) continue;
      const denom=Math.max(1e-9, Math.max(a.val,b.val));
      const diff=Math.abs(a.val-b.val)/denom;
      const local = diff<=0.02?1: (diff<=0.1?0.5:0);
      if(local>best) best=local;
    }
  }
  return best;
}

/***** ÚDRŽBA / NÁSTROJE *****/
function VymazVystup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(OUTPUT_SHEET);
  const rows = Math.max(0, sh.getLastRow() - 1);
  const cols = TOP_K * 6;
  if (rows > 0) sh.getRange(2, OUTPUT_START_COL, rows, cols).clearContent();
}

function SmazCache() {
  try { CacheService.getScriptCache().remove('IDX_VER'); } catch(_) {}
  SpreadsheetApp.getUi().alert('Cache byla vymazána.');
}

function SmazIndexDrive() {
  deleteIndexFromDrive();
  SpreadsheetApp.getUi().alert('Index na Drive byl smazán. Při příštím spuštění se znovu vytvoří.');
}


/***** LADĚNÍ (volitelné - pro kontrolu konkrétního řádku) *****/
function DebugRow(row = 2) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SOURCE_SHEET);
  const qSheet = ss.getSheetByName(QUERY_SHEET);

  const nRows = src.getLastRow()-1;
  const brandsRaw   = src.getRange(2, SOURCE_BRAND_COL,   nRows,1).getValues().map(r=>safeStr(r[0]));
  const displayRaw  = src.getRange(2, SOURCE_DISPLAY_COL, nRows,1).getValues().map(r=>safeStr(r[0]));
  const altColsCount = SOURCE_ALT_END_COL - SOURCE_ALT_START_COL + 1;
  const altsMatrixRaw = altColsCount>0 ? src.getRange(2, SOURCE_ALT_START_COL, nRows, altColsCount).getValues() : [];

  const brands   = brandsRaw.map(cleanCell);
  const displays = displayRaw.map(cleanCell);
  const altsMatrix = (altsMatrixRaw||[]).map(row=>row.map(v=>cleanCell(v)).filter(Boolean));

  const refDisplay = [];
  const refIndexVals = [];
  const refBrandsNorm = [];
  const oldVals = [];
  for (let i=0;i<nRows;i++){
    const br=brands[i], disp=displays[i], alts=altsMatrix[i]||[];
    const idxStr  = joinClean([br, disp, ...alts]);
    if (!idxStr) continue;
    refDisplay.push(disp);
    refIndexVals.push(idxStr);
    refBrandsNorm.push(normalize(br));
    const firstAlt = alts.length?alts[0]:disp;
    oldVals.push(joinClean([br, firstAlt]) || '');
  }

  const idx = buildOrLoadIndex(refIndexVals, oldVals, refBrandsNorm);
  idx.refItems.forEach((it, j) => it.display = refDisplay[j] || '');

  const qName  = cleanCell(qSheet.getRange(row, QUERY_NAME_COL).getValue());
  const qBrand = cleanCell(qSheet.getRange(row, QUERY_BRAND_COL).getValue());
  const qComb  = joinClean([qBrand, qName]);
  const qp = preprocessName(applySynonyms(qComb, idx.synonymMap), -1);
  const qBrandNorm = normalize(qBrand);

  console.log('DebugRow ' + row);
  console.log('Dotaz: "' + qComb + '"');
  console.log('Specs: ' + JSON.stringify(qp.specs));
  console.log('Tokeny: ' + JSON.stringify(qp.toks));

  var tk = weightTokensByIDF(qp.toks, idx.idf).slice(0, Math.max(3, Math.ceil(qp.toks.length*0.6)));
  var testToken = tk[0];

  var refItems = idx.refItems;
  var cands = (idxGet(idx.tokenIndex, testToken)||[]).slice(0,100);
  if (cands.length === 0) cands = [0,1,2,3,4];

  var scored = cands.map(function(id){
      var r = refItems[id];
      var detail = similarityScoreDetailed(qp, r);
      var bScore = brandMatchScore(qBrandNorm, r.brand);
      var final = detail.total + BRAND_WEIGHT*bScore;
      var celkem = Math.round(final*100);
      var nazev = Math.round(detail.nameP*100);
      var fuzzy = Math.round(detail.fuzzyP*100);
      var param = detail.paramP !== null ? Math.round(detail.paramP*100) : null;
      return { id: id, name: (r.display||''), celkem: celkem, nazev: nazev, fuzzy: fuzzy, param: param,
        stav: interpretMatch(nazev, fuzzy, param, celkem), brand: r.brand };
    })
    .sort(function(a,b){ return b.celkem-a.celkem; }).slice(0,5);

  console.log('--- TOP 5 VYSLEDKU PRO RADEK ' + row + ' ---');
  scored.forEach(function(s) {
    console.log('Celkem: ' + s.celkem + '% | Nazev: ' + s.nazev + '% | Fuzzy: ' + s.fuzzy + '% | Param: ' + (s.param !== null ? s.param + '%' : 'N/A') + ' | Stav: ' + s.stav + ' | "' + s.name + '" | Znacka: ' + s.brand);
  });
}
