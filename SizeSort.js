/*******************************
 * SizeSortV1
 * - Menu: Seradit...
 * - Prompt dialog (bez HTML - nevyzaduje OAuth showModalDialog)
 * - Razeni: nejprve podle nazvu (text bez cisla+jednotky), uvnitr podle normalizovane hodnoty
 * - Podporovane jednotky: ml/l, g/kg, mm/cm/m
 *******************************/

/** Otevre prompt dialog pro vyber sloupce a spusti razeni */
function SizeSortV1OpenSortDialog() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastCol = sheet.getLastColumn();

  // 1. Ma hlavicku?
  var headerResponse = ui.alert(
    'Hlavicka',
    'Ma list hlavicku (prvni radek)?',
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (headerResponse === ui.Button.CANCEL) return;
  var hasHeader = (headerResponse === ui.Button.YES);

  // 2. Rozsah sloupcu k razeni (ktere sloupce se maji presouvat)
  var rangeResponse = ui.prompt(
    'Rozsah sloupcu',
    'Ktere sloupce seradit? (napr. A:G nebo A-G)\nPrazdne = vsechny sloupce.',
    ui.ButtonSet.OK_CANCEL
  );
  if (rangeResponse.getSelectedButton() !== ui.Button.OK) return;

  var rangeInput = rangeResponse.getResponseText().trim().toUpperCase();
  var startColIdx = 1;
  var endColIdx = lastCol;

  if (rangeInput) {
    var parts = rangeInput.split(/[\s:\-]+/);
    if (parts.length >= 2) {
      var p1 = parseInt(parts[0], 10);
      var p2 = parseInt(parts[1], 10);
      startColIdx = isNaN(p1) ? SizeSortV1LetterToColIndex(parts[0]) : p1;
      endColIdx = isNaN(p2) ? SizeSortV1LetterToColIndex(parts[1]) : p2;
    } else if (parts.length === 1) {
      var p = parseInt(parts[0], 10);
      endColIdx = isNaN(p) ? SizeSortV1LetterToColIndex(parts[0]) : p;
    }

    if (!startColIdx || !endColIdx || startColIdx < 1 || endColIdx < 1 || startColIdx > endColIdx) {
      ui.alert('Neplatny rozsah sloupcu: ' + rangeResponse.getResponseText());
      return;
    }
  }

  // 3. Vyber sloupec k razeni
  var colResponse = ui.prompt(
    'Sloupec k razeni',
    'Podle ktereho sloupce seradit? (pismeno nebo cislo, napr. B nebo 2)\nRozsah: ' + SizeSortV1ColIndexToLetter(startColIdx) + ':' + SizeSortV1ColIndexToLetter(endColIdx),
    ui.ButtonSet.OK_CANCEL
  );
  if (colResponse.getSelectedButton() !== ui.Button.OK) return;

  var colInput = colResponse.getResponseText().trim().toUpperCase();
  var colIndex = parseInt(colInput, 10);
  if (isNaN(colIndex)) {
    colIndex = SizeSortV1LetterToColIndex(colInput);
  }
  if (!colIndex || colIndex < 1 || colIndex > lastCol) {
    ui.alert('Neplatny sloupec: ' + colResponse.getResponseText());
    return;
  }

  // 4. Smer razeni
  var dirResponse = ui.alert(
    'Smer razeni',
    'Vzestupne (Ano) nebo Sestupne (Ne)?',
    ui.ButtonSet.YES_NO
  );
  var ascending = (dirResponse === ui.Button.YES);

  // Spustit razeni
  SizeSortV1SortByMeasuredValue(colIndex, hasHeader, ascending, startColIdx, endColIdx);
}

/** Prevod pismena sloupce na cislo: A=1, B=2, ..., Z=26, AA=27 */
function SizeSortV1LetterToColIndex(letter) {
  if (!letter || !/^[A-Z]+$/.test(letter)) return NaN;
  var result = 0;
  for (var i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

/** Prevod cisla sloupce na pismeno: 1=A, 2=B, ..., 27=AA */
function SizeSortV1ColIndexToLetter(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Hlavni razeni - hierarchicky: nazev -> velikost
 * startColIdx/endColIdx urcuji rozsah sloupcu ktere se presouvaji (default vsechny)
 */
function SizeSortV1SortByMeasuredValue(colIndex, hasHeader, ascending, startColIdx, endColIdx) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (!startColIdx) startColIdx = 1;
  if (!endColIdx) endColIdx = sheet.getLastColumn();
  if (lastRow < 2) return;
  if (colIndex < startColIdx || colIndex > endColIdx) {
    SpreadsheetApp.getUi().alert('Sloupec k razeni (' + SizeSortV1ColIndexToLetter(colIndex) + ') neni v rozsahu ' + SizeSortV1ColIndexToLetter(startColIdx) + ':' + SizeSortV1ColIndexToLetter(endColIdx));
    return;
  }

  var startRow = hasHeader ? 2 : 1;
  var numRows = lastRow - startRow + 1;
  var numCols = endColIdx - startColIdx + 1;

  var all = sheet.getRange(startRow, startColIdx, numRows, numCols).getValues();
  // keyIdx relativne k rozsahu
  var keyIdx = colIndex - startColIdx;

  var keyed = all.map(function(row) {
    var cellText = (row[keyIdx] == null ? '' : String(row[keyIdx])).toLowerCase().trim();
    var normVal = SizeSortV1NormalizeToBaseUnit(cellText);
    var nameText = cellText.replace(/(\d+(?:[.,]\d+)?\s*(ml|l|kg|g|mm|cm|m)\b)/gi, '').replace(/[,-]+$/, '').trim();
    return {
      keyText: nameText,
      keyNum: Number.isFinite(normVal) ? normVal : null,
      row: row
    };
  });

  keyed.sort(function(a, b) {
    var textCmp = a.keyText.localeCompare(b.keyText, 'cs', { sensitivity: 'base' });
    if (textCmp !== 0) return textCmp;

    var aNum = a.keyNum, bNum = b.keyNum;
    if (aNum === null && bNum === null) return 0;
    if (aNum === null) return 1;
    if (bNum === null) return -1;
    return ascending ? (aNum - bNum) : (bNum - aNum);
  });

  var sorted = keyed.map(function(x) { return x.row; });
  sheet.getRange(startRow, startColIdx, sorted.length, numCols).setValues(sorted);
  SpreadsheetApp.getActive().toast('Razeni hotovo');
}

/**
 * Normalizace textu do zakladni jednotky:
 * objem -> ml, hmotnost -> g, delka -> mm
 */
function SizeSortV1NormalizeToBaseUnit(value) {
  if (value == null) return NaN;
  var raw = String(value).toLowerCase().trim();

  var re = /(\d+(?:[.,]\d+)?)\s*(ml|l|kg|g|mm|cm|m)\b/g;
  var match, last = null;
  while ((match = re.exec(raw)) !== null) {
    last = match;
  }
  if (!last) return NaN;

  var num = parseFloat(last[1].replace(',', '.'));
  var unit = last[2];
  if (!Number.isFinite(num)) return NaN;

  switch (unit) {
    case 'ml': return num;
    case 'l':  return num * 1000;
    case 'g':  return num;
    case 'kg': return num * 1000;
    case 'mm': return num;
    case 'cm': return num * 10;
    case 'm':  return num * 1000;
    default:   return NaN;
  }
}
