/******************************* //// https://chatgpt.com/share/e/691307d8-5b30-800b-b22f-226b14261bd2
 * SizeSortV1 – JEDEN SOUBOR
 * - Menu: Řazení dle veličiny → Seřadit…
 * - Dialog: výběr sloupce, "má hlavičku", směr řazení
 * - Řazení: nejprve podle názvu (text bez čísla+jednotky), uvnitř podle normalizované hodnoty
 * - Podporované jednotky: ml/l, g/kg, mm/cm/m
 *******************************/



/** Otevře dialog (HTML je inline ve stringu) */
function SizeSortV1OpenSortDialog() {
  const SizeSortV1Html = `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <meta charset="UTF-8">
    <style>
      body { font-family: Arial, sans-serif; padding: 14px; }
      label { display:block; margin-top:10px; }
      .SizeSortV1Row { display:flex; gap:10px; align-items:center; }
      .SizeSortV1Actions { margin-top:16px; display:flex; gap:10px; }
      select, input[type="checkbox"], input[type="radio"] { font-size:14px; }
      h3 { margin:0 0 8px 0; }
    </style>
  </head>
 <body>
    <h3>Seřadit list podle veličiny</h3>

    <label>Sloupec k řazení:</label>
    <div class="SizeSortV1Row">
      <select id="SizeSortV1Col"></select>
      <span id="SizeSortV1ColLabel"></span>
    </div>

    <label class="SizeSortV1Row">
      <input type="checkbox" id="SizeSortV1HasHeader" checked>
      Má list hlavičku (první řádek)?
    </label>
     <label>Směr řazení:</label>
    <div class="SizeSortV1Row">
      <label class="SizeSortV1Row"><input type="radio" name="SizeSortV1Dir" value="asc" checked> Vzestupně</label>
      <label class="SizeSortV1Row"><input type="radio" name="SizeSortV1Dir" value="desc"> Sestupně</label>
    </div>

    <div class="SizeSortV1Actions">
      <button onclick="SizeSortV1RunSort()">Seřadit</button>
      <button onclick="google.script.host.close()">Zavřít</button>
    </div>

    <script>
      function SizeSortV1ColIndexToLetter(n){
        let s = '';
        while(n>0){
          let r = (n-1)%26;
          s = String.fromCharCode(65+r) + s;
          n = Math.floor((n-1)/26);
        }
        return s;
      }
      function SizeSortV1InitColumnsCount(count){
        const sel = document.getElementById('SizeSortV1Col');
        sel.innerHTML = '';
        for (let i=1; i<=count; i++){
          const opt = document.createElement('option');
          opt.value = i;
          opt.textContent = SizeSortV1ColIndexToLetter(i) + ' (' + i + ')';
          sel.appendChild(opt);
        }
        SizeSortV1UpdateLabel();
        sel.addEventListener('change', SizeSortV1UpdateLabel);
      }
      function SizeSortV1UpdateLabel(){
        const sel = document.getElementById('SizeSortV1Col');
        const idx = parseInt(sel.value,10);
        document.getElementById('SizeSortV1ColLabel').textContent = 'Vybrán: ' + (isNaN(idx)?'—': (idx + ' [' + SizeSortV1ColIndexToLetter(idx) + ']'));
      }
      function SizeSortV1RunSort(){
        const SizeSortV1Col = parseInt(document.getElementById('SizeSortV1Col').value,10);
        const SizeSortV1HasHeader = document.getElementById('SizeSortV1HasHeader').checked;
        const SizeSortV1Dir = document.querySelector('input[name="SizeSortV1Dir"]:checked').value;
        const SizeSortV1Ascending = (SizeSortV1Dir === 'asc');
        if (!SizeSortV1Col || SizeSortV1Col < 1){
          alert('Vyber sloupec.');
          return;
        }
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .SizeSortV1SortByMeasuredValue(SizeSortV1Col, SizeSortV1HasHeader, SizeSortV1Ascending);
      }
      google.script.run.withSuccessHandler(SizeSortV1InitColumnsCount).SizeSortV1GetActiveSheetColumnCount();
    </script>
  </body>
</html>
  `;
  const SizeSortV1Dialog = HtmlService.createHtmlOutput(SizeSortV1Html)
    .setWidth(440)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(SizeSortV1Dialog, 'Seřadit list podle veličiny');
}

/** Počet sloupců aktivního listu (pro dialog) */
function SizeSortV1GetActiveSheetColumnCount() {
  const SizeSortV1Sheet = SpreadsheetApp.getActiveSheet();
  return SizeSortV1Sheet.getLastColumn();
}

/**
 * Hlavní řazení – hierarchicky: název → velikost
 * @param {number} SizeSortV1ColumnIndex 1 = A, 2 = B, ...
 * @param {boolean} SizeSortV1HasHeader true = první řádek je hlavička
 * @param {boolean} SizeSortV1Ascending true = vzestupně, false = sestupně
 */
function SizeSortV1SortByMeasuredValue(SizeSortV1ColumnIndex, SizeSortV1HasHeader, SizeSortV1Ascending) {
  const SizeSortV1Sheet = SpreadsheetApp.getActiveSheet();
  const SizeSortV1LastRow = SizeSortV1Sheet.getLastRow();
  const SizeSortV1LastCol = SizeSortV1Sheet.getLastColumn();
  if (SizeSortV1LastRow < 2) return;

  const SizeSortV1StartRow = SizeSortV1HasHeader ? 2 : 1;
  const SizeSortV1NumRows = SizeSortV1LastRow - SizeSortV1StartRow + 1;

  const SizeSortV1All = SizeSortV1Sheet.getRange(SizeSortV1StartRow, 1, SizeSortV1NumRows, SizeSortV1LastCol).getValues();
  const SizeSortV1KeyIndex = SizeSortV1ColumnIndex - 1;

  const SizeSortV1Keyed = SizeSortV1All.map((SizeSortV1Row) => {
    const SizeSortV1CellText = (SizeSortV1Row[SizeSortV1KeyIndex] ?? '').toString().toLowerCase().trim();
    const SizeSortV1NormVal = SizeSortV1NormalizeToBaseUnit(SizeSortV1CellText);
    // název = text bez „číslo+jednotka“ (bere poslední výskyt pro hodnotu, ale tady pro název odstraníme VŠECHNY výskyty)
    const SizeSortV1NameText = SizeSortV1CellText.replace(/(\d+(?:[.,]\d+)?\s*(ml|l|kg|g|mm|cm|m)\b)/gi, '').replace(/[,-]+$/,'').trim();
    return {
      SizeSortV1KeyText: SizeSortV1NameText,
      SizeSortV1KeyNum: Number.isFinite(SizeSortV1NormVal) ? SizeSortV1NormVal : null,
      SizeSortV1Row
    };
  });

  SizeSortV1Keyed.sort((a, b) => {
    const SizeSortV1TextCmp = a.SizeSortV1KeyText.localeCompare(b.SizeSortV1KeyText, 'cs', {sensitivity:'base'});
    if (SizeSortV1TextCmp !== 0) return SizeSortV1TextCmp; // různé názvy

    const SizeSortV1A = a.SizeSortV1KeyNum, SizeSortV1B = b.SizeSortV1KeyNum;
    if (SizeSortV1A === null && SizeSortV1B === null) return 0;
    if (SizeSortV1A === null) return 1;
    if (SizeSortV1B === null) return -1;
    return SizeSortV1Ascending ? (SizeSortV1A - SizeSortV1B) : (SizeSortV1B - SizeSortV1A);
  });

  const SizeSortV1Sorted = SizeSortV1Keyed.map(x => x.SizeSortV1Row);
  SizeSortV1Sheet.getRange(SizeSortV1StartRow, 1, SizeSortV1Sorted.length, SizeSortV1LastCol).setValues(SizeSortV1Sorted);
  SpreadsheetApp.getActive().toast('Řazení hotovo');
}

/**
 * Normalizace textu do základní jednotky:
 * - objem → ml (l * 1000)
 * - hmotnost → g (kg * 1000)
 * - délka → mm (m * 1000, cm * 10)
 * Vrací Number nebo NaN, pokud nic nenajde.
 */
function SizeSortV1NormalizeToBaseUnit(SizeSortV1Value) {
  if (SizeSortV1Value == null) return NaN;
  const SizeSortV1Raw = String(SizeSortV1Value).toLowerCase().trim();

  // Najdi POSLEDNÍ výskyt: číslo (+ desetinná tečka/čárka) + volitelná mezera + jednotka
  const SizeSortV1Re = /(\d+(?:[.,]\d+)?)\s*(ml|l|kg|g|mm|cm|m)\b/g;
  let SizeSortV1Match, SizeSortV1Last = null;
  while ((SizeSortV1Match = SizeSortV1Re.exec(SizeSortV1Raw)) !== null) {
    SizeSortV1Last = SizeSortV1Match;
  }
  if (!SizeSortV1Last) return NaN;

  const SizeSortV1Num = parseFloat(SizeSortV1Last[1].replace(',', '.'));
  const SizeSortV1Unit = SizeSortV1Last[2];

  if (!Number.isFinite(SizeSortV1Num)) return NaN;

  switch (SizeSortV1Unit) {
    // OBJEM -> ml
    case 'ml': return SizeSortV1Num;
    case 'l':  return SizeSortV1Num * 1000;

    // HMOTNOST -> g
    case 'g':  return SizeSortV1Num;
    case 'kg': return SizeSortV1Num * 1000;

    // DÉLKA -> mm
    case 'mm': return SizeSortV1Num;
    case 'cm': return SizeSortV1Num * 10;
    case 'm':  return SizeSortV1Num * 1000;

    default:   return NaN;
  }
}
