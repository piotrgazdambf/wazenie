/******************* KARTA WAZENIA (export bierze aktywny KW/KWG + dopisuje PLS + dopisuje PS) *******************
 * - Jeśli aktywny arkusz to "KW" albo "KWG" -> eksportuje aktywny.
 * - Jeśli aktywny nie jest KW/KWG -> NIE eksportuje i NIE czyści (komunikat).
 * - PRZED eksportem: waliduje czy wymagane tabelki jakości są uzupełnione (wg przeznaczenia i liczby odmian).
 * - Po eksporcie:
 *   1) dopisuje do PLS tyle wierszy ile jest odmian (odmiana + skrzynie)
 *      + dopisuje do kolumny H PEŁNĄ LINIĘ DOSTAWCY z F6 (np. "012 - Sortpak")
 *   2) dopisuje do PS parametry surowców (A:H) tyle wierszy ile jest odmian
 *
 * PS (NOWA KOLEJNOŚĆ):
 * A NR DOSTAWY
 * B ODMIANA
 * C ILOŚĆ SKRZYŃ (teraz J/K)
 * D WAGA NETTO (po zwrocie) + " kg"
 * E ZWROT (KG) = Jxx - Kxx + " kg"
 * F BRIX
 * G TWARDOŚĆ
 * H KALIBER ↓ 68mm %
 *
 * - Jeśli BRIX/TWARDOŚĆ/KALIBER/WAGA NETTO/ZWROT brak -> "ND"
 *
 * - Po udanym eksporcie czyści wskazane zakresy w aktywnym arkuszu, ale NIGDY nie usuwa formuł.
 * - Po czyszczeniu chowa WSZYSTKIE tabelki:
 *   KW : 24–46
 *   KWG: 18–39
 **************************************************************************************/

/** KONFIG – folder i MCR muszą być udostępnione każdemu kontu, które ma korzystać ze skryptu */
const KW_EXPORT_CONFIG = {
  PARENT_FOLDER_ID: "1IyBg4JjXiKx1RSH5zKOIZ9N2zlONHRhh",

  MENU_NAME: "Karta ważenia",
  MENU_ITEM: "Prześlij kartę ważenia",

  PLS_SHEET_NAME: "PLS",
  PLS_VIEW_SHEET_NAME: "PLS_VIEW",
  STANY_SUROWCOWE_SHEET_NAME: "DOSTAWCY SUROWCE",
  AKCJE_SKRZYN_SHEET_NAME: "AKCJE SKRZYN",
  STANY_SKRZYN_SHEET_NAME: "STANY SKRZYN",

  PS_SHEET_NAME: "PS",

  SUPPLIER_RANGE_A1: "F6:I6",
  DELIVERY_RANGE_A1: "F7:I7",

  DATE_FORMAT: "dd.MM.yyyy",
  TIME_FORMAT: "HH:mm",

  DELIVERY_PAD: 4,

  MCR_SPREADSHEET_ID: "1LcvYcLots1pU4uUPMexktyLcFV-NlJkbNb0tNBmb-K8",
  MCR_SHEET_NAME: "Raport Akcji Surowca",
};

/******************* JEDYNE onOpen W PROJEKCIE (ROUTER MENU) *******************/
function onOpen(e) {
  try { KW_onOpenMenu_(); } catch (err) {}
  try { if (typeof QR_onOpenMenu_ === "function") QR_onOpenMenu_(); } catch (err) {}
  try { STANY_onOpenMenu_(); } catch (err) {}
  try { WYDANIE_onOpenMenu_(); } catch (err) {}
  try { if (typeof PDKW_onOpenMenu_ === "function") PDKW_onOpenMenu_(); } catch (err) {}
}

/******************* JEDYNE onEdit W PROJEKCIE (ROUTER) *******************/
function onEdit(e) {
  try { KW_onEdit_AutoBlocks_(e); } catch (err) {}
  try { if (typeof TABLEKW_onEdit_ === "function") TABLEKW_onEdit_(e); } catch (err) {}
  try { if (typeof PDKW_WSG_onEdit_ === "function") PDKW_WSG_onEdit_(e); } catch (err) {}
}

/******************* MENU: KARTA WAŻENIA *******************/
function KW_onOpenMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(KW_EXPORT_CONFIG.MENU_NAME)
    .addItem(KW_EXPORT_CONFIG.MENU_ITEM, "KW_EXPORT_CREATE_FILE_FROM_KW_AND_SELECT_IN_VIEW")
    .addToUi();
}

// Kompatybilność: jeśli w arkuszu był stary wpis menu, przekieruj go na właściwą funkcję.
function KW_EXPORT_WITH_PROGRESS_() {
  return KW_EXPORT_CREATE_FILE_FROM_KW_AND_SELECT_IN_VIEW();
}

/******************* MENU: STANY SUROWCOWE *******************/
function STANY_onOpenMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("Stany")
    .addItem("PRZEŚLIJ DO STANÓW", "STANY_PRZESLIJ_DO_STANOW")
    .addToUi();
}

/** Osobne menu na pasku – Wydanie skrzyń (tylko arkusz STANY SKRZYN). */
function WYDANIE_onOpenMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("Wydanie")
    .addItem("Wydaj skrzynie", "STANY_WYDANIE_SKRZYN")
    .addToUi();
}

/**
 * Wydanie skrzyń: tylko w arkuszu STANY SKRZYN, zaznaczona komórka = dostawca (kolumna A).
 * Pokazuje dialog HTML: dostawca, ptaszki drewniane/plastikowe, ilości, przycisk Wydaj -> dopisuje wiersz do AKCJE SKRZYN (Zejscie).
 */
function STANY_WYDANIE_SKRZYN() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== KW_EXPORT_CONFIG.STANY_SKRZYN_SHEET_NAME) {
    ui.alert("Wydanie działa tylko w arkuszu „STANY SKRZYN”. Otwórz ten arkusz, zaznacz komórkę z dostawcą (kolumna A) i wybierz ponownie.");
    return;
  }
  const cell = sh.getActiveCell();
  if (!cell || cell.getColumn() !== 1) {
    ui.alert("Zaznacz komórkę w kolumnie A (dostawca).");
    return;
  }
  const dostawca = String(cell.getDisplayValue() || "").trim();
  if (!dostawca) {
    ui.alert("W zaznaczonej komórce brak dostawcy.");
    return;
  }
  const html = STANY_WYDANIE_HTML_(dostawca);
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(440).setHeight(380), "Wydanie skrzyń");
}

function STANY_WYDANIE_HTML_(dostawca) {
  const escaped = (dostawca || "").replace(/\\/g, "\\\\").replace(/"/g, "&quot;").replace(/'/g, "&#39;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/&/g, "&amp;");
  const displayName = (dostawca || "").replace(/&/g, "&amp;");
  return '<!DOCTYPE html><html><head><meta charset="utf-8"/>' +
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{margin:0;font-family:"Segoe UI",system-ui,-apple-system,sans-serif;background:linear-gradient(145deg,#f0f4f8 0%,#e2e8f0 100%);min-height:100vh;padding:24px;display:flex;align-items:center;justify-content:center;color:#1a202c}' +
    '.card{background:#fff;border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,.08),0 2px 8px rgba(0,0,0,.04);overflow:hidden;max-width:400px;width:100%}' +
    '.card-header{background:linear-gradient(135deg,#1e3a5f 0%,#2d5a87 100%);color:#fff;padding:20px 24px;font-size:15px;font-weight:600;line-height:1.4;letter-spacing:.02em}' +
    '.card-header small{display:block;font-size:10px;font-weight:500;text-transform:uppercase;letter-spacing:.12em;opacity:.85;margin-bottom:6px}' +
    '.card-body{padding:28px 24px}' +
    '.label-top{display:block;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#718096;margin-bottom:16px}' +
    '.row{margin-bottom:20px}' +
    '.row:last-of-type{margin-bottom:0}' +
    'label{display:flex;align-items:center;gap:10px;cursor:pointer;font-size:14px;color:#2d3748}' +
    'input[type=checkbox]{width:18px;height:18px;accent-color:#2d5a87;cursor:pointer}' +
    'input[type=number]{width:80px;padding:10px 12px;border:1px solid #e2e8f0;border-radius:8px;font-size:15px;font-weight:500;text-align:right;margin-left:auto}' +
    'input[type=number]:focus{outline:none;border-color:#2d5a87;box-shadow:0 0 0 3px rgba(45,90,135,.15)}' +
    '.btn{width:100%;margin-top:28px;padding:14px 24px;background:linear-gradient(135deg,#1e3a5f 0%,#2d5a87 100%);color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;letter-spacing:.03em;transition:transform .05s ease,box-shadow .2s ease}' +
    '.btn:hover{box-shadow:0 6px 20px rgba(45,90,135,.35);transform:translateY(-1px)}' +
    '.btn:active{transform:translateY(0)}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<div class="card-header"><small>Dostawca</small>' + displayName + '</div>' +
    '<div class="card-body">' +
    '<span class="label-top">Rodzaj i ilość skrzyń</span>' +
    '<input type="hidden" id="dostawca" value="' + escaped + '"/>' +
    '<div class="row"><label><input type="checkbox" id="cbDrewniane"/> Skrzynie drewniane</label><input type="number" id="qtyDrewniane" min="0" value="0"/></div>' +
    '<div class="row"><label><input type="checkbox" id="cbPlastikowe"/> Skrzynie plastikowe</label><input type="number" id="qtyPlastikowe" min="0" value="0"/></div>' +
    '<button type="button" class="btn" id="btnWydaj">Wydaj</button>' +
    '</div></div>' +
    '<script>' +
    'document.getElementById("btnWydaj").onclick=function(){' +
    'var d=document.getElementById("dostawca").value;' +
    'var qD=document.getElementById("cbDrewniane").checked?parseInt(document.getElementById("qtyDrewniane").value,10)||0:0;' +
    'var qP=document.getElementById("cbPlastikowe").checked?parseInt(document.getElementById("qtyPlastikowe").value,10)||0:0;' +
    'if(qD===0&&qP===0){alert("Zaznacz rodzaj skrzyń i wpisz ilość.");return;}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).withFailureHandler(function(e){alert("Błąd: "+(e&&e.message?e.message:String(e)));}).STANY_WYDANIE_APPEND_ROW(d,qD,qP);' +
    '};</script></body></html>';
}

/** Dopisuje wiersz do AKCJE SKRZYN: A=dostawca, B=timestamp, C=Zejscie, D=ilość drewnianych, E=ilość plastikowych. */
function STANY_WYDANIE_APPEND_ROW(dostawca, qtyDrewniane, qtyPlastikowe) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KW_EXPORT_CONFIG.AKCJE_SKRZYN_SHEET_NAME);
  if (!sh) throw new Error('Brak arkusza "' + KW_EXPORT_CONFIG.AKCJE_SKRZYN_SHEET_NAME + '".');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Warsaw", "dd.MM.yyyy HH:mm");
  const qD = Math.max(0, parseInt(qtyDrewniane, 10) || 0);
  const qP = Math.max(0, parseInt(qtyPlastikowe, 10) || 0);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 5).setValues([["DOSTAWCA", "TIMESTAMP", "AKCJA", "ILOŚĆ SKRZYŃ DREWNIANYCH", "ILOŚĆ SKRZYŃ PLASTIKOWYCH"]]);
    sh.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  const nextRow = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(nextRow, 1, 1, 5).setValues([[String(dostawca || "").trim(), timestamp, "Zejscie", qD, qP]]);
  SpreadsheetApp.flush();
}

/******************* AUTO SHOW/HIDE BLOKÓW (KW + KWG) *******************/
function KW_onEdit_AutoBlocks_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!sh) return;

  const name = sh.getName();
  if (name !== "KW" && name !== "KWG") return;

  KW_APPLY_BLOCK_VISIBILITY_(sh);
}

function KW_APPLY_BLOCK_VISIBILITY_(sh) {
  const sheetName = sh.getName();
  const layout = typeof getLayout_ === "function" ? getLayout_(sheetName) : null;
  const purposeShort = typeof getPurposeShortFromLot_ === "function" ? getPurposeShortFromLot_(sh) : "";
  const blocks = typeof getEffectiveBlockShowHideRanges_ === "function"
    ? getEffectiveBlockShowHideRanges_(sheetName, purposeShort)
    : (layout ? layout.blockShowHideRanges : null);
  if (!layout || !blocks || blocks.length < 4) return;

  const hasAny = (a1) => {
    const vals = sh.getRange(a1).getDisplayValues();
    for (let r = 0; r < vals.length; r++) {
      for (let c = 0; c < vals[r].length; c++) {
        if (String(vals[r][c] || "").trim() !== "") return true;
      }
    }
    return false;
  };

  const nonEmpty = (a1) => String(sh.getRange(a1).getDisplayValue() || "").trim() !== "";

  const showHide = (fromRow, toRow, shouldShow) => {
    const count = toRow - fromRow + 1;
    if (shouldShow) {
      try { sh.showRows(fromRow, count); } catch (e) { if (e && (e.message || e.toString)) Logger.log("KW_APPLY_BLOCK_VISIBILITY showRows: " + (e.message || e.toString())); }
    } else {
      try { sh.hideRows(fromRow, count); } catch (e) { if (e && (e.message || e.toString)) Logger.log("KW_APPLY_BLOCK_VISIBILITY hideRows: " + (e.message || e.toString())); }
    }
  };

  const vf = layout.varietyRowFirst;
  const show1 = nonEmpty("J" + vf) || nonEmpty("K" + vf);
  const show2 = hasAny("E" + (vf + 1) + ":G" + (vf + 1));
  const show3 = hasAny("E" + (vf + 2) + ":G" + (vf + 2));
  const show4 = hasAny("E" + (vf + 3) + ":G" + (vf + 3));

  showHide(blocks[0][0], blocks[0][1], show1);
  showHide(blocks[1][0], blocks[1][1], show2);
  showHide(blocks[2][0], blocks[2][1], show3);
  showHide(blocks[3][0], blocks[3][1], show4);

  SpreadsheetApp.flush();
}

/******************* WALIDACJA TABEL JAKOŚCI PRZED EKSPORTEM *******************
 * Teraz: wartości w E muszą być LICZBĄ (może być 0, może być po przecinku , lub .)
 ********************************************************************************/

/** true jeśli tekst to liczba (np "0", "12", "12,5", "12.5") */
function isNumericInput_(v) {
  const s = String(v == null ? "" : v).trim();
  if (s === "") return false;
  const norm = s.replace(/\s+/g, "").replace(",", ".");
  // tylko liczba (opcjonalnie ujemna, opcjonalnie część dziesiętna)
  return /^-?\d+(\.\d+)?$/.test(norm);
}

function KW_VALIDATE_BEFORE_EXPORT_(kw) {
  const ui = SpreadsheetApp.getUi();
  const sheetName = kw.getName();

  const lotText = String(kw.getRange("F7:I7").getDisplayValue() || "").trim();
  const purpose = getPurposeShortFromLot_(kw); // "P","S","O",...
  const p = String(purpose || "").toUpperCase().trim();

  // ile pól wymagamy w tabelce jakości (kolumna E)
  // WYJĄTEK: KWG + zaznaczony J3 (RYLEX) lub K3 (GRÓJECKA) w WSG:
  // wymagamy tylko BRIX i TWARDOŚĆ (2 pola): E19/E20, E25/E26, E31/E32, E37/E38
  let requiredCount = 2;
  let forceSimpleKWGValidation = false;
  if (sheetName === "KWG") {
    try {
      const wsg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WSG");
      if (wsg) forceSimpleKWGValidation = !!wsg.getRange("J3").getValue() || !!wsg.getRange("K3").getValue();
    } catch (e) {
      if (e && (e.message || e.toString)) Logger.log("KW_VALIDATE_BEFORE_EXPORT_ WSG flags: " + (e.message || e.toString()));
    }
  }
  if (!forceSimpleKWGValidation) {
    if (p === "S") requiredCount = 3;
    else if (p === "O") requiredCount = 4;
    else if (p === "P") requiredCount = 2;
    else requiredCount = 2;
  } else {
    // KWG + RYLEX/GRÓJECKA:
    // P -> tylko BRIX (1 pole)
    // S/O -> BRIX + TWARDOŚĆ (2 pola)
    requiredCount = (p === "P") ? 1 : 2;
  }

  const layout = typeof getLayout_ === "function" ? getLayout_(sheetName) : null;
  const starts = layout
    ? (typeof getEffectiveQualityStarts_ === "function" ? getEffectiveQualityStarts_(sheetName, p) : layout.qualityStarts)
    : (sheetName === "KWG" ? [19, 25, 31, 37] : [24, 30, 36, 42]);

  const vf = layout ? layout.varietyRowFirst : (sheetName === "KW" ? 19 : 12);
  const jVf = String(kw.getRange("J" + vf).getDisplayValue() || "").trim() !== "";
  const kVf = String(kw.getRange("K" + vf).getDisplayValue() || "").trim() !== "";
  const mustFillFlags = [
    jVf || kVf,
    String(kw.getRange("E" + (vf + 1)).getDisplayValue() || "").trim() !== "",
    String(kw.getRange("E" + (vf + 2)).getDisplayValue() || "").trim() !== "",
    String(kw.getRange("E" + (vf + 3)).getDisplayValue() || "").trim() !== "",
  ];

  if (!lotText) return true;

  const missing = [];
  const notNumeric = [];

  for (let i = 0; i < 4; i++) {
    if (!mustFillFlags[i]) continue;

    const startRow = starts[i];
    for (let k = 0; k < requiredCount; k++) {
      const a1 = `E${startRow + k}`;
      const v = String(kw.getRange(a1).getDisplayValue() || "").trim();

      if (v === "") {
        missing.push(a1);
        continue;
      }
      if (!isNumericInput_(v)) {
        notNumeric.push(a1);
      }
    }
  }

  if (missing.length || notNumeric.length) {
    let msg = "Nie można przesłać nieuzupełnionej karty ważenia.\n\n";

    if (missing.length) {
      msg += "Uzupełnij wymagane pola (kolumna E):\n" + missing.join(", ") + "\n\n";
    }
    if (notNumeric.length) {
      msg += "Te pola muszą być liczbą (np. 0, 12, 12,5):\n" + notNumeric.join(", ");
    }

    ui.alert(msg.trim());
    return false;
  }

  return true;
}

/******************* NOWE: WALIDACJA SKRZYŃ J/K PRZED EKSPORTEM *******************
 * - jeśli jest odmiana -> oba pola (J i K) muszą być uzupełnione liczbą
 ********************************************************************************/
function KW_VALIDATE_CRATES_JK_BEFORE_EXPORT_(kw) {
  const ui = SpreadsheetApp.getUi();
  const sheetName = kw.getName();
  const layout = typeof getLayout_ === "function" ? getLayout_(sheetName) : null;
  const vf = layout ? layout.varietyRowFirst : (sheetName === "KW" ? 19 : 12);

  const varCells = ["E" + vf, "E" + (vf + 1), "E" + (vf + 2), "E" + (vf + 3)];
  const jCells   = ["J" + vf, "J" + (vf + 1), "J" + (vf + 2), "J" + (vf + 3)];
  const kCells   = ["K" + vf, "K" + (vf + 1), "K" + (vf + 2), "K" + (vf + 3)];

  const bad = [];

  for (let i = 0; i < 4; i++) {
    const variety = String(kw.getRange(varCells[i]).getDisplayValue() || "").trim();
    if (!variety) continue;

    const jv = String(kw.getRange(jCells[i]).getDisplayValue() || "").trim();
    const kv = String(kw.getRange(kCells[i]).getDisplayValue() || "").trim();

    const jEmpty = (jv === "");
    const kEmpty = (kv === "");

    // muszą iść w parze: oba albo żadne (ale przy odmianie: wymagamy oba)
    if (jEmpty || kEmpty) {
      bad.push(`${jCells[i]}/${kCells[i]}`);
      continue;
    }

    // oba muszą być liczbą
    if (!isNumericInput_(jv) || !isNumericInput_(kv)) {
      bad.push(`${jCells[i]}/${kCells[i]}`);
      continue;
    }
  }

  if (bad.length) {
    ui.alert(
      "Błąd: Ilość skrzyń musi być podana jako PARA J/K.\n\n" +
      "Dla każdej odmiany uzupełnij obie komórki (J oraz K) liczbą.\n" +
      "Problemy w: " + bad.join(", ")
    );
    return false;
  }

  return true;
}

/******************* GŁÓWNA FUNKCJA *******************/
function KW_EXPORT_CREATE_FILE_FROM_KW_AND_SELECT_IN_VIEW() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    ui.alert("Skrypt jest w trakcie wykonywania. Spróbuj ponownie za chwilę.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone() || "Europe/Warsaw";

    let kw = ss.getActiveSheet();
    const activeName = kw ? kw.getName() : "";

    if (!(activeName === "KW" || activeName === "KWG")) {
      ui.alert("Info", "Nie jesteś w Karcie Ważenia (KW/KWG). Przejdź do arkusza KW albo KWG i spróbuj ponownie.", ui.ButtonSet.OK);
      return;
    }
    if (!kw) throw new Error('Brak arkusza źródłowego (KW/KWG).');

    // >>> blokada wysyłki jeśli tabelki jakości nie są uzupełnione / nie są liczbami
    if (!KW_VALIDATE_BEFORE_EXPORT_(kw)) return;

    // >>> NOWE: blokada jeśli skrzynie nie są w parach J/K
    if (!KW_VALIDATE_CRATES_JK_BEFORE_EXPORT_(kw)) return;

    const srcMaxRows = kw.getMaxRows();
    const srcMaxCols = kw.getMaxColumns();

    // 1) Dostawca
    const supplierRaw = String(kw.getRange(KW_EXPORT_CONFIG.SUPPLIER_RANGE_A1).getDisplayValue() || "").trim();
    const supplierParsed = parseSupplier_(supplierRaw);
    const supplierNo = supplierParsed.no || "???";
    const supplierName = supplierParsed.name || "BRAK_DOSTAWCY";

    // 2) Numer dostawy (3–6 znak z F7:I7)
    const deliveryRaw = String(kw.getRange(KW_EXPORT_CONFIG.DELIVERY_RANGE_A1).getDisplayValue() || "").trim();
    const deliveryNo = extractDeliveryNo_3to6_(deliveryRaw, KW_EXPORT_CONFIG.DELIVERY_PAD);

    // 3) Timestamp
    const now = new Date();
    const dateStr = Utilities.formatDate(now, tz, KW_EXPORT_CONFIG.DATE_FORMAT);
    const timeStr = Utilities.formatDate(now, tz, KW_EXPORT_CONFIG.TIME_FORMAT);

    // 4) Nazwa logiczna
    const logicalName = `${deliveryNo}/${supplierNo} - ${supplierName} ${dateStr} ${timeStr}`
      .replace(new RegExp("\\s+", "g"), " ")
      .trim();

    const fileName = sanitizeFileName_(logicalName);

    // 6) Utwórz plik i przenieś do folderu miesiąca (wymaga udostępnienia folderu wszystkim użytkownikom skryptu)
    let parentFolder, monthFolder;
    try {
      parentFolder = DriveApp.getFolderById(KW_EXPORT_CONFIG.PARENT_FOLDER_ID);
      monthFolder = getOrCreateMonthFolder_(parentFolder, now);
    } catch (e) {
      ui.alert(
        "Brak uprawnień",
        "Ten skrypt wymaga dostępu do folderu na Dysku Google (eksport Kart Ważenia).\n\nPoproś właściciela arkusza o udostępnienie folderu eksportu Twojemu kontu z uprawnieniem do dodawania plików.",
        ui.ButtonSet.OK
      );
      return;
    }

    const newSS = SpreadsheetApp.create(fileName);
    const newFile = DriveApp.getFileById(newSS.getId());

    monthFolder.addFile(newFile);
    try { DriveApp.getRootFolder().removeFile(newFile); } catch (e) { if (e && (e.message || e.toString)) Logger.log("KW_EXPORT removeFile: " + (e.message || e.toString())); }

    // 7) Skopiuj KW/KWG do nowego pliku
    const copied = kw.copyTo(newSS);
    copied.setName(kw.getName());
    newSS.setActiveSheet(copied);

    syncSheetGridSize_(copied, srcMaxRows, srcMaxCols);

    const all = newSS.getSheets();
    for (const sh of all) {
      if (sh.getSheetId() !== copied.getSheetId()) newSS.deleteSheet(sh);
    }

    SpreadsheetApp.flush();

    const exportUrl = newSS.getUrl();

    // 8) PLS
    // NR DOSTAWY do PLS bierzemy z WSG!C4 (po sekwencji), z fallbackiem do deliveryNo
    let deliveryNoForPLS = String(deliveryNo || "").trim();
    try {
      const shWSG = ss.getSheetByName("WSG");
      if (shWSG) {
        const c4Raw = String(shWSG.getRange("C4").getDisplayValue() || "").trim();
        const c4Digits = c4Raw.replace(/[^\d]/g, "");
        if (c4Digits) deliveryNoForPLS = String(parseInt(c4Digits, 10));
      }
    } catch (e) {
      if (e && (e.message || e.toString)) Logger.log("KW_EXPORT deliveryNoForPLS from WSG C4: " + (e.message || e.toString()));
    }

    const firstLotText = appendVarietyRowsToPLS_(ss, kw, exportUrl, deliveryNoForPLS, supplierName);

    // 9) PS
    appendVarietyRowsToPS_(ss, kw);

    // 10) PLS_VIEW
    if (firstLotText) {
      selectLotInPLSView_(ss, firstLotText);
    } else {
      const shView = ss.getSheetByName(KW_EXPORT_CONFIG.PLS_VIEW_SHEET_NAME);
      if (shView) ss.setActiveSheet(shView);
    }

    // CZYSZCZENIE: bez usuwania formuł
    clearAfterExport_KeepFormulas_(kw);

    // schowaj po eksporcie
    let tablesHidden = false;
    if (typeof HIDE_ALL_BLOCKS_AFTER_EXPORT_ === "function") {
      tablesHidden = !!HIDE_ALL_BLOCKS_AFTER_EXPORT_(kw);
    }

    if (typeof adjustAllQualityTables_ === "function") {
      adjustAllQualityTables_(kw);
    }
    if (typeof adjustRightPriceRows_ === "function") {
      adjustRightPriceRows_(kw);
    }

    ui.alert(
      "Utworzono i zapisano Kartę Ważenia:\n" +
      logicalName + "\n\n" +
      "Przesłano do folderu miesiąca: " + monthFolder.getName() + "\n" +
      "Dopisano wiersze do PLS (odmiany + skrzynie) i podlinkowano.\n" +
      "Dopisano wiersze do PS (parametry surowców).\n\n" +
      "Wyczyszczono dane w arkuszu " + kw.getName() + " (bez usuwania formuł)" +
      (tablesHidden ? " i schowano tabelki." : " (tabelki NIE zostały schowane).")
    );

  } catch (err) {
    SpreadsheetApp.getUi().alert("Błąd: " + (err && err.message ? err.message : String(err)));
  } finally {
    try { lock.releaseLock(); } catch (e) { if (e && (e.message || e.toString)) Logger.log("KW_EXPORT releaseLock: " + (e.message || e.toString())); }
  }
}

/**
 * ===================== PLS =====================
 * Kolumny PLS:
 * A LOT
 * B Data (F5 display)
 * C Nr dostawy (deliveryNo)
 * D Dostawca (nazwa jak wcześniej)
 * E Przeznaczenie (pełna nazwa)
 * F Odmiana
 * G Ilość skrzyń (TERAZ: J/K)
 * H PEŁNA LINIA DOSTAWCY z F6:I6 (np. "012 - Sortpak")
 */
function appendVarietyRowsToPLS_(ss, kw, exportUrl, deliveryNo, supplierName) {
  const shPLS = ss.getSheetByName(KW_EXPORT_CONFIG.PLS_SHEET_NAME);
  if (!shPLS) throw new Error('Brak arkusza: "' + KW_EXPORT_CONFIG.PLS_SHEET_NAME + '"');

  const lotText = String(kw.getRange("F7:I7").getDisplayValue() || "").trim();
  if (!lotText) return "";

  const purposeName = purposeFromLotFull_(lotText);
  const dateText = String(kw.getRange("F5").getDisplayValue() || "").trim();

  const sheetName = kw.getName();
  const layout = typeof getLayout_ === "function" ? getLayout_(sheetName) : null;
  const purposeShort = getPurposeShortFromLot_(kw, lotText);
  const qs = layout
    ? (typeof getEffectiveQualityStarts_ === "function" ? getEffectiveQualityStarts_(sheetName, purposeShort) : layout.qualityStarts)
    : (sheetName === "KW" ? [24, 30, 36, 42] : [19, 25, 31, 37]);
  const vf = layout ? layout.varietyRowFirst : (sheetName === "KW" ? 19 : 12);
  const nettoAfterCells = qs.map(s => "K" + (s + 1));
  // Owoc z WSG (F4) – zapisywany w M1 przez PDKW; E19/G19 bywa nadpisane przez odmianę (np. Ligol)
  let owoc = String(kw.getRange("M1").getDisplayValue() || "").trim();
  if (!owoc) {
    const fruitRangeA1 = "E" + vf + ":G" + vf;
    owoc = String(kw.getRange(fruitRangeA1).getDisplayValue() || "").trim();
  }

  const supplierFull = String(kw.getRange(KW_EXPORT_CONFIG.SUPPLIER_RANGE_A1).getDisplayValue() || "").trim();
  const supplierNameOnly = String(supplierName || "").trim();

  const pairsWithNet = [];
  for (let i = 0; i < 4; i++) {
    const r = vf + i;
    const variety = String(kw.getRange("E" + r).getDisplayValue() || "").trim();
    if (!variety) continue;
    const crates = cratesForExport_(kw, r);
    const afterRaw = String(kw.getRange(nettoAfterCells[i]).getDisplayValue() || "").trim();
    const afterNum = extractNumber_(afterRaw);
    const wagaNetto = afterNum ? (normalizeIntKg_(afterNum) + " kg") : "";
    pairsWithNet.push({ variety, crates, wagaNetto });
  }
  if (!pairsWithNet.length) return "";

  const pairCount = pairsWithNet.length;
  const normalizeLotForPLS_ = (s) => String(s || "").trim().replace(/\s*-\s*/g, "-");
  const lotNormalized = normalizeLotForPLS_(lotText);
  const lotForVariety = (baseLot, index) => {
    if (pairCount <= 1 || index === 0) return baseLot;
    return baseLot + String(index + 1);
  };
  const allLotsForRemoval = (baseLot) => {
    if (!baseLot) return [];
    if (pairCount <= 1) return [baseLot];
    const arr = [baseLot];
    for (let i = 1; i < pairCount; i++) arr.push(baseLot + String(i + 1));
    return arr;
  };

  removeEmptyVarietyRowsForLotsFromPLS_(shPLS, allLotsForRemoval(lotNormalized));

  const lastRow = Math.max(shPLS.getLastRow(), 2);
  let targetRow = (lastRow < 3) ? 3 : (lastRow + 1);

  const buildRow = (lot, p) => [
    lot,
    dateText,
    String(deliveryNo || "").trim(),
    supplierNameOnly,
    String(purposeName || "").trim(),
    String(p.variety || "").trim(),
    String(p.crates || "").trim(),
    (supplierFull || supplierNameOnly),
    String(p.wagaNetto || "").trim(),
    owoc
  ];

  const rowsMain = pairsWithNet.map((p, i) => buildRow(lotForVariety(lotNormalized, i), p));
  const rows = rowsMain;

  const numRows = rows.length;
  shPLS.getRange(targetRow, 1, numRows, 10).setValues(rows);

  for (let i = 0; i < numRows; i++) {
    const cell = shPLS.getRange(targetRow + i, 1);
    const txt = String(rows[i][0] || "").trim();
    if (txt) setRichLink_(cell, txt, exportUrl);
  }

  SpreadsheetApp.flush();
  return lotNormalized;
}

/**
 * FIX (NOWE): usuwa z PLS puste/placeholderowe wiersze dla podanych LOT-ów:
 * - A (LOT) == lotText
 * - F (odmiana) == "" oraz G (skrzynie) == ""
 */
function removeEmptyVarietyRowsForLotsFromPLS_(shPLS, lotTexts) {
  const lastRow = shPLS.getLastRow();
  if (lastRow < 3) return;

  const vals = shPLS.getRange(3, 1, lastRow - 2, 7).getDisplayValues(); // A:G
  const needles = (lotTexts || []).map(x => String(x || "").trim()).filter(Boolean);
  if (!needles.length) return;

  const rowsToDelete = [];

  for (let i = 0; i < vals.length; i++) {
    const lot = String(vals[i][0] || "").trim();
    if (needles.indexOf(lot) === -1) continue;

    const variety = String(vals[i][5] || "").trim(); // F
    const crates  = String(vals[i][6] || "").trim(); // G

    if (variety === "" && crates === "") {
      rowsToDelete.push(3 + i);
    }
  }

  rowsToDelete.sort((a, b) => b - a).forEach(r => {
    try { shPLS.deleteRow(r); } catch (e) { if (e && (e.message || e.toString)) Logger.log("removeEmptyVarietyRowsForLotsFromPLS_ deleteRow: " + (e.message || e.toString())); }
  });

  SpreadsheetApp.flush();
}

/** Mapowanie owocu na kolumnę startową (1-based) w DOSTAWCY SUROWCE. Nowy owoc = 29 (AC). */
var STANY_OWOC_COLUMNS = {
  "jabłko": 1, "jablko": 1,
  "gruszka": 5,
  "wiśnia": 9, "wisnia": 9,
  "rabarbar": 13,
  "truskawka": 17,
  "marchewka": 21,
  "mango": 25
};
var STANY_NOWY_OWOC_START_COL = 29;

function STANY_getStartColForOwoc_(owoc) {
  const key = String(owoc || "").trim().toLowerCase().replace(/ł/g, "l").replace(/ś/g, "s").replace(/ę/g, "e").replace(/ą/g, "a").replace(/ń/g, "n").replace(/ź|ż/g, "z");
  return STANY_OWOC_COLUMNS[key] != null ? STANY_OWOC_COLUMNS[key] : STANY_NOWY_OWOC_START_COL;
}

function STANY_getNextRowInBlock_(sheet, startCol) {
  let lastDataRow = 2;
  const maxR = Math.max(sheet.getLastRow(), 3);
  for (let r = 3; r <= maxR; r++) {
    const v = String(sheet.getRange(r, startCol).getDisplayValue() || "").trim();
    if (v !== "") lastDataRow = r;
  }
  return lastDataRow + 1;
}

/**
 * PRZEŚLIJ DO STANÓW: z zaznaczonej komórki w kolumnie A w PLS lub PLS_VIEW
 * dopisuje wiersz do DOSTAWCY SUROWCE (odmiana, waga netto, dostawca), wysyła do MCR i ustawia STATUS w K na "PRZESŁANO".
 */
function STANY_PRZESLIJ_DO_STANOW() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (!sh) {
    ui.alert("Brak aktywnego arkusza.");
    return;
  }
  const sheetName = sh.getName();
  if (sheetName !== KW_EXPORT_CONFIG.PLS_SHEET_NAME && sheetName !== KW_EXPORT_CONFIG.PLS_VIEW_SHEET_NAME) {
    ui.alert("Prześlij do stanów tylko z arkusza PLS lub PLS_VIEW. Zaznacz komórkę w kolumnie A (wiersz z danymi) i wybierz ponownie.");
    return;
  }
  const cell = sh.getActiveCell();
  if (!cell) {
    ui.alert("Zaznacz komórkę w kolumnie A w wierszu, który chcesz przesłać do Stanów.");
    return;
  }
  const dataRow = cell.getRow();
  const col = cell.getColumn();
  if (col !== 1) {
    ui.alert("Zaznacz komórkę w kolumnie A (LOT).");
    return;
  }
  const maxCol = Math.max(sh.getLastColumn(), 11);
  const rowVals = sh.getRange(dataRow, 1, 1, maxCol).getDisplayValues()[0];
  const lotQR = String(rowVals[0] || "").trim();
  const statusK = String(rowVals[10] || "").trim();
  if (statusK === "PRZESŁANO") {
    ui.alert("BŁĄD: TE STANY ZOSTAŁY JUŻ PRZESŁANE.", "Prześlij do Stanów", ui.ButtonSet.OK);
    return;
  }
  const odmiana = String(rowVals[5] || "").trim();
  const dostawca = String(rowVals[7] || "").trim();
  const wagaNetto = String(rowVals[8] || "").trim();
  const przeznaczenie = String(rowVals[4] || "").trim();
  let owoc = String(rowVals[9] || "").trim();
  if (!odmiana) {
    ui.alert("W tym wierszu brak odmiany (kolumna F). Uzupełnij PLS i spróbuj ponownie.");
    return;
  }
  if (!owoc) owoc = odmiana;

  // W którym wierszu w PLS wpisać STATUS: z PLS = ten wiersz, z PLS_VIEW = szukamy tego samego LOT w PLS (wiersz na dole)
  let shPLS = null;
  let plsRow = 0;
  const normalizeLotForCompare_ = (x) => {
    // przy importRange mogą dojść różnice w białych znakach / spacji przed myślnikiem
    return String(x || "")
      .trim()
      .replace(/\u00A0/g, " ")
      .replace(/\s*-\s*/g, "-");
  };
  if (sheetName === KW_EXPORT_CONFIG.PLS_SHEET_NAME) {
    shPLS = sh;
    plsRow = dataRow;
  } else {
    shPLS = ss.getSheetByName(KW_EXPORT_CONFIG.PLS_SHEET_NAME);
    if (shPLS && lotQR) {
      const lastR = shPLS.getLastRow();
      const dataStart = 4;
      if (lastR >= dataStart) {
        const colA = shPLS.getRange(dataStart, 1, lastR, 1).getDisplayValues();
        const needle = normalizeLotForCompare_(lotQR);
        for (let r = colA.length - 1; r >= 0; r--) {
          const cellLot = normalizeLotForCompare_(colA[r][0]);
          if (cellLot === needle) {
            plsRow = dataStart + r;
            break;
          }
        }
      }
    }
  }
  const startCol = STANY_getStartColForOwoc_(owoc);
  const shDostawcy = ss.getSheetByName(KW_EXPORT_CONFIG.STANY_SUROWCOWE_SHEET_NAME) || ss.insertSheet(KW_EXPORT_CONFIG.STANY_SUROWCOWE_SHEET_NAME);
  if (startCol === STANY_NOWY_OWOC_START_COL) {
    const ac2 = String(shDostawcy.getRange(2, 29).getDisplayValue() || "").trim();
    if (!ac2) {
      shDostawcy.getRange(1, 30, 1, 2).merge().setValue(owoc);
      shDostawcy.getRange(2, 29, 1, 3).setValues([[owoc, "Ilość", "Dostawca"]]);
    }
  }
  const insertRowD = STANY_getNextRowInBlock_(shDostawcy, startCol);
  shDostawcy.getRange(insertRowD, startCol, 1, 3).setValues([[odmiana, wagaNetto, dostawca]]);

  SpreadsheetApp.flush();

  // Raport Akcji Surowca: dopisz wiersz do zewnętrznego arkusza Magazyn Centralny Roboczy (MCR)
  const timestampMCR = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Warsaw", "dd.MM.yyyy HH:mm");
  const wagaNettoMCR = (function () {
    const raw = String(wagaNetto || "").trim().replace(/\s*kg\s*$/i, "");
    const n = parseFloat(String(extractNumber_(raw) || "").replace(",", "."));
    return (!isNaN(n) && isFinite(n)) ? n : "";
  })();
  try {
    const ssMCR = SpreadsheetApp.openById(KW_EXPORT_CONFIG.MCR_SPREADSHEET_ID);
    const shMCR = ssMCR.getSheetByName(KW_EXPORT_CONFIG.MCR_SHEET_NAME);
    if (shMCR) {
      const lastRowMCR = Math.max(shMCR.getLastRow(), 1);
      shMCR.getRange(lastRowMCR + 1, 1, 1, 7).setValues([[lotQR, timestampMCR, "Przyjecie", wagaNettoMCR, owoc, odmiana, przeznaczenie]]);
    } else {
      Logger.log("STANY MCR: brak arkusza o nazwie " + KW_EXPORT_CONFIG.MCR_SHEET_NAME);
    }
  } catch (e) {
    if (e && (e.message || e.toString)) Logger.log("STANY MCR append: " + (e.message || e.toString()));
    SpreadsheetApp.getUi().alert(
      "Brak uprawnień",
      "Dodano wiersz do Stanów w tym arkuszu, ale nie udało się dopisać do Raportu Akcji Surowca (MCR).\n\nPoproś właściciela arkusza o udostępnienie zewnętrznego pliku „Raport Akcji Surowca” Twojemu kontu.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }

  // STATUS zawsze w PLS, w kolumnie K wiersza z tym LOT (z PLS_VIEW szukamy tego samego wiersza w PLS)
  if (shPLS && plsRow) shPLS.getRange("K" + plsRow).setValue("PRZESŁANO");

  // AKCJE SKRZYN: A=dostawca, B=TIMESTAMP (dd.MM.yyyy godzina), C=AKCJA, D=skrzynie drewniane, E=skrzynie plastikowe
  const cratesG = String(rowVals[6] || "").trim();
  const slashIdx = cratesG.indexOf("/");
  const skrzynieDrewniane = slashIdx >= 0 ? cratesG.substring(0, slashIdx).trim() : cratesG;
  const skrzyniePlastikowe = slashIdx >= 0 ? cratesG.substring(slashIdx + 1).trim() : "";
  const timestampAkcje = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Warsaw", "dd.MM.yyyy HH:mm");
  const shAkcjeSkrzyn = ss.getSheetByName(KW_EXPORT_CONFIG.AKCJE_SKRZYN_SHEET_NAME) || ss.insertSheet(KW_EXPORT_CONFIG.AKCJE_SKRZYN_SHEET_NAME);
  if (shAkcjeSkrzyn.getLastRow() < 1) {
    shAkcjeSkrzyn.getRange(1, 1, 1, 5).setValues([["DOSTAWCA", "TIMESTAMP", "AKCJA", "ILOŚĆ SKRZYŃ DREWNIANYCH", "ILOŚĆ SKRZYŃ PLASTIKOWYCH"]]);
    shAkcjeSkrzyn.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  const nextRowAkcje = Math.max(shAkcjeSkrzyn.getLastRow(), 1) + 1;
  shAkcjeSkrzyn.getRange(nextRowAkcje, 1, 1, 5).setValues([[dostawca, timestampAkcje, "Przyjecie", skrzynieDrewniane, skrzyniePlastikowe]]);

  ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange("A" + dataRow));
  ss.toast("Dodano do Stanów: " + odmiana + " (" + owoc + ")", "Stany", 4);
}

/**
 * Pobiera pary (odmiana, skrzynie) z KW/KWG
 * NOWE: skrzynie są jako "J/K" (np. "81/20")
 */
function getVarietyCratesPairs_(kw) {
  const name = kw.getName();
  const layout = typeof getLayout_ === "function" ? getLayout_(name) : null;
  const vf = layout ? layout.varietyRowFirst : (name === "KW" ? 19 : 12);
  const out = [];

  for (let i = 0; i < 4; i++) {
    const r = vf + i;
    const variety = String(kw.getRange("E" + r).getDisplayValue() || "").trim();
    if (!variety) continue;

    const crates = cratesForExport_(kw, r);
    out.push({ variety, crates });
  }
  return out;
}

/** Dla KWG: skrzynie eksportujemy jako (J-L)/(K-M). Dla KW: bez zmian J/K. */
function cratesForExport_(kw, rowNo) {
  const sheetName = kw.getName();
  const jRaw = String(kw.getRange("J" + rowNo).getDisplayValue() || "").trim();
  const kRaw = String(kw.getRange("K" + rowNo).getDisplayValue() || "").trim();
  if (jRaw === "" || kRaw === "") return "";

  if (sheetName !== "KWG") return jRaw + "/" + kRaw;

  const lRaw = String(kw.getRange("L" + rowNo).getDisplayValue() || "").trim();
  const mRaw = String(kw.getRange("M" + rowNo).getDisplayValue() || "").trim();

  const toNum = (s) => {
    const n = parseFloat(String(s || "").replace(",", "."));
    return isNaN(n) ? 0 : n;
  };
  const toIntStr = (n) => String(Math.max(0, Math.round(n)));

  const jOut = toNum(jRaw) - toNum(lRaw);
  const kOut = toNum(kRaw) - toNum(mRaw);
  return toIntStr(jOut) + "/" + toIntStr(kOut);
}

function purposeFromLotFull_(lotText) {
  const s = String(lotText || "").trim();
  if (!s) return "";

  const m = s.match(/([A-Za-z])\s*$/);
  const code = (m && m[1]) ? String(m[1]).toUpperCase() : "";

  if (code === "S") return "Sok";
  if (code === "O") return "Obieranie";
  if (code === "F") return "Świeże";
  if (code === "P") return "Przecier";
  return "";
}

/**
 * ===================== PS =====================
 * A NR DOSTAWY
 * B ODMIANA
 * C ILOŚĆ SKRZYŃ (TERAZ: J/K)
 * D WAGA NETTO (po zwrocie) + " kg"
 * E ZWROT (KG) + " kg"
 * F BRIX
 * G TWARDOŚĆ
 * H KALIBER ↓ 68mm %
 *
 * Uwaga: PS dopisuje tylko dla LOT głównego (parametry surowca z KW).
 */
function appendVarietyRowsToPS_(ss, kw) {
  const shPS = ss.getSheetByName(KW_EXPORT_CONFIG.PS_SHEET_NAME);
  if (!shPS) throw new Error('Brak arkusza: "' + KW_EXPORT_CONFIG.PS_SHEET_NAME + '"');

  const sheetName = kw.getName();
  const layout = typeof getLayout_ === "function" ? getLayout_(sheetName) : null;
  const lotFull = String(kw.getRange("F7:I7").getDisplayValue() || "").trim();
  if (!lotFull) return;
  const purposeShort = getPurposeShortFromLot_(kw, lotFull);
  const qs = layout
    ? (typeof getEffectiveQualityStarts_ === "function" ? getEffectiveQualityStarts_(sheetName, purposeShort) : layout.qualityStarts)
    : (sheetName === "KW" ? [24, 30, 36, 42] : [19, 25, 31, 37]);
  const tr = layout
    ? (typeof getEffectiveTwardoscRows_ === "function" ? getEffectiveTwardoscRows_(sheetName, purposeShort) : layout.twardoscRows)
    : (sheetName === "KW" ? [26, 32, 38, 44] : [21, 27, 33, 39]);
  const vf = layout ? layout.varietyRowFirst : (sheetName === "KW" ? 19 : 12);

  const nrDostawy = extractLotCodeOnly_(lotFull);

  const varCells = ["E" + vf, "E" + (vf + 1), "E" + (vf + 2), "E" + (vf + 3)];
  const brixCells = qs.map(s => "E" + s);
  const twardCells = tr.map(r => "E" + r);
  const kaliberCells = qs.map(s => "E" + (s + 3));
  const totalNettoCells = qs.map(s => "J" + s);
  const nettoAfterCells = qs.map(s => "K" + (s + 1));

  const rows = [];
  for (let i = 0; i < 4; i++) {
    const variety = String(kw.getRange(varCells[i]).getDisplayValue() || "").trim();
    if (!variety) continue;

    // skrzynie: KW=J/K, KWG=(J-L)/(K-M)
    const crates = cratesForExport_(kw, vf + i);

    // WAGA NETTO po zwrocie
    const afterRaw = String(kw.getRange(nettoAfterCells[i]).getDisplayValue() || "").trim();
    const afterNum = extractNumber_(afterRaw);
    const afterOut = afterNum ? (normalizeIntKg_(afterNum) + " kg") : "ND";

    // ZWROT w KG = (przed - po)
    const beforeRaw = String(kw.getRange(totalNettoCells[i]).getDisplayValue() || "").trim();
    const beforeNum = extractNumber_(beforeRaw);

    let zwrotOut = "ND";
    if (beforeNum && afterNum) {
      const diff = toNumber_(beforeNum) - toNumber_(afterNum);
      if (isFinite(diff)) zwrotOut = (normalizeIntKg_(diff) + " kg");
    }

    const brix = valOrND_(String(kw.getRange(brixCells[i]).getDisplayValue() || "").trim());
    const tward = valOrND_(String(kw.getRange(twardCells[i]).getDisplayValue() || "").trim());
    const kaliber = valOrND_(String(kw.getRange(kaliberCells[i]).getDisplayValue() || "").trim());

    rows.push([
      nrDostawy,
      variety,
      crates,
      afterOut,
      zwrotOut,
      brix,
      tward,
      kaliber
    ]);
  }

  if (!rows.length) return;

  const lastRow = Math.max(shPS.getLastRow(), 2);
  const targetRow = (lastRow < 3) ? 3 : (lastRow + 1);

  shPS.getRange(targetRow, 1, rows.length, 8).setValues(rows);
  SpreadsheetApp.flush();
}

function valOrND_(v) {
  const s = String(v == null ? "" : v).trim();
  return s === "" ? "ND" : s;
}

function extractLotCodeOnly_(lotRaw) {
  const s = String(lotRaw || "").trim();
  if (!s) return "";
  const idx = s.indexOf(" - ");
  if (idx !== -1) return s.substring(0, idx).trim();

  const idx2 = s.indexOf("-");
  if (idx2 !== -1) return s.substring(0, idx2).trim();

  return s.trim();
}

/** Wyciąga pierwszą liczbę z tekstu */
function extractNumber_(txt) {
  const s = String(txt || "");
  const m = s.match(/(\d+(?:[.,]\d+)?)/);
  if (!m) return "";
  return String(m[1]).replace(",", ".");
}

function toNumber_(s) {
  const n = parseFloat(String(s).replace(",", "."));
  return isNaN(n) ? NaN : n;
}

function normalizeIntKg_(num) {
  const n = Math.round(Number(num));
  return isFinite(n) ? String(n) : "";
}

/******************* CZYSZCZENIE PO EKSPORCIE (NIE USUWA FORMUŁ) *******************/
function clearAfterExport_KeepFormulas_(sh) {
  const name = sh.getName();
  const layout = typeof getLayout_ === "function" ? getLayout_(name) : null;
  const vf = layout ? layout.varietyRowFirst : (name === "KW" ? 19 : 12);
  const vl = layout ? layout.varietyRowLast : (name === "KW" ? 22 : 15);
  const lotText = String(sh.getRange("F7:I7").getDisplayValue() || "").trim();
  const purposeShort = typeof getPurposeShortFromLot_ === "function" ? getPurposeShortFromLot_(sh, lotText) : "";
  const qs = layout
    ? (typeof getEffectiveQualityStarts_ === "function" ? getEffectiveQualityStarts_(name, purposeShort) : layout.qualityStarts)
    : (name === "KW" ? [24, 30, 36, 42] : [19, 25, 31, 37]);

  clearValuesPreserveFormulas_(sh.getRange("F5:I7"));

  if (name === "KW") {
    clearValuesPreserveFormulas_(sh.getRange("E11:K12"));
    clearValuesPreserveFormulas_(sh.getRange("E15"));
    clearValuesPreserveFormulas_(sh.getRange("E16"));
    clearValuesPreserveFormulas_(sh.getRange("E17:K17"));
    clearValuesPreserveFormulas_(sh.getRange("E18:K18"));
  } else {
    clearValuesPreserveFormulas_(sh.getRange("E11:K11"));
  }
  clearValuesPreserveFormulas_(sh.getRange("E12:E13"));

  clearValuesPreserveFormulas_(sh.getRange("E" + vf + ":G" + vl));
  clearValuesPreserveFormulas_(sh.getRange("J" + vf + ":K" + vl));
  // KWG: dodatkowe skrzynie MBF w L/M też czyścimy po eksporcie
  clearValuesPreserveFormulas_(sh.getRange("L" + vf + ":M" + vl));

  qs.forEach(s => clearValuesPreserveFormulas_(sh.getRange("E" + s + ":E" + (s + 3))));
  clearValuesPreserveFormulas_(sh.getRange("M1"));
}

function clearValuesPreserveFormulas_(range) {
  const formulas = range.getFormulas();
  const values = range.getValues();

  let changed = false;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const f = formulas[r][c];
      const hasFormula = f && String(f).trim() !== "";
      if (!hasFormula) {
        if (values[r][c] !== "") {
          values[r][c] = "";
          changed = true;
        }
      }
    }
  }
  if (changed) range.setValues(values);
}

/******************* PLS_VIEW: znajdź LOT i zaznacz *******************/
function selectLotInPLSView_(ss, lotText) {
  const shView = ss.getSheetByName(KW_EXPORT_CONFIG.PLS_VIEW_SHEET_NAME);
  if (!shView) throw new Error('Brak arkusza: "' + KW_EXPORT_CONFIG.PLS_VIEW_SHEET_NAME + '"');

  const lastRow = Math.max(shView.getLastRow(), 3);
  const range = shView.getRange(3, 1, Math.max(1, lastRow - 2), 1);
  const vals = range.getDisplayValues();

  const needle = String(lotText || "").trim();
  if (!needle) {
    ss.setActiveSheet(shView);
    return;
  }

  let foundRow = -1;
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || "").trim();
    if (v === needle) { foundRow = 3 + i; break; }
  }

  ss.setActiveSheet(shView);

  if (foundRow !== -1) {
    const cell = shView.getRange(foundRow, 1);
    shView.setActiveRange(cell);
  } else {
    shView.setActiveRange(shView.getRange("A3"));
  }

  SpreadsheetApp.flush();
}

/******************* RICH LINK *******************/
function setRichLink_(cell, text, url) {
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setLinkUrl(0, text.length, String(url))
    .build();
  cell.setRichTextValue(rich);
  cell.setFontWeight("bold");
}

/******************* SYNC ROZMIARU SIATKI *******************/
function syncSheetGridSize_(dstSheet, srcMaxRows, srcMaxCols) {
  const dstRows = dstSheet.getMaxRows();
  if (dstRows < srcMaxRows) dstSheet.insertRowsAfter(dstRows, srcMaxRows - dstRows);
  else if (dstRows > srcMaxRows) dstSheet.deleteRows(srcMaxRows + 1, dstRows - srcMaxRows);

  const dstCols = dstSheet.getMaxColumns();
  if (dstCols < srcMaxCols) dstSheet.insertColumnsAfter(dstCols, srcMaxCols - dstCols);
  else if (dstCols > srcMaxCols) dstSheet.deleteColumns(srcMaxCols + 1, dstCols - srcMaxCols);

  SpreadsheetApp.flush();
}

/******************* PARSING *******************/
function parseSupplier_(raw) {
  const s = String(raw || "").trim();
  if (!s) return { no: "", name: "" };

  const m = s.match(new RegExp("^\\s*(\\d+)\\s*-\\s*(.+)\\s*$"));
  if (m) return { no: m[1], name: m[2] };

  return { no: "", name: s };
}

function extractDeliveryNo_3to6_(raw, padLen) {
  const s = String(raw || "");
  let part = "";
  if (s.length >= 6) part = s.substring(2, 6);
  else if (s.length > 2) part = s.substring(2);
  else part = s;

  const digits = String(part).replace(new RegExp("\\D+", "g"), "");
  if (!digits) return "0".repeat(padLen || 4);

  const n = digits.slice(-(padLen || 4));
  return n.padStart(padLen || 4, "0");
}

function sanitizeFileName_(name) {
  return String(name || "")
    .replace(new RegExp("\\/", "g"), "／")
    .replace(new RegExp('[\\\\\\?\\*\\[\\]\\|\\"]', "g"), " ")
    .replace(new RegExp("\\s+", "g"), " ")
    .trim()
    .slice(0, 180);
}

/******************* FOLDER MIESIĄCA *******************/
function getOrCreateMonthFolder_(parentFolder, dateObj) {
  const monthsPL = [
    "Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
    "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"
  ];

  const tz = Session.getScriptTimeZone() || "Europe/Warsaw";
  const year = parseInt(Utilities.formatDate(dateObj, tz, "yyyy"), 10);
  const monthIndex = parseInt(Utilities.formatDate(dateObj, tz, "M"), 10) - 1;

  const folderName = `${monthsPL[monthIndex]} ${year}`;

  const it = parentFolder.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();

  return parentFolder.createFolder(folderName);
}

/******************* HIDE BLOKÓW PO EKSPORCIE *******************/
function HIDE_ALL_BLOCKS_AFTER_EXPORT_(sh) {
  if (!sh) return false;
  const layout = typeof getLayout_ === "function" ? getLayout_(sh.getName()) : null;
  if (!layout || !layout.blocksHideFromTo) return false;

  const [startRow, endRow] = layout.blocksHideFromTo;
  const howMany = endRow - startRow + 1;

  try { sh.showRows(startRow, howMany); } catch (e) { if (e && (e.message || e.toString)) Logger.log("HIDE_ALL_BLOCKS_AFTER_EXPORT_: " + (e.message || e.toString())); }
  sh.hideRows(startRow, howMany);

  SpreadsheetApp.flush();
  return true;
}