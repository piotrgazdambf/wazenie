/******************* ETYKIETA (PLS_VIEW – QR LOT) *******************
 * ZMIANY:
 * - Działa TYLKO w arkuszu PLS_VIEW.
 * - Zaznaczasz LOT w kolumnie A (aktywna komórka).
 * - Dane z tego samego wiersza:
 *    * Odmiana/owoc: kolumna F
 *    * LOT: kolumna A
 *    * Data: kolumna B
 *    * Numerek dostawcy + dostawca: kolumna H
 *    * Przeznaczenie: kolumna E
 * - Na etykietę trafia pełny LOT z kolumny A (bez skracania), np. C/0167/024/26-O, C/0167/024/26-O2.
 * - QR koduje LOT z A2 (cały wpis z PLS).
 *
 * UWAGA:
 * - Ten plik NIE ma onOpen(). Menu dodaje router z innego skryptu:
 *   if (typeof QR_onOpenMenu_ === "function") QR_onOpenMenu_();
 ********************************************************************/

const MENU_NAME = "Stwórz QR";
const MENU_ITEM_CREATE = "Utwórz etykietę LOT (QR)";
const MENU_ITEM_SAVE_TEMPLATE = "Zapisz formatkę (raz)";

const SOURCE_SHEET_NAME = "PLS_VIEW";
const PRINT_SHEET_NAME = "ETYKIETASUROWCOWA";

const TEMPLATE_SHEET_NAME = "__TEMPLATE_ETYKIETASUROWCOWA";
const TEMPLATE_RANGE_A1 = "A1:C4";

const QR_DISPLAY_SIZE_PX = 230;
const QR_FETCH_SIZE_PX = 2000;
const QR_MARGIN = 1;

const TEMPLATE_DIM_PROP_KEY = "RAW_LABEL_TEMPLATE_DIM_V1";

/**
 * Menu w Arkuszu (wywoływane przez router onOpen w innym pliku)
 * Z KW/KWG: otwiera dialog wyboru odmiany i generuje etykietę. Z PLS_VIEW: jak dotąd (z zaznaczonego wiersza).
 */
function QR_onOpenMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem(MENU_ITEM_CREATE, "CREATE_LABEL_OR_FROM_KW")
    .addItem("Etykieta QR z Karty Ważenia (wybierz odmianę)", "KW_QR_OPEN_DIALOG_")
    .addSeparator()
    .addItem(MENU_ITEM_SAVE_TEMPLATE, "SAVE_TEMPLATE_ONCE")
    .addToUi();
}

/**
 * Router: w KW/KWG → dialog z listą odmian (E19:E22 / E14:E17); w PLS_VIEW → CREATE_LABEL.
 */
function CREATE_LABEL_OR_FROM_KW() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (!sh) {
    ui.alert("Brak aktywnego arkusza.");
    return;
  }
  const sheetName = sh.getName();
  const sheetUpper = String(sheetName || "").trim().toUpperCase();
  if (sheetUpper === "KW" || sheetUpper === "KWG") {
    KW_QR_OPEN_DIALOG_();
    return;
  }
  if (sheetName === SOURCE_SHEET_NAME) {
    CREATE_LABEL();
    return;
  }
  ui.alert(
    "Info",
    `Etykietę QR tworzysz z arkusza ${SOURCE_SHEET_NAME} (zaznacz LOT w kolumnie A) albo z Karty Ważenia (KW/KWG).`,
    ui.ButtonSet.OK
  );
}

/**
 * QR z A2 (A2 = pełny LOT z PLS, np. C/0167/024/26-O2)
 */
function qrFromA2FormulaPL_(displayPx, fetchPx, margin) {
  const baseUrl = `https://quickchart.io/qr?size=${fetchPx}&margin=${margin}&text=`;
  return `=IMAGE("${baseUrl}"&ENCODEURL(A2);4;${displayPx};${displayPx})`;
}

/**
 * Normalizacja LOT na etykietę:
 * - usuwa wszystkie "-" (myślniki)
 * - usuwa końcowe O/P/S/F (jeśli występuje)
 */
function normalizeLotForLabel_(lot) {
  let s = String(lot || "").trim();
  s = s.replace(/-/g, "");        // usuń myślniki
  s = s.replace(/[OPSF]$/i, "");  // usuń końcowe O/P/S/F
  return s;
}

function SAVE_TEMPLATE_ONCE() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shPrint = ss.getSheetByName(PRINT_SHEET_NAME);
  if (!shPrint) {
    ui.alert(`Brak arkusza "${PRINT_SHEET_NAME}". Utwórz go i ustaw formatkę.`);
    return;
  }

  let shTpl = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!shTpl) shTpl = ss.insertSheet(TEMPLATE_SHEET_NAME);
  shTpl.hideSheet();

  shTpl.getRange(TEMPLATE_RANGE_A1).breakApart();
  shTpl.getRange(TEMPLATE_RANGE_A1).clear();

  const src = shPrint.getRange(TEMPLATE_RANGE_A1);
  const dst = shTpl.getRange(TEMPLATE_RANGE_A1);

  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  dst.clearContent();

  const dim = {
    colWidths: [1, 2, 3].map(c => shPrint.getColumnWidth(c)),
    rowHeights: [1, 2, 3, 4].map(r => shPrint.getRowHeight(r)),
  };
  PropertiesService.getDocumentProperties().setProperty(
    TEMPLATE_DIM_PROP_KEY,
    JSON.stringify(dim)
  );

  ui.alert("Zapisano formatkę. Teraz generowanie będzie przywracało wygląd.");
}

function RESTORE_TEMPLATE_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPrint = ss.getSheetByName(PRINT_SHEET_NAME);
  const shTpl = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!shPrint || !shTpl) return false;

  const rawDim = PropertiesService.getDocumentProperties().getProperty(
    TEMPLATE_DIM_PROP_KEY
  );
  if (rawDim) {
    try {
      const dim = JSON.parse(rawDim);
      if (dim.colWidths && dim.colWidths.length === 3) {
        shPrint.setColumnWidth(1, dim.colWidths[0]);
        shPrint.setColumnWidth(2, dim.colWidths[1]);
        shPrint.setColumnWidth(3, dim.colWidths[2]);
      }
      if (dim.rowHeights && dim.rowHeights.length === 4) {
        shPrint.setRowHeight(1, dim.rowHeights[0]);
        shPrint.setRowHeight(2, dim.rowHeights[1]);
        shPrint.setRowHeight(3, dim.rowHeights[2]);
        shPrint.setRowHeight(4, dim.rowHeights[3]);
      }
    } catch (e) { if (e && (e.message || e.toString)) Logger.log("RESTORE_TEMPLATE_ dim: " + (e.message || e.toString())); }
  }

  const dst = shPrint.getRange(TEMPLATE_RANGE_A1);
  dst.breakApart();
  const src = shTpl.getRange(TEMPLATE_RANGE_A1);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  return true;
}

/**
 * CREATE_LABEL:
 * - tylko PLS_VIEW
 * - aktywna komórka musi być w kolumnie A (LOT)
 * - odmiana: F
 * - lot: A (po normalizacji: bez "-" i bez końcowego O/P/S/F)
 * - data: B
 * - dostawca + nr: H
 * - przeznaczenie: E
 */
function CREATE_LABEL() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sh = ss.getActiveSheet();
  if (!sh) return ui.alert("Brak aktywnego arkusza.");

  const sheetName = sh.getName();
  const sheetUpper = String(sheetName || "").trim().toUpperCase();
  // Gdy wywołano z KW/KWG (np. stare menu) – przekieruj do dialogu z listą odmian
  if (sheetUpper === "KW" || sheetUpper === "KWG") {
    KW_QR_OPEN_DIALOG_();
    return;
  }

  if (sheetName !== SOURCE_SHEET_NAME) {
    ui.alert(
      "Info",
      "Aktywny arkusz: \"" + (sheetName || "") + "\".\n\nEtykietę QR tworzysz z arkusza " + SOURCE_SHEET_NAME + " (zaznacz LOT w kolumnie A) albo z Karty Ważenia (KW/KWG).",
      ui.ButtonSet.OK
    );
    return;
  }

  const cell = sh.getActiveCell();
  if (!cell) return ui.alert(`Zaznacz LOT w kolumnie A w arkuszu ${SOURCE_SHEET_NAME}.`);

  const row = cell.getRow();
  const col = cell.getColumn();

  // Musi być kolumna A
  if (col !== 1) {
    ui.alert("Info", "Zaznacz komórkę z LOT-em w kolumnie A (to właśnie wybierasz).", ui.ButtonSet.OK);
    return;
  }

  const lotOriginal = String(cell.getDisplayValue() || "").trim();
  if (!lotOriginal) {
    ui.alert("Brak LOT", "Zaznaczona komórka w kolumnie A jest pusta (brak LOT).", ui.ButtonSet.OK);
    return;
  }

  const varietyText = String(sh.getRange(row, 6).getDisplayValue() || "").trim(); // F
  if (!varietyText) {
    ui.alert("Brak odmiany", "W tym wierszu brakuje odmiany/owocu w kolumnie F.", ui.ButtonSet.OK);
    return;
  }

  const dateText = getDateTextFromCell_(sh.getRange(row, 2)); // B
  const supplierLine = String(sh.getRange(row, 8).getDisplayValue() || "").trim(); // H
  if (!supplierLine) {
    ui.alert("Brak dostawcy", "W tym wierszu brakuje numerku dostawcy/dostawcy w kolumnie H.", ui.ButtonSet.OK);
    return;
  }

  const purposeRaw = String(sh.getRange(row, 5).getDisplayValue() || "").trim(); // E
  const purposeLine = "Przeznaczenie: " + (purposeRaw || "");

  const shPrint = getOrCreateSheet_(ss, PRINT_SHEET_NAME);

  const ok = RESTORE_TEMPLATE_();
  if (!ok) {
    ui.alert(`Najpierw: ${MENU_NAME} → ${MENU_ITEM_SAVE_TEMPLATE}`);
    return;
  }

  // GÓRA: odmiana/owoc
  shPrint.getRange("B1").setValue(varietyText);

  // A2: pełny LOT z PLS (np. C/0167/024/26-O2) – każdy kod indywidualny, QR koduje to samo
  shPrint.getRange("A2").setValue(lotOriginal);

  // QR: koduje LOT z A2 (pełny kod)
  shPrint.getRange("B2").setFormula(qrFromA2FormulaPL_(QR_DISPLAY_SIZE_PX, QR_FETCH_SIZE_PX, QR_MARGIN));

  // PRAWY PION: data
  shPrint.getRange("C2").setValue(dateText);

  // DÓŁ: dostawca/nr + przeznaczenie
  shPrint.getRange("A3:C3").setValue(supplierLine);
  shPrint.getRange("A4:C4").setValue(purposeLine);

  SpreadsheetApp.flush();
  ss.setActiveSheet(shPrint);
  shPrint.setActiveSelection("B2");
}

/**
 * Otwiera dialog wyboru odmiany z Karty Ważenia (KW/KWG).
 * Odmiany: KW E19:E22, KWG E14:E17. Kod etykiety: pierwsza odmiana = LOT, druga = LOT2, itd.
 */
function KW_QR_OPEN_DIALOG_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (!sh) return;
  const sheetName = sh.getName();
  const sheetUpper = String(sheetName || "").trim().toUpperCase();
  if (sheetUpper !== "KW" && sheetUpper !== "KWG") {
    SpreadsheetApp.getUi().alert("Info", "Otwórz arkusz KW lub KWG i wybierz ponownie: Stwórz QR → Etykieta QR z Karty Ważenia.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const layout = typeof getLayout_ === "function" ? getLayout_(sheetUpper) : null;
  const vf = layout ? layout.varietyRowFirst : (sheetUpper === "KW" ? 19 : 14);
  const vl = layout ? layout.varietyRowLast : (sheetUpper === "KW" ? 22 : 17);

  const varieties = [];
  for (let r = vf; r <= vl; r++) {
    const v = String(sh.getRange("E" + r).getDisplayValue() || "").trim();
    if (v) varieties.push(v);
  }
  if (!varieties.length) {
    ui.alert("Brak odmian", "Uzupełnij odmiany w komórkach E" + vf + "–E" + vl + " w Karcie Ważenia.", ui.ButtonSet.OK);
    return;
  }

  const baseLot = String(sh.getRange("F7:I7").getDisplayValue() || "").trim();
  if (!baseLot) {
    ui.alert("Brak LOT", "Uzupełnij LOT w F7:I7 w Karcie Ważenia.", ui.ButtonSet.OK);
    return;
  }
  const dateText = getDateTextFromCell_(sh.getRange("F5"));
  const supplierLine = String(sh.getRange("F6:I6").getDisplayValue() || "").trim();
  const lotFull = baseLot;
  const purposeName = typeof purposeFromLotFull_ === "function" ? purposeFromLotFull_(lotFull) : "";
  const purposeLine = "Przeznaczenie: " + (purposeName || "");

  const html = KW_QR_HTML_(varieties, baseLot, dateText, supplierLine, purposeLine);
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(380), "Etykieta QR z Karty Ważenia");
}

function KW_QR_HTML_(varieties, baseLot, dateText, supplierLine, purposeLine) {
  const escapeAttr = (s) => String(s || "").replace(/\\/g, "\\\\").replace(/"/g, "&quot;").replace(/'/g, "&#39;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/&/g, "&amp;");
  const options = varieties.map((v, i) => '<option value="' + escapeAttr(v) + '" data-idx="' + i + '">' + escapeAttr(v) + '</option>').join("");
  const baseLotEsc = escapeAttr(baseLot);
  const dateEsc = escapeAttr(dateText);
  const supplierEsc = escapeAttr(supplierLine);
  const purposeEsc = escapeAttr(purposeLine);

  return '<!DOCTYPE html><html><head><meta charset="utf-8"/>' +
    '<style>' +
    '*{box-sizing:border-box} body{margin:0;font-family:"Segoe UI",system-ui,sans-serif;background:#f0f4f8;padding:20px;color:#1a202c}' +
    '.box{background:#fff;border-radius:10px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:20px;max-width:380px}' +
    'h3{margin:0 0 14px 0;font-size:14px;color:#4a5568}' +
    'select{width:100%;padding:10px 12px;border:1px solid #e2e8f0;border-radius:8px;font-size:14px;margin-bottom:12px}' +
    '.code{margin:12px 0;padding:10px;background:#edf2f7;border-radius:6px;font-family:monospace;font-size:13px;word-break:break-all}' +
    '.btn{width:100%;padding:12px;background:linear-gradient(135deg,#1e3a5f,#2d5a87);color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;margin-top:16px}' +
    '.btn:hover{opacity:.95}' +
    '</style></head><body><div class="box">' +
    '<h3>Wybierz odmianę</h3>' +
    '<select id="sel">' + options + '</select>' +
    '<div class="code" id="codePreview">Kod etykiety: ' + escapeAttr(baseLot) + '</div>' +
    '<input type="hidden" id="baseLot" value="' + baseLotEsc + '"/>' +
    '<input type="hidden" id="dateText" value="' + dateEsc + '"/>' +
    '<input type="hidden" id="supplierLine" value="' + supplierEsc + '"/>' +
    '<input type="hidden" id="purposeLine" value="' + purposeEsc + '"/>' +
    '<button type="button" class="btn" id="btnSend">Prześlij</button>' +
    '</div><script>' +
    'var sel=document.getElementById("sel");var code=document.getElementById("codePreview");var base=document.getElementById("baseLot").value;' +
    'function lotForIdx(i){return base+(i===0?"":String(i+1));}' +
    'function updatePreview(){var i=parseInt(sel.options[sel.selectedIndex].getAttribute("data-idx"),10);code.textContent="Kod etykiety: "+lotForIdx(i);}' +
    'sel.onchange=updatePreview;updatePreview();' +
    'document.getElementById("btnSend").onclick=function(){' +
    'var variety=sel.options[sel.selectedIndex].value;var i=parseInt(sel.options[sel.selectedIndex].getAttribute("data-idx"),10);var lot=lotForIdx(i);' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();}).withFailureHandler(function(e){alert("Błąd: "+(e&&e.message?e.message:String(e)));})' +
    '.APPLY_LABEL_FROM_KW_(variety,lot,document.getElementById("dateText").value,document.getElementById("supplierLine").value,document.getElementById("purposeLine").value);' +
    '};</script></body></html>';
}

/**
 * Zapisuje etykietę na ETYKIETASUROWCOWA (jak CREATE_LABEL) i zeruje licznik następnej dostawy (WSG C4 = 0001).
 * Gdy brak zapisanej formatki – i tak wpisuje dane i przełącza na arkusz etykiet (możesz zapisać formatkę później).
 */
function APPLY_LABEL_FROM_KW_(varietyText, lotForLabel, dateText, supplierLine, purposeLine) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPrint = getOrCreateSheet_(ss, PRINT_SHEET_NAME);

  RESTORE_TEMPLATE_(); // jeśli formatka zapisana – przywróć; jeśli nie – nic, i tak wpisujemy dane

  shPrint.getRange("B1").setValue(String(varietyText || "").trim());
  shPrint.getRange("A2").setValue(String(lotForLabel || "").trim());
  shPrint.getRange("B2").setFormula(qrFromA2FormulaPL_(QR_DISPLAY_SIZE_PX, QR_FETCH_SIZE_PX, QR_MARGIN));
  shPrint.getRange("C2").setValue(String(dateText || "").trim());
  shPrint.getRange("A3:C3").setValue(String(supplierLine || "").trim());
  shPrint.getRange("A4:C4").setValue(String(purposeLine || "").trim());

  const shWSG = ss.getSheetByName("WSG");
  if (shWSG) {
    const c4 = shWSG.getRange("C4");
    c4.setValue(1);
    c4.setNumberFormat("0000");
  }

  SpreadsheetApp.flush();
  ss.setActiveSheet(shPrint);
  shPrint.setActiveSelection("B2");
  ss.toast("Etykieta wpisana. Przełączono na arkusz „" + PRINT_SHEET_NAME + "\".", "QR z KW", 4);
}

/**
 * Data z komórki (jeśli Date -> dd.MM.yyyy, inaczej display)
 */
function getDateTextFromCell_(range) {
  const v = range.getValue();
  const dv = String(range.getDisplayValue() || "").trim();
  const tz = Session.getScriptTimeZone() || "Europe/Warsaw";

  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, tz, "dd.MM.yyyy");
  }
  return dv;
}

/**
 * Helpers
 */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}