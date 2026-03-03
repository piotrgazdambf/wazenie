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
 */
function QR_onOpenMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem(MENU_ITEM_CREATE, "CREATE_LABEL")
    .addSeparator()
    .addItem(MENU_ITEM_SAVE_TEMPLATE, "SAVE_TEMPLATE_ONCE")
    .addToUi();
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
  if (sheetName !== SOURCE_SHEET_NAME) {
    ui.alert(
      "Info",
      `Etykietę QR tworzysz tylko z arkusza ${SOURCE_SHEET_NAME}. Przejdź do ${SOURCE_SHEET_NAME} i zaznacz LOT w kolumnie A.`,
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