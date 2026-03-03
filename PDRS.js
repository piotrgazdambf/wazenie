/******************* PDRS (BEZ ZMIAN – 1:1 JAK PRZYSŁAŁEŚ) *******************
 * Ten moduł zostawiam bez zmian.
 *******************************************************************************/

const CFG_PDRS = {
  SOURCE_SHEET_MBS: "MBS",
  SOURCE_SHEET_PLS_VIEW: "PLS_VIEW",
  TARGET_SPREADSHEET_ID: "181Hg9h_CuXKV5z0oOzkoErTwUSY7icEhxZIgMh9OYPE",
  TARGET_SHEET_NAME: "Raporty Surowiec",

  MBS_STATUS_COL: 12,
  SENT_MARK: "✅",

  TARGET_LOT_COL: 8,
  /** Kolumny MBS D:K (4–11) kopiowane do targetu od kolumny 9 (1 wiersz, 8 kolumn). */
  MBS_COLS_FROM: 4,
  MBS_COLS_COUNT: 8,
  TARGET_MBS_START_COL: 9,

  MBS_STRIKE_FROM_COL: 1,
  MBS_STRIKE_TO_COL: 12
};

function PDRS() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shMBS = ss.getSheetByName(CFG_PDRS.SOURCE_SHEET_MBS);
  const shPLS = ss.getSheetByName(CFG_PDRS.SOURCE_SHEET_PLS_VIEW);
  if (!shMBS) throw new Error(`Brak arkusza: ${CFG_PDRS.SOURCE_SHEET_MBS}`);
  if (!shPLS) throw new Error(`Brak arkusza: ${CFG_PDRS.SOURCE_SHEET_PLS_VIEW}`);

  const range = ss.getActiveRange();
  if (!range) throw new Error("Zaznacz komórkę z LOTem.");

  const lot = String(range.getValue() ?? "").trim();
  if (!lot) throw new Error("Zaznaczona komórka jest pusta – zaznacz LOT.");

  const mbsRow = range.getRow();
  const statusCell = shMBS.getRange(mbsRow, CFG_PDRS.MBS_STATUS_COL);
  const statusVal = String(statusCell.getValue() ?? "").trim();

  if (statusVal === CFG_PDRS.SENT_MARK) {
    ui.alert("Info", `Ten LOT (${lot}) został już przesłany do Raporty Surowiec.`, ui.ButtonSet.OK);
    return;
  }

  const targetSS = SpreadsheetApp.openById(CFG_PDRS.TARGET_SPREADSHEET_ID);
  const shTarget = targetSS.getSheetByName(CFG_PDRS.TARGET_SHEET_NAME);
  if (!shTarget) throw new Error(`Brak arkusza docelowego: ${CFG_PDRS.TARGET_SHEET_NAME}`);

  const lastRowTarget = shTarget.getLastRow();
  if (lastRowTarget >= 2) {
    const foundInTarget = shTarget
      .getRange(2, CFG_PDRS.TARGET_LOT_COL, lastRowTarget - 1, 1)
      .createTextFinder(lot)
      .matchEntireCell(true)
      .findNext();

    if (foundInTarget) {
      const width = CFG_PDRS.MBS_STRIKE_TO_COL - CFG_PDRS.MBS_STRIKE_FROM_COL + 1;
      shMBS.getRange(mbsRow, CFG_PDRS.MBS_STRIKE_FROM_COL, 1, width)
        .setFontLine("line-through")
        .setFontColor("#ff0000");

      ui.alert(
        "Info",
        `Ten LOT (${lot}) już znajduje się w Raporty Surowiec.\nTo jest zduplikowany wpis w MBS – został przekreślony i NIE został wysłany.`,
        ui.ButtonSet.OK
      );
      return;
    }
  }

  const response = ui.alert(
    "Potwierdzenie",
    "Czy na pewno chcesz zagliźnić?",
    ui.ButtonSet.OK_CANCEL
  );
  if (response !== ui.Button.OK) return;

  const mbsDToK = shMBS.getRange(mbsRow, CFG_PDRS.MBS_COLS_FROM, 1, CFG_PDRS.MBS_COLS_COUNT).getValues()[0];

  const lastRowPLS = shPLS.getLastRow();
  if (lastRowPLS < 2) throw new Error("PLS_VIEW jest pusty.");

  const found = shPLS
    .getRange(1, 1, lastRowPLS, 1)
    .createTextFinder(lot)
    .matchEntireCell(true)
    .findNext();

  if (!found) throw new Error(`Nie znaleziono LOTu "${lot}" w PLS_VIEW.`);

  const plsRow = found.getRow();

  const dataDostawy = shPLS.getRange(plsRow, 2).getValue();
  const nrDostawy   = shPLS.getRange(plsRow, 3).getValue();
  const dostawca    = shPLS.getRange(plsRow, 4).getValue();
  const przezn      = shPLS.getRange(plsRow, 5).getValue();

  const targetRow = Math.max(shTarget.getLastRow() + 1, 2);

  shTarget.getRange(targetRow, 2).setValue(dataDostawy);
  shTarget.getRange(targetRow, 3).setValue(nrDostawy);
  shTarget.getRange(targetRow, 6).setValue(przezn);
  shTarget.getRange(targetRow, 7).setValue(dostawca);
  shTarget.getRange(targetRow, 8).setValue(lot);
  shTarget.getRange(targetRow, CFG_PDRS.TARGET_MBS_START_COL, 1, CFG_PDRS.MBS_COLS_COUNT).setValues([mbsDToK]);

  statusCell.setValue(CFG_PDRS.SENT_MARK);
}
