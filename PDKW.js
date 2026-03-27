/******************* PDKW (PEŁNA WERSJA + FIX KWG: tabelka 19-22, owoc E14:G14, "ZWROT w %") *******************
 * WSG:
 * - B4 data, C4 nr dostawy, D4 dostawca, E4 przeznaczenie, F4 owoc
 * - LOT liczony FORMUŁĄ w H3 (skrypt tylko odczytuje)
 * - RYLEX: wymagany dodatkowy kod w J3 (kolumna J odkrywana tylko dla RYLEX)
 *
 * KW:
 * - F5 data, F6:I6 dostawca, F7:I7 LOT
 * - owoc -> E19:G19
 * - tabelka jakości -> B24:E27
 *
 * KWG (SHIFT +2):
 * - F5 data, F6:I6 dostawca, F7:I7 LOT
 * - owoc -> E14:G14
 * - tabelka jakości -> B19:E22
 *************************************************************************************************************************/

function PDKW() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ===== LOCK =====
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    ui.alert("Skrypt już się wykonuje. Spróbuj ponownie za chwilę.");
    return;
  }

  try {
    const shWSG = ss.getSheetByName("WSG");
    const shKW  = ss.getSheetByName("KW");
    const shKWG = ss.getSheetByName("KWG");
    const shSUP = ss.getSheetByName("SUPPORTSUROWCE");
    const shPLS = ss.getSheetByName("PLS");

    if (!shWSG) throw new Error("Brak arkusza: WSG");
    if (!shKW)  throw new Error("Brak arkusza: KW");
    if (!shSUP) throw new Error("Brak arkusza: SUPPORTSUROWCE");
    if (!shPLS) throw new Error("Brak arkusza: PLS");

    const SUPPLIER_GROJECKA = "GRÓJECKA MBF";
    const SUPPLIER_MBF = "MBF";
    const SUPPLIER_RYLEX = "RYLEX";
    /** Owoc w F4: tylko te wartości prowadzą do arkusza KW; wszystko inne (marchewka, mango, itd.) → KWG */
    const FRUITS_TO_KW = ["jabłko", "gruszka", "jablko"];

    const PREVIEW_LOT_CELL = "H3";

    const RYLEX_LABEL_CELL = "J2";
    const RYLEX_INPUT_CELL = "J3";
    const RYLEX_LABEL_TEXT = "Dodatkowy numer (RYLEX)";
    const RYLEX_COL = 10; // J

    // ✅ NOWE: gdzie wpisujemy dodatkowy numer RYLEX w Karcie Ważenia
    const RYLEX_DEST_CELL = "J7"; // <— tutaj ma wpadać
    const RYLEX_LABEL_RANGE = "J6:K6";
    const RYLEX_BORDER_RANGE = "J6:K7";

    const tz = Session.getScriptTimeZone() || "Europe/Warsaw";

    function parseSheetDate(displayStr, rawValue) {
      if (rawValue instanceof Date && !isNaN(rawValue.getTime())) return rawValue;

      if (typeof rawValue === "number" && !isNaN(rawValue)) {
        const ms = Math.round((rawValue - 25569) * 86400 * 1000);
        const d = new Date(ms);
        return isNaN(d.getTime()) ? null : d;
      }

      if (displayStr === "" || displayStr == null) return null;

      let s = String(displayStr)
        .trim()
        .replace(/[\u200B-\u200D\uFEFF]/g, "")
        .replace(/[‐-‒–—―]/g, "-");

      let m = s.match(/^(\d{4})\D(\d{1,2})\D(\d{1,2})$/);
      if (m) {
        const yyyy = parseInt(m[1], 10);
        const mm = parseInt(m[2], 10);
        const dd = parseInt(m[3], 10);
        const d = new Date(yyyy, mm - 1, dd);
        if (d.getFullYear() === yyyy && d.getMonth() === (mm - 1) && d.getDate() === dd) return d;
        return null;
      }

      m = s.match(/^(\d{1,2})\D(\d{1,2})\D(\d{4})$/);
      if (m) {
        const dd = parseInt(m[1], 10);
        const mm = parseInt(m[2], 10);
        const yyyy = parseInt(m[3], 10);
        const d = new Date(yyyy, mm - 1, dd);
        if (d.getFullYear() === yyyy && d.getMonth() === (mm - 1) && d.getDate() === dd) return d;
        return null;
      }

      const p = Date.parse(s);
      if (!isNaN(p)) {
        const d = new Date(p);
        return isNaN(d.getTime()) ? null : d;
      }
      return null;
    }

    // ===================== WSG: źródła =====================
    const dateCell = shWSG.getRange("B4");
    const rawDate = dateCell.getValue();
    const dispDate = dateCell.getDisplayValue();

    const c4Cell = shWSG.getRange("C4");
    const delivNoRaw = String(c4Cell.getDisplayValue() || "").trim();

    const supplierName = String(shWSG.getRange("D4").getDisplayValue() || "").trim();
    const purposeName  = String(shWSG.getRange("E4").getDisplayValue() || "").trim();
    let fruitName      = String(shWSG.getRange("F4").getDisplayValue() || "").trim();
    const forceRylexByCheckbox = !!shWSG.getRange("K3").getValue();
    if (forceRylexByCheckbox || supplierName === SUPPLIER_RYLEX) {
      fruitName = "Jabłko";
      try { shWSG.getRange("F4").setValue("Jabłko"); } catch (e) { if (e && (e.message || e.toString)) Logger.log("PDKW force Jabłko F4: " + (e.message || e.toString())); }
    }

    if (!dispDate || !delivNoRaw || !supplierName || !purposeName) {
      ss.toast("PDKW: uzupełnij WSG!B4 (data), C4 (nr dostawy), D4 (dostawca), E4 (przeznaczenie).", "Brak danych", 6);
      return;
    }

    // ===== wybór KW / KWG =====
    // 1) Checkbox K3 (RYLEX) lub L3 (GRÓJECKA) zaznaczony -> zawsze KWG
    // 2) Dostawca GRÓJECKA MBF / MBF / RYLEX -> zawsze KWG
    // 3) Owoc (F4) inny niż Jabłko/Gruszka (np. marchewka, mango) -> KWG
    const forceKWG_RYLEX = !!shWSG.getRange("K3").getValue();
    const forceKWG_GROJECKA = !!shWSG.getRange("L3").getValue();
    const fruitNorm = String(fruitName || "").toLowerCase().replace(/\s+/g, " ").trim();
    const isFruitForKW = FRUITS_TO_KW.some(f => fruitNorm === f);
    const isKWG =
      forceKWG_RYLEX ||
      forceKWG_GROJECKA ||
      supplierName === SUPPLIER_GROJECKA ||
      supplierName === SUPPLIER_MBF ||
      supplierName === SUPPLIER_RYLEX ||
      !isFruitForKW;
    let shDestKW = isKWG ? shKWG : shKW;
    if (isKWG && !shKWG) {
      throw new Error('Brak arkusza: KWG (wymagany dla RYLEX/GRÓJECKA/MBF, przy zaznaczonym K3/L3 lub gdy w F4 wybrano owoc inny niż Jabłko/Gruszka).');
    }

    // ===== RYLEX – J3 opcjonalny (bez blokady) =====
    let rylexExtra = "";
    if (supplierName === SUPPLIER_RYLEX) {
      try { shWSG.showColumns(RYLEX_COL); } catch (e) { if (e && (e.message || e.toString)) Logger.log("PDKW showColumns: " + (e.message || e.toString())); }
      shWSG.getRange(RYLEX_LABEL_CELL).setValue(RYLEX_LABEL_TEXT);

      const extraCell = shWSG.getRange(RYLEX_INPUT_CELL);
      const extraVal = String(extraCell.getDisplayValue() || "").trim();
      rylexExtra = extraVal; // opcjonalnie: jeśli wpisane, przenosimy do KW/KWG
    } else {
      shWSG.getRange(RYLEX_LABEL_CELL).clearContent();
      shWSG.getRange(RYLEX_INPUT_CELL).clearContent();
      try { shWSG.hideColumns(RYLEX_COL); } catch (e) { if (e && (e.message || e.toString)) Logger.log("PDKW hideColumns: " + (e.message || e.toString())); }
    }

    // ===== data =====
    const dateVal = parseSheetDate(dispDate, rawDate);
    if (!dateVal) throw new Error(`Nie mogę zinterpretować daty z WSG!B4. Widzę: "${dispDate}"`);
    const prettyDate = Utilities.formatDate(dateVal, tz, "dd.MM.yyyy");

    // ===== kod dostawcy z SUPPORTSUROWCE (A:B) =====
    const supAB = shSUP.getRange(1, 1, shSUP.getLastRow(), 2).getValues();
    let supplierCode = "";
    for (let i = 0; i < supAB.length; i++) {
      const name = (supAB[i][0] ?? "").toString().trim();
      if (name === supplierName) {
        const rawCode = supAB[i][1];
        const num = parseInt(rawCode, 10);
        supplierCode = !isNaN(num)
          ? String(num).padStart(3, "0")
          : (rawCode ?? "").toString().trim().padStart(3, "0");
        break;
      }
    }
    if (!supplierCode) throw new Error(`Nie znaleziono dostawcy "${supplierName}" w SUPPORTSUROWCE!A:A`);

    const supplierBase = `${supplierCode} - ${supplierName}`.replace(/\s+/g, " ").trim();

    // ✅ NOWE: F6 zawsze czysty dostawca, a RYLEX idzie do J7
    const supplierForKW = supplierBase;

    // ===== skrót przeznaczenia z SUPPORTSUROWCE (D:E) =====
    // getRange(row, column, numRows, numColumns): wiersze 1..lastRow, kolumny D(4) i E(5) = 2 kolumny
    const supDE = shSUP.getRange(1, 4, shSUP.getLastRow(), 2).getValues();
    let purposeShort = "";
    for (let i = 0; i < supDE.length; i++) {
      const pName = (supDE[i][0] ?? "").toString().trim();
      if (pName === purposeName) {
        purposeShort = (supDE[i][1] ?? "").toString().trim();
        break;
      }
    }
    if (!purposeShort) throw new Error(`Nie znaleziono skrótu przeznaczenia dla "${purposeName}" w SUPPORTSUROWCE!D:E`);

    // ===== ŚWIEŻE tylko dla GRÓJECKA =====
    if (String(purposeShort).toUpperCase() === "F" && supplierName !== SUPPLIER_GROJECKA) {
      ui.alert("Blokada", 'Przeznaczenie "ŚWIEŻE (F)" jest dozwolone tylko dla dostawcy GRÓJECKA MBF.', ui.ButtonSet.OK);
      return;
    }

    // ===== LOT: odczyt z H3 =====
    SpreadsheetApp.flush();
    const lotMain = String(shWSG.getRange(PREVIEW_LOT_CELL).getDisplayValue() || "").trim();
    if (!lotMain) {
      throw new Error('Brak LOT w WSG!H3. Wstaw formułę do H3 (generator LOT) i upewnij się, że B4-E4 są wypełnione.');
    }

    // ===== Zapis do KW/KWG =====
    shDestKW.getRange("F5").setValue(prettyDate);
    shDestKW.getRange("F6").setValue(supplierForKW);
    shDestKW.getRange("F7").setValue(lotMain);

    // ✅ NOWE: RYLEX -> J6:K6 etykieta "Numer dla RYLEX", J7:K7 kod, obramowanie J6:K7
    if (supplierName === SUPPLIER_RYLEX && rylexExtra) {
      shDestKW.getRange(RYLEX_LABEL_RANGE).setValue("Numer dla RYLEX");
      shDestKW.getRange(RYLEX_DEST_CELL).setValue(String(rylexExtra).trim());
      shDestKW.getRange(RYLEX_BORDER_RANGE).setBorder(true, true, true, true, true, true);
    } else {
      shDestKW.getRange(RYLEX_BORDER_RANGE).clearContent();
      shDestKW.getRange(RYLEX_BORDER_RANGE).setBorder(false, false, false, false, false, false);
    }

    // ===== OWOC: inne miejsce KW vs KWG =====
    const fruitRangeA1 = isKWG ? "E14:G14" : "E19:G19"; // SHIFT +2
    const fruitTarget = shDestKW.getRange(fruitRangeA1);

    if (!fruitName) {
      fruitTarget.clearContent();
      shDestKW.getRange("M1").clearContent();
    } else if (fruitName.toLowerCase() === "jabłko" || fruitName.toLowerCase() === "jablko") {
      fruitTarget.setValue("Wpisz odmianę Jabłka");
      shDestKW.getRange("M1").setValue("Jabłko");
    } else {
      fruitTarget.setValue(fruitName);
      shDestKW.getRange("M1").setValue(fruitName);
    }

    // ===== tabelka jakości: inne miejsce KW vs KWG =====
    const tableStartRow = isKWG ? 19 : 24; // SHIFT +2
    KW_BUILD_QUALITY_TABLE_(shDestKW, purposeShort, tableStartRow);

    // PLS: wiersze dopisuje tylko "Prześlij kartę ważenia" (appendVarietyRowsToPLS_) z pełnymi danymi i znormalizowanym LOT – tu nic nie dopisujemy

    // ===== AUTO-NUMER: następny numer w C4 (tylko gdy K3/L3 NIE zaznaczone) =====
    const manualMode = !!shWSG.getRange("K3").getValue() || !!shWSG.getRange("L3").getValue();
    if (!manualMode) {
      const digitsOnly = delivNoRaw.replace(/[^\d]/g, "");
      const currentNum = parseInt(digitsOnly, 10);
      let nextNum = isNaN(currentNum) ? 1 : (currentNum + 1);
      if (nextNum > 9999) nextNum = 1;

      c4Cell.setValue(nextNum);
      c4Cell.setNumberFormat("0000");
      c4Cell.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(1, 9999)
          .setAllowInvalid(false)
          .build()
      );

      // ===== ochrona warning-only na C4 =====
      try {
        const protections = shWSG.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        protections.forEach(p => {
          const r = p.getRange();
          if (r && r.getRange && r.getA1Notation && r.getA1Notation() === "C4") {
            try { p.remove(); } catch (e) { if (e && (e.message || e.toString)) Logger.log("PDKW protection remove: " + (e.message || e.toString())); }
          }
        });
        const p = c4Cell.protect();
        p.setDescription("AUTO NR DOSTAWY (nie edytować ręcznie)");
        p.setWarningOnly(true);
      } catch (e) { if (e && (e.message || e.toString)) Logger.log("PDKW C4 protect: " + (e.message || e.toString())); }
    }

    // ===== Czyść po wysyłce: B4, D4, E4, F4 =====
    const inputRange = shWSG.getRange("B4:F4");
    const savedValidations = inputRange.getDataValidations();
    const savedNumberFormats = inputRange.getNumberFormats();

    shWSG.getRangeList(["B4", "D4", "E4", "F4"]).clearContent();

    inputRange.setDataValidations(savedValidations);
    inputRange.setNumberFormats(savedNumberFormats);

    shWSG.getRange("B4")
      .setNumberFormat("yyyy-mm-dd")
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireDate()
          .setAllowInvalid(false)
          .build()
      );

    // czyść okienko RYLEX po sukcesie
    shWSG.getRange(RYLEX_LABEL_CELL).clearContent();
    shWSG.getRange(RYLEX_INPUT_CELL).clearContent();
    try { shWSG.hideColumns(RYLEX_COL); } catch (e) { if (e && (e.message || e.toString)) Logger.log("PDKW hideColumns: " + (e.message || e.toString())); }

    // ===== PRZEJDŹ DO KW/KWG =====
    ss.setActiveSheet(shDestKW);
    shDestKW.setActiveSelection(shDestKW.getRange("A1"));
    SpreadsheetApp.flush();

    ss.toast(`PDKW: przeniesiono dane do ${shDestKW.getName()} ✅ | LOT: ${lotMain}`, "OK", 5);

  } catch (err) {
    SpreadsheetApp.getUi().alert("Błąd: " + (err && err.message ? err.message : String(err)));
  } finally {
    try { lock.releaseLock(); } catch (e) { if (e && (e.message || e.toString)) Logger.log("PDKW releaseLock: " + (e.message || e.toString())); }
  }
}

/**
 * WSG K3/L3 (checkboxy):
 * - jeśli zaznaczony K3 lub L3 -> C4 czyścimy i umożliwiamy wpis ręczny
 * - jeśli oba odznaczone -> przywracamy poprzednią wartość C4 (sprzed zaznaczenia)
 */
function PDKW_WSG_onEdit_(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!sh || sh.getName() !== "WSG") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const isSwitchCell = (row === 3 && (col === 11 || col === 12)); // K3/L3
  const isSupplierCell = (row === 4 && col === 4); // D4
  const isFruitCell = (row === 4 && col === 6); // F4
  if (!isSwitchCell && !isSupplierCell && !isFruitCell) return;

  const props = PropertiesService.getDocumentProperties();
  const key = "WSG_C4_BEFORE_MANUAL_MODE";
  const c4 = sh.getRange("C4");
  const manualMode = !!sh.getRange("K3").getValue() || !!sh.getRange("L3").getValue();

  const removeC4Protections_ = () => {
    try {
      const protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      protections.forEach(p => {
        const r = p.getRange();
        if (r && r.getA1Notation && r.getA1Notation() === "C4") {
          try { p.remove(); } catch (err) { if (err && (err.message || err.toString)) Logger.log("PDKW_WSG_onEdit_ remove protection: " + (err.message || err.toString())); }
        }
      });
    } catch (err) {
      if (err && (err.message || err.toString)) Logger.log("PDKW_WSG_onEdit_ protections: " + (err.message || err.toString()));
    }
  };

  if (isSwitchCell) {
    if (manualMode) {
      // zapisz poprzednią wartość tylko raz (pierwsze wejście w tryb ręczny)
      if (!props.getProperty(key)) {
        props.setProperty(key, String(c4.getDisplayValue() || "").trim());
      }
      removeC4Protections_();
      c4.clearDataValidations();
      c4.setValue("");
      SpreadsheetApp.flush();
    } else {
      // oba odznaczone -> wróć do auto i przywróć wartość sprzed trybu ręcznego
      const prev = String(props.getProperty(key) || "").trim();
      if (prev !== "") c4.setValue(prev);
      c4.setNumberFormat("0000");
      c4.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(1, 9999)
          .setAllowInvalid(false)
          .build()
      );
      removeC4Protections_();
      try {
        const p = c4.protect();
        p.setDescription("AUTO NR DOSTAWY (nie edytować ręcznie)");
        p.setWarningOnly(true);
      } catch (err) {
        if (err && (err.message || err.toString)) Logger.log("PDKW_WSG_onEdit_ protect C4: " + (err.message || err.toString()));
      }
      props.deleteProperty(key);
      SpreadsheetApp.flush();
    }
  }

  // RYLEX: K3 zaznaczone LUB dostawca D4 = RYLEX -> F4 zawsze "Jabłko"
  const supplier = String(sh.getRange("D4").getDisplayValue() || "").trim().toUpperCase();
  const forceJablko = !!sh.getRange("K3").getValue() || supplier === "RYLEX";
  if (forceJablko) {
    const f4 = sh.getRange("F4");
    const fruit = String(f4.getDisplayValue() || "").trim();
    if (fruit !== "Jabłko" && fruit !== "Jablko") {
      f4.setValue("Jabłko");
      SpreadsheetApp.getActiveSpreadsheet().toast("Dla RYLEX owoc jest zablokowany: Jabłko.", "WSG", 3);
    }
  }
}

/**
 * Tabelka jakości:
 * - Obszar: B{start}:E{start+3}
 * - B = nr, C:D = etykieta (scalone), E = wartość
 * - BRIX i ZWROT w % zawsze
 * - TWARDOŚĆ tylko S/O
 * - KALIBER <68mm % tylko O
 */
function KW_BUILD_QUALITY_TABLE_(shKW, purposeShort, startRow) {
  const p = String(purposeShort || "").toUpperCase().trim();

  const startCol = 2; // B
  const numRows = 4;
  const numCols = 4;  // B:C:D:E

  const full = shKW.getRange(startRow, startCol, numRows, numCols);

  // czyść wszystko żeby nie zostawały stare ramki / scalenia
  try { full.breakApart(); } catch (e) { if (e && (e.message || e.toString)) Logger.log("KW_BUILD_QUALITY_TABLE_ breakApart: " + (e.message || e.toString())); }
  full.clearContent();
  full.clearFormat();

  full.setFontSize(14).setVerticalAlignment("middle");

  const rows = [
    { idx: 0, show: true,               label: "BRIX" },
    { idx: 1, show: true,               label: "ZWROT w %" },
    { idx: 2, show: (p === "S" || p === "O"), label: "TWARDOŚĆ" },
    { idx: 3, show: (p === "O"),        label: "KALIBER PONIŻEJ 68mm w %" }
  ];

  rows.forEach(r => {
    const row = startRow + r.idx;
    if (!r.show) return;

    shKW.getRange(row, startCol)
      .setValue(r.idx + 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    const labelCell = shKW.getRange(row, startCol + 1, 1, 2);
    labelCell.merge();
    labelCell.setValue(r.label);
    labelCell.setFontWeight("bold");
    labelCell.setHorizontalAlignment("left");
    labelCell.setWrap(true);

    const valueCell = shKW.getRange(row, startCol + 3);
    valueCell.setHorizontalAlignment("left");
    valueCell.setWrap(true);

    shKW.getRange(row, startCol, 1, numCols).setBorder(true, true, true, true, true, true);
  });
}