/**
 * Jedna źródłowa konfiguracja layoutu KW i KWG (wiersze odmian, tabele jakości, ceny, twardość).
 * Używana w TABELEKW.js i KARTA_WAZENIA.js.
 */
var KW_KWG_LAYOUT = {
  KW: {
    varietyRowFirst: 19,
    varietyRowLast: 22,
    qualityStarts: [24, 30, 36, 42],
    priceRanges: ["H26:K27", "H32:K33", "H38:K39", "H44:K45"],
    bottomBorders: ["H25:K25", "H31:K31", "H37:K37", "H43:K43"],
    twardoscRows: [26, 32, 38, 44],
    blocksHideFromTo: [24, 46],
    blockShowHideRanges: [[24, 28], [29, 34], [35, 40], [41, 46]]
  },
  KWG: {
    varietyRowFirst: 14,
    varietyRowLast: 17,
    qualityStarts: [19, 25, 31, 37],
    blockShowHideRanges: [[18, 21], [22, 27], [28, 33], [34, 39]],
    blockShowHideRangesObieranie: [[18, 23], [24, 29], [30, 35], [36, 41]],
    priceRanges: ["H21:K22", "H27:K28", "H33:K34", "H39:K40"],
    bottomBorders: ["H20:K20", "H26:K26", "H32:K32", "H38:K38"],
    twardoscRows: [21, 27, 33, 39],
    blocksHideFromTo: [18, 39]
  }
};

function getLayout_(sheetName) {
  return sheetName === "KWG" ? KW_KWG_LAYOUT.KWG : KW_KWG_LAYOUT.KW;
}

/** Dla KWG zawsze qualityStarts [19,25,31,37] (BRIX/ZWROT/TWARDOŚĆ/KALIBER w 19–22, 25–28, 31–34, 37–40). */
function getEffectiveQualityStarts_(sheetName, purposeShort) {
  const layout = getLayout_(sheetName);
  if (!layout) return sheetName === "KWG" ? [19, 25, 31, 37] : [24, 30, 36, 42];
  return layout.qualityStarts;
}

/** Dla KWG + Obieranie (O) zwraca blockShowHideRangesObieranie; w pozostałych przypadkach layout.blockShowHideRanges. */
function getEffectiveBlockShowHideRanges_(sheetName, purposeShort) {
  const layout = getLayout_(sheetName);
  if (!layout || !layout.blockShowHideRanges) return null;
  if (sheetName === "KWG" && String(purposeShort || "").toUpperCase() === "O" && layout.blockShowHideRangesObieranie)
    return layout.blockShowHideRangesObieranie;
  return layout.blockShowHideRanges;
}

/** Dla KWG zawsze twardoscRows [21,27,33,39]. */
function getEffectiveTwardoscRows_(sheetName, purposeShort) {
  const layout = getLayout_(sheetName);
  if (!layout) return sheetName === "KWG" ? [21, 27, 33, 39] : [26, 32, 38, 44];
  return layout.twardoscRows;
}

function TABLEKW_onEdit_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const sheetName = sh.getName();
  if (sheetName !== "KW" && sheetName !== "KWG") return;

  const layout = getLayout_(sheetName);
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const inRect = (r1, r2, c1, c2) => (row >= r1 && row <= r2 && col >= c1 && col <= c2);

  const isVarietyAreaKW = (sheetName === "KW" && inRect(layout.varietyRowFirst, layout.varietyRowLast, 5, 7));
  const isCratesAreaKW  = (sheetName === "KW" && inRect(layout.varietyRowFirst, layout.varietyRowLast, 10, 11));
  const isVarietyAreaKWG = (sheetName === "KWG" && inRect(layout.varietyRowFirst, layout.varietyRowLast, 5, 7));
  const isCratesAreaKWG  = (sheetName === "KWG" && inRect(layout.varietyRowFirst, layout.varietyRowLast, 10, 11));

  const isLotArea = (row === 7 && col >= 6 && col <= 9);

  const isTwardoscCell = (col === 5 && layout.twardoscRows.indexOf(row) !== -1);

  if (isTwardoscCell) {
    const props = PropertiesService.getDocumentProperties();
    const guardKey = "KW_TWARDOSC_GUARD";
    if (props.getProperty(guardKey) === "1") return;

    const lotTextT = String(sh.getRange("F7:I7").getDisplayValue() || "").trim();
    const p = String(getPurposeShortFromLot_(sh, lotTextT) || "").toUpperCase();

    if (p === "S" || p === "O") {
      const raw = String(e.range.getDisplayValue() || "").trim();
      const num = parseNumeric_(raw);

      if (isFinite(num) && num < 4.5) {
        const ui = SpreadsheetApp.getUi();
        const res = ui.alert(
          "ZA MAŁA TWARDOŚĆ!",
          "Przekazać na przecier?",
          ui.ButtonSet.YES_NO
        );

        if (res === ui.Button.YES) {
          const lotRange = sh.getRange("F7:I7");
          if (lotTextT) {
            const newLot = replaceLastPurposeLetter_(lotTextT, p, "P");
            if (newLot !== lotTextT) {
              props.setProperty(guardKey, "1");
              try {
                lotRange.setValue(newLot);
                SpreadsheetApp.flush();
              } finally {
                props.deleteProperty(guardKey);
              }
              adjustAllQualityTables_(sh, getPurposeShortFromLot_(sh, newLot));
              adjustRightPriceRows_(sh, getPurposeShortFromLot_(sh, newLot));
              SpreadsheetApp.flush();
            }
          }
        }
      }
    }
    return;
  }

  // Jedno odczytanie F7:I7 dla wszystkich reakcji (odmiany/skrzynie/LOT)
  const lotTextOnce = (isVarietyAreaKW || isVarietyAreaKWG || isCratesAreaKW || isCratesAreaKWG || isLotArea)
    ? String(sh.getRange("F7:I7").getDisplayValue() || "").trim()
    : "";
  const purposeShortOnce = getPurposeShortFromLot_(sh, lotTextOnce);

  if (isVarietyAreaKW || isVarietyAreaKWG) {
    adjustAllQualityTables_(sh, purposeShortOnce);
    adjustRightPriceRows_(sh, purposeShortOnce);
    SpreadsheetApp.flush();
    return;
  }

  if (isCratesAreaKW || isCratesAreaKWG) {
    adjustRightPriceRows_(sh, purposeShortOnce);
    SpreadsheetApp.flush();
    return;
  }

  if (isLotArea) {
    adjustAllQualityTables_(sh, purposeShortOnce);
    adjustRightPriceRows_(sh, purposeShortOnce);
    SpreadsheetApp.flush();
    return;
  }
}

/**
 * Dopasowuje układ tabelek jakości do S/O/F/P (i innych).
 * purposeShortOptional – jeśli podany, używany zamiast odczytu z arkusza.
 */
function adjustAllQualityTables_(sh, purposeShortOptional) {
  const sheetName = sh.getName();
  const layout = getLayout_(sheetName);
  const purposeShort = (purposeShortOptional !== undefined && purposeShortOptional !== null)
    ? purposeShortOptional
    : getPurposeShortFromLot_(sh);
  const starts = getEffectiveQualityStarts_(sheetName, purposeShort);

  buildQualityTable_(sh, purposeShort, starts[0]);

  for (let i = 1; i < starts.length; i++) {
    const sr = starts[i];
    if (sh.isRowHiddenByUser(sr)) continue;
    buildQualityTable_(sh, purposeShort, sr);
  }
}

/**
 * Ukrywa/pokazuje (wizualnie) wiersze "W cenie zakupu" i "W cenie obniżonej"
 * POKAZUJ TYLKO dla O (Obieranie); dla S i reszty ukryte.
 * purposeShortOptional – jeśli podany, używany zamiast odczytu z arkusza.
 */
function adjustRightPriceRows_(sh, purposeShortOptional) {
  const sheetName = sh.getName();
  const p = (purposeShortOptional !== undefined && purposeShortOptional !== null)
    ? purposeShortOptional
    : getPurposeShortFromLot_(sh);

  const isObieranie = (String(p).toUpperCase() === "O");
  let simpleKWGMode = false;
  if (sheetName === "KWG") {
    try {
      const wsg = sh.getParent().getSheetByName("WSG");
      if (wsg) simpleKWGMode = !!wsg.getRange("J3").getValue() || !!wsg.getRange("K3").getValue();
    } catch (e) {
      if (e && (e.message || e.toString)) Logger.log("TABELEKW adjustRightPriceRows_ simpleKWGMode: " + (e.message || e.toString()));
    }
  }
  const shouldHide = simpleKWGMode || !isObieranie; // w trybie uproszczonym zawsze ukryj

  const ranges = getPriceRanges_(sheetName);
  ranges.forEach(a1 => {
    const r = sh.getRange(a1);
    if (shouldHide) hideVisually_(r);
    else showVisually_(r);
  });

  // FIX: jak ukryte -> wymuś dolne ramki (separator jak w przecierze)
  if (shouldHide) {
    enforcePrzecierBottomBordersKW_(sh);
    enforcePrzecierBottomBordersKWG_(sh);
  }

  SpreadsheetApp.flush();
}

/**
 * Zakresy (2 wiersze) do "ukrywania" dla każdej odmiany (KW i KWG).
 */
function getPriceRanges_(sheetName) {
  return getLayout_(sheetName).priceRanges;
}

/**
 * Ukrycie bez kasowania formuł: białe tło + biały font + brak ramek.
 */
function hideVisually_(range) {
  range
    .setBackground("#ffffff")
    .setFontColor("#ffffff")
    .setBorder(false, false, false, false, false, false);
}

/**
 * Przywrócenie: czarny font + białe tło + ramki.
 */
function showVisually_(range) {
  range
    .setBackground("#ffffff")
    .setFontColor("#000000")
    .setBorder(true, true, true, true, true, true);
}

/**
 * "Separator" dolnych ramek (KW) – zostają zawsze gdy prawa strona jest ukryta
 */
function enforcePrzecierBottomBordersKW_(sh) {
  if (!sh || sh.getName() !== "KW") return;
  getLayout_("KW").bottomBorders.forEach(a1 => {
    sh.getRange(a1).setBorder(null, null, true, null, null, null);
  });
}

/**
 * "Separator" dolnych ramek (KWG) – zostają zawsze gdy prawa strona jest ukryta
 */
function enforcePrzecierBottomBordersKWG_(sh) {
  if (!sh || sh.getName() !== "KWG") return;
  getLayout_("KWG").bottomBorders.forEach(a1 => {
    sh.getRange(a1).setBorder(null, null, true, null, null, null);
  });
}

/**
 * Litera przeznaczenia z końca LOT-a w F7:I7.
 * Jeśli podano lotText, używa go zamiast odczytu z arkusza (optymalizacja: jeden odczyt w onEdit).
 */
function getPurposeShortFromLot_(sh, lotText) {
  const s = (lotText !== undefined && lotText !== null)
    ? String(lotText).trim()
    : String(sh.getRange("F7:I7").getDisplayValue() || "").trim();
  if (!s) return "";
  let m = s.match(/-\s*([A-Za-z])\s*$/);
  if (m && m[1]) return String(m[1]).toUpperCase();
  m = s.match(/([A-Za-z])\s*$/);
  if (m && m[1]) return String(m[1]).toUpperCase();
  return "";
}

/**
 * Buduje tabelkę jakości w obszarze B{start}:E{start+3}
 * FIX: wymusza dół na ostatnim widocznym wierszu.
 */
function buildQualityTable_(sh, purposeShort, startRow) {
  const p = String(purposeShort || "").toUpperCase().trim();

  const startCol = 2; // B
  const numRows = 4;
  const numCols = 4; // B:C:D:E

  const full = sh.getRange(startRow, startCol, numRows, numCols);

  // zachowaj wartości w kolumnie E
  const oldValues = full.getValues();
  const oldE = oldValues.map(r => r[3]);

  try { full.breakApart(); } catch (e) { if (e && (e.message || e.toString)) Logger.log("TABELEKW buildQualityTable_ breakApart: " + (e.message || e.toString())); }
  full.clearContent();

  // KWG + zaznaczone J3 (RYLEX) lub K3 (GRÓJECKA) w WSG:
  // - bez ZWROT i bez KALIBER
  // - BRIX zawsze
  // - TWARDOŚĆ tylko dla S/O (dla P tylko BRIX)
  let simpleKWGMode = false;
  if (sh.getName() === "KWG") {
    try {
      const wsg = sh.getParent().getSheetByName("WSG");
      if (wsg) {
        simpleKWGMode = !!wsg.getRange("J3").getValue() || !!wsg.getRange("K3").getValue();
      }
    } catch (e) {
      if (e && (e.message || e.toString)) Logger.log("TABELEKW simpleKWGMode: " + (e.message || e.toString()));
    }
  }

  const defs = simpleKWGMode
    ? [
        { sourceIdx: 0, show: true,                     label: "BRIX" },
        { sourceIdx: 2, show: (p === "S" || p === "O"), label: "TWARDOŚĆ" }
      ]
    : [
        { sourceIdx: 0, show: true,                     label: "BRIX" },
        { sourceIdx: 1, show: true,                     label: "ZWROT w %" },
        { sourceIdx: 2, show: (p === "S" || p === "O"), label: "TWARDOŚĆ" },
        { sourceIdx: 3, show: (p === "O"),              label: "KALIBER PONIŻEJ 68mm w %" }
      ];

  let visualRow = 0;
  defs.forEach(d => {
    if (!d.show) return;

    const rr = startRow + visualRow;
    const rowRange = sh.getRange(rr, startCol, 1, numCols);

    sh.getRange(rr, startCol).setValue(visualRow + 1);

    const labelCell = sh.getRange(rr, startCol + 1, 1, 2);
    labelCell.merge();
    labelCell.setValue(d.label);

    const valueCell = sh.getRange(rr, startCol + 3);
    if (oldE[d.sourceIdx] !== "" && oldE[d.sourceIdx] !== null) valueCell.setValue(oldE[d.sourceIdx]);

    rowRange.setBorder(true, true, true, true, true, true);
    visualRow++;
  });

  // wyczyść i odramuj nieużyte wiersze, żeby nie było "pustych tabelek"
  for (let i = visualRow; i < 4; i++) {
    const rr = startRow + i;
    const rowRange = sh.getRange(rr, startCol, 1, numCols);
    try { rowRange.breakApart(); } catch (e) {}
    rowRange.clearContent();
    rowRange.setBorder(false, false, false, false, false, false);
  }

  if (visualRow > 0) {
    const rr = startRow + visualRow - 1;
    sh.getRange(rr, startCol, 1, numCols).setBorder(null, null, true, null, null, null);
  }
}

/** parsuje liczbę z wpisu (akceptuje "," i ".") */
function parseNumeric_(v) {
  const s = String(v == null ? "" : v).trim().replace(/\s+/g, "").replace(",", ".");
  if (s === "") return NaN;
  if (!/^-?\d+(\.\d+)?$/.test(s)) return NaN;
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

/**
 * Zamienia ostatnią literę przeznaczenia w LOT (np. ... S -> ... P, ... O -> ... P)
 * Obsługuje:
 * - "... - S" / "... - O"
 * - "... S" / "... O"
 */
function replaceLastPurposeLetter_(lotText, fromLetter, toLetter) {
  const s = String(lotText || "").trim();
  if (!s) return s;

  const from = String(fromLetter || "").toUpperCase();
  const to = String(toLetter || "").toUpperCase();

  // wariant z myślnikiem: "- X"
  let m = s.match(/-\s*([A-Za-z])\s*$/);
  if (m && m[1] && String(m[1]).toUpperCase() === from) {
    return s.replace(/-\s*[A-Za-z]\s*$/, "- " + to);
  }

  // wariant bez myślnika: "... X"
  m = s.match(/([A-Za-z])\s*$/);
  if (m && m[1] && String(m[1]).toUpperCase() === from) {
    return s.replace(/[A-Za-z]\s*$/, to);
  }

  return s;
}