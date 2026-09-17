/**
 * Kentucky Lottery PPS four-tab builder (browser + Node).
 * Ports .cursor/skills/ky-lottery-pps for Delivery, Odds Table,
 * Summary(Delivery), and Prize Breakdown. Excel calculates formulas on open.
 */

const ExcelJS = globalThis.ExcelJS;

export const DELIVERY_SHEETS = [
  "Delivery",
  "Odds Table",
  "Summary(Delivery)",
  "Prize Breakdown",
];

export const PRICE_GRIDS = {
  mmj3: [0.5, 1, 2, 5, 10, 20, 30, 50],
  "no-jp": [0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 30, 50],
  "daily-streak": [0.5, 1, 2, 3, 5, 10, 20, 30, 50],
  ssj: null,
};

const STANDARD_COLORS = {
  instanta: "FFDDEBF7",
  bonusa: "FFE2EFDA",
  freeplaya: "FFFFF2CC",
  freeplayb: "FFFFE699",
  freeplayc: "FFFFD966",
};

const CONFIDENTIAL =
  "The information contained in this document and all attached documents is strictly confidential " +
  "and contains proprietary information.\n" +
  "It is provided solely for use by the designated recipients and is subject to the terms of any " +
  "confidentiality obligations or non-disclosure agreements between the parties.\n" +
  "All other use is strictly prohibited.";

const X_SUFFIX = /\s+x\s+\d+(?:\.\d+)?\s*$/i;
const NO_FILL = { type: "pattern", pattern: "none" };
const NO_BORDER = {};

function isNumeric(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function asNumber(value, fallback = 0) {
  if (isNumeric(value)) return value;
  if (typeof value === "string" && value.trim() !== "") {
    const n = Number(value);
    if (Number.isFinite(n)) return n;
  }
  return fallback;
}

function cleanFloat(value, digits = 10) {
  const n = asNumber(value);
  if (Math.abs(n - Math.round(n)) < 10 ** -digits) return Math.round(n);
  return Number(n.toFixed(digits));
}

function colLetter(n) {
  let result = "";
  let number = n;
  while (number > 0) {
    const rem = (number - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function colNumber(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function parseAddr(addr) {
  const m = String(addr).match(/^([A-Z]+)(\d+)$/i);
  if (!m) return { col: 1, row: 1 };
  return { col: colNumber(m[1]), row: Number(m[2]) };
}

export function cellResult(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (v == null || v === "") return null;
  if (typeof v !== "object") return v;
  if (Array.isArray(v)) return v.map((p) => p.text ?? "").join("");
  if (v.richText) return v.richText.map((p) => p.text ?? "").join("");
  if (v.hyperlink) return v.text ?? v.hyperlink;
  if (v.error) return null;
  if ("formula" in v || "sharedFormula" in v) {
    return v.result === undefined ? null : v.result;
  }
  return v;
}

function cellFormula(cell) {
  const v = cell?.value;
  if (v && typeof v === "object" && v.formula) return `=${v.formula}`;
  if (v && typeof v === "object" && v.sharedFormula) return `=${v.sharedFormula}`;
  if (typeof v === "string" && v.startsWith("=")) return v;
  return null;
}

function setValue(ws, addr, value) {
  const cell = typeof addr === "string" ? ws.getCell(addr) : ws.getCell(addr.row, addr.col);
  if (typeof value === "string" && value.startsWith("=")) {
    cell.value = { formula: value.slice(1) };
  } else {
    cell.value = value;
  }
  return cell;
}

function clone(obj) {
  if (obj == null) return obj;
  return JSON.parse(JSON.stringify(obj));
}

function copyStyle(src, dst) {
  if (!src || !dst) return;
  try {
    if (src.font && Object.keys(src.font).length) dst.font = clone(src.font);
    if (src.fill && src.fill.type) dst.fill = clone(src.fill);
    if (src.border && Object.keys(src.border).length) dst.border = clone(src.border);
    if (src.alignment && Object.keys(src.alignment).length) dst.alignment = clone(src.alignment);
    if (src.protection && Object.keys(src.protection).length) dst.protection = clone(src.protection);
    if (src.numFmt) dst.numFmt = src.numFmt;
  } catch {
    /* theme-bound styles can fail to clone; values still write */
  }
}

function copyRowStyle(srcWs, dstWs, srcRow, dstRow, minCol, maxCol) {
  const sRow = srcWs.getRow(srcRow);
  const dRow = dstWs.getRow(dstRow);
  if (sRow.height) dRow.height = sRow.height;
  for (let col = minCol; col <= maxCol; col += 1) {
    copyStyle(srcWs.getCell(srcRow, col), dstWs.getCell(dstRow, col));
  }
}

function clearFillBorder(cell) {
  cell.fill = clone(NO_FILL);
  cell.border = clone(NO_BORDER);
}

function copyColumnWidths(srcWs, dstWs, maxCol) {
  for (let col = 1; col <= maxCol; col += 1) {
    const width = srcWs.getColumn(col).width;
    if (width) dstWs.getColumn(col).width = width;
  }
}

function copySheetChrome(srcWs, dstWs) {
  if (srcWs.properties) dstWs.properties = { ...dstWs.properties, ...clone(srcWs.properties) };
  if (srcWs.pageSetup) dstWs.pageSetup = { ...srcWs.pageSetup };
  if (srcWs.headerFooter) dstWs.headerFooter = clone(srcWs.headerFooter);
  if (srcWs.views?.length) dstWs.views = clone(srcWs.views);
  dstWs.showGridLines = srcWs.showGridLines;
}

function localizeFrequencyRefs(formula) {
  return String(formula).replace(/'?Frequency'?!\$?N\$?(\d+)/gi, (_, n) => `$M$${n}`);
}

function shiftFormula(formula, dCol, dRow) {
  return String(formula).replace(
    /(^|[^A-Z0-9_])(\$?)([A-Z]{1,3})(\$?)(\d+)/g,
    (match, pre, colAbs, col, rowAbs, row) => {
      let c = colNumber(col);
      let r = Number(row);
      if (!colAbs) c += dCol;
      if (!rowAbs) r += dRow;
      if (c < 1 || r < 1) return match;
      return `${pre}${colAbs}${colLetter(c)}${rowAbs}${r}`;
    },
  );
}

function translatedFormula(formula, origin, destination) {
  if (!formula || !String(formula).startsWith("=")) return formula;
  const text = localizeFrequencyRefs(formula);
  const from = parseAddr(origin);
  const to = parseAddr(destination);
  return shiftFormula(text, to.col - from.col, to.row - from.row);
}

function methodPrefix(method) {
  const token = String(method || "")
    .split("|")[0]
    .trim();
  return token.split("-")[0].trim().toLowerCase();
}

function methodForColor(method) {
  return String(method || "").replace(X_SUFFIX, "").trim();
}

function singleTierColorPrefix(method) {
  const tokens = methodForColor(method)
    .split("|")
    .map((t) => t.trim())
    .filter(Boolean);
  if (tokens.length !== 1) return null;
  const prefix = tokens[0].split("-")[0].trim().toLowerCase();
  if (prefix.startsWith("freeplay") || prefix.startsWith("instant") || prefix.startsWith("bonus")) {
    return prefix;
  }
  return null;
}

function argbFill(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function winMethodFills(winMethods) {
  const fills = {};
  for (const [key, argb] of Object.entries(STANDARD_COLORS)) fills[key] = argbFill(argb);
  if (!winMethods) return fills;
  const max = Math.min(winMethods.rowCount || 0, 400);
  for (let row = 5; row <= max; row += 1) {
    const method = cellResult(winMethods.getCell(row, 3));
    if (!method) continue;
    const prefix = methodPrefix(method);
    if (
      (prefix.startsWith("freeplay") || prefix.startsWith("instant") || prefix.startsWith("bonus")) &&
      !fills[prefix]
    ) {
      const fill = winMethods.getCell(row, 3).fill;
      if (fill && fill.fgColor) fills[prefix] = clone(fill);
    }
  }
  return fills;
}

function applyMethodFill(cell, method, fills) {
  const prefix = singleTierColorPrefix(method);
  if (prefix && fills[prefix]) cell.fill = clone(fills[prefix]);
  else cell.fill = clone(NO_FILL);
}

function sheet(wb, name) {
  return wb.getWorksheet(name) || wb.worksheets.find((ws) => ws.name.toLowerCase() === name.toLowerCase());
}

function requireSheet(wb, name) {
  const ws = sheet(wb, name);
  if (!ws) throw new Error(`Missing sheet: ${name}`);
  return ws;
}

function identityText(freq) {
  const parts = [];
  for (let row = 2; row <= 8; row += 1) {
    for (let col = 12; col <= 16; col += 1) {
      const v = cellResult(freq.getCell(row, col));
      if (v != null && v !== "") parts.push(String(v));
    }
    for (const col of [11, 12]) {
      const v = cellResult(freq.getCell(row, col));
      if (v != null && v !== "") parts.push(String(v));
    }
  }
  for (const addr of ["L2", "L3", "L4", "M4", "N4"]) {
    const v = cellResult(freq.getCell(addr));
    if (v != null && v !== "") parts.push(String(v));
  }
  return parts.join(" ");
}

function isKentucky(identity, filename) {
  const blob = `${identity} ${filename}`.toLowerCase();
  return blob.includes("kentucky") || /(^|[^a-z])ky([^a-z]|$)/.test(blob) || blob.includes("ky_");
}

function progressiveJackpots(pj) {
  if (!pj) return [];
  const entries = [];
  for (let row = 11; row <= 13; row += 1) {
    const number = cellResult(pj.getCell(row, 2));
    const name = cellResult(pj.getCell(row, 3));
    const prize = cellResult(pj.getCell(row, 5));
    const wins = cellResult(pj.getCell(row, 6));
    const odds = cellResult(pj.getCell(row, 8));
    if (number == null || number === "" || prize == null || prize === "" || odds == null || odds === "") {
      continue;
    }
    if (!isNumeric(asNumber(prize, NaN))) continue;
    entries.push({
      row,
      name: String(name || "").trim(),
      label: `JP${entries.length + 1}`,
      wins: cleanFloat(wins),
      prize: cleanFloat(prize),
      odds: cleanFloat(odds),
    });
  }
  return entries;
}

function ssjJackpots(entries) {
  const byName = {};
  for (const item of entries) {
    const key = item.name.toLowerCase();
    if (key) byName[key] = item;
  }
  if (!byName.small || !byName.large) {
    throw new Error(`SSJ jackpots must include small and large, found ${entries.map((e) => e.name).join(", ") || "none"}`);
  }
  return {
    small: byName.small,
    large: byName.large,
    oddsOrder: [
      { ...byName.large, label: "JP1" },
      { ...byName.small, label: "JP2" },
    ],
    summaryJp1: byName.small,
    summaryJp2: byName.large,
  };
}

function detectSchema({ identity, filename, jpEntries, hasPjSheet }) {
  const blob = `${identity} ${filename}`.toLowerCase();
  if (!jpEntries.length) {
    if (/daily\s*streak|dailystreak/.test(blob)) return "daily-streak";
    return "no-jp";
  }
  if (!hasPjSheet) return "no-jp";
  const names = jpEntries.map((e) => e.name.toLowerCase());
  if (names.includes("small") && names.includes("large")) return "ssj";
  if (/\bssj\b/.test(blob)) return "ssj";
  if (/\bmmj3\b/.test(blob)) return "mmj3";
  if (jpEntries.length <= 2) return "mmj3";
  throw new Error("Unsupported jackpot layout (more than two JP rows, and not SSJ small/large).");
}

function sourceRows(freq) {
  const rows = [];
  const max = Math.min(freq.rowCount || 0, 2000);
  for (let row = 5; row <= max; row += 1) {
    const method = cellResult(freq.getCell(row, 3));
    const prize = cellResult(freq.getCell(row, 4));
    const winners = cellResult(freq.getCell(row, 5));
    if (method == null || method === "" || prize == null || prize === "" || winners == null || winners === "") {
      continue;
    }
    const prizeNum = asNumber(prize, NaN);
    if (!Number.isFinite(prizeNum) || prizeNum <= 0) continue;
    const winNum = asNumber(winners, NaN);
    if (!Number.isFinite(winNum)) {
      throw new Error(`Frequency!E${row}: missing/non-numeric Odds up`);
    }
    rows.push({
      sourceRow: row,
      method: String(method),
      prize: cleanFloat(prizeNum),
      winners: cleanFloat(winNum),
    });
  }
  rows.sort((a, b) => b.prize - a.prize || a.sourceRow - b.sourceRow);
  if (!rows.length) throw new Error("No winning Frequency rows found.");
  const seen = new Map();
  for (const row of rows) seen.set(row.method, (seen.get(row.method) || 0) + 1);
  const duplicateMethods = [...seen.entries()].filter(([, n]) => n > 1).map(([k]) => k);
  if (rows.some((r) => r.winners <= 0)) {
    throw new Error("Zero/negative winning frequencies need a deliberate odds policy before building.");
  }
  rows.duplicateMethods = duplicateMethods;
  return rows;
}

function freqNumber(freq, row) {
  return asNumber(cellResult(freq.getCell(row, 14)), NaN);
}

function priceGridBaseCol(grid) {
  const idx = grid.findIndex((p) => Math.abs(p - 1) < 1e-9);
  if (idx < 0) throw new Error("The price grid must contain a $1 base column.");
  return idx + 2;
}

function findDeliveryTemplateRows(templateWs, requireJp) {
  let nonwinRow = null;
  let jpStart = null;
  const max = templateWs.rowCount || 80;
  for (let row = 1; row <= max; row += 1) {
    const method = cellResult(templateWs.getCell(row, 3));
    if (method === "Non Winning tickets" || method === "NON WINNING TIER") nonwinRow = row;
    if (cellResult(templateWs.getCell(row, 2)) === "Base JP Odds Down") jpStart = row;
  }
  if (nonwinRow == null || (requireJp && jpStart == null)) {
    throw new Error("Could not find Delivery template non-winning or JP rows.");
  }
  return { nonwinRow, totalRow: nonwinRow + 1, jpStart };
}

function findOddsTemplateNonwin(templateWs) {
  const maxR = templateWs.rowCount || 80;
  const maxC = templateWs.columnCount || 20;
  for (let row = 1; row <= maxR; row += 1) {
    for (let col = 1; col <= maxC; col += 1) {
      if (cellResult(templateWs.getCell(row, col)) === "NON WINNING") return row;
    }
  }
  return templateWs.rowCount || 20;
}

function summaryLookup(summary, tokens, fallback) {
  if (!summary) return fallback;
  const needles = tokens.map((t) => t.toLowerCase());
  const max = summary.rowCount || 80;
  for (let row = 1; row <= max; row += 1) {
    const label = String(cellResult(summary.getCell(row, 2)) || "").toLowerCase();
    if (needles.every((t) => label.includes(t))) return cellResult(summary.getCell(row, 3));
  }
  return fallback;
}

function parseCategorySpan(label) {
  if (isNumeric(label)) return [label, label];
  const text = String(label || "").trim();
  if (text === "=1" || text === "1") return [1, 1];
  let m = text.match(/^>([0-9.]+)\s+to\s+<?([0-9.]+)/);
  if (m) return [Number(m[1]), Number(m[2])];
  m = text.match(/^>([0-9.]+)$/);
  if (m) return [Number(m[1]), Infinity];
  return null;
}

function summaryDistribution(summary) {
  let breakeven = 0;
  let low = 0;
  let medium = 0;
  let high = 0;
  if (!summary) return { breakeven, low, medium, high };
  const max = summary.rowCount || 80;
  for (let row = 1; row <= max; row += 1) {
    const span = parseCategorySpan(cellResult(summary.getCell(row, 2)));
    if (!span) continue;
    const [lo, hi] = span;
    const valueDist = asNumber(cellResult(summary.getCell(row, 6)));
    const prob = asNumber(cellResult(summary.getCell(row, 4)));
    if (lo === 1 && hi === 1) breakeven = prob;
    if (hi <= 5) low += valueDist;
    else if (lo >= 5 && hi <= 20) medium += valueDist;
    else if (lo >= 20) high += valueDist;
  }
  return {
    breakeven: cleanFloat(breakeven, 12),
    low: cleanFloat(low, 12),
    medium: cleanFloat(medium, 12),
    high: cleanFloat(high, 12),
  };
}

function currencyStar(value) {
  return `$${asNumber(value).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}*`;
}

function insertSheet(wb, name) {
  const existing = sheet(wb, name);
  if (existing) wb.removeWorksheet(existing.id);
  return wb.addWorksheet(name);
}

function orderSheets(wb) {
  const byName = Object.fromEntries(wb.worksheets.map((ws) => [ws.name, ws]));
  const head = ["Frequency", "Progressive Jackpots"].map((n) => byName[n]).filter(Boolean);
  const mid = DELIVERY_SHEETS.map((n) => byName[n]).filter(Boolean);
  const skip = new Set([...head, ...mid]);
  const tail = wb.worksheets.filter((ws) => !skip.has(ws));
  const list = [...head, ...mid, ...tail];
  const sparse = [];
  list.forEach((ws, i) => {
    ws.id = i + 1;
    ws.orderNo = i;
    sparse[i + 1] = ws;
  });
  wb._worksheets = sparse;
}

function groupPrizes(rows, prizeKey) {
  const grouped = new Map();
  for (const item of rows) {
    const prize = cleanFloat(item[prizeKey]);
    if (prize <= 0) continue;
    grouped.set(prize, (grouped.get(prize) || 0) + item.winners);
  }
  return [...grouped.keys()].sort((a, b) => b - a);
}

function buildDelivery(ws, templateWs, src, rows, ctx) {
  const { schema, fills, freq, pj, jpEntries, ssj } = ctx;
  const hasJp = schema === "ssj" || schema === "mmj3";
  const lastWinRow = 3 + rows.length;
  const nonwinRow = lastWinRow + 1;
  const totalRow = nonwinRow + 1;
  const jpStart = hasJp ? totalRow + 2 : null;
  const base = ctx.base;
  const tpl = findDeliveryTemplateRows(templateWs, hasJp);

  copySheetChrome(templateWs, ws);
  copyColumnWidths(templateWs, ws, hasJp ? 24 : 15);
  for (let row = 1; row <= 3; row += 1) copyRowStyle(templateWs, ws, row, row, 1, 24);

  const headers = [
    "DIVISION",
    "METHOD",
    "PRIZE FACTOR",
    "PRICE POINT",
    "EXPECTED WINNERS PER CARD",
    "ODDS PER GRID",
    "TOTAL PRIZE",
    "PRIZE % ",
    "PAYOUT %",
  ];
  headers.forEach((h, i) => setValue(ws, { row: 2, col: i + 2 }, h));

  rows.forEach((item, idx) => {
    const row = 4 + idx;
    const factor = schema === "ssj" ? cleanFloat(item.prize / base) : item.prize;
    setValue(ws, `B${row}`, idx + 1).numFmt = "00";
    setValue(ws, `C${row}`, item.method);
    applyMethodFill(ws.getCell(`C${row}`), item.method, fills);
    setValue(ws, `D${row}`, factor);
    setValue(ws, `E${row}`, `=D${row}*$M$7`);
    setValue(ws, `F${row}`, item.winners);
    setValue(ws, `G${row}`, `=$M$6/F${row}`);
    setValue(ws, `H${row}`, `=F${row}*E${row}`);
    setValue(ws, `I${row}`, `=H${row}/$M$9`);
    setValue(ws, `J${row}`, `=H${row}/$M$8`);
    copyRowStyle(templateWs, ws, 4, row, 2, 10);
    applyMethodFill(ws.getCell(`C${row}`), item.method, fills);
    if (row >= 24) {
      for (let col = 11; col <= 24; col += 1) clearFillBorder(ws.getCell(row, col));
    }
    for (const col of ["D", "E", "F", "G", "H"]) ws.getCell(`${col}${row}`).numFmt = "#,##0.00";
    ws.getCell(`I${row}`).numFmt = "0.00%";
    ws.getCell(`J${row}`).numFmt = "0.00%";
  });

  copyRowStyle(templateWs, ws, tpl.nonwinRow, nonwinRow, 2, 10);
  setValue(ws, `B${nonwinRow}`, rows.length + 1).numFmt = "00";
  setValue(ws, `C${nonwinRow}`, hasJp ? "Non Winning tickets" : "NON WINNING TIER");
  ws.getCell(`C${nonwinRow}`).fill = clone(NO_FILL);
  setValue(ws, `D${nonwinRow}`, 0);
  setValue(ws, `E${nonwinRow}`, `=D${nonwinRow}*$M$7`);
  setValue(ws, `F${nonwinRow}`, hasJp ? `=$M$6-$M$14` : `=$M$6-$M$12`);
  setValue(ws, `G${nonwinRow}`, `=$M$6/F${nonwinRow}`);
  setValue(ws, `H${nonwinRow}`, `=F${nonwinRow}*E${nonwinRow}`);
  setValue(ws, `I${nonwinRow}`, `=H${nonwinRow}/$M$9`);
  setValue(ws, `J${nonwinRow}`, `=H${nonwinRow}/$M$8`);

  copyRowStyle(templateWs, ws, tpl.totalRow, totalRow, 2, 10);
  setValue(ws, `F${totalRow}`, `=SUM(F4:F${nonwinRow})`);
  setValue(ws, `H${totalRow}`, `=SUM(H4:H${nonwinRow})`);
  setValue(ws, `I${totalRow}`, `=SUM(I4:I${nonwinRow})`);
  setValue(ws, `J${totalRow}`, `=SUM(J4:J${nonwinRow})`);
  for (const row of [nonwinRow, totalRow]) {
    for (const col of ["D", "E", "F", "G", "H"]) ws.getCell(`${col}${row}`).numFmt = "#,##0.00";
    ws.getCell(`I${row}`).numFmt = "0.00%";
    ws.getCell(`J${row}`).numFmt = "0.00%";
  }

  for (let row = 2; row <= 23; row += 1) copyRowStyle(templateWs, ws, row, row, 12, 15);
  setValue(ws, "L2", "Prize Structure");
  setValue(ws, "L3", "Kentucky Lottery");
  setValue(ws, "L4", cellResult(freq.getCell("M4")) || cellResult(freq.getCell("L4")) || "");
  setValue(ws, "L6", "Odds Down (ticket quantity if pool based):");
  setValue(ws, "M6", ctx.quantity);
  setValue(ws, "L7", "Base:");
  setValue(ws, "M7", ctx.base);
  setValue(ws, "L8", "Revenus (if pool based):");
  setValue(ws, "M8", "=M7*M6");
  setValue(ws, "L9", "Prize Fund (if pool based):");
  setValue(ws, "M9", `=H${totalRow}`);
  setValue(ws, "L10", "RTP Setting");
  setValue(ws, "M10", ctx.rtp);
  setValue(ws, "L11", "Actual RTP:");
  setValue(ws, "M11", "=M9/M8");
  ws.getCell("M10").numFmt = "0.00%";
  ws.getCell("M11").numFmt = "0.00%";

  if (hasJp) {
    setValue(ws, "N11", '=IF(ABS(M11-M10)<1E-10,"okay","error")');
    setValue(ws, "L12", "JP RTP:");
    if (schema === "ssj") {
      setValue(ws, "M12", `=C${jpStart + 12}+C${jpStart + 19}`);
    } else {
      setValue(ws, "M12", "='Progressive Jackpots'!$C$14+'Progressive Jackpots'!$C$21");
    }
    setValue(ws, "L13", "Total RTP:");
    setValue(ws, "M13", "=M12+M11");
    setValue(ws, "L14", "Winning Tiers Freq");
    setValue(ws, "M14", `=SUM(F4:F${lastWinRow})`);
    setValue(ws, "L15", "Hit Rate:");
    setValue(ws, "M15", "=M6/M14");
    ws.getCell("M12").numFmt = "0.00%";
    ws.getCell("M13").numFmt = "0.00%";
    if (schema === "ssj") {
      setValue(ws, "L17", CONFIDENTIAL);
      try {
        ws.mergeCells("L17:O21");
      } catch {
        /* already merged */
      }
    } else {
      setValue(ws, "L17", "Approx Count:");
      setValue(ws, "M17", cleanFloat(cellResult(freq.getCell("N17"))));
      setValue(ws, "L18", "Winning Tiers");
      setValue(ws, "M18", `=COUNT(F4:F${lastWinRow})`);
      setValue(ws, "L19", "Min per tier");
      setValue(ws, "M19", cleanFloat(cellResult(freq.getCell("N19"))));
      setValue(ws, "L21", "Difference:");
      setValue(ws, "M21", "=M8*M10-M9");
      setValue(ws, "L22", "Actual RTP Check:");
      setValue(ws, "M22", `=SUM(H4:H${lastWinRow})/M8`);
      setValue(ws, "L23", "Actual Hit Rate check");
      setValue(ws, "M23", `=SUM(F4:F${lastWinRow})`);
      setValue(ws, "N23", '=IF(M15=M23,"okay","error")');
      ws.getCell("M22").numFmt = "0.00%";
    }
  } else {
    setValue(ws, "L12", "Winning Tiers Freq");
    setValue(ws, "M12", `=SUM(F4:F${lastWinRow})`);
    setValue(ws, "L13", "Hit Rate:");
    setValue(ws, "M13", "=M6/M12");
    ws.getCell("M13").numFmt = "#,##0.00";
    const l16 = cellResult(templateWs.getCell("L16"));
    if (l16) setValue(ws, "L16", l16);
  }

  if (hasJp && pj) {
    copyRowStyle(templateWs, ws, tpl.totalRow + 1, totalRow + 1, 2, 24);
    for (let srcRow = 2; srcRow <= 32; srcRow += 1) {
      const dstRow = jpStart + (srcRow - 2);
      copyRowStyle(templateWs, ws, tpl.jpStart + (srcRow - 2), dstRow, 2, 24);
      for (let col = 2; col <= 24; col += 1) {
        if (schema === "ssj" && srcRow <= 7 && col >= 5) {
          ws.getCell(dstRow, col).value = null;
          continue;
        }
        const srcCell = pj.getCell(srcRow, col);
        const dstCell = ws.getCell(dstRow, col);
        const formula = cellFormula(srcCell);
        if (formula) {
          dstCell.value = {
            formula: translatedFormula(formula, srcCell.address, dstCell.address).replace(/^=/, ""),
          };
        } else {
          dstCell.value = cellResult(srcCell);
        }
      }
    }
    if (schema === "ssj") {
      setValue(ws, `C${jpStart + 23}`, `=C${jpStart}/F${jpStart + 12}`);
      setValue(ws, `C${jpStart + 24}`, `=1/((1/C${jpStart + 23})+(1/M15))`);
      setValue(ws, `C${jpStart + 25}`, `=M6-(F${jpStart + 12}/(G${jpStart + 10}/M6)+M14)`);
      setValue(ws, `C${jpStart + 26}`, "=M6");
      setValue(ws, `C${jpStart + 27}`, "=M10");
      setValue(ws, `C${jpStart + 28}`, `=C${jpStart + 12}`);
      setValue(ws, `C${jpStart + 29}`, `=C${jpStart + 19}`);
      setValue(ws, `C${jpStart + 30}`, `=C${jpStart + 27}+C${jpStart + 28}+C${jpStart + 29}`);
    }
  }

  try {
    ws.mergeCells("B2:B3");
    ws.mergeCells("C2:C3");
    ws.mergeCells("D2:D3");
    ws.mergeCells("E2:E3");
    ws.mergeCells("F2:F3");
    ws.mergeCells("G2:G3");
    ws.mergeCells("H2:H3");
    ws.mergeCells("I2:I3");
    ws.mergeCells("J2:J3");
  } catch {
    /* template may already encode merges via copy */
  }

  return { lastWinRow, nonwinRow, totalRow, jpStart };
}

function reflowOddsJp(ws, templateWs, jpCount, firstPrizeRow, nonwinRow, maxCol) {
  copyRowStyle(templateWs, ws, 1, 1, 1, maxCol);
  for (let offset = 0; offset < jpCount; offset += 1) {
    copyRowStyle(templateWs, ws, Math.min(2 + offset, 3), 2 + offset, 1, maxCol);
  }
  for (let row = firstPrizeRow; row < nonwinRow; row += 1) {
    copyRowStyle(templateWs, ws, row === firstPrizeRow ? 4 : 5, row, 1, maxCol);
  }
  copyRowStyle(templateWs, ws, findOddsTemplateNonwin(templateWs), nonwinRow, 1, maxCol);
}

function buildOddsSsj(ws, templateWs, rows, delivery, ctx) {
  const prizes = groupPrizes(rows, "prize");
  const lastWin = delivery.lastWinRow;
  copySheetChrome(templateWs, ws);
  ws.getColumn(1).width = templateWs.getColumn(1).width || 14;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 16;
  setValue(ws, "A1", "Ticket Price");
  setValue(ws, "B1", ctx.base).numFmt = "0.##";
  setValue(ws, "C1", "1 in X Odds ");
  ctx.ssj.oddsOrder.forEach((jp, offset) => {
    const row = 2 + offset;
    setValue(ws, `A${row}`, jp.label);
    setValue(ws, `B${row}`, currencyStar(jp.prize));
    setValue(ws, `C${row}`, `='Progressive Jackpots'!$H$${jp.row}`);
  });
  const firstRow = 4;
  prizes.forEach((prize, offset) => {
    const row = firstRow + offset;
    if (offset === 0) setValue(ws, `A${row}`, "Prize");
    setValue(ws, `B${row}`, prize);
    setValue(ws, `C${row}`, `='Delivery'!$M$6/SUMIF('Delivery'!$E$4:$E$${lastWin},B${row},'Delivery'!$F$4:$F$${lastWin})`);
  });
  const nonwinRow = firstRow + prizes.length;
  setValue(ws, `B${nonwinRow}`, "NON WINNING");
  setValue(ws, `C${nonwinRow}`, `='Delivery'!$M$6/'Delivery'!$F$${delivery.nonwinRow}`);
  reflowOddsJp(ws, templateWs, 2, firstRow, nonwinRow, 3);
  for (let row = 1; row <= nonwinRow; row += 1) {
    ws.getCell(row, 2).numFmt = "#,##0.00";
    ws.getCell(row, 3).numFmt = "#,##0.00";
  }
  return { nonwinRow, uniquePrizes: prizes.length };
}

function buildOddsMmj3(ws, templateWs, rows, delivery, ctx) {
  const prizes = groupPrizes(rows, "prize");
  const grid = ctx.priceGrid;
  copySheetChrome(templateWs, ws);
  copyColumnWidths(templateWs, ws, 17);
  setValue(ws, "A1", "Ticket Price");
  grid.forEach((price, idx) => {
    const priceCol = 2 + idx * 2;
    setValue(ws, { row: 1, col: priceCol }, price);
    setValue(ws, { row: 1, col: priceCol + 1 }, "1 in X Odds ");
  });
  ctx.jpEntries.forEach((jp, offset) => {
    const row = 2 + offset;
    setValue(ws, `A${row}`, jp.label);
    setValue(ws, `B${row}`, currencyStar(jp.prize));
    setValue(ws, `D${row}`, currencyStar(jp.prize));
    setValue(ws, `E${row}`, `='Progressive Jackpots'!$H$${jp.row}`);
    setValue(ws, `C${row}`, `=E${row}*($D$1/$B$1)`);
    for (const priceCol of [6, 8, 10, 12, 14, 16]) {
      const p = colLetter(priceCol);
      const o = colLetter(priceCol + 1);
      setValue(ws, `${p}${row}`, cellResult(ws.getCell(`D${row}`)) ?? ws.getCell(`D${row}`).value);
      setValue(ws, `${o}${row}`, `=E${row}/${p}$1`);
    }
  });
  const firstRow = 2 + ctx.jpEntries.length;
  const lastWin = delivery.lastWinRow;
  prizes.forEach((prize, offset) => {
    const row = firstRow + offset;
    if (offset === 0) setValue(ws, `A${row}`, "Prize");
    setValue(ws, `D${row}`, prize);
    setValue(ws, `E${row}`, `='Delivery'!$M$6/SUMIF('Delivery'!$D$4:$D$${lastWin},D${row},'Delivery'!$F$4:$F$${lastWin})`);
    setValue(ws, `B${row}`, `=D${row}*$B$1`);
    setValue(ws, `C${row}`, `=E${row}`);
    for (const priceCol of [6, 8, 10, 12, 14, 16]) {
      const p = colLetter(priceCol);
      const o = colLetter(priceCol + 1);
      setValue(ws, `${p}${row}`, `=D${row}*${p}$1`);
      setValue(ws, `${o}${row}`, `=E${row}`);
    }
  });
  const nonwinRow = firstRow + prizes.length;
  for (const priceCol of [2, 4, 6, 8, 10, 12, 14, 16]) {
    const p = colLetter(priceCol);
    const o = colLetter(priceCol + 1);
    setValue(ws, `${p}${nonwinRow}`, "NON WINNING");
    setValue(ws, `${o}${nonwinRow}`, priceCol === 4 ? `='Delivery'!$M$6/'Delivery'!$F$${delivery.nonwinRow}` : `=E${nonwinRow}`);
  }
  reflowOddsJp(ws, templateWs, ctx.jpEntries.length, firstRow, nonwinRow, 17);
  for (let row = 1; row <= nonwinRow; row += 1) {
    for (let col = 2; col <= 17; col += 1) ws.getCell(row, col).numFmt = "#,##0.00";
  }
  return { nonwinRow, uniquePrizes: prizes.length };
}

function buildOddsNoJp(ws, templateWs, rows, delivery, ctx) {
  const prizes = groupPrizes(rows, "prize");
  const grid = ctx.priceGrid;
  const baseCol = priceGridBaseCol(grid);
  const baseLetter = colLetter(baseCol);
  const oddsCol = 2 + grid.length;
  const oddsLetter = colLetter(oddsCol);
  const lastWin = delivery.lastWinRow;
  copySheetChrome(templateWs, ws);
  copyColumnWidths(templateWs, ws, oddsCol);
  setValue(ws, "A1", "Ticket Price");
  grid.forEach((price, idx) => setValue(ws, { row: 1, col: idx + 2 }, price));
  setValue(ws, { row: 1, col: oddsCol }, "1 in X Odds ");
  prizes.forEach((prize, offset) => {
    const row = 2 + offset;
    if (offset === 0) setValue(ws, `A${row}`, "Prize");
    setValue(ws, { row, col: baseCol }, prize);
    grid.forEach((_, idx) => {
      const col = idx + 2;
      if (col === baseCol) return;
      setValue(ws, { row, col }, `=${baseLetter}${row}*${colLetter(col)}$1`);
    });
    setValue(ws, `${oddsLetter}${row}`, `='Delivery'!$M$6/SUMIF('Delivery'!$D$4:$D$${lastWin},${baseLetter}${row},'Delivery'!$F$4:$F$${lastWin})`);
  });
  const nonwinRow = 2 + prizes.length;
  grid.forEach((_, idx) => setValue(ws, { row: nonwinRow, col: idx + 2 }, "NON WINNING"));
  setValue(ws, `${oddsLetter}${nonwinRow}`, `='Delivery'!$M$6/'Delivery'!$F$${delivery.nonwinRow}`);
  copyRowStyle(templateWs, ws, 1, 1, 1, oddsCol);
  for (let row = 2; row < nonwinRow; row += 1) {
    copyRowStyle(templateWs, ws, row === 2 ? 2 : 3, row, 1, oddsCol);
  }
  copyRowStyle(templateWs, ws, findOddsTemplateNonwin(templateWs), nonwinRow, 1, oddsCol);
  return { nonwinRow, uniquePrizes: prizes.length };
}

function buildSummarySsj(ws, templateWs, ctx, delivery) {
  copySheetChrome(templateWs, ws);
  copyColumnWidths(templateWs, ws, 2);
  const maxRow = Math.min(templateWs.rowCount || 28, 28);
  for (let row = 1; row <= maxRow; row += 1) {
    copyRowStyle(templateWs, ws, row, row, 1, 2);
    setValue(ws, `A${row}`, cellResult(templateWs.getCell(row, 1)));
  }
  const dist = summaryDistribution(ctx.summary);
  const jp1 = ctx.ssj.summaryJp1;
  const jp2 = ctx.ssj.summaryJp2;
  setValue(ws, "B1", ctx.base);
  setValue(ws, "B2", ctx.topPrize);
  setValue(ws, "B3", "v1");
  setValue(ws, "B4", ctx.rtp);
  setValue(ws, "B5", dist.breakeven);
  setValue(ws, "B6", "='Delivery'!$M$15");
  setValue(ws, "B7", "='Delivery'!$M$6");
  setValue(ws, "B8", "x0 - x5");
  setValue(ws, "B9", dist.low);
  setValue(ws, "B10", "x5 - x20");
  setValue(ws, "B11", dist.medium);
  setValue(ws, "B12", "x20 - top");
  setValue(ws, "B13", dist.high);
  setValue(ws, "B14", "=B7*B4*B1");
  setValue(ws, "B15", ctx.base);
  setValue(ws, "B16", cleanFloat(cellResult(ctx.pj.getCell("C2"))));
  setValue(ws, "B17", jp2.wins);
  setValue(ws, "B18", jp1.wins);
  setValue(ws, "B19", 1);
  setValue(ws, "B20", 40);
  setValue(ws, "B21", `='Progressive Jackpots'!$H$${jp1.row}`);
  setValue(ws, "B22", `='Progressive Jackpots'!$H$${jp2.row}`);
  setValue(ws, "B23", "=B16*B15");
  setValue(ws, "B24", "='Progressive Jackpots'!$C$14");
  setValue(ws, "B25", jp1.prize);
  setValue(ws, "B26", jp2.prize);
  setValue(ws, "B27", "='Progressive Jackpots'!$C$21");
  setValue(ws, "B28", "=B27*B17");
  ws.getCell("B4").numFmt = "0.00%";
  ws.getCell("B5").numFmt = "0.00%";
  ws.getCell("B9").numFmt = "0.00%";
  ws.getCell("B11").numFmt = "0.00%";
  ws.getCell("B13").numFmt = "0.00%";
  return dist;
}

function buildSummaryGrid(ws, templateWs, ctx) {
  const grid = ctx.priceGrid;
  const hasJp = ctx.schema === "mmj3";
  const dist = summaryDistribution(ctx.summary);
  const baseCol = priceGridBaseCol(grid);
  const baseLetter = colLetter(baseCol);
  const oddsCell = hasJp ? "$M$15" : "$M$13";
  copySheetChrome(templateWs, ws);
  copyColumnWidths(templateWs, ws, grid.length + 1);
  const labels = [
    "Size of Grid",
    "Top Prize",
    "Version",
    "Payout",
    "Breakeven",
    "Odds",
    "Size of Grid",
    "Low Tier Winning Range",
    "Percentage in Low Tier",
    "Medium Tier Winning Range",
    "Percentage in Medium Tier",
    "High Tier Winning Range",
    "Percentage in High Tier",
    "Total Pool Amount",
  ];
  labels.forEach((label, i) => {
    copyRowStyle(templateWs, ws, i + 1, i + 1, 1, grid.length + 1);
    setValue(ws, `A${i + 1}`, label);
  });
  grid.forEach((price, idx) => {
    const col = colLetter(idx + 2);
    const isBase = idx + 2 === baseCol;
    setValue(ws, { row: 1, col: idx + 2 }, price);
    setValue(ws, `${col}2`, isBase ? ctx.topPrize : `=${baseLetter}$2*${col}$1`);
    setValue(ws, `${col}3`, "v1");
    setValue(ws, `${col}4`, isBase ? ctx.rtp : `=${baseLetter}$4`);
    setValue(ws, `${col}5`, isBase ? dist.breakeven : `=${baseLetter}$5`);
    setValue(ws, `${col}6`, isBase ? `='Delivery'!${oddsCell}` : `=${baseLetter}$6`);
    setValue(ws, `${col}7`, "='Delivery'!$M$6");
    setValue(ws, `${col}8`, "x0 - x5");
    setValue(ws, `${col}9`, isBase ? dist.low : `=${baseLetter}$9`);
    setValue(ws, `${col}10`, "x5 - x20");
    setValue(ws, `${col}11`, isBase ? dist.medium : `=${baseLetter}$11`);
    setValue(ws, `${col}12`, "x20 - top");
    setValue(ws, `${col}13`, isBase ? dist.high : `=${baseLetter}$13`);
    setValue(ws, `${col}14`, `=${col}7*${col}4*${col}1`);
  });
  if (!hasJp) return dist;

  setValue(ws, "A15", "Progressive Jackpot Info");
  grid.forEach((price, idx) => setValue(ws, { row: 15, col: idx + 2 }, price));
  const jp1 = ctx.jpEntries[0];
  const jp2 = ctx.jpEntries[1];
  if (!jp2) {
    const jpLabels = [
      "Size of Grid",
      "JP1 Wins/Grid",
      "JP1 Wins/Price Ratio",
      "Odds of JP1 Win",
      "Revenue/JP win",
      "Jackpot Reset %",
      "JP1 Reset Value",
      "Jackpot Increment %",
      "Jackpot Increment Amount",
    ];
    jpLabels.forEach((label, i) => setValue(ws, `A${16 + i}`, label));
    grid.forEach((_, idx) => {
      const col = colLetter(idx + 2);
      const isBase = idx + 2 === baseCol;
      setValue(ws, `${col}16`, cleanFloat(cellResult(ctx.pj.getCell("C2"))));
      setValue(ws, `${col}17`, isBase ? jp1.wins : `=${baseLetter}$17*${col}$15`);
      setValue(ws, `${col}18`, isBase ? 1 : `=${baseLetter}$18`);
      setValue(ws, `${col}19`, isBase ? jp1.odds : `=${baseLetter}$19/${col}$15`);
      setValue(ws, `${col}20`, `=${col}16*${col}15`);
      setValue(ws, `${col}21`, isBase ? "='Progressive Jackpots'!$C$14" : `=${baseLetter}$21`);
      setValue(ws, `${col}22`, isBase ? jp1.prize : `=${baseLetter}$22`);
      setValue(ws, `${col}23`, isBase ? "='Progressive Jackpots'!$C$21" : `=${baseLetter}$23`);
      setValue(ws, `${col}24`, `=${col}23*${col}15`);
    });
    return dist;
  }
  const jpLabels = [
    "Size of Grid",
    "JP1 Wins/Grid",
    "JP2 Wins/Grid",
    "JP1 Wins/Price Ratio",
    "JP2 Wins/Price Ratio",
    "Odds of JP1 Win",
    "Odds of JP2 Win",
    "Revenue/JP win",
    "Jackpot Reset %",
    "JP1 Reset Value",
    "JP2 Reset Value",
    "Jackpot Increment %",
    "Jackpot Increment Amount",
  ];
  jpLabels.forEach((label, i) => setValue(ws, `A${16 + i}`, label));
  grid.forEach((_, idx) => {
    const col = colLetter(idx + 2);
    const isBase = idx + 2 === baseCol;
    setValue(ws, `${col}16`, cleanFloat(cellResult(ctx.pj.getCell("C2"))));
    setValue(ws, `${col}17`, isBase ? jp1.wins : `=${baseLetter}$17*${col}$15`);
    setValue(ws, `${col}18`, isBase ? jp2.wins : `=${baseLetter}$18*${col}$15`);
    setValue(ws, `${col}19`, isBase ? 1 : `=${baseLetter}$19`);
    setValue(ws, `${col}20`, isBase ? 1 : `=${baseLetter}$20`);
    setValue(ws, `${col}21`, isBase ? jp1.odds : `=${baseLetter}$21/${col}$15`);
    setValue(ws, `${col}22`, isBase ? jp2.odds : `=${baseLetter}$22/${col}$15`);
    setValue(ws, `${col}23`, `=${col}16*${col}15`);
    setValue(ws, `${col}24`, isBase ? "='Progressive Jackpots'!$C$14" : `=${baseLetter}$24`);
    setValue(ws, `${col}25`, isBase ? jp1.prize : `=${baseLetter}$25`);
    setValue(ws, `${col}26`, isBase ? jp2.prize : `=${baseLetter}$26`);
    setValue(ws, `${col}27`, isBase ? "='Progressive Jackpots'!$C$21" : `=${baseLetter}$27`);
    setValue(ws, `${col}28`, `=${col}27*${col}15`);
  });
  return dist;
}

function payTableMatrix(pay) {
  const quantityCols = [];
  for (let col = 4; col <= (pay.columnCount || 20); col += 1) {
    const v = cellResult(pay.getCell(3, col));
    if (isNumeric(asNumber(v, NaN))) quantityCols.push(col);
  }
  const quantities = quantityCols.map((col) => Number(cellResult(pay.getCell(3, col))));
  const symbolCodes = {};
  const matrix = {};
  const max = pay.rowCount || 20;
  for (let row = 4; row <= max; row += 1) {
    const label = String(cellResult(pay.getCell(row, 2)) || "").trim();
    if (label.toLowerCase() === "bonus") break;
    const symbol = cellResult(pay.getCell(row, 2));
    const code = cellResult(pay.getCell(row, 3));
    if (symbol == null || symbol === "" || !isNumeric(asNumber(code, NaN))) continue;
    const hasQty = quantityCols.some((col) => isNumeric(asNumber(cellResult(pay.getCell(row, col)), NaN)));
    if (!hasQty) continue;
    const codeNum = Number(code);
    symbolCodes[codeNum] = String(symbol);
    for (const col of quantityCols) {
      const qty = Number(cellResult(pay.getCell(3, col)));
      matrix[`${codeNum}:${qty}`] = cleanFloat(cellResult(pay.getCell(row, col)));
    }
  }
  return { symbolCodes, quantities, matrix };
}

function symbolLabel(value, symbolCodes) {
  if (value == null || value === "") return "";
  if (isNumeric(value)) return String(symbolCodes[value] ?? value);
  return String(value)
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const n = Number(part);
      return Number.isFinite(n) ? String(symbolCodes[n] ?? part) : part;
    })
    .join("/");
}

function buildBonusLabels(pay, wm) {
  const labels = {};
  const maxPay = pay?.rowCount || 0;
  for (let row = 1; row <= maxPay; row += 1) {
    const method = cellResult(pay.getCell(row, 2));
    if (!method) continue;
    const prefix = methodPrefix(method);
    const label = cellResult(pay.getCell(row, 5));
    if (label && prefix.startsWith("bonus") && !labels[prefix]) labels[prefix] = label;
  }
  const maxWm = wm?.rowCount || 0;
  for (let row = 5; row <= maxWm; row += 1) {
    const method = cellResult(wm.getCell(row, 3));
    if (!method) continue;
    const prefix = methodPrefix(method);
    const note = cellResult(wm.getCell(row, 8));
    if (prefix.startsWith("bonus") && note && !labels[prefix]) labels[prefix] = note;
  }
  return labels;
}

function buildPrizeBreakdown(ws, templateWs, ctx) {
  const pay = ctx.pay;
  const wm = ctx.winMethods;
  const free = ctx.freePlays;
  copySheetChrome(templateWs, ws);
  copyColumnWidths(templateWs, ws, 16);
  copyRowStyle(templateWs, ws, 2, 2, 2, 6);
  setValue(ws, "B2", "Tier");
  setValue(ws, "C2", "Win Methods");
  setValue(ws, "D2", "Prize");
  setValue(ws, "E2", "Symbol");
  setValue(ws, "F2", "Quantity of Symbols");
  const bonusLabels = buildBonusLabels(pay, wm);
  if (!bonusLabels.bonusa && pay) {
    bonusLabels.bonusa = String(cellResult(pay.getCell("B13")) || "Bonus");
  }
  const { symbolCodes, quantities, matrix } = payTableMatrix(pay);
  const freeplayLookup = {};
  if (free) {
    const max = free.rowCount || 0;
    for (let row = 3; row <= max; row += 1) {
      const method = cellResult(free.getCell(row, 4));
      const rounds = cellResult(free.getCell(row, 3));
      if (method) {
        freeplayLookup[String(method)] = rounds;
        freeplayLookup[String(method).toLowerCase()] = rounds;
      }
    }
  }
  let outRow = 3;
  const maxWm = wm.rowCount || 0;
  for (let winRow = 5; winRow <= maxWm; winRow += 1) {
    const methodRaw = cellResult(wm.getCell(winRow, 3));
    const prize = cellResult(wm.getCell(winRow, 6));
    if (!methodRaw || methodRaw === "0" || prize == null || prize === "") continue;
    if (!isNumeric(asNumber(prize, NaN))) continue;
    const method = String(methodRaw);
    const prefix = methodPrefix(method);
    const noteValue = cellResult(wm.getCell(winRow, 8)) || "";
    let symbol;
    let qty;
    if (prefix === "main" || prefix.startsWith("wild")) {
      symbol = symbolLabel(cellResult(wm.getCell(winRow, 4)), symbolCodes);
      qty = cellResult(wm.getCell(winRow, 5));
    } else if (prefix.startsWith("instant")) {
      symbol = "Instant Win";
      qty = 1;
    } else if (prefix.startsWith("bonus")) {
      symbol = bonusLabels[prefix] || "Bonus";
      qty = isNumeric(asNumber(cellResult(wm.getCell(winRow, 5)), NaN))
        ? cellResult(wm.getCell(winRow, 5))
        : 1;
    } else if (prefix.startsWith("freeplay")) {
      symbol = "Free Plays";
      qty = freeplayLookup[method] ?? freeplayLookup[method.toLowerCase()] ?? 10;
    } else {
      symbol = prefix;
      qty = isNumeric(asNumber(cellResult(wm.getCell(winRow, 5)), NaN))
        ? cellResult(wm.getCell(winRow, 5))
        : 1;
    }
    copyRowStyle(templateWs, ws, 3, outRow, 2, 6);
    setValue(ws, `B${outRow}`, cellResult(wm.getCell(winRow, 2)));
    setValue(ws, `C${outRow}`, method);
    applyMethodFill(ws.getCell(`C${outRow}`), method, ctx.fills);
    setValue(ws, `D${outRow}`, cleanFloat(prize)).numFmt = "#,##0.00";
    setValue(ws, `E${outRow}`, symbol);
    setValue(ws, `F${outRow}`, qty);
    if (noteValue && String(noteValue) !== String(symbol) && prefix !== "instant") {
      setValue(ws, `G${outRow}`, noteValue);
    }
    outRow += 1;
  }
  setValue(ws, "H8", "Quantity of Symbols");
  const codes = Object.keys(symbolCodes)
    .map(Number)
    .sort((a, b) => a - b);
  codes.forEach((code, idx) => setValue(ws, { row: 8, col: 9 + idx }, symbolCodes[code]));
  quantities.forEach((qty, rIdx) => {
    setValue(ws, { row: 9 + rIdx, col: 8 }, qty);
    codes.forEach((code, cIdx) => {
      setValue(ws, { row: 9 + rIdx, col: 9 + cIdx }, matrix[`${code}:${qty}`]);
    });
  });
  const matrixLastRow = 8 + quantities.length;
  const matrixLastCol = 8 + codes.length;
  for (let row = 8; row <= matrixLastRow; row += 1) {
    for (let col = 8; col <= matrixLastCol; col += 1) {
      copyStyle(templateWs.getCell(row, col), ws.getCell(row, col));
    }
  }
  return outRow - 1;
}

async function loadWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

export async function inspectPps(buffer, filename = "upload.xlsx") {
  const wb = await loadWorkbook(buffer);
  const freq = requireSheet(wb, "Frequency");
  const identity = identityText(freq);
  const pj = sheet(wb, "Progressive Jackpots");
  const jpEntries = progressiveJackpots(pj);
  const schema = detectSchema({ identity, filename, jpEntries, hasPjSheet: Boolean(pj) });
  const quantity = freqNumber(freq, 6);
  const base = freqNumber(freq, 7);
  const rtp = freqNumber(freq, 10);
  if (!Number.isFinite(quantity) || !Number.isFinite(base) || quantity <= 0 || base <= 0) {
    throw new Error("Frequency N6 pool / N7 ticket price is missing or zero.");
  }
  const rows = sourceRows(freq);
  const missingCache = rows.length && rows.some((r) => !r.method);
  const wins = rows.reduce((s, r) => s + r.winners, 0);
  const fund = rows.reduce((s, r) => s + r.prize * r.winners, 0);
  const actualRtp = schema === "ssj" ? fund / (quantity * base) : fund / quantity;
  return {
    filename,
    identity,
    kentucky: isKentucky(identity, filename),
    schema,
    sheets: wb.worksheets.map((ws) => ws.name),
    quantity,
    base,
    rtp: Number.isFinite(rtp) ? rtp : null,
    winningTiers: rows.length,
    uniquePrizes: groupPrizes(rows, "prize").length,
    duplicateMethods: rows.duplicateMethods || [],
    wins,
    hitRate: quantity / wins,
    prizeFund: fund,
    actualRtp,
    jpCount: jpEntries.length,
    jpNames: jpEntries.map((e) => e.name || e.label),
    existingDelivery: DELIVERY_SHEETS.filter((n) => sheet(wb, n)),
    missingMethodCache: Boolean(missingCache),
  };
}

export async function buildPps(buffer, filename, options = {}) {
  if (!ExcelJS) throw new Error("ExcelJS is not loaded.");
  const templates = options.templates || {};
  const wb = await loadWorkbook(buffer);
  const freq = requireSheet(wb, "Frequency");
  const identity = identityText(freq);
  const kentucky = isKentucky(identity, filename);
  if (!kentucky && !options.allowNonKentucky) {
    throw new Error(
      "Workbook labels / filename do not establish Kentucky identity. Tick 'This is a Kentucky PPS' only if that is intentional.",
    );
  }
  const pj = sheet(wb, "Progressive Jackpots");
  const jpEntries = progressiveJackpots(pj);
  let schema = options.schema && options.schema !== "auto"
    ? options.schema
    : detectSchema({ identity, filename, jpEntries, hasPjSheet: Boolean(pj) });
  const quantity = freqNumber(freq, 6);
  const base = freqNumber(freq, 7);
  const rtp = freqNumber(freq, 10);
  if (!Number.isFinite(quantity) || !Number.isFinite(base) || quantity <= 0 || base <= 0) {
    throw new Error("Frequency N6 pool / N7 ticket price is missing or zero.");
  }
  if (!Number.isFinite(rtp)) throw new Error("Frequency N10 RTP setting is missing.");
  const rows = sourceRows(freq);
  const summary = sheet(wb, "Summary");
  const winMethods = requireSheet(wb, "Win Methods");
  const pay = requireSheet(wb, "Pay Table");
  const freePlays = sheet(wb, "Free Plays");
  const fills = winMethodFills(winMethods);
  let ssj = null;
  if (schema === "ssj") ssj = ssjJackpots(jpEntries);

  const priceGrid = PRICE_GRIDS[schema] || null;
  const templateName = schema === "ssj" ? "ssj" : schema === "mmj3" ? "mmj3" : "no-jp";
  const templateBuffer = templates[templateName] || templates.mmj3 || templates["no-jp"];
  if (!templateBuffer) throw new Error(`Template not loaded: ${templateName}`);
  const templateWb = await loadWorkbook(templateBuffer);

  const deliveryTemplate = requireSheet(templateWb, "Delivery");
  const oddsTemplate = requireSheet(templateWb, "Odds Table");
  const summaryTemplate = requireSheet(templateWb, "Summary(Delivery)");
  const prizeTemplate = requireSheet(templateWb, "Prize Breakdown");

  const topPrize = cleanFloat(
    summaryLookup(summary, ["max prize"], cellResult(summary?.getCell("C6")) || cellResult(summary?.getCell("C7"))),
  );

  const ctx = {
    schema,
    filename,
    identity,
    quantity,
    base,
    rtp,
    topPrize,
    priceGrid,
    jpEntries,
    ssj,
    freq,
    pj,
    summary,
    winMethods,
    pay,
    freePlays,
    fills,
  };

  const deliveryWs = insertSheet(wb, "Delivery");
  const oddsWs = insertSheet(wb, "Odds Table");
  const summaryWs = insertSheet(wb, "Summary(Delivery)");
  const prizeWs = insertSheet(wb, "Prize Breakdown");

  const delivery = buildDelivery(deliveryWs, deliveryTemplate, wb, rows, ctx);
  let odds;
  if (schema === "ssj") odds = buildOddsSsj(oddsWs, oddsTemplate, rows, delivery, ctx);
  else if (schema === "mmj3") odds = buildOddsMmj3(oddsWs, oddsTemplate, rows, delivery, ctx);
  else odds = buildOddsNoJp(oddsWs, oddsTemplate, rows, delivery, ctx);
  const dist =
    schema === "ssj"
      ? buildSummarySsj(summaryWs, summaryTemplate, ctx, delivery)
      : buildSummaryGrid(summaryWs, summaryTemplate, ctx);
  const prizeLast = buildPrizeBreakdown(prizeWs, prizeTemplate, ctx);
  orderSheets(wb);

  wb.calcProperties = wb.calcProperties || {};
  wb.calcProperties.fullCalcOnLoad = true;

  const wins = rows.reduce((s, r) => s + r.winners, 0);
  const fund = rows.reduce((s, r) => s + r.prize * r.winners, 0);
  const outBuffer = await wb.xlsx.writeBuffer();
  const report = {
    filename,
    identity,
    kentucky,
    schema,
    quantity,
    base,
    rtp,
    winningTiers: rows.length,
    uniquePrizes: odds.uniquePrizes,
    duplicateMethods: rows.duplicateMethods || [],
    wins,
    hitRate: quantity / wins,
    prizeFund: fund,
    actualRtp: schema === "ssj" ? fund / (quantity * base) : fund / (quantity * (schema === "mmj3" ? 1 : base)),
    dist,
    delivery,
    oddsLast: odds.nonwinRow,
    prizeLast,
    sheets: wb.worksheets.map((ws) => ws.name),
  };
  return { buffer: outBuffer, report };
}

export function outputFilename(name) {
  const base = String(name || "pps.xlsx").replace(/\.xlsx$/i, "");
  if (/_delivery$/i.test(base)) return `${base}.xlsx`;
  return `${base}_Delivery.xlsx`;
}
