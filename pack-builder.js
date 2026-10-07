/**
 * NC / GA / PA / NH / VA Delivery builder (browser + Node).
 * Ports .cursor/skills/pps-delivery-conversion. One Delivery sheet, not KY four-tab.
 * PA and NH use the Georgia layout. VA is the NC layout with the Frequency summary
 * above the tier table instead of beside it. Number formats are plain: no _ or * padding.
 */

const ExcelJS = globalThis.ExcelJS;

export const DELIVERY_TAB_COLOR = { theme: 5, tint: 0.7999816888943144 };

export const PACK_LOTTERIES = {
  NC: { code: "NC", layout: "NC", name: "NC Education Lottery" },
  GA: { code: "GA", layout: "GA", name: "Georgia Lottery" },
  PA: { code: "PA", layout: "GA", name: "Pennsylvania Lottery" },
  NH: { code: "NH", layout: "GA", name: "New Hampshire Lottery" },
  VA: { code: "VA", layout: "VA", name: "Virginia Lottery" },
  DC: { code: "DC", layout: "DC", name: "DC Lottery" },
};

export const DC_PRICE_OPTIONS = [0.1, 0.2, 0.5, 1, 2, 3, 5, 10, 20, 30, 50];
export const DC_DEFAULT_PRICE_POINTS = [0.5, 1, 2, 3, 5, 10, 20, 30, 50];

export const PACK_SHEET = "Delivery";

const FMT_INT = "#,##0";
const FMT_MONEY = "#,##0.00";
const FMT_PCT = "0.00%";
const FMT_HIT = "0.00";
const FMT_DATE_GA = "[$-409]mmmm d, yyyy";

const CONFIDENTIAL =
  "The information contained in this document and all attached documents is strictly confidential " +
  "and contains proprietary information.\n" +
  "It is provided solely for use by the designated recipients and is subject to the terms of any " +
  "confidentiality obligations or non-disclosure agreements between the parties.\n" +
  "All other use is strictly prohibited.";

const NC_WIDTHS = {
  2: 18.44,
  3: 68.66,
  4: 13.33,
  5: 12.66,
  6: 15.33,
  7: 14.78,
  8: 21.44,
  9: 21.11,
  10: 14.78,
  11: 32.78,
  12: 17.55,
  13: 12.55,
  14: 21.11,
};

const GA_WIDTHS = {
  1: 30.66,
  2: 15,
  3: 13,
  4: 13,
  5: 16,
  6: 44.78,
  7: 15,
  8: 15,
  9: 13,
  10: 20.89,
  11: 15,
  12: 13,
  13: 13,
  14: 13,
  15: 13,
  16: 13,
};

const DC_WIDTHS = {
  2: 17.11,
  3: 55.33,
  4: 13.33,
  5: 12.66,
  6: 15.33,
  7: 17.66,
  8: 14.78,
  9: 12,
  10: 12,
  11: 37.11,
  12: 23.44,
  13: 17.55,
  14: 12.55,
  15: 21.11,
  16: 16.66,
  17: 18.33,
  18: 12.33,
  19: 18.78,
  20: 16.11,
  21: 19,
  22: 23.44,
};

const VA_WIDTHS = {
  2: 25.89,
  3: 55.33,
  4: 13.33,
  5: 12.66,
  6: 15.33,
  7: 17.66,
  8: 14.78,
  9: 12,
};

const MEDIUM_BOTTOM = {
  bottom: { style: "medium", color: { indexed: 64 } },
};

const TINT_80 = 0.7999816888943144;
const FONT_TITLE = { name: "Arial Black", size: 10, bold: true };
const FONT_GROUP = { name: "Geneva", size: 10 };
const FONT_HELV8 = { name: "Helv", size: 8 };
const FONT_HELV8B = { name: "Helv", size: 8, bold: true };
const FONT_HELV10 = { name: "Helv", size: 10 };
const FONT_HELV10B = { name: "Helv", size: 10, bold: true };
const CENTER = { horizontal: "center" };

function themeFill(theme) {
  return { type: "pattern", pattern: "solid", fgColor: { theme, tint: TINT_80 } };
}

export function isPackLottery(id) {
  return Boolean(PACK_LOTTERIES[String(id || "").toUpperCase()]);
}

function cellResult(cell) {
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

function sheet(wb, name) {
  return wb.getWorksheet(name) || wb.worksheets.find((ws) => ws.name.toLowerCase() === name.toLowerCase());
}

function requireSheet(wb, name) {
  const ws = sheet(wb, name);
  if (!ws) throw new Error(`Missing sheet: ${name}`);
  return ws;
}

async function loadWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
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

function cloneJson(obj) {
  if (obj == null) return obj;
  try {
    return JSON.parse(JSON.stringify(obj));
  } catch {
    return null;
  }
}

function put(ws, addr, value, style = {}) {
  const cell = typeof addr === "string" ? ws.getCell(addr) : ws.getCell(addr.row, addr.col);
  cell.style = cloneJson(style) || {};
  if (value !== undefined) setValue(ws, cell.address, value);
  return cell;
}

function addBottomEdge(ws, row, fromCol, toCol) {
  for (let col = fromCol; col <= toCol; col += 1) {
    const cell = ws.getCell(row, col);
    const value = cell.value;
    const style = cloneJson(cell.style) || {};
    style.border = { ...(style.border || {}), ...cloneJson(MEDIUM_BOTTOM) };
    cell.style = style;
    if (value != null) cell.value = value;
  }
}

function hasFill(fill) {
  return Boolean(fill && fill.type === "pattern" && fill.pattern && String(fill.pattern).toLowerCase() !== "none");
}

function hasBorder(border) {
  return Boolean(border && ["left", "right", "top", "bottom"].some((edge) => border[edge]?.style));
}

function copyCellStyle(src, dst) {
  const style = {};
  if (src.font) style.font = cloneJson(src.font);
  if (hasFill(src.fill)) style.fill = cloneJson(src.fill);
  if (hasBorder(src.border)) style.border = cloneJson(src.border);
  if (src.alignment) style.alignment = cloneJson(src.alignment);
  if (src.numFmt && src.numFmt !== "General") style.numFmt = plainFormatCode(src.numFmt);
  dst.style = style;
}

function insertSheet(wb, name) {
  const existing = sheet(wb, name);
  if (existing) wb.removeWorksheet(existing.id);
  return wb.addWorksheet(name);
}

function labelText(freq, row) {
  return String(cellResult(freq.getCell(row, 13)) || "").trim();
}

function storedLabel(freq, row, fallback) {
  if (!row) return fallback;
  const text = cellResult(freq.getCell(row, 13));
  return text == null || text === "" ? fallback : String(text);
}

function findLabelRow(freq, test) {
  const max = Math.min(freq.rowCount || 40, 40);
  for (let row = 1; row <= max; row += 1) {
    const label = labelText(freq, row);
    if (label && test(label)) return row;
  }
  return null;
}

function valueAtLabel(freq, test, fallback = null) {
  const row = findLabelRow(freq, test);
  if (!row) return fallback;
  const value = cellResult(freq.getCell(row, 14));
  return value == null || value === "" ? fallback : value;
}

function formulaAtLabel(freq, test) {
  const row = findLabelRow(freq, test);
  if (!row) return null;
  return cellFormula(freq.getCell(row, 14));
}

function startsWithCi(text, prefix) {
  return String(text).toLowerCase().startsWith(String(prefix).toLowerCase());
}

export function excelSerialFromYymmdd(stamp) {
  const s = String(stamp || "");
  if (!/^\d{6}$/.test(s)) return null;
  const year = 2000 + Number(s.slice(0, 2));
  const month = Number(s.slice(2, 4));
  const day = Number(s.slice(4, 6));
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const utc = Date.UTC(year, month - 1, day);
  const epoch = Date.UTC(1899, 11, 30);
  return Math.round((utc - epoch) / 86400000);
}

export function packDateSerial(filename) {
  const m = String(filename || "").match(/^(\d{6})_/);
  return excelSerialFromYymmdd(m?.[1]) ?? excelSerialFromYymmdd("260102");
}

export function packOutputFilename(name, lottery = "NC") {
  const spec = PACK_LOTTERIES[String(lottery).toUpperCase()] || PACK_LOTTERIES.NC;
  const raw = String(name || "pps.xlsx").replace(/\.xlsx$/i, "");
  const m = raw.match(/^(\d{6})_([A-Za-z]{2,4})_(.+)_PPS_(.+)$/);
  if (m) {
    const game = m[3].replace(/\(Delivery\)/gi, "");
    return `${m[1]}_${spec.code}_${game}(Delivery)_PPS_${m[4]}.xlsx`;
  }
  if (/\(Delivery\)/i.test(raw)) return `${raw}.xlsx`;
  return `${raw}(Delivery).xlsx`;
}

export function sortGaWinning(rows) {
  return [...rows].sort((a, b) => a.prize - b.prize || a.sourceRow - b.sourceRow);
}

function captureFill(cell) {
  const fill = cell?.fill;
  if (!fill || fill.type == null) return null;
  const pattern = String(fill.pattern || "").toLowerCase();
  if (!pattern || pattern === "none") return null;
  const cloned = cloneJson(fill);
  if (!cloned) return null;
  const font = cell.font || {};
  return {
    fill: cloned,
    fontColor: cloneJson(font.color),
    bold: Boolean(font.bold),
  };
}

function applyMethodFill(cell, captured, baseFont) {
  cell.font = { ...baseFont };
  if (!captured?.fill) return;
  cell.fill = captured.fill;
  cell.font = {
    ...baseFont,
    bold: captured.bold || Boolean(baseFont.bold),
    ...(captured.fontColor ? { color: captured.fontColor } : {}),
  };
}

function setColWidths(ws, widths) {
  for (const [col, width] of Object.entries(widths)) {
    ws.getColumn(Number(col)).width = width;
  }
}

function applyTabColor(ws) {
  ws.properties = { ...ws.properties, tabColor: { ...DELIVERY_TAB_COLOR } };
}

function applyConfidential(ws, range, fontName, size = 8) {
  const start = String(range).split(":")[0];
  setValue(ws, start, CONFIDENTIAL);
  try {
    ws.mergeCells(range);
  } catch {
    /* already merged */
  }
  const cell = ws.getCell(start);
  cell.font = { name: fontName, size, bold: true };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function orderPackSheets(wb, layout) {
  const byName = Object.fromEntries(wb.worksheets.map((ws) => [ws.name, ws]));
  const names =
    layout === "VA" || layout === "DC" ? ["Frequency", PACK_SHEET] : ["Frequency", "Progressive Jackpots", PACK_SHEET];
  const head = names.map((n) => byName[n]).filter(Boolean);
  const skip = new Set(head);
  const tail = wb.worksheets.filter((ws) => !skip.has(ws));
  const list = [...head, ...tail];
  const sparse = [];
  list.forEach((ws, i) => {
    ws.id = i + 1;
    ws.orderNo = i;
    sparse[i + 1] = ws;
  });
  wb._worksheets = sparse;
}

function disableFullCalcOnLoad(wb) {
  wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: false };
}

export function frequencyWinningRows(freq) {
  const rows = [];
  const zeroFrequency = [];
  const max = Math.min(freq.rowCount || 0, 2500);
  const nonwin = {
    sourceRow: 4,
    tier: asNumber(cellResult(freq.getCell(4, 2)), 0),
    method: cellResult(freq.getCell(4, 3)),
    prize: asNumber(cellResult(freq.getCell(4, 4)), 0),
    winners: asNumber(cellResult(freq.getCell(4, 5)), 0),
    fill: captureFill(freq.getCell(4, 3)),
  };
  for (let row = 5; row <= max; row += 1) {
    const method = cellResult(freq.getCell(row, 3));
    const prize = cellResult(freq.getCell(row, 4));
    const winners = cellResult(freq.getCell(row, 5));
    if (method == null || method === "") continue;
    if (prize == null || prize === "") continue;
    const winNum = asNumber(winners, NaN);
    if (!Number.isFinite(winNum)) continue;
    const methodText = String(method);
    if (winNum === 0) {
      zeroFrequency.push(methodText);
      continue;
    }
    if (winNum < 0) {
      throw new Error(`Frequency has negative Odds up at row ${row} (${methodText}).`);
    }
    const prizeNum = asNumber(prize, NaN);
    if (!Number.isFinite(prizeNum)) continue;
    rows.push({
      sourceRow: row,
      tier: asNumber(cellResult(freq.getCell(row, 2)), rows.length + 1),
      method: methodText,
      prize: cleanFloat(prizeNum),
      winners: cleanFloat(winNum),
      fill: captureFill(freq.getCell(row, 3)),
    });
  }
  if (!rows.length) throw new Error("No winning Frequency rows found.");
  return { nonwin, rows, zeroFrequency };
}

function lookupMeta(freq, lottery) {
  const spec = PACK_LOTTERIES[lottery];
  const jurisdiction = cellResult(freq.getCell("M3")) || spec.name;
  const title = cellResult(freq.getCell("M4")) || "";
  const pool = asNumber(
    valueAtLabel(freq, (l) => startsWithCi(l, "odds down"), cellResult(freq.getCell("N6"))),
    NaN,
  );
  const base = asNumber(
    valueAtLabel(freq, (l) => /^base\s*:?/i.test(l), cellResult(freq.getCell("N7"))),
    1,
  );
  const rrpRow = findLabelRow(freq, (l) => startsWithCi(l, "rrp"));
  const rrp = rrpRow ? asNumber(cellResult(freq.getCell(rrpRow, 14)), NaN) : null;
  const buy = Number.isFinite(rrp) && rrp > 0;
  const rtp = asNumber(
    valueAtLabel(freq, (l) => startsWithCi(l, "rtp setting"), cellResult(freq.getCell("N10"))),
    NaN,
  );
  const hasJp = Boolean(findLabelRow(freq, (l) => startsWithCi(l, "jp rtp")));
  const hitRow =
    findLabelRow(freq, (l) => /^hit (rate|freq)\s*(:|\(main game\)|$)/i.test(l)) ||
    findLabelRow(freq, (l) => /hit (rate|freq)/i.test(l) && !/^hit rate\s*\(\+/i.test(l));
  const freqRows = {
    pool: findLabelRow(freq, (l) => startsWithCi(l, "odds down")) || 6,
    base: findLabelRow(freq, (l) => /^base\s*:?$/i.test(l)),
    rrp: rrpRow,
    revenue: findLabelRow(freq, (l) => startsWithCi(l, "revenu")),
    fund: findLabelRow(freq, (l) => startsWithCi(l, "prize fund")),
    rtp: findLabelRow(freq, (l) => startsWithCi(l, "rtp setting")) || 10,
    actual: findLabelRow(freq, (l) => /^(actual|base) rtp/i.test(l)),
    jp: findLabelRow(freq, (l) => startsWithCi(l, "jp rtp")),
    total: findLabelRow(freq, (l) => startsWithCi(l, "total rtp")),
    wins: findLabelRow(freq, (l) => startsWithCi(l, "winning tiers")) || 14,
    hit: hitRow || 15,
  };
  const hits = [];
  for (let row = 1; row <= 40; row += 1) {
    const label = labelText(freq, row);
    if (!/^hit rate\s*\(\+/i.test(label)) continue;
    const formula = cellFormula(freq.getCell(row, 14));
    if (formula) hits.push({ row, label: storedLabel(freq, row, label), formula });
  }
  return {
    spec,
    jurisdiction: String(jurisdiction || spec.name),
    title: String(title || ""),
    pool,
    base: Number.isFinite(base) && base > 0 ? base : 1,
    rrp: buy ? rrp : null,
    buy,
    rtp,
    hasJp,
    freqRows,
    hits,
    jpRtpFormula: freqRows.jp ? cellFormula(freq.getCell(freqRows.jp, 14)) : null,
  };
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

function jackpotFreqMap(layout, meta, rows) {
  const map = new Map();
  const { pool, rtp, wins, hit, actual } = meta.freqRows;
  if (layout === "VA") {
    map.set(`N${pool}`, "$C$6");
    map.set(`N${rtp}`, `$C$${rows.rtpSet}`);
    if (actual) map.set(`N${actual}`, `$C$${rows.actual}`);
    map.set(`N${wins}`, `$C$${rows.winFreq}`);
    map.set(`N${hit}`, `$C$${rows.hit}`);
  } else if (layout === "NC" || layout === "DC") {
    map.set(`N${pool}`, "$L$6");
    map.set(`N${rtp}`, `$L$${rows.rtpSet}`);
    if (actual) map.set(`N${actual}`, `$L$${rows.actual}`);
    map.set(`N${wins}`, `$L$${rows.winFreq}`);
    map.set(`N${hit}`, `$L$${rows.hit}`);
  } else {
    map.set(`N${pool}`, "$B$5");
    map.set(`N${rtp}`, "$B$9");
    map.set(`N${wins}`, `$B$${rows.winFreq}`);
    map.set(`N${hit}`, `$B$${rows.hit}`);
  }
  return map;
}

function rewriteJackpotFormula(formula, dRow, freqMap) {
  const localized = String(formula).replace(
    /'?Frequency'?!(\$?)([A-Z]{1,3})(\$?)(\d+)/gi,
    (match, _colAbs, col, _rowAbs, row) => freqMap.get(`${col.toUpperCase()}${row}`) || match,
  );
  return shiftFormula(localized.replace(/^=/, ""), 0, dRow);
}

function copyJackpotSetting(ws, pj, jpStart, freqMap, ownWidths = {}) {
  const dRow = jpStart - 2;
  const maxRow = Math.min(pj.rowCount || 32, 80);
  const maxCol = Math.min(Math.max(pj.columnCount || 24, 24), 30);
  for (let srcRow = 2; srcRow <= maxRow; srcRow += 1) {
    const dstRow = jpStart + (srcRow - 2);
    const height = pj.getRow(srcRow).height;
    if (height) ws.getRow(dstRow).height = height;
    for (let col = 2; col <= maxCol; col += 1) {
      if (srcRow <= 7 && col >= 5) continue;
      const src = pj.getCell(srcRow, col);
      const raw = src.value;
      const empty = raw == null || raw === "";
      if (empty && !hasBorder(src.border) && !hasFill(src.fill)) continue;
      const text = cellResult(src);
      if (typeof text === "string" && /json setup/i.test(text)) continue;
      const dst = ws.getCell(dstRow, col);
      copyCellStyle(src, dst);
      if (empty) continue;
      const formula = cellFormula(src);
      if (formula) dst.value = { formula: rewriteJackpotFormula(formula, dRow, freqMap) };
      else dst.value = text;
    }
  }
  for (let col = 2; col <= maxCol; col += 1) {
    if (ownWidths[col]) continue;
    const width = pj.getColumn(col).width;
    if (width) ws.getColumn(col).width = width;
  }
  return jpStart;
}

function hasJackpotSetting(pj) {
  if (!pj) return false;
  for (let row = 2; row <= 32; row += 1) {
    for (let col = 2; col <= 4; col += 1) {
      const v = pj.getCell(row, col).value;
      if (v != null && v !== "") return true;
    }
  }
  return false;
}

function resolveJackpot(wb, spec, meta) {
  if (spec.layout === "VA" && meta.hasJp) {
    meta.hasJp = hasJackpotSetting(sheet(wb, "Progressive Jackpots"));
  }
  return meta;
}

function rewritePjRefs(formula, dRow) {
  const text = String(formula || "").replace(/^=/, "");
  if (!/Progressive Jackpots'?!/i.test(text)) return "";
  return text.replace(/'?Progressive Jackpots'?!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/gi, (_, ref) =>
    ref.replace(/\$?([A-Z]{1,3})\$?(\d+)/g, (_m, col, row) => `${col}${Number(row) + dRow}`),
  );
}

function linkJackpotSummary(ws, meta, jpStart) {
  const sum = ncSummaryRows(meta.buy, true);
  const linked = rewritePjRefs(meta.jpRtpFormula, jpStart - 2);
  setValue(ws, `L${sum.jp}`, `=${linked || `C${jpStart + 28}+C${jpStart + 29}`}`);
  ws.getCell(`L${sum.jp}`).numFmt = FMT_PCT;
  ws.getCell(`L${sum.jp}`).font = { ...FONT_HELV10 };
}

const GA_INPUT_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { theme: 7, tint: 0.7999816888943144 },
};
const GA_STAKE_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC000" } };
const GA_SPLIT_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { theme: 9, tint: 0.7999816888943144 },
};
const GA_BASE_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { theme: 2, tint: -0.0999786370433668 },
};

function styleCell(cell, font, fmt, fill) {
  cell.font = font;
  if (fmt) cell.numFmt = fmt;
  if (fill) cell.fill = fill;
}

function readGaJackpotTiers(pj) {
  const tiers = [];
  for (let row = 11; row <= 13; row += 1) {
    const number = cellResult(pj.getCell(row, 2));
    const name = cellResult(pj.getCell(row, 3));
    if (number == null || number === "" || name == null || name === "") break;
    const oddsUp = asNumber(cellResult(pj.getCell(row, 13)), NaN);
    const seed = asNumber(cellResult(pj.getCell(row, 14)), NaN);
    const trigger = asNumber(cellResult(pj.getCell(row, 15)), NaN);
    if (![oddsUp, seed, trigger].every(Number.isFinite)) {
      throw new Error(
        `Progressive Jackpots row ${row} is missing cached Odds Up, Seed, or Target trigger. Open the PPS in Excel and save it.`,
      );
    }
    tiers.push({ number, name: String(name), oddsUp, seed, trigger });
  }
  if (!tiers.length) throw new Error("Progressive Jackpots has no jackpot tiers.");
  return tiers;
}

function buildGaJackpotSetting(ws, pj, jpStart) {
  const oddsDown = asNumber(cellResult(pj.getCell("C2")), NaN);
  const stake = asNumber(cellResult(pj.getCell("C3")), NaN);
  if (!Number.isFinite(oddsDown) || !Number.isFinite(stake)) {
    throw new Error("Progressive Jackpots Base JP Odds Down or JP Stake is missing.");
  }
  const tiers = readGaJackpotTiers(pj);
  const n = tiers.length;
  const cal = { name: "Calibri", size: 10 };
  const calB = { name: "Calibri", size: 10, bold: true };
  const helv = { name: "Helv", size: 8 };
  const helvB = { name: "Helv", size: 8, bold: true };

  const baseDown = jpStart;
  const stakeRow = jpStart + 1;
  const revenue = jpStart + 2;
  const jackpot = jpStart + 3;
  const jackpotPct = jpStart + 4;
  const baseFirst = jpStart + 2;
  const baseLast = baseFirst + n - 1;
  const baseTotal = baseLast + 1;
  const seedLabel = jackpotPct + 2;
  const seedHeader = seedLabel + 1;
  const seedFirst = seedHeader + 1;
  const seedLast = seedFirst + n - 1;
  const seedRtp = seedLast + 1;
  const contribLabel = seedRtp + 2;
  const contribHeader = contribLabel + 1;
  const contribFirst = contribHeader + 1;
  const contribLast = contribFirst + n - 1;
  const contribSum = contribLast + 1;
  const contribRtp = contribSum + 1;
  const triggerLabel = contribRtp + 2;
  const triggerHeader = triggerLabel + 1;
  const triggerFirst = triggerHeader + 1;

  const header = [
    [baseDown, "Base JP Odds Down", oddsDown, FMT_INT, GA_INPUT_FILL],
    [stakeRow, "Stake", stake, "0.00", GA_STAKE_FILL],
    [revenue, "Revenue", `=D${baseDown}*D${stakeRow}`, FMT_INT, null],
    [jackpot, "Jackpot", `=D${stakeRow}*D${jackpotPct}`, "0.00", null],
    [jackpotPct, "Jackpot %", `=D${seedRtp}+D${contribRtp}`, FMT_PCT, GA_INPUT_FILL],
  ];
  for (const [row, label, value, fmt, fill] of header) {
    styleCell(setValue(ws, `B${row}`, label), cal, null, null);
    styleCell(setValue(ws, `D${row}`, value), cal, fmt, fill);
  }

  styleCell(setValue(ws, `F${baseDown}`, "BASE"), helvB, null, null);
  for (const [col, label] of [
    ["F", "Odds Up"],
    ["G", "Odds Down"],
    ["H", "1 in x Odds"],
    ["I", "Prize Cost"],
  ]) {
    styleCell(setValue(ws, `${col}${stakeRow}`, label), helvB, null, GA_BASE_FILL);
  }
  tiers.forEach((tier, idx) => {
    const row = baseFirst + idx;
    const seedRow = seedFirst + idx;
    styleCell(setValue(ws, `F${row}`, tier.oddsUp), helv, FMT_MONEY, GA_INPUT_FILL);
    styleCell(setValue(ws, `G${row}`, `=D${baseDown}`), helv, FMT_INT, GA_INPUT_FILL);
    styleCell(setValue(ws, `H${row}`, `=G${row}/F${row}`), helv, FMT_INT, GA_INPUT_FILL);
    styleCell(setValue(ws, `I${row}`, `=F${row}*E${seedRow}`), helv, FMT_INT, GA_INPUT_FILL);
  });
  styleCell(setValue(ws, `F${baseTotal}`, `=SUM(F${baseFirst}:F${baseLast})`), helvB, FMT_MONEY, GA_INPUT_FILL);
  styleCell(setValue(ws, `I${baseTotal}`, `=SUM(I${baseFirst}:I${baseLast})`), helv, FMT_INT, GA_INPUT_FILL);

  styleCell(setValue(ws, `B${seedLabel}`, "Seed"), calB, null, null);
  styleCell(setValue(ws, `K${seedLabel}`, `=D${jackpotPct}-D${contribRtp}`), calB, FMT_PCT, null);
  const seedHeads = [
    ["B", "Number"],
    ["D", "Name"],
    ["E", "Prize"],
    ["F", "Base Odds Up"],
    ["G", "Odds Down"],
    ["H", "Base 1 in x Odds"],
    ["I", "Prize Cost"],
    ["J", "Contribution split"],
    ["K", "Contribution"],
  ];
  for (const [col, label] of seedHeads) styleCell(setValue(ws, `${col}${seedHeader}`, label), calB, null, null);
  for (const [col, label] of seedHeads) {
    if (col === "E") styleCell(setValue(ws, `${col}${contribHeader}`, "Average Prize"), calB, null, null);
    else styleCell(setValue(ws, `${col}${contribHeader}`, label), calB, null, null);
  }
  tiers.forEach((tier, idx) => {
    const row = seedFirst + idx;
    const baseRow = baseFirst + idx;
    const triggerRow = triggerFirst + idx;
    const contribRow = contribFirst + idx;
    styleCell(setValue(ws, `B${row}`, tier.number), cal, null, null);
    styleCell(setValue(ws, `D${row}`, tier.name), cal, null, null);
    styleCell(setValue(ws, `E${row}`, tier.seed), cal, FMT_INT, GA_INPUT_FILL);
    styleCell(setValue(ws, `F${row}`, `=F${baseRow}*D${stakeRow}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `G${row}`, `=D${baseDown}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `H${row}`, `=G${row}/F${row}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `I${row}`, `=E${row}*F${row}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `J${row}`, `=I${row}/I${seedRtp}`), cal, "0.0000000", GA_SPLIT_FILL);
    styleCell(setValue(ws, `K${row}`, `=K${seedLabel}*J${row}*D${stakeRow}`), cal, FMT_MONEY, null);

    styleCell(setValue(ws, `B${contribRow}`, tier.number), cal, null, null);
    styleCell(setValue(ws, `D${contribRow}`, tier.name), cal, null, null);
    styleCell(setValue(ws, `E${contribRow}`, `=E${triggerRow}-E${row}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `F${contribRow}`, `=F${baseRow}*D${stakeRow}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `G${contribRow}`, `=D${baseDown}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `H${contribRow}`, `=G${contribRow}/F${contribRow}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `I${contribRow}`, `=E${contribRow}*F${contribRow}`), cal, FMT_INT, null);
    styleCell(setValue(ws, `J${contribRow}`, `=I${contribRow}/I${contribSum}`), cal, "0.0000000", GA_SPLIT_FILL);
    styleCell(setValue(ws, `K${contribRow}`, `=K${contribLabel}*J${contribRow}*D${stakeRow}`), cal, FMT_MONEY, null);

    styleCell(setValue(ws, `B${triggerRow}`, tier.number), cal, null, null);
    styleCell(setValue(ws, `D${triggerRow}`, tier.name), cal, null, null);
    styleCell(setValue(ws, `E${triggerRow}`, tier.trigger), cal, FMT_INT, GA_INPUT_FILL);
  });

  styleCell(setValue(ws, `B${seedRtp}`, "RTP"), calB, null, null);
  styleCell(setValue(ws, `D${seedRtp}`, `=I${seedRtp}/D${revenue}`), calB, FMT_PCT, null);
  styleCell(setValue(ws, `F${seedRtp}`, `=SUM(F${seedFirst}:F${seedLast})`), cal, FMT_INT, null);
  styleCell(setValue(ws, `H${seedRtp}`, `=D${baseDown}/F${seedRtp}`), cal, FMT_INT, null);
  styleCell(setValue(ws, `I${seedRtp}`, `=SUM(I${seedFirst}:I${seedLast})`), calB, FMT_INT, null);
  styleCell(setValue(ws, `J${seedRtp}`, `=SUM(J${seedFirst}:J${seedLast})`), calB, FMT_INT, null);
  styleCell(setValue(ws, `K${seedRtp}`, `=SUM(K${seedFirst}:K${seedLast})`), calB, FMT_MONEY, null);

  styleCell(setValue(ws, `B${contribLabel}`, "Jackpot Contributions (excluding seed)"), calB, null, null);
  styleCell(
    setValue(ws, `K${contribLabel}`, `=D${jackpotPct}-I${seedRtp}/D${revenue}`),
    calB,
    FMT_PCT,
    null,
  );
  styleCell(setValue(ws, `H${contribSum}`, `=SUM(H${contribFirst}:H${contribLast})`), cal, FMT_INT, null);
  styleCell(setValue(ws, `I${contribSum}`, `=SUM(I${contribFirst}:I${contribLast})`), calB, FMT_INT, null);
  styleCell(setValue(ws, `J${contribSum}`, `=SUM(J${contribFirst}:J${contribLast})`), calB, FMT_INT, null);
  styleCell(setValue(ws, `K${contribSum}`, `=SUM(K${contribFirst}:K${contribLast})`), calB, FMT_MONEY, null);
  styleCell(setValue(ws, `B${contribRtp}`, "RTP"), calB, null, null);
  styleCell(setValue(ws, `D${contribRtp}`, `=I${contribSum}/D${revenue}`), calB, FMT_PCT, null);

  styleCell(setValue(ws, `B${triggerLabel}`, "Target trigger"), calB, null, null);
  styleCell(setValue(ws, `B${triggerHeader}`, "Number"), cal, null, null);
  styleCell(setValue(ws, `D${triggerHeader}`, "Name"), cal, null, null);
  styleCell(setValue(ws, `E${triggerHeader}`, " Trigger Value"), cal, null, null);
  applyGaJackpotBorders(ws, {
    baseDown,
    stakeRow,
    jackpotPct,
    baseFirst,
    baseLast,
    baseTotal,
    seedHeader,
    seedFirst,
    seedLast,
    contribHeader,
    contribFirst,
    contribLast,
    triggerHeader,
    triggerFirst,
    triggerLast: triggerFirst + n - 1,
  });

  return { jpStart, seedRtp, contribRtp, triggerLabel };
}

const MEDIUM_EDGE = { style: "medium", color: { indexed: 64 } };

function paintRowEdges(ws, row, fromCol, toCol, flags) {
  for (let col = fromCol; col <= toCol; col += 1) {
    const sides = [];
    if (flags.left && col === fromCol) sides.push("left");
    if (flags.right && col === toCol) sides.push("right");
    if (flags.top) sides.push("top");
    if (flags.bottom) sides.push("bottom");
    if (!sides.length) continue;
    const border = {};
    for (const side of sides) border[side] = { ...MEDIUM_EDGE };
    ws.getCell(row, col).border = border;
  }
}

function outlineDataRows(ws, first, last, fromCol, toCol) {
  for (let row = first; row <= last; row += 1) {
    paintRowEdges(ws, row, fromCol, toCol, {
      left: true,
      right: true,
      top: row === first,
      bottom: row === last,
    });
  }
}

function applyGaJackpotBorders(ws, rows) {
  const {
    baseDown,
    stakeRow,
    jackpotPct,
    baseFirst,
    baseLast,
    baseTotal,
    seedHeader,
    seedFirst,
    seedLast,
    contribHeader,
    contribFirst,
    contribLast,
    triggerHeader,
    triggerFirst,
    triggerLast,
  } = rows;
  for (let row = baseDown; row <= jackpotPct; row += 1) {
    paintRowEdges(ws, row, 2, 4, {
      left: true,
      right: true,
      top: row === baseDown,
      bottom: row === jackpotPct,
    });
  }
  paintRowEdges(ws, stakeRow, 6, 9, { left: true, right: true, top: true, bottom: true });
  for (let row = baseFirst; row <= baseLast; row += 1) {
    paintRowEdges(ws, row, 6, 9, { left: true, right: true });
  }
  paintRowEdges(ws, baseTotal, 6, 9, { left: true, right: true, bottom: true });
  paintRowEdges(ws, seedHeader, 2, 8, { left: true, top: true });
  paintRowEdges(ws, seedHeader, 9, 10, { top: true, bottom: true });
  paintRowEdges(ws, seedHeader, 11, 11, { right: true, top: true, bottom: true });
  outlineDataRows(ws, seedFirst, seedLast, 2, 11);
  paintRowEdges(ws, contribHeader, 2, 11, { left: true, right: true, top: true, bottom: true });
  outlineDataRows(ws, contribFirst, contribLast, 2, 11);
  paintRowEdges(ws, triggerHeader, 2, 5, { left: true, right: true, top: true, bottom: true });
  outlineDataRows(ws, triggerFirst, triggerLast, 2, 5);
}

function assertPool(meta) {
  if (!Number.isFinite(meta.pool) || meta.pool <= 0) {
    throw new Error("Frequency Odds Down (pool) is missing or zero.");
  }
  if (!Number.isFinite(meta.rtp)) {
    throw new Error("Frequency RTP Setting is missing.");
  }
}

function ncSummaryRows(buy, hasJp) {
  const insert = (buy ? 1 : 0) + (hasJp ? 2 : 0);
  return {
    revenue: 8 + (buy ? 1 : 0),
    fund: 9 + (buy ? 1 : 0),
    rtpSet: 10 + (buy ? 1 : 0),
    actual: 11 + (buy ? 1 : 0),
    jp: hasJp ? 12 + (buy ? 1 : 0) : null,
    totalRtp: hasJp ? 13 + (buy ? 1 : 0) : null,
    winFreq: 12 + insert,
    hit: 13 + insert,
    gameInfo: 16 + insert,
  };
}

function applyNcNumberFormats(ws, last, tot, buy, hasJp) {
  const sum = ncSummaryRows(buy, hasJp);
  for (let row = 4; row <= last; row += 1) {
    ws.getCell(`B${row}`).numFmt = "00";
    ws.getCell(`D${row}`).numFmt = FMT_MONEY;
    ws.getCell(`E${row}`).numFmt = FMT_INT;
    ws.getCell(`F${row}`).numFmt = FMT_INT;
    ws.getCell(`G${row}`).numFmt = FMT_MONEY;
    ws.getCell(`H${row}`).numFmt = FMT_MONEY;
    ws.getCell(`I${row}`).numFmt = FMT_PCT;
  }
  ws.getCell(`E${tot}`).numFmt = FMT_MONEY;
  ws.getCell(`G${tot}`).numFmt = FMT_MONEY;
  ws.getCell(`I${tot}`).numFmt = FMT_PCT;
  ws.getCell("L6").numFmt = FMT_INT;
  ws.getCell("L7").numFmt = FMT_MONEY;
  if (buy) ws.getCell("L8").numFmt = "0";
  ws.getCell(`L${sum.revenue}`).numFmt = FMT_MONEY;
  ws.getCell(`L${sum.fund}`).numFmt = FMT_MONEY;
  ws.getCell(`L${sum.rtpSet}`).numFmt = FMT_PCT;
  ws.getCell(`L${sum.actual}`).numFmt = FMT_PCT;
  if (sum.jp) ws.getCell(`L${sum.jp}`).numFmt = FMT_PCT;
  if (sum.totalRtp) ws.getCell(`L${sum.totalRtp}`).numFmt = FMT_PCT;
  ws.getCell(`L${sum.winFreq}`).numFmt = FMT_INT;
  ws.getCell(`L${sum.hit}`).numFmt = FMT_MONEY;
}

function buildNcDelivery(ws, freq, meta, source, options = {}) {
  const { nonwin, rows } = source;
  const dc = meta.spec.layout === "DC";
  const last = 4 + rows.length;
  const tot = last + 1;
  const buy = meta.buy;
  const hasJp = meta.hasJp;
  const sum = ncSummaryRows(buy, hasJp);
  const fundCell = `L$${sum.fund}`;
  const poolCell = "L6";
  const helv8 = { name: "Helv", size: 8 };
  const helv8b = { name: "Helv", size: 8, bold: true };
  const helv10 = { name: "Helv", size: 10 };

  setColWidths(ws, dc ? DC_WIDTHS : NC_WIDTHS);

  setValue(ws, "B2", "TIER");
  setValue(ws, "B3", "NUMBER");
  setValue(ws, "C3", "WIN METHOD");
  setValue(ws, "D3", "PRIZE");
  setValue(ws, "E2", "ODDS");
  setValue(ws, "E3", "UP");
  setValue(ws, "F2", "ODDS");
  setValue(ws, "F3", "DOWN");
  setValue(ws, "G2", "PRIZE");
  setValue(ws, "G3", "COST");
  setValue(ws, "H3", "1 in X odds");
  setValue(ws, "I2", "% OF ");
  setValue(ws, "I3", "PRIZE FUND");
  for (const addr of ["B2", "E2", "F2", "G2"]) {
    ws.getCell(addr).font = { ...FONT_GROUP };
    ws.getCell(addr).alignment = { ...CENTER };
  }
  ws.getCell("I2").font = { ...FONT_GROUP, bold: true };
  ws.getCell("I2").alignment = { ...CENTER };
  for (const addr of ["B3", "C3", "D3", "E3", "F3", "G3", "H3", "I3"]) {
    ws.getCell(addr).font = { ...helv8b };
    ws.getCell(addr).alignment = { ...CENTER };
    ws.getCell(addr).border = cloneJson(MEDIUM_BOTTOM);
  }
  ws.getRow(2).height = 16.2;
  ws.getRow(3).height = 16.8;
  ws.getRow(4).height = 16.2;

  setValue(ws, "B4", 0);
  setValue(ws, "C4", nonwin.method == null ? 0 : nonwin.method);
  applyMethodFill(ws.getCell("C4"), nonwin.fill, helv10);
  setValue(ws, "D4", cleanFloat(nonwin.prize));
  const odds = (row) => (dc ? `=F${row}/E${row}` : `=IFERROR(F${row}/E${row},0)`);
  setValue(ws, "E4", `=F4-SUM(E5:E${last})`);
  setValue(ws, "F4", dc ? "=$L$6" : meta.pool);
  setValue(ws, "G4", "=E4*D4");
  setValue(ws, "H4", odds(4));
  if (!dc) setValue(ws, "I4", `=IFERROR(G4/${fundCell},0)`);

  rows.forEach((item, idx) => {
    const row = 5 + idx;
    setValue(ws, `B${row}`, item.tier);
    setValue(ws, `C${row}`, item.method);
    applyMethodFill(ws.getCell(`C${row}`), item.fill, helv10);
    setValue(ws, `D${row}`, item.prize);
    setValue(ws, `E${row}`, item.winners);
    setValue(ws, `F${row}`, dc ? "=$L$6" : meta.pool);
    setValue(ws, `G${row}`, `=E${row}*D${row}`);
    setValue(ws, `H${row}`, odds(row));
    setValue(ws, `I${row}`, `=G${row}/${fundCell}`);
  });

  setValue(ws, `E${tot}`, `=SUM(E4:E${last})`);
  setValue(ws, `G${tot}`, `=SUM(G4:G${last})`);
  setValue(ws, `I${tot}`, dc ? `=SUM(I4:I${last})` : `=G${tot}/${fundCell}`);

  setValue(ws, "K2", "Prize Structure").font = { ...FONT_TITLE };
  setValue(ws, "K3", meta.jurisdiction).font = { ...FONT_TITLE };
  setValue(ws, "K4", meta.title).font = { ...FONT_TITLE };
  const fr = meta.freqRows;
  const kLabels = [
    [6, fr.pool, "Odds Down (ticket quantity if pool based):"],
    [7, fr.base, "Base:"],
  ];
  if (buy) kLabels.push([8, fr.rrp, "RRP(x):"]);
  kLabels.push(
    [sum.revenue, fr.revenue, "Revenus (if pool based):"],
    [sum.fund, fr.fund, "Prize Fund (if pool based):"],
    [sum.rtpSet, fr.rtp, "RTP Setting:"],
    [sum.actual, fr.actual, "Actual RTP:"],
  );
  if (hasJp) {
    kLabels.push([sum.jp, fr.jp, "JP RTP:"], [sum.totalRtp, fr.total, "Total RTP:"]);
  }
  kLabels.push([sum.winFreq, fr.wins, "Winning Tiers Freq"], [sum.hit, fr.hit, "Hit Rate:"]);
  for (const [row, freqRow, fallback] of kLabels) {
    setValue(ws, `K${row}`, storedLabel(freq, freqRow, fallback)).font = { ...helv8b };
    ws.getCell(`K${row}`).alignment = { horizontal: "left" };
  }

  setValue(ws, "L6", meta.pool);
  setValue(ws, "L7", meta.base);
  if (buy) {
    setValue(ws, "L8", meta.rrp);
    setValue(ws, "L9", "=L6*L7*L8");
  } else {
    setValue(ws, "L8", "=L6*L7");
  }
  setValue(ws, `L${sum.fund}`, dc ? `=G${tot}` : `=SUM(G5:G${last})`);
  setValue(ws, `L${sum.rtpSet}`, meta.rtp);
  setValue(ws, `L${sum.actual}`, dc ? `=SUM(G5:G${last})/L${sum.revenue}` : `=L${sum.fund}/L${sum.revenue}`);
  setValue(
    ws,
    `M${sum.actual}`,
    `=IF(ROUND(L${sum.rtpSet},6)=ROUND(L${sum.actual},6),"okay","error")`,
  );
  if (hasJp) {
    setValue(ws, `L${sum.jp}`, "=0");
    setValue(ws, `L${sum.totalRtp}`, `=L${sum.jp}+L${sum.actual}`);
  }
  setValue(ws, `L${sum.winFreq}`, `=SUM(E5:E${last})`);
  setValue(ws, `L${sum.hit}`, `=${poolCell}/L${sum.hit - 1}`);

  let confRange;
  let pricePoints = null;
  if (dc) {
    pricePoints = writeDcGameSheet(ws, sum, last, options.pricePoints);
    confRange = `K${sum.hit + 3}:N${sum.hit + 7}`;
  } else {
    const info = writeNcGameInfo(ws, meta, sum, last);
    confRange = `K${info + 13}:N${info + 17}`;
    ws.getRow(info + 13).height = 18;
  }

  for (const addr of ["L6", "L7", `L${sum.revenue}`, `L${sum.fund}`, `L${sum.rtpSet}`, `L${sum.actual}`, `L${sum.winFreq}`, `L${sum.hit}`]) {
    ws.getCell(addr).font = { ...helv10 };
  }
  ws.getCell("L6").alignment = { horizontal: "right" };
  if (buy) ws.getCell("L8").font = { ...helv10 };
  if (hasJp) {
    ws.getCell(`L${sum.jp}`).font = { ...helv10 };
    ws.getCell(`L${sum.totalRtp}`).font = { ...helv10 };
  }
  for (let row = 4; row <= last; row += 1) {
    ws.getCell(`B${row}`).font = { ...helv8b };
    ws.getCell(`B${row}`).alignment = { ...CENTER };
    ws.getCell(`C${row}`).alignment = { ...CENTER };
    for (const col of ["D", "E", "F", "G", "H", "I"]) {
      ws.getCell(`${col}${row}`).font = { ...helv8 };
    }
  }
  for (const col of ["E", "G", "I"]) ws.getCell(`${col}${tot}`).font = { ...helv8 };
  applyNcNumberFormats(ws, last, tot, buy, hasJp);
  for (let col = 2; col <= 9; col += 1) {
    ws.getCell(last, col).border = cloneJson(MEDIUM_BOTTOM);
  }
  applyConfidential(ws, confRange, "Helv", 8);
  return { last, tot, fundCell, sum, pricePoints };
}

function formatPricePoints(prices) {
  return prices.map((p) => `$${Number(p).toFixed(2)}`).join(", ");
}

function writeDcGameSheet(ws, sum, last, pricePoints) {
  const prices = (pricePoints?.length ? pricePoints : DC_DEFAULT_PRICE_POINTS).map(Number).sort((a, b) => a - b);
  const top = sum.hit + 10;
  const drng = `$D$5:$D$${last}`;
  const grng = `$G$5:$G$${last}`;
  const mult = `${drng}/$L$7`;
  const fund = `$L$${top + 3}`;
  const label = (bold) => ({ font: bold ? FONT_HELV8B : FONT_GROUP, alignment: bold ? { horizontal: "left" } : undefined });
  const lines = [
    ["Total Tickets", true, "=L6", { numFmt: FMT_INT }],
    ["Retail Price", true, "=L7", { numFmt: FMT_MONEY }],
    ["Revenue", false, `=L${top + 1}*L${top}`, { numFmt: FMT_MONEY }],
    ["Prize Fund", true, `=L${sum.fund}`, { numFmt: FMT_MONEY }],
    ["Payout", true, `=L${top + 3}/L${top + 2}`, { numFmt: FMT_PCT }],
    ["Odds", false, `="1 in "&TEXT(L${sum.hit},"0.00")`, { alignment: CENTER }],
    ["Top Prize", false, `=TEXT(MAX(${drng})/$L$7,"#,##0")&"x"`, { alignment: CENTER }],
    ["Max Top Prize $ Value", false, `=MAX(${drng})/$L$7*${prices[prices.length - 1]}`, { numFmt: FMT_INT }],
    ["Price Points", false, formatPricePoints(prices), {}],
    [null, false, null, {}],
    ["Prize (as multiple of stake)", false, "Value Distribution", {}],
    ["NIL", "band", "-", { numFmt: FMT_PCT }],
    [">0 to <1", "band", `=SUMPRODUCT((${drng}>0)*(${mult}<1)*${grng})/${fund}`, { numFmt: FMT_PCT }],
    [1, "band", `=SUMPRODUCT((${mult}=1)*${grng})/${fund}`, { numFmt: FMT_PCT }],
  ];
  const bands = [[1, 2], [2, 5], [5, 10], [10, 20], [20, 30], [30, 50], [50, 100]];
  for (const [lo, hi] of bands) {
    lines.push([`>${lo} to ${hi}`, "band", `=SUMPRODUCT((${mult}>${lo})*(${mult}<=${hi})*${grng})/${fund}`, { numFmt: FMT_PCT }]);
  }
  lines.push([">100", "band", `=SUMPRODUCT((${mult}>100)*${grng})/${fund}`, { numFmt: FMT_PCT }]);
  const end = top + lines.length - 1;
  lines.forEach(([text, kind, value, extra], idx) => {
    const row = top + idx;
    const edge = (side) => ({
      [side]: { style: "medium", color: { indexed: 64 } },
      ...(row === top ? { top: { style: "medium", color: { indexed: 64 } } } : {}),
      ...(row === end ? { bottom: { style: "medium", color: { indexed: 64 } } } : {}),
    });
    const kStyle = kind === "band" ? { font: FONT_GROUP, alignment: CENTER } : label(kind === true);
    put(ws, `K${row}`, text ?? undefined, { ...kStyle, border: edge("left") });
    put(ws, `L${row}`, value ?? undefined, { font: FONT_GROUP, ...extra, border: edge("right") });
  });
  return prices;
}

function writeNcGameInfo(ws, meta, sum, last) {
  const rng = `$C$5:$C$${last}`;
  const drng = `$D$5:$D$${last}`;
  const erng = `$E$5:$E$${last}`;
  const grng = `$G$5:$G$${last}`;
  const poolCell = "L6";
  const helv8b = FONT_HELV8B;
  const helv10 = FONT_HELV10;
  const revenueCell = `L${sum.revenue}`;
  const info = sum.gameInfo;
  setValue(ws, `K${info}`, `Game info for ${meta.spec.code}`).font = { ...FONT_HELV10B };
  ws.getCell(`K${info}`).alignment = { horizontal: "left" };
  const infoLabels = [
    [info + 1, "Bonus Odds"],
    [info + 2, "Avg. Bonus Prize"],
    [info + 3, "Free Game Odds"],
    [info + 4, "Avg. Free Game Prize"],
    [info + 5, "Top Prize Odds"],
    [info + 6, "Top Prize Amount"],
    [info + 7, "RTP 0-5x"],
    [info + 8, "RTP 5-10x"],
    [info + 9, "RTP 10-25x"],
    [info + 10, "RTP 25x+"],
  ];
  for (const [row, text] of infoLabels) {
    setValue(ws, `K${row}`, text).font = { ...helv8b };
    ws.getCell(`K${row}`).alignment = { horizontal: "left" };
  }
  for (const row of [info + 1, info + 2]) ws.getCell(`K${row}`).fill = themeFill(9);
  for (const row of [info + 3, info + 4]) ws.getCell(`K${row}`).fill = themeFill(7);
  setValue(
    ws,
    `L${info + 1}`,
    `=IF(SUMIF(${rng},"*bonus*",${erng})=0,"-",${poolCell}/SUMIF(${rng},"*bonus*",${erng}))`,
  );
  setValue(
    ws,
    `L${info + 2}`,
    `=IF(SUMIF(${rng},"*bonus*",${erng})=0,"-",SUMPRODUCT(ISNUMBER(SEARCH("bonus",${rng}))*(${grng}))/SUMIF(${rng},"*bonus*",${erng}))`,
  );
  setValue(
    ws,
    `L${info + 3}`,
    `=IF(SUMIF(${rng},"*freeplay*",${erng})=0,"-",${poolCell}/SUMIF(${rng},"*freeplay*",${erng}))`,
  );
  setValue(
    ws,
    `L${info + 4}`,
    `=IF(SUMIF(${rng},"*freeplay*",${erng})=0,"-",SUMPRODUCT(ISNUMBER(SEARCH("freeplay",${rng}))*(${grng}))/SUMIF(${rng},"*freeplay*",${erng}))`,
  );
  setValue(ws, `L${info + 6}`, `=MAX(${drng})`);
  setValue(ws, `L${info + 5}`, `=IFERROR(${poolCell}/SUMIF(${drng},L${info + 6},${erng}),"-")`);
  const bandDen = revenueCell;
  setValue(
    ws,
    `L${info + 7}`,
    `=IFERROR(SUMPRODUCT((${drng}>0)*(${drng}/$L$7<5)*(${grng}))/${bandDen},0)`,
  );
  setValue(
    ws,
    `L${info + 8}`,
    `=IFERROR(SUMPRODUCT((${drng}/$L$7>=5)*(${drng}/$L$7<10)*(${grng}))/${bandDen},0)`,
  );
  setValue(
    ws,
    `L${info + 9}`,
    `=IFERROR(SUMPRODUCT((${drng}/$L$7>=10)*(${drng}/$L$7<25)*(${grng}))/${bandDen},0)`,
  );
  setValue(
    ws,
    `L${info + 10}`,
    `=IFERROR(SUMPRODUCT((${drng}/$L$7>=25)*(${grng}))/${bandDen},0)`,
  );
  for (let row = info + 1; row <= info + 6; row += 1) {
    ws.getCell(`L${row}`).numFmt = FMT_MONEY;
  }
  for (let row = info + 7; row <= info + 10; row += 1) {
    ws.getCell(`L${row}`).numFmt = FMT_PCT;
  }
  for (let row = info + 1; row <= info + 10; row += 1) {
    ws.getCell(`L${row}`).font = row > info + 6 ? { ...FONT_GROUP } : { ...helv10 };
  }
  return info;
}

function vaSummaryRows(buy, hasJp) {
  const b = buy ? 1 : 0;
  const j = hasJp ? 2 : 0;
  return {
    pool: 6,
    base: 7,
    rrp: buy ? 8 : null,
    revenue: 8 + b,
    fund: 9 + b,
    rtpSet: 10 + b,
    actual: 11 + b,
    jp: hasJp ? 12 + b : null,
    totalRtp: hasJp ? 13 + b : null,
    winFreq: 12 + b + j,
    hit: 13 + b + j,
  };
}

function vaHitFormula(formula, meta, sum) {
  return String(formula)
    .replace(/^=/, "")
    .replace(/(^|[^A-Za-z0-9!'$])(\$?)N(\$?)(\d+)/g, (_, pre, _c, _r, row) => {
      const r = Number(row);
      return r === meta.freqRows.hit ? `${pre}$C$${sum.hit}` : `${pre}Frequency!$N$${r}`;
    });
}

function buildVaDelivery(ws, freq, meta, source) {
  const { nonwin, rows } = source;
  const fr = meta.freqRows;
  const sum = vaSummaryRows(meta.buy, meta.hasJp);
  const group = sum.hit + 3;
  const sub = group + 1;
  const nw = group + 2;
  const first = nw + 1;
  const last = nw + rows.length;
  const tot = last + 1;
  const hits = meta.hits || [];
  const widths = {
    ...VA_WIDTHS,
    ...(hits.length ? { 5: 20.66 } : {}),
    ...(meta.hasJp ? {} : { 10: 12, 11: 37.11 }),
  };
  setColWidths(ws, widths);

  const label = { font: FONT_HELV8B, alignment: { horizontal: "left" } };
  const value = (numFmt) => ({ font: FONT_HELV10, numFmt });

  put(ws, "B2", "Prize Structure", { font: FONT_TITLE });
  put(ws, "B3", meta.jurisdiction, { font: FONT_TITLE });
  put(ws, "B4", meta.title, { font: FONT_TITLE });
  for (const row of [2, 3, 4]) ws.getRow(row).height = 16.2;

  const lines = [
    [sum.pool, fr.pool, "Odds Down (ticket quantity if pool based):", meta.pool, { ...value(FMT_INT), alignment: { horizontal: "right" } }],
    [sum.base, fr.base, "Base:", meta.base, value(FMT_MONEY)],
  ];
  if (meta.buy) lines.push([sum.rrp, fr.rrp, "RRP(x):", meta.rrp, value(FMT_MONEY)]);
  lines.push(
    [sum.revenue, fr.revenue, "Revenus (if pool based):", meta.buy ? "=C6*C7*C8" : "=C6*C7", value(FMT_MONEY)],
    [sum.fund, fr.fund, "Prize Fund (if pool based):", `=G${tot}`, value(FMT_MONEY)],
    [sum.rtpSet, fr.rtp, "RTP Setting", meta.rtp, value(FMT_PCT)],
    [sum.actual, fr.actual, "Actual RTP:", `=SUM(G${first}:G${last})/C${sum.revenue}`, value(FMT_PCT)],
  );
  if (meta.hasJp) {
    lines.push(
      [sum.jp, fr.jp, "JP RTP", "=0", value(FMT_PCT)],
      [sum.totalRtp, fr.total, "Total RTP", `=C${sum.jp}+C${sum.actual}`, value(FMT_PCT)],
    );
  }
  lines.push(
    [sum.winFreq, fr.wins, "Winning Tiers Freq", `=SUM(E${first}:E${last})`, value(FMT_INT)],
    [sum.hit, fr.hit, "Hit Rate:", `=C6/C${sum.winFreq}`, value(FMT_MONEY)],
  );
  for (const [row, freqRow, fallback, v, style] of lines) {
    put(ws, `B${row}`, storedLabel(freq, freqRow, fallback), label);
    put(ws, `C${row}`, v, style);
  }

  hits.forEach((hit, idx) => {
    const row = 9 + idx;
    put(ws, `E${row}`, hit.label, label);
    put(ws, `F${row}`, `=${vaHitFormula(hit.formula, meta, sum)}`, { font: FONT_HELV8, numFmt: FMT_MONEY });
  });

  const groupTexts = { B: "TIER", E: "ODDS", F: "ODDS", G: "PRIZE", I: "% OF " };
  for (const [col, text] of Object.entries(groupTexts)) {
    put(ws, `${col}${group}`, text, { font: col === "I" ? { ...FONT_GROUP, bold: true } : FONT_GROUP, alignment: CENTER });
  }
  const subTexts = ["NUMBER", "WIN METHOD", "PRIZE", "UP", "DOWN", "COST", "1 in X odds", "PRIZE FUND"];
  subTexts.forEach((text, idx) => {
    put(ws, { row: sub, col: 2 + idx }, text, { font: FONT_HELV8B, alignment: CENTER, border: MEDIUM_BOTTOM });
  });
  ws.getRow(group).height = 16.8;
  ws.getRow(sub).height = 16.8;
  ws.getRow(nw).height = 16.2;

  const tierStyle = { font: FONT_HELV8B, alignment: CENTER, numFmt: "00" };
  const num = (numFmt) => ({ font: FONT_HELV8, numFmt });
  put(ws, `B${nw}`, 0, tierStyle);
  put(ws, `C${nw}`, nonwin.method == null ? 0 : nonwin.method, { alignment: CENTER });
  applyMethodFill(ws.getCell(`C${nw}`), nonwin.fill, FONT_HELV10);
  put(ws, `D${nw}`, cleanFloat(nonwin.prize), num(FMT_MONEY));
  put(ws, `E${nw}`, `=F${nw}-SUM(E${first}:E${last})`, num(FMT_INT));
  put(ws, `F${nw}`, "=$C$6", num(FMT_INT));
  put(ws, `G${nw}`, `=E${nw}*D${nw}`, num(FMT_MONEY));
  put(ws, `H${nw}`, `=F${nw}/E${nw}`, num(FMT_MONEY));

  rows.forEach((item, idx) => {
    const row = first + idx;
    put(ws, `B${row}`, item.tier, tierStyle);
    put(ws, `C${row}`, item.method, { alignment: CENTER });
    applyMethodFill(ws.getCell(`C${row}`), item.fill, FONT_HELV10);
    put(ws, `D${row}`, item.prize, num(FMT_MONEY));
    put(ws, `E${row}`, item.winners, num(FMT_INT));
    put(ws, `F${row}`, "=$C$6", num(FMT_INT));
    put(ws, `G${row}`, `=E${row}*D${row}`, num(FMT_MONEY));
    put(ws, `H${row}`, `=F${row}/E${row}`, num(FMT_MONEY));
    put(ws, `I${row}`, `=G${row}/C$${sum.fund}`, num(FMT_PCT));
  });
  addBottomEdge(ws, last, 2, 9);

  put(ws, `E${tot}`, `=SUM(E${nw}:E${last})`, num(FMT_MONEY));
  put(ws, `G${tot}`, `=SUM(G${nw}:G${last})`, num(FMT_MONEY));
  put(ws, `I${tot}`, `=SUM(I${nw}:I${last})`, num(FMT_PCT));

  applyConfidential(ws, hits.length ? "H9:K13" : "E9:H13", "Helv", 8);
  return { last, tot, first, sum, widths };
}

function gaLeftRow(hasJp) {
  const shift = hasJp ? 4 : 0;
  return {
    shift,
    winFreq: 10 + shift,
    hit: 11 + shift,
    odds: 13 + shift,
    bonus: 14 + shift,
    freeplay: 15 + shift,
    rrp: 16 + shift,
    avg: 17 + shift,
    bonusAvg: 18 + shift,
    freeplayAvg: 19 + shift,
    confidential: 22 + shift,
  };
}

function buildGaDelivery(ws, freq, meta, source, dateSerial) {
  const { nonwin, rows: unsorted } = source;
  const rows = sortGaWinning(unsorted);
  const last = 3 + rows.length;
  const tot = last + 1;
  const hasJp = meta.hasJp;
  const left = gaLeftRow(hasJp);
  const calibri = { name: "Calibri", size: 10 };
  const calibriB = { name: "Calibri", size: 10, bold: true };

  setColWidths(ws, GA_WIDTHS);
  setValue(ws, "A1", meta.jurisdiction).font = { ...calibriB };
  const titleFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { theme: 4, tint: 0.7999816888943144 },
  };
  setValue(ws, "A2", meta.title).font = { ...calibriB };
  ws.getCell("A2").fill = titleFill;
  setValue(ws, "A3", dateSerial).numFmt = FMT_DATE_GA;
  ws.getCell("A3").font = { ...calibri };
  ws.getCell("A3").fill = titleFill;

  const leftLabels = [
    [5, "Odds Down (ticket quantity if pool based):"],
    [6, "Base:"],
    [7, "Revenus (if pool based):"],
    [8, "Prize Fund (if pool based):"],
    [9, "Main Game RTP:"],
  ];
  if (hasJp) {
    leftLabels.push(
      [10, "JP Seed  RTP:"],
      [11, "JP Contribution RTP:"],
      [12, "Total JP RTP:"],
      [13, "Total Game RTP:"],
    );
  }
  leftLabels.push(
    [left.winFreq, "Winning Tiers Freq"],
    [left.hit, "Hit Rate:"],
    [left.odds, "Odds (excluding JP)"],
    [left.bonus, "Bonus "],
    [left.freeplay, "Freeplay "],
    [left.avg, "Avg Win (excluding JP)"],
    [left.bonusAvg, "Bonus"],
    [left.freeplayAvg, "Freeplay"],
  );
  if (meta.buy) leftLabels.push([left.rrp, "RRP(x):"]);
  for (const [row, text] of leftLabels) {
    setValue(ws, `A${row}`, text).font = { ...calibriB };
  }

  setValue(ws, "B5", "=M3");
  setValue(ws, "B6", meta.base);
  setValue(ws, "B7", meta.buy ? `=B5*B6*B${left.rrp}` : "=B5*B6");
  setValue(ws, "B8", `=N${tot}`);
  setValue(ws, "B9", "=B8/B7");
  if (hasJp) {
    setValue(ws, "B10", "='Progressive Jackpots'!C30");
    setValue(ws, "B11", "='Progressive Jackpots'!C31");
    setValue(ws, "B12", "=B10+B11");
    setValue(ws, "B13", "=B9+B12");
  }
  setValue(ws, `B${left.winFreq}`, `=SUM(L4:L${last})`);
  setValue(ws, `B${left.hit}`, `=B5/B${left.winFreq}`);
  if (meta.buy) setValue(ws, `B${left.rrp}`, meta.rrp);

  const fWin = `$F$4:$F$${last}`;
  const lWin = `$L$4:$L$${last}`;
  const fAll = `$F$3:$F$${last}`;
  const lAll = `$L$3:$L$${last}`;
  const nAll = `$N$3:$N$${last}`;
  setValue(
    ws,
    `B${left.bonus}`,
    `=IF(SUMIF(${fWin},"*bonus*",${lWin})=0,"-",$B$5/SUMIF(${fWin},"*bonus*",${lWin}))`,
  );
  setValue(
    ws,
    `B${left.freeplay}`,
    `=IF(SUMIF(${fWin},"*freeplay*",${lWin})=0,"-",$B$5/SUMIF(${fWin},"*freeplay*",${lWin}))`,
  );
  setValue(
    ws,
    `B${left.bonusAvg}`,
    `=IF(SUMIF(${fAll},"*bonus*",${lAll})=0,"-",SUMIF(${fAll},"*bonus*",${nAll})/SUMIF(${fAll},"*bonus*",${lAll}))`,
  );
  setValue(
    ws,
    `B${left.freeplayAvg}`,
    `=IF(SUMIF(${fAll},"*freeplay*",${lAll})=0,"-",SUMIF(${fAll},"*freeplay*",${nAll})/SUMIF(${fAll},"*freeplay*",${lAll}))`,
  );

  setValue(ws, "P1", "% OF ");
  setValue(ws, "E2", "order");
  setValue(ws, "F2", "Tiers");
  setValue(ws, "G2", "Prize");
  setValue(ws, "H2", "1 in x Odds");
  setValue(ws, "I2", "Prize");
  setValue(ws, "J2", "Chances Are 1 in X");
  setValue(ws, "K2", "Winners");
  setValue(ws, "L2", "Odds Up");
  setValue(ws, "M2", "Odds Down");
  setValue(ws, "N2", "Prize Cost");
  setValue(ws, "P2", "Prize Fund");
  for (const col of ["P1", "E2", "F2", "G2", "H2", "I2", "J2", "K2", "L2", "M2", "N2", "P2"]) {
    ws.getCell(col).font = { ...calibriB };
  }
  for (let col = 5; col <= 16; col += 1) {
    const cell = ws.getCell(2, col);
    cell.border = { ...MEDIUM_BOTTOM };
  }

  setValue(ws, "E3", 0);
  setValue(ws, "F3", nonwin.method == null ? 0 : nonwin.method);
  applyMethodFill(ws.getCell("F3"), nonwin.fill, calibri);
  setValue(ws, "G3", cleanFloat(nonwin.prize));
  setValue(ws, "H3", "=IFERROR(M3/L3,0)");
  setValue(ws, "L3", `=B5-SUM(L4:L${last})`);
  setValue(ws, "M3", meta.pool);
  setValue(ws, "N3", "=G3*L3");

  rows.forEach((item, idx) => {
    const row = 4 + idx;
    setValue(ws, `E${row}`, idx + 1);
    setValue(ws, `F${row}`, item.method);
    applyMethodFill(ws.getCell(`F${row}`), item.fill, calibri);
    setValue(ws, `G${row}`, item.prize);
    setValue(ws, `H${row}`, `=IFERROR(M${row}/L${row},0)`);
    setValue(ws, `I${row}`, `=IF(G${row}<>G${row + 1},G${row},"")`);
    setValue(ws, `J${row}`, `=IFERROR($B$5/SUMIF($G$3:$G$${last},I${row},$L$3:$L$${last}),"")`);
    setValue(ws, `K${row}`, `=IF(I${row}<>"",SUMIF($G$3:$G$${last},I${row},$L$3:$L$${last}),"")`);
    setValue(ws, `L${row}`, item.winners);
    setValue(ws, `M${row}`, meta.pool);
    setValue(ws, `N${row}`, `=G${row}*L${row}`);
    setValue(ws, `O${row}`, `=IF(K${row}<>"",G${row}*K${row},"")`);
    setValue(ws, `P${row}`, `=IF(O${row}<>"",O${row}/$B$8,"")`);
  });

  for (let col = 5; col <= 16; col += 1) {
    ws.getCell(last, col).border = { ...MEDIUM_BOTTOM };
  }

  setValue(ws, `K${tot}`, `=SUM(K3:K${last})`);
  setValue(ws, `L${tot}`, `=SUM(L3:L${last})`);
  setValue(ws, `N${tot}`, `=SUM(N3:N${last})`);
  setValue(ws, `O${tot}`, `=SUM(O3:O${last})`);
  setValue(ws, `P${tot}`, `=SUM(P3:P${last})`);

  for (let row = 3; row <= last; row += 1) {
    ws.getCell(`E${row}`).numFmt = FMT_INT;
    ws.getCell(`G${row}`).numFmt = FMT_MONEY;
    ws.getCell(`H${row}`).numFmt = FMT_MONEY;
    ws.getCell(`L${row}`).numFmt = FMT_INT;
    ws.getCell(`M${row}`).numFmt = FMT_INT;
    ws.getCell(`N${row}`).numFmt = FMT_MONEY;
    if (row >= 4) {
      ws.getCell(`I${row}`).numFmt = FMT_MONEY;
      ws.getCell(`J${row}`).numFmt = FMT_MONEY;
      ws.getCell(`K${row}`).numFmt = FMT_INT;
      ws.getCell(`O${row}`).numFmt = FMT_MONEY;
      ws.getCell(`P${row}`).numFmt = FMT_PCT;
    }
  }
  ws.getCell(`K${tot}`).numFmt = FMT_INT;
  ws.getCell(`L${tot}`).numFmt = FMT_INT;
  ws.getCell(`N${tot}`).numFmt = FMT_MONEY;
  ws.getCell(`O${tot}`).numFmt = FMT_MONEY;
  ws.getCell(`P${tot}`).numFmt = FMT_PCT;
  ws.getCell("B5").numFmt = FMT_INT;
  ws.getCell("B6").numFmt = "0";
  ws.getCell("B7").numFmt = FMT_MONEY;
  ws.getCell("B8").numFmt = FMT_MONEY;
  ws.getCell("B9").numFmt = FMT_PCT;
  ws.getCell(`B${left.winFreq}`).numFmt = FMT_INT;
  ws.getCell(`B${left.hit}`).numFmt = FMT_HIT;
  ws.getCell(`B${left.bonus}`).numFmt = FMT_HIT;
  ws.getCell(`B${left.freeplay}`).numFmt = FMT_HIT;
  ws.getCell(`B${left.bonusAvg}`).numFmt = FMT_MONEY;
  ws.getCell(`B${left.freeplayAvg}`).numFmt = FMT_MONEY;
  if (hasJp) {
    for (const row of [10, 11, 12, 13]) ws.getCell(`B${row}`).numFmt = FMT_PCT;
  }
  if (meta.buy) ws.getCell(`B${left.rrp}`).numFmt = "0";

  for (let row = 3; row <= last; row += 1) {
    for (const col of ["E", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]) {
      const cell = ws.getCell(`${col}${row}`);
      if (cell.font && Object.keys(cell.font).length) continue;
      cell.font = { ...calibri };
    }
  }

  if (last >= 4) {
    ws.addConditionalFormatting({
      ref: `P4:P${last}`,
      rules: [
        {
          type: "colorScale",
          cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
          color: [{ argb: "FFF8696B" }, { argb: "FFFFEB84" }, { argb: "FF63BE7B" }],
        },
      ],
    });
  }

  applyConfidential(ws, `A${left.confidential}:D${left.confidential + 4}`, "Calibri", 8);
  return { last, tot, sorted: rows, left };
}

function independentStats(rows, meta) {
  const wins = rows.reduce((s, r) => s + r.winners, 0);
  const fund = rows.reduce((s, r) => s + r.prize * r.winners, 0);
  const revenue = meta.pool * meta.base * (meta.buy ? meta.rrp : 1);
  return { wins, fund, revenue, actualRtp: revenue ? fund / revenue : null, hitRate: wins ? meta.pool / wins : null };
}

export async function inspectPack(buffer, filename = "upload.xlsx", lottery = "NC") {
  const code = String(lottery || "NC").toUpperCase();
  const spec = PACK_LOTTERIES[code];
  if (!spec) throw new Error(`Unsupported pack lottery: ${lottery}`);
  const wb = await loadWorkbook(buffer);
  const freq = requireSheet(wb, "Frequency");
  const meta = resolveJackpot(wb, spec, lookupMeta(freq, code));
  assertPool(meta);
  const source = frequencyWinningRows(freq);
  const stats = independentStats(source.rows, meta);
  return {
    filename,
    lottery: code,
    layout: spec.layout,
    jurisdiction: meta.jurisdiction,
    title: meta.title,
    identity: `${meta.jurisdiction} ${meta.title}`.trim(),
    schema: spec.layout,
    buy: meta.buy,
    rrp: meta.rrp,
    hasJp: meta.hasJp,
    quantity: meta.pool,
    base: meta.base,
    rtp: meta.rtp,
    winningTiers: source.rows.length,
    zeroFrequency: source.zeroFrequency,
    wins: stats.wins,
    hitRate: stats.hitRate,
    prizeFund: stats.fund,
    actualRtp: stats.actualRtp,
    sheets: wb.worksheets.map((ws) => ws.name),
    existingDelivery: sheet(wb, PACK_SHEET) ? [PACK_SHEET] : [],
    jpCount: meta.hasJp ? 1 : 0,
    jpNames: meta.hasJp ? ["Progressive Jackpots"] : [],
    sideHits: meta.hits.map((h) => h.label),
  };
}

export async function buildPack(buffer, filename, options = {}) {
  if (!ExcelJS) throw new Error("ExcelJS is not loaded.");
  const code = String(options.lottery || "NC").toUpperCase();
  const spec = PACK_LOTTERIES[code];
  if (!spec) throw new Error(`Unsupported pack lottery: ${options.lottery}`);
  const wb = await loadWorkbook(buffer);
  const freq = requireSheet(wb, "Frequency");
  const meta = resolveJackpot(wb, spec, lookupMeta(freq, code));
  assertPool(meta);
  const source = frequencyWinningRows(freq);
  const ws = insertSheet(wb, PACK_SHEET);
  let built;
  if (spec.layout === "NC" || spec.layout === "DC") {
    built = buildNcDelivery(ws, freq, meta, source, { pricePoints: options.pricePoints });
  } else if (spec.layout === "VA") built = buildVaDelivery(ws, freq, meta, source);
  else built = buildGaDelivery(ws, freq, meta, source, packDateSerial(filename));
  if (meta.hasJp) {
    const pj = requireSheet(wb, "Progressive Jackpots");
    if (spec.layout === "NC" || spec.layout === "DC") {
      const jpStart = built.tot + (spec.layout === "DC" ? 3 : 2);
      const widths = spec.layout === "DC" ? DC_WIDTHS : NC_WIDTHS;
      copyJackpotSetting(ws, pj, jpStart, jackpotFreqMap(spec.layout, meta, built.sum), widths);
      linkJackpotSummary(ws, meta, jpStart);
      built.jpStart = jpStart;
    } else if (spec.layout === "VA") {
      const jpStart = built.tot + 3;
      const sum = built.sum;
      copyJackpotSetting(ws, pj, jpStart, jackpotFreqMap(spec.layout, meta, sum), built.widths);
      const linked = meta.jpRtpFormula ? rewritePjRefs(meta.jpRtpFormula, jpStart - 2) : "";
      setValue(ws, `C${sum.jp}`, `=${linked || `C${jpStart + 28}+C${jpStart + 29}`}`);
      built.jpStart = jpStart;
    } else {
      const jpStart = built.tot + 2;
      const placed = buildGaJackpotSetting(ws, pj, jpStart);
      setValue(ws, "B10", `=D${placed.seedRtp}`);
      setValue(ws, "B11", `=D${placed.contribRtp}`);
      ws.getCell("B10").numFmt = FMT_PCT;
      ws.getCell("B11").numFmt = FMT_PCT;
      ws.getCell("B10").font = { name: "Calibri", size: 10 };
      ws.getCell("B11").font = { name: "Calibri", size: 10 };
      built.jpStart = placed.jpStart;
      built.seedRtp = placed.seedRtp;
      built.contribRtp = placed.contribRtp;
    }
  }
  applyTabColor(ws);
  orderPackSheets(wb, spec.layout);
  disableFullCalcOnLoad(wb);
  const stats = independentStats(source.rows, meta);
  const rawBuffer = await wb.xlsx.writeBuffer();
  const outBuffer = await patchPackXlsx(rawBuffer);
  return {
    buffer: outBuffer,
    report: {
      filename,
      lottery: code,
      layout: spec.layout,
      schema: spec.layout,
      identity: `${meta.jurisdiction} ${meta.title}`.trim(),
      jurisdiction: meta.jurisdiction,
      title: meta.title,
      buy: meta.buy,
      rrp: meta.rrp,
      hasJp: meta.hasJp,
      quantity: meta.pool,
      base: meta.base,
      rtp: meta.rtp,
      winningTiers: source.rows.length,
      uniquePrizes: new Set(source.rows.map((r) => r.prize)).size,
      zeroFrequency: source.zeroFrequency,
      wins: stats.wins,
      hitRate: stats.hitRate,
      prizeFund: stats.fund,
      actualRtp: stats.actualRtp,
      delivery: built,
      sheets: wb.worksheets.map((ws2) => ws2.name),
    },
  };
}

export function assertPlainNumberFormats(ws, label = "Delivery") {
  const maxRow = Math.min(ws.rowCount || 0, 2500);
  const maxCol = Math.min(ws.columnCount || 0, 24);
  for (let row = 1; row <= maxRow; row += 1) {
    for (let col = 1; col <= maxCol; col += 1) {
      const fmt = ws.getCell(row, col).numFmt;
      if (fmt && /[_*]/.test(String(fmt))) {
        throw new Error(`${label} ${ws.getCell(row, col).address} has accounting format ${fmt}`);
      }
    }
  }
}

const CRC32_TABLE = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(data) {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i += 1) crc = CRC32_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(view, offset) {
  return view.getUint16(offset, true);
}

function u32(view, offset) {
  return view.getUint32(offset, true);
}

function asUint8(buffer) {
  if (buffer instanceof Uint8Array) return buffer;
  return new Uint8Array(buffer);
}

async function inflateRaw(bytes) {
  const ds = new DecompressionStream("deflate-raw");
  return new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(ds)).arrayBuffer());
}

async function deflateRaw(bytes) {
  const cs = new CompressionStream("deflate-raw");
  return new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(cs)).arrayBuffer());
}

async function unzipEntries(buffer) {
  const bytes = asUint8(buffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  const start = Math.max(0, bytes.length - 22 - 65536);
  for (let i = bytes.length - 22; i >= start; i -= 1) {
    if (u32(view, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("Workbook is not a zip package.");
  const count = u16(view, eocd + 10);
  const cdOffset = u32(view, eocd + 16);
  if (cdOffset === 0xffffffff) throw new Error("ZIP64 workbooks are not supported.");
  const files = new Map();
  let ptr = cdOffset;
  for (let n = 0; n < count; n += 1) {
    if (u32(view, ptr) !== 0x02014b50) throw new Error("Corrupt xlsx central directory.");
    const method = u16(view, ptr + 10);
    const compSize = u32(view, ptr + 20);
    const uncompSize = u32(view, ptr + 24);
    const nameLen = u16(view, ptr + 28);
    const extraLen = u16(view, ptr + 30);
    const commentLen = u16(view, ptr + 32);
    const localOff = u32(view, ptr + 42);
    const name = new TextDecoder().decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen));
    const localNameLen = u16(view, localOff + 26);
    const localExtra = u16(view, localOff + 28);
    const dataStart = localOff + 30 + localNameLen + localExtra;
    const packed = bytes.subarray(dataStart, dataStart + compSize);
    let data;
    if (method === 0 || packed.length === 0 || uncompSize === 0) data = new Uint8Array(packed);
    else if (method === 8) data = await inflateRaw(packed);
    else throw new Error(`Unsupported zip method ${method} for ${name}`);
    files.set(name, data);
    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

function concatBytes(chunks) {
  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

function dosDate() {
  const d = new Date();
  const time = (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2);
  const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  return { time, date };
}

async function zipEntries(files) {
  const { time, date } = dosDate();
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const [name, data] of files) {
    const nameBytes = new TextEncoder().encode(name);
    const store = data.length === 0;
    const compressed = store ? data : await deflateRaw(data);
    const method = store ? 0 : 8;
    const crc = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length + compressed.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(8, method, true);
    lv.setUint16(10, time, true);
    lv.setUint16(12, date, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, compressed.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    local.set(compressed, 30 + nameBytes.length);
    const central = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(10, method, true);
    cv.setUint16(12, time, true);
    cv.setUint16(14, date, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, compressed.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    locals.push(local);
    centrals.push(central);
    offset += local.length;
  }
  const cd = concatBytes(centrals);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, files.size, true);
  ev.setUint16(10, files.size, true);
  ev.setUint32(12, cd.length, true);
  ev.setUint32(16, offset, true);
  return concatBytes([...locals, cd, eocd]);
}

function decodeText(bytes) {
  return new TextDecoder("utf-8").decode(bytes);
}

function encodeText(text) {
  return new TextEncoder().encode(text);
}

function plainFormatCode(code) {
  if (!/[_*]/.test(code)) return code;
  const first = code.split(";")[0];
  if (first.includes("%")) return "0.00%";
  const match = first.match(/\.([0#]+)/);
  if (match) return `#,##0.${"0".repeat(match[1].length)}`;
  return "#,##0";
}

function patchStylesXml(xml) {
  let text = xml.replace(/formatCode="([^"]*)"/g, (_, code) => `formatCode="${plainFormatCode(code)}"`);
  const idMap = { 37: 3, 38: 3, 41: 3, 42: 3, 39: 4, 40: 4, 43: 4, 44: 4 };
  for (const [oldId, newId] of Object.entries(idMap)) {
    text = text.replaceAll(`numFmtId="${oldId}"`, `numFmtId="${newId}"`);
  }
  return text;
}

function patchWorkbookXml(xml) {
  const repl = (tag) => {
    let next = tag.replace(/\s+calcCompleted="[^"]*"/g, "").replace(/\s+fullCalcOnLoad="[^"]*"/g, "");
    if (/calcMode=/.test(next)) next = next.replace(/calcMode="[^"]*"/, 'calcMode="auto"');
    else next = next.replace("<calcPr", '<calcPr calcMode="auto"');
    if (/calcOnSave=/.test(next)) next = next.replace(/calcOnSave="[^"]*"/, 'calcOnSave="1"');
    else next = next.replace("<calcPr", '<calcPr calcOnSave="1"');
    return next;
  };
  if (!xml.includes("<calcPr")) {
    return xml.replace("</workbook>", '<calcPr calcMode="auto" calcOnSave="1"/></workbook>');
  }
  return xml.replace(/<calcPr\b[^>]*\/>/g, repl);
}

function sheetNamesFromWorkbook(xml) {
  return [...xml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/g)].map((match) => match[1]);
}

function stripInternalWorkbookIndex(xml, sheetNames) {
  let text = xml;
  for (const name of sheetNames) {
    const quoted = `'${name.replaceAll("'", "''")}'`;
    text = text.replaceAll(`[1]${quoted}!`, `${quoted}!`);
    if (!name.includes(" ") && !name.includes("'")) text = text.replaceAll(`[1]${name}!`, `${name}!`);
  }
  return text;
}

export async function patchPackXlsx(buffer) {
  const files = await unzipEntries(buffer);
  const styles = files.get("xl/styles.xml");
  const book = files.get("xl/workbook.xml");
  const sheetNames = book ? sheetNamesFromWorkbook(decodeText(book)) : [];
  if (styles) files.set("xl/styles.xml", encodeText(patchStylesXml(decodeText(styles))));
  if (book) files.set("xl/workbook.xml", encodeText(patchWorkbookXml(decodeText(book))));
  for (const [name, data] of files) {
    if (!name.startsWith("xl/worksheets/sheet") || !name.endsWith(".xml")) continue;
    const text = decodeText(data);
    if (!text.includes("[1]")) continue;
    files.set(name, encodeText(stripInternalWorkbookIndex(text, sheetNames)));
  }
  return zipEntries(files);
}

export async function readZipText(buffer, name) {
  const files = await unzipEntries(buffer);
  const data = files.get(name);
  return data ? decodeText(data) : null;
}
