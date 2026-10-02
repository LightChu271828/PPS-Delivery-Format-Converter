/**
 * NC / GA / PA / NH Delivery builder (browser + Node).
 * Ports .cursor/skills/pps-delivery-conversion. One Delivery sheet, not KY four-tab.
 * PA and NH use the Georgia layout. Number formats are plain: no _ or * padding.
 */

const ExcelJS = globalThis.ExcelJS;

export const DELIVERY_TAB_COLOR = { theme: 5, tint: 0.7999816888943144 };

export const PACK_LOTTERIES = {
  NC: { code: "NC", layout: "NC", name: "NC Education Lottery" },
  GA: { code: "GA", layout: "GA", name: "Georgia Lottery" },
  PA: { code: "PA", layout: "GA", name: "Pennsylvania Lottery" },
  NH: { code: "NH", layout: "GA", name: "New Hampshire Lottery" },
};

export const PACK_SHEET = "Delivery";

const FMT_INT = "#,##0";
const FMT_MONEY = "#,##0.00";
const FMT_PCT = "0.00%";
const FMT_HIT = "0.00";
const FMT_DATE_NC = "yymmdd";
const FMT_DATE_GA = "[$-409]mmmm d, yyyy";

const CONFIDENTIAL =
  "The information contained in this document and all attached documents is strictly confidential " +
  "and contains proprietary information.\n" +
  "It is provided solely for use by the designated recipients and is subject to the terms of any " +
  "confidentiality obligations or non-disclosure agreements between the parties.\n" +
  "All other use is strictly prohibited.";

const NC_WIDTHS = {
  1: 13,
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

const MEDIUM_BOTTOM = {
  bottom: { style: "medium", color: { indexed: 64 } },
};

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

function insertSheet(wb, name) {
  const existing = sheet(wb, name);
  if (existing) wb.removeWorksheet(existing.id);
  return wb.addWorksheet(name);
}

function labelText(freq, row) {
  return String(cellResult(freq.getCell(row, 13)) || "").trim();
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

function orderPackSheets(wb) {
  const byName = Object.fromEntries(wb.worksheets.map((ws) => [ws.name, ws]));
  const head = ["Frequency", "Progressive Jackpots", PACK_SHEET].map((n) => byName[n]).filter(Boolean);
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
  const jpFormula =
    formulaAtLabel(freq, (l) => startsWithCi(l, "jp rtp")) ||
    "='Progressive Jackpots'!C30+'Progressive Jackpots'!C31";
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
    jpFormula,
  };
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
    gameInfo: 15 + insert,
  };
}

function applyNcNumberFormats(ws, last, tot, buy, hasJp) {
  const sum = ncSummaryRows(buy, hasJp);
  for (let row = 4; row <= last; row += 1) {
    ws.getCell(`D${row}`).numFmt = FMT_MONEY;
    ws.getCell(`E${row}`).numFmt = FMT_INT;
    ws.getCell(`F${row}`).numFmt = FMT_INT;
    ws.getCell(`G${row}`).numFmt = FMT_MONEY;
    ws.getCell(`H${row}`).numFmt = FMT_MONEY;
    ws.getCell(`I${row}`).numFmt = FMT_PCT;
  }
  ws.getCell(`E${tot}`).numFmt = FMT_INT;
  ws.getCell(`G${tot}`).numFmt = FMT_MONEY;
  ws.getCell(`I${tot}`).numFmt = FMT_PCT;
  ws.getCell("L6").numFmt = FMT_INT;
  ws.getCell("L7").numFmt = "0";
  if (buy) ws.getCell("L8").numFmt = "0";
  ws.getCell(`L${sum.revenue}`).numFmt = FMT_MONEY;
  ws.getCell(`L${sum.fund}`).numFmt = FMT_MONEY;
  ws.getCell(`L${sum.rtpSet}`).numFmt = FMT_PCT;
  ws.getCell(`L${sum.actual}`).numFmt = FMT_PCT;
  if (sum.jp) ws.getCell(`L${sum.jp}`).numFmt = FMT_PCT;
  if (sum.totalRtp) ws.getCell(`L${sum.totalRtp}`).numFmt = FMT_PCT;
  ws.getCell(`L${sum.winFreq}`).numFmt = FMT_INT;
  ws.getCell(`L${sum.hit}`).numFmt = FMT_HIT;
  ws.getCell("A1").numFmt = FMT_DATE_NC;
}

function buildNcDelivery(ws, freq, meta, source, dateSerial) {
  const { nonwin, rows } = source;
  const last = 4 + rows.length;
  const tot = last + 1;
  const buy = meta.buy;
  const hasJp = meta.hasJp;
  const sum = ncSummaryRows(buy, hasJp);
  const fundCell = `L$${sum.fund}`;
  const poolCell = "L6";
  const revenueCell = `L${sum.revenue}`;
  const rng = `$C$5:$C$${last}`;
  const drng = `$D$5:$D$${last}`;
  const erng = `$E$5:$E$${last}`;
  const grng = `$G$5:$G$${last}`;
  const helv8 = { name: "Helv", size: 8 };
  const helv8b = { name: "Helv", size: 8, bold: true };
  const helv10 = { name: "Helv", size: 10 };

  setColWidths(ws, NC_WIDTHS);
  setValue(ws, "A1", dateSerial).numFmt = FMT_DATE_NC;
  ws.getCell("A1").font = { ...helv8 };

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
  for (const addr of ["B2", "B3", "C3", "D3", "E2", "E3", "F2", "F3", "G2", "G3", "H3", "I2", "I3"]) {
    ws.getCell(addr).font = { ...helv8b };
  }

  setValue(ws, "B4", 0);
  setValue(ws, "C4", nonwin.method == null ? 0 : nonwin.method);
  applyMethodFill(ws.getCell("C4"), nonwin.fill, helv10);
  setValue(ws, "D4", cleanFloat(nonwin.prize));
  setValue(ws, "E4", `=F4-SUM(E5:E${last})`);
  setValue(ws, "F4", meta.pool);
  setValue(ws, "G4", "=E4*D4");
  setValue(ws, "H4", "=IFERROR(F4/E4,0)");
  setValue(ws, "I4", `=IFERROR(G4/${fundCell},0)`);

  rows.forEach((item, idx) => {
    const row = 5 + idx;
    setValue(ws, `B${row}`, item.tier);
    setValue(ws, `C${row}`, item.method);
    applyMethodFill(ws.getCell(`C${row}`), item.fill, helv10);
    setValue(ws, `D${row}`, item.prize);
    setValue(ws, `E${row}`, item.winners);
    setValue(ws, `F${row}`, meta.pool);
    setValue(ws, `G${row}`, `=E${row}*D${row}`);
    setValue(ws, `H${row}`, `=IFERROR(F${row}/E${row},0)`);
    setValue(ws, `I${row}`, `=G${row}/${fundCell}`);
  });

  setValue(ws, `E${tot}`, `=SUM(E4:E${last})`);
  setValue(ws, `G${tot}`, `=SUM(G4:G${last})`);
  setValue(ws, `I${tot}`, `=G${tot}/${fundCell}`);

  setValue(ws, "K2", "Prize Structure").font = { ...helv8b };
  setValue(ws, "K3", meta.jurisdiction).font = { ...helv8b };
  setValue(ws, "K4", meta.title).font = { ...helv8b };
  const kLabels = [
    [6, "Odds Down (ticket quantity if pool based):"],
    [7, "Base:"],
  ];
  if (buy) kLabels.push([8, "RRP(x):"]);
  kLabels.push(
    [sum.revenue, "Revenus (if pool based):"],
    [sum.fund, "Prize Fund (if pool based):"],
    [sum.rtpSet, "RTP Setting:"],
    [sum.actual, "Actual RTP:"],
  );
  if (hasJp) {
    kLabels.push([sum.jp, "JP RTP:"], [sum.totalRtp, "Total RTP:"]);
  }
  kLabels.push([sum.winFreq, "Winning Tiers Freq"], [sum.hit, "Hit Rate:"]);
  for (const [row, text] of kLabels) {
    setValue(ws, `K${row}`, text).font = { ...helv8b };
  }

  setValue(ws, "L6", meta.pool);
  setValue(ws, "L7", meta.base);
  if (buy) {
    setValue(ws, "L8", meta.rrp);
    setValue(ws, "L9", "=L6*L7*L8");
  } else {
    setValue(ws, "L8", "=L6*L7");
  }
  setValue(ws, `L${sum.fund}`, `=SUM(G5:G${last})`);
  setValue(ws, `L${sum.rtpSet}`, meta.rtp);
  setValue(ws, `L${sum.actual}`, `=L${sum.fund}/L${sum.revenue}`);
  setValue(
    ws,
    `M${sum.actual}`,
    `=IF(ROUND(L${sum.rtpSet},6)=ROUND(L${sum.actual},6),"okay","error")`,
  );
  if (hasJp) {
    setValue(ws, `L${sum.jp}`, meta.jpFormula);
    setValue(ws, `L${sum.totalRtp}`, `=L${sum.jp}+L${sum.actual}`);
  }
  setValue(ws, `L${sum.winFreq}`, `=SUM(E5:E${last})`);
  setValue(ws, `L${sum.hit}`, `=${poolCell}/L${sum.hit - 1}`);

  const info = sum.gameInfo;
  setValue(ws, `K${info}`, `Game info for ${meta.spec.code}`).font = { ...helv8b };
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
  }
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
    ws.getCell(`L${row}`).numFmt = row === info + 6 ? FMT_MONEY : FMT_HIT;
  }
  for (let row = info + 7; row <= info + 10; row += 1) {
    ws.getCell(`L${row}`).numFmt = FMT_PCT;
  }

  for (const addr of ["L6", "L7", `L${sum.revenue}`, `L${sum.fund}`, `L${sum.rtpSet}`, `L${sum.actual}`, `L${sum.winFreq}`, `L${sum.hit}`]) {
    ws.getCell(addr).font = { ...helv8 };
  }
  if (buy) ws.getCell("L8").font = { ...helv8 };
  if (hasJp) {
    ws.getCell(`L${sum.jp}`).font = { ...helv8 };
    ws.getCell(`L${sum.totalRtp}`).font = { ...helv8 };
  }
  for (let row = 4; row <= last; row += 1) {
    for (const col of ["B", "D", "E", "F", "G", "H", "I"]) {
      const cell = ws.getCell(`${col}${row}`);
      cell.font = { ...helv8 };
    }
  }
  applyNcNumberFormats(ws, last, tot, buy, hasJp);
  const confRow = info + 14;
  applyConfidential(ws, `K${confRow}:N${confRow + 4}`, "Helv", 8);
  ws.getRow(confRow).height = 18;
  return { last, tot, fundCell, info };
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
  setValue(ws, "A2", meta.title).font = { ...calibriB };
  setValue(ws, "A3", dateSerial).numFmt = FMT_DATE_GA;
  ws.getCell("A3").font = { ...calibri };

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
  const meta = lookupMeta(freq, code);
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
  };
}

export async function buildPack(buffer, filename, options = {}) {
  if (!ExcelJS) throw new Error("ExcelJS is not loaded.");
  const code = String(options.lottery || "NC").toUpperCase();
  const spec = PACK_LOTTERIES[code];
  if (!spec) throw new Error(`Unsupported pack lottery: ${options.lottery}`);
  const wb = await loadWorkbook(buffer);
  const freq = requireSheet(wb, "Frequency");
  const meta = lookupMeta(freq, code);
  assertPool(meta);
  const source = frequencyWinningRows(freq);
  const dateSerial = packDateSerial(filename);
  const ws = insertSheet(wb, PACK_SHEET);
  const built =
    spec.layout === "NC"
      ? buildNcDelivery(ws, freq, meta, source, dateSerial)
      : buildGaDelivery(ws, freq, meta, source, dateSerial);
  applyTabColor(ws);
  orderPackSheets(wb);
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

export async function patchPackXlsx(buffer) {
  const files = await unzipEntries(buffer);
  const styles = files.get("xl/styles.xml");
  const book = files.get("xl/workbook.xml");
  if (styles) files.set("xl/styles.xml", encodeText(patchStylesXml(decodeText(styles))));
  if (book) files.set("xl/workbook.xml", encodeText(patchWorkbookXml(decodeText(book))));
  return zipEntries(files);
}

export async function readZipText(buffer, name) {
  const files = await unzipEntries(buffer);
  const data = files.get(name);
  return data ? decodeText(data) : null;
}
