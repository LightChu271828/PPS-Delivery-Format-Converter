/**
 * Virginia Lottery delivery workbook.
 * Two uploaded PPS files become one workbook with Main Game and Side Bet sheets.
 * Layout follows Cash Vault Boost 2: Frequency in columns B–I from row 19, then the
 * Progressive Jackpots block. Styles come from assets/templates/va-delivery.xlsx.
 * Number formats stay plain: no accounting _ or * padding.
 */

const ExcelJS = globalThis.ExcelJS;

const FMT_INT = "#,##0";
const FMT_MONEY = "#,##0.00";
const FMT_PCT = "0.00%";

const TEMPLATE_ROWS = {
  "Main Game": { nonwin: 19, win: 20, total: 220, jp: 223, jpRows: 31 },
  "Side Bet": { nonwin: 19, win: 20, total: 76, jp: 79, jpRows: 31 },
};

const CONFIDENTIAL =
  "The information contained in this document and all attached documents is strictly confidential " +
  "and contains proprietary information.\n" +
  "It is provided solely for use by the designated recipients and is subject to the terms of any " +
  "confidentiality obligations or non-disclosure agreements between the parties.\n" +
  "All other use is strictly prohibited.";

function plainFormat(fmt) {
  if (!fmt || fmt === "General" || !/[_*]/.test(String(fmt))) return fmt || "General";
  const first = String(fmt).split(";")[0];
  if (first.includes("%")) return FMT_PCT;
  const match = first.match(/\.([0#]+)/);
  if (match) return `#,##0.${"0".repeat(match[1].length)}`;
  return FMT_INT;
}

function cellResult(cell) {
  if (!cell) return null;
  const value = cell.value;
  if (value == null || value === "") return null;
  if (typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map((part) => part.text ?? "").join("");
  if (value.richText) return value.richText.map((part) => part.text ?? "").join("");
  if (value.hyperlink) return value.text ?? value.hyperlink;
  if (value.error) return null;
  if ("formula" in value || "sharedFormula" in value) return value.result === undefined ? null : value.result;
  return value.result === undefined ? null : value.result;
}

function cellFormula(cell) {
  const value = cell?.value;
  if (value && typeof value === "object" && value.formula) return String(value.formula).replace(/^=/, "");
  return null;
}

function paint(src, dst, numFmt) {
  if (!src) return;
  const style = {};
  if (src.font) style.font = JSON.parse(JSON.stringify(src.font));
  const fill = src.fill;
  if (fill && fill.type === "pattern" && fill.pattern && String(fill.pattern).toLowerCase() !== "none") {
    style.fill = JSON.parse(JSON.stringify(fill));
  }
  const border = src.border;
  if (border && ["left", "right", "top", "bottom"].some((edge) => border[edge]?.style)) {
    style.border = JSON.parse(JSON.stringify(border));
  }
  if (src.alignment) style.alignment = JSON.parse(JSON.stringify(src.alignment));
  if (numFmt) style.numFmt = numFmt;
  else if (src.numFmt && src.numFmt !== "General") style.numFmt = plainFormat(src.numFmt);
  dst.style = style;
}

function copyHeight(srcWs, srcRow, dstWs, dstRow) {
  const height = srcWs.getRow(srcRow).height;
  if (height) dstWs.getRow(dstRow).height = height;
}

async function loadWorkbook(buffer) {
  if (!ExcelJS) throw new Error("ExcelJS is not loaded.");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

function labelAt(freq, row) {
  return String(cellResult(freq.getCell(row, 13)) ?? "").trim();
}

export function vaRole(filename, title = "") {
  const text = `${filename || ""} ${title || ""}`;
  if (/side\s*bet/i.test(text)) return "side";
  if (/main\s*game/i.test(text)) return "main";
  return null;
}

export function vaOutputFilename(name) {
  const raw = String(name || "VA_PPS.xlsx").replace(/\.xlsx$/i, "");
  if (/\(main\s*game\)/i.test(raw)) return `${raw.replace(/\(main\s*game\)/i, "(Delivery)")}.xlsx`;
  const match = raw.match(/^(.*)_PPS_(.+)$/i);
  if (match && !/\(Delivery\)/i.test(match[1])) return `${match[1]}(Delivery)_PPS_${match[2]}.xlsx`;
  if (/\(Delivery\)/i.test(raw)) return `${raw}.xlsx`;
  return `${raw}(Delivery).xlsx`;
}

function shiftFormula(formula, offset) {
  let body = String(formula || "").replace(/^=/, "");
  const tokens = [
    ["'Frequency'!$N$16", "__HIT__"],
    ["'Frequency'!$N$15", "__HIT__"],
    ["'Frequency'!$N$14", "__WINS__"],
    ["'Frequency'!$N$11", "__BASE__"],
    ["'Frequency'!$N$6", "__POOL__"],
    ["Frequency!$N$16", "__HIT__"],
    ["Frequency!$N$15", "__HIT__"],
    ["Frequency!$N$14", "__WINS__"],
    ["Frequency!$N$11", "__BASE__"],
    ["Frequency!$N$6", "__POOL__"],
    ["Frequency!N16", "__HIT__"],
    ["Frequency!N15", "__HIT__"],
    ["Frequency!N14", "__WINS__"],
    ["Frequency!N11", "__BASE__"],
    ["Frequency!N6", "__POOL__"],
  ];
  for (const [from, token] of tokens) body = body.replaceAll(from, token);
  body = body.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g, (_, colAbs, letters, rowAbs, row) => {
    return `${colAbs}${letters}${rowAbs}${Number(row) + offset}`;
  });
  return body
    .replaceAll("__POOL__", "$C$6")
    .replaceAll("__WINS__", "$C$14")
    .replaceAll("__HIT__", "$C$15")
    .replaceAll("__BASE__", "$C$11");
}

function readTiers(freq, filename) {
  let last = 0;
  const max = Math.min(freq.rowCount || 0, 2000);
  for (let row = 4; row <= max; row += 1) {
    const method = cellResult(freq.getCell(row, 3));
    if (method == null || method === "") continue;
    last = row;
  }
  if (last < 4) throw new Error(`${filename}: Frequency has no tier rows.`);
  const rows = [];
  for (let row = 4; row <= last; row += 1) {
    const method = cellResult(freq.getCell(row, 3));
    const prizeRaw = cellResult(freq.getCell(row, 4));
    const winners = cellResult(freq.getCell(row, 5));
    const number = cellResult(freq.getCell(row, 2));
    const prize = prizeRaw == null && row === 4 ? 0 : prizeRaw;
    if (method == null || method === "" || prize == null || (row > 4 && winners == null)) {
      throw new Error(
        `${filename}: Frequency row ${row} has no cached method, prize, or Odds up. Open it in Excel, calculate, and save.`,
      );
    }
    rows.push({
      number: row === 4 ? 0 : number ?? row - 4,
      method: String(method),
      prize: Number(prize),
      winners: row === 4 ? null : Number(winners),
    });
  }
  return rows;
}

function readHits(freq) {
  const hits = [];
  for (let row = 1; row <= 40; row += 1) {
    const label = labelAt(freq, row);
    if (!label.startsWith("Hit Rate(") || /main game/i.test(label)) continue;
    const formula = cellFormula(freq.getCell(row, 14));
    if (!formula) {
      throw new Error(`${label} needs a formula in Frequency N${row}.`);
    }
    hits.push({ label, formula });
  }
  return hits;
}

function readJackpot(pj) {
  if (!pj) return [];
  const cells = [];
  const max = Math.min(pj.rowCount || 0, 80);
  let found = false;
  for (let row = 1; row <= Math.min(max, 15); row += 1) {
    if (cellResult(pj.getCell(row, 2)) === "Base JP Odds Down") found = true;
  }
  if (!found) return [];
  for (let row = 2; row <= max; row += 1) {
    for (let col = 2; col <= 23; col += 1) {
      if (row <= 7 && col >= 5) continue;
      const cell = pj.getCell(row, col);
      const formula = cellFormula(cell);
      if (formula) {
        cells.push({ row, col, formula });
        continue;
      }
      const value = cellResult(cell);
      if (value == null || value === "") continue;
      cells.push({ row, col, value });
    }
  }
  return cells;
}

function readGame(wb, filename) {
  const freq = wb.getWorksheet("Frequency");
  if (!freq) throw new Error(`${filename}: Frequency sheet is missing.`);
  const rows = readTiers(freq, filename);
  const pool = cellResult(freq.getCell(6, 14));
  const base = cellResult(freq.getCell(7, 14));
  let rtp = null;
  for (let row = 1; row <= 40; row += 1) {
    const label = labelAt(freq, row);
    if (label.startsWith("RTP Setting")) rtp = cellResult(freq.getCell(row, 14));
  }
  if (!Number.isFinite(Number(pool)) || !Number.isFinite(Number(base))) {
    throw new Error(`${filename}: Frequency pool (N6) or base (N7) is missing.`);
  }
  if (!Number.isFinite(Number(rtp))) throw new Error(`${filename}: Frequency RTP Setting is missing.`);
  const winning = rows.slice(1);
  const wins = winning.reduce((sum, row) => sum + (Number(row.winners) || 0), 0);
  const prizeFund = rows.reduce((sum, row) => sum + (Number(row.prize) || 0) * (Number(row.winners) || 0), 0);
  const title = cellResult(freq.getCell(4, 13)) || filename;
  return {
    filename,
    title: String(title),
    role: vaRole(filename, title),
    pool: Number(pool),
    base: Number(base),
    rtp: Number(rtp),
    rows,
    hits: readHits(freq),
    jackpot: readJackpot(wb.getWorksheet("Progressive Jackpots")),
    tiers: winning.length,
    wins,
    prizeFund,
    hitRate: wins ? pool / wins : null,
    actualRtp: pool * base ? prizeFund / (pool * base) : null,
    zeroFrequency: winning.filter((row) => row.winners === 0).length,
  };
}

export async function inspectVaGame(buffer, filename = "upload.xlsx") {
  const wb = await loadWorkbook(buffer);
  return readGame(wb, filename);
}

function writeMeta(ws, templateWs, game, lastWin, total, jpStart) {
  const text = [
    [2, 2, "Prize Structure"],
    [3, 2, "Virginia Lottery"],
    [4, 2, game.title],
    [6, 2, "Odds Down (ticket quantity if pool based):"],
    [7, 2, "Base:"],
    [8, 2, "Revenus (if pool based):"],
    [9, 2, "Prize Fund (if pool based):"],
    [10, 2, "RTP Setting "],
    [11, 2, "Base RTP:"],
    [12, 2, "JP RTP:"],
    [13, 2, "Total RTP"],
    [14, 2, "Winning Tiers Freq:"],
    [15, 2, "Hit Rate:"],
  ];
  for (const [row, col, value] of text) {
    const cell = ws.getCell(row, col);
    paint(templateWs.getCell(row, col), cell);
    cell.value = value;
    copyHeight(templateWs, row, ws, row);
  }
  const values = [
    [6, 3, game.pool, FMT_INT],
    [7, 3, game.base, FMT_MONEY],
    [8, 3, { formula: "C6*C7" }, FMT_MONEY],
    [9, 3, { formula: `SUM(G19:G${lastWin})` }, FMT_MONEY],
    [10, 3, game.rtp, FMT_PCT],
    [11, 3, { formula: `G${total}/C8` }, FMT_PCT],
    [12, 3, game.jackpot.length ? { formula: `C${jpStart + 28}+C${jpStart + 29}` } : null, FMT_PCT],
    [13, 3, { formula: "C11+C12" }, FMT_PCT],
    [14, 3, { formula: `SUM(E20:E${lastWin})` }, FMT_INT],
    [15, 3, { formula: "C6/C14" }, FMT_MONEY],
  ];
  for (const [row, col, value, fmt] of values) {
    const cell = ws.getCell(row, col);
    paint(templateWs.getCell(row, col), cell, fmt);
    if (value != null) cell.value = value;
  }
}

function writeTable(ws, templateWs, spec, rows, lastWin, total) {
  const headers = [
    [17, 2, "TIER"],
    [17, 5, "ODDS"],
    [17, 6, "ODDS"],
    [17, 7, "PRIZE"],
    [17, 9, "% OF "],
    [18, 2, "NUMBER"],
    [18, 3, "WIN METHOD"],
    [18, 4, "PRIZE"],
    [18, 5, "UP"],
    [18, 6, "DOWN"],
    [18, 7, "COST"],
    [18, 8, "1 in X Odds"],
    [18, 9, "PRIZE FUND"],
  ];
  for (const [row, col, value] of headers) {
    const cell = ws.getCell(row, col);
    paint(templateWs.getCell(row, col), cell);
    cell.value = value;
  }
  copyHeight(templateWs, 17, ws, 17);
  copyHeight(templateWs, 18, ws, 18);
  rows.forEach((row, index) => {
    const dest = 19 + index;
    const srcRow = index === 0 ? spec.nonwin : spec.win;
    const formats = {
      2: "00",
      4: FMT_MONEY,
      5: FMT_INT,
      6: FMT_INT,
      7: FMT_MONEY,
      8: FMT_MONEY,
      9: FMT_PCT,
    };
    for (let col = 2; col <= 9; col += 1) {
      paint(templateWs.getCell(srcRow, col), ws.getCell(dest, col), formats[col]);
    }
    ws.getCell(dest, 2).value = row.number;
    ws.getCell(dest, 3).value = row.method;
    ws.getCell(dest, 4).value = row.prize;
    ws.getCell(dest, 5).value = index === 0 ? { formula: `F${dest}-SUM(E20:E${lastWin})` } : row.winners;
    ws.getCell(dest, 6).value = { formula: "$C$6" };
    ws.getCell(dest, 7).value = { formula: `E${dest}*D${dest}` };
    ws.getCell(dest, 8).value = { formula: `F${dest}/E${dest}` };
    if (index) ws.getCell(dest, 9).value = { formula: `G${dest}/$C$9` };
    if (index === 0) copyHeight(templateWs, srcRow, ws, dest);
  });
  for (const [col, fmt, letter] of [
    [5, FMT_INT, "E"],
    [7, FMT_MONEY, "G"],
    [9, FMT_PCT, "I"],
  ]) {
    const cell = ws.getCell(total, col);
    paint(templateWs.getCell(spec.total, col), cell, fmt);
    cell.value = { formula: `SUM(${letter}19:${letter}${lastWin})` };
  }
}

function writeHits(ws, templateWs, hits) {
  hits.forEach((hit, index) => {
    const row = 9 + index;
    const label = ws.getCell(row, 5);
    paint(templateWs.getCell(9, 5), label);
    label.value = hit.label;
    let body = hit.formula.replace(/^=/, "");
    body = body.replaceAll("$N$16", "$C$15").replaceAll("$N$15", "'Side Bet'!$C$15");
    body = body.replaceAll("N$16", "$C$15").replaceAll("N$15", "'Side Bet'!$C$15");
    const value = ws.getCell(row, 6);
    paint(templateWs.getCell(9, 6), value, FMT_MONEY);
    value.value = { formula: body };
  });
}

function writeConfidential(ws, templateWs, range, anchor) {
  const [start, end] = range.split(":");
  ws.mergeCells(`${start}:${end}`);
  const cell = ws.getCell(start);
  paint(templateWs.getCell(anchor), cell);
  cell.value = CONFIDENTIAL;
  cell.alignment = { ...(cell.alignment || {}), horizontal: "center", vertical: "middle", wrapText: true };
}

function writeJackpot(ws, templateWs, spec, jackpot, jpStart) {
  for (let offset = 0; offset < spec.jpRows; offset += 1) {
    const srcRow = spec.jp + offset;
    const dstRow = jpStart + offset;
    copyHeight(templateWs, srcRow, ws, dstRow);
    for (let col = 2; col <= 23; col += 1) {
      paint(templateWs.getCell(srcRow, col), ws.getCell(dstRow, col));
    }
  }
  const shift = jpStart - 2;
  for (const entry of jackpot) {
    const cell = ws.getCell(entry.row + shift, entry.col);
    if (entry.formula) cell.value = { formula: shiftFormula(entry.formula, shift) };
    else cell.value = entry.value;
  }
}

function buildSheet(wb, templateWb, sheetName, game, withHits) {
  const templateWs = templateWb.getWorksheet(sheetName);
  if (!templateWs) throw new Error(`VA template is missing ${sheetName}.`);
  const spec = TEMPLATE_ROWS[sheetName];
  const ws = wb.addWorksheet(sheetName);
  for (let col = 1; col <= 23; col += 1) {
    const width = templateWs.getColumn(col).width;
    if (width) ws.getColumn(col).width = width;
  }
  const lastWin = 19 + game.rows.length - 1;
  const total = lastWin + 1;
  const jpStart = game.jackpot.length ? total + 3 : null;
  writeMeta(ws, templateWs, game, lastWin, total, jpStart || 0);
  writeTable(ws, templateWs, spec, game.rows, lastWin, total);
  if (withHits) writeHits(ws, templateWs, game.hits);
  if (jpStart) writeJackpot(ws, templateWs, spec, game.jackpot, jpStart);
  return { lastWin, total, jpStart, tiers: game.tiers };
}

export async function buildVa(mainBuffer, sideBuffer, mainName, sideName, options = {}) {
  if (!options.template) throw new Error("VA template is not loaded.");
  const templateWb = await loadWorkbook(options.template);
  const mainWb = await loadWorkbook(mainBuffer);
  const sideWb = await loadWorkbook(sideBuffer);
  const main = readGame(mainWb, mainName || "Main Game.xlsx");
  const side = readGame(sideWb, sideName || "Side Bet.xlsx");
  const out = new ExcelJS.Workbook();
  out.calcProperties = { ...(out.calcProperties || {}), calcMode: "auto", fullCalcOnLoad: false };
  const mainPlace = buildSheet(out, templateWb, "Main Game", main, true);
  const sidePlace = buildSheet(out, templateWb, "Side Bet", side, false);
  writeConfidential(out.getWorksheet("Main Game"), templateWb.getWorksheet("Main Game"), "H9:K13", "H9");
  writeConfidential(out.getWorksheet("Side Bet"), templateWb.getWorksheet("Side Bet"), "E8:H12", "E8");
  const raw = await out.xlsx.writeBuffer();
  const { patchPackXlsx } = await import("./pack-builder.js");
  const buffer = await patchPackXlsx(raw);
  return {
    buffer,
    report: {
      lottery: "VA",
      main: { ...main, ...mainPlace },
      side: { ...side, ...sidePlace },
    },
  };
}
