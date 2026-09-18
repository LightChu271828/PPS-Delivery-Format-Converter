import { createRequire } from "node:module";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const ExcelJS = require("./vendor/exceljs.min.js");
if (!ExcelJS?.Workbook) throw new Error("ExcelJS did not expose Workbook");
globalThis.ExcelJS = ExcelJS;

const {
  buildPps,
  inspectPps,
  findJackpotType,
  detectSchema,
  DEFAULT_PRICE_GRID,
  DAILY_STREAK_PRICE_GRID,
  resolvePriceGrid,
} = await import("./builder.js");

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

assert(findJackpotType("260916_KY_XtheMoneySSJ($2)") === "SSJ", "SSJ from filename");
assert(findJackpotType("KingKongJackpot(MMJ3)Standard") === "MMJ3", "MMJ3 before MMJ");
assert(findJackpotType("Something(MMJ)_PPS") === "MMJ", "MMJ not MMJ3");
assert(findJackpotType("Game(ChatterJP)") === "ChatterJP", "ChatterJP casing");
assert(findJackpotType("ssj") === "SSJ", "detect is case-insensitive, id is SSJ");
assert(findJackpotType("mmj3") === "MMJ3", "mmj3 -> MMJ3");
assert(
  detectSchema({
    identity: "PremiumGold(SSJ) Standard",
    filename: "260707_KY_test(MMJ)_PPS_083_004.xlsx",
    jpEntries: [{ row: 11 }],
  }) === "MMJ",
  "filename MMJ beats leftover SSJ labels",
);
assert(detectSchema({ filename: "DailyStreakBooster", jpEntries: [] }) === "Daily Streak");
assert(detectSchema({ filename: "America250", jpEntries: [] }) === "No jackpot");
assert(JSON.stringify(resolvePriceGrid("MMJ")) === JSON.stringify(DEFAULT_PRICE_GRID), "MMJ default grid");
assert(
  JSON.stringify(resolvePriceGrid("Daily Streak")) === JSON.stringify(DAILY_STREAK_PRICE_GRID),
  "Daily Streak keeps $3 unless overridden",
);
assert(JSON.stringify(resolvePriceGrid("MMJ", [50, 1, 0.5])) === JSON.stringify([0.5, 1, 50]), "selected prices sort");
try {
  resolvePriceGrid("MMJ", [0.5, 2, 5]);
  throw new Error("expected missing $1 error");
} catch (err) {
  if (!/\$1 base column/.test(err.message)) throw err;
}
try {
  detectSchema({ filename: "mystery-jp.xlsx", jpEntries: [{ row: 11 }] });
  throw new Error("expected missing jackpot-type error");
} catch (err) {
  if (!/MMJ, MMJ3, SSJ, or ChatterJP/.test(err.message)) throw err;
}

const templates = {
  mmj3: await readFile(path.join(root, "assets/templates/mmj3.xlsx")),
  "no-jp": await readFile(path.join(root, "assets/templates/no-jp.xlsx")),
};

const source =
  process.argv[2] ||
  path.join(root, "..", "260915 X the Money(SSJ)", "260916_KY_XtheMoneySSJ($2)_PPS_083_004.xlsx");

const buf = await readFile(source);
const info = await inspectPps(buf, path.basename(source));
console.log("inspect", { schema: info.schema, priceGrid: info.priceGrid, jp: info.jpNames });
assert(info.schema === "SSJ", `expected SSJ, got ${info.schema}`);
assert(JSON.stringify(info.priceGrid) === JSON.stringify(DEFAULT_PRICE_GRID), "SSJ uses default 8-price grid");

const { buffer, report } = await buildPps(buf, path.basename(source), { templates });
console.log("report", { schema: report.schema, priceGrid: report.priceGrid, uniquePrizes: report.uniquePrizes });
assert(report.schema === "SSJ", `written schema ${report.schema}`);
const outDir = path.join(root, "test-out");
await mkdir(outDir, { recursive: true });
const outPath = path.join(outDir, path.basename(source).replace(/\.xlsx$/i, "_Delivery.xlsx"));
await writeFile(outPath, Buffer.from(buffer));
const check = new ExcelJS.Workbook();
await check.xlsx.load(buffer);
const names = check.worksheets.map((ws) => ws.name);
for (const need of ["Delivery", "Odds Table", "Summary(Delivery)", "Prize Breakdown"]) {
  if (!names.includes(need)) throw new Error(`missing ${need}`);
}
const odds = check.getWorksheet("Odds Table");
const prices = [];
for (let col = 2; col <= 16; col += 2) prices.push(odds.getCell(1, col).value);
assert(JSON.stringify(prices) === JSON.stringify(DEFAULT_PRICE_GRID), `odds prices ${prices}`);
const summary = check.getWorksheet("Summary(Delivery)");
const summaryPrices = [];
for (let col = 2; col <= 9; col += 1) summaryPrices.push(summary.getCell(1, col).value);
assert(JSON.stringify(summaryPrices) === JSON.stringify(DEFAULT_PRICE_GRID), `summary prices ${summaryPrices}`);
const delivery = check.getWorksheet("Delivery");
const lastWin = report.delivery.lastWinRow;
const c4 = delivery.getCell("C4").value;
const f4 = delivery.getCell("F4").value;
console.log("saved", outPath);
console.log("sheets", names);
console.log("Delivery C4", c4, "F4", f4, "lastWin", lastWin);
if (report.winningTiers < 10) throw new Error("too few winning tiers");
if (Math.abs(report.actualRtp - 0.83) > 0.02) {
  console.warn("RTP off cached independent sum", report.actualRtp);
}

const scrooge = path.join(
  process.env.USERPROFILE,
  "OneDrive - Instant Win Gaming Ltd",
  "PPS Excels",
  "KY",
  "260707_KY_Scrooge(MMJ)_PPS_083_004.xlsx",
);
const scroogeBuf = await readFile(scrooge);
const scroogeInfo = await inspectPps(scroogeBuf, path.basename(scrooge));
assert(scroogeInfo.schema === "MMJ", `Scrooge schema ${scroogeInfo.schema}`);
const scroogeBuilt = await buildPps(scroogeBuf, path.basename(scrooge), { templates });
const scroogeWb = new ExcelJS.Workbook();
await scroogeWb.xlsx.load(scroogeBuilt.buffer);
const scroogeDelivery = scroogeWb.getWorksheet("Delivery");
assert(scroogeDelivery.getCell("L17").value == null, `Delivery L17 should be empty, got ${scroogeDelivery.getCell("L17").value}`);
assert(scroogeDelivery.getCell("L23").value == null, "Delivery hit-rate check should be gone");
const jpStart = scroogeBuilt.report.delivery.jpStart;
assert(jpStart, "missing Delivery JP start");
for (let row = jpStart; row <= jpStart + 6; row += 1) {
  for (let col = 5; col <= 10; col += 1) {
    const v = scroogeDelivery.getCell(row, col).value;
    assert(v == null || v === "", `json setup leftover ${row},${col}: ${v}`);
  }
}
const hitRate = scroogeDelivery.getCell("M15").value;
assert(hitRate?.formula === "ROUND(M6/M14,2)", `hit rate formula ${JSON.stringify(hitRate)}`);
const scroogeSummary = scroogeWb.getWorksheet("Summary(Delivery)");
const jp1Ratio = scroogeSummary.getCell("C19").value;
const jp2Ratio = scroogeSummary.getCell("C20").value;
assert(jp1Ratio?.formula === "C17/C15", `JP1 ratio ${JSON.stringify(jp1Ratio)}`);
assert(jp2Ratio?.formula === "C18/C15", `JP2 ratio ${JSON.stringify(jp2Ratio)}`);
const jpBorder = scroogeSummary.getCell("B15").border;
assert(jpBorder?.top?.style || jpBorder?.bottom?.style || jpBorder?.left?.style, "Summary row 15 needs borders");
assert(scroogeSummary.getCell("B28").border?.bottom?.style, "Summary row 28 needs a bottom border");
const scroogePb = scroogeWb.getWorksheet("Prize Breakdown");
let bonusRow = null;
let freeplayRow = null;
for (let row = 3; row <= 400; row += 1) {
  const method = scroogePb.getCell(row, 3).value;
  if (!method) continue;
  if (!bonusRow && String(method).startsWith("Bonus")) bonusRow = row;
  if (!freeplayRow && String(method).toLowerCase().startsWith("freeplay")) freeplayRow = row;
  if (bonusRow && freeplayRow) break;
}
assert(scroogePb.getCell(3, 2).value === 1, `first tier ${scroogePb.getCell(3, 2).value}`);
assert(scroogePb.getCell(4, 2).value === 2, `second tier ${scroogePb.getCell(4, 2).value}`);
assert(scroogePb.getCell(bonusRow, 2).value === bonusRow - 2, `bonus tier id ${scroogePb.getCell(bonusRow, 2).value}`);
assert(scroogePb.getCell(bonusRow, 5).value === "Bonus", `Bonus symbol ${scroogePb.getCell(bonusRow, 5).value}`);
assert(scroogePb.getCell(bonusRow, 6).value === 3, `Bonus qty ${scroogePb.getCell(bonusRow, 6).value}`);
assert(freeplayRow, "missing FreePlay prize-breakdown row");
assert(scroogePb.getCell(freeplayRow, 5).value === "Free Plays", `FreePlay symbol ${scroogePb.getCell(freeplayRow, 5).value}`);
assert(scroogePb.getCell(freeplayRow, 6).value === 3, `FreePlay qty ${scroogePb.getCell(freeplayRow, 6).value}`);
const scroogeWm = scroogeWb.getWorksheet("Win Methods");
let freeplayFormulas = 0;
for (let row = 5; row <= 400; row += 1) {
  const method = scroogeWm.getCell(row, 3).value;
  if (!method) continue;
  const identity = String(method).split("|")[0].trim();
  if (!identity.toLowerCase().startsWith("freeplay")) continue;
  const cell = scroogeWm.getCell(row, 6).value;
  const formula = typeof cell === "object" ? String(cell.formula || "") : String(cell || "");
  assert(/SUMPRODUCT\(SUMIF\(/i.test(formula), `FreePlay F${row} should be SUMPRODUCT, got ${formula}`);
  assert(!formula.includes("{") && !/VLOOKUP/i.test(formula), `FreePlay F${row} still CSE/VLOOKUP: ${formula}`);
  freeplayFormulas += 1;
}
assert(freeplayFormulas > 0, "expected rewritten FreePlay formulas");
const scroogeOut = path.join(outDir, path.basename(scrooge).replace(/\.xlsx$/i, "_Delivery.xlsx"));
await writeFile(scroogeOut, Buffer.from(scroogeBuilt.buffer));
console.log("scrooge saved", scroogeOut, "bonus", bonusRow, "freeplay", freeplayRow);

const fiesta = path.join(
  process.env.USERPROFILE,
  "OneDrive - Instant Win Gaming Ltd",
  "PPS Excels",
  "KY",
  "260710_KY_FiestaPepperPayout(SSJ)_PPS_083_004.xlsx",
);
const fiestaBuf = await readFile(fiesta);
const fiestaInfo = await inspectPps(fiestaBuf, path.basename(fiesta));
assert(fiestaInfo.schema === "SSJ", `Fiesta schema ${fiestaInfo.schema}`);
assert(fiestaInfo.zeroFrequency?.length === 11, `Fiesta zero-frequency ${fiestaInfo.zeroFrequency?.length}`);
const fiestaBuilt = await buildPps(fiestaBuf, path.basename(fiesta), { templates });
assert(fiestaBuilt.report.winningTiers === fiestaInfo.winningTiers, "Fiesta inspect/build tier mismatch");
const fiestaOut = path.join(outDir, path.basename(fiesta).replace(/\.xlsx$/i, "_Delivery.xlsx"));
await writeFile(fiestaOut, Buffer.from(fiestaBuilt.buffer));
console.log("fiesta saved", fiestaOut, "tiers", fiestaBuilt.report.winningTiers, "skipped", fiestaInfo.zeroFrequency.length);
console.log("ok");
