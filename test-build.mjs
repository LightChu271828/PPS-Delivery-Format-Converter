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
  priceGridBaseCol,
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
assert(JSON.stringify(resolvePriceGrid("MMJ", [5, 10, 20, 30, 50])) === JSON.stringify([5, 10, 20, 30, 50]), "grid without $1");
assert(priceGridBaseCol(DEFAULT_PRICE_GRID) === 3, "$1 stays the base when present");
assert(priceGridBaseCol([5, 10, 20, 30, 50]) === 2, "first selected price is the base without $1");
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

const subset = [5, 10, 20, 30, 50];
const subsetBuilt = await buildPps(buf, path.basename(source), { templates, priceGrid: subset });
assert(JSON.stringify(subsetBuilt.report.priceGrid) === JSON.stringify(subset), "SSJ subset grid");
const subsetWb = new ExcelJS.Workbook();
await subsetWb.xlsx.load(subsetBuilt.buffer);
const subsetOdds = subsetWb.getWorksheet("Odds Table");
const subsetOddsPrices = [];
for (let col = 2; col <= 10; col += 2) subsetOddsPrices.push(subsetOdds.getCell(1, col).value);
assert(JSON.stringify(subsetOddsPrices) === JSON.stringify(subset), `SSJ subset odds ${subsetOddsPrices}`);
const subsetJpOdds = subsetOdds.getCell("C2").value;
assert(/Progressive Jackpots/.test(subsetJpOdds?.formula || ""), `JP odds ${JSON.stringify(subsetJpOdds)}`);

const fortune = path.join(
  process.env.USERPROFILE,
  "OneDrive - Instant Win Gaming Ltd",
  "PPS Excels",
  "KY",
  "260601_KY_FirstClassFortune_PPS_085 - test.xlsx",
);
const fortuneBuf = await readFile(fortune);
const fortuneInfo = await inspectPps(fortuneBuf, path.basename(fortune));
assert(fortuneInfo.schema === "No jackpot", `Fortune schema ${fortuneInfo.schema}`);
const fortuneBuilt = await buildPps(fortuneBuf, path.basename(fortune), { templates, priceGrid: subset });
assert(JSON.stringify(fortuneBuilt.report.priceGrid) === JSON.stringify(subset), "Fortune subset grid");
const fortuneWb = new ExcelJS.Workbook();
await fortuneWb.xlsx.load(fortuneBuilt.buffer);
const fortuneOdds = fortuneWb.getWorksheet("Odds Table");
const fortunePrices = [];
for (let col = 2; col <= 6; col += 1) fortunePrices.push(fortuneOdds.getCell(1, col).value);
assert(JSON.stringify(fortunePrices) === JSON.stringify(subset), `Fortune odds prices ${fortunePrices}`);
const fortunePrize = fortuneOdds.getCell("B2").value;
assert(typeof fortunePrize === "number" && fortunePrize > 0, `Fortune $5 prize ${fortunePrize}`);
const fortuneScaled = fortuneOdds.getCell("C2").value;
assert(
  /B2\*C\$1\/B\$1/.test(fortuneScaled?.formula || ""),
  `Fortune scaled prize ${JSON.stringify(fortuneScaled)}`,
);
const fortuneSummary = fortuneWb.getWorksheet("Summary(Delivery)");
assert(fortuneSummary.getCell("B1").value === 5, "Fortune summary starts at $5");
assert(typeof fortuneSummary.getCell("B2").value === "number", `Fortune summary top prize ${fortuneSummary.getCell("B2").value}`);
const fortuneOut = path.join(outDir, path.basename(fortune).replace(/\.xlsx$/i, "_Delivery.xlsx"));
await writeFile(fortuneOut, Buffer.from(fortuneBuilt.buffer));
console.log("fortune saved", fortuneOut, "grid", fortuneBuilt.report.priceGrid);
console.log("ok");
