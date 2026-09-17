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
console.log("ok");
