import { createRequire } from "node:module";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const ExcelJS = require("./vendor/exceljs.min.js");
if (!ExcelJS?.Workbook) throw new Error("ExcelJS did not expose Workbook");
globalThis.ExcelJS = ExcelJS;

const { buildVa, inspectVaGame, vaOutputFilename, vaRole } = await import("./va-builder.js");
const { assertPlainNumberFormats, readZipText } = await import("./pack-builder.js");

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

function formulaOf(cell) {
  const value = cell?.value;
  if (value && typeof value === "object" && value.formula) return `=${value.formula}`;
  return value;
}

assert(vaRole("261006_VA_YearOfFireGoatSupremeJP(MainGame)_PPS_084_003.xlsx") === "main", "main role");
assert(vaRole("anything(SideBet).xlsx") === "side", "side role");
assert(
  vaOutputFilename("261006_VA_YearOfFireGoatSupremeJP(MainGame)_PPS_084_003.xlsx") ===
    "261006_VA_YearOfFireGoatSupremeJP(Delivery)_PPS_084_003.xlsx",
  "VA output name",
);
assert(
  vaOutputFilename("261006_VA_YearOfFireGoatSupremeJP(SideBet)_PPS_084_003.xlsx") ===
    "261006_VA_YearOfFireGoatSupremeJP(SideBet)(Delivery)_PPS_084_003.xlsx",
  "VA side output name",
);

const onedrive = path.join(process.env.USERPROFILE, "OneDrive - Instant Win Gaming Ltd", "PPS Excels", "VA");
const goat = path.join(onedrive, "261006 Year of Fire Goat(SupremeJP)");
const mainName = "261006_VA_YearOfFireGoatSupremeJP(MainGame)_PPS_084_003.xlsx";
const sideName = "261006_VA_YearOfFireGoatSupremeJP(SideBet)_PPS_084_003.xlsx";
const mainBuf = await readFile(path.join(goat, mainName));
const sideBuf = await readFile(path.join(goat, sideName));
const template = await readFile(path.join(root, "assets/templates/va-delivery.xlsx"));

const mainInfo = await inspectVaGame(mainBuf, mainName);
const sideInfo = await inspectVaGame(sideBuf, sideName);
assert(mainInfo.tiers === 203, `main tiers ${mainInfo.tiers}`);
assert(sideInfo.tiers === 55, `side tiers ${sideInfo.tiers}`);
assert(mainInfo.hits.length === 5, `hits ${mainInfo.hits.length}`);
assert(mainInfo.zeroFrequency === 4, `zeros ${mainInfo.zeroFrequency}`);
assert(mainInfo.prizeFund > 0 && Math.abs(mainInfo.actualRtp - mainInfo.prizeFund / mainInfo.pool) < 1e-9, "main fund");
assert(mainInfo.sideHit && mainInfo.sideHit.value > 1, `side hit cache ${mainInfo.sideHit?.value}`);
assert(sideInfo.hitRate > 1 && sideInfo.hits.length === 0, `side hit ${sideInfo.hitRate}`);

const built = await buildVa(mainBuf, mainName, { template });
assert(built.report.sheet === "Main Game", built.report.sheet);
assert(built.report.jpStart === 226, `jp ${built.report.jpStart}`);
const sideBuilt = await buildVa(sideBuf, sideName, { template });
assert(sideBuilt.report.sheet === "Side Bet", sideBuilt.report.sheet);
assert(sideBuilt.report.jpStart === 78, `side jp ${sideBuilt.report.jpStart}`);
assert(sideBuilt.report.hits.length === 0, "side file has no extra hit rates");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(built.buffer);
assert(wb.worksheets.map((ws) => ws.name).join(",") === "Main Game", "sheet names");
const mg = wb.getWorksheet("Main Game");
const sideWb = new ExcelJS.Workbook();
await sideWb.xlsx.load(sideBuilt.buffer);
const sb = sideWb.getWorksheet("Side Bet");
assertPlainNumberFormats(mg, "VA main");
assertPlainNumberFormats(sb, "VA side");
assert(mg.getCell("B18").value === "NUMBER", mg.getCell("B18").value);
assert(mg.getCell("C18").value === "WIN METHOD", mg.getCell("C18").value);
assert(mg.getCell("B4").value === "Year of Fire Goat Supreme JP(Main game)", mg.getCell("B4").value);
assert(formulaOf(mg.getCell("C8")) === "=C6*C7", formulaOf(mg.getCell("C8")));
assert(formulaOf(mg.getCell("C9")) === "=SUM(G19:G222)", formulaOf(mg.getCell("C9")));
assert(formulaOf(mg.getCell("C12")) === "=C254+C255", formulaOf(mg.getCell("C12")));
assert(formulaOf(mg.getCell("E19")) === "=F19-SUM(E20:E222)", formulaOf(mg.getCell("E19")));
assert(mg.getCell("C222").value === "bonusB-10|||||||||", mg.getCell("C222").value);
assert(mg.getCell("E222").value === 0, "zero odds row kept");
assert(mg.getCell("E7").value === "Side bet Hit Rate", mg.getCell("E7").value);
assert(mg.getCell("F7").value === mainInfo.sideHit.value, mg.getCell("F7").value);
assert(String(formulaOf(mg.getCell("F9"))).includes("$F$7"), formulaOf(mg.getCell("F9")));
assert(String(formulaOf(mg.getCell("F9"))).includes("$C$15"), formulaOf(mg.getCell("F9")));
assert(!String(formulaOf(mg.getCell("F9"))).includes("'Side Bet'"), formulaOf(mg.getCell("F9")));
assert(String(mg.getCell("H9").value || "").includes("confidential"), "main confidential");
assert(mg.getCell("E13").value === "Hit Rate(+5 Side bets):", mg.getCell("E13").value);
assert(mg.getCell("C226").value === 20000000, mg.getCell("C226").value);
assert(mg.getCell("M235").value === 50, mg.getCell("M235").value);
assert(mg.getCell("N236").value === 100000, mg.getCell("N236").value);
assert(mg.getCell("B233").fill?.fgColor?.theme === 5, `seed fill ${JSON.stringify(mg.getCell("B233").fill)}`);
assert(mg.getCell("C226").fill?.fgColor?.theme === 7, `jp input fill ${JSON.stringify(mg.getCell("C226").fill)}`);
assert(mg.getCell("B233").border?.left?.style === "thin", "seed border");
assert(!mg.getCell("B232").value, "blank row before Seed");
assert(mg.getCell("B232").fill?.pattern !== "solid", "blank row has no fill");
assert(String(formulaOf(mg.getCell("C251"))).includes("$C$6"), formulaOf(mg.getCell("C251")));
assert(!String(formulaOf(mg.getCell("C251"))).includes("Frequency"), formulaOf(mg.getCell("C251")));
assert(formulaOf(sb.getCell("C12")) === "=C106+C107", formulaOf(sb.getCell("C12")));
assert(sb.getCell("C74").value === "bonusB-13|||||||||", sb.getCell("C74").value);
assert(!String(sb.getCell("E7").value || "").startsWith("Hit Rate"), "side has no hit-rate block");
assert(String(sb.getCell("E8").value || "").includes("confidential"), "side confidential");

const book = await readZipText(built.buffer, "xl/workbook.xml");
assert(book.includes('calcMode="auto"'), book.match(/<calcPr\b[^>]*>/)?.[0]);
assert(!book.includes('fullCalcOnLoad="1"') && !book.includes('fullCalcOnLoad="true"'), "full calc");
console.log("va ok", built.report.tiers, sideBuilt.report.tiers);
