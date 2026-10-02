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
  buildPack,
  inspectPack,
  packOutputFilename,
  excelSerialFromYymmdd,
  assertPlainNumberFormats,
  isPackLottery,
  readZipText,
} = await import("./pack-builder.js");

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

function formulaOf(cell) {
  const v = cell?.value;
  if (v && typeof v === "object" && v.formula) return `=${v.formula}`;
  if (typeof v === "string") return v;
  return v;
}

const onedrive = path.join(process.env.USERPROFILE, "OneDrive - Instant Win Gaming Ltd", "PPS Excels");
const outDir = path.join(root, "test-out");
await mkdir(outDir, { recursive: true });

assert(isPackLottery("NC") && isPackLottery("ga") && isPackLottery("PA") && isPackLottery("NH"), "pack lottery ids");
assert(!isPackLottery("KY"), "KY is not a pack lottery");
assert(excelSerialFromYymmdd("261002") === 46297, `261002 serial ${excelSerialFromYymmdd("261002")}`);
assert(
  packOutputFilename("260626_NC_PiggyBreaker_PPS_086.xlsx", "NC") ===
    "260626_NC_PiggyBreaker(Delivery)_PPS_086.xlsx",
  "NC output name",
);
assert(
  packOutputFilename("260821_NC_VIPElite(SSJ)_PPS_085_004.xlsx", "NC") ===
    "260821_NC_VIPElite(SSJ)(Delivery)_PPS_085_004.xlsx",
  "NC JP output name",
);
assert(
  packOutputFilename("260428_GA_LibertyLuck_PPS_078.xlsx", "PA") ===
    "260428_PA_LibertyLuck(Delivery)_PPS_078.xlsx",
  "PA output name uses selected lottery",
);

async function loadBuilt(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

const piggyPath = path.join(onedrive, "NC", "260626_NC_PiggyBreaker_PPS_086.xlsx");
const piggyBuf = await readFile(piggyPath);
const piggyInfo = await inspectPack(piggyBuf, path.basename(piggyPath), "NC");
assert(piggyInfo.layout === "NC", `Piggy layout ${piggyInfo.layout}`);
assert(!piggyInfo.hasJp, "Piggy should not be JP");
assert(!piggyInfo.buy, "Piggy is Standard");
const piggyBuilt = await buildPack(piggyBuf, path.basename(piggyPath), { lottery: "NC" });
assert(piggyBuilt.report.winningTiers === piggyInfo.winningTiers, "Piggy inspect/build tier mismatch");
const piggyWb = await loadBuilt(piggyBuilt.buffer);
const piggyD = piggyWb.getWorksheet("Delivery");
assert(piggyD, "missing NC Delivery");
assertPlainNumberFormats(piggyD, "PiggyBreaker");
assert(formulaOf(piggyD.getCell("E4")) === `=F4-SUM(E5:E${4 + piggyInfo.winningTiers})`, `Piggy E4 ${formulaOf(piggyD.getCell("E4"))}`);
assert(piggyD.getCell("F4").value === 2000000, `Piggy F4 ${piggyD.getCell("F4").value}`);
assert(String(formulaOf(piggyD.getCell("I4"))).includes("L$9"), `Piggy I4 ${formulaOf(piggyD.getCell("I4"))}`);
assert(piggyD.getCell("K15").value === "Game info for NC", `Piggy game info ${piggyD.getCell("K15").value}`);
assert(String(formulaOf(piggyD.getCell("L8"))) === "=L6*L7", `Piggy revenue ${formulaOf(piggyD.getCell("L8"))}`);
assert(String(formulaOf(piggyD.getCell("M11"))).includes("okay"), `Piggy flag ${formulaOf(piggyD.getCell("M11"))}`);
assert(piggyD.getCell("A1").numFmt === "yymmdd", `Piggy A1 fmt ${piggyD.getCell("A1").numFmt}`);
assert(String(piggyD.getCell("C5").value || "").length > 0, "Piggy first method empty");
assert(!String(piggyD.getCell("C5").value).includes(" x 1"), "Piggy method gained x 1");
const piggyTab = piggyD.properties?.tabColor;
assert(piggyTab?.theme === 5, `Piggy tab color ${JSON.stringify(piggyTab)}`);
await writeFile(path.join(outDir, packOutputFilename(path.basename(piggyPath), "NC")), Buffer.from(piggyBuilt.buffer));
console.log("piggy", piggyInfo.winningTiers, "fund", piggyBuilt.report.prizeFund, "rtp", piggyBuilt.report.actualRtp);

const vipPath = path.join(onedrive, "NC", "260821_NC_VIPElite(SSJ)_PPS_085_004.xlsx");
const vipBuf = await readFile(vipPath);
const vipInfo = await inspectPack(vipBuf, path.basename(vipPath), "NC");
assert(vipInfo.hasJp, "VIPElite should be JP");
const vipBuilt = await buildPack(vipBuf, path.basename(vipPath), { lottery: "NC" });
const vipWb = await loadBuilt(vipBuilt.buffer);
const vipD = vipWb.getWorksheet("Delivery");
assertPlainNumberFormats(vipD, "VIPElite");
assert(vipD.getCell("K12").value === "JP RTP:", `VIPElite K12 ${vipD.getCell("K12").value}`);
const vipJp = vipBuilt.report.delivery.jpStart;
assert(vipD.getCell(vipJp, 2).value === "Base JP Odds Down", `VIPElite JP start ${vipD.getCell(vipJp, 2).value}`);
assert(formulaOf(vipD.getCell("L12")) === `=C${vipJp + 28}+C${vipJp + 29}`, `VIPElite L12 ${formulaOf(vipD.getCell("L12"))}`);
assert(!String(vipD.getCell(vipJp, 5).value || "").toLowerCase().includes("json"), "VIPElite copied json setup");
assert(String(formulaOf(vipD.getCell(vipJp + 24, 3)) || "").includes("$L$15"), `VIPElite hit link ${formulaOf(vipD.getCell(vipJp + 24, 3))}`);
assert(vipD.getCell("K13").value === "Total RTP:", `VIPElite K13 ${vipD.getCell("K13").value}`);
assert(vipD.getCell("K17").value === "Game info for NC", `VIPElite game info ${vipD.getCell("K17").value}`);
await writeFile(path.join(outDir, packOutputFilename(path.basename(vipPath), "NC")), Buffer.from(vipBuilt.buffer));
console.log("vipelite", vipInfo.winningTiers, "jp", true);

const libPath = path.join(onedrive, "GA", "260428_GA_LibertyLuck_PPS_078.xlsx");
const libBuf = await readFile(libPath);
const libInfo = await inspectPack(libBuf, path.basename(libPath), "GA");
assert(libInfo.layout === "GA", "Liberty layout");
assert(!libInfo.hasJp, "Liberty should not be JP");
const libBuilt = await buildPack(libBuf, path.basename(libPath), { lottery: "GA" });
const libWb = await loadBuilt(libBuilt.buffer);
const libD = libWb.getWorksheet("Delivery");
assertPlainNumberFormats(libD, "LibertyLuck");
assert(libD.getCell("E2").value === "order", `Liberty E2 ${libD.getCell("E2").value}`);
assert(libD.getCell("E3").value === 0, "Liberty non-win order");
assert(libD.getCell("E4").value === 1, "Liberty first win order");
const libLast = 3 + libInfo.winningTiers;
assert(String(formulaOf(libD.getCell("L3"))) === `=B5-SUM(L4:L${libLast})`, `Liberty L3 ${formulaOf(libD.getCell("L3"))}`);
assert(String(formulaOf(libD.getCell("B5"))) === "=M3", `Liberty B5 ${formulaOf(libD.getCell("B5"))}`);
assert(String(formulaOf(libD.getCell("B7"))) === "=B5*B6", `Liberty B7 ${formulaOf(libD.getCell("B7"))}`);
assert(String(formulaOf(libD.getCell(`P${libLast + 1}`))) === `=SUM(P3:P${libLast})`, `Liberty P total ${formulaOf(libD.getCell(`P${libLast + 1}`))}`);
assert(String(libD.getCell("A14").value).startsWith("Bonus"), `Liberty Bonus ${libD.getCell("A14").value}`);
assert(String(libD.getCell("A15").value).startsWith("Freeplay"), `Liberty Freeplay ${libD.getCell("A15").value}`);
assert(String(libD.getCell("A22").value || "").includes("strictly confidential"), "Liberty confidential");
const p4 = libD.getCell("G4").value;
const p5 = libD.getCell("G5").value;
assert(typeof p4 === "number" && typeof p5 === "number" && p4 <= p5, `Liberty prizes not sorted ${p4} ${p5}`);
assert(libD.getCell("E2").border?.bottom?.style === "medium", "Liberty header underline");
assert(libD.getCell(libLast, 5).border?.bottom?.style === "medium", "Liberty last-win underline");
assert(libD.conditionalFormattings?.length || libD._conditionalFormattings?.length, "Liberty missing heat scale");
await writeFile(path.join(outDir, packOutputFilename(path.basename(libPath), "GA")), Buffer.from(libBuilt.buffer));
console.log("liberty", libInfo.winningTiers, "first prize", p4);

const xmasPath = path.join(onedrive, "GA", "260917_GA_ChristmasCash(MMJ)_PPS_078_004.xlsx");
const xmasBuf = await readFile(xmasPath);
const xmasInfo = await inspectPack(xmasBuf, path.basename(xmasPath), "GA");
assert(xmasInfo.hasJp, "ChristmasCash should be JP");
const xmasBuilt = await buildPack(xmasBuf, path.basename(xmasPath), { lottery: "GA" });
const xmasWb = await loadBuilt(xmasBuilt.buffer);
const xmasD = xmasWb.getWorksheet("Delivery");
assertPlainNumberFormats(xmasD, "ChristmasCash");
assert(xmasD.getCell("E2").value === "order", "ChristmasCash table starts at E");
assert(xmasD.getCell("A10").value === "JP Seed  RTP:", `Xmas A10 ${xmasD.getCell("A10").value}`);
const xmasJp = xmasBuilt.report.delivery.jpStart;
assert(xmasD.getCell(xmasJp, 2).value === "Base JP Odds Down", `Xmas JP start ${xmasD.getCell(xmasJp, 2).value}`);
assert(formulaOf(xmasD.getCell("B10")) === `=C${xmasJp + 12}`, `Xmas B10 ${formulaOf(xmasD.getCell("B10"))}`);
assert(formulaOf(xmasD.getCell("B11")) === `=C${xmasJp + 19}`, `Xmas B11 ${formulaOf(xmasD.getCell("B11"))}`);
assert(!String(xmasD.getCell(xmasJp, 5).value || "").toLowerCase().includes("json"), "Xmas copied json setup");
assert(xmasD.getCell("A13").value === "Total Game RTP:", `Xmas A13 ${xmasD.getCell("A13").value}`);
assert(String(xmasD.getCell("A18").value).startsWith("Bonus"), `Xmas Bonus ${xmasD.getCell("A18").value}`);
assert(String(xmasD.getCell("A26").value || "").includes("strictly confidential"), "Xmas confidential shifted");
await writeFile(path.join(outDir, packOutputFilename(path.basename(xmasPath), "GA")), Buffer.from(xmasBuilt.buffer));
console.log("christmas", xmasInfo.winningTiers, "jp rows ok");

const paBuilt = await buildPack(libBuf, path.basename(libPath), { lottery: "PA" });
const paWb = await loadBuilt(paBuilt.buffer);
const paD = paWb.getWorksheet("Delivery");
assert(paD.getCell("E2").value === "order", "PA uses GA layout");
assert(paBuilt.report.layout === "GA", "PA layout flag");
assert(
  packOutputFilename(path.basename(libPath), "NH") === "260428_NH_LibertyLuck(Delivery)_PPS_078.xlsx",
  "NH filename",
);
console.log("pa/nh layout ok");

async function assertSafePackage(buffer, label) {
  const book = await readZipText(buffer, "xl/workbook.xml");
  const styles = await readZipText(buffer, "xl/styles.xml");
  const calc = book.match(/<calcPr\b[^>]*\/>/);
  assert(calc, `${label} missing calcPr`);
  assert(/calcMode="auto"/.test(calc[0]), `${label} calcMode ${calc[0]}`);
  assert(/calcOnSave="1"/.test(calc[0]), `${label} calcOnSave ${calc[0]}`);
  assert(!/fullCalcOnLoad/.test(calc[0]), `${label} fullCalcOnLoad ${calc[0]}`);
  const bad = [...styles.matchAll(/formatCode="([^"]*)"/g)].map((m) => m[1]).filter((c) => /[_*]/.test(c));
  assert(bad.length === 0, `${label} accounting formats ${bad.slice(0, 4)}`);
}

await assertSafePackage(piggyBuilt.buffer, "PiggyBreaker");
await assertSafePackage(libBuilt.buffer, "LibertyLuck");
await assertSafePackage(xmasBuilt.buffer, "ChristmasCash");
console.log("pack ok");
