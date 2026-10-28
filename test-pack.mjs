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
  lnzSheetName,
} = await import("./pack-builder.js");

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}

function formulaOf(cell) {
  const v = cell?.value;
  if (v && typeof v === "object" && v.formula) return `=${v.formula}`;
  if (v && typeof v === "object" && v.sharedFormula && cell.formula) return `=${cell.formula}`;
  if (typeof v === "string") return v;
  return v;
}

const onedrive = path.join(process.env.USERPROFILE, "OneDrive - Instant Win Gaming Ltd", "PPS Excels");
const outDir = path.join(root, "test-out");
await mkdir(outDir, { recursive: true });

assert(
  ["NC", "ga", "PA", "NH", "VA", "DC", "LNZ"].every((id) => isPackLottery(id)),
  "pack lottery ids",
);
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
const piggyLast = 5 + piggyInfo.winningTiers;
const piggyTot = piggyLast + 1;
assert(piggyBuilt.report.delivery.tot === piggyTot, `Piggy total row ${piggyBuilt.report.delivery.tot}`);
assert(piggyD.getCell("A1").value == null, `Piggy A1 ${piggyD.getCell("A1").value}`);
assert(
  piggyD.getCell("C1").value === "Piggy Breaker Prize Structure for North Carolina Education Lottery",
  `Piggy title ${piggyD.getCell("C1").value}`,
);
assert(piggyD.getCell("C1").font?.name === "Aptos Narrow" && piggyD.getCell("C1").font?.size === 16 && piggyD.getCell("C1").font?.bold, `Piggy C1 font ${JSON.stringify(piggyD.getCell("C1").font)}`);
const ncHeads = {
  A3: "TIER", A4: "Number", B4: "Prize", C4: "Win Method", D4: "Odds", E4: "Probability",
  F3: "Expected", F4: "Return", G3: "Prize", G4: "Cost", H4: "1 in X Odds",
  I3: "Number in", I4: "Pool", J3: "% of", J4: "Prize Fund",
};
for (const [addr, text] of Object.entries(ncHeads)) {
  assert(piggyD.getCell(addr).value === text, `Piggy ${addr} ${piggyD.getCell(addr).value}`);
}
for (const col of ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]) {
  assert(piggyD.getCell(`${col}4`).border?.bottom?.style === "thin", `Piggy ${col}4 underline`);
  assert(piggyD.getCell(`${col}${piggyLast}`).border?.bottom?.style === "thin", `Piggy ${col}${piggyLast} underline`);
}
assert(piggyD.getCell("K4").value == null && piggyD.getCell("K2").value == null, "Piggy has old K column content");
assert(piggyD.getCell("A5").value === 0 && piggyD.getCell("A5").numFmt === "00", "Piggy non-win tier on row 5");
assert(formulaOf(piggyD.getCell("D5")) === `=I5-SUM(D6:D${piggyLast})`, `Piggy D5 ${formulaOf(piggyD.getCell("D5"))}`);
assert(piggyD.getCell("I5").value === 2000000, `Piggy I5 ${piggyD.getCell("I5").value}`);
assert(formulaOf(piggyD.getCell("J5")) === "=IFERROR(G5/M$8,0)", `Piggy J5 ${formulaOf(piggyD.getCell("J5"))}`);
assert(formulaOf(piggyD.getCell("E6")) === "=D6/I6", `Piggy E6 ${formulaOf(piggyD.getCell("E6"))}`);
assert(formulaOf(piggyD.getCell("F6")) === "=E6*B6", `Piggy F6 ${formulaOf(piggyD.getCell("F6"))}`);
assert(formulaOf(piggyD.getCell("G6")) === "=D6*B6", `Piggy G6 ${formulaOf(piggyD.getCell("G6"))}`);
assert(formulaOf(piggyD.getCell("H6")) === "=IFERROR(I6/D6,0)", `Piggy H6 ${formulaOf(piggyD.getCell("H6"))}`);
assert(formulaOf(piggyD.getCell("J6")) === "=G6/M$8", `Piggy J6 ${formulaOf(piggyD.getCell("J6"))}`);
assert(typeof piggyD.getCell("D6").value === "number", "Piggy D6 winners should be a value");
for (const col of ["D", "E", "F", "G"]) {
  assert(formulaOf(piggyD.getCell(`${col}${piggyTot}`)) === `=SUM(${col}5:${col}${piggyLast})`, `Piggy ${col}${piggyTot} total`);
}
const piggySummary = {
  L5: "Odds Down (ticket quantity if pool based):", L6: "Base:", L7: "Revenue:", L8: "Prize Fund:",
  L9: "RTP Setting:", L10: "Base RTP:", L11: "Winning Tiers Freq:", L12: "Hit Rate:", L15: "Game info for NC",
};
for (const [addr, text] of Object.entries(piggySummary)) {
  assert(piggyD.getCell(addr).value === text, `Piggy ${addr} ${piggyD.getCell(addr).value}`);
}
assert(piggyD.getCell("M5").value === 2000000, `Piggy M5 ${piggyD.getCell("M5").value}`);
assert(formulaOf(piggyD.getCell("M7")) === "=M5*M6", `Piggy revenue ${formulaOf(piggyD.getCell("M7"))}`);
assert(formulaOf(piggyD.getCell("M8")) === `=SUM(G6:G${piggyLast})`, `Piggy fund ${formulaOf(piggyD.getCell("M8"))}`);
assert(formulaOf(piggyD.getCell("M10")) === `=G${piggyTot}/M7`, `Piggy base RTP ${formulaOf(piggyD.getCell("M10"))}`);
assert(formulaOf(piggyD.getCell("M11")) === `=SUM(D6:D${piggyLast})`, `Piggy wins ${formulaOf(piggyD.getCell("M11"))}`);
assert(formulaOf(piggyD.getCell("M12")) === "=M5/M11", `Piggy hit ${formulaOf(piggyD.getCell("M12"))}`);
assert(String(formulaOf(piggyD.getCell("M16"))).includes("$M$5/SUMIF($C$6:$C$"), `Piggy bonus odds ${formulaOf(piggyD.getCell("M16"))}`);
for (const [addr, theme] of [["L16", 9], ["L17", 9], ["L18", 7], ["L19", 7]]) {
  assert(piggyD.getCell(addr).fill?.fgColor?.theme === theme, `Piggy ${addr} fill ${JSON.stringify(piggyD.getCell(addr).fill)}`);
}
assert(String(piggyD.getCell("L27").value || "").includes("strictly confidential"), "Piggy confidential at L27");
assert(piggyD.getCell(`A${piggyTot + 2}`).value == null, "Piggy non-JP should not carry the JP note");
for (let row = 1; row <= 40; row += 1) {
  assert(!String(formulaOf(piggyD.getCell(`N${row}`)) || "").includes("okay"), `Piggy N${row} kept the okay flag`);
}
const piggyWidths = Array.from({ length: 10 }, (_, i) => piggyD.getColumn(i + 1).width || 0);
assert(piggyWidths[2] === Math.max(...piggyWidths), `Piggy Win Method is not the widest column ${piggyWidths}`);
assert(String(piggyD.getCell(`C${piggyLast}`).value || "").length > 0, "Piggy last method lost its value");
assert(String(piggyD.getCell("C6").value || "").length > 0, "Piggy first method empty");
assert(!String(piggyD.getCell("C6").value).includes(" x 1"), "Piggy method gained x 1");
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
const vipSummary = {
  L9: "RTP Setting:", L10: "Base RTP:", L11: "JP RTP:", L12: "Total RTP:",
  L13: "Winning Tiers Freq:", L14: "Hit Rate:", L17: "Game info for NC",
};
for (const [addr, text] of Object.entries(vipSummary)) {
  assert(vipD.getCell(addr).value === text, `VIPElite ${addr} ${vipD.getCell(addr).value}`);
}
const vipTot = 5 + vipInfo.winningTiers + 1;
const vipJp = vipBuilt.report.delivery.jpStart;
assert(vipJp === vipTot + 4, `VIPElite JP start ${vipJp}`);
assert(String(vipD.getCell(`A${vipTot + 2}`).value || "").startsWith("Any progressive jackpot information"), "VIPElite JP note");
assert(vipD.getCell(`A${vipTot + 1}`).value == null && vipD.getCell(`A${vipTot + 3}`).value == null, "VIPElite note spacing");
assert(vipD.getCell(vipJp, 2).value === "Base JP Odds Down", `VIPElite JP start ${vipD.getCell(vipJp, 2).value}`);
assert(formulaOf(vipD.getCell("M11")) === `=C${vipJp + 28}+C${vipJp + 29}`, `VIPElite M11 ${formulaOf(vipD.getCell("M11"))}`);
assert(formulaOf(vipD.getCell("M12")) === "=M11+M10", `VIPElite M12 ${formulaOf(vipD.getCell("M12"))}`);
assert(formulaOf(vipD.getCell("M14")) === "=M5/M13", `VIPElite M14 ${formulaOf(vipD.getCell("M14"))}`);
assert(!String(vipD.getCell(vipJp, 5).value || "").toLowerCase().includes("json"), "VIPElite copied json setup");
assert(String(formulaOf(vipD.getCell(vipJp + 24, 3)) || "").includes("$M$14"), `VIPElite hit link ${formulaOf(vipD.getCell(vipJp + 24, 3))}`);
assert(String(vipD.getCell("L29").value || "").includes("strictly confidential"), "VIPElite confidential at L29");
const vipSrc = await loadBuilt(vipBuf);
const vipPj = vipSrc.getWorksheet("Progressive Jackpots");
let vipStyled = 0;
for (let r = 2; r <= 32; r += 1) {
  for (let c = 2; c <= 4; c += 1) {
    const src = vipPj.getCell(r, c);
    const dst = vipD.getCell(vipJp + r - 2, c);
    for (const edge of ["left", "right", "top", "bottom"]) {
      assert((src.border?.[edge]?.style || null) === (dst.border?.[edge]?.style || null), `VIPElite JP border ${dst.address} ${edge}`);
    }
    if (src.fill?.pattern === "solid") {
      assert(dst.fill?.fgColor?.theme === src.fill.fgColor?.theme, `VIPElite JP fill ${dst.address}`);
      vipStyled += 1;
    }
  }
}
assert(vipStyled > 0, "VIPElite PJ has no fills to compare");
await assertSafePackage(vipBuilt.buffer, "VIPElite");
await writeFile(path.join(outDir, packOutputFilename(path.basename(vipPath), "NC")), Buffer.from(vipBuilt.buffer));
console.log("vipelite", vipInfo.winningTiers, "jp", true);

const qfcName = "260928_NC_QuestForCash(BuyBonus$25)_PPS_087.xlsx";
const qfcBuf = await readFile(path.join(onedrive, "NC", "260928 Quest for Cash", qfcName));
const qfcInfo = await inspectPack(qfcBuf, qfcName, "NC");
assert(qfcInfo.buy && qfcInfo.rrp === 25 && !qfcInfo.hasJp, `QfC buy ${qfcInfo.buy} rrp ${qfcInfo.rrp}`);
const qfcBuilt = await buildPack(qfcBuf, qfcName, { lottery: "NC" });
const qfcD = (await loadBuilt(qfcBuilt.buffer)).getWorksheet("Delivery");
assertPlainNumberFormats(qfcD, "QuestForCash buy");
const qfcLast = 5 + qfcInfo.winningTiers;
const qfcSummary = {
  L6: "Base:", L7: "RRP(x):", L8: "Revenue:", L9: "Prize Fund:", L10: "RTP Setting:",
  L11: "Base RTP:", L12: "Winning Tiers Freq:", L13: "Hit Rate:", L16: "Game info for NC",
};
for (const [addr, text] of Object.entries(qfcSummary)) {
  assert(qfcD.getCell(addr).value === text, `QfC ${addr} ${qfcD.getCell(addr).value}`);
}
assert(qfcD.getCell("M7").value === 25, `QfC RRP ${qfcD.getCell("M7").value}`);
assert(formulaOf(qfcD.getCell("M8")) === "=M5*M6*M7", `QfC revenue ${formulaOf(qfcD.getCell("M8"))}`);
assert(formulaOf(qfcD.getCell("M11")) === `=G${qfcLast + 1}/M8`, `QfC base RTP ${formulaOf(qfcD.getCell("M11"))}`);
assert(formulaOf(qfcD.getCell(`J${qfcLast}`)) === `=G${qfcLast}/M$9`, `QfC share ${formulaOf(qfcD.getCell(`J${qfcLast}`))}`);
assert(String(qfcD.getCell("L28").value || "").includes("strictly confidential"), "QfC confidential at L28");
await assertSafePackage(qfcBuilt.buffer, "QuestForCash buy");
await writeFile(path.join(outDir, packOutputFilename(qfcName, "NC")), Buffer.from(qfcBuilt.buffer));
console.log("quest for cash buy", qfcInfo.winningTiers, "rrp", qfcInfo.rrp);

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

const xmasPath = path.join(onedrive, "GA", "261009_GA_ChristmasCash(MMJ)_PPS_078_004.xlsx");
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
const xmasSeed = xmasBuilt.report.delivery.seedRtp;
const xmasContrib = xmasBuilt.report.delivery.contribRtp;
assert(xmasD.getCell(xmasJp, 2).value === "Base JP Odds Down", `Xmas JP start ${xmasD.getCell(xmasJp, 2).value}`);
assert(xmasD.getCell(xmasJp, 4).value === 50000000, `Xmas odds down ${xmasD.getCell(xmasJp, 4).value}`);
assert(xmasD.getCell(xmasJp + 1, 2).value === "Stake", `Xmas stake label ${xmasD.getCell(xmasJp + 1, 2).value}`);
assert(xmasD.getCell("A2").fill?.fgColor?.theme === 4, "Christmas A2 fill");
assert(xmasD.getCell("A3").fill?.fgColor?.theme === 4, "Christmas A3 fill");
assert(xmasD.getCell(xmasJp, 2).border?.left?.style === "medium", "JP setting left border");
assert(xmasD.getCell(xmasJp, 4).border?.top?.style === "medium", "JP setting top border");
assert(xmasD.getCell(xmasJp, 6).value === "BASE", `Xmas BASE ${xmasD.getCell(xmasJp, 6).value}`);
assert(formulaOf(xmasD.getCell("B10")) === `=D${xmasSeed}`, `Xmas B10 ${formulaOf(xmasD.getCell("B10"))}`);
assert(formulaOf(xmasD.getCell("B11")) === `=D${xmasContrib}`, `Xmas B11 ${formulaOf(xmasD.getCell("B11"))}`);
assert(xmasD.getCell(xmasSeed, 2).value === "RTP", "Xmas seed RTP label");
assert(String(xmasD.getCell(xmasBuilt.report.delivery.jpStart, 2).value) !== "Overal Game Summary");
let xmasTrigger = null;
for (let row = xmasJp; row <= xmasJp + 40; row += 1) {
  if (xmasD.getCell(row, 2).value === "Target trigger") xmasTrigger = row;
  assert(xmasD.getCell(row, 2).value !== "Overal Game Summary", "GA JP block should not copy the NC overall summary");
}
assert(xmasTrigger, "Xmas missing Target trigger");
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

const mm5Path = path.join(onedrive, "VA", "260908_VA_MerryMatch5_PPS_086.xlsx");
const mm5Buf = await readFile(mm5Path);
const mm5Info = await inspectPack(mm5Buf, path.basename(mm5Path), "VA");
assert(mm5Info.layout === "VA" && !mm5Info.hasJp && !mm5Info.buy, `MerryMatch5 info ${JSON.stringify(mm5Info.layout)}`);
const mm5Built = await buildPack(mm5Buf, path.basename(mm5Path), { lottery: "VA" });
const mm5Wb = await loadBuilt(mm5Built.buffer);
const mm5D = mm5Wb.getWorksheet("Delivery");
assertPlainNumberFormats(mm5D, "MerryMatch5");
const mm5Src = await loadBuilt(mm5Buf);
const mm5SrcNames = mm5Src.worksheets.map((ws) => ws.name);
const mm5Names = mm5Wb.worksheets.map((ws) => ws.name);
assert(mm5Names.length === mm5SrcNames.length, `MerryMatch5 lost tabs ${mm5Names}`);
assert(mm5Names[0] === "Frequency" && mm5Names[1] === "Delivery", `MerryMatch5 order ${mm5Names}`);
for (const name of mm5SrcNames) assert(mm5Names.includes(name), `MerryMatch5 dropped ${name}`);
const mm5Ref = mm5Src.getWorksheet("Delivery");
const mm5Tot = mm5Built.report.delivery.tot;
assert(mm5Tot === 18 + mm5Info.winningTiers + 1, `MerryMatch5 total row ${mm5Tot}`);
let mm5Diff = [];
for (let r = 2; r <= mm5Tot; r += 1) {
  for (let c = 2; c <= 9; c += 1) {
    const a = formulaOf(mm5D.getCell(r, c));
    const b = formulaOf(mm5Ref.getCell(r, c));
    const same = a === b || (a == null && b == null) || (typeof a === "number" && typeof b === "number" && Math.abs(a - b) < 1e-9);
    if (!same) mm5Diff.push(`${mm5D.getCell(r, c).address} ${JSON.stringify(a)} vs ${JSON.stringify(b)}`);
  }
}
assert(mm5Diff.length === 0, `MerryMatch5 differs from the reference Delivery:\n${mm5Diff.slice(0, 12).join("\n")}`);
for (const col of ["B", "C", "D", "E", "F", "G", "H", "I"]) {
  assert(mm5D.getCell(`${col}17`).border?.bottom?.style === "medium", `MerryMatch5 ${col}17 underline`);
  assert(mm5D.getCell(`${col}${mm5Tot - 1}`).border?.bottom?.style === "medium", `MerryMatch5 ${col}${mm5Tot - 1} underline`);
}
assert(!mm5D.getCell("A1").value, "MerryMatch5 A1 should be empty");
assert(String(mm5D.getCell("E9").value || "").includes("strictly confidential"), "MerryMatch5 confidential at E9");
for (let r = 6; r <= 13; r += 1) {
  assert(!/jp rtp|total rtp/i.test(String(mm5D.getCell(r, 2).value || "")), `MerryMatch5 has JP row at B${r}`);
}
assert(mm5D.properties?.tabColor?.theme === 5, "MerryMatch5 tab color");
await assertSafePackage(mm5Built.buffer, "MerryMatch5");
await writeFile(path.join(outDir, packOutputFilename(path.basename(mm5Path), "VA")), Buffer.from(mm5Built.buffer));
console.log("merrymatch5", mm5Info.winningTiers, "matches reference");

const goatDir = path.join(onedrive, "VA", "261006 Year of Fire Goat(SupremeJP)");
const goatMainName = "261006_VA_YearOfFireGoatSupremeJP(MainGame)_PPS_084_003.xlsx";
const goatMainBuf = await readFile(path.join(goatDir, goatMainName));
const goatInfo = await inspectPack(goatMainBuf, goatMainName, "VA");
assert(goatInfo.hasJp, "Fire Goat main should be JP");
assert(goatInfo.winningTiers === 199, `Fire Goat main tiers ${goatInfo.winningTiers}`);
assert(goatInfo.sideHits.length === 5, `Fire Goat side hits ${goatInfo.sideHits.length}`);
const goatBuilt = await buildPack(goatMainBuf, goatMainName, { lottery: "VA" });
const goatWb = await loadBuilt(goatBuilt.buffer);
const goatD = goatWb.getWorksheet("Delivery");
assertPlainNumberFormats(goatD, "FireGoat main");
const goatSrc = await loadBuilt(goatMainBuf);
assert(goatWb.worksheets.length === goatSrc.worksheets.length + 1, "Fire Goat lost tabs");
assert(goatD.getCell("B12").value.trim().replace(/:$/, "") === "JP RTP", `Fire Goat B12 ${goatD.getCell("B12").value}`);
assert(goatD.getCell("B13").value.trim().replace(/:$/, "") === "Total RTP", `Fire Goat B13 ${goatD.getCell("B13").value}`);
assert(goatD.getCell("B19").value === "NUMBER", `Fire Goat B19 ${goatD.getCell("B19").value}`);
const goatTot = goatBuilt.report.delivery.tot;
const goatJp = goatBuilt.report.delivery.jpStart;
assert(goatTot === 20 + 199 + 1 && goatJp === goatTot + 3, `Fire Goat rows ${goatTot} ${goatJp}`);
assert(formulaOf(goatD.getCell("C12")) === `=C${goatJp + 28}+C${goatJp + 29}`, `Fire Goat C12 ${formulaOf(goatD.getCell("C12"))}`);
assert(formulaOf(goatD.getCell("C13")) === "=C12+C11", `Fire Goat C13 ${formulaOf(goatD.getCell("C13"))}`);
assert(goatD.getCell(goatJp, 3).value === 20000000, `Fire Goat JP odds ${goatD.getCell(goatJp, 3).value}`);
assert(goatD.getCell(goatJp + 9, 13).value === 50, `Fire Goat M ${goatD.getCell(goatJp + 9, 13).value}`);
assert(goatD.getCell(goatJp + 7, 2).fill?.fgColor?.theme === 5, "Fire Goat seed fill");
assert(goatD.getCell(goatJp, 3).fill?.fgColor?.theme === 7, "Fire Goat JP input fill");
assert(goatD.getCell(goatJp + 7, 2).border?.left?.style === "thin", "Fire Goat seed border");
assert(goatD.getCell(goatJp + 6, 2).fill?.pattern !== "solid", "Fire Goat blank row has no fill");
assert(String(formulaOf(goatD.getCell(goatJp + 25, 3))).includes("$C$6"), formulaOf(goatD.getCell(goatJp + 25, 3)));
assert(!/Frequency/.test(String(formulaOf(goatD.getCell(goatJp + 25, 3)))), formulaOf(goatD.getCell(goatJp + 25, 3)));
assert(String(goatD.getCell("E9").value).startsWith("Hit Rate(+"), `Fire Goat E9 ${goatD.getCell("E9").value}`);
const goatF9 = String(formulaOf(goatD.getCell("F9")));
assert(goatF9.includes("$C$15") && goatF9.includes("Frequency!$N$15"), `Fire Goat F9 ${goatF9}`);
assert(String(goatD.getCell("H9").value || "").includes("strictly confidential"), "Fire Goat confidential at H9");
await assertSafePackage(goatBuilt.buffer, "FireGoat main");
await writeFile(path.join(outDir, packOutputFilename(goatMainName, "VA")), Buffer.from(goatBuilt.buffer));

const goatSideName = "261006_VA_YearOfFireGoatSupremeJP(SideBet)_PPS_084_003.xlsx";
const goatSideBuf = await readFile(path.join(goatDir, goatSideName));
const sideBuilt = await buildPack(goatSideBuf, goatSideName, { lottery: "VA" });
const sideWb = await loadBuilt(sideBuilt.buffer);
const sideD = sideWb.getWorksheet("Delivery");
assertPlainNumberFormats(sideD, "FireGoat side");
assert(sideBuilt.report.hasJp, "Fire Goat side should be JP");
const sideJp = sideBuilt.report.delivery.jpStart;
assert(formulaOf(sideD.getCell("C12")) === `=C${sideJp + 28}+C${sideJp + 29}`, `Fire Goat side C12 ${formulaOf(sideD.getCell("C12"))}`);
assert(!String(sideD.getCell("E9").value || "").startsWith("Hit Rate"), "Fire Goat side has no hit-rate block");
assert(String(sideD.getCell("E9").value || "").includes("strictly confidential"), "Fire Goat side confidential at E9");
assert(/hit rate/i.test(String(sideD.getCell("B15").value)), `Fire Goat side B15 ${sideD.getCell("B15").value}`);
await assertSafePackage(sideBuilt.buffer, "FireGoat side");
await writeFile(path.join(outDir, packOutputFilename(goatSideName, "VA")), Buffer.from(sideBuilt.buffer));
console.log("fire goat", goatInfo.winningTiers, sideBuilt.report.winningTiers, "jp", goatJp, sideJp);

function sameCell(a, b) {
  const x = formulaOf(a);
  const y = formulaOf(b);
  if (x === y || (x == null && y == null)) return true;
  if (typeof x === "number" && typeof y === "number") return Math.abs(x - y) < 1e-9;
  if (typeof x === "string" && typeof y === "string") return x.replaceAll("$", "").trim() === y.replaceAll("$", "").trim();
  return false;
}

function compareBlock(out, ref, rows, cols, label) {
  const diff = [];
  for (const r of rows) {
    for (const c of cols) {
      if (!sameCell(out.getCell(r, c), ref.getCell(r, c))) {
        diff.push(`${out.getCell(r, c).address} ${JSON.stringify(formulaOf(out.getCell(r, c)))} vs ${JSON.stringify(formulaOf(ref.getCell(r, c)))}`);
      }
    }
  }
  assert(diff.length === 0, `${label} differs from the reference:\n${diff.slice(0, 12).join("\n")}`);
}

const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);

function checkDcBox(out, ref, top, last, label) {
  const labels = ["Total Tickets", "Retail Price", "Revenue", "Prize Fund", "Payout", "Odds", "Top Prize", "Max Top Prize $ Value", "Price Points", null, "Prize (as multiple of stake)", "NIL", ">0 to <1", 1, ">1 to 2", ">2 to 5", ">5 to 10", ">10 to 20", ">20 to 30", ">30 to 50", ">50 to 100", ">100"];
  labels.forEach((text, i) => {
    assert((out.getCell(top + i, 11).value ?? null) === text, `${label} K${top + i} ${out.getCell(top + i, 11).value}`);
    assert((ref.getCell(top + i, 11).value ?? null) === text, `${label} reference K${top + i} ${ref.getCell(top + i, 11).value}`);
    assert(out.getCell(top + i, 11).border?.left?.style === "medium", `${label} K${top + i} left edge`);
    assert(out.getCell(top + i, 12).border?.right?.style === "medium", `${label} L${top + i} right edge`);
  });
  const end = top + labels.length - 1;
  for (const col of [11, 12]) {
    assert(out.getCell(top, col).border?.top?.style === "medium", `${label} box top`);
    assert(out.getCell(end, col).border?.bottom?.style === "medium", `${label} box bottom`);
  }
  let fund = 0;
  let wins = 0;
  let maxPrize = 0;
  const base = out.getCell("L7").value;
  const prizes = [];
  for (let r = 5; r <= last; r += 1) {
    const d = out.getCell(r, 4).value;
    const e = out.getCell(r, 5).value;
    fund += d * e;
    wins += e;
    maxPrize = Math.max(maxPrize, d);
    prizes.push([d / base, d * e]);
  }
  const share = (test) => prizes.filter(([m]) => test(m)).reduce((s, [, g]) => s + g, 0) / fund;
  const expected = [
    share((m) => m > 0 && m < 1),
    share((m) => m === 1),
    ...[[1, 2], [2, 5], [5, 10], [10, 20], [20, 30], [30, 50], [50, 100]].map(([lo, hi]) => share((m) => m > lo && m <= hi)),
    share((m) => m > 100),
  ];
  expected.forEach((v, i) => {
    const refValue = ref.getCell(top + 12 + i, 12).value;
    assert(Math.abs(v - refValue) < 1e-9, `${label} band ${labels[12 + i]} ${v} vs reference ${refValue}`);
  });
  const odds = `1 in ${(out.getCell("L6").value / wins).toFixed(2)}`;
  assert(odds === ref.getCell(top + 5, 12).value, `${label} odds ${odds} vs ${ref.getCell(top + 5, 12).value}`);
  const topText = `${(maxPrize / base).toLocaleString("en-US")}x`;
  assert(topText === String(ref.getCell(top + 6, 12).value).trim(), `${label} top prize ${topText}`);
  assert(String(formulaOf(out.getCell(top + 5, 12))).includes("TEXT("), `${label} odds formula`);
  return { maxPrize, base, end };
}

const dcDir = path.join(onedrive, "DC");
const dcMmName = "260812_DC_MerryMatch5_PPS_086.xlsx";
const dcMmBuf = await readFile(path.join(dcDir, dcMmName));
const dcMmInfo = await inspectPack(dcMmBuf, dcMmName, "DC");
assert(dcMmInfo.layout === "DC" && !dcMmInfo.hasJp, `DC MerryMatch5 layout ${dcMmInfo.layout} jp ${dcMmInfo.hasJp}`);
const dcMmBuilt = await buildPack(dcMmBuf, dcMmName, { lottery: "DC" });
const dcMmWb = await loadBuilt(dcMmBuilt.buffer);
const dcMmD = dcMmWb.getWorksheet("Delivery");
assertPlainNumberFormats(dcMmD, "DC MerryMatch5");
const dcMmSrc = await loadBuilt(dcMmBuf);
const dcMmRef = dcMmSrc.getWorksheet("Delivery");
const dcMmNames = dcMmWb.worksheets.map((w) => w.name);
assert(dcMmNames.join("|") === dcMmSrc.worksheets.map((w) => w.name).join("|"), `DC MerryMatch5 tabs ${dcMmNames}`);
const dcMmLast = 4 + dcMmInfo.winningTiers;
compareBlock(dcMmD, dcMmRef, range(2, dcMmLast + 1), range(2, 9), "DC MerryMatch5 table");
compareBlock(dcMmD, dcMmRef, range(2, 13), [11, 12], "DC MerryMatch5 summary");
assert(!dcMmD.getCell("A1").value, "DC A1 should be empty");
assert(String(dcMmD.getCell("K16").value || "").includes("strictly confidential"), "DC MerryMatch5 confidential at K16");
assert(!/game info/i.test(JSON.stringify(dcMmD.getColumn(11).values)), "DC should not write Game info");
const mmBox = checkDcBox(dcMmD, dcMmRef, 23, dcMmLast, "DC MerryMatch5");
assert(dcMmD.getCell("L31").value === String(dcMmRef.getCell("L31").value).trim(), `DC price points ${dcMmD.getCell("L31").value}`);
assert(mmBox.maxPrize / mmBox.base * 50 === dcMmRef.getCell("L30").value, "DC MerryMatch5 max top prize value");
for (const col of ["B", "C", "D", "E", "F", "G", "H", "I"]) {
  assert(dcMmD.getCell(`${col}3`).border?.bottom?.style === "medium", `DC ${col}3 underline`);
  assert(dcMmD.getCell(`${col}${dcMmLast}`).border?.bottom?.style === "medium", `DC ${col}${dcMmLast} underline`);
}
await assertSafePackage(dcMmBuilt.buffer, "DC MerryMatch5");
await writeFile(path.join(outDir, packOutputFilename(dcMmName, "DC")), Buffer.from(dcMmBuilt.buffer));
console.log("dc merrymatch5", dcMmInfo.winningTiers, "matches reference");

const dcPhName = "260603_DC_Pharaoh'sDestiny(ChatterJP)_PPS_084_005.xlsx";
const dcPhBuf = await readFile(path.join(dcDir, dcPhName));
const dcPhInfo = await inspectPack(dcPhBuf, dcPhName, "DC");
assert(dcPhInfo.hasJp, "Pharaoh's Destiny should be JP");
const phPrices = [0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 30];
const dcPhBuilt = await buildPack(dcPhBuf, dcPhName, { lottery: "DC", pricePoints: phPrices });
const dcPhWb = await loadBuilt(dcPhBuilt.buffer);
const dcPhD = dcPhWb.getWorksheet("Delivery");
assertPlainNumberFormats(dcPhD, "DC Pharaoh");
const dcPhSrc = await loadBuilt(dcPhBuf);
const dcPhRef = dcPhSrc.getWorksheet("Delivery");
const phLast = 4 + dcPhInfo.winningTiers;
const phJp = dcPhBuilt.report.delivery.jpStart;
assert(phJp === phLast + 4, `Pharaoh JP start ${phJp}`);
compareBlock(dcPhD, dcPhRef, range(2, phLast + 1), range(2, 9), "DC Pharaoh table");
compareBlock(dcPhD, dcPhRef, range(2, 15), [11], "DC Pharaoh summary labels");
compareBlock(dcPhD, dcPhRef, [6, 7, 9, 10, 12, 13, 14, 15], [12], "DC Pharaoh summary values");
compareBlock(dcPhD, dcPhRef, range(phJp, phJp + 30).filter((r) => r !== phJp + 27), range(2, 16), "DC Pharaoh JP block");
assert(dcPhD.getCell(phJp + 27, 2).value === "Main Game RTP", "Pharaoh Main Game RTP row");
assert(formulaOf(dcPhD.getCell(phJp + 27, 3)) === "=$L$11", `Pharaoh Main Game RTP follows PJ (Frequency N11) ${formulaOf(dcPhD.getCell(phJp + 27, 3))}`);
const phPj = dcPhSrc.getWorksheet("Progressive Jackpots");
for (let r = 2; r <= 32; r += 1) {
  for (let c = 2; c <= 4; c += 1) {
    const src = phPj.getCell(r, c);
    const dst = dcPhD.getCell(phJp + r - 2, c);
    for (const edge of ["left", "right", "top", "bottom"]) {
      assert((src.border?.[edge]?.style || null) === (dst.border?.[edge]?.style || null), `Pharaoh JP border ${dst.address} ${edge}`);
    }
    if (src.fill?.pattern === "solid") assert(dst.fill?.fgColor?.theme === src.fill.fgColor?.theme, `Pharaoh JP fill ${dst.address}`);
  }
}
assert(String(dcPhD.getCell("K18").value || "").includes("strictly confidential"), "Pharaoh confidential at K18");
const phBox = checkDcBox(dcPhD, dcPhRef, 25, phLast, "DC Pharaoh");
assert(dcPhD.getCell("L33").value === "$0.10, $0.20, $0.50, $1.00, $2.00, $5.00, $10.00, $20.00, $30.00", `Pharaoh price points ${dcPhD.getCell("L33").value}`);
assert(phBox.maxPrize / phBox.base * 30 === dcPhRef.getCell("L32").value, "Pharaoh max top prize value");
assert(dcPhWb.worksheets[1].name === "Delivery", "Pharaoh Delivery after Frequency");
await assertSafePackage(dcPhBuilt.buffer, "DC Pharaoh");
await writeFile(path.join(outDir, packOutputFilename(dcPhName, "DC")), Buffer.from(dcPhBuilt.buffer));
console.log("dc pharaoh", dcPhInfo.winningTiers, "jp", phJp);

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
for (let i = 1; i <= 12; i += 1) {
  const xml = await readZipText(xmasBuilt.buffer, `xl/worksheets/sheet${i}.xml`);
  if (xml) assert(!xml.includes("[1]"), `sheet${i} has a broken [1] sheet reference`);
}

const cached = (cell) => {
  const v = cell.value;
  return v && typeof v === "object" && "result" in v ? v.result : v;
};

function payList(ws) {
  const map = new Map();
  for (let r = 3; r <= 18; r += 1) {
    for (let c = 6; c <= 20; c += 1) {
      const name = cached(ws.getCell(r, c));
      const value = cached(ws.getCell(r, c + 1));
      if (typeof name === "string" && /-\d+$/.test(name) && typeof value === "number") map.set(name.toLowerCase(), value);
    }
  }
  return map;
}

function sidePrizes(ws) {
  const list = [];
  for (let r = 21; typeof cached(ws.getCell(`N${r}`)) === "number"; r += 1) list.push(cached(ws.getCell(`N${r}`)));
  return list;
}

const lnzCases = [
  ["260708_LNZ_ShowMuchFun_PPS_062.xlsx", [1, 2, 3], "Show Much Fun"],
  ["260706_LNZ_CashClash_PPS_062.xlsx", [1, 2, 5], "Cash Explosion"],
  ["260128_LNZ_LabRichesV2($3.0)_PPS_062.xlsx", [3], "Lab Riches($3.0)"],
];
for (const [name, prices, title] of lnzCases) {
  const buf = await readFile(path.join(onedrive, "LNZ", name));
  const info = await inspectPack(buf, name, "LNZ");
  assert(info.layout === "LNZ" && !info.unsupported, `${name} inspect ${info.unsupported}`);
  assert(info.existingDelivery.length > 0, `${name} existing Delivery sheets`);
  const built = await buildPack(buf, name, { lottery: "LNZ", pricePoints: prices });
  const names = prices.map(lnzSheetName);
  assert(built.report.deliverySheets.join() === names.join(), `${name} sheets ${built.report.deliverySheets}`);
  const out = await loadBuilt(built.buffer);
  const order = out.worksheets.map((w) => w.name);
  assert(order.slice(0, names.length + 1).join() === ["Frequency", ...names].join(), `${name} order ${order}`);
  assert(order.filter((n) => /^delivery/i.test(n)).length === names.length, `${name} old Delivery sheets left: ${order}`);
  const src = await loadBuilt(buf);
  for (const s of built.report.delivery.sheets) {
    const ws = out.getWorksheet(s.name);
    const ref = src.worksheets.find((w) => /^delivery/i.test(w.name) && Number(cached(w.getCell("C3"))) === s.price);
    assert(ref, `${name} has no sample sheet for $${s.price}`);
    assert(s.last - s.first + 1 === info.winningTiers, `${s.name} tier rows`);
    for (let r = s.first; r <= s.last; r += 1) {
      for (const col of ["A", "B", "E", "F"]) {
        const a = cached(ws.getCell(`${col}${r}`));
        const b = cached(ref.getCell(`${col}${r}`));
        assert(String(a) === String(b), `${name} ${s.name} ${col}${r} ${a} vs ${b}`);
      }
    }
    for (const addr of ["C3", "C4", "C6", "C8", "C9", "C10"]) {
      assert(cached(ws.getCell(addr)) === cached(ref.getCell(addr)), `${name} ${s.name} ${addr}`);
    }
    const sideRef = sidePrizes(ref);
    if (sideRef.length) assert(sidePrizes(ws).join() === sideRef.join(), `${name} ${s.name} side prizes`);
    const payOut = payList(ws);
    const payRef = payList(ref);
    assert(payOut.size === payRef.size, `${name} ${s.name} pay list ${payOut.size} vs ${payRef.size}`);
    for (const [k, v] of payRef) assert(payOut.get(k) === v, `${name} ${s.name} pay ${k} ${payOut.get(k)} vs ${v}`);
    const t = s.tot;
    const f = (addr) => formulaOf(ws.getCell(addr));
    assert(f("A1") === `="${title}. Payout "&TEXT(100*C5,"###.00")&"%. Odds 1 in "&TEXT(D${t},"###.00")`, `${s.name} A1 ${f("A1")}`);
    assert(f("C5") === `=(G${t}/C6/C3)*C4` && f("C7") === `=C6-E${t}`, `${s.name} C5/C7`);
    assert(f(`E${t + 1}`) === `=SUMIF(F22:F${s.last},C$3,E22:E${s.last})/E${t}`, `${s.name} break even ${f(`E${t + 1}`)}`);
    assert(f("O21") === `=SUMIF($F$22:$F$${s.last},N21,$E$22:$E$${s.last})`, `${s.name} O21 ${f("O21")}`);
    assert(f("L22") === "=IF(F22<C$9,\"\",H22)", `${s.name} L22 ${f("L22")}`);
    assert(ws.getCell(`G${t + 5}`).value == null, `${s.name} scratch cells below the total`);
    assertPlainNumberFormats(ws, `${name} ${s.name}`);
  }
  await assertSafePackage(built.buffer, `LNZ ${name}`);
  await writeFile(path.join(outDir, packOutputFilename(name, "LNZ")), Buffer.from(built.buffer));
  console.log("lnz", name, info.winningTiers, names.join(" "));
}
const lnzJp = await inspectPack(vipBuf, path.basename(vipPath), "LNZ");
assert(/jackpot/i.test(lnzJp.unsupported || ""), "LNZ flags jackpot games");
let lnzJpThrew = false;
try {
  await buildPack(vipBuf, path.basename(vipPath), { lottery: "LNZ", pricePoints: [1] });
} catch {
  lnzJpThrew = true;
}
assert(lnzJpThrew, "LNZ refuses jackpot games");
console.log("pack ok");
