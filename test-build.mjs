import { createRequire } from "node:module";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const ExcelJS = require("./vendor/exceljs.min.js");
globalThis.ExcelJS = ExcelJS;

const { buildPps, inspectPps } = await import("./builder.js");

const root = path.dirname(fileURLToPath(import.meta.url));
const templates = {
  mmj3: await readFile(path.join(root, "assets/templates/mmj3.xlsx")),
  "no-jp": await readFile(path.join(root, "assets/templates/no-jp.xlsx")),
  ssj: await readFile(path.join(root, "assets/templates/summary-ssj.xlsx")),
};

const source =
  process.argv[2] ||
  path.join(root, "..", "260915 X the Money(SSJ)", "260916_KY_XtheMoneySSJ($2)_PPS_083_004.xlsx");

const buf = await readFile(source);
const info = await inspectPps(buf, path.basename(source));
console.log("inspect", info);
const { buffer, report } = await buildPps(buf, path.basename(source), { templates });
console.log("report", report);
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
