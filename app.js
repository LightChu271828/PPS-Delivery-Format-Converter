import { buildPps, inspectPps, outputFilename } from "./builder.js?v=20261009-1";
import {
  buildPack,
  inspectPack,
  isPackLottery,
  packOutputFilename,
  DC_PRICE_OPTIONS,
  DC_DEFAULT_PRICE_POINTS,
} from "./pack-builder.js?v=20261009-1";

const $ = (id) => document.getElementById(id);

const state = {
  templates: null,
  file: null,
  buffer: null,
  inspect: null,
  output: null,
  outputName: null,
  lottery: "KY",
};

function toast(msg) {
  const el = $("toast");
  el.textContent = msg;
  el.classList.add("on");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.textContent && el.classList.remove("on"), 2200);
}

function fmt(n, digits = 3) {
  if (n == null || Number.isNaN(n)) return "—";
  return Number(n).toLocaleString("en-US", { maximumFractionDigits: digits });
}

function pct(n) {
  if (n == null || Number.isNaN(n)) return "—";
  return `${(n * 100).toFixed(2)}%`;
}

function stat(value, label, cls = "") {
  return `<div class="stat"><div class="v ${cls}">${value}</div><div class="l">${label}</div></div>`;
}

function packMode() {
  return isPackLottery(state.lottery);
}

function dcMode() {
  return state.lottery === "DC";
}

function show(el, on) {
  el.classList.toggle("hidden", !on);
  el.hidden = !on;
}

function syncLotteryUi() {
  document.querySelectorAll(".markets button[data-lottery]").forEach((btn) => {
    const on = btn.dataset.lottery === state.lottery;
    btn.classList.toggle("on", on);
    btn.setAttribute("aria-pressed", on ? "true" : "false");
  });
  show($("priceGrid"), !packMode());
  show($("dcPriceGrid"), dcMode());
  $("buildBtn").textContent = packMode() ? "Add Delivery sheet" : "Add Delivery tabs";
  const kyTabs = $("kyResultTabs");
  const packTabs = $("packResultTabs");
  if (kyTabs) show(kyTabs, !packMode());
  if (packTabs) show(packTabs, packMode());
}

async function fetchTemplate(name) {
  const res = await fetch(`./assets/templates/${name}`);
  if (!res.ok) throw new Error(`Could not load ${name}`);
  return res.arrayBuffer();
}

async function loadTemplates() {
  const pill = $("templateStatus");
  try {
    const [mmj3, nojp] = await Promise.all([fetchTemplate("mmj3.xlsx"), fetchTemplate("no-jp.xlsx")]);
    state.templates = { mmj3, "no-jp": nojp };
    pill.textContent = "Templates ready";
    pill.className = "pill ok";
  } catch (err) {
    pill.textContent = "Template load failed";
    pill.className = "pill bad";
    throw err;
  }
}

function canBuild() {
  if (!state.file || !state.inspect) return false;
  if (dcMode()) return selectedDcPrices().length > 0;
  return packMode() || Boolean(state.templates);
}

function setFile(file) {
  state.file = file;
  state.buffer = null;
  state.inspect = null;
  state.output = null;
  state.outputName = null;
  $("fileLabel").textContent = file ? `${file.name} · ${(file.size / 1024).toFixed(0)} KB` : "";
  $("downloadBtn").disabled = true;
  $("downloadBtn").classList.remove("ready");
  $("inspectBlock").classList.add("hidden");
  $("inspectBlock").hidden = true;
  $("resultBlock").classList.add("hidden");
  $("resultBlock").hidden = true;
  $("buildBtn").disabled = true;
  $("statusHint").textContent = file ? "Reading Frequency…" : "Choose a PPS to inspect it before building.";
}

async function inspectSelected() {
  if (!state.file) return;
  try {
    state.buffer = await state.file.arrayBuffer();
    state.inspect = packMode()
      ? await inspectPack(state.buffer, state.file.name, state.lottery)
      : await inspectPps(state.buffer, state.file.name);
    renderInspect(state.inspect);
    $("buildBtn").disabled = !canBuild();
    $("statusHint").textContent = packMode()
      ? `Looks readable. Add the ${state.lottery} Delivery sheet, then download.`
      : "Looks readable. Add Delivery tabs, then download.";
  } catch (err) {
    $("inspectBlock").classList.remove("hidden");
    $("inspectBlock").hidden = false;
    $("inspectStats").innerHTML = stat("Failed", "Inspect", "bad");
    $("inspectNote").className = "note bad";
    $("inspectNote").innerHTML = `<b>${err.message}</b>`;
    $("buildBtn").disabled = true;
    $("statusHint").textContent = err.message;
  }
}

function renderInspect(info) {
  $("inspectBlock").classList.remove("hidden");
  $("inspectBlock").hidden = false;
  if (packMode()) {
    $("inspectStats").innerHTML = [
      stat(info.layout || "—", "Layout"),
      stat(info.winningTiers, "Winning tiers"),
      stat(fmt(info.wins, 0), "Wins / pool"),
      stat(fmt(info.hitRate, 3), "Hit rate"),
      stat(pct(info.rtp ?? info.actualRtp), "RTP setting"),
      stat(info.hasJp ? "Yes" : "No", "Jackpot RTP", info.hasJp ? "ok" : ""),
    ].join("");
    const bits = [];
    bits.push(`<b>${info.jurisdiction || state.lottery}</b>${info.title ? ` · ${info.title}` : ""}.`);
    bits.push(`Ticket <b>$${info.base}</b>, pool <b>${fmt(info.quantity, 0)}</b>.`);
    if (info.buy) bits.push(`RRP <b>${info.rrp}x</b>.`);
    if (info.layout === "GA") bits.push("Winning tiers sort by prize, smallest first.");
    else bits.push("Winning tiers keep Frequency order.");
    if (info.layout === "VA" && info.sideHits?.length) {
      bits.push(`Side-bet hit rates: <b>${info.sideHits.length}</b>.`);
    }
    if (info.layout === "DC") {
      const prices = selectedDcPrices();
      bits.push(
        prices.length
          ? `Price points: <b>${prices.map((p) => `$${p.toFixed(2)}`).join(", ")}</b>. Max Top Prize uses $${prices[prices.length - 1].toFixed(2)}.`
          : "Pick at least one price point.",
      );
    }
    if (info.existingDelivery?.length) {
      bits.push(`Will replace existing ${info.existingDelivery.map((n) => `<code>${n}</code>`).join(", ")}.`);
    }
    if (info.zeroFrequency?.length) {
      bits.push(`Skipped ${info.zeroFrequency.length} Frequency row(s) with 0 Odds up.`);
    }
    $("inspectNote").className = "note";
    $("inspectNote").innerHTML = bits.join(" ");
    return;
  }
  const kyClass = info.kentucky ? "ok" : "warn";
  $("inspectStats").innerHTML = [
    stat(info.schema || "—", "Schema", info.schemaError ? "warn" : ""),
    stat(info.winningTiers, "Winning tiers"),
    stat(fmt(info.wins, 0), "Wins / pool"),
    stat(fmt(info.hitRate, 3), "Hit rate"),
    stat(pct(info.rtp ?? info.actualRtp), "RTP setting"),
    stat(info.kentucky ? "Yes" : "Not found", "Kentucky identity", kyClass),
  ].join("");
  const bits = [];
  bits.push(`Ticket <b>$${info.base}</b>, pool <b>${fmt(info.quantity, 0)}</b>.`);
  const grid = selectedPriceGrid();
  if (grid.length) bits.push(`Price grid: <b>${grid.join(", ")}</b>.`);
  if (info.jpCount) bits.push(`Jackpots: <b>${info.jpNames.join(", ") || info.jpCount}</b>.`);
  if (info.schemaError) bits.push(info.schemaError);
  if (info.existingDelivery.length) {
    bits.push(`Will replace existing ${info.existingDelivery.map((n) => `<code>${n}</code>`).join(", ")}.`);
  }
  if (info.zeroFrequency?.length) {
    bits.push(`Skipped ${info.zeroFrequency.length} Frequency row(s) with 0 Odds up.`);
  }
  if (!info.kentucky) {
    bits.push("Kentucky identity was not found in the labels or filename.");
  }
  $("inspectNote").className = info.schemaError ? "note warn" : info.kentucky ? "note" : "note warn";
  $("inspectNote").innerHTML = bits.join(" ");
}

function renderResult(report) {
  $("resultBlock").classList.remove("hidden");
  $("resultBlock").hidden = false;
  syncLotteryUi();
  if (packMode()) {
    $("resultStats").innerHTML = [
      stat(report.layout, "Written as"),
      stat(report.winningTiers, "Delivery win rows"),
      stat(report.uniquePrizes, "Prize groups"),
      stat(fmt(report.wins, 0), "Wins"),
      stat(fmt(report.hitRate, 3), "Hit rate"),
      stat(pct(report.actualRtp), "Indep. RTP from Frequency"),
    ].join("");
    return;
  }
  $("resultStats").innerHTML = [
    stat(report.schema, "Written as"),
    stat(report.winningTiers, "Delivery win rows"),
    stat(report.uniquePrizes, "Odds prize groups"),
    stat(fmt(report.wins, 0), "Wins"),
    stat(fmt(report.hitRate, 3), "Hit rate"),
    stat(pct(report.actualRtp), "Indep. RTP from Frequency"),
  ].join("");
}

async function build() {
  if (!state.buffer) return;
  if (!packMode() && !state.templates) return;
  $("buildBtn").disabled = true;
  $("statusHint").textContent = packMode() ? "Building Delivery sheet…" : "Building Delivery tabs…";
  try {
    if (packMode()) {
      const { buffer, report } = await buildPack(state.buffer, state.file.name, {
        lottery: state.lottery,
        pricePoints: dcMode() ? selectedDcPrices() : undefined,
      });
      state.output = buffer;
      state.outputName = packOutputFilename(state.file.name, state.lottery);
      renderResult(report);
      $("downloadBtn").disabled = false;
      $("downloadBtn").classList.add("ready");
      $("statusHint").textContent = `Ready: ${state.outputName}`;
      toast("Delivery sheet added");
    } else {
      const { buffer, report } = await buildPps(state.buffer, state.file.name, {
        templates: state.templates,
        priceGrid: selectedPriceGrid(),
      });
      state.output = buffer;
      state.outputName = outputFilename(state.file.name);
      renderResult(report);
      $("downloadBtn").disabled = false;
      $("downloadBtn").classList.add("ready");
      $("statusHint").textContent = `Ready: ${state.outputName}`;
      toast("Delivery tabs added");
    }
  } catch (err) {
    $("resultBlock").classList.remove("hidden");
    $("resultBlock").hidden = false;
    $("resultStats").innerHTML = stat("Failed", "Build", "bad");
    $("statusHint").textContent = err.message;
    toast(err.message);
  } finally {
    $("buildBtn").disabled = !canBuild();
  }
}

function download() {
  if (!state.output) return;
  const blob = new Blob([state.output], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = state.outputName;
  a.click();
  URL.revokeObjectURL(a.href);
}

function selectedPriceGrid() {
  return [...document.querySelectorAll("#priceGridOptions input[type=checkbox]:checked")].map((el) =>
    Number(el.value),
  );
}

function resetPriceGrid() {
  document.querySelectorAll("#priceGridOptions input[type=checkbox]").forEach((el) => {
    el.checked = true;
  });
  document.querySelectorAll("#dcPriceOptions input[type=checkbox]").forEach((el) => {
    el.checked = DC_DEFAULT_PRICE_POINTS.includes(Number(el.value));
  });
}

function renderDcPriceOptions() {
  $("dcPriceOptions").innerHTML = DC_PRICE_OPTIONS.map(
    (p) =>
      `<label><input type="checkbox" value="${p}"${DC_DEFAULT_PRICE_POINTS.includes(p) ? " checked" : ""}> ${p}</label>`,
  ).join("");
}

function selectedDcPrices() {
  return [...document.querySelectorAll("#dcPriceOptions input[type=checkbox]:checked")]
    .map((el) => Number(el.value))
    .sort((a, b) => a - b);
}

function reset() {
  $("file").value = "";
  resetPriceGrid();
  setFile(null);
  toast("Reset");
}

function setLottery(id) {
  if (!id || id === state.lottery) return;
  state.lottery = id;
  state.output = null;
  state.outputName = null;
  $("downloadBtn").disabled = true;
  $("downloadBtn").classList.remove("ready");
  $("resultBlock").classList.add("hidden");
  $("resultBlock").hidden = true;
  syncLotteryUi();
  if (state.file) inspectSelected();
}

const drop = $("drop");
drop.addEventListener("click", () => $("file").click());
drop.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") $("file").click();
});
["dragenter", "dragover"].forEach((ev) => {
  drop.addEventListener(ev, (e) => {
    e.preventDefault();
    drop.classList.add("drag");
  });
});
["dragleave", "drop"].forEach((ev) => {
  drop.addEventListener(ev, (e) => {
    e.preventDefault();
    drop.classList.remove("drag");
  });
});
drop.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files?.[0];
  if (file) {
    setFile(file);
    inspectSelected();
  }
});
$("file").addEventListener("change", () => {
  const file = $("file").files?.[0];
  if (file) {
    setFile(file);
    inspectSelected();
  }
});
$("buildBtn").addEventListener("click", build);
$("downloadBtn").addEventListener("click", download);
$("resetBtn").addEventListener("click", reset);
$("priceGridOptions").addEventListener("change", () => {
  if (state.inspect && !packMode()) renderInspect(state.inspect);
});
$("dcPriceOptions").addEventListener("change", () => {
  if (state.inspect && dcMode()) renderInspect(state.inspect);
  $("buildBtn").disabled = !canBuild();
});
document.querySelector(".markets").addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-lottery]");
  if (!btn || btn.disabled) return;
  setLottery(btn.dataset.lottery);
});

renderDcPriceOptions();
syncLotteryUi();
loadTemplates()
  .then(() => {
    $("buildBtn").disabled = !canBuild();
  })
  .catch((err) => {
    $("statusHint").textContent = err.message;
    if (state.inspect && packMode()) {
      $("buildBtn").disabled = false;
    }
  });
