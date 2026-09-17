import { buildPps, inspectPps, outputFilename } from "./builder.js?v=20260917-7";

const $ = (id) => document.getElementById(id);

const state = {
  templates: null,
  file: null,
  buffer: null,
  inspect: null,
  output: null,
  outputName: null,
};

function toast(msg) {
  const el = $("toast");
  el.textContent = msg;
  el.classList.add("on");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.classList.remove("on"), 2200);
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

async function fetchTemplate(name) {
  const res = await fetch(`./assets/templates/${name}`);
  if (!res.ok) throw new Error(`Could not load ${name}`);
  return res.arrayBuffer();
}

async function loadTemplates() {
  const pill = $("templateStatus");
  try {
    const [mmj3, nojp] = await Promise.all([
      fetchTemplate("mmj3.xlsx"),
      fetchTemplate("no-jp.xlsx"),
    ]);
    state.templates = { mmj3, "no-jp": nojp };
    pill.textContent = "Templates ready";
    pill.className = "pill ok";
  } catch (err) {
    pill.textContent = "Template load failed";
    pill.className = "pill bad";
    throw err;
  }
}

function setFile(file) {
  state.file = file;
  state.buffer = null;
  state.inspect = null;
  state.output = null;
  state.outputName = null;
  $("fileLabel").textContent = file ? `${file.name} · ${(file.size / 1024).toFixed(0)} KB` : "";
  $("downloadBtn").disabled = true;
    $("inspectBlock").classList.add("hidden");
    $("inspectBlock").hidden = true;
    $("resultBlock").classList.add("hidden");
    $("resultBlock").hidden = true;
  $("buildBtn").disabled = !file || !state.templates;
  $("statusHint").textContent = file ? "Reading Frequency…" : "Choose a PPS to inspect it before building.";
}

async function inspectSelected() {
  if (!state.file) return;
  try {
    state.buffer = await state.file.arrayBuffer();
    state.inspect = await inspectPps(state.buffer, state.file.name);
    renderInspect(state.inspect);
    $("buildBtn").disabled = !state.templates;
    $("statusHint").textContent = "Looks readable. Add the four tabs, then download.";
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
  if (info.priceGrid?.length) bits.push(`Price grid: <b>${info.priceGrid.join(", ")}</b>.`);
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
  if (!state.buffer || !state.templates) return;
  $("buildBtn").disabled = true;
  $("statusHint").textContent = "Building four tabs…";
  try {
    const { buffer, report } = await buildPps(state.buffer, state.file.name, {
      templates: state.templates,
    });
    state.output = buffer;
    state.outputName = outputFilename(state.file.name);
    renderResult(report);
    $("downloadBtn").disabled = false;
    $("statusHint").textContent = `Ready: ${state.outputName}`;
    toast("Four tabs added");
  } catch (err) {
    $("resultBlock").classList.remove("hidden");
    $("resultBlock").hidden = false;
    $("resultStats").innerHTML = stat("Failed", "Build", "bad");
    $("statusHint").textContent = err.message;
    toast(err.message);
  } finally {
    $("buildBtn").disabled = false;
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

function reset() {
  $("file").value = "";
  setFile(null);
  toast("Reset");
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

loadTemplates().catch((err) => {
  $("statusHint").textContent = err.message;
});
