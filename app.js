const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.26.2/full/";

const DAILY_OPTIONS = [
  ["00", "气象站面板（总图）"],
  ["01", "概览"],
  ["02", "PM 与能见度"],
  ["03", "相关性热图"],
  ["04", "风向风速"],
  ["05", "分布直方图"],
  ["data", "清洗后的数据表（CSV）"],
];

const PERIOD_OPTIONS = [
  ["figure", "周报 / 月报总图"],
  ["data", "清洗后的数据表（CSV）"],
];

const fileInput = document.getElementById("file");
const fileName = document.getElementById("fileName");
const drop = document.getElementById("drop");
const checks = document.getElementById("checks");
const runBtn = document.getElementById("run");
const statusEl = document.getElementById("status");
const progress = document.getElementById("progress");
const results = document.getElementById("results");

let pyodide = null;

function currentMode() {
  const el = document.querySelector('input[name="mode"]:checked');
  return el ? el.value : "daily";
}

function isPeriod() {
  return currentMode() !== "daily";
}

// ---------------------------------------------------------------- UI

function buildChecks() {
  const options = isPeriod() ? PERIOD_OPTIONS : DAILY_OPTIONS;
  checks.innerHTML = options.map(([code, label]) => `
    <label class="check">
      <input type="checkbox" value="${code}" checked>
      <span>${label}</span>
    </label>
  `).join("");
}

function selectedFigs() {
  return [...checks.querySelectorAll("input:checked")].map(el => el.value);
}

function setStatus(text, isError = false) {
  statusEl.textContent = text || "";
  statusEl.className = "status" + (isError ? " err" : "");
}

function setBusy(busy) {
  runBtn.disabled = busy || !fileInput.files.length;
  progress.hidden = !busy;
  if (busy) progress.value = 10;
}

function updateFileLabel() {
  const files = [...fileInput.files];
  if (!files.length) {
    fileName.textContent = "";
    runBtn.disabled = true;
    return;
  }
  if (isPeriod()) {
    fileName.textContent = `已选 ${files.length} 个文件`;
  } else {
    fileName.textContent = files[0].name;
  }
  runBtn.disabled = false;
}

// ---------------------------------------------------------------- Pyodide

async function initPyodide() {
  if (pyodide) return pyodide;

  setStatus("正在下载运行环境（约 35 MB，仅首次）...");
  if (!window.loadPyodide) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = PYODIDE_URL + "pyodide.js";
      s.onload = resolve;
      s.onerror = () => reject(new Error("无法加载 Pyodide，请检查网络"));
      document.head.appendChild(s);
    });
  }

  pyodide = await loadPyodide({ indexURL: PYODIDE_URL });

  setStatus("安装 pandas / matplotlib ...");
  progress.value = 45;
  await pyodide.loadPackage(["pandas", "matplotlib"]);

  setStatus("加载中文字体 ...");
  progress.value = 70;
  try {
    const fontResp = await fetch("fonts/msyh.ttc");
    if (fontResp.ok) {
      pyodide.FS.writeFile("/tmp/msyh.ttc", new Uint8Array(await fontResp.arrayBuffer()));
    }
  } catch (err) {
    console.warn("字体加载失败", err);
  }

  setStatus("加载处理脚本 ...");
  progress.value = 85;
  const daily = await (await fetch("data_processing.py?v=3")).text();
  pyodide.FS.writeFile("/home/pyodide/data_processing.py", daily);
  const period = await (await fetch("data_processing_weekly.py?v=2")).text();
  pyodide.FS.writeFile("/home/pyodide/data_processing_weekly.py", period);

  progress.value = 90;
  setStatus("运行环境就绪");
  return pyodide;
}

// ---------------------------------------------------------------- 生成 Python 代码

function commonPrelude() {
  return `
import sys, types, json, base64, os
from pathlib import Path

sys.modules["subprocess"] = types.ModuleType("subprocess")
_shutil = types.ModuleType("shutil")
_shutil.which = lambda *a, **k: None
_shutil.get_terminal_size = lambda *a, **k: os.terminal_size((80, 24))
sys.modules["shutil"] = _shutil

sys.path.insert(0, "/home/pyodide")
import pandas as pd
pd.read_excel = lambda path, **kwargs: pd.read_csv(path, dtype=str, encoding="utf-8-sig")

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

font_path = "/tmp/msyh.ttc"
if os.path.exists(font_path):
    font_manager.fontManager.addfont(font_path)
    plt.rcParams["font.family"] = font_manager.FontProperties(fname=font_path).get_name()
plt.rcParams["axes.unicode_minus"] = False
`;
}

function buildDailyScript(figs, dpi) {
  const figList = figs.map(f => JSON.stringify(f)).join(", ");
  return commonPrelude() + `
import data_processing as dp

df = dp.load_data(Path("/tmp/input_0.csv"))
if df.empty:
    raise RuntimeError("表格里没有有效数据")

out = Path("/tmp/out")
out.mkdir(parents=True, exist_ok=True)
day = df["时间"].max().strftime("%Y-%m-%d")
figs = [${figList}]
made = {}

for code, label in dp.FIGURE_CHOICES:
    if code not in figs or code == "video":
        continue
    if code == "data":
        p = out / f"{day}.csv"
        df.to_csv(p, index=False, encoding="utf-8-sig")
        made[p.name] = ("csv", p.read_bytes())
        continue
    p = dp.PLOTTERS[code](df, out, ${dpi}, prefix=f"{day}_")
    if p:
        made[p.name] = ("png", p.read_bytes())

json.dumps({k: {"kind": v[0], "data": base64.b64encode(v[1]).decode()} for k, v in made.items()})
`;
}

function buildPeriodScript(figs, dpi, mode, count) {
  const figList = figs.map(f => JSON.stringify(f)).join(", ");
  return commonPrelude() + `
import data_processing_weekly as dw

files = [Path(f"/tmp/input_{i}.csv") for i in range(${count})]
df = dw.merge_files(files)
if df.empty:
    raise RuntimeError("合并后没有有效数据")

folder, _ = dw.period_paths(df, ${JSON.stringify(mode)})
out = Path(f"/tmp/{folder}")
out.mkdir(parents=True, exist_ok=True)
figs = [${figList}]
made = {}

if "figure" in figs:
    p = dw.plot_period_panel(df, out, ${JSON.stringify(mode)}, len(files), ${dpi})
    if p:
        made[p.name] = ("png", p.read_bytes())

if "data" in figs:
    p = out / f"{folder}.csv"
    df.to_csv(p, index=False, encoding="utf-8-sig")
    made[p.name] = ("csv", p.read_bytes())

json.dumps({k: {"kind": v[0], "data": base64.b64encode(v[1]).decode()} for k, v in made.items()})
`;
}

// ---------------------------------------------------------------- 处理

async function fileToCsv(file) {
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_csv(sheet);
}

function renderFiles(files) {
  results.innerHTML = "";
  const entries = Object.entries(files);
  if (!entries.length) {
    setStatus("没有生成任何文件，请检查勾选项", true);
    return;
  }
  for (const [name, item] of entries) {
    const bytes = Uint8Array.from(atob(item.data), c => c.charCodeAt(0));
    const type = item.kind === "png" ? "image/png" : "text/csv";
    const blob = new Blob([bytes], { type });
    const url = URL.createObjectURL(blob);

    const card = document.createElement("div");
    card.className = "shot";
    if (item.kind === "png") {
      card.innerHTML = `
        <a href="${url}" target="_blank" rel="noopener">
          <img src="${url}" alt="${name}">
        </a>
        <div class="bar"><span>${name}</span><a href="${url}" download="${name}">下载</a></div>`;
    } else {
      card.innerHTML = `
        <div class="bar" style="padding:18px 14px;">
          <span>${name}</span>
          <a href="${url}" download="${name}">下载 CSV</a>
        </div>`;
    }
    results.appendChild(card);
  }
  setStatus(`完成：${entries.length} 个文件`);
}

async function run() {
  const files = [...fileInput.files];
  if (!files.length) return;
  const figs = selectedFigs();
  if (!figs.length) {
    setStatus("至少勾选一项要生成的内容", true);
    return;
  }

  try {
    setBusy(true);
    results.innerHTML = "";
    const py = await initPyodide();

    progress.value = 95;
    setStatus("读取 Excel ...");
    const usePeriod = isPeriod();
    const list = usePeriod ? files : files.slice(0, 1);
    for (let i = 0; i < list.length; i++) {
      const csv = await fileToCsv(list[i]);
      py.FS.writeFile(`/tmp/input_${i}.csv`, new TextEncoder().encode(csv));
    }

    setStatus(usePeriod
      ? `正在合并 ${list.length} 个文件并出图，手机上大约需要 20–60 秒 ...`
      : "正在处理，手机上大约需要 10–30 秒 ...");

    const script = usePeriod
      ? buildPeriodScript(figs, 120, currentMode(), list.length)
      : buildDailyScript(figs, 120);
    const result = py.runPython(script);
    renderFiles(JSON.parse(result));
  } catch (err) {
    console.error(err);
    const msg = String(err && err.message ? err.message : err);
    setStatus("处理失败：" + msg.split("\n").slice(-3).join(" "), true);
  } finally {
    setBusy(false);
    progress.value = 100;
  }
}

// ---------------------------------------------------------------- 事件

drop.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", updateFileLabel);

["dragenter", "dragover"].forEach(evt =>
  drop.addEventListener(evt, e => { e.preventDefault(); drop.classList.add("dragover"); })
);
["dragleave", "drop"].forEach(evt =>
  drop.addEventListener(evt, e => { e.preventDefault(); drop.classList.remove("dragover"); })
);
drop.addEventListener("drop", e => {
  const files = [...e.dataTransfer.files];
  if (!files.length) return;
  const dt = new DataTransfer();
  files.forEach(f => dt.items.add(f));
  fileInput.files = dt.files;
  updateFileLabel();
});

document.querySelectorAll('input[name="mode"]').forEach(el =>
  el.addEventListener("change", () => {
    buildChecks();
    updateFileLabel();
  })
);

runBtn.addEventListener("click", run);

buildChecks();

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  });
}
