const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.26.2/full/";

const OPTIONS = [
  ["00", "气象站面板（总图）"],
  ["01", "概览"],
  ["02", "PM 与能见度"],
  ["03", "相关性热图"],
  ["04", "风向风速"],
  ["05", "分布直方图"],
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

// ---------------------------------------------------------------- UI

function buildChecks() {
  checks.innerHTML = OPTIONS.map(([code, label]) => `
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
  if (busy) {
    progress.value = 10;
  }
}

function pickFile(file) {
  if (!file) return;
  fileInput.files = toFileList(file);
  fileName.textContent = file.name;
  runBtn.disabled = false;
}

// 把拖入的文件塞回 input（DataTransfer）
function toFileList(file) {
  const dt = new DataTransfer();
  dt.items.add(file);
  return dt.files;
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
  progress.value = 75;
  try {
    const fontResp = await fetch("fonts/msyh.ttc");
    if (fontResp.ok) {
      const fontBuf = new Uint8Array(await fontResp.arrayBuffer());
      pyodide.FS.writeFile("/tmp/msyh.ttc", fontBuf);
    }
  } catch (err) {
    console.warn("字体加载失败，中文可能显示为方框", err);
  }

  const src = await (await fetch("data_processing.py?v=2")).text();
  pyodide.FS.writeFile("/home/pyodide/data_processing.py", src);

  progress.value = 90;
  setStatus("运行环境就绪");
  return pyodide;
}

// ---------------------------------------------------------------- 处理

function buildScript(figs, dpi) {
  const figList = figs.map(f => JSON.stringify(f)).join(", ");
  return `
import sys, types, json, base64, os
from pathlib import Path

# Pyodide 没有 subprocess / shutil.which，补两个空模块（处理脚本只在桌面版用到）
sys.modules["subprocess"] = types.ModuleType("subprocess")
_shutil = types.ModuleType("shutil")
_shutil.which = lambda *a, **k: None
_shutil.get_terminal_size = lambda *a, **k: os.terminal_size((80, 24))
sys.modules["shutil"] = _shutil

sys.path.insert(0, "/home/pyodide")
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
import data_processing as dp

font_path = "/tmp/msyh.ttc"
if os.path.exists(font_path):
    font_manager.fontManager.addfont(font_path)
    plt.rcParams["font.family"] = font_manager.FontProperties(fname=font_path).get_name()
    plt.rcParams["axes.unicode_minus"] = False

# Pyodide 里没有 openpyxl，前端已转成 CSV，这里用 read_csv 代替 read_excel
import pandas as pd
pd.read_excel = lambda path, **kwargs: pd.read_csv(path, dtype=str, encoding="utf-8-sig")

df = dp.load_data(Path("/tmp/input.csv"))
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
  const file = fileInput.files[0];
  if (!file) return;
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
    // 用 SheetJS 在前端解析 xlsx，转成 CSV 交给 Python（免装 openpyxl）
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    py.FS.writeFile("/tmp/input.csv", new TextEncoder().encode(csv));

    setStatus("正在处理，手机上大约需要 10–30 秒 ...");
    const result = py.runPython(buildScript(figs, 120));
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
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) {
    fileName.textContent = fileInput.files[0].name;
    runBtn.disabled = false;
  }
});

["dragenter", "dragover"].forEach(evt =>
  drop.addEventListener(evt, e => { e.preventDefault(); drop.classList.add("dragover"); })
);
["dragleave", "drop"].forEach(evt =>
  drop.addEventListener(evt, e => { e.preventDefault(); drop.classList.remove("dragover"); })
);
drop.addEventListener("drop", e => {
  const f = e.dataTransfer.files[0];
  if (f) pickFile(f);
});

runBtn.addEventListener("click", run);

buildChecks();

// PWA
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  });
}
