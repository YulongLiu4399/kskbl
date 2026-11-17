"""把多天（或多份）导出表合并，生成一张周报图；也支持月报。

用法（conda 环境 data）:
    conda activate data
    python data_processing_weekly.py                        # 弹窗多选文件，再选输出目录
    python data_processing_weekly.py 01.xlsx 02.xlsx -o 输出  # 命令行指定
    python data_processing_weekly.py --period monthly        # 生成月报
    python data_processing_weekly.py --no-gui 01.xlsx 02.xlsx -o 输出

输入支持一次选中七天的多个 Excel 文件；脚本会逐个读取、按时间合并、去重、排序，
然后画一张 9 面板的综合图，输出到：
    <输出目录>/weekly_起_止/weekly_起_止.png      （月报为 monthly_YYYY-MM/...）

需要与 data_processing.py 放在同一目录（复用其中的数据清洗函数）。
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# 复用同目录下 data_processing.py 里的清洗函数与面板配置
sys.path.insert(0, str(Path(__file__).resolve().parent))
from data_processing import (  # noqa: E402
    STATION_PANELS,
    load_data,
    normalize_rain,
    plot_wind_video,
    setup_chinese_font,
)


# ---------------------------------------------------------------- 选择文件


def _tkinter_available() -> bool:
    try:
        import tkinter  # noqa: F401
        from tkinter import filedialog  # noqa: F401
    except Exception:
        return False
    return True


def _dialog_root():
    import tkinter as tk

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root


def pick_files_via_dialog(initial_dir: Path) -> list[Path]:
    """弹出多选文件框，返回选中的 Excel 列表；取消返回空列表。"""
    from tkinter import filedialog

    root = _dialog_root()
    try:
        chosen = filedialog.askopenfilenames(
            title="选择要合并的 Excel 文件（可按住 Ctrl 多选）",
            initialdir=str(initial_dir),
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
    finally:
        root.destroy()
    return [Path(p) for p in chosen]


def pick_dir_via_dialog(initial_dir: Path) -> Path | None:
    from tkinter import filedialog

    root = _dialog_root()
    try:
        chosen = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=str(initial_dir),
            mustexist=False,
        )
    finally:
        root.destroy()
    return Path(chosen) if chosen else None


def find_all_excel(folder: Path) -> list[Path]:
    files = [
        p
        for p in folder.iterdir()
        if p.is_file()
        and not p.name.startswith("~$")
        and p.suffix.lower() in {".xlsx", ".xls", ".xlsm"}
    ]
    return sorted(files, key=lambda p: p.name)


def resolve_files(args_files: list[Path], use_gui: bool) -> list[Path]:
    if args_files:
        paths = []
        for item in args_files:
            if item.is_dir():
                paths.extend(find_all_excel(item))
            elif item.exists():
                paths.append(item)
            else:
                raise SystemExit(f"找不到文件：{item}")
        if not paths:
            raise SystemExit("没有可用的 Excel 文件")
        return paths

    if use_gui and _tkinter_available():
        picked = pick_files_via_dialog(Path(__file__).resolve().parent)
        if not picked:
            raise SystemExit("已取消选择文件")
        return picked

    raise SystemExit("没有指定输入文件（命令行给文件，或去掉 --no-gui 用弹窗选择）")


def resolve_outdir(arg: Path | None, first_file: Path, use_gui: bool) -> Path:
    if arg is not None:
        return arg
    if use_gui and _tkinter_available():
        picked = pick_dir_via_dialog(first_file.parent)
        if picked is not None:
            return picked
    return first_file.parent


# ---------------------------------------------------------------- 合并数据


def merge_files(paths: list[Path], stats: dict | None = None) -> pd.DataFrame:
    frames = []
    per_file = []
    for path in paths:
        part = load_data(path, zero_rain_at_start=False)
        per_file.append((path.name, len(part)))
        print(f"  {path.name}: {len(part)} 条")
        frames.append(part)

    df = pd.concat(frames, ignore_index=True)
    before = len(df)
    df = df.drop_duplicates(subset=["时间"], keep="last")
    df = df.sort_values("时间").reset_index(drop=True)
    df = normalize_rain(df)  # 合并后从整段窗口起点归零，各文件之间连续累计
    removed = before - len(df)
    if removed:
        print(f"去掉重复时间戳 {removed} 条")
    if stats is not None:
        stats["per_file"] = per_file
        stats["duplicates"] = removed
        stats["merged_rows"] = len(df)
    return df


def _cn_date(ts) -> str:
    """把时间写成“11 月 11 日”。"""
    return f"{ts.month} 月 {ts.day} 日"


def period_paths(df: pd.DataFrame, period: str) -> tuple[str, str]:
    """返回（文件夹名/文件名主干, 标题里的时间文字）。"""
    t0, t1 = df["时间"].min(), df["时间"].max()
    span = f"{_cn_date(t0)}—{_cn_date(t1)}"
    if period == "monthly":
        return f"monthly_{t0:%Y-%m}", f"{t0.year} 年 {t0.month} 月（{span}）"
    if period == "multi":
        return f"multi_{t0:%Y-%m-%d}_{t1:%Y-%m-%d}", span
    iso = t0.isocalendar()
    return (
        f"weekly_{t0:%Y-%m-%d}_{t1:%Y-%m-%d}",
        f"{iso.year} 年第 {iso.week} 周（{span}）",
    )


# ---------------------------------------------------------------- 画图


def _stat_text(series: pd.Series) -> str:
    s = series.dropna()
    if s.empty:
        return "无数据"
    return f"均值 {s.mean():.2f}\n最大 {s.max():.2f}\n最小 {s.min():.2f}"


def plot_period_panel(
    df: pd.DataFrame, outdir: Path, period: str, source_count: int, dpi: int
) -> Path:
    """一张纵向堆叠的综合图，标题标明周报/月报。"""
    panels = [
        {"kind": kind, "col": col, "label": label, "color": color}
        for col, label, color, kind in STATION_PANELS
        if col in df.columns
    ]

    pm_cols = [c for c in ("PM10", "PM2.5(ug/m³)") if c in df.columns]
    if pm_cols:
        idx = next(
            (i for i, p in enumerate(panels) if p["col"] == "光电总辐射(w/m2)"), -1
        )
        panels.insert(idx + 1, {"kind": "pm", "col": None, "label": "颗粒物 (μg/m³)", "color": None})

    n = len(panels)
    fig, axes = plt.subplots(n, 1, figsize=(12.5, 1.55 * n + 0.9), squeeze=False)
    axes = axes[:, 0]
    t0, t1 = df["时间"].min(), df["时间"].max()

    for ax, panel in zip(axes, panels):
        kind = panel["kind"]

        if kind == "pm":
            for col, color, name in (
                ("PM10", "#2c3e50", "PM10"),
                ("PM2.5(ug/m³)", "#c0392b", "PM2.5"),
            ):
                if col in df.columns:
                    ax.plot(df["时间"], df[col], lw=0.9, color=color, label=name)
            ax.legend(loc="upper left", frameon=False, fontsize=9, ncol=2, handlelength=1.6)
            stats = []
            for col, name in (("PM10", "PM10"), ("PM2.5(ug/m³)", "PM2.5")):
                if col in df.columns and df[col].notna().any():
                    s = df[col].dropna()
                    stats.append(f"{name} 均 {s.mean():.1f} 高 {s.max():.1f} 低 {s.min():.1f}")
            stat_text = "\n".join(stats)
        else:
            series = df[panel["col"]]
            if kind == "scatter":
                # 静风时风向无意义，过滤风速 < 0.2 m/s 的点
                mask = pd.Series(True, index=df.index)
                if "风速(m/s)" in df.columns:
                    positive = df["风速(m/s)"].fillna(0) >= 0.2
                    if positive.any():
                        mask = positive
                ax.scatter(df.loc[mask, "时间"], series[mask], s=3,
                           color=panel["color"], alpha=0.6, lw=0)
                ax.set_ylim(-10, 370)
                ax.set_yticks([0, 90, 180, 270, 360])
            else:
                style = "steps-post" if panel["col"] == "累计雨量(mm)" else "default"
                ax.plot(df["时间"], series, lw=0.8,
                        color=panel["color"], drawstyle=style)
                if panel["col"] == "累计雨量(mm)" and series.dropna().eq(0).all():
                    ax.set_ylim(-0.02, 0.2)
            # 累计雨量是累加量，均值/最值意义不大，右侧不再标注
            stat_text = "" if panel["col"] == "累计雨量(mm)" else _stat_text(series)

        ax.set_ylabel(panel["label"], fontsize=9.5)
        ax.set_xlim(t0, t1)
        ax.margins(x=0)
        ax.grid(alpha=0.28, lw=0.6)
        ax.tick_params(labelsize=8, length=3)
        ax.text(
            1.015, 0.5, stat_text, transform=ax.transAxes, ha="left", va="center",
            fontsize=8, color="#444444", linespacing=1.6, clip_on=False,
        )

    locator = mdates.AutoDateLocator(minticks=4, maxticks=9)
    formatter = mdates.DateFormatter("%m-%d\n%H:%M")
    for i, ax in enumerate(axes):
        ax.xaxis.set_major_locator(locator)
        ax.xaxis.set_major_formatter(formatter)
        ax.tick_params(labelbottom=(i == n - 1))
    axes[-1].set_xlabel("时间", fontsize=10)

    title_word = {"weekly": "周报", "monthly": "月报"}.get(period, "多日连续")
    _, span_text = period_paths(df, period)
    fig.suptitle(
        f"南京大学大气科学学院气象站 · {title_word} {span_text}",
        fontsize=17, fontweight="bold", y=0.988,
    )
    fig.text(
        0.07, 0.006,
        f"观测记录：{len(df)} 条；{df.attrs.get('processing_note', '未做重采样或平滑')}",
        fontsize=9, color="#555555",
    )
    fig.subplots_adjust(left=0.075, right=0.85, top=0.962, bottom=0.075, hspace=0.16)

    path = outdir / f"{outdir.name}.png"
    fig.savefig(path, dpi=dpi)
    plt.close(fig)
    return path


# ---------------------------------------------------------------- main


def main() -> None:
    parser = argparse.ArgumentParser(
        description="合并多份 Excel 数据，生成一张周报/月报图（不带参数会弹窗多选文件）"
    )
    parser.add_argument("files", nargs="*", type=Path, help="要合并的 Excel 文件或目录；省略则弹窗多选")
    parser.add_argument("-o", "--outdir", type=Path, default=None, help="输出目录；省略则弹窗选择")
    parser.add_argument("--period", choices=["weekly", "monthly", "multi"], default="weekly",
                        help="周报、月报或多日连续（任意天数合并）")
    parser.add_argument("--figs", default=None, help="要生成的内容：figure,data,video；默认 figure+data")
    parser.add_argument("--dpi", type=int, default=130, help="图片分辨率（默认 130）")
    parser.add_argument("--no-gui", action="store_true", help="不弹窗，必须用参数指定文件")
    args = parser.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")

    setup_chinese_font()

    use_gui = not args.no_gui
    files = resolve_files(args.files, use_gui)
    print(f"共选中 {len(files)} 个文件：")
    for p in files:
        print(f"  - {p}")

    stats: dict = {}
    df = merge_files(files, stats=stats)

    base_outdir = resolve_outdir(args.outdir, files[0], use_gui)
    folder_name, span_text = period_paths(df, args.period)
    outdir = base_outdir / folder_name
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"输出：{outdir}")
    print(f"合并后 {len(df)} 条，时间范围 {df['时间'].min():%Y-%m-%d %H:%M} ~ {df['时间'].max():%Y-%m-%d %H:%M}")

    # 简单检查时间断档，方便发现漏传的文件
    gaps = df["时间"].diff().dropna()
    big = gaps[gaps > pd.Timedelta(minutes=10)]
    if not big.empty:
        print(f"注意：有 {len(big)} 处超过 10 分钟的空档，最大 {big.max()}")

    selected = [c.strip() for c in args.figs.split(",")] if args.figs else ["figure", "data"]
    print("生成内容：" + "、".join(selected))

    made: list[Path] = []
    if "figure" in selected:
        path = plot_period_panel(df, outdir, args.period, len(files), args.dpi)
        made.append(path)
        print("已生成", path)

    # 同时保存一份合并、清洗后的数据，文件名与文件夹一致
    if "data" in selected:
        data_path = outdir / f"{folder_name}.csv"
        df.to_csv(data_path, index=False, encoding="utf-8-sig")
        made.append(data_path)
        print("已保存数据", data_path)

    if "video" in selected:
        path = plot_wind_video(df, outdir, args.dpi, prefix=f"{folder_name}_")
        if path:
            made.append(path)
        print("已生成" if path else "跳过（数据不足）", path or "plot_wind_video")

    # 处理日志：记录每个输入文件、去重、断档和输出清单
    gap_text = f"{len(big)} 处超过 10 分钟，最大 {big.max()}" if not big.empty else "无"
    log_lines = [
        f"运行时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
        f"脚本：{Path(__file__).name}",
        f"合并文件（{len(files)} 个）：",
        *[f"  {name}: {rows} 条" for name, rows in stats.get("per_file", [])],
        f"去重删除：{stats.get('duplicates', 0)} 条",
        f"合并后记录：{len(df)} 条",
        f"时间范围：{df['时间'].min():%Y-%m-%d %H:%M:%S} ~ {df['时间'].max():%Y-%m-%d %H:%M:%S}",
        f"时间断档：{gap_text}",
        "",
        "各列缺失数量：",
        df.isna().sum().to_string(),
        "",
        "参数：",
        f"  --period {args.period}",
        f"  --figs {','.join(selected)}",
        f"  --dpi {args.dpi}",
        "",
        "输出文件：",
        *[f"  {p.name}" for p in made],
    ]
    log_path = outdir / f"{folder_name}_处理日志.txt"
    log_path.write_text(chr(10).join(log_lines) + chr(10), encoding="utf-8")
    print("已保存日志", log_path)


if __name__ == "__main__":
    main()
