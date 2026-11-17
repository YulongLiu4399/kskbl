#!/usr/bin/env python

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import animation, font_manager


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


DEFAULT_INPUT = (
    _app_dir()
    / ".reasonix/attachments/clipboard-20261112-124020.889436-000001.xlsx"
)
DEFAULT_OUTDIR = _app_dir() / "charts"

__version__ = "1.1.0"
APP_NAME = "气象数据处理工具"

SUMMARY_LABELS = {"最大值", "最小值", "平均值", "最大值(max)", "最小值(min)"}
SUMMARY_KEYWORDS = ("最大值", "最小值", "平均值", "平均", "总计", "合计", "汇总")
TIME_COLUMN_NAMES = {"时间", "时间戳", "日期时间", "datetime", "time"}
# 气象设备常见的缺测哨兵值：低于 -999 或高于 99999 一律当空值
MISSING_LOW = -999
MISSING_HIGH = 99999

LABELS = {
    "空气湿度(%)": ("空气湿度", "%"),
    "温度(℃)": ("温度", "℃"),
    "电源(V)": ("电源电压", "V"),
    "气压(KPa)": ("气压", "KPa"),
    "信号强度(dBm)": ("信号强度", "dBm"),
    "PM10": ("PM10", "μg/m³"),
    "风速(m/s)": ("风速", "m/s"),
    "累计雨量(mm)": ("累计雨量", "mm"),
    "风向(°)": ("风向", "°"),
    "光电总辐射(w/m2)": ("光电总辐射", "W/m²"),
    "PM2.5(ug/m³)": ("PM2.5", "μg/m³"),
    "能见度(m)": ("能见度", "m"),
}

OVERVIEW_COLUMNS = [
    "温度(℃)",
    "空气湿度(%)",
    "气压(KPa)",
    "PM2.5(ug/m³)",
    "PM10",
    "能见度(m)",
    "风速(m/s)",
    "光电总辐射(w/m2)",
    "电源(V)",
    "信号强度(dBm)",
    "累计雨量(mm)",
]


def setup_chinese_font() -> None:
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in ("Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Source Han Sans SC"):
        if name in installed:
            plt.rcParams["font.family"] = name
            break
    plt.rcParams["axes.unicode_minus"] = False


setup_chinese_font()




def _find_time_column(columns) -> str:
    """优先按列名识别时间列，找不到才退回第一列。"""
    for col in columns:
        if str(col).strip().lower() in TIME_COLUMN_NAMES:
            return col
    return columns[0]


def _is_summary_label(value) -> bool:
    text = str(value).strip()
    return bool(text) and any(key in text for key in SUMMARY_KEYWORDS)


def load_data(path: Path, stats: dict | None = None, zero_rain_at_start: bool = True) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=0, dtype=str)
    raw.columns = [str(c).strip() for c in raw.columns]
    if stats is not None:
        stats["raw_rows"] = len(raw)

    time_col = _find_time_column(raw.columns)
    df = raw.copy()
    df[time_col] = df[time_col].astype(str).str.strip()
    before = len(df)
    df = df[~df[time_col].map(_is_summary_label)].copy()
    if stats is not None:
        stats["summary_rows"] = before - len(df)

    df["时间"] = pd.to_datetime(df[time_col], errors="coerce")
    before = len(df)
    df = df.dropna(subset=["时间"])
    if stats is not None:
        stats["bad_time_rows"] = before - len(df)
    if time_col != "时间":
        df = df.drop(columns=[time_col])

    text_cols = {"风向方位", "天气状况"}
    for col in df.columns:
        if col == "时间":
            continue
        if col in text_cols:
            df[col] = df[col].astype(str).str.strip().replace({"nan": np.nan, "": np.nan})
            numeric = pd.to_numeric(df[col], errors="coerce")
            if col == "天气状况" and numeric.notna().all():
                df[col] = numeric.astype(int)
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 缺测哨兵值（-999、999999 等）转成 NaN，避免拉爆曲线和统计量
    value_cols = [c for c in df.columns if c != "时间" and c not in text_cols]
    if value_cols:
        missing = df[value_cols].le(MISSING_LOW) | df[value_cols].ge(MISSING_HIGH)
        if stats is not None:
            stats["missing_values"] = int(missing.to_numpy().sum())
        df[value_cols] = df[value_cols].mask(missing)

    # 同一时间戳只保留最后一条
    before = len(df)
    df = df.drop_duplicates(subset=["时间"], keep="last")
    if stats is not None:
        stats["duplicates"] = before - len(df)
    df = df.sort_values("时间").reset_index(drop=True)

    if stats is not None:
        stats["missing_columns"] = [c for c in OVERVIEW_COLUMNS if c not in df.columns]

    if zero_rain_at_start:
        df = normalize_rain(df)
    return df


def normalize_rain(df: pd.DataFrame) -> pd.DataFrame:
    """累计雨量以本次数据窗口的第一条记录为 0：只在文件起点归零，跨天继续累计。"""
    col = "累计雨量(mm)"
    if col not in df.columns or df.empty:
        return df

    df = df.copy()
    rain = df[col].dropna()
    if rain.empty:
        return df
    df[col] = (df[col] - rain.iloc[0]).clip(lower=0)
    return df


def _last_valid(series: pd.Series) -> float:
    s = series.dropna()
    return float(s.iloc[-1]) if not s.empty else np.nan


def _last_text(series: pd.Series):
    s = series.dropna()
    return s.iloc[-1] if not s.empty else np.nan


def _circular_mean(series: pd.Series) -> float:
    """风向圆周平均：359° 与 1° 应得到 0°，直接算术平均会错成 180°。"""
    s = series.dropna().to_numpy(dtype=float)
    if s.size == 0:
        return np.nan
    rad = np.deg2rad(s)
    return float(np.rad2deg(np.arctan2(np.sin(rad).mean(), np.cos(rad).mean())) % 360)


def resample_data(df: pd.DataFrame, minutes: int) -> pd.DataFrame:
    """把高频观测聚合到 N 分钟。

    - 累计雨量取区间末值（累计量不能平均）
    - 风向按圆周平均，避免 0°/360° 跨越出错
    - 风向方位、天气状况取区间内最后一个有效值
    - 其余数值列取均值
    """
    if not minutes or minutes <= 0 or df.empty:
        return df

    work = df.set_index("时间")
    agg = {}
    for col in work.columns:
        if col == "累计雨量(mm)":
            agg[col] = _last_valid
        elif col == "风向(°)":
            agg[col] = _circular_mean
        elif col in ("风向方位", "天气状况") or not pd.api.types.is_numeric_dtype(work[col]):
            agg[col] = _last_text
        else:
            agg[col] = "mean"

    out = work.resample(f"{int(minutes)}min").agg(agg).dropna(how="all").reset_index()
    note = f"已重采样至 {int(minutes)} 分钟"
    prev = df.attrs.get("processing_note")
    out.attrs = dict(df.attrs)
    out.attrs["processing_note"] = f"{prev}；{note}" if prev else note
    return out


def smooth_data(df: pd.DataFrame, window: int) -> pd.DataFrame:
    """滑动平均平滑；累计雨量、风向和文本列不参与，窗口自动取奇数。"""
    if not window or window <= 1 or df.empty:
        return df

    window = int(window)
    if window % 2 == 0:
        window += 1
    skip = {"累计雨量(mm)", "风向(°)", "风向方位", "天气状况"}
    out = df.copy()
    for col in out.columns:
        if col == "时间" or col in skip:
            continue
        if pd.api.types.is_numeric_dtype(out[col]):
            out[col] = out[col].rolling(window, center=True, min_periods=1).mean()

    note = f"已做 {window} 点滑动平均"
    prev = out.attrs.get("processing_note")
    out.attrs["processing_note"] = f"{prev}；{note}" if prev else note
    return out


def numeric_frame(df: pd.DataFrame) -> pd.DataFrame:
    return df.select_dtypes(include=[np.number])


def short_label(col: str) -> tuple[str, str]:
    return LABELS.get(col, (col, ""))


def with_unit(col: str) -> str:
    name, unit = short_label(col)
    return f"{name} ({unit})" if unit else name


def time_axis(ax, fmt: str = "%m-%d %H:%M", maxticks: int = 6) -> None:
    ax.xaxis.set_major_formatter(mdates.DateFormatter(fmt))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=3, maxticks=maxticks))
    for label in ax.get_xticklabels():
        label.set_rotation(30)
        label.set_ha("right")



STATION_PANELS = [
    ("温度(℃)", "温度 (℃)", "#c0392b", "line"),
    ("空气湿度(%)", "相对湿度 (%)", "#2471a3", "line"),
    ("气压(KPa)", "气压 (KPa)", "#6c5ce7", "line"),
    ("风速(m/s)", "风速 (m/s)", "#17a589", "line"),
    ("风向(°)", "风向 (°)", "#c9a227", "scatter"),
    ("光电总辐射(w/m2)", "太阳总辐射 (W/m²)", "#8e44ad", "line"),
    ("能见度(m)", "能见度 (m)", "#2471a3", "line"),
    ("累计雨量(mm)", "累计雨量 (mm)", "#27ae60", "line"),
]


def _stat_lines(series: pd.Series) -> str:
    s = series.dropna()
    if s.empty:
        return "无数据"
    return f"均值 {s.mean():.2f}\n最大 {s.max():.2f}\n最小 {s.min():.2f}"


def plot_station_panel(df: pd.DataFrame, outdir: Path, dpi: int, prefix: str = "") -> Path:
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

    wind_handles = []
    for ax, panel in zip(axes, panels):
        kind = panel["kind"]

        if kind == "pm":
            for col, color, name in (
                ("PM10", "#2c3e50", "PM10"),
                ("PM2.5(ug/m³)", "#c0392b", "PM2.5"),
            ):
                if col in df.columns:
                    ax.plot(df["时间"], df[col], lw=0.9, color=color, label=name)
            ax.legend(
                loc="upper left", frameon=False, fontsize=9, ncol=2, handlelength=1.6
            )
            stats = []
            for col, name in (("PM10", "PM10"), ("PM2.5(ug/m³)", "PM2.5")):
                if col in df.columns and df[col].notna().any():
                    s = df[col].dropna()
                    stats.append(
                        f"{name} 均 {s.mean():.1f} 高 {s.max():.1f} 低 {s.min():.1f}"
                    )
            stat_text = "\n".join(stats)
        else:
            series = df[panel["col"]]
            if kind == "scatter":
                # 静风时风向没有物理意义：有风的点用主色，静风用浅灰小点，避免 0° 堆积
                if "风速(m/s)" in df.columns:
                    calm = df["风速(m/s)"].fillna(0) < 0.2
                else:
                    calm = pd.Series(False, index=df.index)
                if calm.any():
                    h_quiet = ax.scatter(df.loc[calm, "时间"], series[calm], s=3,
                                         color="#c8c8c8", alpha=0.7, lw=0, label="静风 (<0.2 m/s)")
                    wind_handles.append(h_quiet)
                h_active = ax.scatter(df.loc[~calm, "时间"], series[~calm], s=6,
                                      color=panel["color"], alpha=0.85, lw=0, label="有风 (≥0.2 m/s)")
                wind_handles.append(h_active)
                ax.set_ylim(-10, 370)
                ax.set_yticks([0, 90, 180, 270, 360])
            else:
                # 累计雨量是阶梯式累加，用 steps-post 更符合物理意义
                style = "steps-post" if panel["col"] == "累计雨量(mm)" else "default"
                ax.plot(df["时间"], series, lw=0.9,
                        color=panel["color"], drawstyle=style)
                if panel["col"] == "累计雨量(mm)" and series.dropna().eq(0).all():
                    ax.set_ylim(-0.02, 0.2)
            # 累计雨量是累加量，均值/最值意义不大，右侧不再标注
            stat_text = "" if panel["col"] == "累计雨量(mm)" else _stat_lines(series)

        ax.set_ylabel(panel["label"], fontsize=9.5)
        ax.set_xlim(t0, t1)
        ax.margins(x=0)
        ax.grid(alpha=0.28, lw=0.6)
        ax.tick_params(labelsize=8, length=3)
        ax.text(
            1.015, 0.5, stat_text, transform=ax.transAxes, ha="left", va="center",
            fontsize=8, color="#444444", linespacing=1.6, clip_on=False,
        )

    locator = mdates.HourLocator(byhour=range(0, 24, 2))
    formatter = mdates.DateFormatter("%m-%d\n%H:%M")
    for i, ax in enumerate(axes):
        ax.xaxis.set_major_locator(locator)
        ax.xaxis.set_major_formatter(formatter)
        ax.tick_params(labelbottom=(i == n - 1))
    axes[-1].set_xlabel("时间", fontsize=10)

    if wind_handles:
        fig.legend(handles=wind_handles, labels=[h.get_label() for h in wind_handles],
                   loc="lower center", bbox_to_anchor=(0.5, 0.042),
                   ncol=2, frameon=False, fontsize=9)

    fig.suptitle(
        f"南京大学大气科学学院气象站（{t0:%Y-%m-%d %H:%M:%S}--{t1:%Y-%m-%d %H:%M:%S}）",
        fontsize=17, fontweight="bold", y=0.988,
    )
    fig.text(
        0.07, 0.006,
        f"绘制记录：{len(df)} 条；{df.attrs.get('processing_note', '未做重采样或平滑')}",
        fontsize=9, color="#555555",
    )
    fig.subplots_adjust(left=0.075, right=0.85, top=0.962, bottom=0.105, hspace=0.16)

    stem = prefix.rstrip("_") if prefix else "00_气象站面板"
    path = outdir / f"{stem}.png"
    fig.savefig(path, dpi=dpi)
    plt.close(fig)
    return path




def plot_overview(df: pd.DataFrame, outdir: Path, dpi: int, prefix: str = "") -> Path:
    cols = [c for c in OVERVIEW_COLUMNS if c in df.columns]
    ncols = 3
    nrows = int(np.ceil(len(cols) / ncols))
    fig, axes = plt.subplots(
        nrows, ncols, figsize=(15, 2.5 * nrows), constrained_layout=True
    )
    axes = np.atleast_1d(axes).ravel()
    t0, t1 = df["时间"].min(), df["时间"].max()

    for ax, col in zip(axes, cols):
        series = df[col]
        ax.plot(df["时间"], series, lw=1.0, color="#1f77b4")
        ax.set_title(with_unit(col), fontsize=10)
        ax.set_xlim(t0, t1)
        ax.grid(alpha=0.3)
        ax.tick_params(labelsize=8)
        if series.nunique(dropna=True) <= 1:
            value = series.dropna().iloc[0] if series.notna().any() else 0
            ax.set_ylim(value - 1, value + 1)
            ax.text(
                0.5, 0.5, f"全程恒定：{value:g}", transform=ax.transAxes,
                ha="center", va="center", fontsize=10, color="gray",
            )
        elif series.notna().any():
            mean = series.mean()
            ax.axhline(mean, color="#d62728", ls="--", lw=0.8, alpha=0.7)
            ax.text(
                0.995, 0.95, f"均值 {mean:.2f}", transform=ax.transAxes,
                ha="right", va="top", fontsize=8, color="#d62728",
            )

    for ax in axes[len(cols):]:
        ax.set_visible(False)

    for idx, ax in enumerate(axes[: len(cols)]):
        time_axis(ax)
        ax.tick_params(labelbottom=idx // ncols == nrows - 1)

    fig.suptitle("环境监测数据概览（红色虚线为均值）", fontsize=13)
    path = outdir / f"{prefix}01_概览.png"
    fig.savefig(path, dpi=dpi)
    plt.close(fig)
    return path




def plot_pm(df: pd.DataFrame, outdir: Path, dpi: int, prefix: str = "") -> Path | None:
    have = [c for c in ("PM2.5(ug/m³)", "PM10", "能见度(m)") if c in df.columns]
    if len(have) < 2:
        return None

    fig, axes = plt.subplots(
        2, 1, figsize=(13, 7), gridspec_kw={"height_ratios": (2, 1)}, constrained_layout=True
    )

    ax = axes[0]
    for col, color in zip(
        [c for c in ("PM2.5(ug/m³)", "PM10") if c in df.columns],
        ["#d62728", "#ff7f0e"],
    ):
        ax.plot(df["时间"], df[col], lw=1.0, color=color, label=with_unit(col))
    ax.set_ylabel("浓度 (μg/m³)")
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(alpha=0.3)
    ax.set_title("颗粒物浓度与能见度")

    ax2 = axes[1]
    if "能见度(m)" in df.columns:
        ax2.plot(df["时间"], df["能见度(m)"], lw=1.0, color="#2ca02c")
        ax2.set_ylabel("能见度 (m)")
        ax2.grid(alpha=0.3)
        if "PM2.5(ug/m³)" in df.columns:
            ax3 = ax2.twinx()
            ax3.plot(df["时间"], df["PM2.5(ug/m³)"], lw=0.8, color="#d62728", alpha=0.45)
            ax3.set_ylabel("PM2.5 (μg/m³)", color="#d62728")
            ax3.tick_params(axis="y", colors="#d62728")
        ax2.set_title("能见度（叠加 PM2.5 作对照）")
    elif {"PM2.5(ug/m³)", "PM10"}.issubset(df.columns):
        ratio = df["PM2.5(ug/m³)"] / df["PM10"].replace(0, np.nan)
        ax2.plot(df["时间"], ratio, lw=1.0, color="#2ca02c")
        ax2.set_ylabel("PM2.5 / PM10")
        ax2.set_title("PM2.5 占 PM10 的比例")
        ax2.grid(alpha=0.3)
    time_axis(ax2)
    axes[0].tick_params(labelbottom=False)
    t0, t1 = df["时间"].min(), df["时间"].max()
    for ax_ in (ax, ax2):
        ax_.set_xlim(t0, t1)

    path = outdir / f"{prefix}02_PM与能见度.png"
    fig.savefig(path, dpi=dpi)
    plt.close(fig)
    return path




def plot_correlation(df: pd.DataFrame, outdir: Path, dpi: int, prefix: str = "") -> Path | None:
    num = numeric_frame(df).drop(columns=["天气状况"], errors="ignore")
    num = num.loc[:, num.notna().sum() > 2]
    num = num.loc[:, num.std(numeric_only=True) > 0]
    if num.shape[1] < 2:
        return None

    corr = num.corr(method="pearson")
    labels = [short_label(c)[0] for c in corr.columns]

    fig, ax = plt.subplots(figsize=(8.5, 7))
    im = ax.imshow(corr.values, cmap="RdBu_r", vmin=-1, vmax=1)
    ax.set_xticks(range(len(labels)), labels, rotation=45, ha="right", fontsize=9)
    ax.set_yticks(range(len(labels)), labels, fontsize=9)
    for i in range(len(labels)):
        for j in range(len(labels)):
            v = corr.values[i, j]
            if np.isfinite(v):
                ax.text(
                    j, i, f"{v:.2f}", ha="center", va="center", fontsize=7,
                    color="white" if abs(v) > 0.6 else "black",
                )
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04, label="Pearson 相关系数")
    ax.set_title("指标相关性矩阵")
    fig.tight_layout()
    path = outdir / f"{prefix}03_相关性热图.png"
    fig.savefig(path, dpi=dpi)
    plt.close(fig)
    return path




def plot_wind(df: pd.DataFrame, outdir: Path, dpi: int, prefix: str = "") -> Path | None:
    if not {"风速(m/s)", "风向(°)"}.issubset(df.columns):
        return None

    wind = df[["时间", "风速(m/s)", "风向(°)"]].dropna()
    if wind.empty:
        return None
    # 静风时风向无物理意义：有风的点正常显示，静风用浅灰小点，图例区分
    calm = (wind["风速(m/s)"] < 0.2).to_numpy()
    active, quiet = wind[~calm], wind[calm]

    fig = plt.figure(figsize=(13, 5.5), constrained_layout=True)
    ax1 = fig.add_subplot(1, 2, 1, projection="polar")
    ax1.set_theta_zero_location("N")
    ax1.set_theta_direction(-1)
    theta = np.deg2rad(wind["风向(°)"].to_numpy())
    r = wind["风速(m/s)"].to_numpy()
    h_active = ax1.scatter(theta[~calm], r[~calm], s=14, alpha=0.5, c=r[~calm], cmap="viridis",
                           label="有风 (≥0.2 m/s)")
    h_quiet = None
    if not quiet.empty:
        h_quiet = ax1.scatter(theta[calm], r[calm], s=5, alpha=0.5, color="#c8c8c8",
                              label="静风 (<0.2 m/s)")
    ax1.set_title("风向-风速散点（极坐标，0°=北）")
    ax1.set_rlabel_position(112.5)
    ax1.tick_params(labelsize=8)

    ax2 = fig.add_subplot(1, 2, 2)
    # 用散点而不是折线：0°/360° 是同一方向，折线会在 359°→1° 处画出贯穿整图的斜线
    ax2.scatter(active["时间"], active["风向(°)"], s=4, alpha=0.6, color="#9467bd", lw=0,
                label="有风 (≥0.2 m/s)")
    if not quiet.empty:
        ax2.scatter(quiet["时间"], quiet["风向(°)"], s=3, alpha=0.6, color="#c8c8c8", lw=0,
                    label="静风 (<0.2 m/s)")
    ax2.set_ylim(0, 360)
    ax2.set_yticks([0, 90, 180, 270, 360])
    ax2.set_ylabel("风向 (°，0°=北，顺时针)")
    ax2.set_title("风向随时间变化")
    ax2.grid(alpha=0.3)
    time_axis(ax2)

    handles = [h_active] + ([h_quiet] if h_quiet is not None else [])
    fig.legend(handles=handles, labels=[h.get_label() for h in handles],
               loc="lower center", ncol=2, frameon=False, fontsize=9)

    path = outdir / f"{prefix}04_风向风速.png"
    fig.savefig(path, dpi=dpi)
    plt.close(fig)
    return path




def _ffmpeg_exe() -> str | None:
    exe = shutil.which("ffmpeg")
    if exe:
        return exe
    base = Path(getattr(sys, "_MEIPASS", _app_dir()))
    for candidate in (base / "ffmpeg.exe", Path(sys.executable).parent / "Library" / "bin" / "ffmpeg.exe"):
        if candidate.exists():
            return str(candidate)
    return None


def plot_wind_video(
    df: pd.DataFrame, outdir: Path, dpi: int, prefix: str = "",
    fps: int = 8, max_frames: int = 120,
) -> Path | None:
    cols = ["时间", "风向(°)", "风速(m/s)"]
    if not set(cols).issubset(df.columns):
        return None

    wind = df[cols].dropna().reset_index(drop=True)
    if len(wind) < 10:
        return None

    frame_idx = np.linspace(0, len(wind) - 1, min(max_frames, len(wind))).astype(int)
    theta = np.deg2rad(wind["风向(°)"].to_numpy())
    speed = wind["风速(m/s)"].to_numpy()
    times = wind["时间"].to_numpy()

    fig = plt.figure(figsize=(12, 5.5), constrained_layout=True)
    ax_polar = fig.add_subplot(1, 2, 1, projection="polar")
    ax_polar.set_theta_zero_location("N")
    ax_polar.set_theta_direction(-1)
    rmax = max(float(speed.max()) * 1.2, 1.0)
    ax_polar.set_ylim(0, rmax)
    ax_polar.set_title("当前风向（0°=北，顺时针）", fontsize=10)
    trail, = ax_polar.plot([], [], ".", ms=2, color="#bbbbbb", alpha=0.6)
    arrow, = ax_polar.plot([], [], "-", lw=2.2, color="#c0392b")
    info = ax_polar.text(
        0.5, -0.12, "", transform=ax_polar.transAxes,
        ha="center", va="top", fontsize=9, color="#444444",
    )

    ax_line = fig.add_subplot(1, 2, 2)
    ax_line.plot(wind["时间"], wind["风向(°)"], lw=0.6, color="#999999")
    ax_line.set_ylim(0, 360)
    ax_line.set_yticks([0, 90, 180, 270, 360])
    ax_line.set_ylabel("风向 (°)")
    ax_line.set_title("风向随时间变化（红点=当前时刻）", fontsize=10)
    ax_line.grid(alpha=0.3)
    cursor = ax_line.axvline(wind["时间"].iloc[0], color="#c0392b", lw=1.2)
    point, = ax_line.plot([], [], "o", color="#c0392b", ms=5)
    time_axis(ax_line)

    t0, t1 = wind["时间"].min(), wind["时间"].max()
    fig.suptitle(
        f"南京大学大气科学学院气象站 · 风向变化（{t0:%Y-%m-%d %H:%M}--{t1:%Y-%m-%d %H:%M}）",
        fontsize=14, fontweight="bold",
    )

    def update(n: int):
        i = frame_idx[n]
        now = pd.Timestamp(times[i])
        trail.set_data(theta[: i + 1], speed[: i + 1])
        arrow.set_data([theta[i], theta[i]], [0, rmax * 0.85])
        cursor.set_xdata([now, now])
        point.set_data([now], [speed[i]])
        info.set_text(
            f"{now:%m-%d %H:%M}   风向 {wind['风向(°)'].iloc[i]:.0f}°   "
            f"风速 {speed[i]:.1f} m/s"
        )
        return trail, arrow, cursor, point, info

    anim = animation.FuncAnimation(fig, update, frames=len(frame_idx), interval=1000 / fps)

    ffmpeg = _ffmpeg_exe()
    if ffmpeg:
        matplotlib.rcParams["animation.ffmpeg_path"] = ffmpeg
        path = outdir / f"{prefix}06_风向变化.mp4"
        writer = animation.FFMpegWriter(
            fps=fps, bitrate=2400,
            extra_args=["-pix_fmt", "yuv420p", "-profile:v", "main", "-movflags", "+faststart"],
        )
        anim.save(path, writer=writer, dpi=dpi)
    else:
        path = outdir / f"{prefix}06_风向变化.gif"
        anim.save(path, writer=animation.PillowWriter(fps=fps), dpi=dpi)

    plt.close(fig)
    return path


def plot_distributions(df: pd.DataFrame, outdir: Path, dpi: int, prefix: str = "") -> Path | None:
    cols = [c for c in OVERVIEW_COLUMNS if c in df.columns and df[c].notna().any()]
    cols = [c for c in cols if df[c].nunique() > 1]
    if not cols:
        return None

    ncols = 4
    nrows = int(np.ceil(len(cols) / ncols))
    fig, axes = plt.subplots(
        nrows, ncols, figsize=(14, 2.4 * nrows), constrained_layout=True
    )
    axes = np.atleast_1d(axes).ravel()

    for ax, col in zip(axes, cols):
        s = df[col].dropna()
        ax.hist(s, bins=40, color="#1f77b4", alpha=0.8, edgecolor="none")
        ax.axvline(s.mean(), color="#d62728", ls="--", lw=1.0)
        ax.axvline(s.median(), color="#2ca02c", ls=":", lw=1.0)
        ax.set_title(
            f"{short_label(col)[0]}  n={len(s)}  "
            f"均值={s.mean():.2f} 中位数={s.median():.2f}",
            fontsize=9,
        )
        ax.tick_params(labelsize=7)
        ax.grid(alpha=0.25)

    for ax in axes[len(cols):]:
        ax.set_visible(False)

    fig.suptitle("各指标分布（红虚线=均值，绿点线=中位数）", fontsize=13)
    path = outdir / f"{prefix}05_分布直方图.png"
    fig.savefig(path, dpi=dpi)
    plt.close(fig)
    return path




def _tkinter_available() -> bool:
    try:
        import tkinter
        from tkinter import filedialog
    except Exception:
        return False
    return True


def _dialog_root():
    import tkinter as tk

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    return root


def pick_file_via_dialog(initial_dir: Path) -> Path | None:
    from tkinter import filedialog

    root = _dialog_root()
    try:
        chosen = filedialog.askopenfilename(
            title="选择气象数据 Excel 文件",
            initialdir=str(initial_dir),
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
    finally:
        root.destroy()
    return Path(chosen) if chosen else None


def pick_dir_via_dialog(initial_dir: Path) -> Path | None:
    from tkinter import filedialog

    root = _dialog_root()
    try:
        chosen = filedialog.askdirectory(
            title="选择图表输出目录",
            initialdir=str(initial_dir),
            mustexist=False,
        )
    finally:
        root.destroy()
    return Path(chosen) if chosen else None


def find_latest_excel(folder: Path) -> Path | None:
    files = [
        p
        for p in folder.iterdir()
        if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in {".xlsx", ".xls", ".xlsm"}
    ]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def resolve_input(arg: Path | None, use_gui: bool) -> Path:
    if arg is not None:
        if arg.is_dir():
            found = find_latest_excel(arg)
            if found is None:
                raise SystemExit(f"目录里没有 Excel 文件：{arg}")
            print(f"目录中最新的 Excel：{found.name}")
            return found
        if not arg.exists():
            raise SystemExit(f"找不到输入文件：{arg}")
        return arg

    if use_gui and _tkinter_available():
        picked = pick_file_via_dialog(DEFAULT_INPUT.parent)
        if picked is None:
            raise SystemExit("已取消选择输入文件")
        return picked

    if DEFAULT_INPUT.exists():
        return DEFAULT_INPUT
    return find_latest_excel(_app_dir()) or DEFAULT_INPUT


def resolve_outdir(arg: Path | None, input_path: Path, use_gui: bool) -> Path:
    if arg is not None:
        return arg

    if use_gui and _tkinter_available():
        picked = pick_dir_via_dialog(input_path.parent / "charts")
        if picked is not None:
            return picked

    return DEFAULT_OUTDIR


# ---------------------------------------------------------------- 生成哪些内容

FIGURE_CHOICES = [
    ("00", "气象站面板（总图）"),
    ("01", "概览"),
    ("02", "PM 与能见度"),
    ("03", "相关性热图"),
    ("04", "风向风速"),
    ("05", "分布直方图"),
    ("video", "风向变化动画（MP4/GIF）"),
    ("data", "清洗后的数据表"),
]
FIGURE_LABELS = dict(FIGURE_CHOICES)
PLOTTERS = {
    "00": plot_station_panel,
    "01": plot_overview,
    "02": plot_pm,
    "03": plot_correlation,
    "04": plot_wind,
    "05": plot_distributions,
    "video": plot_wind_video,
}


def pick_figures_via_dialog() -> list[str] | None:
    import tkinter as tk

    root = tk.Tk()
    root.title("选择要生成的内容")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    vars_ = {}
    for code, name in FIGURE_CHOICES:
        var = tk.BooleanVar(value=(code != "video"))
        vars_[code] = var
        tk.Checkbutton(root, text=f"{code}  {name}", variable=var, anchor="w").pack(
            fill="x", padx=16, pady=2
        )

    result: list[str] | None = None

    def confirm() -> None:
        nonlocal result
        result = [code for code, var in vars_.items() if var.get()]
        root.destroy()

    tk.Button(root, text="开始生成", command=confirm, width=16).pack(pady=10)
    root.mainloop()
    return result


def resolve_figures(spec: str | None, use_gui: bool) -> list[str]:
    valid = [code for code, _ in FIGURE_CHOICES]
    if spec:
        chosen = [c.strip() for c in spec.split(",") if c.strip()]
        bad = [c for c in chosen if c not in valid]
        if bad:
            raise SystemExit(f"未知的编号：{', '.join(bad)}；可用：{', '.join(valid)}")
        return chosen

    if use_gui and _tkinter_available():
        picked = pick_figures_via_dialog()
        if picked is None:
            raise SystemExit("已取消生成")
        return picked

    return [c for c in valid if c != "video"]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="绘制环境监测 Excel 数据的图表（不带参数会弹窗选择输入/输出）"
    )
    parser.add_argument(
        "input", nargs="?", type=Path, default=None,
        help="Excel 文件或所在目录；省略则弹窗选择",
    )
    parser.add_argument(
        "-o", "--outdir", type=Path, default=None,
        help="图表输出目录；省略则弹窗选择（取消则用脚本旁的 charts/）",
    )
    parser.add_argument("--dpi", type=int, default=130, help="图片分辨率（默认 130）")
    parser.add_argument("--resample", type=int, default=0, metavar="分钟",
                        help="重采样间隔（分钟），0 表示不重采样")
    parser.add_argument("--smooth", type=int, default=0, metavar="点数",
                        help="滑动平均窗口（建议奇数，>=3），0 表示不平滑")
    parser.add_argument(
        "--figs", default=None,
        help="要生成的内容，逗号分隔：00,01,02,03,04,05,data；默认全部",
    )
    parser.add_argument("--no-gui", action="store_true", help="不弹窗，使用默认选项")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")

    setup_chinese_font()

    use_gui = not args.no_gui
    input_path = resolve_input(args.input, use_gui)

    stats: dict = {}
    df = load_data(input_path, stats=stats)

    if args.resample and args.resample > 0:
        df = resample_data(df, args.resample)
        print(f"已重采样至 {args.resample} 分钟，共 {len(df)} 条")
    if args.smooth and args.smooth > 1:
        df = smooth_data(df, args.smooth)
        print(f"已做 {args.smooth} 点滑动平均")

    day = df["时间"].max().strftime("%Y-%m-%d")
    outdir = resolve_outdir(args.outdir, input_path, use_gui) / day
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"输入：{input_path}")
    print(f"输出：{outdir}")
    span = f"{df['时间'].min():%Y-%m-%d %H:%M} ~ {df['时间'].max():%Y-%m-%d %H:%M}"
    print(f"有效记录 {len(df)} 条，时间范围 {span}")
    print(f"数值列缺失情况：\n{df.isna().sum().to_string()}")

    selected = resolve_figures(args.figs, use_gui)
    print("生成内容：" + "、".join(f"{c} {FIGURE_LABELS[c]}" for c in selected))

    made: list[Path] = []
    for code, name in FIGURE_CHOICES:
        if code not in selected:
            continue
        if code == "data":
            data_path = outdir / f"{day}.csv"
            df.to_csv(data_path, index=False, encoding="utf-8-sig")
            made.append(data_path)
            print("已保存数据", data_path)
            continue
        path = PLOTTERS[code](df, outdir, args.dpi, prefix=f"{day}_")
        if path:
            made.append(path)
        print("已生成" if path else "跳过（数据不足）", path or PLOTTERS[code].__name__)

    missing_cols = [c for c in OVERVIEW_COLUMNS if c not in df.columns]
    log_lines = [
        f"运行时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
        f"脚本：{Path(__file__).name}",
        f"输入文件：{input_path}",
        f"原始行数：{stats.get('raw_rows', 0)}",
        f"剔除汇总行：{stats.get('summary_rows', 0)}",
        f"时间无法解析：{stats.get('bad_time_rows', 0)}",
        f"去重删除：{stats.get('duplicates', 0)}",
        f"缺测值替换：{stats.get('missing_values', 0)}",
        f"有效记录：{len(df)}",
        f"时间范围：{df['时间'].min():%Y-%m-%d %H:%M:%S} ~ {df['时间'].max():%Y-%m-%d %H:%M:%S}",
        "",
        "各列缺失数量：",
        df.isna().sum().to_string(),
        "",
        "参数：",
        f"  --figs {','.join(selected)}",
        f"  --dpi {args.dpi}",
        f"  --resample {args.resample}",
        f"  --smooth {args.smooth}",
    ]
    if missing_cols:
        log_lines += ["", "缺少的指标列：", *[f"  {c}" for c in missing_cols]]
    log_lines += ["", "输出文件：", *[f"  {p.name}" for p in made]]

    log_path = outdir / f"{day}_处理日志.txt"
    log_path.write_text(chr(10).join(log_lines) + chr(10), encoding="utf-8")
    print("已保存日志", log_path)


if __name__ == "__main__":
    main()
