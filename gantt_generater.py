# -*- coding: utf-8 -*-
# Gantt from Excel — clean left border (single black line), grid clipped to chart, legend bottom-right outside
# 改善重點：
# 1) 不再用 name 當 y 軸索引，避免重複/NaN 被折疊 → 保證每列都會畫
# 2) Excel 讀取指定 openpyxl，缺少時提示安裝
# 3) 日期解析更健壯（YYYY/MM/DD、YYYY-MM-DD、YYYY.MM.DD、MM/DD 等）
# 4) 名稱清洗（strip / NaN → 任務N），可選擇附加 (2)、(3) 避免視覺重名

import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Patch
from matplotlib.gridspec import GridSpec
from datetime import datetime, timedelta, date
from matplotlib import font_manager, rcParams
from matplotlib.lines import Line2D
from collections import defaultdict

# -------------------- FONT --------------------
# Windows 建議直接用微軟正黑體；若要固定字型檔，填路徑 (e.g. r"C:\Windows\Fonts\msjh.ttc")
FONT_FILE = None
if FONT_FILE:
    assert os.path.exists(FONT_FILE), f"Font file not found: {FONT_FILE}"
    font_manager.fontManager.addfont(FONT_FILE)
    rcParams["font.family"] = font_manager.FontProperties(fname=FONT_FILE).get_name()
else:
    rcParams["font.sans-serif"] = ["Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", "SimHei"]
rcParams["axes.unicode_minus"] = False

# -------------------- DEFINE (可調參數) --------------------
INPUT_XLSX = "tasks_gantt.xlsx"   # Excel（sheet: tasks；欄位 name/start/(end或days)/progress(0~1 可省)）
OUTPUT_IMG = "output.png"         # 輸出圖

SHOW_TITLE = False
TITLE = "專案時程"

FIG_WIDTH_INCH = 15
ROW_HEIGHT_INCH = 0.3     # 每列高度 (inch)
TOP_BOTTOM_MARGIN = 0.5   # 上下額外留白 (inch)
LEFT_LABEL_COL_RATIO  = 1.65        # 左欄（主題）寬
RIGHT_CHART_COL_RATIO = 8.35        # 右欄（圖表）寬

LEFT_LABEL_FONT_SIZE = 10
HEADER_FONT_SIZE = LEFT_LABEL_FONT_SIZE
BAR_HEIGHT = 0.55

# 顏色
COLOR_LABEL_TEXT    = "#000000"
COLOR_HEADER_TEXT   = "#A6AEC5"
COLOR_WEEKEND       = "#FFF0F0"
COLOR_TODAY         = "#FF0000"
COLOR_BAR_PLANNED   = "#3DB9D3"
COLOR_BAR_DONE_100  = "#4a0ce7"
COLOR_BAR_DONE_PART = "#2898B0"
GRID_COLOR_MAJOR    = "#d0d0d0"     # 主要格線（每 N 天 / 每列）
GRID_COLOR_MINOR    = "#eaeaea"     # 每日細格線

# 時間軸 / 版面
MAX_XLABELS = 16                # 日期文字最多顯示數量（格線仍每日畫）
DRAW_WEEKENDS = True
DRAW_TODAY_BAND = True
TODAY_DATE = date.today()
TAIL_EXTRA_DAYS = 3             # x 軸尾端額外多留天數

# 圖例（放圖表外右下）
LEGEND_BBOX = (1.005, -0.04)
SUBPLOT_BTM = 0.17

SHOW_PROCESS = False
CUSTOM_START_DATE = None   # e.g. "2025-01-01"
CUSTOM_END_DATE   = None   # e.g. "2025-03-31"
LABEL_DEDUP_SUFFIX = True  # 若名稱重複，是否自動加 (2)、(3)...

# -------------------- 小工具 --------------------
def _parse_date(x):
    """盡量容錯解析日期，支援 YYYY/MM/DD、YYYY-MM-DD、YYYY.MM.DD、MM/DD（自動補今年）"""
    if isinstance(x, datetime):
        return x
    if isinstance(x, (date, )):
        return datetime(x.year, x.month, x.day)

    s = str(x).strip()
    if not s or s.lower() == "nan":
        raise ValueError("空白日期")

    # 年月日
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    # 月/日 → 補今年
    for fmt in ("%m/%d", "%m-%d", "%m.%d"):
        try:
            d = datetime.strptime(s, fmt)
            return datetime(TODAY_DATE.year, d.month, d.day)
        except ValueError:
            pass

    raise ValueError(f"無法解析日期: {s}")

def _daterange(d0, d1):
    cur = d0
    while cur <= d1:
        yield cur
        cur += timedelta(days=1)

def _tick_step(total_units: int, fig_width_inch: float, max_labels: int = MAX_XLABELS) -> int:
    if total_units <= 0:
        return 1
    labels_per_inch = 1.2
    allowed_labels = int(fig_width_inch * labels_per_inch)
    allowed_labels = min(allowed_labels, max_labels)
    if allowed_labels <= 0:
        allowed_labels = 1
    return max(1, (total_units + allowed_labels - 1) // allowed_labels)

# -------------------- 主流程 --------------------
def render_gantt_from_excel(xlsx_path, out_png):
    # 讀 Excel（要求 openpyxl）
    try:
        df = pd.read_excel(xlsx_path, sheet_name="tasks", engine="openpyxl")
    except ImportError:
        raise SystemExit(
            "缺少 openpyxl：請先在目前虛擬環境執行\n"
            "    python -m pip install openpyxl\n"
        )
    except FileNotFoundError:
        raise SystemExit(f"找不到檔案：{xlsx_path}")
    except ValueError as e:
        raise SystemExit(f"讀取工作表錯誤：{e}（請確認存在 sheet: 'tasks'）")

    # 欄位檢查
    required_any = (("end" in df.columns) or ("days" in df.columns))
    if "name" not in df.columns or "start" not in df.columns or not required_any:
        raise SystemExit("Excel 需包含欄位：name, start, 並且 (end 或 days) 至少其一。可選欄位：progress(0~1)")

    # 清洗欄位
    # 名稱：去空白，NaN → 任務N；可選擇去重後加 (2)、(3) 做視覺辨識
    names_raw = [str(x).strip() if pd.notna(x) else "" for x in df["name"].tolist()]
    names = []
    seen = defaultdict(int)
    for i, n in enumerate(names_raw):
        label = n if n else f"任務{i+1}"
        if LABEL_DEDUP_SUFFIX:
            seen[label] += 1
            names.append(label if seen[label] == 1 else f"{label} ({seen[label]})")
        else:
            names.append(label)

    # 解析日期欄
    try:
        df["start"] = df["start"].apply(_parse_date)
        if "end" in df.columns and "days" in df.columns:
            df["end"] = df.apply(
                lambda r: _parse_date(r["end"]) if pd.notna(r["end"])
                        else r["start"] + pd.to_timedelta(int(r["days"]) - 1, unit="D"),
                axis=1
            )
        elif "end" in df.columns:
            df["end"] = df["end"].apply(_parse_date)
        else:  # 有 days
            df["end"] = df["start"] + pd.to_timedelta(df["days"].astype(int) - 1, unit="D")

        if "progress" not in df.columns:
            df["progress"] = 0.0
        else:
            # 轉成 [0,1] 區間的浮點數
            df["progress"] = pd.to_numeric(df["progress"], errors="coerce").fillna(0.0)
            df["progress"] = df["progress"].clip(lower=0.0, upper=1.0)
    except Exception as e:
        raise SystemExit(f"日期/欄位解析失敗：{e}")

    # 組 records
    tasks = df[["start", "end", "progress"]].to_dict("records")
    if not tasks:
        raise SystemExit("Excel 裡沒有任務資料（sheet: tasks）")

    # 時間範圍
    proj_start = min(t["start"] for t in tasks)
    proj_end   = max(t["end"]   for t in tasks)

    # 手動覆蓋
    if CUSTOM_START_DATE:
        proj_start = _parse_date(CUSTOM_START_DATE)
    if CUSTOM_END_DATE:
        proj_end   = _parse_date(CUSTOM_END_DATE)

    # 畫布高度依列數調整
    num_tasks = len(tasks)
    FIG_HEIGHT_INCH = ROW_HEIGHT_INCH * num_tasks + TOP_BOTTOM_MARGIN

    # ===== 兩欄版面：左主題 / 右圖表 =====
    fig = plt.figure(figsize=(FIG_WIDTH_INCH, FIG_HEIGHT_INCH))
    gs = GridSpec(1, 2, width_ratios=[LEFT_LABEL_COL_RATIO, RIGHT_CHART_COL_RATIO], figure=fig)
    ax_left  = fig.add_subplot(gs[0, 0])
    ax_chart = fig.add_subplot(gs[0, 1])

    # y 範圍（用 enumerate 索引確保每列一行）
    ymin, ymax = -0.5, num_tasks - 0.5
    ax_left.set_ylim(ymin, ymax)
    ax_chart.set_ylim(ymin, ymax)
    ax_chart.invert_yaxis()
    ax_left.invert_yaxis()

    # ===== 左欄（主題）=====
    ax_left.set_xlim(0, 1)
    ax_left.set_xticks([])
    y_pos = list(range(num_tasks))
    ax_left.set_yticks(y_pos)
    ax_left.set_yticklabels(names, fontsize=LEFT_LABEL_FONT_SIZE, color=COLOR_LABEL_TEXT, ha="left")
    for lab in ax_left.get_yticklabels():
        lab.set_x(0.07)
    ax_left.set_title("Marc時程規劃", fontsize=HEADER_FONT_SIZE, color=COLOR_HEADER_TEXT, loc="left", pad=10)
    for s in ["top", "bottom", "left", "right"]:
        ax_left.spines[s].set_visible(False)
    ax_left.tick_params(axis="y", length=0)

    # ===== 右欄（圖表）=====
    ax_chart.xaxis.tick_top()
    ax_chart.xaxis.set_label_position("top")
    for s in ["left", "top", "right", "bottom"]:
        ax_chart.spines[s].set_visible(False)

    # x 範圍（尾端加 TAIL_EXTRA_DAYS）
    x_left  = mdates.date2num(proj_start) - 0.5
    x_right = mdates.date2num(proj_end + timedelta(days=TAIL_EXTRA_DAYS)) + 1.5
    ax_chart.set_xlim(x_left, x_right)

    # ==== 底層背景：週末、今天 ====
    if DRAW_WEEKENDS:
        for d in _daterange(proj_start, proj_end + timedelta(days=TAIL_EXTRA_DAYS)):
            if d.weekday() >= 5:
                x0 = mdates.date2num(datetime(d.year, d.month, d.day))
                if x_left < x0 < x_right:
                    ax_chart.axvspan(x0, x0 + 1, facecolor=COLOR_WEEKEND, zorder=0, linewidth=0)

    if DRAW_TODAY_BAND and proj_start.date() <= TODAY_DATE <= (proj_end + timedelta(days=TAIL_EXTRA_DAYS)).date():
        xt0 = mdates.date2num(datetime(TODAY_DATE.year, TODAY_DATE.month, TODAY_DATE.day))
        if x_left < xt0 < x_right:
            ax_chart.axvspan(xt0, xt0 + 1, facecolor=COLOR_TODAY, zorder=0, linewidth=0)

    # ==== Bars（已完成/未完成）====
    for idx, t in enumerate(tasks):
        s = mdates.date2num(t["start"])
        e = mdates.date2num(t["end"]) + 1  # inclusive
        y = idx - BAR_HEIGHT/2
        p = float(t.get("progress", 0.0))
        if p >= 1.0:
            ax_chart.broken_barh([(s, e - s)], (y, BAR_HEIGHT), facecolor=COLOR_BAR_DONE_100, zorder=2)
        else:
            if SHOW_PROCESS:
                done = s + (e - s) * p
                if p > 0:
                    ax_chart.broken_barh([(s, done - s)], (y, BAR_HEIGHT), facecolor=COLOR_BAR_DONE_PART, zorder=2)
                ax_chart.broken_barh([(done, e - done)], (y, BAR_HEIGHT), facecolor=COLOR_BAR_PLANNED, zorder=2)
            else:
                ax_chart.broken_barh([(s, e - s)], (y, BAR_HEIGHT), facecolor=COLOR_BAR_PLANNED, zorder=2)

    # ==== 直橫格線：僅在表格內畫（避免越界）====
    EPS = 1e-7
    total_days = ((proj_end + timedelta(days=TAIL_EXTRA_DAYS)).date() - proj_start.date()).days + 1
    major_step = _tick_step(total_days, FIG_WIDTH_INCH)

    # 日期文字（只放 major）
    ax_chart.xaxis.set_major_locator(mdates.DayLocator(interval=major_step))
    ax_chart.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))

    # 每日細格線（自行畫）
    cur = proj_start
    while cur <= proj_end + timedelta(days=TAIL_EXTRA_DAYS):
        x = mdates.date2num(datetime(cur.year, cur.month, cur.day))
        if (x_left + EPS) < x < (x_right - EPS):
            ax_chart.vlines(x, ymin, ymax, colors=GRID_COLOR_MINOR, linewidth=0.9, zorder=1)
        cur += timedelta(days=1)

    # 每 N 天粗格線（major）
    cur = proj_start
    idx = 0
    while cur <= proj_end + timedelta(days=TAIL_EXTRA_DAYS):
        if idx % major_step == 0:
            x = mdates.date2num(datetime(cur.year, cur.month, cur.day))
            if (x_left + EPS) < x < (x_right - EPS):
                ax_chart.vlines(x, ymin, ymax, colors=GRID_COLOR_MAJOR, linewidth=1.2, zorder=1)
        cur += timedelta(days=1)
        idx += 1

    # 橫線：每列上下邊界（限制在表格內）
    for i in range(num_tasks + 1):
        ax_chart.hlines(i - 0.5, x_left + EPS, x_right - EPS, colors=GRID_COLOR_MAJOR, linewidth=1.2, zorder=1)

    # ===== 黑色外框（確保在最上層）=====
    fig.add_artist(Line2D([x_left,  x_right], [ymax,  ymax], transform=ax_chart.transData, color="black", lw=1.2, zorder=999))
    fig.add_artist(Line2D([x_left,  x_right], [ymin,  ymin], transform=ax_chart.transData, color="black", lw=1.2, zorder=999))
    fig.add_artist(Line2D([x_left,  x_left], [ymin,  ymax], transform=ax_chart.transData, color="black", lw=2.0, zorder=999))
    fig.add_artist(Line2D([x_right, x_right], [ymin,  ymax], transform=ax_chart.transData, color="black", lw=1.2, zorder=999))

    # 右欄 y 軸不顯示標籤
    ax_chart.set_yticks([])

    # Title（可開關）
    if SHOW_TITLE:
        fig.suptitle(TITLE, fontsize=12, y=0.98)

    # 邊距
    fig.subplots_adjust(left=0.07, right=0.97, top=0.93, bottom=SUBPLOT_BTM)
    fig.tight_layout()

    # ===== 用 figure 座標一次畫完外框與中間分隔線 =====
    bL = ax_left.get_position()
    bC = ax_chart.get_position()

    x_left_outer   = bL.x0
    x_middle_split = bL.x1
    x_right_outer  = bC.x1
    y_bottom       = bL.y0
    y_top          = bL.y1

    # 橫向黑框（整排跨左右欄）
    for y in (y_top, y_bottom):
        fig.add_artist(Line2D([x_left_outer, x_right_outer], [y, y], transform=fig.transFigure, color="black", lw=1.2, zorder=10))

    # 左右側黑框直線
    fig.add_artist(Line2D([x_left_outer, x_left_outer], [y_bottom, y_top], transform=fig.transFigure, color="black", lw=1.2, zorder=10))
    fig.add_artist(Line2D([x_right_outer, x_right_outer], [y_bottom, y_top], transform=fig.transFigure, color="black", lw=1.2, zorder=10))

    # 灰色橫線（跨左右欄，從最左黑框到最右黑框）
    row_levels = [i - 0.5 for i in range(num_tasks + 1)]
    for ydata in row_levels:
        y_fig = ax_chart.transData.transform((0, ydata))[1]
        y_fig = fig.transFigure.inverted().transform((0, y_fig))[1]
        fig.add_artist(Line2D([x_left_outer, x_right_outer], [y_fig, y_fig], transform=fig.transFigure, color=GRID_COLOR_MAJOR, lw=1.0, zorder=5))

    # 輸出
    fig.savefig(out_png, dpi=180)
    plt.close(fig)

# 直接執行
if __name__ == "__main__":
    render_gantt_from_excel(INPUT_XLSX, OUTPUT_IMG)
