# -*- coding: utf-8 -*-
"""拉曼单谱成图（通用版，用户定稿样式）

用法：
    python raman_plot.py <数据.txt> [输出名主干] [x下限] [x上限]

例：
    python raman_plot.py <数据目录>/20261023-1-a.txt
    python raman_plot.py 20261029-2-b.txt 20261029-2-b 1000 2000

样式（用户逐条确认过的定稿）：
- X 轴 1000-2000，主刻度 500，次刻度 250；Y 轴无刻度无标签
- 上轴 + 右轴封闭方框，不带刻度
- 线条 #CEF297 低饱和绿，线宽 2.0
- 图例右上，内容 = 文件名主干，Arial 22
- 图例下方 gap=170 处一行 I<sub>D</sub>/I<sub>G</sub>=X.XX，Arial 22
- 页面裁到内容外接框 + 四周 150 设备单位（0.25 英寸）留白，即"紧凑页边距"

交付物：
- <输出名>.tif   300 dpi，尺寸 = 页面实际尺寸 x 300
- <输出名>.opju  Origin 工程源文件（含数据表和成图）

拟合只用来读 D/G 两个峰高算比值，图上不画任何拟合曲线。
口径：全谱 + 常量基线（自由参数）+ 两个 Gaussian（与 Origin Peak Analysis 一致）。
ID/IG 保留两位有效数字，四舍六入五成双。
"""
import os
import sys

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from origin_plot import OriginPlot          # noqa: E402
from raman_fit import recipe_A, sig2       # noqa: E402

OUT = HERE
LINE_RGB = "206,242,151"   # #CEF297 低饱和绿
LINE_W = 1000              # set %C -w 1000 -> Line.Width 2.0
KEEP_PAGE = "Book1"        # 用户自己手做的项目页，绝不删
MARGIN = 150               # 页边距，设备单位 1/600 英寸 -> 0.25 in -> 300dpi 下 75 px
DPI = 300


def read_xy(path):
    """读制表符/空格分隔、无表头的两列文本，x 统一翻成升序。"""
    pts = []
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            parts = line.split()
            if len(parts) >= 2:
                try:
                    pts.append((float(parts[0]), float(parts[1])))
                except ValueError:
                    pass
    if not pts:
        raise SystemExit("没有读到数据点: %s" % path)
    pts.sort(key=lambda p: p[0])
    return pts


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 1
    src = os.path.abspath(argv[1])
    stem = os.path.splitext(os.path.basename(src))[0]
    out_name = argv[2] if len(argv) > 2 else stem
    x_lo = float(argv[3]) if len(argv) > 3 else 1000.0
    x_hi = float(argv[4]) if len(argv) > 4 else 2000.0

    pts = read_xy(src)
    rows = [[x, y] for x, y in pts]
    print("文件:", src)
    print("点数:", len(rows), " x:", rows[0][0], "->", rows[-1][0])

    # 只取两个峰高，不画拟合曲线
    x = np.array([p[0] for p in pts])
    y = np.array([p[1] for p in pts])
    fit = recipe_A(x, y)
    hD, hG = fit["popt"][0], fit["popt"][3]
    ratio = hD / hG
    label = r"I\-(D)/I\-(G)=%s" % sig2(ratio)   # \-( ) 是 Origin 的下标标记
    print("hD=%.1f hG=%.1f  ID/IG=%.4f -> %s" % (hD, hG, ratio, sig2(ratio)))

    op = OriginPlot()
    for p in list(op.app.Pages):
        if p.Name != KEEP_PAGE:
            try:
                op.delete_page(p.Name)
            except Exception:
                pass

    book = op.new_book(rows, long_names=["Raman Shift", stem])
    graph = op.line_plot(book, ncols=1, colors=[LINE_RGB], widths=[LINE_W])
    op.style(
        xlabel=r"Raman Shift (cm\+(-1))",
        ylabel="Intensity (a.u.)",
        label_size=22,
        legend_size=22,
        legend_pos="tr",
    )
    op.box_axes(
        xrange=(x_lo, x_hi),
        xinc=500,
        hide_y_ticks=True,
        hide_y_labels=True,
        hide_top_ticks=True,
        x_minor=1,      # 主刻度 500，1 条次刻度 -> 落在 250
        y_minor=0,
    )
    op.axis_font(xtick_pt=22, font_index=1)     # 1 = Arial
    op.set_legend(r"\l(1) " + stem, size=22)
    op.place_legend("tr", pad=140)              # 图例变宽后重摆，避免贴边框
    op.place_y_title(gap=22)
    box = op.text_below_legend(label, size=22, font="Arial", gap=170)
    print("标注:", label, "->", box)

    # 裁掉多余页面留白，四周只留 MARGIN
    pw, ph = op.trim_page(margin=MARGIN)
    print("页面: %.0f x %.0f 设备单位 = %.2f x %.2f 英寸" % (pw, ph, pw / 600.0, ph / 600.0))

    # 出图前视觉自检：图例/标注不能压在曲线上（用户明确要求）
    ok, report = op.check_overlap()
    for nm, n, verdict in report:
        print("  自检 %-10s 深色像素 %5d  %s" % (nm, n, verdict))
    if not ok:
        # 标注压线：自动换个不压线的位置再验一次
        print("  自检未过，自动重摆标注…")
        op.place_idig_clear(label, size=22, font="Arial")
        ok, report = op.check_overlap()
        for nm, n, verdict in report:
            print("  复检 %-10s 深色像素 %5d  %s" % (nm, n, verdict))
    if not ok:
        print("!! 视觉自检未通过：图例或标注与曲线/轴线重叠，先调开再出图")
        print("!! 调法：op.place_legend('tr', pad=...) 或 text_below_legend(gap=...)")
        return 2
    print("视觉自检通过：图例和标注都没有压线")

    files, cm, px = op.export_dpi(OUT, name=out_name, dpi=DPI, types=("tif",))
    print("导出: %.2f x %.2f cm -> %d x %d px @ %d dpi" % (cm[0], cm[1], px[0], px[1], DPI))
    for f, ok, sz in files:
        print("  ", f, ok, sz)

    # .opju 源文件（必须反斜杠绝对路径）
    # 安全：已存在就拒绝，绝不覆盖用户自己的工程文件
    proj = os.path.join(OUT, out_name + ".opju")
    if os.path.exists(proj):
        print("工程已存在，跳过保存（避免覆盖）:", proj)
        print("  需要覆盖请先手动备份并改名，或传 --force")
        if "--force" not in sys.argv:
            ok, exists = False, True
        else:
            ok, exists = op.save_project(proj, overwrite=True)
    else:
        ok, exists = op.save_project(proj)
    print("工程:", proj, ok, exists)
    print("book:", book, " graph:", graph)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
