# -*- coding: utf-8 -*-
"""碳材料拉曼 D 带 / G 带双峰拟合（通用版）

用法
----
    python raman_fit.py <数据文件.txt> [输出.csv]

主口径 A（与 Origin Peak Analysis 默认设置一致，作为报告值）
- 窗口：全谱（数据本身的 x 范围）
- 基线：常量 b0，作为拟合参数放开（初值取窗口中位数）
- 线型：两个 Gaussian
- ID/IG = 拟合峰高比（峰高从基线算起）

对照口径 B（诊断用）
- 基线固定在两侧真实平坦区（左 1000-1100、右 1800-1900 的中位数连线），不参与拟合
- 窗口 1100-1800，两个 Gaussian

输出
----
- CSV 六列：x, raw, fit, D+base, G+base, baseline（主口径 A 的曲线）
"""
import os
import sys
from decimal import Decimal, ROUND_HALF_EVEN

import numpy as np
from scipy.optimize import curve_fit

# CSV 默认输出到脚本所在目录
DEFAULT_OUT_DIR = os.path.dirname(os.path.abspath(__file__))


def sig2(value):
    """保留两位有效数字，四舍六入五成双（banker's rounding，逢五取偶）。

    >>> sig2(0.67901)   -> '0.68'
    >>> sig2(0.54855)   -> '0.55'
    >>> sig2(0.125)     -> '0.12'   # 5 前一位是偶数 2，舍
    >>> sig2(0.135)     -> '0.14'   # 5 前一位是奇数 3，入
    """
    if value == 0:
        return "0.0"
    d = Decimal(repr(float(value)))
    q = Decimal(1).scaleb(d.adjusted() - 1)      # 第二位有效数字所在的量级
    return str(d.quantize(q, rounding=ROUND_HALF_EVEN))
ANCH_L = (1000.0, 1100.0)
ANCH_R = (1800.0, 1900.0)


def load(path):
    pts = []
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            p = line.split()
            if len(p) >= 2:
                try:
                    pts.append((float(p[0]), float(p[1])))
                except ValueError:
                    pass
    pts.sort()
    a = np.array(pts)
    return a[:, 0], a[:, 1]


def gauss(x, h, xc, w):
    return h * np.exp(-4.0 * np.log(2.0) * ((x - xc) / w) ** 2)


def lor(x, h, xc, w):
    return h / (1.0 + 4.0 * ((x - xc) / w) ** 2)


def med(x, y, lo, hi):
    m = (x >= lo) & (x <= hi)
    if not m.any():
        return float(np.median(y))
    return float(np.median(y[m]))


def guess_centers(x, y):
    """给峰位初值：在 D/G 窗口内各取最大值位置。"""
    def pick(lo, hi, fallback):
        m = (x >= lo) & (x <= hi)
        return float(x[m][np.argmax(y[m])]) if m.any() else fallback
    return pick(1280, 1450, 1360.0), pick(1500, 1650, 1575.0)


def recipe_A(x, y, shape=gauss):
    """Origin 口径：常量基线（自由参数）+ 两个 Gaussian，全谱窗口。"""
    xs, ys = x, y
    cD0, cG0 = guess_centers(x, y)

    def model(xx, hD, cD, wD, hG, cG, wG, b0):
        return b0 + shape(xx, hD, cD, wD) + shape(xx, hG, cG, wG)

    span = float(ys.max() - ys.min())
    p0 = [span * .5, cD0, 240, span * .7, cG0, 85, float(np.median(ys))]
    lo = [0, 1280, 40, 0, 1500, 30, -np.inf]
    hi = [np.inf, 1450, 400, np.inf, 1650, 200, np.inf]
    popt, pcov = curve_fit(model, xs, ys, p0=p0, bounds=(lo, hi), maxfev=600000)
    perr = np.sqrt(np.diag(pcov))
    fv = model(xs, *popt)
    r2 = 1 - np.sum((ys - fv) ** 2) / np.sum((ys - ys.mean()) ** 2)
    b0 = float(popt[6])
    return dict(popt=popt, perr=perr, r2=r2, xs=xs, ys=ys, model=model, f=shape,
                bg=lambda xx: np.full_like(np.asarray(xx, dtype=float), b0),
                tag="A  Origin 口径：常量基线 + Gaussian，全谱")


def recipe_B(x, y, shape=gauss):
    """对照口径：两侧平坦区锚定基线（固定），两个 Gaussian，窗口 1100-1800。"""
    xl, xr = 0.5 * sum(ANCH_L), 0.5 * sum(ANCH_R)
    yl, yr = med(x, y, *ANCH_L), med(x, y, *ANCH_R)
    sl = (yr - yl) / (xr - xl)
    ic = yl - sl * xl
    bg = lambda xx: ic + sl * np.asarray(xx, dtype=float)

    m = (x >= 1100.0) & (x <= 1800.0)
    xs, ys = x[m], y[m]
    cD0, cG0 = guess_centers(xs, ys)

    def model(xx, hD, cD, wD, hG, cG, wG):
        return bg(xx) + shape(xx, hD, cD, wD) + shape(xx, hG, cG, wG)

    span = float(ys.max() - ys.min())
    p0 = [span * .5, cD0, 230, span * .7, cG0, 95]
    lo = [0, 1280, 40, 0, 1500, 30]
    hi = [np.inf, 1450, 400, np.inf, 1650, 200]
    popt, pcov = curve_fit(model, xs, ys, p0=p0, bounds=(lo, hi), maxfev=600000)
    perr = np.sqrt(np.diag(pcov))
    fv = model(xs, *popt)
    r2 = 1 - np.sum((ys - fv) ** 2) / np.sum((ys - ys.mean()) ** 2)
    return dict(popt=popt, perr=perr, r2=r2, xs=xs, ys=ys, model=model, f=shape,
                bg=bg, tag="B  对照：平坦区锚定基线 + Gaussian，窗口 1100-1800")


def report(r, shape_name="Gaussian"):
    hD, cD, wD, hG, cG, wG = r["popt"][:6]
    eD, eG = r["perr"][0], r["perr"][3]
    k = np.sqrt(np.pi / (4 * np.log(2))) if shape_name == "Gaussian" else np.pi / 2
    print("  " + r["tag"])
    print("    基线: bg@D(%.1f)=%.1f  bg@G(%.1f)=%.1f" %
          (cD, float(r["bg"](cD)), cG, float(r["bg"](cG))))
    print("    D: 峰位=%8.2f+-%.2f  半高宽=%7.2f+-%.2f  峰高=%8.1f+-%.1f" %
          (cD, r["perr"][1], wD, r["perr"][2], hD, eD))
    print("    G: 峰位=%8.2f+-%.2f  半高宽=%7.2f+-%.2f  峰高=%8.1f+-%.1f" %
          (cG, r["perr"][4], wG, r["perr"][5], hG, eG))
    ratio = hD / hG
    print("    R2=%.5f   ID/IG(峰高)=%.4f -> %s (两位有效数字)   ID/IG(面积)=%.4f" %
          (r["r2"], ratio, sig2(ratio), (hD * wD * k) / (hG * wG * k)))
    return ratio


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 1
    src = argv[1]
    if not os.path.exists(src):
        print("找不到文件:", src)
        return 1
    out = argv[2] if len(argv) > 2 else os.path.join(
        DEFAULT_OUT_DIR, os.path.splitext(os.path.basename(src))[0] + "_fit.csv")

    x, y = load(src)
    print("file:", src)
    print("N=%d  x %.1f..%.1f  y %.1f..%.1f" % (len(x), x[0], x[-1], y.min(), y.max()))
    print()
    rA = recipe_A(x, y)
    rB = recipe_B(x, y)
    a = report(rA)
    print()
    b = report(rB)
    print()
    print("主口径 ID/IG = %.4f -> %s ；对照口径 = %.4f -> %s" % (a, sig2(a), b, sig2(b)))

    f, bg = rA["f"], rA["bg"]
    hD, cD, wD, hG, cG, wG = rA["popt"][:6]
    xs, ys = rA["xs"], rA["ys"]
    rows = np.column_stack([
        xs, ys,
        bg(xs) + f(xs, hD, cD, wD) + f(xs, hG, cG, wG),
        bg(xs) + f(xs, hD, cD, wD),
        bg(xs) + f(xs, hG, cG, wG),
        bg(xs),
    ])
    np.savetxt(out, rows, delimiter=",",
               header="x,raw,fit,D+base,G+base,baseline", comments="")
    print("saved", out)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
