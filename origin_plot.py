# -*- coding: utf-8 -*-
"""Origin 自动化绘图工具（OriginPro 2021 + COM/LabTalk）

用法要点
--------
- Origin 已经打开时直接 Dispatch("Origin.ApplicationSI")，会连上同一个可见实例。
- 数据用 PutWorksheet 一次性写入（2 万行约 0.6 秒），比逐格赋值快得多。
- 所有作图用 LabTalk 的 Execute() 完成。
- 导出前必须先让目标图成为活动窗口（win -a 图名），否则 expGraph 返回 False。
- 路径用正斜杠 /，带空格的目录也能正常导出。
- set %C 每调用一次就前进到下一条曲线，按颜色1、线宽1、颜色2、线宽2 的顺序连续发即可。
- 图例位置用 COM 的 Left/Top（页面设备单位，1/600 英寸）设置，LabTalk 的 legend.x/y 不生效。
"""

import os
import win32com.client as win32


class OriginPlot:
    def __init__(self, visible=True):
        self.app = win32.Dispatch("Origin.ApplicationSI")
        self._visible = visible

    # ---------- 基础 ----------
    def execute(self, cmd):
        """执行一条 LabTalk 命令，返回 Origin 的返回值。"""
        return self.app.Execute(cmd)

    def pages(self):
        return [(p.Name, p.Type) for p in self.app.Pages]

    def graphs(self):
        return [p.Name for p in self.app.Pages if p.Type == 3]

    def books(self):
        return [p.Name for p in self.app.Pages if p.Type == 2]

    def _new_page(self, cmd):
        before = set(p.Name for p in self.app.Pages)
        self.app.Execute(cmd)
        for p in self.app.Pages:
            if p.Name not in before:
                return p.Name
        raise RuntimeError("新建页面失败: " + cmd)

    # ---------- 数据 ----------
    def new_book(self, rows, long_names=None, short_name=None):
        """把二维列表 rows 写入新工作簿，返回工作簿名。

        rows: [[x, y1, y2, ...], ...]
        long_names: 每列的长名称（作图时作为轴标题和图例）
        """
        cmd = 'newbook option:=lsname;'
        if short_name:
            cmd = 'newbook name:="%s" option:=lsname;' % short_name
        book = self._new_page(cmd)
        self.app.PutWorksheet(book, [[float(v) for v in row] for row in rows])
        if long_names:
            for i, nm in enumerate(long_names, start=1):
                self.app.Execute('wks.col%d.lname$="%s";' % (i, nm))
        return book

    def delete_page(self, name):
        self.app.Execute('win -c "%s";' % name)

    # ---------- 作图 ----------
    def line_plot(self, book, ncols, colors=None, widths=None):
        """以第 1 列为 X，第 2..ncols+1 列各画一条线，返回图名。

        ncols 是曲线条数，不是列号。
        """
        self.app.Execute('win -a "%s";' % book)
        before = set(p.Name for p in self.app.Pages)
        self.app.Execute('plotxy iy:=(1,2:%d) plot:=200;' % (ncols + 1))
        gname = None
        for p in self.app.Pages:
            if p.Name not in before and p.Type == 3:
                gname = p.Name
        if gname is None:
            gname = self.graphs()[-1]
        self.app.Execute('win -a "%s"; layer -a;' % gname)
        # 关键：set %C 每调用一次就前进到下一条曲线，
        # 所以按「颜色1、线宽1、颜色2、线宽2…」的顺序连续发即可，
        # 不要用 layer -i / set [book]1!n，那两个语法在这里不生效。
        for i in range(ncols):
            clr = colors[i] if colors and i < len(colors) else "auto"
            wid = widths[i] if widths and i < len(widths) else None
            self.app.Execute('set %%C -cl color(%s);' % clr)
            if wid:
                self.app.Execute('set %%C -w %d;' % wid)
        return gname

    def style(self, xlabel=None, ylabel=None, xrange=None, yrange=None,
              label_size=26, legend_size=20, legend_pos="tr", legend_pad=120):
        """设置轴标题、范围、字号，并把图例放到不压数据的位置。

        legend_pos: "tr" 右上 / "tl" 左上 / "br" 右下 / "bl" 左下
        """
        if xrange:
            self.app.Execute('layer.x.from = %g; layer.x.to = %g;' % xrange)
        if yrange:
            self.app.Execute('layer.y.from = %g; layer.y.to = %g;' % yrange)
        if xlabel:
            self.app.Execute('xb.text$ = "%s";' % xlabel)
        if ylabel:
            self.app.Execute('yl.text$ = "%s";' % ylabel)
        self.app.Execute('xb.fsize = %d; yl.fsize = %d;' % (label_size, label_size))
        self.app.Execute('legend.fsize = %d;' % legend_size)
        if legend_pos:
            self.place_legend(legend_pos, legend_pad)

    def place_legend(self, where="tr", pad=120):
        """按页面设备单位（1/600 英寸）摆放图例，避开数据区。"""
        pg = self.app.ActivePage
        ly = pg.Layers(0)
        lgn = None
        for i in range(ly.GraphObjects.Count):
            o = ly.GraphObjects(i)
            if o.Name == "Legend":
                lgn = o
                break
        if lgn is None:
            return
        PW, PH = pg.Width, pg.Height          # 英寸
        lL, lT, lW, lH = [ly.GetNumProp(x) for x in ("Left", "Top", "Width", "Height")]
        x0, y0 = lL / 100.0 * PW * 600, lT / 100.0 * PH * 600
        x1, y1 = (lL + lW) / 100.0 * PW * 600, (lT + lH) / 100.0 * PH * 600
        w, h = lgn.Width, lgn.Height
        if where == "tr":
            lgn.Left, lgn.Top = int(x1 - w - pad), int(y0 + pad)
        elif where == "tl":
            lgn.Left, lgn.Top = int(x0 + pad), int(y0 + pad)
        elif where == "br":
            lgn.Left, lgn.Top = int(x1 - w - pad), int(y1 - h - pad)
        else:
            lgn.Left, lgn.Top = int(x0 + pad), int(y1 - h - pad)

    def box_axes(self, xrange=None, xinc=None, yrange=None,
                 hide_y_ticks=True, hide_y_labels=False, hide_top_ticks=True,
                 x_minor=0, y_minor=0):
        """画成封闭方框：四边轴线都在，但上轴/右轴不带刻度。

        - layer.x/y.showAxes = 3   四边全开（1=只下/左，2=只上/右，3=全部）
        - layer.y.tickLength = 0   左轴刻度线去掉（标签另由 showLabels 控制）
        - layer.y2.tickLength = 0  右轴刻度线去掉
        - layer.x2.tickLength = 0  上轴刻度线去掉
        - layer.y.showLabels = 0   左轴数字标签隐藏（实测可写可回读）
        - x_minor / y_minor       两主刻度之间的次刻度条数；主刻度 500 时
                                  1 表示次刻度落在 250 处
        """
        self.app.Execute('layer.x.showAxes = 3; layer.y.showAxes = 3;')
        if xrange:
            self.app.Execute('layer.x.from = %g; layer.x.to = %g;' % xrange)
        if xinc:
            self.app.Execute('layer.x.inc = %g;' % xinc)
        if yrange:
            self.app.Execute('layer.y.from = %g; layer.y.to = %g;' % yrange)
        if hide_y_ticks:
            self.app.Execute('layer.y.tickLength = 0; layer.y2.tickLength = 0;')
        if hide_top_ticks:
            self.app.Execute('layer.x2.tickLength = 0;')
        if hide_y_labels:
            self.app.Execute('layer.y.showLabels = 0;')
        self.app.Execute('layer.x.minorTicks = %d; layer.y.minorTicks = %d;'
                         % (x_minor, y_minor))

    def axis_font(self, xtick_pt=None, ytick_pt=None, font_index=None,
                  xtitle_pt=None, ytitle_pt=None):
        """设置刻度标签与轴标题的字号/字体。

        font_index 是 Origin 的字体索引（本机 1 = Arial，也是默认值）；
        LabTalk 的 xb.font$ / SetStrProp("font", 名字) 在本环境都无效，
        只能走 X.Label.Font / fsize 这类数值属性。
        """
        if font_index is not None:
            self.app.Execute('layer.x.label.font = %d; layer.y.label.font = %d;'
                             % (font_index, font_index))
        if xtick_pt is not None:
            self.app.Execute('layer.x.label.pt = %g;' % xtick_pt)
        if ytick_pt is not None:
            self.app.Execute('layer.y.label.pt = %g;' % ytick_pt)
        if xtitle_pt is not None:
            self.app.Execute('xb.fsize = %g;' % xtitle_pt)
        if ytitle_pt is not None:
            self.app.Execute('yl.fsize = %g;' % ytitle_pt)

    def place_y_title(self, gap=26):
        """把 Y 轴标题往左挪，拉开与轴线的距离。

        实测：LabTalk 的 yl.x 是页面设备单位，把它调小 N 个单位，
        导出 1600px 宽的图上标题与轴线的空隙就多约 N 像素（1:1）。
        gap 是想要留出的空隙像素数；默认 26 与参考图
        （空隙/字高 ≈ 0.64）相当。
        """
        cur = self.app.LTVar("yl.x")
        if not cur:
            return False
        self.app.Execute("yl.x = %g;" % (cur - gap))
        return True

    def set_legend(self, text, size=None, frame=False):
        r"""直接改写图例文字（例如 \l(1) 20261023-1-a）。

        frame=False 时去掉图例边框（legend.showframe = 0）；
        frame=True 则保留边框。底色另由 legend.background 控制。
        """
        pg = self.app.ActivePage
        ly = pg.Layers(0)
        for i in range(ly.GraphObjects.Count):
            o = ly.GraphObjects(i)
            if o.Name == "Legend":
                o.Text = text
                if size:
                    self.app.Execute('legend.fsize = %g;' % size)
                self.app.Execute('legend.showframe = %d;' % (1 if frame else 0))
                return True
        return False

    def add_text(self, name, text, size=22, font="Arial", left=None, top=None):
        """在图上新建一个文字对象。

        实测要点：
        - 语法是 label -s ... -n 名字 "内容"，且 **选项必须写在 -n 之前**，
          否则选项会被当成文字内容的一部分（"X" 前多出 '-fs 22 '）。
        - 字体用 -fn Arial 传入（Origin 自己换算成字体索引），
          LabTalk 的 .font$ / .fontname$ 赋值在本环境无效。
        - 位置只能用 COM 的 Left / Top（页面设备单位 1/600 英寸），
          LabTalk 的 .x / .y 赋值返回 True 但位置不动。
        """
        if left is not None or top is not None:
            self.app.Execute('label -s -fs %g -fn %s -n %s "%s";'
                             % (size, font, name, text))
        else:
            self.app.Execute('label -s -fs %g -fn %s -n %s "%s";'
                             % (size, font, name, text))
        obj = self._find_object(name)
        if obj is None:
            return None
        if left is not None:
            obj.Left = int(left)
        if top is not None:
            obj.Top = int(top)
        return obj

    def _find_object(self, name):
        ly = self.app.ActivePage.Layers(0)
        for i in range(ly.GraphObjects.Count):
            o = ly.GraphObjects(i)
            if o.Name == name:
                return o
        return None

    def text_below_legend(self, text, size=22, font="Arial", gap=40, center=True):
        """在图例正下方加一行文字（例如 ID/IG=0.55），并与图例水平居中对齐。

        返回 (Left, Top, Width, Height)；找不到图例时返回 None。
        """
        lgn = self._find_object("Legend")
        if lgn is None:
            return None
        lL, lT, lW, lH = lgn.Left, lgn.Top, lgn.Width, lgn.Height
        obj = self.add_text("idig_text", text, size=size, font=font,
                            top=lT + lH + gap)
        if obj is None:
            return None
        if center:
            obj.Left = int(lL + (lW - obj.Width) / 2.0)
        else:
            obj.Left = int(lL)
        return obj.Left, obj.Top, obj.Width, obj.Height

    def place_idig_clear(self, text, size=22, font="Arial",
                         gaps=(170, 220, 270, 330), aligns=("right", "center", "left")):
        """在图例正下方放标注，并自动挑一个**不压曲线**的位置。

        为什么需要：图例贴在右上角是安全的（曲线峰值到不了那么高），
        但图例正下方往往正好是峰顶附近，标注会盖在曲线上。
        这里按 间距×对齐 逐格试，每试一次就渲染自检，第一个干净的就用。

        返回 (Left, Top, Width, Height)；都不干净时返回最后一个尝试的位置。
        """
        lgn = self._find_object("Legend")
        if lgn is None:
            return None
        lL, lT, lW, lH = lgn.Left, lgn.Top, lgn.Width, lgn.Height
        last = None
        for gap in gaps:
            for align in aligns:
                obj = self.add_text("idig_text", text, size=size, font=font,
                                    top=lT + lH + gap)
                if obj is None:
                    continue
                if align == "right":
                    obj.Left = int(lL + lW - obj.Width)
                elif align == "center":
                    obj.Left = int(lL + (lW - obj.Width) / 2.0)
                else:
                    obj.Left = int(lL)
                ok, _ = self.check_overlap(names=("idig_text",))
                last = (obj.Left, obj.Top, obj.Width, obj.Height)
                if ok:
                    return last
        return last

    # ---------- 页面裁剪 ----------
    def page_size(self):
        """返回 (宽, 高)，单位是页面设备单位 1/600 英寸。"""
        return self.app.LTVar("page.width"), self.app.LTVar("page.height")

    def content_bbox(self):
        """返回 (L, T, R, B)：图层 + 所有可见图形对象的外接矩形，设备单位。"""
        pw, ph = self.page_size()
        ly = self.app.ActivePage.Layers(0)
        L = self.app.LTVar("layer.left") * pw / 100.0
        T = self.app.LTVar("layer.top") * ph / 100.0
        R = L + self.app.LTVar("layer.width") * pw / 100.0
        B = T + self.app.LTVar("layer.height") * ph / 100.0
        for i in range(ly.GraphObjects.Count):
            o = ly.GraphObjects(i)
            if o.Name.startswith("_") or not o.Show or o.Width <= 0 or o.Height <= 0:
                continue
            L = min(L, o.Left)
            T = min(T, o.Top)
            R = max(R, o.Left + o.Width)
            B = max(B, o.Top + o.Height)
        return L, T, R, B

    def _probe_ink(self, tmpdir=None, probe_px=1400):
        """渲染一张探测 PNG，返回墨迹包围盒在**设备单位**下的 (L, T, R, B)。

        为什么不用 content_bbox()：刻度标签（如 X 轴最右的 "2000"）不是
        GraphObject，取不到 bbox，纯几何裁剪会把它们切掉。必须靠真实像素。
        """
        tmpdir = tmpdir or os.environ.get("TEMP", ".").replace("\\", "/")
        tmp = os.path.join(tmpdir, "_origin_trim_probe")
        png = tmp + ".png"
        if os.path.exists(png):
            os.remove(png)
        pw, ph = self.page_size()
        px_w = int(probe_px)
        px_h = max(1, int(round(probe_px * ph / pw)))
        self.app.Execute(
            'expGraph type:=png path:="%s" filename:="_origin_trim_probe" overwrite:=replace '
            'tr1.unit:=2 tr1.Width:=%d tr1.Height:=%d;' % (tmpdir, px_w, px_h))
        if not os.path.exists(png):
            return None
        try:
            from PIL import Image
            import numpy as np
        except ImportError:
            return None
        a = np.array(Image.open(png).convert("L"))
        h, w = a.shape
        mask = a < 240
        if not mask.any():
            return None
        ys, xs = np.where(mask)
        sx, sy = pw / float(w), ph / float(h)
        os.remove(png)
        return xs.min() * sx, ys.min() * sy, (xs.max() + 1) * sx, (ys.max() + 1) * sy

    def trim_page(self, margin=150, iters=3):
        """把页面裁到**实际渲染出的墨迹**外接框 + 四周 margin（设备单位 1/600 英寸）。

        关键：**先改 page 尺寸，再按物理尺寸重算 layer 的百分比**，
        否则图层会跟着页面一起缩放，图就变形了。
        每次迭代都重新渲染探测，直到墨迹离四边都是 margin。
        """
        for _ in range(iters):
            ink = self._probe_ink()
            if ink is None:
                return self.page_size()
            il, it_, ir, ib = ink
            pw, ph = self.page_size()

            # 当前图层的物理几何
            L = self.app.LTVar("layer.left") * pw / 100.0
            T = self.app.LTVar("layer.top") * ph / 100.0
            W = self.app.LTVar("layer.width") * pw / 100.0
            H = self.app.LTVar("layer.height") * ph / 100.0

            npw = (ir - il) + 2 * margin
            nph = (ib - it_) + 2 * margin
            if abs(npw - pw) < 2 and abs(nph - ph) < 2 and \
               abs(il - margin) < 2 and abs(it_ - margin) < 2:
                break

            self.app.Execute("page.width=%g; page.height=%g;" % (npw, nph))
            self.app.Execute("layer.width=%g; layer.height=%g;"
                             % (W / npw * 100.0, H / nph * 100.0))
            self.app.Execute("layer.left=%g; layer.top=%g;"
                             % ((L + (margin - il)) / npw * 100.0,
                                (T + (margin - it_)) / nph * 100.0))
        return self.page_size()

    # ---------- 出图前视觉自检 ----------
    def check_overlap(self, names=("Legend", "idig_text"), probe_px=1400,
                      pad_px=4, dark_threshold=240, min_dark=6, tmpdir=None):
        """检查图例/文字对象是否压在曲线或坐标轴上（出图前必跑）。

        做法：把这些对象临时隐藏 -> 渲染一张探测 PNG -> 数它们原占位置里
        还有多少深色像素。有 = 曲线/轴线穿过那块地方 = 互相覆盖。
        数完把 Show 恢复原样。

        返回 (ok, report)：
          ok      True 表示所有对象都干净
          report  [(name, 深色像素数, 判定), ...]，便于直接打印给用户看
        """
        try:
            from PIL import Image
            import numpy as np
        except ImportError:
            return True, [("(跳过)", 0, "缺少 Pillow/numpy，未做视觉自检")]

        pg = self.app.ActivePage
        ly = pg.Layers(0)
        objs = []
        for i in range(ly.GraphObjects.Count):
            o = ly.GraphObjects(i)
            if o.Name in names and o.Show:
                objs.append((o.Name, o.Left, o.Top, o.Width, o.Height))
        if not objs:
            return True, []

        # 临时移出页面。实测 o.Show = False 对 expGraph 无效（对象照旧渲染出来），
        # 只有把 Left 挪到页面外才能真正让它消失。
        saved = {}
        for i in range(ly.GraphObjects.Count):
            o = ly.GraphObjects(i)
            if o.Name in names:
                saved[o.Name] = (o.Left, o.Top)
                o.Left = -60000
        try:
            tmpdir = tmpdir or os.environ.get("TEMP", ".").replace("\\", "/")
            tmp = os.path.join(tmpdir, "_origin_overlap_probe")
            png = tmp + ".png"
            if os.path.exists(png):
                os.remove(png)
            pw, ph = self.page_size()
            px_w = int(probe_px)
            px_h = max(1, int(round(probe_px * ph / pw)))
            self.app.Execute(
                'expGraph type:=png path:="%s" filename:="_origin_overlap_probe" '
                'overwrite:=replace tr1.unit:=2 tr1.Width:=%d tr1.Height:=%d;'
                % (tmpdir, px_w, px_h))
            if not os.path.exists(png):
                return True, [("(跳过)", 0, "探测图未生成，未做视觉自检")]
            arr = np.array(Image.open(png).convert("L"))
            os.remove(png)
        finally:
            for i in range(ly.GraphObjects.Count):
                o = ly.GraphObjects(i)
                if o.Name in saved:
                    o.Left, o.Top = saved[o.Name]

        h, w = arr.shape
        sx, sy = w / float(pw), h / float(ph)
        report = []
        ok = True
        for name, L, T, W, H in objs:
            x0 = max(0, int((L - pad_px / sx) * sx))
            x1 = min(w, int((L + W + pad_px / sx) * sx))
            y0 = max(0, int((T - pad_px / sy) * sy))
            y1 = min(h, int((T + H + pad_px / sy) * sy))
            if x1 <= x0 or y1 <= y0:
                report.append((name, 0, "越界，跳过"))
                continue
            sub = arr[y0:y1, x0:x1]
            n = int((sub < dark_threshold).sum())
            bad = n >= min_dark
            ok = ok and not bad
            report.append((name, n, "压在曲线上！" if bad else "干净"))
        return ok, report

    # ---------- 导出 ----------
    def export(self, outdir, name="figure", width=1600, height=1200, types=("png",)):
        """导出当前激活的图。outdir 为正斜杠绝对目录，不要含空格。

        width/height 是像素尺寸。注意这会强制拉伸成该尺寸，
        页面长宽比和它不一致时图会被拉变形 —— 要精确控制请用 export_dpi()。
        """
        out = []
        base = name
        for ext in types:
            f = os.path.join(outdir, base + "." + ext)
            if os.path.exists(f):
                os.remove(f)
            self.app.Execute(
                'expGraph type:=%s path:="%s" filename:="%s" overwrite:=replace '
                'tr1.unit:=2 tr1.Width:=%d tr1.Height:=%d;'
                % (ext, outdir, base, width, height)
            )
            out.append((f, os.path.exists(f), os.path.getsize(f) if os.path.exists(f) else 0))
        return out

    def export_dpi(self, outdir, name="figure", dpi=300, types=("tif",)):
        """按页面实际尺寸 + 指定 dpi 导出，保持长宽比不变形。

        实测：expGraph 的 tr1.DPI 参数会让命令整体失败（返回 False 且不生成文件），
        只能用 tr1.unit:=1（单位 cm）或 unit:=2（像素）间接控制。
        tif 默认就是 300 dpi，这里按 dpi 反算像素，再验证写回的分辨率。
        """
        pw, ph = self.page_size()
        cm_w, cm_h = pw / 600.0 * 2.54, ph / 600.0 * 2.54
        px_w, px_h = int(round(cm_w / 2.54 * dpi)), int(round(cm_h / 2.54 * dpi))
        out = []
        base = name
        for ext in types:
            f = os.path.join(outdir, base + "." + ext)
            if os.path.exists(f):
                os.remove(f)
            self.app.Execute(
                'expGraph type:=%s path:="%s" filename:="%s" overwrite:=replace '
                'tr1.unit:=2 tr1.Width:=%d tr1.Height:=%d;'
                % (ext, outdir, base, px_w, px_h)
            )
            out.append((f, os.path.exists(f), os.path.getsize(f) if os.path.exists(f) else 0))
        return out, (cm_w, cm_h), (px_w, px_h)

    # ---------- 保存工程 ----------
    def save_project(self, path, overwrite=False):
        """另存为 .opju。path 必须是**反斜杠**的绝对路径。

        实测：app.Save() 不传参数返回 False 什么也不做；
        传正斜杠路径会被当成相对文件名拼到 Documents 下并弹「文件正在使用」错误框；
        传反斜杠绝对路径才成功。

        安全：目标已存在时默认拒绝（overwrite=False），避免把用户自己的工程覆盖掉。
        2026-11-18 就是因为无脑 Save 到已存在的 raman_1a.opju，覆盖了用户 1.5 MB 的拟合工程。
        """
        p = path.replace("/", "\\")
        if os.path.exists(p) and not overwrite:
            raise FileExistsError(
                "目标工程已存在，拒绝覆盖: %s\n"
                "（确需覆盖请显式传 overwrite=True，或换个输出名）" % p
            )
        return self.app.Save(p), os.path.exists(p)