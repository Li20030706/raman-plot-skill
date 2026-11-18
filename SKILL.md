---
name: raman-plot
description: "用 COM 控制本机 Origin 绘制碳材料拉曼单谱（D 带 / G 带，1000–2000 cm⁻¹，含 I_D/I_G 标注）。当用户要求把拉曼 txt 数据画成图、批量出图、或改 ID/IG 标注时使用。"
---

# 碳材料拉曼单谱成图（Origin COM 路线）

## ⚠️ 适用范围与局限（先读这一节）

**这个 skill 只适用于碳材料（石墨、石墨烯、碳纳米管、无定形碳等）的 D 带 / G 带拉曼谱。**

| 限制 | 说明 |
|---|---|
| **X 轴固定 1000–2000 cm⁻¹** | 只覆盖 D 带（~1350）和 G 带（~1580）。其他拉曼振动峰——如 Si 520、TiO₂ 144/399/639、MoS₂ 383/408、有机物指纹区 400–1000——**全部落在绘图区间之外，画不出来**。 |
| **只认两个峰** | 拟合固定为「常量基线 + 两个 Gaussian」，峰位约束 D 1280–1450、G 1500–1650。谱上多于两个峰、或峰不在这两个窗口内，结果无意义。 |
| **只报 I_D/I_G** | 不做峰位/半高宽/晶粒尺寸的正式报告，也不做去卷积、分峰归属。 |
| **不是通用拉曼工具** | 换材料（无机氧化物、聚合物、矿物）请另写流程，不要硬套本脚本。 |

真要用于别的材料，需要改的地方：`raman_plot.py` 的 `x_lo/x_hi`（命令行第 3、4 个参数可临时覆盖）、
`origin_plot.box_axes` 的 `xinc`，以及 `raman_fit.py` 里 `recipe_A` 的峰位窗口和峰数。

---

用户已经定稿样式，**不要重新调样式**，直接跑脚本。

## 一条命令搞定

```bash
cd <本 skill 目录>
python raman_plot.py <数据.txt> [输出名主干] [x下限] [x上限]
```

默认 X 轴 1000–2000，输出到脚本同目录：

- **`<名>.tif`** —— 300 dpi，页边距紧凑（四周 0.25 英寸）—— 交付图
- **`<名>.opju`** —— Origin 工程源文件 —— 交付源文件

## 交付物规范

| 交付物 | 规格 |
|---|---|
| `.opju` | Origin 工程，含数据表 + 成图 |
| `.tif` | 300 dpi，页边距紧凑（四周 150 设备单位 = 0.25 in = 75 px） |

### 命名规则：跟 txt 完全一致

**txt 文件叫什么，`.tif` 和 `.opju` 就叫什么**，只换扩展名，不加 `_v2`/`_check` 之类后缀。

- `20261023-1-a.txt` → `20261023-1-a.tif` + `20261023-1-a.opju`
- 脚本里 `out_name` **默认就是 txt 的主干名**，直接 `python raman_plot.py <数据.txt>` 即可，不用传第二个参数。
- 需要试参数时另建临时名，验证完删掉，别污染交付目录。

### 保存 .opju —— 只能用反斜杠绝对路径

```python
op.save_project(r"<输出目录>\figure.opju")   # 内部把 / 转成 \
```

- `app.Save()` 无参数 → 返回 False，什么都不做
- `app.Save("D:/.../x.opju")` 正斜杠 → 被拼到 `Documents\OriginLab\User Files\` 下，弹「文件正在使用」错误框，
  **且会卡死后续所有 COM 调用**（症状：整个脚本超时）。遇到就 `EnumWindows` 找 `#32770` 发 `WM_CLOSE`。
- LabTalk `save -i` / `doc -s` / `saveas` 全部无效
- 失败会留下 `*.opju.oggu` 垃圾文件

### 紧凑页边距 —— 必须靠渲染探测，不能用几何 bbox

```python
op.trim_page(margin=150)   # 设备单位 1/600 in
```

- **几何 bbox 会切掉刻度标签**：`GraphObject` 里没有 X 轴 "2000"、Y 轴数字这些对象，纯几何裁剪会把它们切一半。
- 正确做法：`_probe_ink()` 渲染 1400px 探测 PNG → 灰度阈值 240 取墨迹框 → 换算回设备单位。
- **先改 `page.width/height`，再按物理尺寸重算 `layer` 百分比**，反过来图层会跟着页面缩放导致变形。
- 迭代 3 次收敛。

### TIF 的 300 dpi

- `expGraph type:=tif` **默认就是 300 dpi**，不需要也不能显式设置。
- `tr1.DPI:=300` 会让整条命令失败（返回 False 且不生成文件）。
- 像素尺寸按 `页面英寸 × dpi` 反算，用 `tr1.unit:=2 tr1.Width:=W tr1.Height:=H` 给出。
- `tr1.unit:=0` 也失败，`unit:=1`（cm）可用但会被四舍五入。

## 环境前提

- Origin 必须**已经打开**（OriginPro 2021，可见窗口）。
- 用 `win32com.client.Dispatch("Origin.ApplicationSI")` 连正在运行的实例，不会新开进程。
- `GetActiveObject` 会报 `com_error -2147221021`，**这是正常的**，不要改用它。
- 依赖：pywin32、numpy、scipy、Pillow。

## 定稿样式（逐条确认过，别改）

| 项 | 值 |
|---|---|
| X 轴 | 1000–2000，主刻度 500，次刻度 250 |
| Y 轴 | 无刻度、无数值标签，只留轴线 |
| 上轴 + 右轴 | 有轴线、无刻度，形成封闭方框 |
| 线色 | `#CEF297` → LabTalk `206,242,151` |
| 线宽 | `set %C -w 1000` = 2.0 pt |
| 轴标题 | `Raman Shift (cm\+(-1))` / `Intensity (a.u.)`，Arial 22 |
| 图例 | 右上，内容 = 文件名主干，Arial 22，pad=140 |
| ID/IG 标注 | 图例下方 **gap=170**，Arial 22，`r"I\-(D)/I\-(G)=X.XX"` |

## 拟合口径（只用来读峰高）

- 全谱 + 常量基线（自由参数）+ 两个 Gaussian —— 与 Origin Peak Analysis 一致。
- 峰位约束：D 1280–1450，G 1500–1650；FWHM D 40–400，G 30–200。
- `ID/IG = hD / hG`（**从基线算起的峰高**，不是面积比）。
- **图上不画任何拟合曲线**，只画原始谱。拟合只是为了读出峰高。

## 报数规则

两位有效数字，四舍六入五成双（banker's rounding），用 `raman_fit.sig2()`。
例：0.6790 → 0.68，0.5486 → 0.55，0.125 → 0.12。

口径必须在报告里写明，因为换窗口/基线/线型 ID/IG 最多能差 20%。

## 关键 LabTalk 坑（踩过，别重试）

- `set %C` 是**会前进且不复位**的游标。颜色/线宽必须一次性全交给 `line_plot`，否则会从第 N 条开始把前面的样式冲掉。
- 文字对象：`label -s -fn Arial -fs 22 -n 名字 "内容";` —— **所有选项必须写在 `-n` 之前**，写在后面会被整串当成文字内容。
- 字体不能用 LabTalk 赋值（`b1.font$="Arial"` 返回 True 但读回永远 2.0），只能用 `-fn`。
- 位置只能用 COM 的 `obj.Left` / `obj.Top`（单位 1/600 英寸）；LabTalk 的 `b1.x=300` 无效。
- `page.maximize=1` 不是最大化窗口，而是**新建一张空图**。要最大化得用 win32gui 对 MDI 子窗口发 `SW_MAXIMIZE`。
- `expGraph` 前必须 `win -a "图名"` 激活目标页，否则静默返回 False 且不生成文件。
- `expWks` / `impasc` 在本环境失效，数据一律走 `PutWorksheet`。
- 删除页面用 `win -c "页面名"`。
- 残留 `#32770` 模态框会让 `app.Pages(i)` 抛 `com_error -2147418113`，用 `win32gui.EnumWindows` + `PostMessage(WM_CLOSE)` 清掉。

## 数据安全

- 若当前工程里有用户手工建的页面（例如 `Book1`），**绝对不能删**。脚本里一律 `if p.Name != KEEP_PAGE`。
  其他页面是本工具生成的，可以清。
- **绝不覆盖已存在的 .opju。** `save_project(path, overwrite=False)` 目标存在即抛 `FileExistsError`；
  `raman_plot.py` 遇到同名 `.opju` 跳过保存，只有显式 `--force` 才覆盖。
  曾因为无脑 Save 到已存在的工程，把一份 1.5 MB 的 Origin 工程覆盖成 35 KB（Origin 无自动备份、无卷影副本、回收站也没有，**不可恢复**）。

## 出图前必须做视觉验证

**最终出图之前，一定要确认图例和标注没有压在曲线上。**

### 自动自检（脚本已内置）

```python
ok, report = op.check_overlap(names=("Legend", "idig_text"))
# report -> [("Legend", 深色像素数, "干净"/"压在曲线上！"), ...]
```

做法：把图例/标注**临时挪到页面外**（`o.Left = -60000`）→ 渲染探测 PNG →
数它们原占位置里还有多少深色像素。有深色 = 曲线或轴线穿过那块地方。

- **`o.Show = False` 对 `expGraph` 无效**，对象照旧渲染出来 —— 必须挪坐标。
- `raman_plot.py` 里自检不通过会 **直接 `return 2` 停下，不出图**。
- 标注位置用 `op.place_idig_clear(label)` 自动找：按 间距×对齐 逐格试，第一个干净的就用。
  图例贴右上角一般安全，但**图例正下方往往正好是峰顶**，标注会盖上去。

### 人工视觉验证（必须做）

自动自检只数深色像素，判断不了"挤不挤"。出图后**必须把图打开看一眼**：

```python
from PIL import Image
Image.open("<名>.tif").convert("RGB").save("_check.png")   # 再目视
```

重点看：图例文字有没有被右边框切掉、标注和峰顶的距离够不够、X 轴最右的刻度标签有没有被裁。

### 顺序坑：trim 之后要重摆一次

`trim_page()` 改了 `page.width/height`，图例和文字的**绝对坐标会跟着错位**。
所以流程必须是：`place_legend` → `text_below_legend` → `trim_page` → **`place_legend` 再摆一次** → `place_idig_clear` → `check_overlap` → 导出。

## 出图后必须问用户是否满意

1. 出完图**先问用户满不满意**，不要自己认定完成。
2. 用户说满意 → 交付 `.tif` + `.opju`。
3. 用户不满意 / 要微调 → **把 `.opju` 工程文件在 Origin 里打开给用户自己拖**：

```python
op.app.Execute('doc -o "%s";' % r"<输出目录>\figure.opju")   # 打开工程
op.app.Execute('win -a "Graph1";')                            # 激活图窗口
```

用户在 Origin 窗口里手动微调（拖图例、改字号、挪标注），**不要抢着改脚本**。
用户调完再让他说一声，需要时再把调整值回写进脚本。

坑：`doc -o` 会弹「是否保存当前工程」的模态框（class `#32770`，标题就是 "OriginPro 2021"），
**弹框时整个 Origin 主窗口会变成不可见（vis=0）**，COM 调用也会挂住。
要先起一个后台线程轮询 `EnumWindows` 找 `#32770` 并 `PostMessage(WM_CLOSE)`，再发 `doc -o`。

## 文件

- `raman_plot.py` —— 通用成图（推荐入口）
- `origin_plot.py` —— COM 封装类 `OriginPlot`
- `raman_fit.py` —— 拟合与报数，`recipe_A` / `recipe_B` / `sig2`

## 批量出图

对每个文件调一次 `raman_plot.py`。注意每次都会清掉除 `KEEP_PAGE` 外的所有 Origin 页面，所以**后一锅会覆盖前一锅的 Origin 页面**（导出文件不受影响）。需要多张图同时留在 Origin 里时，要改脚本里那段清理逻辑。

## 示例结果（口径 A，碳材料）

| 文件 | ID/IG |
|---|---|
| 20261023-1-a | 0.68 |
| 20261023-1-b | 0.66 |
| 20261029-1-a | 0.93 |
| 20261029-1-b | 0.90 |
| 20261029-2-a | 0.42 |
| 20261029-2-b | 0.55 |
