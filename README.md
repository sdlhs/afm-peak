# afm-peak

JPK 原子力显微镜（AFM）力谱（force–distance）曲线的**判峰**工具，两个单文件 Python 脚本。

Peak picking for JPK AFM force–distance curves — one GUI for manual per-curve marking with
WLC fitting, one batch screener that reads tens of thousands of raw `.jpk-force` files directly.

| 脚本 | 作用 | 平台 |
|---|---|---|
| [`txt-wlc-peak.py`](txt-wlc-peak.py) | 单条精修：左键拖框自动找峰 → 增删峰 → `F13` 回原始数据取精确值并写 Excel；同时叠一簇 WLC 曲线拟合 p / ΔL | Windows（用 Excel COM 写表） |
| [`txt-multipeak.py`](txt-multipeak.py) | 批量筛选：直读上万条 `.jpk-force`，按「≥N 个峰」筛出多峰曲线并复制到子文件夹 | 任意平台，只要 numpy |

两者**判峰口径完全一致**（同一套 `process()` / `detect()`），只是输入不同：
前者读 JPK Data Processing 导出的 txt，后者直接读 `.jpk-force` 压缩包，不用先手工导出。

## 判峰方法

核心是「在力曲线本身找显著极大值」，不是找二阶导数的折点：

```
σ   = 1.4826 × MAD(Δy) / √2        每条曲线自己估噪声（稳健标准差）
thr = 12 × σ                        阈值 = 12σ，相对量，不用手工调
峰  = find_peaks(y, prominence ≥ thr, 最小间距 8 nm)
```

三处设计让同一套参数能跨数据使用：

1. **窗口按 nm 定义，按点距自适应**。粗平滑 3 nm、精平滑 1 nm、峰顶搜索 ±2 nm，
   程序自己用 `dx = median(Δx)` 换算成点数。所以 256 点/段和 4000 点/段用同一套参数。
2. **阈值相对化**。`12σ` 随每条曲线的噪声水平浮动，低力数据（设定值 0.6 pN）不会因为
   绝对阈值过高而漏判，大信号数据也不会因为阈值过低而把噪声当峰。
3. **先剔脱开峰**。`drop_rupture()` 把回撤末端那次完整的拉脱大峰去掉，再数剩下的峰，
   这样「一条曲线上有几个分子被依次拉开」才数得准。

## 环境

```bash
pip install -r requirements.txt
```

- `txt-wlc-peak.py`：Python 3.8+，**Windows**，需要本机装 **Microsoft Excel**（用 `win32com`
  往打开的表里写）。依赖 `numpy`、`keyboard`、`pywin32`。Tk 界面用标准库 `tkinter`。
- `txt-multipeak.py`：Python 3.8+，任意平台，只依赖 `numpy`（tkinter 用于弹窗选目录，缺失时可用命令行传路径）。

## 用法（单条精修 `txt-wlc-peak.py`）

在 **Spyder 里直接 F5**，或命令行：

```bash
python txt-wlc-peak.py                 # 弹窗选文件夹
python txt-wlc-peak.py "D:\data\exp1"  # 或直接给路径
```

依次读文件夹里的 `*.txt`，每个文件：

| 操作 | 效果 |
|---|---|
| **左键拖框** | 只在这个框里找峰（画面上留一个绿虚线框；再拖一个就换区域） |
| **左键单击** | 添加一个峰（自动吸到附近最高点；附近已有峰则不重复加） |
| **右键单击** | 去掉鼠标圈住的那个峰（立即生效） |
| 中键单击 | 在曲线上放绿圈锚点，用它反解 WLC 的 L₀（= Igor 的 cursor B） |
| 中键拖动 | 平移（按住滚轮）；滚轮 = 缩放；`Home` = 复位视图 |
| `←` `→` | 锚点沿曲线逐点移动（`Shift`+方向键 = 10 点） |
| `↑` `↓` | 深蓝曲线上下平移 1 pN（旋钮区滚轮 / 左右拖同效） |
| `p` / `ΔL` / `N` 滑块 | WLC 曲线族参数，拖动实时重画 |
| `F13` | 定稿：回原始数据在 ±2 nm 内取最大值 → 写 Excel → 下一个文件 |
| 左 `Ctrl` | 跳过这个文件（不写行） |
| `Backspace` | 回到上一个文件（已改过的峰保留，再按 `F13` 覆盖它原来那一行） |
| 小键盘 `+` / `-` | 阈值直接跳到「峰数会变」的那一档（`+` 少一个峰、`-` 多一个峰） |
| `Esc` | 退出 |

打开一个文件时**先不标峰**（画面上一个红圈都没有），先看曲线，再用左键框你要找峰的那一段。

### 输出方式：直接写进已经打开的 Excel

不生成文件，而是往你**正开着的那张表**的活动单元格写。

- 先打开工作簿，把光标停在开始写的单元格（程序以它为起点向右写）；
- 每按一次 `F13`，一行一条曲线：

  | A | B | C | D | E | F | G | … |
  |---|---|---|---|---|---|---|---|
  | 序号 | p | ΔL | *(空)* | 峰1 | 峰2 | 峰3 | … |

  序号 = 该文件在文件夹里的次序；p / ΔL 是当前 WLC 参数（保留 2 位小数）；
  峰高四舍五入到 1 pN；写完回读校验，光标自动移到下一行同一列。
- 选中的格落在已经写过的行上时，会自动跳到最后一个已写行的下一行，不覆盖别人；
- 退出时把文件夹记在脚本同目录的 `txtpeak_last.txt`，下次启动自动带上。

### WLC 拟合

叠一簇蠕虫状链（WLC）曲线，用来读链的持续长度 p 和断裂间距 ΔL：

```
f(x) = (kBT/p) · [ 0.25·(1 − x/Lc)⁻² − 0.25 + x/Lc ]      kBT = 4.1 pN·nm
```

- 中键在曲线上点一下 → 用这个点反解 L₀；
- 画 `N` 条曲线，第 i 条的 `Lc = L₀ + i·ΔL`（ΔL 为负 → 一族逐渐变短的链）；
- 滑块范围：p `0…1.2`（实际取 `max(0.01, …)`，避开除零）、ΔL `−36…−20`（nm，步长 0.01）、N `1…20`；
- 每写成功一行，除了 Excel 里那两格，还会在 **Excel 工作簿所在目录**（拿不到就用曲线目录）
  追加一句 `wlc fit results.txt`：`curve \t p \t DeltaL`，同一条曲线退格重做只改原行；
- 参数存在脚本同目录的 `wlc_cfg.txt`，下次启动自动恢复（两个 txt 都不进版本库）。

## 用法（批量筛选 `txt-multipeak.py`）

给一个**总目录**，程序把**它下面第一层的每个子目录**当成一批，各自在自己的目录里建
`multipeak\`，把多峰文件复制进去：

```bash
python txt-multipeak.py                                 # 弹窗选总目录
python txt-multipeak.py "D:\data\2026-11-R4"            # 指定总目录
python txt-multipeak.py "D:\data\2026-11-R4" --dry-run  # 只统计，不复制
python txt-multipeak.py --test "D:\data\exp1"           # 自检：和 JPK 导出的 txt 比峰
```

跳过一个子目录的三种情况：里面的 `.jpk-force` 少于 `--min-files`（默认 10）条；
目录名带 `_processed`（JPK 处理输出，不是原始数据）；目录名就是输出目录本身。

| 参数 | 含义 |
|---|---|
| `--min-peaks N` | 回撤段保留几个峰算多峰（默认 3，2 峰不算） |
| `--min-files N` | 少于 N 条 `.jpk-force` 的子目录整个跳过（默认 10） |
| `--out NAME` | 输出子文件夹名（默认 `multipeak`） |
| `--jobs N` | 并行进程数（默认多核；Spyder/IPython 里自动单进程） |
| `--dry-run` | 只统计，不复制文件 |
| `--test DIR` | 自检：同一批曲线分别用原生解码和 JPK 导出的 txt 判峰，比对峰数与峰高 |

每个目录输出三个文件：`多峰清单.csv`（多峰文件，按峰数降序）、
`_全部统计.csv`（所有文件的峰数 / σ / 阈值，方便自己重新定阈值）、
`_跳过.txt`（读不了或点数太少的文件及原因）。

## 数据格式

- `txt-wlc-peak.py`：输入是**每条曲线一个 txt**，两列（高度、力）或三列（高度、力、高度2 —— 用最后一列当横轴），
  程序按「前半段 = 去程、后半段 = 回程」切开。单位给 m / N 也行，程序按数量级自动换算成 nm / pN。
  若导出时把 pause 段也写进去（列数不一致）会读不了，用只导 extend + retract 的 txt 即可。
  表头里的 `springConstant` 用来做悬臂修正（缺省 33.5 pN/nm）。
- `txt-multipeak.py`：输入是 JPK 的 `.jpk-force`（zip 容器，通道数据为大端 int16，倍率从
  `shared-data/header.properties` 的 `lcd-info.N` 换算链乘出来，不写死）。
- 判峰只用**回撤段**（retract）；去程段只用于显示和基线。

## 已知限制

- 峰顶位置精度为 ±dx/2（dx 为点距）；峰高取采样点上的最大值。
- 判据是「显著的局部极大值」，对**台阶型事件**（斜率突变但力不回升）不敏感。
- 点数不足 200 的段会被跳过（多见于采集中断），批量模式下记在 `_跳过.txt`。
- `.jpk-force` 里**不含**弹簧常数（springConstant），它存在 JPK 软件的标定里；
  所以批量输出里的 k 列可能为空，峰位/峰数不受影响。
- `keyboard` 在部分系统上需要管理员权限才能全局抓 F13 / Ctrl；Tk 窗口有焦点时不影响。

## 许可

MIT，见 [LICENSE](LICENSE)。
