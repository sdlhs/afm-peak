# afm-peak

JPK 原子力显微镜（AFM）力谱曲线的**判峰**工具，两个单文件 Python 脚本。

Peak picking for JPK AFM force–distance curves — one GUI for manual per-curve marking,
one batch screener that reads tens of thousands of raw `.jpk-force` files directly.

| 脚本 | 作用 | 平台 |
|---|---|---|
| [`txtpeak_gui.py`](txtpeak_gui.py) | 单条精修：拖框自动找峰 → 右键增删 → F13 回原始数据取精确值并写 Excel | Windows（用 Excel COM 写表） |
| [`txt-multipeak.py`](txt-multipeak.py) | 批量筛选：直读上万条 `.jpk-force`，按「≥N 个峰」筛出多峰曲线并复制到子文件夹 | 任意平台，只要 numpy |

两者**判峰口径一致**（同一套 `process()` / `detect()`），只是输入不同：
前者读 JPK Data Processing 导出的 9 列 txt，后者直接读 `.jpk-force` 压缩包，不用先手工导出。

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

- `txtpeak_gui.py`：Python 3.8+，**Windows**，需要本机装 **Microsoft Excel**（用 `win32com` 写 xlsx）；
  依赖 `numpy`、`keyboard`、`pywin32`。
- `txt-multipeak.py`：Python 3.8+，任意平台，只依赖 `numpy`（tkinter 用于弹窗选目录，缺失时可用命令行传路径）。

## 用法

### 单条精修（`txtpeak_gui.py`）

```bash
python txtpeak_gui.py                 # 弹窗选文件夹
python txtpeak_gui.py "D:\data\exp1"  # 或直接给路径
```

会依次读文件夹里的 `*.txt`，每个文件：

| 操作 | 效果 |
|---|---|
| 左键拖框 | 在框内自动找峰并标红圈 |
| 右键点峰 | 删掉这个峰 |
| 双击右键 | 加一个峰（自动吸到附近最高点） |
| 中键拖动 / 滚轮 | 平移 / 缩放 |
| `F13` | 定稿：回原始数据在 ±2 nm 内取最大值 → 写 Excel → 下一个 |
| 左 `Ctrl` | 跳过这个文件 |
| `Backspace` | 回到上一个文件（已改过的峰保留，再按 F13 覆盖原来那一行） |
| 小键盘 `+` / `-` | 阈值跳到「峰数会变」的那一档 |
| `Esc` | 退出 |

输出方式比较特别：**不生成文件，而是直接写进已经打开的那张 Excel 表**。

- 先用 Excel 打开你的工作簿，把光标停在要开始写的单元格上（程序以它为起点，向右写）；
- 每按一次 `F13`：该行第一格写序号（= 文件名在文件夹里的次序），右边依次写各峰高度；
- 峰值取整到 1 pN 再写，写完回读校验，不一致会提示；光标自动移到下一行同一列；
- 如果选中的格落在已经写过的行上，会自动往下让到第一个空行，不覆盖别人；
- 退出时把文件夹记在脚本同目录的 `txtpeak_last.txt`，下次启动自动带上。

### 批量筛选（`txt-multipeak.py`）

在 Spyder 里直接 `F5` 就会弹文件夹选择框；也可以命令行：

```bash
python txt-multipeak.py                              # 弹窗选总目录
python txt-multipeak.py "D:\data\2026-11-R4"         # 指定总目录
python txt-multipeak.py "D:\data\2026-11-R4" --dry-run   # 只统计，不复制
python txt-multipeak.py --test "D:\data\exp1"        # 自检：和 JPK 导出的 txt 比峰
```

给一个**总目录**，程序会把**它下面第一层的每个子目录**当成一批，各自在自己的目录里建
`multipeak\` 子文件夹，把多峰文件复制进去。跳过一个子目录的三种情况：

1. 里面的 `.jpk-force` 少于 `--min-files`（默认 10）条；
2. 目录名带 `_processed`（JPK 处理输出，不是原始数据）；
3. 目录名就是输出目录本身（`multipeak` / `多峰`）。

参数：

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

- 输入：JPK 的 `.jpk-force`（zip 容器，通道数据为大端 int16，倍率从
  `shared-data/header.properties` 的 `lcd-info.N` 换算链乘出来，不写死）。
- 判峰只用**回撤段**（retract）；去程段只用于显示和基线。
- 也兼容 JPK Data Processing 导出的 9 列 txt（`txtpeak_gui.py` 仅支持这种）。

## 已知限制

- `.jpk-force` 里**不含**弹簧常数（springConstant），它存在 JPK 软件的标定里；
  所以批量输出里的 k 列可能为空，峰位/峰数不受影响。
- 点数不足 200 的段会被跳过（多见于采集中断），记在 `_跳过.txt`。
- 峰顶位置精度为 ±dx/2；峰高取采样点上的最大值。
- 判据是「显著的局部极大值」，对**台阶型事件**（斜率突变但力不回升）不敏感。

## 许可

MIT，见 [LICENSE](LICENSE)。
