# -*- coding: utf-8 -*-    最后修改: 2026-11-24 15:10
"""
txtpeak.py — 直接读 JPK 原始 txt → 画峰图 → 人工定峰 → 回原始数据取精确值 → 写 Excel

单文件版：不 import 任何本地模块。这样
  1) Spyder 不会再打 "Reloaded modules" 的红字；
  2) 不会因为 peakcap.py 被改过而互相影响；
  3) 整个程序只创建一个 Tk 根，反复运行也不会起不了窗。

为什么比截图强：截图受像素量化限制，这里峰值直接从 8000 点的原始数组里取。

流程
  启动 → 选文件夹（也可以在命令行直接给路径）→ 依次读每个 *.txt
  每个文件：处理（平台归零 + 峰翻上来）→ 弹窗（**先不标峰，等你框选**）
      左键拖框   = 只在这个框里找峰（画面上留一个绿虚线框；再拖一个就换区域）
                  脱离峰（最后那个大断裂峰）默认已自动剔除
      左键单击   = 添加一个峰（自动吸到附近最高点；附近已有峰则不重复加）
      右键单击   = 去掉鼠标圈住的那个峰（立即生效，不再等 330 ms）
      中键单击   = 放绿圈锚点，用它解出 L₀（= Igor 的 cursor B）
      中键拖动   = 平移（滚轮按住）；滚轮 = 缩放；Home = 复位视图
      方向键 ←→  = 锚点沿曲线逐点移动（Shift+方向键 = 10 点）
      方向键 ↑↓  = 深蓝曲线上下平移 1 pN（旋钮区滚轮/左右拖同效）
      p / ΔL 滑块 = WLC 曲线族的参数，拖动实时重画
      F13        = 定稿：回原始数据在 ±SEARCH_NM 内找最大值 → 写 Excel
                    （A=序号  B=p  C=ΔL  D=空  E 起=各峰高度，一行一条曲线）→ 下一个文件
      左 Ctrl    = 跳过这个文件（不写行）
      退格       = 回到上一个文件（跳错了/想重看时用），已改过的峰会保留；
                   再按 F13 会覆盖它原来那一行，不另起一行
      小键盘 +/- = 阈值直接跳到"峰数会变"的那一档（+ 少一个峰、- 多一个峰）
      Esc        = 退出整个程序
  除报错外不在控制台打印任何东西，退出时只说一句。

处理规则完全沿用 画图.py：
  baseline = mean(retract_force[-1000:])
  x        = height - min(extend_height)
  y        = -(force - baseline)          ← 这个负号把向下的峰翻上来
"""

import glob
import os
import re
import sys
import time

import numpy as np
import tkinter as tk
from tkinter import filedialog

# ---------------------------------------------------------------- 可调参数
SEARCH_NM = 2.0         # 定稿/找峰时在粗定位点 ±多少 nm 内找最大值（纯 nm，与点数无关）
COARSE_NM = 3.0         # 找峰（定峰数）用的平滑窗半宽（nm）。放宽点能压掉噪声假峰
SMOOTH_NM = 1.0         # 精定位峰顶用的平滑窗半宽（nm）。必须 <= SEARCH_NM，否则 ±2nm 够不着
ADD_SNAP_NM = 2.0       # 右键加点时，在点击处 ±多少 nm 内吸到局部最高点
                        # （按 nm 算，不能按点数：点间距大的文件会吸到隔壁峰上）
PROM_SIGMA = 12.0       # 自动找峰阈值 = PROM_SIGMA × 噪声 σ（小键盘 +/- 可实时调）
BASE_WIN = 200          # 找平台时用的窗口长度（点）
BASE_TOL = 40.0         # 窗口 std 小于这个就认为是平台（pN）
BASE_REGION = 0.40      # 只在数组开头这 40% 里找（别跑到曲线中段去）
# 回撤段末尾零点依次试这些窗口长度（Igor 原版只用第一个 200）。
# 前一个不平（std >= BASE_TOL）就退到下一个；全都不平就用第一个并告警。
TAIL_WINS = (200, 50, 20)
YMIN = 0.0              # 峰高绝对下限（pN）。0 = 不限，低力数据也能测；
                        # 杂峰靠"突出度阈值"（12×噪声）过滤就够了，不靠这个
EDIT_TOL = 2.0          # 判断"同一个峰"的 x 容差（nm）
NUDGE_STEP = 1.0        # ↑/↓ 和旋钮每格平移多少 pN
KNOB_REF_PN = 100.0     # 旋钮视觉基准：每 KNOB_REF_PN 位移 KNOB_REF_PX 像素
KNOB_REF_PX = 130.0     # 相当于原来圆盘版的"转 45°/100pN"（只影响画出来的位移）
KNOB_TOOTH_PITCH = 24   # 相邻滚花齿的间距（px）

# ---------------------------------------------------------------- WLC 拟合
KBT = 4.1                       # pN·nm
L0_MIN, L0_MAX = 20.0, 1500.0
WLC_CFG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "wlc_cfg.txt")
WLC = {"p": 0.4, "dl": -28.0, "n": 8, "l0": 300.0}
P_MIN, P_MAX = 0.0, 1.2          # p 滑块范围
P_EPS = 0.01                     # p 的实际下限（0 会让 WLC 除零）


def wlc_force(x, p, L):
    """Igor 的 LVFitWLC：f(x) = (kBT/p)(0.25(1-x/L)^-2 - 0.25 + x/L)"""
    if p <= 1e-6 or L <= 1.0:
        return np.full(np.shape(np.asarray(x, float)), np.nan)
    u = np.clip(np.asarray(x, float) / L, -0.99, 0.9995)
    return (KBT / p) * (0.25 / (1.0 - u) ** 2 - 0.25 + u)


def solve_L(xm, y, p):
    """给一个点，反解让它落在 WLC 上的 L（Igor 的 FindLc 同款）"""
    if xm <= 0 or y <= 0:
        return None
    lo, hi = xm * 1.0001, xm * 5.0
    for _ in range(120):
        mid = 0.5 * (lo + hi)
        if wlc_force(xm, p, mid) > y:
            lo = mid
        else:
            hi = mid
        if hi - lo < 1e-6:
            break
    return 0.5 * (lo + hi)


def read_header(path):
    """从 JPK 表头读悬臂常数(k, pN/nm) 和扫描速度(v, µm/s)。"""
    meta = {"k": 33.5, "v": None}
    zs = ze = ts = None
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            for _ in range(500):
                ln = fh.readline()
                if not ln or not ln.startswith("#"):
                    break
                m = re.search(r"(?i)springConstant:\s*([0-9.eE+-]+)", ln)
                if m:
                    kv = float(m.group(1)) * 1e3
                    if 1e-4 < kv < 1e6:
                        meta["k"] = kv
                m = re.search(r"(?i)relative-z-start:\s*([0-9.eE+-]+)", ln)
                if m:
                    zs = float(m.group(1))
                m = re.search(r"(?i)relative-z-end:\s*([0-9.eE+-]+)", ln)
                if m:
                    ze = float(m.group(1))
                m = re.search(r"(?i)retraceScanTime:\s*([0-9.eE+-]+)", ln)
                if m:
                    ts = float(m.group(1))
                if ts is None:
                    m = re.search(r"(?i)retract-scan-time:\s*([0-9.eE+-]+)", ln)
                    if m:
                        ts = float(m.group(1))
    except Exception:
        pass
    if zs is not None and ze is not None and ts and ts > 0:
        meta["v"] = abs(zs - ze) / ts * 1e6          # m/s -> µm/s
    return meta


def read_k(path, default=33.5):
    """从 JPK 表头读悬臂弹性常数（N/m -> pN/nm）"""
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            for _ in range(400):
                ln = fh.readline()
                if not ln or not ln.startswith("#"):
                    break
                m = re.search(r"(?i)springConstant:\s*([0-9.eE+-]+)", ln)
                if m:
                    v = float(m.group(1)) * 1e3
                    if 1e-4 < v < 1e6:
                        return v
    except Exception:
        pass
    return default


def load_wlc_cfg():
    """读配置并钳制到合法范围——配置坏了也不会出现 p=0 / N=0 / ΔL=0。"""
    try:
        with open(WLC_CFG, "r", encoding="utf-8") as fh:
            for ln in fh:
                key, _, val = ln.strip().partition("=")
                if key in WLC:
                    try:
                        WLC[key] = type(WLC[key])(float(val))
                    except Exception:
                        pass
    except Exception:
        pass
    WLC["p"] = round(max(P_EPS, min(P_MAX, float(WLC["p"]))), 2)
    WLC["dl"] = round(max(-36.0, min(-20.0, float(WLC["dl"]))), 2)
    WLC["n"] = max(1, min(20, int(WLC["n"])))
    WLC["l0"] = max(L0_MIN, min(L0_MAX, float(WLC["l0"])))


def save_wlc_cfg():
    try:
        with open(WLC_CFG, "w", encoding="utf-8") as fh:
            fh.write("p=%.4f\ndl=%.4f\nn=%d\nl0=%.4f\n"
                     % (WLC["p"], WLC["dl"], int(WLC["n"]), WLC["l0"]))
    except Exception:
        pass


def wlc_result_dir(fallback):
    """结果 txt 放 Excel 同目录；拿不到就用曲线所在目录"""
    d = ""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
            d = str(xl.ActiveWorkbook.Path or "")
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        d = ""
    return d or fallback


def append_wlc_result(dirpath, name, pval, dl):
    """记一条 WLC 结果（curve / p / DeltaL）。

    退格重做同一条曲线时不会追加重复行，只把原来那一行改掉。
    """
    path = os.path.join(dirpath, "wlc fit results.txt")
    head = "curve\tp\tDeltaL\n"
    line = "%s\t%.2f\t%.2f\n" % (name, pval, dl)
    rows = []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            rows = fh.readlines()
    except Exception:
        rows = []
    out, hit, first = [], False, True
    for ln in rows:
        if first and ln.lstrip("\ufeff").strip() == "curve\tp\tDeltaL":
            out.append(head)                       # 表头规整回来
            first = False
            continue
        first = False
        if ln.split("\t")[0] == name:              # 同名曲线：换成新值
            out.append(line)
            hit = True
        else:
            out.append(ln if ln.endswith("\n") else ln + "\n")
    if not rows or out[:1] != [head]:
        out.insert(0, head)
    if not hit:
        out.append(line)
    with open(path, "w", encoding="utf-8") as fh:
        fh.writelines(out)
    return path


# 脱离峰（最后那个大断裂峰）：默认不计入，只保留解折叠峰
DROP_RUPTURE = True
RUPTURE_RATIO = 1.35    # 最后一个峰高出前面峰中位数的多少倍 -> 脱离峰
RUPTURE_RECOVER = 0.30  # 或者：最后一个峰之后曲线最高点不到它 30% -> 脱离峰

# 基线：照 Igor 的 FXImport 各自归零（两条曲线各用各的零点，见 process()）。
# 平台法 / 差分当零点都已取消（2026-10-30）。


# ================================================================
# 基础函数（原本在 forcepeak.py 里，为了不 import 本地文件而内联）
# ================================================================
def _nice_step(span, target=8):
    """给人看的刻度步长。候选里包含 2.5 和 4，这样 2000 pN 量程出 400 的格子
    （只用 1/2/5 会跳到 500）。"""
    if span <= 0:
        return 1.0
    raw = span / float(max(1, target))
    mag = 10.0 ** np.floor(np.log10(raw))
    for m in (1.0, 2.0, 2.5, 4.0, 5.0, 10.0):
        if m * mag >= raw:
            return float(m * mag)
    return float(10.0 * mag)


def smooth(y, w):
    if w is None or w < 3:
        return y.copy()
    w = int(w) // 2 * 2 + 1
    k = np.ones(w) / w
    return np.convolve(np.pad(y, (w // 2, w // 2), mode="edge"), k, mode="valid")


def noise_sigma(y):
    d = np.diff(y)
    mad = np.median(np.abs(d - np.median(d)))
    return 1.4826 * mad / np.sqrt(2.0)


def prominences(y):
    n = len(y)
    prom = np.zeros(n)
    for i in np.nonzero((y[1:-1] >= y[:-2]) & (y[1:-1] >= y[2:]))[0] + 1:
        p = y[i]
        j, lm = i, y[i]
        while j > 0 and y[j] <= p:
            lm = min(lm, y[j])
            j -= 1
        j, rm = i, y[i]
        while j < n - 1 and y[j] <= p:
            rm = min(rm, y[j])
            j += 1
        prom[i] = p - max(lm, rm)
    return prom


def find_peaks(x, y, prom_min, min_dist=0.0):
    pr = prominences(y)
    cand = np.nonzero((y[1:-1] >= y[:-2]) & (y[1:-1] >= y[2:])
                      & (pr[1:-1] >= prom_min))[0] + 1
    keep = []
    for i in sorted(cand, key=lambda i: -y[i]):
        if all(abs(x[i] - x[j]) >= min_dist for j in keep):
            keep.append(i)
    keep.sort()
    return keep, pr


def _col_letter(n):
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def write_group(vals, seq=None, row=None, col=None, clear_to=0, taken=None,
                pval=None, dl=None):
    """把一组峰高写进 Excel，一行一条曲线：

        A=序号   B=p   C=ΔL   D=空着   E 起 = 峰高，从左往右排。

    p / ΔL 不传就不写（B、C 留空），E 列的位置固定不变。

    序号是**文件在这个文件夹里的序号**（预览标题「19/32」的那个 19），
    由调用方传进来，跟 Excel 里已有的数字无关。

    不传 row/col：用当前选中格当这一行的第一格（序号位），
                 写完光标换到下一行的同一列，接着等高；
                 如果选中格落在已经写过的那一行上（点了一下 Excel），
                 自动往下让到第一个空行，不覆盖别人。
    传了 row/col：覆盖写回那一行（退格重做用），光标不动。
    clear_to    ：这一行上次写到过几个峰，重写后峰变少时把多出来的旧值清掉。
    pval / dl   ：WLC 那两个参数，写在第 2、3 格；第 4 格留空，峰从第 5 格起。
    taken       ：已经写过的行号集合，用来判断选中格是不是占着的。
    写完回读校验，返回 (提示语, 行号, 起始列, 序号)；出错时后三个是 None。
    """
    try:
        import pythoncom
        import win32com.client
    except Exception as e:
        return ("没装 pywin32：%s" % e, None, None, None)

    from decimal import Decimal, ROUND_HALF_UP
    vals = [int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            for v in vals]

    redo = (row is not None) and (col is not None)
    pythoncom.CoInitialize()
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        try:
            sheet = xl.ActiveSheet.Name
            have_name = True
        except Exception:
            sheet, have_name = "?", False
        if redo:
            r, c = int(row), int(col)
        else:
            cell = xl.ActiveCell
            r, c = int(cell.Row), int(cell.Column)
            if taken and r in taken:               # 点到写过的那一行：往下让开
                r = max(taken) + 1
        seq = 1 if seq is None else int(seq)
        xl.Cells(r, c).Value = seq                 # A：序号
        has_w = (pval is not None) or (dl is not None)
        if has_w:                                  # B：p   C：ΔL
            xl.Cells(r, c + 1).Value = (None if pval is None
                                        else round(float(pval), 2))
            xl.Cells(r, c + 2).Value = (None if dl is None
                                        else round(float(dl), 2))
        xl.Cells(r, c + 3).Value = None            # D：空一格
        for i, v in enumerate(vals):               # E 起：峰高
            xl.Cells(r, c + 4 + i).Value = v
        for i in range(len(vals), max(int(clear_to or 0), 0)):
            xl.Cells(r, c + 4 + i).Value = None    # 重写比原来短，清掉尾巴
        back = [xl.Cells(r, c).Value]
        if has_w:
            back.append(xl.Cells(r, c + 1).Value)
            back.append(xl.Cells(r, c + 2).Value)
        for i in range(len(vals)):
            try:
                back.append(xl.Cells(r, c + 4 + i).Value)
            except Exception:
                back.append(None)
    except Exception as e:
        return ("写 Excel 出错：%s（数据可能没写进去）" % e, None, None, None)

    if not redo:
        try:
            xl.Cells(r + 1, c).Select()            # 下一行同一列
        except Exception:
            pass
    pythoncom.CoUninitialize()

    bad = 0
    want = [seq] + ([pval, dl] if has_w else []) + vals
    for v, b in zip(want, back):
        try:
            if abs(float(v) - float(b)) > 1e-9:
                bad += 1
        except Exception:
            bad += 1
    c_end = c + 4 + len(vals) - 1          # 峰恒定从 E 列起，所以末列与 p/ΔL 无关
    where = ("%s!%s%d:%s%d" % (sheet, _col_letter(c), r,
                               _col_letter(c_end), r)
             if have_name else "第 %d 行第 %d..%d 列" % (r, c, c_end))
    check = "回读校验通过" if bad == 0 else "！回读有 %d 个对不上" % bad
    far = ("覆盖原行，光标没动" if redo
           else "光标到 %s%d" % (_col_letter(c), r + 1))
    tail = ("序号%d + p%.2f + ΔL%.2f + %d 个峰" % (seq, float(pval), float(dl),
                                                  len(vals))
            if has_w else "序号%d + %d 个峰" % (seq, len(vals)))
    return ("已写入 %s 到 %s（%s；%s）" % (tail, where, check, far), r, c, seq)


# ================================================================
# 读数据 / 处理
# ================================================================
def load_jpk(path):
    """读 JPK TEXT EXPORT：跳过 # 头，2 列(高度,力) 或 3 列(高度,力,高度2)。"""
    d = np.loadtxt(path, comments="#")
    if d.ndim == 1:
        d = d.reshape(-1, 2)
    if d.shape[1] >= 3:
        h, f = d[:, 2], d[:, 1]
    else:
        h, f = d[:, 0], d[:, 1]
    if np.nanmax(np.abs(h)) < 1e-3:
        h = h * 1e9
    if np.nanmax(np.abs(f)) < 1e-3:
        f = f * 1e12
    n = len(d) // 2
    if n < 200:
        raise ValueError("点数太少（%d）" % len(d))
    return {"hex": h[:n], "fex": f[:n], "hre": h[n:], "fre": f[n:]}


def flat_window_base(arr, win=BASE_WIN, tol=BASE_TOL, region=BASE_REGION):
    """在数组**开头** region 那一段里找第一个平坦窗口，返回 (中位数, 是否找到)。

    从前往后、以 win 点为窗口、半窗为步长扫；第一个 std < tol 的窗口就用它的中位数。
    找不到（数据在漂 / 被截断）就用最前面一个窗口的中位数兜底，并返回 False。
    """
    n = len(arr)
    if n < 20:
        return float(np.median(arr)), False
    win = int(min(win, max(20, n // 4)))
    step = max(1, win // 2)
    hi = int(n * region)                  # 只在开头这一段里找，别跑到曲线中段去
    start = 0
    while start + win <= hi:
        w = arr[start:start + win]
        if float(w.std()) < tol:
            return float(np.median(w)), True
        start += step
    return float(np.median(arr[:min(max(win, n // 10), n)])), False


def tail_zero(y, wins=TAIL_WINS, tol=BASE_TOL):
    """回撤段末尾的零点：按 wins 依次试窗口长度，用第一个够平的。

    为什么要有降级：回撤段末尾那一段如果正好压着"脱开"的陡降或某个尖峰，
    200 点的 std 会很大（实测 std 400+），均值被拉高几百 pN，
    整条曲线压到负区、峰翻不过 0。这时用更短的窗口（常常已经很平，std≈5）
    反而准。默认链是 200 → 50 → 20。

    返回 (零点, 用了多少点, 提示语)。提示语为空串表示"就是标准做法，没问题"。
    """
    n = len(y)
    if n < 20:
        return float(np.mean(y)), n, "点数太少（%d），零点用全段均值" % n
    tried = []                              # [(窗口长度, std), ...]
    for w in wins:
        seg = y[-min(int(w), n):]
        sd = float(seg.std())
        tried.append((len(seg), sd))
        if sd < tol:
            if len(tried) == 1:             # 第一档就平：标准做法，不用提示
                return float(np.mean(seg)), len(seg), ""
            return (float(np.mean(seg)), len(seg),
                    "%s，改用末 %d 点（std %.0f）"
                    % (u"、".join("末尾 %d 点不平（std %.0f）" % t
                                 for t in tried[:-1]), len(seg), sd))
    # 全都不平：退回第一档（Igor 原版做法），并告警
    seg = y[-min(int(wins[0]), n):]
    return (float(np.mean(seg)), len(seg),
            "%s，全都平不了，零点可能偏，结果仅供参考"
            % u"、".join("末 %d 点 std %.0f" % t for t in tried))


def diff_curve(raw):
    """Igor 的 BaselineCorr 做法：retract[n] - approach(同一压电位置)。

    JPF_Analysis_sdlh.ipf 里是 W_Corr[n] = W[n] - Wb[Points-n]，
    W=Force_F、Wb=Force_B，按折返点对齐后逐点相减，用来消掉漂移和共模背景。

    返回值只用 ok（差分曲线有没有一段平坦的参考区），当"这文件质量如何"的参考；
    零点**不用**它 —— 两条曲线各自归零，见 process()。
    """
    m = min(len(raw["hex"]), len(raw["hre"]))
    ha, fa = raw["hex"][:m], raw["fex"][:m]
    hb, fb = raw["hre"][:m], raw["fre"][:m]
    turn = min(ha.min(), hb.min())
    # 把 approach 插值到 retract 的"离折返点距离"坐标上（同一距离＝同一压电位置）
    fa_i = np.interp(hb - turn, (ha - turn)[::-1], fa[::-1])
    _, ok = flat_window_base(fb - fa_i)
    return ok


def process(raw, k=33.5):
    """两条曲线各自归零 + 悬臂修正。

    横轴用 Extension = (头高度 − 零点) − F/k，和 Igor 的 FXImport 一致：
        Ext_for_old = Ext_for_old*1e9 - Force_for_old/spring_constant
    """
    z0 = float(np.min(raw["hex"]))
    # 照 Igor 的 FXImport：两条曲线各自归零，互不影响
    Y_ex = -raw["fex"]
    Y_re = -raw["fre"]
    # 悬臂修正：Extension = 头高度 − F/k（WLC 版一直有这一步）
    X_ex = (raw["hex"] - z0) - Y_ex / k
    X_re = (raw["hre"] - z0) - Y_re / k
    ex_base = float(np.mean(Y_ex[:200]))
    re_base, re_win, base_note = tail_zero(Y_re)
    Y_ex = Y_ex - ex_base
    Y_re = Y_re - re_base
    # 横轴各自从 0 起（Igor 用 Ef_ini=靠近段末点、Eb_ini=回撤段首点）
    X_ex = X_ex - float(X_ex[-1])
    X_re = X_re - float(X_re[0])
    # 差分曲线只用来判断质量，不当零点
    ok = diff_curve(raw)
    return dict(D_ex=X_ex, Y_ex=Y_ex, D_re=X_re, Y_re=Y_re,
                base=re_base, base_win=re_win, base_note=base_note,
                z0=z0, base_ok=ok, mode="igor", kk=k)


def detect(pr, prom=None):
    """只在 Retract 段找峰。峰位回到原始数据里定（±SEARCH_NM）。"""
    D, Y = pr["D_re"], pr["Y_re"]
    sig = float(noise_sigma(Y))
    dx = float(np.median(np.diff(D))) or 1.0
    # 找峰（决定峰数）用粗平滑。窗按 nm 定，不能按点数 —— 点间距随扫描速率差 16 倍。
    # 这个窗不能窄：窄了噪声尖峰的突出度会顶过阈值，凭空多冒出假峰。
    wc = max(3, int(round(COARSE_NM / dx)) * 2 + 1)
    ys = smooth(Y, wc)
    thr = prom if prom else max(YMIN, PROM_SIGMA * sig)
    idx, _ = find_peaks(D, ys, thr, 8.0)
    # 峰位要回到原始数据里定，分两步：
    #   1) 粗定位点 ±2*COARSE_NM 内用细平滑（SMOOTH_NM）找峰顶 —— 抵消粗平滑的拖偏
    #   2) 峰顶 ±SEARCH_NM 内取原始数据最大值 —— 这才是最终峰高
    # 两个窗都按 nm 算，折成点数会随扫描速率变，等于没定死。
    wf = max(3, int(round(SMOOTH_NM / dx)) * 2 + 1)
    ys_f = smooth(Y, wf)
    # 稀疏数据（点间距 > SEARCH_NM）时 ±SEARCH_NM 连相邻点都够不到，
    # 只能停在平滑后的坡上 —— 必须至少覆盖相邻一个点，否则报出来的是坡上的值。
    win = max(SEARCH_NM, dx)
    peaks = []
    for i in idx:
        ii0 = np.nonzero(np.abs(D - D[i]) <= COARSE_NM * 2)[0]
        if len(ii0) == 0:
            continue
        k = ii0[int(np.argmax(ys_f[ii0]))]
        ii = np.nonzero(np.abs(D - D[k]) <= win)[0]
        # 点很稀时（dx > 2nm）±win 可能只框住 k 自己，取到的是坡上的值（不是峰）。
        # 所以把窗再往外推，直到最大值落在窗中间为止 —— 保证拿到的是真正的局部最大。
        lo, hi = int(ii[0]), int(ii[-1])
        for _ in range(8):
            j = lo + int(np.argmax(Y[lo:hi + 1]))
            if lo < j < hi:
                break
            nlo, nhi = max(0, lo - 1), min(len(D) - 1, hi + 1)
            if nlo == lo and nhi == hi:
                break
            lo, hi = nlo, nhi
        if Y[j] >= YMIN:
            peaks.append([float(D[j]), float(Y[j])])
    return drop_rupture(pr, peaks), sig, thr


def peak_ladder(pr):
    """所有局部极大值的突出度（升序）+ 噪声 σ。

    +/- 调阈值时靠它直接跳到"峰数会变"的那个值，不用一档一档试。
    """
    D, Y = pr["D_re"], pr["Y_re"]
    sig = float(noise_sigma(Y))
    dx = float(np.median(np.diff(D))) or 1.0
    # 必须跟 detect 用同一个粗平滑窗，否则算出来的突出度跟实际峰数对不上，
    # 调阈值时"跳到下一个会变峰数的值"就会跳错档。
    w = max(3, int(round(COARSE_NM / dx)) * 2 + 1)
    ys = smooth(Y, w)
    pr_all = prominences(ys)
    mask = (ys[1:-1] >= ys[:-2]) & (ys[1:-1] >= ys[2:])
    vals = pr_all[1:-1][mask]
    vals = vals[vals > 0]
    return np.sort(vals), sig


def step_threshold(ladder, cur, direction, count_now, re_detect, tries=20):
    """从当前阈值 cur 出发，往 direction 方向找第一个"峰数会变"的阈值。

    返回 (peaks, sig, thr) 或 None（已经没得更变了）。
    """
    if ladder is None or len(ladder) == 0:
        return None
    if direction > 0:                       # 升阈值：越过最小的那个峰，让它掉出去
        cands = ladder[ladder > cur + 1e-6]
    else:                                   # 降阈值：刚好把下一个被排除的峰包进来
        cands = ladder[ladder < cur - 1e-6][::-1]
    for c in cands[:tries]:
        t = float(c) + 1e-3 if direction > 0 else float(c) * 0.9999
        if t <= 0:
            continue
        got = re_detect(t)
        if got is not None and len(got[0]) != count_now:
            return got
    return None


def drop_rupture(pr, peaks):
    """剔除脱离峰。它只可能是最后一个峰，所以只判断最后一个，别的峰一概不动。

    两个判据满足任一即是脱离峰：
      ① 它比前面峰的中位数高出 RUPTURE_RATIO 倍（典型的大断裂峰）；
      ② 它之后曲线就平了（回不到峰值的 30%）。用来兜住那种"脱离力很低"的情况，
         例如有的文件分子在 691 pN 就断了，倍数只有 1.25，靠 ① 会漏。
    永远只删最后一个，所以调阈值时绝不会把中段的峰误删。
    """
    if not DROP_RUPTURE or not peaks:
        return peaks
    xl, yl = peaks[-1]
    big = False
    if len(peaks) >= 2:
        med = float(np.median([p[1] for p in peaks[:-1]]))
        big = med > 0 and yl > RUPTURE_RATIO * med
    flat = False
    try:
        D, Y = pr["D_re"], pr["Y_re"]
        after = Y[D > xl + 1.0]
        flat = len(after) > 0 and yl > 0 and float(after.max()) < RUPTURE_RECOVER * yl
    except Exception:
        pass
    return peaks[:-1] if (big or flat) else peaks


def refine(pr, peaks, window=SEARCH_NM):
    D, Y = pr["D_re"], pr["Y_re"]
    # 稀疏数据（点间距 > window）时 window 连相邻点都够不到，只会取到坡上的值。
    # 跟 detect() 一样至少覆盖相邻一个点；密集数据（dx < window）仍保持严格的 window。
    dx = float(np.median(np.diff(D))) or 1.0
    win = max(window, dx)
    out = []
    for x0, _ in peaks:
        m = np.abs(D - x0) <= win
        if not m.any():
            j = int(np.argmin(np.abs(D - x0)))
            out.append((float(D[j]), float(Y[j])))
            continue
        idxs = np.nonzero(m)[0]
        # 同 detect：稀疏数据下窗口可能只框住坡上的点，往外推到最大值落中间为止。
        lo, hi = int(idxs[0]), int(idxs[-1])
        for _ in range(8):
            j = lo + int(np.argmax(Y[lo:hi + 1]))
            if lo < j < hi:
                break
            nlo, nhi = max(0, lo - 1), min(len(D) - 1, hi + 1)
            if nlo == lo and nhi == hi:
                break
            lo, hi = nlo, nhi
        out.append((float(D[j]), float(Y[j])))
    return out


def apply_edits(cands, edits):
    """把用户手动删掉/补上的峰套用回新一次自动识别的结果。"""
    dead = edits.get("deleted") or []
    add = edits.get("added") or []
    out = []
    for x, y in cands:
        if any(abs(x - d) <= EDIT_TOL for d in dead):
            continue
        out.append((float(x), float(y)))
    for ax, ay in add:
        if all(abs(ax - x) > EDIT_TOL for x, _ in out):
            out.append((float(ax), float(ay)))
    out.sort(key=lambda t: t[0])
    return out


# ================================================================
# 预览窗
# ================================================================
PREVIEW_WIN = None


def preview(fname, idx_tot, pr, peaks, sig, thr, parent, st,
            re_detect=None, edits=None, ladder=None, kk=33.5, wlc=None,
            meta=None):
    if wlc is None:
        wlc = WLC
    global PREVIEW_WIN, PROM_SIGMA
    if PREVIEW_WIN is not None:
        try:
            PREVIEW_WIN.destroy()
        except Exception:
            pass
        PREVIEW_WIN = None

    try:
        sw, sh = parent.winfo_screenwidth(), parent.winfo_screenheight()
    except Exception:
        sw, sh = 1920, 1080
    BARS = 193                        # 底部：按钮条 + 滑块条 + L₀ 条 + 刻度盘
    W = max(900, int(min(sw * 0.92, 2400)))
    WINH = max(460, min(int(sh * 0.92), 1250))
    H = max(360, WINH - BARS)         # 画布吃掉全部剩余高度
    ml, mr, mt, mb = 100, 32, 50, 56
    FS = max(11, min(16, int(W / 150)))

    win = tk.Toplevel(parent)
    PREVIEW_WIN = win
    st["job"] = None
    win.geometry("%dx%d+20+20" % (W + 16, WINH))
    win.resizable(False, False)
    win.attributes("-topmost", True)

    bar = tk.Frame(win)
    bar.pack(side="bottom", fill="x", padx=8, pady=6)
    tk.Button(bar, text="跳过此文件 (左 Ctrl)", command=lambda: setjob("skip"),
              font=("Microsoft YaHei", FS - 1)).pack(side="right")
    tk.Button(bar, text="上一个 (退格)", command=lambda: setjob("back"),
              font=("Microsoft YaHei", FS - 1)).pack(side="right", padx=8)
    tk.Button(bar, text="定稿写 Excel (F13)", command=lambda: setjob("accept"),
              font=("Microsoft YaHei", FS - 1)).pack(side="right")
    # 零点提示：默认空（正常情况不占地方），有问题时才显示红字
    base_lab = tk.Label(win, anchor="w", font=("Microsoft YaHei", FS - 1),
                        fg="#c00000", bg="#fff4e5", padx=6, pady=3)
    _note0 = (pr or {}).get("base_note") or ""
    if _note0:
        base_lab.config(text="⚠ 零点：" + _note0)
        base_lab.pack(side="bottom", fill="x", padx=8, pady=(0, 4))

    lab = tk.Label(bar, anchor="w", font=("Consolas", FS - 1), fg="#333")
    lab.pack(side="left", fill="x", expand=True)

    wbar = tk.Frame(win)
    wbar.pack(side="bottom", fill="x", padx=8, pady=(0, 2))
    v_p = tk.DoubleVar(value=float(wlc["p"]))
    v_dl = tk.DoubleVar(value=float(wlc["dl"]))
    v_n = tk.IntVar(value=int(wlc["n"]))
    SLEN = max(160, min(420, int(W * 0.26)))     # 基准长度
    SLEN_P = int(SLEN * 0.8)                     # p 短一点
    SLEN_D = int(SLEN * 1.2)                     # ΔL 长一点

    def on_wlc(_=None):
        """滑块/文本框一变就把值搬进 wlc，然后重画。

        直接读控件自己的值（s_p.get()），不依赖变量同步——
        之前有环境里 Scale 的 command 不触发，拖了没反应。
        """
        try:
            pp = float(s_p.get())
            dd = float(s_dl.get())
        except Exception:
            try:
                pp = float(v_p.get())
                dd = float(v_dl.get())
            except Exception:
                return
        try:
            nn = int(float(str(sp_n.get()).strip() or 8))
        except Exception:
            try:
                nn = int(v_n.get())
            except Exception:
                nn = 8
        wlc["p"] = round(max(P_EPS, min(P_MAX, pp)), 2)
        wlc["dl"] = round(dd, 2)
        wlc["n"] = max(1, min(20, nn))
        # 改了 p 之后 L₀ 必须重解，否则曲线就不经过锚点圈了
        # （对应 Igor 的 G_SetLtoCursor：参数一动就把 L 对到光标上）
        try:
            if ui.get("anchor") is not None:
                _ix = int(ui["anchor"])
                _L = solve_L(float(D_re[_ix]), float(Y_re[_ix]), float(wlc["p"]))
                if _L is not None:
                    wlc["l0"] = max(L0_MIN, min(L0_MAX, _L))
        except Exception:
            pass
        try:
            lab_p.config(text="%.2f" % wlc["p"])
            lab_dl.config(text="%.2f" % wlc["dl"])
        except Exception:
            pass
        draw()

    tk.Label(wbar, text="p", width=2,
             font=("Consolas", FS - 1)).pack(side="left")
    s_p = tk.Scale(wbar, from_=P_MIN, to=P_MAX, resolution=0.01, orient="horizontal",
                   length=SLEN_P, width=12, variable=v_p, showvalue=False,
                   command=on_wlc)
    s_p.pack(side="left", padx=(0, 4))
    lab_p = tk.Label(wbar, text="%.2f" % v_p.get(), width=5,
                     font=("Consolas", FS - 1))
    lab_p.pack(side="left", padx=(0, 14))

    tk.Label(wbar, text="ΔL", width=3,
             font=("Consolas", FS - 1)).pack(side="left")
    s_dl = tk.Scale(wbar, from_=-36.0, to=-20.0, resolution=0.01, orient="horizontal",
                    length=SLEN_D, width=12, variable=v_dl, showvalue=False,
                    command=on_wlc)
    s_dl.pack(side="left", padx=(0, 4))
    lab_dl = tk.Label(wbar, text="%.2f" % v_dl.get(), width=6,
                      font=("Consolas", FS - 1))
    lab_dl.pack(side="left", padx=(0, 14))

    tk.Label(wbar, text="N", width=2,
             font=("Consolas", FS - 1)).pack(side="left")
    sp_n = tk.Spinbox(wbar, from_=1, to=20, width=3, textvariable=v_n,
                      font=("Consolas", FS - 1), command=on_wlc)
    sp_n.pack(side="left", padx=(0, 6))
    for ev in ("<Return>", "<FocusOut>", "<KeyRelease>"):
        sp_n.bind(ev, lambda e: on_wlc())

    # 不依赖 Scale 的 command：直接把鼠标事件也接上，拖一下就同步一次
    def sync_from_scale(*_a):
        try:
            v_p.set(s_p.get())
            v_dl.set(s_dl.get())
        except Exception:
            pass
        on_wlc()

    for _s in (s_p, s_dl):
        _s.bind("<B1-Motion>", sync_from_scale)
        _s.bind("<ButtonRelease-1>", sync_from_scale)

    wbar2 = tk.Frame(win)
    wbar2.pack(side="bottom", fill="x", padx=8, after=bar)
    lab_wlc = tk.Label(wbar2, text="", anchor="w", fg="#a04000",
                       font=("Consolas", FS - 1))
    lab_wlc.pack(side="left")

    def sync_widgets():
        """开窗后把控件强制对齐到 wlc 的值。

        Tk 的 Scale 偶尔不在创建时读取变量，表现成"滑块停在 0 的位置、
        标签却显示 -28"，一拖就跳回 0。显式 set 一次即可根治。
        """
        try:
            s_p.set(float(wlc["p"]))
            s_dl.set(float(wlc["dl"]))
            v_p.set(float(wlc["p"]))
            v_dl.set(float(wlc["dl"]))
            v_n.set(int(wlc["n"]))
            sp_n.delete(0, "end")
            sp_n.insert(0, str(int(wlc["n"])))
            lab_p.config(text="%.2f" % float(wlc["p"]))
            lab_dl.config(text="%.2f" % float(wlc["dl"]))
        except Exception:
            pass

    # 刻度盘要先 pack（side=bottom），否则后面的 canvas 会把它挤出窗口
    KNOB_H = 78                     # 旋钮区高度（要放得下转盘）
    KNOB_TICKS = 48                 # 转盘外圈滚花齿数
    knob = tk.Canvas(win, height=KNOB_H, bg="#e8e8e8", highlightthickness=1,
                     highlightbackground="#999")
    knob.pack(side="bottom", fill="x", padx=8, pady=(0, 2))
    kui = {"px": None, "drag": False}

    cv = tk.Canvas(win, width=W, height=H, bg="white", highlightthickness=0)
    cv.pack(padx=8, pady=(8, 0))
    # 先把布局跑完，否则下面的 draw() 里 knob 的 winfo_width() 还是小值，
    # 旋钮会被画到左边去（这个坑踩过）。
    try:
        win.update_idletasks()
    except Exception:
        pass

    ed = edits if edits is not None else {"deleted": [], "added": []}
    ed.setdefault("deleted", [])
    ed.setdefault("added", [])
    ui = {"peaks": [list(p) for p in peaks], "sig": sig, "thr": thr,
          "edits": ed, "anchor": None, "off": 0.0,     # off = 深蓝曲线整体平移量(pN)
          "auto": True,                                # True = Y 轴跟着 off 自动重算
          "sel": None}                                 # 框选区域 (xa,xb,ya,yb)，y 存原始坐标

    D_ex, Y_ex = pr["D_ex"], pr["Y_ex"]
    D_re, Y_re = pr["D_re"], pr["Y_re"]

    def auto_view(off=0.0):
        """按当前曲线（含平移量 off）重算视图范围，顶部留 20% 余量。"""
        yr = Y_re + float(off)
        ytop = float(max(Y_ex.max(), yr.max(), 0.0))
        ybot = float(min(Y_ex.min(), yr.min(), 0.0))
        return [float(min(D_ex.min(), D_re.min())),
                float(max(D_ex.max(), D_re.max())),
                ybot,
                ytop * 1.2 if ytop > 0 else ybot + 100.0]

    view = auto_view(0.0)

    def x2p(v):
        return ml + (v - view[0]) / (view[1] - view[0] or 1) * (W - ml - mr)

    # ---- 纵轴分段：0 以上按正常比例，0 以下压缩成 BOTTOM_UNITS 那么高 ----
    BOTTOM_UNITS = 150.0

    def _ylayout():
        hp = float(H - mt - mb)
        vt = max(1e-6, float(view[3]))
        vb = min(0.0, float(view[2]))
        u = hp / (vt + BOTTOM_UNITS)      # 0 以上每 pN 多少像素
        y0 = mt + hp - BOTTOM_UNITS * u   # y=0 所在的像素位置
        return hp, vt, vb, u, y0

    def y2p(v):
        hp, vt, vb, u, y0 = _ylayout()
        if v >= 0 or vb >= -1e-12:
            return y0 - v * u
        span = (0.0 - vb) or 1.0
        return y0 + (-v) / span * (BOTTOM_UNITS * u)

    def p2x(p):
        return view[0] + (p - ml) / (W - ml - mr) * (view[1] - view[0])

    def p2y(py):
        """y2p 的逆：像素 → 数据值。

        本文件的纵轴是分段的（0 以上正常、0 以下压成 BOTTOM_UNITS 那么高），
        所以不能拿比例直接算，要按同一套分段反解，否则框出来的 y 范围是错的。
        """
        hp, vt, vb, u, y0 = _ylayout()
        if py <= y0 or vb >= -1e-12:
            return (y0 - py) / (u or 1.0)
        span = (0.0 - vb) or 1.0
        return -((py - y0) / (BOTTOM_UNITS * u or 1.0)) * span

    def setjob(job):
        st["job"] = job

    def _draw_wlc():
        ncur = int(wlc["n"])
        l0 = float(wlc["l0"])
        for i in range(max(0, ncur)):
            Lw = l0 + float(wlc["dl"]) * i
            if Lw <= 1.0:
                continue
            xm = np.linspace(0.0, view[1], 600)
            yw = wlc_force(xm, float(wlc["p"]), Lw)
            ok = (np.isfinite(yw) & (xm >= view[0]) & (xm <= view[1])
                  & (yw >= view[2]) & (yw <= view[3]))
            seg = []
            for j in np.nonzero(ok)[0]:
                seg.extend((x2p(xm[j]), y2p(yw[j])))
            if len(seg) >= 4:
                cv.create_line(*seg, fill="#e06010", width=1)

    def draw():
        cv.delete("all")
        has_neg = view[2] < -1e-9
        if has_neg:                       # 0 以下那条压缩带先铺底
            cv.create_rectangle(ml, y2p(0.0), W - mr, H - mb,
                                fill="#fafafa", outline="")
        step = _nice_step(view[3], 8)
        g = 0.0
        while g <= view[3] + 1e-9:
            cv.create_line(ml, y2p(g), W - mr, y2p(g), fill="#eeeeee")
            cv.create_text(ml - 8, y2p(g), text=("%g" % g).replace("-0", "0"),
                           anchor="e", fill="#666", font=("Consolas", FS - 1))
            g += step
        if has_neg:                       # 压缩带的最低点标个真实值
            cv.create_text(ml - 8, H - mb, text="%g" % round(float(view[2])),
                           anchor="e", fill="#888",
                           font=("Consolas", FS - 1))
            cv.create_text(W - mr - 6, (y2p(0.0) + H - mb) / 2.0,
                           text="0以下已压缩", anchor="e", fill="#bbb",
                           font=("Microsoft YaHei", max(8, FS - 3)))
        stepx = _nice_step(view[1] - view[0], 9)
        for i in range(int(np.ceil(view[0] / stepx)), int(np.floor(view[1] / stepx)) + 1):
            g = i * stepx + 0.0
            cv.create_line(x2p(g), mt, x2p(g), H - mb, fill="#f2f2f2")
            cv.create_text(x2p(g), H - mb + 12, text=("%g" % g).replace("-0", "0"),
                           fill="#666", font=("Consolas", FS - 1))
        cv.create_rectangle(ml, mt, W - mr, H - mb, outline="#555")
        cv.create_text(ml + (W - ml - mr) / 2, H - 20, text="Extension [nm]",
                       fill="#333", font=("Microsoft YaHei", FS - 1))
        for xx, yy, col in ((D_ex, Y_ex, "#8fb8e8"), (D_re, Y_re, "#1f4fd0")):
            if col == "#1f4fd0":            # 回撤段（深蓝）：整体平移 ui["off"]
                yy = yy + float(ui["off"])
            m = (xx >= view[0]) & (xx <= view[1])
            pts = []
            for a, b in zip(xx[m], yy[m]):
                pts += [x2p(a), y2p(b)]
            if len(pts) >= 4:
                cv.create_line(*pts, fill=col, width=1)
        _off = float(ui["off"])
        for k, (xv, yv) in enumerate(ui["peaks"], 1):
            cx, cy = x2p(xv), y2p(yv + _off)     # 峰存原始 y，画的时候加 offset
            rr = max(7, FS - 3)
            cv.create_oval(cx - rr, cy - rr, cx + rr, cy + rr, outline="#cc0000", width=2)
            cv.create_line(cx - rr - 5, cy, cx + rr + 5, cy, fill="#cc0000")
            cv.create_text(cx, cy - rr - 9, text="#%d %.0f" % (k, yv + _off),
                           fill="#aa0000", font=("Consolas", FS - 1))
        # 框选范围（绿虚线）：峰只在这个框里判过；再拖一个框就换区域
        if ui.get("sel"):
            _xa, _xb, _ya, _yb = ui["sel"]
            cv.create_rectangle(x2p(_xa), y2p(_ya + _off),
                                x2p(_xb), y2p(_yb + _off),
                                outline="#008800", dash=(4, 4), width=1)
        vals = " ".join("%.0f" % (p[1] + _off) for p in ui["peaks"])
        l0 = float(wlc["l0"])
        # WLC 曲线族（橙色）：横轴已是 Extension，直接画 (xm, F)
        try:
            _draw_wlc()
        except Exception:
            pass

        lab_p.config(text="%.2f" % float(wlc["p"]))
        lab_dl.config(text="%.2f" % float(wlc["dl"]))

        # 锚点圈（Igor 的 cursor B）
        if ui["anchor"] is not None and 0 <= ui["anchor"] < len(D_re):
            ax, ay = float(D_re[ui["anchor"]]), float(Y_re[ui["anchor"]])
            r = 7
            cv.create_oval(x2p(ax) - r, y2p(ay) - r, x2p(ax) + r, y2p(ay) + r,
                           outline="#008000", width=2)

        if ui["anchor"] is not None and 0 <= ui["anchor"] < len(D_re):
            lab_wlc.config(text="L₀=%.0f nm（锚点 %0.f,%0.f）"
                           % (l0, D_re[ui["anchor"]], Y_re[ui["anchor"]]))
        else:
            lab_wlc.config(text="L₀=%.0f nm  中键点曲线可指定" % l0)

        _seltxt = "框选区内" if ui.get("sel") else "全图（还没框选）"
        lab.config(text="PROM_SIGMA %.1f（小键盘 +/- 跳档）  %s  峰数 %d   σ≈%.1f pN   阈值 %.0f   →  %s"
                   % (PROM_SIGMA, _seltxt, len(ui["peaks"]), ui["sig"], ui["thr"], vals))
        _m = meta or {}
        _kv = float(_m.get("k", kk))
        _vv = _m.get("v")
        _vtxt = ("v=%.2f µm/s" % _vv) if _vv else "v=?"
        win.title("预览 %d/%d  %s   k=%.2f pN/nm   %s   PROM_SIGMA=%.1f   %d 个峰"
                  % (idx_tot[0], idx_tot[1], fname, _kv, _vtxt,
                     PROM_SIGMA, len(ui["peaks"])))
        try:
            knob_draw()                 # 平移量变了，刻度盘跟着动
        except Exception:
            pass
        if _note0:                      # 零点提示固定在窗口底部，重画不丢
            if not base_lab.winfo_ismapped():
                base_lab.pack(side="bottom", fill="x", padx=8, pady=(0, 4))
            base_lab.config(text="⚠ 零点：" + _note0)

    def anchor_at(ev):
        """左键点曲线：圈吸到最近数据点，用它解出 L₀"""
        if not (ml <= ev.x <= W - mr and mt <= ev.y <= H - mb):
            return
        xw = p2x(ev.x)
        ix = int(np.argmin(np.abs(D_re - xw)))
        set_anchor(ix)

    def set_anchor(ix):
        ix = int(max(0, min(len(D_re) - 1, ix)))
        ui["anchor"] = ix
        L = solve_L(float(D_re[ix]), float(Y_re[ix]), float(wlc["p"]))
        if L is not None:
            wlc["l0"] = max(L0_MIN, min(L0_MAX, L))
        draw()

    def nudge_anchor(d):
        """锚点沿曲线逐点移动（← → 一次 1 点，Shift 一次 10 点）。"""
        if ui["anchor"] is None:
            return
        set_anchor(ui["anchor"] + d)

    def hit_r():
        return max(14, FS + 6)

    def find_peak(ev, mult=1.0):
        """鼠标落在哪个红圈里？用圆距离判断，圈内一律算命中那个点。"""
        best, bd = None, hit_r() * mult
        _off = float(ui["off"])
        for i, (xv, yv) in enumerate(ui["peaks"]):
            d = ((x2p(xv) - ev.x) ** 2 + (y2p(yv + _off) - ev.y) ** 2) ** 0.5
            if d < bd:
                best, bd = i, d
        return best

    def del_at(ev):
        """删掉鼠标圈住的那个峰（点空白处什么都不做）。"""
        if not (ml <= ev.x <= W - mr and mt <= ev.y <= H - mb):
            return
        i = find_peak(ev)
        if i is None:
            return
        x_del = ui["peaks"].pop(i)[0]
        ed["deleted"].append(float(x_del))
        ed["added"] = [t for t in ed["added"] if abs(t[0] - x_del) > EDIT_TOL]
        draw()

    def on_right(ev):
        """右键：圈住哪个峰就删哪个，点了立刻生效（不再等 330 ms 判双击）。"""
        del_at(ev)

    def add_at(ev):
        """右键：加点。附近已有峰就拒绝，避免叠在一起。"""
        if not (ml <= ev.x <= W - mr and mt <= ev.y <= H - mb):
            return
        xw = p2x(ev.x)
        j = int(np.argmin(np.abs(D_re - xw)))
        # 吸附窗口按 nm 算（不是点数）：不同文件点间距差别很大，
        # 这个文件 1.35 nm/点，写死 ±12 点就是 ±16 nm，会被拉去隔壁大峰。
        _dx = float(np.median(np.diff(D_re))) or 1.0
        span = max(1, int(round(ADD_SNAP_NM / _dx)))
        a0, b0 = max(0, j - span), min(len(D_re), j + span + 1)
        j = a0 + int(np.argmax(Y_re[a0:b0]))
        nx, ny = float(D_re[j]), float(Y_re[j])
        # 关键：算出新点位置之后再和已有峰比距离。
        # 不能只比"鼠标离峰圈多远"—— 在很高的空白处点击时鼠标离圈很远，
        # 但吸附进来的新点可能正好落在某个旧点上，那样就重叠了。
        npx = x2p(nx)
        sep = max(20.0, 2.2 * max(7, FS - 3))
        for xv, _yv in ui["peaks"]:
            if abs(x2p(xv) - npx) < sep:
                return
        ui["peaks"].append([nx, ny])
        ui["peaks"].sort(key=lambda t: t[0])
        ed["added"].append((nx, ny))
        ed["deleted"] = [d for d in ed["deleted"] if abs(d - nx) > EDIT_TOL]
        draw()

    # ---------- 框选：左键拖一个框，只框里找峰 ----------
    band = {"p0": None, "rect": None}

    def apply_sel(allp):
        """只留框里的峰。sel 里存的是**原始** y（已经减掉 off），所以直接比。"""
        sel = ui.get("sel")
        if not sel:
            return [[float(q[0]), float(q[1])] for q in allp]
        xa, xb, ya, yb = sel
        return [[float(q[0]), float(q[1])] for q in allp
                if xa <= q[0] <= xb and ya <= q[1] <= yb]

    def re_detect_box(t=None):
        if re_detect is None:
            return None
        pk, sg, th = detect_off(t)          # 按显示高度判 YMIN
        return apply_sel(pk), sg, th

    def on_press1(ev):
        if not (ml <= ev.x <= W - mr and mt <= ev.y <= H - mb):
            return
        band["p0"] = (ev.x, ev.y)

    def on_drag1(ev):
        if band["p0"] is None:
            return
        if band["rect"] is not None:
            cv.delete(band["rect"])
        x0, y0 = band["p0"]
        band["rect"] = cv.create_rectangle(x0, y0, ev.x, ev.y,
                                           outline="#008800", dash=(5, 3), width=2)

    def on_release1(ev):
        p0 = band["p0"]
        band["p0"] = None
        if band["rect"] is not None:
            cv.delete(band["rect"])
            band["rect"] = None
        if p0 is None:
            return
        x0, y0 = p0
        if abs(ev.x - x0) < 8 or abs(ev.y - y0) < 8:
            add_at(ev)                      # 拖得太短 = 单击：加一个峰
            return
        xa, xb = sorted((p2x(x0), p2x(ev.x)))
        _off = float(ui["off"])
        # 存原始坐标：之后拧旋钮改 off，框里的范围也不会跟着漂
        ya, yb = sorted((p2y(y0) - _off, p2y(ev.y) - _off))
        ui["sel"] = (xa, xb, ya, yb)
        got = re_detect_box()
        if got is not None:
            ui["peaks"] = apply_sel(got[0])
            ui["sig"], ui["thr"] = got[1], got[2]
        ui["edits"]["deleted"] = []
        ui["edits"]["added"] = []
        draw()

    def zoom(ev):
        f = 0.82 if ev.delta > 0 else 1.22
        cx = p2x(ev.x)
        view[0], view[1] = cx - (cx - view[0]) * f, cx + (view[1] - cx) * f
        ui["auto"] = False                  # 手动缩放过：之后调高度别动视野
        draw()

    _pan = [None]

    def pan_start(ev):
        _pan[0] = (ev.x, ev.y, list(view))

    def pan_move(ev):
        if not _pan[0]:
            return
        x0, y0, v0 = _pan[0]
        dx = (ev.x - x0) / (W - ml - mr) * (v0[1] - v0[0])
        dy = (ev.y - y0) / (H - mt - mb) * (v0[3] - v0[2])
        view[0], view[1] = v0[0] - dx, v0[1] - dx
        view[2], view[3] = v0[2] + dy, v0[3] + dy
        ui["auto"] = False                  # 手动平移过：之后调高度别动视野
        draw()

    def pan_release(ev):
        """中键几乎没动 = 单击：在曲线上放锚点圈（原来左键的活）。"""
        p = _pan[0]
        _pan[0] = None
        if p is None:
            return
        if abs(ev.x - p[0]) <= 3 and abs(ev.y - p[1]) <= 3:
            anchor_at(ev)

    def reset(ev=None):
        ui["auto"] = True                   # Home 复位：回到自动视图
        view[:] = auto_view(ui["off"])      # 复位也保留当前平移量
        draw()

    # 左键不再放锚点：改成"中键单击"（见下面绑定块）
    # ---------- 深蓝曲线上下平移 ----------
    def redraw_only():
        draw()

    def apply_off(off):
        """设置平移量。

        只有"自动视图"模式（ui["auto"]=True）才把 Y 轴跟着重算、留 20% 余量；
        用户手动缩放/平移过之后就保持当前视野不动 ——
        否则一调高度就把局部放大冲掉了。
        """
        ui["off"] = float(off)
        if ui.get("auto", True):
            view[:] = auto_view(ui["off"])
        knob_draw()
        redraw_only()

    def nudge(dv):
        apply_off(float(ui["off"]) + float(dv))

    def detect_off(t=None):
        """按当前平移量重新找峰。

        `detect()` 里有 `if Y[j] >= YMIN` 的下限判断，用的是**原始** y。
        曲线被抬高 off 之后原始 y 普遍低 off，于是"显示上 >0 但原始 <0"
        的峰会被整片砍掉（踩过：抬高 700 后，显示值 <700 的峰全没了）。
        这里给 detect 一份 Y_re+off 的副本，再把峰 y 减回原始坐标，
        这样既按显示高度过 YMIN，又保持 ui["peaks"] 存原始 y 的约定。
        """
        off = float(ui["off"])
        q = dict(pr)
        q["Y_re"] = Y_re + off
        pk, sg, th = detect(q, prom=t)
        return [[float(x), float(y) - off] for x, y in pk], sg, th

    def knob_draw():
        """画细准焦螺旋的**侧视剖图**（不是正面圆盘）。

        视角：旋钮轴竖直立在屏幕平面里，我们从侧面看它——
        所以滚花变成一排**竖向齿条**，转动时齿条整体横向滚过去。
        这和"以 z 轴转 90°、再以 y 轴转 90°"得到的那张切面一致。

        组成（自下而上）：
          鼓身      —— 灰色长方形 + 上下两道轮缘
          滚花齿    —— 等间距竖向齿，带亮/暗边做出立体感，随 offset 平移
          固定基准  —— 中心上方一个红三角（不动），用来衬托齿在滚
          读数      —— 右侧显示 %+.0f pN
        """
        knob.delete("all")
        w = max(knob.winfo_width(), 200)
        cx, cy = w / 2.0, KNOB_H / 2.0
        off = float(ui["off"])

        DH = 40.0                      # 鼓身高
        top, bot = cy - DH / 2.0, cy + DH / 2.0
        DW = float(min(w * 0.42, 560))  # 鼓身宽
        x0, x1 = cx - DW / 2.0, cx + DW / 2.0

        # 齿条横向位移：每 KNOB_REF_PN 走 KNOB_REF_PX 像素
        shift = off / KNOB_REF_PN * KNOB_REF_PX
        pitch = float(KNOB_TOOTH_PITCH)
        phase = shift % pitch

        # 1) 鼓身底色
        knob.create_rectangle(x0, top, x1, bot, fill="#cfcfcf",
                              outline="")
        # 2) 滚花齿：从 x0-pitch 开始铺，超出鼓身的不画
        k = 0
        tx = x0 - pitch + phase
        while tx < x1:
            l, r = tx, tx + pitch * 0.62
            if r > x0 and l < x1:
                cl = max(l, x0)
                cr = min(r, x1)
                knob.create_rectangle(cl, top + 3, cr, bot - 3,
                                      fill="#a8a8a8", outline="")
                # 立体感：左边提亮，右边压暗
                knob.create_line(cl + 1, top + 3, cl + 1, bot - 3,
                                 fill="#e6e6e6", width=2)
                knob.create_line(cr - 1, top + 3, cr - 1, bot - 3,
                                 fill="#7d7d7d", width=2)
            tx += pitch
            k += 1

        # 3) 上下轮缘（把齿夹住，做出"圆柱侧面"的感觉）
        knob.create_line(x0, top, x1, top, fill="#4a4a4a", width=2)
        knob.create_line(x0, bot, x1, bot, fill="#4a4a4a", width=2)
        knob.create_line(x0, top + 2, x1, top + 2, fill="#f0f0f0", width=1)
        knob.create_line(x0, bot - 2, x1, bot - 2, fill="#909090", width=1)
        # 左右端头（圆柱的两端）
        knob.create_line(x0, top, x0, bot, fill="#4a4a4a", width=2)
        knob.create_line(x1, top, x1, bot, fill="#4a4a4a", width=2)

        # 4) 固定基准指针（不转）：中心上方红三角，尖朝下指住鼓身
        knob.create_polygon(cx, top - 3, cx - 7, top - 13, cx + 7, top - 13,
                            fill="#c00000", outline="#800000")
        # 中心一根细红线贯穿，方便看"哪一格正对着基准"
        knob.create_line(cx, top - 1, cx, bot + 1, fill="#c00000", width=1)

        # 5) 右侧读数与说明
        knob.create_text(x1 + 18, cy - 12, anchor="w",
                         text="%+.0f pN" % off, fill="#c00000",
                         font=("Consolas", 12, "bold"))
        knob.create_text(x1 + 18, cy + 6, anchor="w",
                         text="滚轮 / 左右拖 / ↑ ↓   每次 1 pN",
                         fill="#666", font=("Microsoft YaHei", 9))

    def knob_wheel(ev):
        """滚轮：上滚 +1 pN，下滚 -1 pN（像拧细准焦螺旋一格）。"""
        nudge(NUDGE_STEP if ev.delta > 0 else -NUDGE_STEP)
        return "break"

    def knob_press(ev):
        kui["drag"] = True
        kui["x0"] = ev.x
        kui["off0"] = float(ui["off"])

    def knob_move(ev):
        """左右拖：齿条跟着手走（1 像素 ≈ 1 pN），像拧旋钮。"""
        if not kui["drag"]:
            return
        dx = ev.x - kui.get("x0", ev.x)
        # 手指移动 dx 像素，就位移 dx 像素对应的 pN
        dpn = dx * KNOB_REF_PN / KNOB_REF_PX
        apply_off(round(kui.get("off0", 0.0) + dpn))

    def knob_release(ev=None):
        kui["drag"] = False

    knob.bind("<Button-1>", knob_press)
    knob.bind("<B1-Motion>", knob_move)
    knob.bind("<ButtonRelease-1>", knob_release)
    knob.bind("<MouseWheel>", knob_wheel)
    knob.bind("<Enter>", lambda e: knob.focus_set())
    # 窗口一布局好/被拉伸，就按新宽度重画，保证旋钮居中
    knob.bind("<Configure>", lambda e: knob_draw())

    # 键盘：绑到窗口 + 画布 + 旋钮，这样点过哪里按键都有效。
    # 注意要 return "break"：否则 cv 上触发后还会冒泡到 win 再触发一次（一下跳两格）。
    def bind_key(widget, stop=True):
        def wrap(fn):
            def h(ev):
                fn()
                return "break" if stop else None
            return h

        widget.bind("<Up>", wrap(lambda: nudge(NUDGE_STEP)))
        widget.bind("<Down>", wrap(lambda: nudge(-NUDGE_STEP)))

    bind_key(win, stop=False)          # 窗口本身不需要 break
    bind_key(cv)
    bind_key(knob)
    win.focus_force()                  # 打开就有焦点，不用先点一下

    # ---- 键位：每个键各管一件事，没有任何"等 330 ms 看是不是双击"的判断 ----
    #   左键拖 = 框选找峰      左键单击 = 加一个峰
    #   右键   = 删掉圈住的峰  中键单击 = 放 L₀ 锚点   中键拖 = 平移
    #   滚轮 = 缩放   Home = 复位视图   ↑↓ = 曲线平移 1 pN   ←→ = 锚点逐点移动
    cv.bind("<Button-1>", on_press1)
    cv.bind("<B1-Motion>", on_drag1)
    cv.bind("<ButtonRelease-1>", on_release1)
    cv.bind("<Button-3>", on_right)             # 右键：去掉一个峰（立即）
    cv.bind("<Button-2>", pan_start)            # 中键按下 + 拖动：平移
    cv.bind("<B2-Motion>", pan_move)
    cv.bind("<ButtonRelease-2>", pan_release)   # 中键没拖动 = 放锚点
    cv.bind("<MouseWheel>", zoom)               # 滚轮：缩放

    def home_reset(ev=None):
        reset()
        return "break"

    cv.bind("<Home>", home_reset)               # Home：复位视图
    win.bind("<Home>", home_reset)
    win.bind("<Left>", lambda e: nudge_anchor(-1))      # 方向键：锚点沿曲线移动
    win.bind("<Right>", lambda e: nudge_anchor(+1))
    win.bind("<Shift-Left>", lambda e: nudge_anchor(-10))
    win.bind("<Shift-Right>", lambda e: nudge_anchor(+10))
    win.bind("<Escape>", lambda e: (st.update(quit=True), setjob("quit")))

    draw()
    win.focus_force()
    try:
        win.after(60, sync_widgets)
        win.after(400, sync_widgets)
    except Exception:
        pass

    st["active"] = True
    while st["job"] is None and not st.get("quit"):
        # 小键盘 +/- ：直接跳到"峰数会变"的下一个阈值（跳一档，不是 +1）
        bump = st.get("bump") or 0
        if bump and re_detect is not None:
            st["bump"] = 0
            try:
                got = step_threshold(ladder, ui["thr"], 1 if bump > 0 else -1,
                                     len(ui["peaks"]), re_detect_box)
                if got is not None:
                    newp, ui["sig"], ui["thr"] = got
                    if ui["sig"] > 0:
                        PROM_SIGMA = max(0.1, ui["thr"] / ui["sig"])
                    ui["peaks"] = [list(q) for q in apply_edits(newp, ui["edits"])]
                    draw()
            except Exception:
                pass
        elif bump:
            st["bump"] = 0
        try:
            if not win.winfo_exists():
                break
            win.update()
        except tk.TclError:
            break
        time.sleep(0.02)
    st["active"] = False
    act = st["job"]
    try:
        win.destroy()
    except Exception:
        pass
    PREVIEW_WIN = None
    # 峰保持**原始坐标**返回，偏移量单独带出去：
    # refine() 会从原始 Y_re 重算 y，在这里加 off 会被它覆盖掉（踩过）。
    return [tuple(p) for p in ui["peaks"]], act, float(ui["off"])


# ================================================================
# Tk 收尾：把解释器彻底放开，避免同一内核里再跑一次起不了窗
# ================================================================
def close_tk(root):
    global PREVIEW_WIN
    try:
        if PREVIEW_WIN is not None:
            PREVIEW_WIN.destroy()
    except Exception:
        pass
    PREVIEW_WIN = None
    # 销毁前先把待处理的虚拟事件（ttk 的 <<ThemeChanged>> 之类）吃掉，
    # 否则解释器退出时会冒 "application has been destroyed" 的 Tcl 噪音。
    try:
        root.update_idletasks()
        root.update()
    except Exception:
        pass
    try:
        root.quit()
    except Exception:
        pass
    try:
        root.destroy()
    except Exception:
        pass
    try:
        if getattr(tk, "_default_root", None) is root:
            tk._default_root = None
    except Exception:
        pass
    try:
        import gc
        gc.collect()
    except Exception:
        pass


# ================================================================
# 主流程
# ================================================================
LAST_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "txtpeak_last.txt")


def load_last_folder():
    """返回上次用的文件夹（没有就空串）。第 2 行是老版本存的基线模式，忽略。"""
    folder = ""
    try:
        with open(LAST_FILE, "r", encoding="utf-8") as fh:
            lines = [x.strip() for x in fh.read().splitlines()]
        if lines and os.path.isdir(lines[0]):
            folder = lines[0]
    except Exception:
        pass
    return folder


def save_last_folder(p):
    try:
        with open(LAST_FILE, "w", encoding="utf-8") as fh:
            fh.write(p + "\n")
    except Exception:
        pass


def pick_folder(root, last=""):
    """用 Tk 的文件夹对话框选目录。

    这里**必须** deiconify()：Windows 只允许"已经显示出来的窗口"把自己的
    对话框带到前台。把根窗口藏起来的话，对话框会被别的窗口盖住 ——
    程序卡在模态循环里，既看不见框、Ctrl+C 也传不进去（这个坑踩过）。

    代价是会露出一个空白的 tk 小窗，这个是有意保留的。
    """
    try:
        root.deiconify()
        root.attributes("-topmost", True)
        root.update()
    except Exception:
        pass
    try:
        p = filedialog.askdirectory(
            parent=root,
            initialdir=last if (last and os.path.isdir(last)) else None,
            title="选一个文件夹（里面全是 JPK 导出的 txt）")
    except Exception as e:
        print("[x] 文件夹对话框出错：%s" % e)
        p = ""
    finally:
        try:
            root.attributes("-topmost", False)
            root.withdraw()
        except Exception:
            pass
    return p or ""


def run_folder():
    # 整个程序只创建一个 Tk 根；它同时充当文件夹对话框的父窗口
    root = tk.Tk()
    root.withdraw()

    load_wlc_cfg()
    last_folder = load_last_folder()
    print("基线：照 Igor FXImport 各自归零（靠近段减前200点均值、回撤段减末200点均值）")
    print("写 Excel：光标放在序号那一列的第一个格（如 A2），峰高从右边一格往右排；"
          "第一格写文件在本文件夹里的序号（标题 19/32 就是 19）；"
          "退格重做会覆盖原来那一行")

    folder = ""
    if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]):
        folder = sys.argv[1]
    else:
        folder = pick_folder(root, last_folder)

    if not folder or not os.path.isdir(folder):
        print("没选到文件夹，退出。")
        close_tk(root)
        return
    save_last_folder(folder)
    files = sorted(glob.glob(os.path.join(folder, "*.txt")))
    if not files:
        print("这个文件夹里没有 .txt")
        close_tk(root)
        return

    try:
        import keyboard as kb
    except Exception as e:
        print("[x] 没装 keyboard：%s" % e)
        close_tk(root)
        return

    st = {"job": None, "quit": False, "active": False, "stop": False,
          "bump": 0}

    def hot_f13():
        if st["active"]:
            st["job"] = "accept"

    def hot_skip():
        if st["active"]:
            st["job"] = "skip"

    def hot_back():
        if st["active"]:
            st["job"] = "back"

    def hot_plus():
        if st["active"]:
            st["bump"] = (st.get("bump") or 0) + 1

    def hot_minus():
        if st["active"]:
            st["bump"] = (st.get("bump") or 0) - 1

    def hot_esc():
        st["quit"] = True
        if st["active"]:
            st["job"] = "quit"

    try:
        kb.add_hotkey("f13", hot_f13)
        kb.add_hotkey("left ctrl", hot_skip)
        kb.on_press_key("backspace", lambda e: hot_back())
        kb.on_press_key("esc", lambda e: hot_esc())
        for _n, _f in (("add", hot_plus), ("subtract", hot_minus)):
            try:
                kb.add_hotkey(_n, _f)
            except Exception as e:
                print("[x] 小键盘 %s 注册失败：%s" % (_n, e))
    except Exception as e:
        print("[x] 热键注册失败：%s" % e)
        close_tk(root)
        return

    written = 0
    rewritten = 0
    # 序号直接用"文件在文件夹里的序号"（i+1），跟预览标题 19/32 的 19 一致，
    # 不再自己往下数 —— 跳过的文件在表里就是缺号，正好和标题对得上。
    rows = {}                    # 文件路径 -> (行号, 序号列, 序号, 上次写了几个峰)
    cache = {}

    def load_and_show(i, path):
        try:
            pr = process(load_jpk(path), read_k(path))
        except Exception as e:
            print("[x] 读取失败 %s：%s，跳过" % (os.path.basename(path), e))
            return None
        if not pr.get("base_ok", True):
            print("[!] %s：retract 末尾没有平坦段，零点是用最后一段的中位数估的，"
                  "这一条的峰高可能偏" % os.path.basename(path))
        detected, sig, thr = detect(pr)
        if path in cache:
            old_peaks, old_del, old_add = cache[path]
            peaks = [list(q) for q in old_peaks]
            edits = {"deleted": list(old_del), "added": list(old_add)}
        else:
            peaks = []                      # 新文件：先不标，等你左键框选
            edits = {"deleted": [], "added": []}
        ladder, _ = peak_ladder(pr)
        meta = read_header(path)
        kk = meta["k"]
        # L₀ 初值：从最右边那个峰反解（用自动判出来的峰，不影响显示）
        for xx, yy in reversed(detected):
            L = solve_L(xx, yy, float(WLC["p"]))
            if L is not None:
                WLC["l0"] = max(L0_MIN, min(L0_MAX, L))
                break
        peaks, act, off = preview(os.path.basename(path), (i + 1, len(files)),
                             pr, peaks, sig, thr, root, st,
                             re_detect=lambda t=None: detect(pr, prom=t),
                             edits=edits, ladder=ladder, kk=kk, wlc=WLC,
                             meta=meta)
        cache[path] = ([tuple(q) for q in peaks],
                       list(edits["deleted"]), list(edits["added"]))
        return pr, peaks, act, off

    try:
        i = 0
        while i < len(files):
            if st["quit"]:
                break
            got = load_and_show(i, files[i])
            if got is None:
                i += 1
                continue
            pr, peaks, act, off = got
            if st["quit"] or act in ("quit", None):
                break
            if act == "back":
                if i > 0:
                    i -= 1
                continue
            if act == "skip":
                i += 1
                continue
            if not peaks:
                i += 1
                continue
            fin = refine(pr, peaks)
            # 手动调高度：refine 给的是原始 y，补上偏移量 —— 这才是写进 Excel 的数
            ys = [y + off for _, y in fin]
            old = rows.get(files[i])
            if old is None:                     # 新的一行：序号 = 文件序号
                msg, rr, cc, ss = write_group(
                    ys, seq=i + 1, taken={v[0] for v in rows.values()},
                    pval=float(WLC["p"]), dl=float(WLC["dl"]))
            else:                               # 退格重做：写回原来那一行
                msg, rr, cc, ss = write_group(ys, seq=old[2], row=old[0],
                                              col=old[1], clear_to=old[3],
                                              pval=float(WLC["p"]),
                                              dl=float(WLC["dl"]))
            if not msg.startswith("已写入"):        # 只有出错才说话
                print("[x] " + msg)
            else:
                rows[files[i]] = (rr, cc, ss, len(ys))
                if old is None:
                    written += 1
                else:
                    rewritten += 1
                save_wlc_cfg()
                try:
                    append_wlc_result(
                        wlc_result_dir(folder),
                        os.path.splitext(os.path.basename(files[i]))[0],
                        float(WLC["p"]), float(WLC["dl"]))
                except Exception as e:
                    print("[x] 写 WLC 结果失败：%s" % e)
            i += 1
    finally:
        try:
            kb.unhook_all_hotkeys()
        except Exception:
            pass
        try:
            kb.unhook_all()
        except Exception:
            pass
        close_tk(root)
    print("已退出（新写 %d 行%s），键盘钩子已卸载。"
          % (written, "，退格覆盖重写 %d 次" % rewritten if rewritten else ""))


# ================================================================
# 自检（不弹窗、不碰 Excel）
# ================================================================
def selftest(path):
    print("== 自检 %s ==" % os.path.basename(path))
    pr = process(load_jpk(path))
    print("  点数 %d，基线 %.1f pN，x 范围 %.1f..%.1f nm"
          % (len(pr["D_re"]), pr["base"], pr["D_re"].min(), pr["D_re"].max()))
    peaks, sig, thr = detect(pr)
    print("  σ≈%.1f pN，PROM_SIGMA=%.0f，阈值 %.0f，检出 %d 个峰"
          % (sig, PROM_SIGMA, thr, len(peaks)))
    for k, (x, y) in enumerate(refine(pr, peaks), 1):
        print("    %2d  x=%7.1f nm   y=%7.1f pN  → %d" % (k, x, y, round(y)))


if __name__ == "__main__":
    if "--test" in sys.argv:
        k = sys.argv.index("--test")
        for p in sys.argv[k + 1:]:
            selftest(p)
    else:
        run_folder()
