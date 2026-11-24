# -*- coding: utf-8 -*-  最后修改: 2026-11-24 12:26
"""
txt-multipeak.py —— 从上千/上万条 JPK 原始 .jpk-force 里挑出「多峰」的曲线

跟 txt-peak.py 的关系：**判峰方法照抄它**（同一套 process()/detect()，
同样的 PROM_SIGMA、COARSE_NM、SMOOTH_NM、SEARCH_NM、脱开峰剔除），
区别只是——txt-peak.py 读的是 JPK 软件导出的 txt，这个直接读 .jpk-force 压缩包，
所以不用先手工导出几万个文件。

【最常用：在 Spyder 里直接 F5 运行】会弹一个文件夹选择框，选完就开始扫，
扫完再弹一个结果框告诉你扫了多少条、挑出多少条、存在哪。

【扫哪些目录】给一个**总目录**（例如 D:\\data\\2026-11-R4），
会把**它下面第一层的每个子目录**当成一批，各自在自己的目录里建 multipeak 子文件夹：
    dataset\\exp-1-1ums-0.6pN-1s\\             ← 扫
        └─ multipeak\\  （结果放这儿）
    dataset\\exp-1-..._processed-2026.11.23-20.47.16\\     ← 跳过（见下）
跳过一个子目录的三种情况：
    1) 里面的 .jpk-force **少于 MIN_FILES（默认 10）条**（多半个把两个的残次采集）；
    2) 目录名里带 "_processed"（JPK 处理输出，不是原始数据）；
    3) 目录名就是输出目录本身（multipeak / 多峰 / _多峰），避免重跑套娃。
直接给某一个实验目录（里面有 .jpk-force）也能用，就只扫它自己。

用法（命令行可选）：
    python txt-multipeak.py                      # 弹窗选文件夹
    python txt-multipeak.py "D:\\data\\2026-11-R4"
    python txt-multipeak.py "D:\\...\\2026-11-R4" --dry-run
    python txt-multipeak.py --test "D:\\...\\dir"    # 自检：和 JPK 导出的 txt 比峰
参数：
    --min-peaks N   回撤段保留几个峰就算多峰（**默认 3**，2 峰的不算）
    --min-files N   少于 N 条 .jpk-force 的子目录整个跳过（**默认 10**）
    --dry-run       只统计，不复制文件
    --out NAME      输出子文件夹名（默认 multipeak，英文）
    --jobs N        并行进程数（默认多核；在 Spyder/IPython 里自动用单进程）
    --k V           手动指定弹簧常数（pN/nm，只写进清单，不影响判峰）
    --test DIR      自检模式：把同名 .txt（JPK 导出）当标准答案，比对峰数和峰高
    --follow        再往下钻，把更深的子目录也当一批（默认只到第一层）

结果（每个被扫的实验目录里新建 multipeak 子文件夹）：
    multipeak/*.jpk-force      多峰的那几条原样复制过去
    multipeak/多峰清单.csv      每个被挑中文件的峰数、峰位、峰高（UTF-8 BOM，Excel 直接开）
    multipeak/_跳过.txt        打不开/点数太少的文件，附原因
    multipeak/_全部统计.csv    该目录**每个**文件的峰数、σ、阈值（想自己重新定阈值时看这个）

------------------------------------------------------------
为什么能直接读 .jpk-force（2026-11-24 逆向结果，别改错）
------------------------------------------------------------
.jpk-force 就是个 zip：
    header.properties / shared-data/header.properties
    segments/<n>/segment-header.properties
    segments/<n>/channels/<通道>.dat
通道 .dat 是**大端** int16（Java 写出来的，跟 numpy 默认的小端相反），每个点 2 字节。
物理值 = 原始计数 × 倍率，倍率从 shared-data 的 lcd-info.N 里串起来：
    encoder.scaling.multiplier
      × conversion-set.conversion.<style>.scaling.multiplier   （沿 base-calibration-slot 往上走）
      × ... 直到 base-calibration-slot == conversion-set.conversions.base（通常 "volts"）
这个文件里 vDeflection 的链条是 encoder(3.093364338374452E-4)
→ distance(3.995192986269578E-8) → force(0.032228520471696505)，乘积 = 3.98299E-13 N/计数。
实测：用这套解码，四条通道和 JPK 自己导出的值 corr = 1.00000000、斜率比 = 1.000000。

两个坑（都踩过）：
 1) 别用 "<i2"（小端）读，出来的曲线是乱跳的噪声，还会骗你以为数据溢出要解绕
    （小端等于把大端值乘 256 再加一个随机低字节，看着像"周期性折叠"，其实压根没折叠）。
 2) 常数偏移不用管。文件里有 force-segment-header.baseline 和 force-baseline-adjust，
    JPK 导出时还会再扣一次；但 process() 里两条曲线各自归零（去程减前 200 点均值、
    回程减末尾平坦段均值），常数项自动消掉，只留倍率就够精确。
"""

import glob
import os
import re
import shutil
import sys
import zipfile

import numpy as np

# ================================================================ 可调参数
MIN_PEAKS = 3           # 回撤段（剔掉脱开峰后）还剩 ≥ 几个峰 → 算多峰（2 峰不要）
PROM_SIGMA = 12.0       # 判峰阈值 = PROM_SIGMA × 噪声σ，照抄 txt-peak.py
YMIN = 0.0              # 峰高下限（pN），0 = 归零后的基线以上才算
COARSE_NM = 3.0         # 定峰数用的粗平滑窗（nm）
SMOOTH_NM = 1.0         # 定峰顶用的细平滑窗（nm）
SEARCH_NM = 2.0         # 峰高回原始数据里取的半径（nm）
BASE_TOL = 40.0         # 回撤段末尾"平不平"的 std 容忍（pN）
TAIL_WINS = (200, 50, 20)
DROP_RUPTURE = True     # 剔掉最后一个脱开峰（多峰判据靠"不是脱开峰的那些"）
RUPTURE_RATIO = 1.35
RUPTURE_RECOVER = 0.30

OUT_NAME = u"multipeak"  # 输出子文件夹名（英文，避免路径里出中文）
MIN_FILES = 10          # 子目录里的 .jpk-force 少于这么多条就整个跳过（残次/空目录）
SKIP_TAG = u"_processed"  # 目录名带这个的跳过（JPK 处理输出，不是原始采集）
COPY_FILES = True       # False = 只统计不复制
K_OVERRIDE = None       # --k 手动给弹簧常数（pN/nm），.jpk-force 里没有这项

# ================================================================ 判峰内核
# 下面 smooth / noise_sigma / prominences / find_peaks / flat_window_base /
# tail_zero / diff_curve / drop_rupture / process / detect 全部照抄
# txt-peak.py（2026-11-15 版本），只删掉了跟画图/Excel 有关的函数。
# 要改判峰口径就改上面那几个参数，别在这里动算法，否则两边结果不一致。


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


def flat_window_base(arr, win=200, tol=BASE_TOL, region=0.40):
    n = len(arr)
    if n < 20:
        return float(np.median(arr)), False
    win = int(min(win, max(20, n // 4)))
    step = max(1, win // 2)
    hi = int(n * region)
    start = 0
    while start + win <= hi:
        w = arr[start:start + win]
        if float(w.std()) < tol:
            return float(np.median(w)), True
        start += step
    return float(np.median(arr[:min(max(win, n // 10), n)])), False


def tail_zero(y, wins=TAIL_WINS, tol=BASE_TOL):
    n = len(y)
    if n < 20:
        return float(np.mean(y)), n, u"点数太少（%d），零点用全段均值" % n
    tried = []
    for w in wins:
        seg = y[-min(int(w), n):]
        sd = float(seg.std())
        tried.append((len(seg), sd))
        if sd < tol:
            if len(tried) == 1:
                return float(np.mean(seg)), len(seg), ""
            return (float(np.mean(seg)), len(seg),
                    u"%s，改用末 %d 点（std %.0f）"
                    % (u"、".join(u"末尾 %d 点不平（std %.0f）" % t for t in tried[:-1]),
                       len(seg), sd))
    seg = y[-min(int(wins[0]), n):]
    return (float(np.mean(seg)), len(seg),
            u"%s，全都平不了，零点可能偏，结果仅供参考"
            % u"、".join(u"末 %d 点 std %.0f" % t for t in tried))


def diff_curve(raw):
    m = min(len(raw["hex"]), len(raw["hre"]))
    ha, fa = raw["hex"][:m], raw["fex"][:m]
    hb, fb = raw["hre"][:m], raw["fre"][:m]
    turn = min(ha.min(), hb.min())
    fa_i = np.interp(hb - turn, (ha - turn)[::-1], fa[::-1])
    _, ok = flat_window_base(fb - fa_i)
    return ok


def process(raw):
    z0 = float(np.min(raw["hex"]))
    Y_ex = -raw["fex"]
    Y_re = -raw["fre"]
    ex_base = float(np.mean(Y_ex[:200]))
    re_base, re_win, base_note = tail_zero(Y_re)
    Y_ex = Y_ex - ex_base
    Y_re = Y_re - re_base
    ok = diff_curve(raw)
    return dict(D_ex=raw["hex"] - z0, Y_ex=Y_ex,
                D_re=raw["hre"] - z0, Y_re=Y_re,
                base=re_base, base_win=re_win, base_note=base_note,
                z0=z0, base_ok=ok, mode="igor")


def drop_rupture(pr, peaks):
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


def detect(pr, prom=None):
    D, Y = pr["D_re"], pr["Y_re"]
    sig = float(noise_sigma(Y))
    dx = float(np.median(np.diff(D))) or 1.0
    wc = max(3, int(round(COARSE_NM / dx)) * 2 + 1)
    ys = smooth(Y, wc)
    thr = prom if prom else max(YMIN, PROM_SIGMA * sig)
    idx, _ = find_peaks(D, ys, thr, 8.0)
    wf = max(3, int(round(SMOOTH_NM / dx)) * 2 + 1)
    ys_f = smooth(Y, wf)
    win = max(SEARCH_NM, dx)
    peaks = []
    for i in idx:
        ii0 = np.nonzero(np.abs(D - D[i]) <= COARSE_NM * 2)[0]
        if len(ii0) == 0:
            continue
        k = ii0[int(np.argmax(ys_f[ii0]))]
        ii = np.nonzero(np.abs(D - D[k]) <= win)[0]
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


# ================================================================ 读 .jpk-force
S16 = ">i2"          # 大端 int16 —— 千万别改成 "<i2"


def _props(txt):
    out = {}
    for ln in txt.splitlines():
        ln = ln.strip()
        if not ln or ln.startswith("#") or "=" not in ln:
            continue
        k, v = ln.split("=", 1)
        out[k.strip()] = v.strip()
    return out


def _chains(shared):
    """lcd-info.N 的通道名 → (序号, 总倍率)。倍率沿 base-calibration-slot 串起来。"""
    blocks = {}
    for k, v in shared.items():
        if not k.startswith("lcd-info."):
            continue
        parts = k.split(".", 2)
        if len(parts) < 3:
            continue
        try:
            idx = int(parts[1])
        except ValueError:
            continue
        blocks.setdefault(idx, {})[parts[2]] = v
    res = {}
    for idx, b in blocks.items():
        name = b.get("channel.name")
        if not name:
            continue
        m = float(b.get("encoder.scaling.multiplier", 1.0))
        style = b.get("conversion-set.conversions.default")
        base = b.get("conversion-set.conversions.base", "volts")
        for _ in range(8):
            if not style:
                break
            m *= float(b.get("conversion-set.conversion.%s.scaling.multiplier" % style, 1.0))
            nxt = b.get("conversion-set.conversion.%s.base-calibration-slot" % style)
            if nxt is None or nxt == base:
                break
            style = nxt
        res[name] = (idx, m)
    return res


def _meta(zf):
    """.jpk-force 里的 k(pN/nm)、v(µm/s) —— 只在清单里显示用。

    注意：.jpk-force 本身**不存弹簧常数 k**（JPK 是导出时才从软件标定里补上的），
    所以 k 一般拿不到，除非命令行给 --k。v 从 relative-z-start/end 和扫描时间算。
    """
    meta = {"k": None, "v": None}
    try:
        h = _props(zf.read("header.properties").decode("utf-8", "replace"))

        def get(*suffixes):
            for key, val in h.items():
                for s in suffixes:
                    if key.endswith(s):
                        return val
            return None

        k = get("springConstant", "spring-constant")
        if k:
            kv = float(k) * 1e3
            if 1e-4 < kv < 1e6:
                meta["k"] = kv
        zs = get("force-settings.relative-z-start")
        ze = get("force-settings.relative-z-end")
        ts = get("force-settings.retrace-scan-time", "force-settings.retract-scan-time")
        if zs and ze and ts and float(ts) > 0:
            meta["v"] = abs(float(zs) - float(ze)) / float(ts) * 1e6
    except Exception:
        pass
    if meta["k"] is None and K_OVERRIDE:
        meta["k"] = K_OVERRIDE
    return meta


def load_jpk_force(path):
    """直接读 .jpk-force → txt-peak 的 raw 字典（hex/fex/hre/fre，单位 nm / pN）。

    取「去程(extend)」和「回程(retract)」两段，跳过停顿段。
    返回 (raw, meta)；读不了就抛异常，由调用方记进 _跳过.txt。
    """
    with zipfile.ZipFile(path) as zf:
        shared = _props(zf.read("shared-data/header.properties").decode("utf-8", "replace"))
        ch = _chains(shared)
        hname = None
        for cand in ("measuredHeight", "strainGaugeHeight", "height", "head-height"):
            if cand in ch:
                hname = cand
                break
        if "vDeflection" not in ch or hname is None:
            raise ValueError("缺通道（%s）" % ",".join(sorted(ch)))
        sh = ch[hname][1]
        sf = ch["vDeflection"][1]

        segs = {}
        for nm in zf.namelist():
            m = re.match(r"segments/(\d+)/segment-header\.properties$", nm)
            if not m:
                continue
            p = _props(zf.read(nm).decode("utf-8", "replace"))
            segs[int(m.group(1))] = p
        if not segs:
            raise ValueError("没有 segments/*/segment-header.properties")

        order = sorted(segs)
        ext = ret = None
        for i in order:
            nm = (segs[i].get("force-segment-header.name.name") or "").lower()
            if ext is None and ("extend" in nm or "approach" in nm):
                ext = i
            if ret is None and "retract" in nm:
                ret = i
        if ext is None:
            ext = order[0]
        if ret is None:
            ret = order[-1] if len(order) > 1 else None
        if ret is None or ret == ext:
            raise ValueError("找不到回撤段")

        def chan(seg, name):
            p = segs[seg]
            rel = p.get("channel.%s.data.file.name" % name) or ("channels/%s.dat" % name)
            buf = zf.read("segments/%d/%s" % (seg, rel))
            return np.frombuffer(buf, dtype=S16).astype(np.float64)

        h_ex = chan(ext, hname) * sh * 1e9
        f_ex = chan(ext, "vDeflection") * sf * 1e12
        h_re = chan(ret, hname) * sh * 1e9
        f_re = chan(ret, "vDeflection") * sf * 1e12
        meta = _meta(zf)
    n = min(len(h_ex), len(h_re))
    if n < 200:                      # 跟 txt-peak.py 一个门槛：每段不到 200 点没法归零
        raise ValueError("点数太少（%d）" % n)
    return ({"hex": h_ex, "fex": f_ex, "hre": h_re, "fre": f_re}, meta)


def load_txt_export(path):
    """读 JPK 导出的 txt（9 列），跟 load_jpk_force 返回同一种字典 —— 自检用。"""
    seg = None
    data, cols = {}, {}
    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        for ln in fh:
            if ln.startswith("# segmentIndex:"):
                seg = int(ln.split(":", 1)[1])
                data[seg] = []
            elif ln.startswith("# columns:"):
                cols[seg] = ln.split(":", 1)[1].split()
            elif ln.startswith("#") or not ln.strip():
                continue
            else:
                data[seg].append([float(x) for x in ln.split()])
    if not data:
        raise ValueError("txt 里没有数据段")
    order = sorted(data)
    ext, ret = order[0], order[-1]
    ca, cb = cols[ext], cols[ret]
    ha = np.array([r[ca.index("measuredHeight")] for r in data[ext]]) * 1e9
    fa = np.array([r[ca.index("vDeflection")] for r in data[ext]]) * 1e12
    hb = np.array([r[cb.index("measuredHeight")] for r in data[ret]]) * 1e9
    fb = np.array([r[cb.index("vDeflection")] for r in data[ret]]) * 1e12
    return {"hex": ha, "fex": fa, "hre": hb, "fre": fb}


def judge(path, min_peaks=MIN_PEAKS):
    """一条曲线 → (是否多峰, 峰列表, σ, 阈值, meta)。打不开就抛异常。"""
    raw, meta = load_jpk_force(path)
    pr = process(raw)
    peaks, sig, thr = detect(pr)
    return (len(peaks) >= min_peaks, peaks, sig, thr, meta)


# ================================================================ 扫目录
def _count_force(d):
    """目录本层有多少个 .jpk-force（不看子目录）。"""
    try:
        return sum(1 for f in os.listdir(d) if f.lower().endswith(".jpk-force"))
    except OSError:
        return 0


def _targets(src, follow=False, min_files=MIN_FILES, out_name=None):
    """要处理哪些目录，以及跳过了哪些（连原因）。

    返回 (dirs, skips)，skips = [(路径, 原因)]。
    规则：给总目录 → 只扫它**第一层**的子目录，各自建 multipeak；
    给的目录本层就有 .jpk-force → 只扫它自己。
    """
    src = os.path.abspath(src)
    dirs, skips = [], []
    out_names = tuple({OUT_NAME, u"多峰", u"_多峰", out_name or OUT_NAME})

    def take(p, explicit=False):
        """够格就收，不够格就记到 skips 里。

        explicit=True 表示这个目录是用户**亲手选的**：只挡输出目录，
        不按条数和 _processed 过滤（他明确点了就扫，哪怕只有几条）。
        """
        n = _count_force(p)
        name = os.path.basename(p) or p
        if name in out_names:
            skips.append((p, u"输出目录，不扫"))
            return
        if not explicit:
            if SKIP_TAG and SKIP_TAG in name:
                skips.append((p, u"名字带 %s（JPK 处理输出）" % SKIP_TAG))
                return
            if n == 0:
                skips.append((p, u"没有 .jpk-force"))
                return
            if n < max(1, int(min_files)):
                skips.append((p, u"只有 %d 条 < %d 条，整个跳过"
                              % (n, max(1, int(min_files)))))
                return
        dirs.append(p)

    if follow:
        for root, subdirs, files in os.walk(src):
            subdirs[:] = [d for d in subdirs
                          if not d.startswith(".") and d not in out_names]
            take(root)
        return dirs, skips

    if _count_force(src) > 0:                  # 直接指了一个实验目录
        take(src, explicit=True)
        return dirs, skips

    for d in sorted(os.listdir(src)):          # 总目录 → 看第一层
        p = os.path.join(src, d)
        if not os.path.isdir(p) or d.startswith("."):
            continue
        take(p)
    return dirs, skips


def _csv_writer(path):
    fh = open(path, "w", encoding="utf-8-sig", newline="")
    return fh


def _judge_one(job):
    """并行工作单元：必须放在模块顶层，多进程才认得出。"""
    p, min_peaks = job
    name = os.path.basename(p)
    try:
        ok, peaks, sig, thr, meta = judge(p, min_peaks)
        return (name, len(peaks), peaks, meta, sig, thr, None)
    except Exception as e:
        return (name, -1, [], None, 0.0, 0.0, "%s: %s" % (type(e).__name__, e))


def _iter_jobs(paths, min_peaks, jobs, quiet=False, tag=""):
    """按顺序吐结果；jobs>1 就开多进程。

    多进程一旦出错（Spyder/IPython 里子进程拿不到 __main__ 是常见情形），
    **整批作废重来**，退回单进程 —— 不能一半并行一半串行，否则同一个文件会被算两次。
    """
    jobs = max(1, int(jobs or 1))
    total = len(paths)
    res = []
    if jobs > 1:
        try:
            from concurrent.futures import ProcessPoolExecutor
            with ProcessPoolExecutor(max_workers=jobs) as ex:
                for r in ex.map(_judge_one, [(p, min_peaks) for p in paths], chunksize=16):
                    res.append(r)
                    if not quiet and len(res) % 500 == 0:
                        sys.stdout.write("\r    %d/%d ..." % (len(res), total))
                        sys.stdout.flush()
        except Exception as e:            # 多进程起不来 → 清空，改单进程重算
            res = []
            if not quiet:
                sys.stdout.write("\r    （多进程不可用：%s，改单进程）\n" % e)
    if len(res) == len(paths):
        for r in res:
            yield r
        return
    res, done = [], 0
    for p in paths:
        done += 1
        if not quiet and done % 200 == 0:
            sys.stdout.write("\r    %d/%d ..." % (done, total))
            sys.stdout.flush()
        yield _judge_one((p, min_peaks))


def run_dir(d, min_peaks=MIN_PEAKS, out_name=OUT_NAME, dry=False, quiet=False, jobs=1):
    files = sorted(f for f in os.listdir(d) if f.lower().endswith(".jpk-force"))
    if not files:
        return None
    out = os.path.join(d, out_name)
    got, bad, rows = [], [], []
    paths = [os.path.join(d, f) for f in files]
    for name, n, peaks, meta, sig, thr, err in _iter_jobs(paths, min_peaks, jobs, quiet):
        if err:
            bad.append((name, err))
            continue
        rows.append((name, n, sig, thr, meta))
        if n >= min_peaks:
            got.append((name, peaks, meta))
    if not quiet:
        from collections import Counter
        hist = Counter(r[1] for r in rows)
        top = "  ".join("%d峰:%d" % (k, hist[k]) for k in sorted(hist) if k > 0)
        sys.stdout.write("\r    %d 条扫完，多峰 %d 条%s\n"
                         % (len(files), len(got), ("，跳过 %d" % len(bad)) if bad else ""))
        if top:
            sys.stdout.write("    峰数分布（0峰 %d 条）：%s\n" % (hist.get(0, 0), top))
    if dry:
        return dict(dir=d, n=len(files), multi=len(got), bad=len(bad), out=None)
    if got:
        if not os.path.isdir(out):
            os.makedirs(out)
        if COPY_FILES:
            for name, peaks, meta in got:
                src = os.path.join(d, name)
                dst = os.path.join(out, name)
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)
        # 清单：一行一个文件，峰多的排前面，好先看最典型的那几条
        got.sort(key=lambda t: (-len(t[1]), t[0]))
        with _csv_writer(os.path.join(out, u"多峰清单.csv")) as fh:
            fh.write(u"文件,峰数,峰位(nm),峰高(pN),k(pN/nm),v(um/s)\n")
            for name, peaks, meta in got:
                ks = " / ".join("%.1f" % x for x, _ in peaks)
                ys = " / ".join("%.0f" % y for _, y in peaks)
                fh.write(u'"%s",%d,"%s","%s",%s,%s\n'
                         % (name, len(peaks), ks, ys,
                            "%.4f" % meta["k"] if meta["k"] else "",
                            "%.3f" % meta["v"] if meta["v"] else ""))
        # 全部统计（含不够多峰的），想自己改阈值/阈值时看它
        with _csv_writer(os.path.join(out, u"_全部统计.csv")) as fh:
            fh.write(u"文件,峰数,sigma(pN),阈值(pN),k(pN/nm),v(um/s)\n")
            for name, n, sig, thr, meta in rows:
                fh.write(u'"%s",%d,%.1f,%.1f,%s,%s\n'
                         % (name, n, sig, thr,
                            "%.4f" % meta["k"] if meta["k"] else "",
                            "%.3f" % meta["v"] if meta["v"] else ""))
        if bad:
            with open(os.path.join(out, u"_跳过.txt"), "w", encoding="utf-8-sig") as fh:
                for name, why in bad:
                    fh.write(u"%s\t%s\n" % (name, why))
    return dict(dir=d, n=len(files), multi=len(got), bad=len(bad),
                out=out if got else None)


def test_dir(d):
    """自检：同名 .txt（JPK 导出）当标准答案，比峰数和峰高。"""
    pairs = []
    for f in sorted(os.listdir(d)):
        if not f.lower().endswith(".jpk-force"):
            continue
        t = os.path.join(d, os.path.splitext(f)[0] + ".txt")
        if os.path.exists(t):
            pairs.append((os.path.join(d, f), t))
    if not pairs:
        print("这个目录没有同名 txt 可对照（先导出 txt 才能自检）")
        return 1
    same = diff_peak = 0
    worst = 0.0
    worst_at = ""
    for jpk, txt in pairs:
        try:
            pr = process(load_jpk_force(jpk)[0])
            pk_nat = detect(pr)[0]
            pr2 = process(load_txt_export(txt))
            pk_txt = detect(pr2)[0]
        except Exception as e:
            print("  FAIL  %s  %s: %s" % (os.path.basename(jpk), type(e).__name__, e))
            continue
        if len(pk_nat) == len(pk_txt):
            same += 1
            dp = max([abs(a[1] - b[1]) for a, b in zip(pk_nat, pk_txt)] or [0.0])
            if dp > worst:
                worst, worst_at = dp, os.path.basename(jpk)
            if dp > 1.0:
                diff_peak += 1
            print("  ok    %s  峰数 %d，峰高最大差 %.3f pN"
                  % (os.path.basename(jpk), len(pk_nat), dp))
        else:
            diff_peak += 1
            print("  DIFF  %s  原生 %d 峰 / 导出 %d 峰"
                  % (os.path.basename(jpk), len(pk_nat), len(pk_txt)))
    print("-" * 70)
    print("%d 对，峰数一致 %d 对，峰高差 >1 pN 的 %d 对；最大差 %.3f pN%s"
          % (len(pairs), same, diff_peak, worst, ("（" + worst_at + "）") if worst_at else ""))
    return 0 if diff_peak == 0 else 2


def _in_spyder():
    """在 Spyder / IPython 控制台里跑 —— 多进程子进程拿不到 __main__，默认单进程。"""
    try:
        if "IPython" in sys.modules or "spyder_kernels" in sys.modules:
            return True
        return bool(get_ipython())            # noqa: F821  （只在 IPython 里有定义）
    except Exception:
        return False


def _msgbox(title, text):
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(title, text)
        root.destroy()
    except Exception:
        print(text)


def pick_dir():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        d = filedialog.askdirectory(title=u"选一个装着 .jpk-force 的文件夹")
        root.destroy()
        return d
    except Exception as e:
        print("选文件夹失败（%s），请在命令行里给路径" % e)
        return ""


def main(argv):
    args = list(argv[1:])
    src = ""
    min_peaks = MIN_PEAKS
    out_name = OUT_NAME
    dry = False
    follow = False
    test = None
    jobs = 0
    min_files = MIN_FILES
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--min-peaks" and i + 1 < len(args):
            min_peaks = int(args[i + 1]); i += 2; continue
        if a == "--min-files" and i + 1 < len(args):
            min_files = int(args[i + 1]); i += 2; continue
        if a == "--out" and i + 1 < len(args):
            out_name = args[i + 1]; i += 2; continue
        if a == "--jobs" and i + 1 < len(args):
            jobs = int(args[i + 1]); i += 2; continue
        if a == "--k" and i + 1 < len(args):
            global K_OVERRIDE
            K_OVERRIDE = float(args[i + 1]); i += 2; continue
        if a == "--test":
            test = args[i + 1] if i + 1 < len(args) else ""
            i += 2; continue
        if a == "--dry-run":
            dry = True; i += 1; continue
        if a == "--follow":
            follow = True; i += 1; continue
        if a in ("-h", "--help"):
            print(__doc__); return 0
        src = a; i += 1
    if jobs <= 0:
        if _in_spyder():
            jobs = 1                      # Spyder/IPython 里多进程容易挂，直接用单进程
        else:
            try:
                jobs = max(1, min(8, (os.cpu_count() or 2) - 1))
            except Exception:
                jobs = 1
    if test is not None:
        return test_dir(test or src or pick_dir())
    interactive = False
    if not src:
        src = pick_dir()                  # Spyder 里 F5 运行：这里弹选择框
        if not src:
            return 1
        interactive = True
    dirs, skips = _targets(src, follow, min_files, out_name)
    print("源：%s" % src)
    if skips:
        print("-" * 70)
        print(u"跳过 %d 个子目录：" % len(skips))
        for p, why in skips:
            print(u"  × %s\n      %s" % (os.path.basename(p.rstrip("\\/")) or p, why))
        print("-" * 70)
    if not dirs:
        msg = (u"这个文件夹下没有可以扫的子目录。\n\n%s\n\n"
               u"（规则：子目录里要有 ≥%d 条 .jpk-force；名字带 %s 的跳过）"
               % (src, min_files, SKIP_TAG))
        print(msg)
        if interactive:
            _msgbox(u"没找到数据", msg)
        return 1
    print("批次：%d 个目录（跳过 %d 个）；多峰判据：回撤段剔掉脱开峰后 ≥ %d 个峰；并行 %d 进程"
          % (len(dirs), len(skips), min_peaks, jobs))
    tot = multi = 0
    for k, d in enumerate(dirs, 1):
        print("[%d/%d] %s" % (k, len(dirs), os.path.basename(d) or d))
        r = run_dir(d, min_peaks=min_peaks, out_name=out_name, dry=dry, jobs=jobs)
        if r:
            tot += r["n"]
            multi += r["multi"]
    print("=" * 70)
    print("共 %d 条，多峰 %d 条（%.1f%%）%s"
          % (tot, multi, (100.0 * multi / tot) if tot else 0.0,
             "（--dry-run，没复制）" if dry else ("，已复制到「%s」子文件夹" % out_name)))
    if interactive:
        _msgbox(u"筛选完成",
                u"%d 个目录、共 %d 条曲线\n判为多峰（≥%d 峰）：%d 条（%.1f%%）\n\n%s\n\n%s"
                % (len(dirs), tot, min_peaks, multi,
                   (100.0 * multi / tot) if tot else 0.0,
                   ("位置：" + os.path.join(src, out_name)) if len(dirs) == 1
                   else ("位置：各实验目录下的 %s 子文件夹" % out_name),
                   ("（--dry-run：只统计，没有复制文件）" if dry
                    else u"已复制原文件，清单见 CSV"))
                + (u"\n\n另外跳过 %d 个子目录（详见控制台）" % len(skips) if skips else ""))
    return 0


if __name__ == "__main__":
    main(sys.argv)
