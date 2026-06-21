# -*- coding: utf-8 -*-
"""
词汇联想网络图 — 批量生成 (16 组: 青少年 初二/高二/大一/大三 + 成人 30/40/50/60, 各汉/英)

设计:
  * 以"边列表 CSV"构建无向加权网络
  * 节点单词字号 = 加权度 (weighted degree) -> 体现联想强度/weight
  * 节点单词颜色 = "原始网络结果" docx 中的 community 划分
  * community 环形播种 + 力导向精修 -> "中心核 + 社群花瓣"布局
  * 枢纽词防重叠后处理 (resolve_overlaps)
  * 深色学术配色, 角注网络指标
"""
import csv, math, glob, os, re, ast
import numpy as np
import networkx as nx
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.colors import to_rgba

CJK_FONT = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
DPI = 230
FIGIN = 22.0
OUTDIR = "网络图-输出"

# ----------------------------------------------------------------------------
# docx 解析: 建立 community 注册表
# ----------------------------------------------------------------------------
def parse_docx_blocks(path):
    import docx
    d = docx.Document(path)
    blocks, cur = [], None
    for p in d.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if re.match(r"^(初二|高二|大一|大三)(汉语|英语)$", t) or \
           re.search(r"network&communities$", t):
            cur = {"label": t, "comm": {}, "metrics": {}}
            blocks.append(cur)
            continue
        if cur is None:
            continue
        m = re.match(r"Community\s+(\d+):\s*(\{.*\})\s*$", t)
        if m:
            cid = int(m.group(1))
            try:
                members = ast.literal_eval(m.group(2))
            except Exception:
                members = set(re.findall(r"'([^']*)'", m.group(2)))
            for w in members:
                cur["comm"][w] = cid
            continue
        for key, pat in [("nodes", "节点数"), ("edges", "边数"),
                         ("apl", "平均最短路径"), ("clustering", "集聚系数"),
                         ("density", "密度"), ("avgdeg", "平均度"),
                         ("diameter", "直径"), ("modularity", "Modularity")]:
            if pat in t:
                num = re.findall(r"[-+]?\d*\.?\d+",
                                 t.split("：")[-1].split(":")[-1])
                if num:
                    cur["metrics"][key] = float(num[-1])
    return blocks


def build_registry():
    reg = []
    for f in sorted(glob.glob("原始网络结果-结项/*.docx")):
        if "network&communities" in f or "初二到大三" in f:
            for b in parse_docx_blocks(f):
                if b["comm"]:
                    b["nodeset"] = set(b["comm"])
                    b["src"] = os.path.basename(f)
                    reg.append(b)
    return reg


def match_block(nodes, reg):
    best, bj = None, -1
    for b in reg:
        inter = len(nodes & b["nodeset"])
        uni = len(nodes | b["nodeset"]) or 1
        j = inter / uni
        if j > bj:
            bj, best = j, b
    return best, bj


# ----------------------------------------------------------------------------
def load_graph(fn):
    G = nx.Graph()
    with open(fn, encoding="utf-8", errors="ignore") as f:
        r = csv.reader(f)
        next(r)
        for row in r:
            if len(row) < 3:
                continue
            s, t = row[0].strip(), row[1].strip()
            try:
                w = float(row[2])
            except ValueError:
                continue
            if not s or not t or s == t:
                continue
            if G.has_edge(s, t):
                G[s][t]["weight"] += w
            else:
                G.add_edge(s, t, weight=w)
    return G


def academic_palette(n):
    base = [
        "#4C9CD6", "#E8A33D", "#5CB37F", "#D9685F", "#9B82C4",
        "#C7895B", "#E07FB0", "#8DB04A", "#4FC4C4", "#C5B23C",
        "#7C97C9", "#D2706E", "#6FB59A", "#B57FB0", "#A0A36B",
        "#5BA0B5", "#CC8A4A", "#8FA8D0", "#D49AA0", "#74C08D",
    ]
    while len(base) < n:
        base += base
    return base[:n]


def community_aware_layout(G, comm, seed=42):
    rng = np.random.RandomState(seed)
    cids = sorted(set(comm.values()))
    sizes = {c: sum(1 for v in comm.values() if v == c) for c in cids}
    cids = sorted(cids, key=lambda c: -sizes[c])
    ang = {c: 2 * math.pi * i / len(cids) for i, c in enumerate(cids)}
    R = 7.5
    init = {}
    for n in G.nodes():
        c = comm.get(n, cids[0])
        cx, cy = R * math.cos(ang[c]), R * math.sin(ang[c])
        spread = 0.7 + 0.06 * math.sqrt(sizes[c])
        init[n] = np.array([cx + rng.normal(0, spread),
                            cy + rng.normal(0, spread)])
    k = 1.9 / math.sqrt(G.number_of_nodes())
    return nx.spring_layout(G, pos=init, k=k, iterations=90,
                            weight="weight", seed=seed)


def resolve_overlaps(pos, fsizes, latin, coord_span,
                     topk=110, iters=80):
    """Push the largest labels apart so hub words don't collide."""
    fig_px = FIGIN * DPI
    pxpt = DPI / 72.0
    unit = coord_span / fig_px            # layout units per pixel
    items = sorted(fsizes, key=lambda n: -fsizes[n])[:topk]
    hw, hh = {}, {}
    for n in items:
        h_px = fsizes[n] * pxpt
        cw = (0.62 if latin else 1.02) * h_px
        hw[n] = (len(n) * cw / 2.0) * unit * 1.05
        hh[n] = (h_px / 2.0) * unit * 1.35
    P = {n: np.array(pos[n], dtype=float) for n in items}
    for _ in range(iters):
        moved = False
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                a, b = items[i], items[j]
                d = P[a] - P[b]
                ox = (hw[a] + hw[b]) - abs(d[0])
                oy = (hh[a] + hh[b]) - abs(d[1])
                if ox > 0 and oy > 0:
                    moved = True
                    if ox < oy:
                        sh = ox / 2.0 * (1 if d[0] >= 0 else -1)
                        P[a][0] += sh
                        P[b][0] -= sh
                    else:
                        sh = oy / 2.0 * (1 if d[1] >= 0 else -1)
                        P[a][1] += sh
                        P[b][1] -= sh
        if not moved:
            break
    for n in items:
        pos[n] = P[n]
    return pos


def make_figure(name, edge_csv, block, out_png, latin):
    print(f"[{name}] {edge_csv}")
    G = load_graph(edge_csv)
    raw = block["comm"]
    stripped = {k.strip(): v for k, v in raw.items()}
    comm = {n: raw.get(n, stripped.get(n.strip(), 0)) for n in G.nodes()}
    cids = sorted(set(comm.values()))
    pal = dict(zip(cids, academic_palette(len(cids))))

    wdeg = dict(G.degree(weight="weight"))
    nodes = list(G.nodes())
    vals = np.array([wdeg[n] for n in nodes])
    norm = (vals - vals.min()) / (vals.max() - vals.min() + 1e-9)
    nrm = dict(zip(nodes, norm))

    pos = community_aware_layout(G, comm)
    xs = [p[0] for p in pos.values()]
    ys = [p[1] for p in pos.values()]
    coord_span = max(max(xs) - min(xs), max(ys) - min(ys))

    FMIN, FMAX = 3.8, 40.0
    fsizes = {n: FMIN + (FMAX - FMIN) * (nrm[n] ** 0.62) for n in nodes}
    pos = resolve_overlaps(pos, fsizes, latin, coord_span)

    # metrics (compute modularity if docx lacks it)
    m = dict(block["metrics"])
    if "modularity" not in m:
        groups = {}
        for n, c in comm.items():
            groups.setdefault(c, set()).add(n)
        m["modularity"] = nx.community.modularity(
            G, list(groups.values()), weight="weight")

    # ---- draw ----
    fig, ax = plt.subplots(figsize=(FIGIN, FIGIN), dpi=DPI)
    BG = "#0b0e14"
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.axis("off")

    ew = np.array([d["weight"] for _, _, d in G.edges(data=True)])
    ewn = ew / ew.max()
    nx.draw_networkx_edges(G, pos, ax=ax, edge_color="#9fb3c8",
                           width=0.12 + 0.5 * ewn, alpha=0.045)

    fp_cache = {}

    def fprop(sz):
        key = round(sz, 1)
        if key not in fp_cache:
            fp_cache[key] = (fm.FontProperties(family="DejaVu Sans", size=key)
                             if latin else
                             fm.FontProperties(fname=CJK_FONT, size=key))
        return fp_cache[key]

    for n in sorted(nodes, key=lambda x: fsizes[x]):
        x, y = pos[n]
        sz = fsizes[n]
        c = to_rgba(pal[comm[n]])
        weight = "bold" if sz > 14 else "normal"
        alpha = 0.95 if sz > 9 else (0.78 if sz > 6 else 0.6)
        ax.text(x, y, n, fontproperties=fprop(sz), color=c,
                ha="center", va="center", alpha=alpha,
                fontweight=weight, zorder=3)

    cap = (f"{name}\n"
           f"N = {G.number_of_nodes()} nodes   E = {G.number_of_edges()} edges\n"
           f"communities = {len(cids)}   modularity Q = {m['modularity']:.3f}\n"
           f"density = {m.get('density', float('nan')):.3f}   "
           f"mean degree = {m.get('avgdeg', float('nan')):.1f}")
    ax.text(0.012, 0.985, cap, transform=ax.transAxes,
            fontproperties=fm.FontProperties(fname=CJK_FONT, size=20),
            color="#c7d3e0", va="top", ha="left", linespacing=1.5,
            bbox=dict(boxstyle="round,pad=0.6", fc="#10151f",
                      ec="#2a3344", alpha=0.7))

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(out_png, facecolor=BG, dpi=DPI, bbox_inches="tight",
                pad_inches=0.2)
    plt.close(fig)
    print(f"  -> {out_png}  (comms={len(cids)}, Q={m['modularity']:.3f})")


# ----------------------------------------------------------------------------
# 中文展示名
LABEL_CN = {
    "初二": "初二", "高二": "高二", "大一": "大一", "大三": "大三",
    "30": "30岁组", "40": "40岁组", "50": "50岁组", "60": "60岁组",
}


def display_name(fn):
    base = os.path.basename(fn).replace("-边列表.csv", "")
    m = re.match(r"(初二|高二|大一|大三|30|40|50|60)(汉语|英语)", base)
    if m:
        return LABEL_CN.get(m.group(1), m.group(1)) + m.group(2)
    return base


if __name__ == "__main__":
    os.makedirs(OUTDIR, exist_ok=True)
    reg = build_registry()
    print(f"registry blocks: {len(reg)}\n")
    for csvf in sorted(glob.glob("边列表-结项/*.csv")):
        nodes = set()
        with open(csvf, encoding="utf-8", errors="ignore") as fh:
            r = csv.reader(fh)
            next(r)
            for row in r:
                if len(row) >= 3:
                    nodes.add(row[0].strip())
                    nodes.add(row[1].strip())
        block, j = match_block(nodes, reg)
        latin = "英语" in os.path.basename(csvf)
        name = display_name(csvf)
        print(f"match {name}: {block['src'][:26]} jaccard={j:.3f}")
        out = os.path.join(OUTDIR, f"{name}-网络图.png")
        make_figure(name, csvf, block, out, latin)
    print("\nALL DONE")
