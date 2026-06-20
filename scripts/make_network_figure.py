# -*- coding: utf-8 -*-
"""
词汇联想网络图生成脚本
- 以"边列表 CSV"为基础构建无向加权网络
- 节点标签字号由加权度 (weighted degree) 决定 -> 体现节点 weight/重要性
- 节点颜色由"原始网络结果"中的 community 划分决定
- 学术深色风格，参考已发表 JCL 文章中的网络图
"""
import csv, json, math, random, sys, os
import numpy as np
import networkx as nx
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.colors import to_rgba

random.seed(42)
np.random.seed(42)

CJK_FONT = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
GROUPS = json.load(open("/tmp/groups.json", encoding="utf-8"))

# ----------------------------------------------------------------------------
def load_graph(fn):
    G = nx.Graph()
    with open(fn, encoding="utf-8-sig") as f:
        r = csv.reader(f)
        next(r)
        for row in r:
            if len(row) < 3:
                continue
            s, t, w = row[0].strip(), row[1].strip(), float(row[2])
            if not s or not t or s == t:
                continue
            if G.has_edge(s, t):
                G[s][t]["weight"] += w
            else:
                G.add_edge(s, t, weight=w)
    return G


def community_map(group, nodes):
    raw = GROUPS[group]["comm"]
    stripped = {k.strip(): v for k, v in raw.items()}
    out = {}
    for n in nodes:
        if n in raw:
            out[n] = raw[n]
        elif n.strip() in stripped:
            out[n] = stripped[n.strip()]
        else:
            out[n] = 0
    return out


def academic_palette(n):
    """Muted but distinguishable colours for a dark academic background."""
    base = [
        "#4C9CD6", "#E8A33D", "#5CB37F", "#D9685F", "#9B82C4",
        "#C7895B", "#E07FB0", "#8DB04A", "#4FC4C4", "#C5B23C",
        "#7C97C9", "#D2706E", "#6FB59A", "#B57FB0", "#A0A36B",
        "#5BA0B5", "#CC8A4A", "#8FA8D0",
    ]
    while len(base) < n:
        base += base
    return base[:n]


def community_aware_layout(G, comm, seed=42):
    """Seed nodes around per-community centres on a circle, then refine with
    a force-directed (Fruchterman-Reingold) pass -> dense core + community
    'petals', matching the published aesthetic."""
    rng = np.random.RandomState(seed)
    cids = sorted(set(comm.values()))
    # order communities by size so big ones are spread out
    sizes = {c: sum(1 for v in comm.values() if v == c) for c in cids}
    cids = sorted(cids, key=lambda c: -sizes[c])
    ang = {c: 2 * math.pi * i / len(cids) for i, c in enumerate(cids)}
    R = 7.5
    init = {}
    for n in G.nodes():
        c = comm.get(n, 0)
        cx, cy = R * math.cos(ang[c]), R * math.sin(ang[c])
        # bigger communities occupy more area -> less central crowding
        spread = 0.7 + 0.06 * math.sqrt(sizes[c])
        init[n] = np.array([cx + rng.normal(0, spread),
                            cy + rng.normal(0, spread)])
    k = 1.9 / math.sqrt(G.number_of_nodes())
    pos = nx.spring_layout(
        G, pos=init, k=k, iterations=90, weight="weight", seed=seed
    )
    return pos


def make_figure(group, edge_csv, out_png, latin=False):
    print(f"[{group}] loading {edge_csv}")
    G = load_graph(edge_csv)
    print(f"  nodes={G.number_of_nodes()} edges={G.number_of_edges()}")

    comm = community_map(group, list(G.nodes()))
    cids = sorted(set(comm.values()))
    pal = dict(zip(cids, academic_palette(len(cids))))

    # node weight = weighted degree
    wdeg = dict(G.degree(weight="weight"))
    vals = np.array([wdeg[n] for n in G.nodes()])
    vmin, vmax = vals.min(), vals.max()
    norm = (vals - vmin) / (vmax - vmin + 1e-9)

    pos = community_aware_layout(G, comm)

    # ---- figure ----
    fig, ax = plt.subplots(figsize=(22, 22), dpi=230)
    BG = "#0b0e14"
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)
    ax.axis("off")

    # edges (light, low alpha) -- weight modulates width & opacity slightly
    ew = np.array([d["weight"] for _, _, d in G.edges(data=True)])
    ewn = ew / ew.max()
    nx.draw_networkx_edges(
        G, pos, ax=ax,
        edge_color="#9fb3c8",
        width=0.12 + 0.5 * ewn,
        alpha=0.045,
    )

    # node label font sizes from weighted degree (sqrt scaling -> few big hubs)
    FMIN, FMAX = 4.0, 40.0
    fsizes = {n: FMIN + (FMAX - FMIN) * (norm[i] ** 0.62)
              for i, n in enumerate(G.nodes())}

    fontpath = None if latin else CJK_FONT
    fp_cache = {}

    def fprop(sz):
        key = round(sz, 1)
        if key not in fp_cache:
            if fontpath:
                fp_cache[key] = fm.FontProperties(fname=fontpath, size=key)
            else:
                fp_cache[key] = fm.FontProperties(
                    family="DejaVu Sans", size=key)
        return fp_cache[key]

    # draw small labels first, hubs last (so hubs sit on top)
    order = sorted(G.nodes(), key=lambda n: fsizes[n])
    for n in order:
        x, y = pos[n]
        sz = fsizes[n]
        c = to_rgba(pal[comm.get(n, cids[0])])
        # heavier weight for hub words
        weight = "bold" if sz > 14 else "normal"
        alpha = 0.95 if sz > 9 else (0.78 if sz > 6 else 0.62)
        ax.text(
            x, y, n, fontproperties=fprop(sz),
            color=c, ha="center", va="center",
            alpha=alpha, fontweight=weight, zorder=3,
        )

    # metrics annotation (academic caption style)
    m = GROUPS[group]["metrics"]
    cap = (f"{group}\n"
           f"N = {int(m['nodes'])} nodes   E = {int(m['edges'])} edges\n"
           f"communities = {len(cids)}   modularity Q = {m['modularity']:.3f}\n"
           f"density = {m['density']:.3f}   mean degree = {m['avgdeg']:.1f}")
    txt_fp = fm.FontProperties(fname=CJK_FONT, size=20)
    ax.text(0.012, 0.985, cap, transform=ax.transAxes,
            fontproperties=txt_fp, color="#c7d3e0",
            va="top", ha="left", linespacing=1.5,
            bbox=dict(boxstyle="round,pad=0.6", fc="#10151f",
                      ec="#2a3344", alpha=0.7))

    ax.margins(0.02)
    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    fig.savefig(out_png, facecolor=BG, dpi=230, bbox_inches="tight",
                pad_inches=0.2)
    plt.close(fig)
    print(f"  saved -> {out_png}")


if __name__ == "__main__":
    os.makedirs("网络图-输出", exist_ok=True)
    make_figure("初二汉语", "边列表-结项/初二汉语-边列表.csv",
                "网络图-输出/初二汉语-网络图.png", latin=False)
    make_figure("初二英语", "边列表-结项/初二英语-边列表.csv",
                "网络图-输出/初二英语-网络图.png", latin=True)
