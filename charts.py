"""필수 차트 2종(master 6-3): 수치는 metrics.json에서만 읽는다."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker
from matplotlib import font_manager

NAVY, GREY, BURGUNDY = "#1B365D", "#8C99A6", "#B8293D"
for _f in ("Malgun Gothic", "NanumGothic", "Noto Sans CJK KR", "Noto Sans CJK JP", "Noto Sans CJK TC"):
    if any(_f == x.name for x in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = _f
        break
plt.rcParams["axes.unicode_minus"] = False


def chart_periods(b: dict) -> list:
    """연간 + 최신 누적만 쓴다(전년동기 누적은 뺀다)."""
    ps = b["periods"]
    cum = [i for i, p in enumerate(ps) if "누적" in p]
    return [i for i, p in enumerate(ps) if "누적" not in p] + cum[-1:]


def _style(ax):
    ax.tick_params(labelsize=7.5)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color="#E5E8EC", linewidth=0.6)
    ax.set_axisbelow(True)


def _save(fig, path):
    fig.set_size_inches(6.3, 3.2)
    fig.savefig(path, dpi=300, bbox_inches="tight")
    plt.close(fig)


def _row(b, k, idx):
    v = b["rows"].get(k) or [None] * len(b["periods"])
    return [v[i] for i in idx]


def sales_ebitda(metrics: dict, basis: str, path: str):
    b = metrics["bases"][basis]; idx = chart_periods(b)
    labs = [b["periods"][i].replace("(누적)", "\n누적") for i in idx]
    s, e, m = _row(b, "매출액", idx), _row(b, "EBITDA", idx), _row(b, "EBIT/매출액", idx)
    fig, ax = plt.subplots(); _style(ax)
    x = range(len(idx)); w = 0.38
    b1 = ax.bar([i - w / 2 for i in x], [v or 0 for v in s], w, color=NAVY, label="매출액")
    b2 = ax.bar([i + w / 2 for i in x], [v or 0 for v in e], w, color=GREY, label="EBITDA")
    for bars in (b1, b2):
        ax.bar_label(bars, fmt="{:,.0f}", fontsize=6.5, padding=2)
    ax.set_ylim(min(0, min(v or 0 for v in e) * 1.3), max(v or 0 for v in s) * 1.3)
    ax.set_xticks(list(x), labs, fontsize=7.5); ax.set_ylabel("억원", fontsize=7.5)
    ax2 = ax.twinx(); ax2.spines[["top"]].set_visible(False); ax2.tick_params(labelsize=7.5)
    pm = [v * 100 if v is not None else None for v in m]
    ax2.plot(list(x), pm, color=BURGUNDY, marker="o", label="영업이익률(EBIT/매출액)")
    for i, v in enumerate(pm):
        if v is not None:
            ax2.annotate(f"{v:.1f}%", (i, v), textcoords="offset points", xytext=(0, 8 if i % 2 else -12), ha="center", fontsize=6.5, color=BURGUNDY)
    top = max([v for v in pm if v is not None] or [1])
    ax2.set_ylim(-top * 1.4, top * 1.25); ax2.set_ylabel("%", fontsize=7.5)  # 꺾은선을 막대 위쪽 절반에 둔다
    ax2.set_yticks([t for t in ax2.get_yticks() if t >= 0])
    h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=3, frameon=False, fontsize=7)
    _save(fig, path)


def netdebt(metrics: dict, basis: str, path: str):
    b = metrics["bases"][basis]; idx = chart_periods(b)
    labs = [b["periods"][i].replace("(누적)", "\n누적") for i in idx]
    nd, r = _row(b, "순차입금(총차입금−금융자산)", idx), _row(b, "순차입금/EBITDA(배)", idx)
    eb = _row(b, "EBITDA", idx)
    fig, ax = plt.subplots(); _style(ax)
    x = list(range(len(idx)))
    bars = ax.bar(x, [v or 0 for v in nd], 0.55, color=[BURGUNDY if (v or 0) < 0 else NAVY for v in nd], label="순차입금")
    ax.bar_label(bars, fmt="{:,.0f}", fontsize=6.5, padding=2)
    lo, hi = min(0, min(v or 0 for v in nd)), max(v or 0 for v in nd)
    ax.set_ylim(lo * 1.3, hi * 1.3 if hi > 0 else 1)
    ax.set_xticks(x, labs, fontsize=7.5); ax.set_ylabel("억원", fontsize=7.5)
    ax2 = ax.twinx(); ax2.spines[["top"]].set_visible(False); ax2.tick_params(labelsize=7.5)
    pts = [(i, v) for i, v in zip(x, r) if v is not None and (eb[i] or 0) > 0]
    if pts:
        ax2.plot([p[0] for p in pts], [p[1] for p in pts], color=GREY, marker="o", label="순차입금/EBITDA(배)")
        for i, v in pts:
            ax2.annotate(f"{v:.1f}", (i, v), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=6.5)
        top = max(p[1] for p in pts) or 1
        ax2.set_ylim(-top * 1.4, top * 1.25)  # 꺾은선을 막대 위쪽 절반에 둔다
        ax2.set_yticks([t for t in ax2.get_yticks() if t >= 0])
    for i in x:
        if (eb[i] or 0) <= 0:
            ax2.annotate("N/M", (i, 0), textcoords="offset points", xytext=(0, 4), ha="center", fontsize=6.5, color=BURGUNDY)
    ax2.set_ylabel("배", fontsize=7.5)
    h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=2, frameon=False, fontsize=7)
    _save(fig, path)
