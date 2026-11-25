"""Quick PNG preview of a layout dict, so the layout can be checked without Visio."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

from colors import rgb


def pin_xy(layout, ref, pin):
    for d in layout["devices"]:
        if d["ref"] == ref:
            for p in d["pins"]:
                if p["number"] == pin:
                    x = d["x"] + (d["w"] if p["side"] == "right" else 0)
                    return x, d["y"] + p["dy"], p["side"]
    raise KeyError((ref, pin))


def render(layout, path):
    page = layout["page"]
    fig, ax = plt.subplots(figsize=(max(page["width"], 4), max(page["height"], 3)))
    for d in layout["devices"]:
        ax.add_patch(Rectangle((d["x"], d["y"]), d["w"], d["h"], fill=False, lw=1.5))
        ax.text(d["x"] + d["w"] / 2, d["y"] + 0.05, f"{d['ref']}\n{d['description']}",
                ha="center", va="top", fontsize=8, fontweight="bold")
        for p in d["pins"]:
            right = p["side"] == "right"
            ax.text(d["x"] + (d["w"] - 0.08 if right else 0.08), d["y"] + p["dy"], p["label"],
                    ha="right" if right else "left", va="center", fontsize=7)
    for w in layout["wires"]:
        x1, y1, s1 = pin_xy(layout, w["from"]["ref"], w["from"]["pin"])
        x2, y2, s2 = pin_xy(layout, w["to"]["ref"], w["to"]["pin"])
        mid = (x1 + x2) / 2
        color = tuple(c / 255 for c in rgb(w["color"]))
        ax.plot([x1, mid, mid, x2], [y1, y1, y2, y2], color=color, lw=1.2)
        ax.text(mid, (y1 + y2) / 2, w["label"], fontsize=6, ha="center", va="bottom", color=color)
    ax.set_xlim(0, page["width"])
    ax.set_ylim(page["height"], 0)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.savefig(path, dpi=130, bbox_inches="tight")
    plt.close(fig)
