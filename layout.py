"""Automatic left-to-right layout. Produces a plain dict (JSON-serialisable), units = inches.

Coordinates are screen-style: origin top-left, y grows downward.
"""
import math

from grandalf.graphs import Edge, Graph, Vertex
from grandalf.layouts import SugiyamaLayout

PIN_PITCH = 0.25
HEADER_H = 0.55
CHAR_W = 0.085
MIN_BOX_W = 1.8
LAYER_GAP = 1.8  # horizontal room between columns for wires
NODE_GAP = 0.5  # vertical gap between boxes in a column
COMPONENT_GAP = 0.8
MARGIN = 0.5


class _View:
    def __init__(self, w, h):
        self.w, self.h, self.xy = w, h, (0, 0)


def _pin_label(pin):
    return f"{pin.number} {pin.name}".strip()


def _box_width(dev):
    longest_pin = max((len(_pin_label(p)) for p in dev.pins), default=0)
    header = max(len(dev.ref), len(dev.description))
    return max(MIN_BOX_W, 2 * longest_pin * CHAR_W + 0.4, header * CHAR_W * 1.3 + 0.4)


def _place(sizes, edges):
    """Layered layout. sizes: ref -> (w, h). Returns ref -> (cx, cy), x = rank axis."""
    verts = {ref: Vertex(ref) for ref in sizes}
    for ref, (w, h) in sizes.items():
        # grandalf lays out top-down; swap axes so ranks run left-to-right
        verts[ref].view = _View(h, w)
    graph = Graph(list(verts.values()), [Edge(verts[a], verts[b]) for a, b in edges])

    out, offset = {}, 0.0
    for comp in graph.C:
        layout = SugiyamaLayout(comp)
        layout.xspace, layout.yspace = NODE_GAP, LAYER_GAP
        layout.init_all()
        layout.draw()
        vs = comp.sV
        lo = min(v.view.xy[0] - v.view.w / 2 for v in vs)
        hi = max(v.view.xy[0] + v.view.w / 2 for v in vs)
        rank_lo = min(v.view.xy[1] - v.view.h / 2 for v in vs)
        for v in vs:
            out[v.data] = (v.view.xy[1] - rank_lo, v.view.xy[0] - lo + offset)
        offset += hi - lo + COMPONENT_GAP
    return out


def build_layout(devices, wires):
    by_ref = {d.ref: d for d in devices}
    widths = {d.ref: _box_width(d) for d in devices}
    edges = sorted({(w.from_ref, w.to_ref) for w in wires if w.from_ref != w.to_ref})

    # Pass 1: rough sizes (pins split over two sides) just to learn the column order.
    rough = {
        d.ref: (widths[d.ref], HEADER_H + math.ceil(len(d.pins) / 2) * PIN_PITCH)
        for d in devices
    }
    pos = _place(rough, edges)

    # Decide which side each pin sits on: facing the device it connects to.
    neighbours = {}  # (ref, pin) -> [(other_ref, is_from_end)]
    for w in wires:
        neighbours.setdefault((w.from_ref, w.from_pin), []).append((w.to_ref, True))
        neighbours.setdefault((w.to_ref, w.to_pin), []).append((w.from_ref, False))
    sides = {}
    for d in devices:
        for p in d.pins:
            near = neighbours.get((d.ref, p.number))
            if not near:
                sides[(d.ref, p.number)] = "left"
                continue
            dx = sum(pos[o][0] - pos[d.ref][0] for o, _ in near) / len(near)
            if abs(dx) < 1e-6:
                sides[(d.ref, p.number)] = "right" if near[0][1] else "left"
            else:
                sides[(d.ref, p.number)] = "right" if dx > 0 else "left"

    # Pass 2: real box heights.
    def count(ref, side):
        return sum(1 for p in by_ref[ref].pins if sides[(ref, p.number)] == side)

    sizes = {
        d.ref: (widths[d.ref], HEADER_H + max(count(d.ref, "left"), count(d.ref, "right"), 1) * PIN_PITCH)
        for d in devices
    }
    pos = _place(sizes, edges)

    # Order pins on each side by where their neighbour sits vertically (fewer crossings).
    boxes = []
    for d in devices:
        w, h = sizes[d.ref]
        cx, cy = pos[d.ref]
        pins = []
        for side in ("left", "right"):
            side_pins = [p for p in d.pins if sides[(d.ref, p.number)] == side]

            def key(p):
                near = neighbours.get((d.ref, p.number))
                return sum(pos[o][1] for o, _ in near) / len(near) if near else float("inf")

            for i, p in enumerate(sorted(side_pins, key=key)):
                pins.append(
                    {
                        "number": p.number,
                        "name": p.name,
                        "label": _pin_label(p),
                        "side": side,
                        "dy": HEADER_H + (i + 0.5) * PIN_PITCH,
                    }
                )
        boxes.append(
            {"ref": d.ref, "description": d.description, "x": cx - w / 2, "y": cy - h / 2, "w": w, "h": h, "pins": pins}
        )

    min_x = min(b["x"] for b in boxes)
    min_y = min(b["y"] for b in boxes)
    for b in boxes:
        b["x"] += MARGIN - min_x
        b["y"] += MARGIN - min_y

    return {
        "page": {
            "width": max(b["x"] + b["w"] for b in boxes) + MARGIN,
            "height": max(b["y"] + b["h"] for b in boxes) + MARGIN,
        },
        "devices": boxes,
        "wires": [
            {
                "id": w.id,
                "from": {"ref": w.from_ref, "pin": w.from_pin},
                "to": {"ref": w.to_ref, "pin": w.to_pin},
                "signal": w.signal,
                "color": w.color,
                "gauge": w.gauge,
                "label": " ".join(x for x in (w.id, w.signal) if x),
            }
            for w in wires
        ],
    }
