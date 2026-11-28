"""Draw a layout dict in Visio via COM. Needs desktop Visio + pywin32 (no Excel or pandas required).

Usage: python visio_draw.py layout.json out.vsdx
"""
import json
import os
import sys

import win32com.client

from colors import rgb

# Visio constants
SEC_CONNECTION_PTS = 7
ROW_LAST = -2
SEC_PROP = 243
ALIGN_LEFT, ALIGN_RIGHT = 0, 2
SELECT = 2  # visSelect
PROP_TYPE_BOOL = 3
FMT_STR_NORMAL = 37  # visFmtStrNormal
PIN_H = 0.22

# Wire Shape Data rows: (row name, layout key, shown in the label by default)
WIRE_FIELDS = (("WireID", "id", True), ("Signal", "signal", True), ("Color", "color", False), ("Gauge", "gauge", False))


def _q(value):
    return '"' + str(value).replace('"', '""') + '"'


def _set_prop(shape, name, value):
    if not value:
        return
    shape.AddNamedRow(SEC_PROP, name, 0)
    shape.CellsU(f"Prop.{name}").FormulaU = _q(value)


def _ensure_prop(sheet, name, label=None, type_=None, default=None):
    """Create Prop.<name> if missing (setting label/type/default); never overwrite an existing row."""
    if sheet.CellExistsU(f"Prop.{name}", 0):
        return
    sheet.AddNamedRow(SEC_PROP, name, 0)
    if label:
        sheet.CellsU(f"Prop.{name}.Label").FormulaU = _q(label)
    if type_ is not None:
        sheet.CellsU(f"Prop.{name}.Type").FormulaU = str(type_)
    if default is not None:
        sheet.CellsU(f"Prop.{name}").FormulaU = default


def add_label_toggles(page):
    """Page Shape Data switches (Show WireID / Signal / Color / Gauge) used by every wire label."""
    for name, _, shown in WIRE_FIELDS:
        _ensure_prop(page.PageSheet, f"Show{name}", f"Show wire {name}", PROP_TYPE_BOOL, "TRUE" if shown else "FALSE")


def set_live_label(conn):
    """Replace a wire's text with a field built from its Shape Data and the page toggles."""
    for name, _, _ in WIRE_FIELDS:
        _ensure_prop(conn, name, default='""')
    formula = "&".join(f'IF(ThePage!Prop.Show{name},Prop.{name}&" ","")' for name, _, _ in WIRE_FIELDS)
    conn.Characters.AddCustomFieldU(formula, FMT_STR_NORMAL)


def draw(layout, out_path):
    app = win32com.client.Dispatch("Visio.Application")
    doc = app.Documents.Add("")
    page = doc.Pages(1)
    page_h = layout["page"]["height"]
    page.PageSheet.CellsU("PageWidth").FormulaU = f"{layout['page']['width']} in"
    page.PageSheet.CellsU("PageHeight").FormulaU = f"{page_h} in"
    add_label_toggles(page)

    window = app.ActiveWindow
    pins = {}  # (ref, pin) -> (device group shape, connection point row)
    for d in layout["devices"]:
        left, right = d["x"], d["x"] + d["w"]
        box = page.DrawRectangle(left, page_h - d["y"] - d["h"], right, page_h - d["y"])
        box.Name = f"{d['ref']} body"
        box.Text = f"{d['ref']}\n{d['description']}".strip()
        box.CellsU("VerticalAlign").FormulaU = "0"  # header text at the top
        box.CellsU("FillForegnd").FormulaU = "RGB(255,255,255)"
        members = [box]

        for p in d["pins"]:
            is_right = p["side"] == "right"
            cy = page_h - (d["y"] + p["dy"])
            # pin label shape: the half of the box on the pin's side, flush with the box edge
            x1, x2 = (left + d["w"] / 2, right) if is_right else (left, left + d["w"] / 2)
            pin = page.DrawRectangle(x1, cy - PIN_H / 2, x2, cy + PIN_H / 2)
            pin.Name = f"{d['ref']}.{p['number']}"
            pin.Text = p["label"]
            pin.CellsU("LinePattern").FormulaU = "0"
            pin.CellsU("FillPattern").FormulaU = "0"
            pin.CellsU("Para.HorzAlign").FormulaU = str(ALIGN_RIGHT if is_right else ALIGN_LEFT)
            pin.CellsU("Char.Size").FormulaU = "7 pt"
            pin.CellsU("LeftMargin").FormulaU = "0.06 in"
            pin.CellsU("RightMargin").FormulaU = "0.06 in"
            members.append(pin)

        # Group box + pin labels so they move together, then add one connection point per
        # pin. Each connection point is driven by its label's position: Y follows the label,
        # X snaps to the nearest side of the box. Drag a label and its wire follows.
        window.DeselectAll()
        for m in members:
            window.Select(m, SELECT)
        group = window.Selection.Group()
        group.Name = d["ref"]
        _set_prop(group, "RefDes", d["ref"])
        _set_prop(group, "Description", d["description"])
        for p in d["pins"]:
            row = group.AddRow(SEC_CONNECTION_PTS, ROW_LAST, 0)
            pins[(d["ref"], p["number"])] = (group, row)
            # look the label up after grouping; its ID can change when it is grouped
            label = group.Shapes.ItemU(f"{d['ref']}.{p['number']}")
            try:
                src = f"Sheet.{label.ID}!PinX"
                group.CellsSRC(SEC_CONNECTION_PTS, row, 0).FormulaU = f"IF({src}<Width*0.5,0,Width)"
                group.CellsSRC(SEC_CONNECTION_PTS, row, 1).FormulaU = f"Sheet.{label.ID}!PinY"
            except Exception as e:  # fall back to a fixed connection point at the layout position
                print(f"warning: could not link {d['ref']}.{p['number']} to its label ({e}); using a fixed point")
                group.CellsSRC(SEC_CONNECTION_PTS, row, 0).FormulaU = "Width*1" if p["side"] == "right" else "Width*0"
                group.CellsSRC(SEC_CONNECTION_PTS, row, 1).FormulaU = f"Height*{1 - p['dy'] / d['h']:.6f}"
            try:
                # sub-shapes can't reference their parent by name, only as Sheet.<ID>
                label.CellsU("Para.HorzAlign").FormulaU = (
                    f"IF(PinX<Sheet.{group.ID}!Width*0.5,{ALIGN_LEFT},{ALIGN_RIGHT})"
                )
            except Exception as e:  # cosmetic only; the label keeps its original alignment
                print(f"warning: could not auto-align label {d['ref']}.{p['number']} ({e})")

    connector_tool = app.ConnectorToolDataObject
    for w in layout["wires"]:
        a, a_row = pins[(w["from"]["ref"], w["from"]["pin"])]
        b, b_row = pins[(w["to"]["ref"], w["to"]["pin"])]
        conn = page.Drop(connector_tool, 0, 0)
        conn.CellsU("BeginX").GlueTo(a.CellsSRC(SEC_CONNECTION_PTS, a_row, 0))
        conn.CellsU("EndX").GlueTo(b.CellsSRC(SEC_CONNECTION_PTS, b_row, 0))
        conn.Name = f"Wire {w['id']}"
        r, g, bl = rgb(w["color"])
        conn.CellsU("LineColor").FormulaU = f"RGB({r},{g},{bl})"
        conn.CellsU("EndArrow").FormulaU = "0"
        conn.CellsU("Char.Size").FormulaU = "7 pt"
        for name, key, _ in WIRE_FIELDS:
            _ensure_prop(conn, name, default=_q(w[key]))
        set_live_label(conn)

    doc.SaveAs(os.path.abspath(out_path))
    return doc


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    with open(sys.argv[1]) as f:
        draw(json.load(f), sys.argv[2])
