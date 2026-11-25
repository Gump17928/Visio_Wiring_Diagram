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
PIN_H = 0.22


def _set_prop(shape, name, value):
    if not value:
        return
    shape.AddNamedRow(SEC_PROP, name, 0)
    shape.CellsU(f"Prop.{name}").FormulaU = '"' + str(value).replace('"', '""') + '"'


def draw(layout, out_path):
    app = win32com.client.Dispatch("Visio.Application")
    doc = app.Documents.Add("")
    page = doc.Pages(1)
    page_h = layout["page"]["height"]
    page.PageSheet.CellsU("PageWidth").FormulaU = f"{layout['page']['width']} in"
    page.PageSheet.CellsU("PageHeight").FormulaU = f"{page_h} in"

    pins = {}  # (ref, pin) -> (pin shape, connection point row)
    for d in layout["devices"]:
        left, right = d["x"], d["x"] + d["w"]
        box = page.DrawRectangle(left, page_h - d["y"] - d["h"], right, page_h - d["y"])
        box.Name = d["ref"]
        box.Text = f"{d['ref']}\n{d['description']}".strip()
        box.CellsU("VerticalAlign").FormulaU = "0"  # header text at the top
        box.CellsU("FillForegnd").FormulaU = "RGB(255,255,255)"
        _set_prop(box, "RefDes", d["ref"])
        _set_prop(box, "Description", d["description"])

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
            row = pin.AddRow(SEC_CONNECTION_PTS, ROW_LAST, 0)
            pin.CellsSRC(SEC_CONNECTION_PTS, row, 0).FormulaU = "Width*1" if is_right else "Width*0"
            pin.CellsSRC(SEC_CONNECTION_PTS, row, 1).FormulaU = "Height*0.5"
            pins[(d["ref"], p["number"])] = (pin, row)

    connector_tool = app.ConnectorToolDataObject
    for w in layout["wires"]:
        a, a_row = pins[(w["from"]["ref"], w["from"]["pin"])]
        b, b_row = pins[(w["to"]["ref"], w["to"]["pin"])]
        conn = page.Drop(connector_tool, 0, 0)
        conn.CellsU("BeginX").GlueTo(a.CellsSRC(SEC_CONNECTION_PTS, a_row, 0))
        conn.CellsU("EndX").GlueTo(b.CellsSRC(SEC_CONNECTION_PTS, b_row, 0))
        conn.Name = f"Wire {w['id']}"
        conn.Text = w["label"]
        r, g, bl = rgb(w["color"])
        conn.CellsU("LineColor").FormulaU = f"RGB({r},{g},{bl})"
        conn.CellsU("EndArrow").FormulaU = "0"
        conn.CellsU("Char.Size").FormulaU = "7 pt"
        for key, name in (("id", "WireID"), ("signal", "Signal"), ("color", "Color"), ("gauge", "Gauge")):
            _set_prop(conn, name, w[key])

    doc.SaveAs(os.path.abspath(out_path))
    return doc


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    with open(sys.argv[1]) as f:
        draw(json.load(f), sys.argv[2])
