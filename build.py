"""Spreadsheet -> layout.json (+ preview.png, + Visio drawing).

Examples:
  python build.py wiring.xlsx                      # validate, layout.json + preview.png
  python build.py wiring.xlsx --visio wiring.vsdx  # also draw in Visio (needs Visio + pywin32)
  python build.py devices.xlsx --wires wires.xlsx  # two separate files
"""
import argparse
import json
import sys

import layout as layout_mod
import netlist
import preview


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("devices", help="workbook with the Devices sheet (and Wires sheet unless --wires is given)")
    ap.add_argument("--wires", help="separate workbook with the Wires sheet")
    ap.add_argument("--json", default="layout.json")
    ap.add_argument("--png", default="preview.png")
    ap.add_argument("--visio", help="save a Visio drawing to this .vsdx path")
    args = ap.parse_args()

    try:
        devices, wires, warnings = netlist.load(args.devices, args.wires)
    except netlist.NetlistError as e:
        print("ERRORS:")
        for msg in e.errors:
            print("  -", msg)
        sys.exit(1)
    for msg in warnings:
        print("warning:", msg)

    layout = layout_mod.build_layout(devices, wires)
    with open(args.json, "w") as f:
        json.dump(layout, f, indent=2)
    preview.render(layout, args.png)
    print(f"{len(devices)} devices, {len(wires)} wires -> {args.json}, {args.png}")

    if args.visio:
        import visio_draw

        visio_draw.draw(layout, args.visio)
        print("saved", args.visio)


if __name__ == "__main__":
    main()
