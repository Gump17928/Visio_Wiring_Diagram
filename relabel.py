"""Switch wire labels on an existing drawing to live labels (W# + Signal by default).

Adds the Show WireID / Signal / Color / Gauge toggles to each page's Shape Data and replaces
every wire's text with a field built from its Shape Data. Box layout and hand edits to
positions are untouched; hand-typed wire text is replaced.

Usage:
  python relabel.py drawing.vsdx   # open, update and save the file
  python relabel.py                # update the drawing currently open in Visio (you save it)
"""
import os
import sys

import win32com.client

import visio_draw


def relabel_document(doc):
    count = 0
    for i in range(1, doc.Pages.Count + 1):
        page = doc.Pages(i)
        wires = [page.Shapes(j) for j in range(1, page.Shapes.Count + 1)]
        wires = [s for s in wires if s.CellExistsU("Prop.WireID", 0)]
        if not wires:
            continue
        visio_draw.add_label_toggles(page)
        for shape in wires:
            visio_draw.set_live_label(shape)
        count += len(wires)
    return count


def main():
    if len(sys.argv) > 2:
        sys.exit(__doc__)
    if len(sys.argv) == 2:
        app = win32com.client.Dispatch("Visio.Application")
        doc = app.Documents.Open(os.path.abspath(sys.argv[1]))
        n = relabel_document(doc)
        doc.Save()
        print(f"relabeled {n} wires in {sys.argv[1]} (saved)")
    else:
        try:
            doc = win32com.client.GetActiveObject("Visio.Application").ActiveDocument
        except Exception:
            doc = None
        if doc is None:
            sys.exit("No drawing is open in Visio. Open one, or pass a .vsdx path.")
        n = relabel_document(doc)
        print(f"relabeled {n} wires in {doc.Name} (not saved - save it in Visio)")


if __name__ == "__main__":
    main()
