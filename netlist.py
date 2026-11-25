"""Load and validate the wiring spreadsheets (devices + wires)."""
from dataclasses import dataclass, field

import pandas as pd

DEVICE_REQUIRED = ["RefDes", "Pin"]
DEVICE_OPTIONAL = ["Description", "PinName"]
WIRE_REQUIRED = ["WireID", "FromRef", "FromPin", "ToRef", "ToPin"]
WIRE_OPTIONAL = ["Signal", "Color", "Gauge"]


class NetlistError(Exception):
    def __init__(self, errors):
        self.errors = errors
        super().__init__("\n".join(errors))


@dataclass
class Pin:
    number: str
    name: str = ""


@dataclass
class Device:
    ref: str
    description: str = ""
    pins: list = field(default_factory=list)


@dataclass
class Wire:
    id: str
    from_ref: str
    from_pin: str
    to_ref: str
    to_pin: str
    signal: str = ""
    color: str = ""
    gauge: str = ""


def _read(path, sheet, required, optional):
    xl = pd.ExcelFile(path)
    name = sheet if sheet in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(name, dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise NetlistError([f"Sheet '{name}' in {path}: missing column(s) {missing}"])
    for c in optional:
        if c not in df.columns:
            df[c] = ""
    for c in df.columns:
        df[c] = df[c].str.strip()
    return df[df[required[0]] != ""].reset_index(drop=True)


def load(devices_path, wires_path=None):
    """Return (devices, wires, warnings). Raises NetlistError on hard errors."""
    wires_path = wires_path or devices_path
    ddf = _read(devices_path, "Devices", DEVICE_REQUIRED, DEVICE_OPTIONAL)
    wdf = _read(wires_path, "Wires", WIRE_REQUIRED, WIRE_OPTIONAL)
    errors, warnings = [], []

    devices = {}
    seen_pins = set()
    for i, r in ddf.iterrows():
        row = i + 2  # spreadsheet row number
        dev = devices.setdefault(r.RefDes, Device(r.RefDes))
        if r.Description and not dev.description:
            dev.description = r.Description
        if r.Pin == "":
            errors.append(f"Devices row {row}: {r.RefDes} has no Pin")
            continue
        if (r.RefDes, r.Pin) in seen_pins:
            errors.append(f"Devices row {row}: duplicate pin {r.RefDes}.{r.Pin}")
            continue
        seen_pins.add((r.RefDes, r.Pin))
        dev.pins.append(Pin(r.Pin, r.PinName))

    wires, wire_ids, pin_use = [], set(), {}
    for i, r in wdf.iterrows():
        row = i + 2
        w = Wire(r.WireID, r.FromRef, r.FromPin, r.ToRef, r.ToPin, r.Signal, r.Color, r.Gauge)
        if w.id in wire_ids:
            errors.append(f"Wires row {row}: duplicate WireID {w.id}")
        wire_ids.add(w.id)
        ok = True
        for ref, pin in ((w.from_ref, w.from_pin), (w.to_ref, w.to_pin)):
            if ref not in devices:
                errors.append(f"Wires row {row} ({w.id}): unknown device '{ref}'")
                ok = False
            elif (ref, pin) not in seen_pins:
                errors.append(f"Wires row {row} ({w.id}): {ref} has no pin '{pin}'")
                ok = False
        if (w.from_ref, w.from_pin) == (w.to_ref, w.to_pin):
            errors.append(f"Wires row {row} ({w.id}): connects a pin to itself")
            ok = False
        if ok:
            wires.append(w)
            for end in ((w.from_ref, w.from_pin), (w.to_ref, w.to_pin)):
                pin_use.setdefault(end, []).append(w.id)

    for (ref, pin), ids in pin_use.items():
        if len(ids) > 1:
            warnings.append(f"Pin {ref}.{pin} is used by several wires: {', '.join(ids)}")
    connected = {ref for ref, _ in pin_use}
    for ref in devices:
        if ref not in connected:
            warnings.append(f"Device {ref} has no wires")

    if errors:
        raise NetlistError(errors)
    return list(devices.values()), wires, warnings
