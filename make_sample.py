"""Write a small example workbook (wiring_sample.xlsx) showing the expected sheet layout."""
import pandas as pd

devices = [
    ("PSU1", "24V Supply", "1", "+24V"), ("PSU1", "24V Supply", "2", "0V"), ("PSU1", "24V Supply", "3", "PE"),
    ("PLC1", "Main PLC", "1", "24V IN"), ("PLC1", "Main PLC", "2", "0V IN"),
    ("PLC1", "Main PLC", "3", "DO0"), ("PLC1", "Main PLC", "4", "DI0"), ("PLC1", "Main PLC", "5", "AI0"),
    ("K1", "Contactor", "A1", "Coil +"), ("K1", "Contactor", "A2", "Coil -"),
    ("SW1", "E-Stop", "1", "NC1"), ("SW1", "E-Stop", "2", "NC2"),
    ("S1", "Pressure Sensor", "1", "V+"), ("S1", "Pressure Sensor", "2", "V-"), ("S1", "Pressure Sensor", "3", "OUT"),
    ("J9", "Spare Connector", "1", "Spare"),
]
wires = [
    ("W1", "PSU1", "1", "PLC1", "1", "24V", "RED", "18AWG"),
    ("W2", "PSU1", "2", "PLC1", "2", "0V", "BLK", "18AWG"),
    ("W3", "PLC1", "3", "K1", "A1", "K1_CMD", "ORG", "20AWG"),
    ("W4", "PSU1", "2", "K1", "A2", "0V", "BLK", "20AWG"),
    ("W5", "PSU1", "1", "SW1", "1", "24V", "RED", "20AWG"),
    ("W6", "SW1", "2", "PLC1", "4", "ESTOP", "YEL", "20AWG"),
    ("W7", "PSU1", "1", "S1", "1", "24V", "RED", "22AWG"),
    ("W8", "PSU1", "2", "S1", "2", "0V", "BLK", "22AWG"),
    ("W9", "S1", "3", "PLC1", "5", "PRESS", "BLU", "22AWG"),
]

with pd.ExcelWriter("wiring_sample.xlsx") as xw:
    pd.DataFrame(devices, columns=["RefDes", "Description", "Pin", "PinName"]).to_excel(xw, sheet_name="Devices", index=False)
    pd.DataFrame(
        wires, columns=["WireID", "FromRef", "FromPin", "ToRef", "ToPin", "Signal", "Color", "Gauge"]
    ).to_excel(xw, sheet_name="Wires", index=False)
print("wrote wiring_sample.xlsx")
