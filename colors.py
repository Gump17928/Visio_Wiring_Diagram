"""Wire color names (as written in the spreadsheet) -> RGB."""

WIRE_COLORS = {
    "BLK": (0, 0, 0), "BLACK": (0, 0, 0),
    "RED": (220, 0, 0),
    "BLU": (0, 70, 220), "BLUE": (0, 70, 220),
    "GRN": (0, 150, 0), "GREEN": (0, 150, 0),
    "YEL": (230, 190, 0), "YELLOW": (230, 190, 0),
    "ORG": (240, 120, 0), "ORANGE": (240, 120, 0),
    "WHT": (150, 150, 150), "WHITE": (150, 150, 150),
    "BRN": (120, 70, 20), "BROWN": (120, 70, 20),
    "VIO": (130, 0, 170), "VIOLET": (130, 0, 170), "PUR": (130, 0, 170),
    "GRY": (110, 110, 110), "GRAY": (110, 110, 110), "GREY": (110, 110, 110),
}


def rgb(name):
    """RGB tuple for the first word of a color name, black if unknown."""
    first = name.replace("/", " ").replace("-", " ").split()[:1]
    return WIRE_COLORS.get(first[0].upper(), (0, 0, 0)) if first else (0, 0, 0)
