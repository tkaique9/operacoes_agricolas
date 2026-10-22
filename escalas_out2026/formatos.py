"""Formatação de cargos e horários usada no PDF e no HTML."""
import re


ROMANOS = {"I", "II", "III", "IV", "V", "VI", "VII", "VIII", "MG", "ETA", "RH", "TI", "PCM"}
MINUSC = {"DE", "DA", "DO", "DOS", "DAS", "E", "EM"}


def cargo(txt):
    """OP. DE MAQUINAS AGRICOLAS II (MG) -> Op. de Maquinas Agricolas II (MG)."""
    out = []
    for i, w in enumerate(txt.split()):
        nu = w.strip("()").upper()
        if nu in ROMANOS:
            out.append(w.upper())
        elif i and nu in MINUSC:
            out.append(w.lower())
        elif w.startswith("("):
            out.append("(" + w[1:].capitalize())
        else:
            out.append(w.capitalize())
    return " ".join(out)


def horario(desc):
    """07:00 - 12:00e13:00 - 15:20 6x1 -> 07:00–12:00 / 13:00–15:20 6x1."""
    t = re.sub(r"\s*-\s*(?=\d{2}:)", "–", desc)
    t = re.sub(r"(\d{2}:\d{2})e(\d{2}:\d{2})", r"\1 / \2", t)
    t = re.sub(r"(?i)\s+(as|às)\s+", "–", t)
    t = re.sub(r"(\d{2}:\d{2}) (\d{2}:\d{2})", r"\1–\2", t)
    t = re.sub(r"(?i)\(inter 30 min\)", "(30 min)", t)
    if " - INT " in t:
        t = t.replace(" - INT ", " (pausa ") + ")"
    return t.replace("6X1–", "6x1 ").strip()
