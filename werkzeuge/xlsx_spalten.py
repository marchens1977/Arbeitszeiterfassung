"""Hilfsprogramm: Spalte(n) in ein Excel-Blatt einfügen und ALLE Bezüge mitziehen
(Formeln auf allen Blättern, Namen, bedingte Formatierung, Gültigkeitsprüfungen,
verbundene Zellen, Spaltenbreiten, Fixierung) – so wie Excel es beim Einfügen tut.

Verwendung in Python:
    from xlsx_spalten import spalten_einfuegen
    spalten_einfuegen(wb, 'Dateneingabe', 6)   # neue leere Spalte F, alte F rutscht nach G
"""
import re
from copy import copy
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import column_index_from_string as c2i, get_column_letter as i2c
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.formatting.formatting import ConditionalFormattingList

_REF = re.compile(r"^(?:(?P<sheet>'[^']+'|[A-Za-z0-9_.ÄÖÜäöüß]+)!)?(?P<rng>\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?|\$?\d+:\$?\d+)$")
_PART = re.compile(r"(\$?)([A-Z]{1,3})(\$?\d*)")


def _shift_range(rng, idx, n):
    def f(m):
        col = c2i(m.group(2))
        if col >= idx:
            col += n
        return f"{m.group(1)}{i2c(col)}{m.group(3)}"
    if re.fullmatch(r"\$?\d+:\$?\d+", rng):   # ganze Zeilen – keine Spalten
        return rng
    return ":".join(_PART.sub(f, p, count=1) for p in rng.split(":"))


def formel_verschieben(formel, akt_blatt, ziel_blatt, idx, n):
    """Verschiebt in einer Formel alle Spaltenbezüge auf ziel_blatt ab Spalte idx um n."""
    if not isinstance(formel, str) or not formel.startswith("="):
        return formel
    tok = Tokenizer(formel)
    out = []
    for t in tok.items:
        v = t.value
        if t.type == "OPERAND" and t.subtype == "RANGE":
            m = _REF.match(v)
            if m:
                sheet = m.group("sheet")
                ziel = sheet.strip("'") if sheet else akt_blatt
                if ziel == ziel_blatt:
                    v = (sheet + "!" if sheet else "") + _shift_range(m.group("rng"), idx, n)
        out.append(v)
    return "=" + "".join(out)


def spalten_einfuegen(wb, blatt, idx, n=1):
    ws = wb[blatt]
    # 1) Formeln auf allen Blättern anpassen (vor dem Verschieben der Zellen)
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = formel_verschieben(c.value, w.title, blatt, idx, n)
        # Gültigkeitsprüfungen (Formeln) auf allen Blättern
        for dv in w.data_validations.dataValidation:
            for attr in ("formula1", "formula2"):
                val = getattr(dv, attr)
                if val:
                    setattr(dv, attr, formel_verschieben("=" + val, w.title, blatt, idx, n)[1:])
    # 2) Namen
    for name, dn in list(wb.defined_names.items()):
        dn.attr_text = formel_verschieben("=" + dn.attr_text, "", blatt, idx, n)[1:]
    # 3) Merges merken, entfernen
    merges = [str(m) for m in ws.merged_cells.ranges]
    for m in merges:
        ws.unmerge_cells(m)
    # 4) Zellen verschieben (Werte + Formate)
    ws.insert_cols(idx, n)
    # 5) Merges neu
    for m in merges:
        ws.merge_cells(_shift_range(m, idx, n))
    # 6) Bedingte Formatierung dieses Blatts
    alt = ws.conditional_formatting
    neu = ConditionalFormattingList()
    for cf in alt:
        rng = " ".join(_shift_range(str(r), idx, n) for r in cf.sqref.ranges)
        for rule in cf.rules:
            rule.formula = [formel_verschieben("=" + f, blatt, blatt, idx, n)[1:] for f in (rule.formula or [])]
            neu.add(rng, rule)
    ws.conditional_formatting = neu
    # 7) Gültigkeitsbereiche dieses Blatts
    for dv in ws.data_validations.dataValidation:
        dv.sqref = type(dv.sqref)(" ".join(_shift_range(str(r), idx, n) for r in dv.sqref.ranges))
    # 8) Spaltenbreiten / ausgeblendet
    dims = {k: (v.width, v.hidden) for k, v in ws.column_dimensions.items()}
    for k in list(ws.column_dimensions.keys()):
        del ws.column_dimensions[k]
    for k, (w_, h) in dims.items():
        col = c2i(k)
        nk = i2c(col + n if col >= idx else col)
        ws.column_dimensions[nk].width = w_
        ws.column_dimensions[nk].hidden = h
    # 9) Fixierung
    if ws.freeze_panes:
        ws.freeze_panes = _shift_range(ws.freeze_panes, idx, n)
    return ws


def stil_kopieren(ziel, quelle):
    for a in ("font", "fill", "border", "alignment", "number_format", "protection"):
        setattr(ziel, a, copy(getattr(quelle, a)))


def blatt_umbenennen(wb, alt, neu):
    """Benennt ein Blatt um und ersetzt alle Bezüge darauf (Formeln, Namen, Gültigkeitsprüfungen, bedingte Formatierung)."""
    import re
    muster = re.compile(r"(?<![A-Za-z0-9_.])'?" + re.escape(alt) + r"'?!")
    ziel = (f"'{neu}'!" if re.search(r"[^A-Za-z0-9_]", neu) else f"{neu}!")
    ers = lambda s: muster.sub(ziel, s) if isinstance(s, str) else s
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    c.value = ers(c.value)
        for dv in ws.data_validations.dataValidation:
            dv.formula1 = ers(dv.formula1); dv.formula2 = ers(dv.formula2)
        for cf in ws.conditional_formatting:
            for r in cf.rules:
                r.formula = [ers(f) for f in (r.formula or [])]
    for n, dn in wb.defined_names.items():
        dn.attr_text = ers(dn.attr_text)
    wb[alt].title = neu
