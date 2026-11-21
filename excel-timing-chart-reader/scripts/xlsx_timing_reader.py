#!/usr/bin/env python3
"""
xlsx_timing_reader.py - read timing charts drawn in Excel, in any style, and work out what they say.

Understands timing charts drawn with: shapes/bars, bus hexagons, line waveforms (connectors or line
shapes), cell borders (top/bottom/left/right/diagonal), coloured cells, text boxes, callouts, arrows of
any kind (glued, loose, block arrows, double-headed dimension arrows), vertical marker lines, groups,
legends, and native charts / pictures nearby (reported, not parsed as waveforms).

Commands
  inventory BOOK                       what is on every tab
  scene     BOOK [--sheet S]           split each tab's drawings into regions, say which look like timing
                                       charts (and what else is there: legends, notes, flowcharts, pictures)
                                       -> scene_<tab>.json, overview_<tab>.png
  cells     BOOK --sheet S [--range A1:Z40] [--fills] [--borders]
  timing    BOOK --sheet S (--region R1[,R2] | --box C3:P20) [options]
            -> lanes, states over time, every transition (event), explicit arrows snapped to the
               transitions they connect, measured intervals, markers, clocks, setup/hold, handshakes,
               INFERRED links where no arrow was drawn, and a list of everything missing or unclear
               -> timing_*.md / .json / .png

Standard library only; matplotlib (optional) draws the PNGs.
"""
from __future__ import annotations
import argparse, json, math, os, re, sys, zipfile
from collections import defaultdict
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
PX, PT = 9525, 12700                         # EMU per pixel / per point
LINE_GEOMS = {"line", "straightConnector1", "bentConnector2", "bentConnector3", "bentConnector4",
              "bentConnector5", "curvedConnector2", "curvedConnector3", "curvedConnector4", "curvedConnector5"}
BLOCK_ARROWS = {"rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow",
                "notchedRightArrow", "stripedRightArrow"}
DECISION = {"diamond", "flowChartDecision"}
BUS_GEOMS = {"hexagon", "octagon", "parallelogram", "chevron", "homePlate", "flowChartPreparation", "trapezoid"}
FLOWISH = DECISION | {"flowChartProcess", "flowChartTerminator", "flowChartAlternateProcess", "flowChartConnector",
                      "flowChartDocument", "flowChartInputOutput", "flowChartPredefinedProcess", "flowChartMagneticDisk",
                      "flowChartPreparation", "flowChartManualInput", "ellipse", "parallelogram", "can"}
CALLOUT = re.compile(r"(callout|wedge|cloud)", re.I)
DATE_FMT_IDS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))
THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
               "hlink", "folHlink"]                                       # cell 'theme=' index order
SCHEME_ALIAS = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}


def q(t):
    p, n = t.split(":"); return f"{{{NS[p]}}}{n}"


def local(el):
    return el.tag.split("}")[-1]


def col_letter(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s


def col_index(s):
    n = 0
    for ch in s.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def parse_ref(ref):                          # 'C3' -> (row 1-based, col 0-based)
    m = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", ref.strip())
    return int(m.group(2)), col_index(m.group(1))


def fnum(v):
    if v is None:
        return "-"
    s = f"{v:.3f}".rstrip("0").rstrip(".")
    return "0" if s == "-0" else s


def serial_to_iso(v):
    d = datetime(1899, 12, 30) + timedelta(days=float(v))
    return d.strftime("%Y-%m-%d %H:%M") if abs(v - round(v)) > 1e-6 else d.strftime("%Y-%m-%d")


# ============================================================================ workbook + cells
class Grid:
    def __init__(self):
        self.def_col, self.def_row = 64 * PX, 20 * PX
        self.col_w, self.row_h = {}, {}
        self.cells, self.fill, self.dates, self.merges, self.border = {}, {}, set(), [], {}
        self._ce, self._re = [0], [0]

    def cx(self, c):
        while len(self._ce) <= c:
            i = len(self._ce) - 1; self._ce.append(self._ce[-1] + self.col_w.get(i, self.def_col))
        return self._ce[c]

    def ry(self, r):
        while len(self._re) <= r:
            i = len(self._re) - 1; self._re.append(self._re[-1] + self.row_h.get(i, self.def_row))
        return self._re[r]

    def col_at(self, x):
        c = 0
        while self.cx(c + 1) <= x:
            c += 1
        w = self.cx(c + 1) - self.cx(c)
        return c + ((x - self.cx(c)) / w if w else 0)

    def row_at(self, y):
        r = 0
        while self.ry(r + 1) <= y:
            r += 1
        h = self.ry(r + 1) - self.ry(r)
        return r + ((y - self.ry(r)) / h if h else 0)

    def ref(self, x, y):
        return f"{col_letter(int(self.col_at(x)))}{int(self.row_at(y)) + 1}"

    def rng(self, box):
        return f"{self.ref(box[0], box[1])}:{self.ref(max(box[0], box[2] - 1), max(box[1], box[3] - 1))}"

    def cell_box(self, r, c, r2=None, c2=None):           # 1-based row, 0-based col
        r2, c2 = r2 or r, c if c2 is None else c2
        return (self.cx(c), self.ry(r - 1), self.cx(c2 + 1), self.ry(r2))

    def show(self, r, c):
        v = self.cells.get((r, c))
        if isinstance(v, float):
            return serial_to_iso(v) if (r, c) in self.dates else fnum(v)
        return v


def width_px(w):
    return math.floor((256 * w + math.floor(128 / 7)) / 256 * 7)


class Book:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())
        self.shared = []
        if "xl/sharedStrings.xml" in self.names:
            for si in self.xml("xl/sharedStrings.xml").findall("main:si", NS):
                self.shared.append("".join(t.text or "" for t in si.iter(q("main:t"))))
        self.theme = self._theme()
        self.xf_fill, self.xf_date, self.xf_border = self._styles()
        rels = self.rels("xl/workbook.xml")
        self.sheets = []
        for s in self.xml("xl/workbook.xml").find("main:sheets", NS):
            self.sheets.append((s.get("name"), rels[s.get(q("r:id"))][1], s.get("state", "visible")))

    def xml(self, p):
        return ET.fromstring(self.z.read(p))

    def rels(self, part):
        d, f = os.path.split(part)
        rp, out = f"{d}/_rels/{f}.rels", {}
        if rp in self.names:
            for r in self.xml(rp).findall("rel:Relationship", NS):
                t = r.get("Target")
                full = t.lstrip("/") if t.startswith("/") else os.path.normpath(os.path.join(d, t)).replace("\\", "/")
                out[r.get("Id")] = (r.get("Type").rsplit("/", 1)[-1], full)
        return out

    def _theme(self):
        out = {}
        for n in self.names:
            if n.startswith("xl/theme/") and n.endswith(".xml"):
                cs = self.xml(n).find(".//a:clrScheme", NS)
                for el in (cs if cs is not None else []):
                    c = el.find("a:srgbClr", NS)
                    s = el.find("a:sysClr", NS)
                    out[local(el)] = "#" + (c.get("val") if c is not None else (s.get("lastClr", "000000") if s is not None else "000000"))
                break
        return out

    def _styles(self):
        fills, xf_fill, xf_date, borders, xf_border = [], [], [], [], []
        if "xl/styles.xml" not in self.names:
            return xf_fill, xf_date, xf_border
        st = self.xml("xl/styles.xml")
        custom = {}
        nf = st.find("main:numFmts", NS)
        for n in (list(nf) if nf is not None else []):
            custom[int(n.get("numFmtId"))] = n.get("formatCode", "")
        fl = st.find("main:fills", NS)
        for f in (list(fl) if fl is not None else []):
            pf = f.find("main:patternFill", NS)
            col = None
            if pf is not None and pf.get("patternType") not in (None, "none", "gray125"):
                fg = pf.find("main:fgColor", NS)
                col = self.cell_color(fg)
            fills.append(col)
        bl = st.find("main:borders", NS)
        for b in (list(bl) if bl is not None else []):
            sides = {}
            for side in ("left", "right", "top", "bottom"):
                e = b.find(f"main:{side}", NS)
                if e is not None and e.get("style") not in (None, "none"):
                    sides[side] = e.get("style")
            dg = b.find("main:diagonal", NS)
            if dg is not None and dg.get("style") not in (None, "none"):
                if b.get("diagonalUp") in ("1", "true"):
                    sides["diagUp"] = dg.get("style")
                if b.get("diagonalDown") in ("1", "true"):
                    sides["diagDown"] = dg.get("style")
            borders.append(sides)
        xfs = st.find("main:cellXfs", NS)
        for x in (list(xfs) if xfs is not None else []):
            fid, nid, bid = int(x.get("fillId", 0)), int(x.get("numFmtId", 0)), int(x.get("borderId", 0))
            xf_border.append(borders[bid] if bid < len(borders) else {})
            xf_fill.append(fills[fid] if fid < len(fills) else None)
            code = re.sub(r'"[^"]*"|\[[^\]]*\]', "", custom.get(nid, ""))
            xf_date.append(nid in DATE_FMT_IDS or bool(re.search(r"[dy]", code, re.I)))
        return xf_fill, xf_date, xf_border

    def cell_color(self, el):
        if el is None:
            return None
        if el.get("rgb"):
            return "#" + el.get("rgb")[-6:].upper()
        if el.get("theme") is not None:
            i = int(el.get("theme"))
            return self.theme.get(THEME_ORDER[i], f"theme:{i}") if i < len(THEME_ORDER) else f"theme:{i}"
        if el.get("indexed") is not None:
            return {"64": None, "65": None, "8": "#000000", "9": "#FFFFFF", "10": "#FF0000", "11": "#00FF00",
                    "12": "#0000FF", "13": "#FFFF00"}.get(el.get("indexed"), f"indexed:{el.get('indexed')}")
        return None

    def sheet_path(self, name):
        for n, p, _ in self.sheets:
            if n == name:
                return p
        sys.exit(f"No sheet named {name!r}. Sheets: {[n for n, _, _ in self.sheets]}")

    def grid(self, sp):
        root, g = self.xml(sp), Grid()
        f = root.find("main:sheetFormatPr", NS)
        if f is not None:
            if f.get("defaultColWidth"):
                g.def_col = width_px(float(f.get("defaultColWidth"))) * PX
            elif f.get("baseColWidth"):
                g.def_col = math.ceil((float(f.get("baseColWidth")) * 7 + 5) / 8) * 8 * PX
            if f.get("defaultRowHeight"):
                g.def_row = float(f.get("defaultRowHeight")) * PT
        cols = root.find("main:cols", NS)
        for c in (list(cols) if cols is not None else []):
            for i in range(int(c.get("min")) - 1, int(c.get("max"))):
                g.col_w[i] = 0 if c.get("hidden") in ("1", "true") else (
                    width_px(float(c.get("width"))) * PX if c.get("width") else g.def_col)
        data = root.find("main:sheetData", NS)
        for r in (list(data) if data is not None else []):
            ri = int(r.get("r")) - 1
            if r.get("ht"):
                g.row_h[ri] = float(r.get("ht")) * PT
            if r.get("hidden") in ("1", "true"):
                g.row_h[ri] = 0
            for c in r.findall("main:c", NS):
                row, col = parse_ref(c.get("r"))
                s = int(c.get("s", 0))
                if s < len(self.xf_fill) and self.xf_fill[s]:
                    g.fill[(row, col)] = self.xf_fill[s]
                if s < len(self.xf_border) and self.xf_border[s]:
                    g.border[(row, col)] = self.xf_border[s]
                t, v = c.get("t"), c.find("main:v", NS)
                if t == "s" and v is not None:
                    val = self.shared[int(v.text)]
                elif t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(q("main:t")))
                elif v is not None:
                    try:
                        val = float(v.text)
                        if s < len(self.xf_date) and self.xf_date[s]:
                            g.dates.add((row, col))
                    except ValueError:
                        val = v.text
                else:
                    continue
                if val != "":
                    g.cells[(row, col)] = val
        mc = root.find("main:mergeCells", NS)
        for m in (list(mc) if mc is not None else []):
            a, b = m.get("ref").split(":")
            (r0, c0), (r1, c1) = parse_ref(a), parse_ref(b)
            g.merges.append((r0, c0, r1, c1))
        return g

    def drawing(self, sp):
        for typ, full in self.rels(sp).values():
            if typ == "drawing":
                return full
        return None

    def scheme(self, name):
        return self.theme.get(SCHEME_ALIAS.get(name, name), f"scheme:{name}")


# ============================================================================ drawing objects
def _xfrm(pr):
    x = pr.find("a:xfrm", NS) if pr is not None else None
    if x is None or x.find("a:off", NS) is None:
        return None
    o, e = x.find("a:off", NS), x.find("a:ext", NS)
    d = dict(x=float(o.get("x")), y=float(o.get("y")), cx=float(e.get("cx")), cy=float(e.get("cy")),
             fh=x.get("flipH") in ("1", "true"), fv=x.get("flipV") in ("1", "true"), rot=int(x.get("rot", 0)) / 60000)
    co, ce = x.find("a:chOff", NS), x.find("a:chExt", NS)
    if co is not None:
        d.update(chx=float(co.get("x")), chy=float(co.get("y")), chcx=float(ce.get("cx")), chcy=float(ce.get("cy")))
    return d


def _unwrap(children):
    for el in children:
        if local(el) == "AlternateContent":
            br = el.find("mc:Choice", NS)
            if br is None or len(br) == 0:
                br = el.find("mc:Fallback", NS)
            if br is not None:
                yield from _unwrap(list(br))
        else:
            yield el


def resolve_ref(book, ref):
    """'Sheet 1'!$B$2:$B$4 -> list of cell values (reads the workbook itself)."""
    m = re.fullmatch(r"'?(.*?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", ref or "")
    if not m:
        return []
    sh = next((p for n, p, _ in book.sheets if n == m.group(1)), None)
    if not sh:
        return []
    g = book.grid(sh)
    c0, r0 = col_index(m.group(2)), int(m.group(3))
    c1, r1 = (col_index(m.group(4)), int(m.group(5))) if m.group(4) else (c0, r0)
    return [g.show(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)]


def read_chart(book, path):
    root = book.xml(path)
    title = " ".join(t.text or "" for t in root.findall("c:chart/c:title//a:t", NS)).strip()
    plots = []
    pa = root.find(".//c:plotArea", NS)
    for pl in (list(pa) if pa is not None else []):
        if not local(pl).endswith("Chart"):
            continue
        d = dict(type=local(pl), series=[])
        for k in ("barDir", "grouping"):
            e = pl.find(f"c:{k}", NS)
            if e is not None:
                d[k] = e.get("val")
        for s in pl.findall("c:ser", NS):
            def pts(tag):
                e = s.find(tag, NS)
                if e is None:
                    return None, []
                f = e.find(".//c:f", NS)
                vals = []
                for p in e.findall(".//c:pt", NS):
                    v = p.find("c:v", NS).text
                    try:
                        vals.append(float(v))
                    except (TypeError, ValueError):
                        vals.append(v)
                return (f.text if f is not None else None), vals
            nref, nm = pts("c:tx")
            cref, cat = pts("c:cat")
            if cref is None and not cat:
                cref, cat = pts("c:xVal")
            vref, val = pts("c:val")
            if vref is None and not val:
                vref, val = pts("c:yVal")
            sp = s.find("c:spPr", NS)
            nm = nm or resolve_ref(book, nref)
            cat = cat or resolve_ref(book, cref)
            val = val or resolve_ref(book, vref)
            d["series"].append(dict(name=nm[0] if nm else None, name_ref=nref, categories=cat, categories_ref=cref,
                                    values=val, values_ref=vref,
                                    invisible=sp is not None and sp.find("a:noFill", NS) is not None))
        plots.append(d)
    return dict(title=title, plots=plots)


def read_objects(book, dpath, g):
    root, objs = book.xml(dpath), []
    drels = book.rels(dpath)
    z = [0]

    def color(el):
        if el is None:
            return ""
        if el.find("a:noFill", NS) is not None:
            return "none"
        if el.find("a:pattFill", NS) is not None:
            return "pattern"                                          # hatched = unknown / don't-care
        sf = el.find("a:solidFill", NS)
        if sf is None:
            gf = el.find("a:gradFill", NS)
            if gf is not None:
                sf = gf.find(".//a:gs", NS)
        if sf is not None:
            c = sf.find("a:srgbClr", NS)
            if c is not None:
                return "#" + c.get("val").upper()
            c = sf.find("a:schemeClr", NS)
            if c is not None:
                return book.scheme(c.get("val")) + ("~" if len(c) else "")       # ~ = tinted/shaded variant
            c = sf.find("a:sysClr", NS)
            if c is not None:
                return "#" + c.get("lastClr", "000000")
        return ""

    def text(el):
        tb = el.find("xdr:txBody", NS)
        if tb is None:
            return ""
        paras = ["".join(t.text or "" for t in p.iter(q("a:t"))) for p in tb.iter(q("a:p"))]
        return " / ".join(p.strip() for p in paras if p.strip())

    def anchor_box(a):
        def pt(e):
            return (g.cx(int(e.find("xdr:col", NS).text)) + int(e.find("xdr:colOff", NS).text),
                    g.ry(int(e.find("xdr:row", NS).text)) + int(e.find("xdr:rowOff", NS).text))
        t = local(a)
        if t == "twoCellAnchor":
            return (*pt(a.find("xdr:from", NS)), *pt(a.find("xdr:to", NS)))
        if t == "oneCellAnchor":
            x, y = pt(a.find("xdr:from", NS)); e = a.find("xdr:ext", NS)
            return x, y, x + float(e.get("cx")), y + float(e.get("cy"))
        if t == "absoluteAnchor":
            p, e = a.find("xdr:pos", NS), a.find("xdr:ext", NS)
            x, y = float(p.get("x")), float(p.get("y"))
            return x, y, x + float(e.get("cx")), y + float(e.get("cy"))
        return None

    def visit(el, box, xf, group):
        t = local(el)
        z[0] += 1
        if t == "grpSp":
            gx = _xfrm(el.find("xdr:grpSpPr", NS))
            nv = el.find("xdr:nvGrpSpPr/xdr:cNvPr", NS)
            gid = (nv.get("name") or nv.get("id")) if nv is not None else None
            if group:
                gid = f"{group}/{gid}"
            sx = (box[2] - box[0]) / gx["chcx"] if gx and gx.get("chcx") else 1
            sy = (box[3] - box[1]) / gx["chcy"] if gx and gx.get("chcy") else 1
            for ch in _unwrap(list(el)):
                ct = local(ch)
                if ct not in ("sp", "cxnSp", "grpSp", "pic", "graphicFrame"):
                    continue
                pr = ch.find("xdr:grpSpPr" if ct == "grpSp" else ("xdr:xfrm" if ct == "graphicFrame" else "xdr:spPr"), NS)
                cx = _xfrm(pr) if ct != "graphicFrame" else None
                if ct == "graphicFrame" and pr is not None and pr.find("a:off", NS) is not None:
                    o_, e_ = pr.find("a:off", NS), pr.find("a:ext", NS)
                    cx = dict(x=float(o_.get("x")), y=float(o_.get("y")), cx=float(e_.get("cx")), cy=float(e_.get("cy")))
                if cx is None or gx is None or "chx" not in gx:
                    continue
                f = lambda X, Y: (box[0] + (X - gx["chx"]) * sx, box[1] + (Y - gx["chy"]) * sy)
                a, b = f(cx["x"], cx["y"]), f(cx["x"] + cx["cx"], cx["y"] + cx["cy"])
                visit(ch, (a[0], a[1], b[0], b[1]), cx, gid)
            return
        if t in ("pic", "graphicFrame"):
            nv = el.find(".//xdr:cNvPr", NS)
            o = dict(id=nv.get("id"), name=nv.get("name", ""), kind="picture", geom="", text=nv.get("descr", "") or "",
                     fill="", line="", dash="", width=0, group=group, box=box, z=z[0])
            if t == "pic":
                b = el.find(".//a:blip", NS)
                rid = b.get(q("r:embed")) if b is not None else None
                o["media"] = drels.get(rid, (None, None))[1]
            else:
                ch = el.find(".//c:chart", NS)
                if ch is not None:
                    o["kind"] = "chart"
                    path = drels.get(ch.get(q("r:id")), (None, None))[1]
                    o["chart"] = read_chart(book, path) if path in book.names else {}
                else:
                    o["kind"] = "object"
            objs.append(o)
            return
        if t not in ("sp", "cxnSp"):
            return
        nv = el.find("xdr:nvSpPr" if t == "sp" else "xdr:nvCxnSpPr", NS)
        c = nv.find("xdr:cNvPr", NS)
        pr = el.find("xdr:spPr", NS)
        pg = pr.find("a:prstGeom", NS) if pr is not None else None
        geom = pg.get("prst") if pg is not None else ("custom" if pr is not None and pr.find("a:custGeom", NS) is not None else "")
        xf = xf or _xfrm(pr) or {}
        ln = pr.find("a:ln", NS) if pr is not None else None
        # style may come from the theme style block (<xdr:style>) when spPr has none
        stl = el.find("xdr:style", NS)
        fill = color(pr)
        if not fill and stl is not None:
            fr = stl.find("a:fillRef", NS)
            sc = fr.find("a:schemeClr", NS) if fr is not None else None
            fill = book.scheme(sc.get("val")) if sc is not None and fr.get("idx") != "0" else fill
        line = color(ln)
        if not line and stl is not None:
            lr = stl.find("a:lnRef", NS)
            sc = lr.find("a:schemeClr", NS) if lr is not None else None
            line = book.scheme(sc.get("val")) if sc is not None and lr.get("idx") != "0" else line
        dash = ln.find("a:prstDash", NS) if ln is not None else None
        o = dict(id=c.get("id"), name=c.get("name", ""), geom=geom, text=text(el), fill=fill, line=line,
                 dash=dash.get("val") if dash is not None else "solid",
                 width=round(int(ln.get("w", 0)) / PT, 2) if ln is not None and ln.get("w") else 0,
                 group=group, box=box, z=z[0], rot=round(xf.get("rot", 0), 1))
        if t == "cxnSp" or geom in LINE_GEOMS:
            x0, y0, x1, y1 = box
            xs, xe = (x1, x0) if xf.get("fh") else (x0, x1)
            ys, ye = (y1, y0) if xf.get("fv") else (y0, y1)
            hs = he = ""
            if ln is not None:
                h, tl = ln.find("a:headEnd", NS), ln.find("a:tailEnd", NS)
                hs = h.get("type", "none") if h is not None else "none"
                he = tl.get("type", "none") if tl is not None else "none"
            hs, he = ("" if hs in ("", "none") else hs), ("" if he in ("", "none") else he)
            gs = ge = None
            cc = nv.find("xdr:cNvCxnSpPr", NS)
            if cc is not None:
                s, e = cc.find("a:stCxn", NS), cc.find("a:endCxn", NS)
                gs, ge = (s.get("id") if s is not None else None), (e.get("id") if e is not None else None)
            if hs and not he:                                         # head drawn at the start: flip direction
                xs, ys, xe, ye, gs, ge, hs, he = xe, ye, xs, ys, ge, gs, he, hs
            o.update(kind="arrow" if (hs or he) else "line", p0=(xs, ys), p1=(xe, ye), glue=(gs, ge),
                     double=bool(hs and he), head=he or hs, connector=t == "cxnSp" and not geom.startswith("line"),
                     bent=geom.startswith(("bent", "curved")))
        elif geom in BLOCK_ARROWS:
            x0, y0, x1, y1 = box; cy, cxm = (y0 + y1) / 2, (x0 + x1) / 2
            p0, p1 = {"leftArrow": ((x1, cy), (x0, cy)), "upArrow": ((cxm, y1), (cxm, y0)),
                      "downArrow": ((cxm, y0), (cxm, y1)), "upDownArrow": ((cxm, y0), (cxm, y1))
                      }.get(geom, ((x0, cy), (x1, cy)))
            o.update(kind="arrow", p0=p0, p1=p1, glue=(None, None), double=geom in ("leftRightArrow", "upDownArrow"),
                     head="block", connector=False, bent=False)
        else:
            sp_attr = nv.find("xdr:cNvSpPr", NS)
            is_tb = sp_attr is not None and sp_attr.get("txBox") in ("1", "true")
            if CALLOUT.search(geom or ""):
                kind = "callout"
            elif is_tb or (fill in ("none", "") and line in ("none", "") and o["text"]):
                kind = "text"
            else:
                kind = "shape"
            o["kind"] = kind
        objs.append(o)

    for a in _unwrap(list(root)):
        box = anchor_box(a)
        if box is None:
            continue
        for el in _unwrap(list(a)):
            if local(el) in ("sp", "cxnSp", "grpSp", "pic", "graphicFrame"):
                visit(el, box, None, None)
    for o in objs:
        x0, y0, x1, y1 = o["box"]
        o["box"] = (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
    return objs


def cell_bars(g):
    """Runs of same-coloured cells in a row become 'cellbar' objects (Gantt charts drawn with cell fills)."""
    out, by_row = [], defaultdict(list)
    for (r, c), col in g.fill.items():
        if col and col.upper() not in ("#FFFFFF",):
            by_row[r].append((c, col))
    n = 0
    for r, cs in sorted(by_row.items()):
        cs.sort()
        run = [cs[0]]
        for c, col in cs[1:] + [(None, None)]:
            if c is not None and c == run[-1][0] + 1 and col == run[-1][1]:
                run.append((c, col)); continue
            c0, c1 = run[0][0], run[-1][0]
            txt = " ".join(str(g.show(r, k)) for k in range(c0, c1 + 1) if (r, k) in g.cells)
            n += 1
            out.append(dict(id=f"cell{n}", name=f"{col_letter(c0)}{r}:{col_letter(c1)}{r}", kind="cellbar", geom="cells",
                            text=txt, fill=run[0][1], line="", dash="", width=0, group=None,
                            box=g.cell_box(r, c0, r, c1), z=0))
            if c is not None:
                run = [(c, col)]
    return out


DASHED_BORDERS = {"dashed", "dotted", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
                  "mediumDashDotDot", "slantDashDot"}


def _bline(n, p0, p1, st):
    return dict(id=f"border{n}", name=f"cell border ({st})", kind="line", geom="border", text="", fill="", line="border",
                dash="dash" if st in DASHED_BORDERS else "solid", width=0, group=None, z=0, p0=p0, p1=p1,
                glue=(None, None), double=False, head="", connector=False, bent=False,
                box=(min(p0[0], p1[0]), min(p0[1], p1[1]), max(p0[0], p1[0]), max(p0[1], p1[1])))


def border_lines(g):
    """Cell borders as line objects. Waveforms are often drawn with top/bottom (levels), left/right (edges)
    and diagonal (slanted transitions) borders. Fully boxed cells (table grids) are ignored."""
    H, V, out = defaultdict(list), defaultdict(list), []
    n = 0
    for (r, c), sides in g.border.items():
        if all(k in sides for k in ("left", "right", "top", "bottom")):
            continue
        x0, y0, x1, y1 = g.cell_box(r, c)
        for side, st in sides.items():
            if side == "top":
                H[y0].append((x0, x1, st))
            elif side == "bottom":
                H[y1].append((x0, x1, st))
            elif side == "left":
                V[x0].append((y0, y1, st))
            elif side == "right":
                V[x1].append((y0, y1, st))
            else:
                n += 1
                out.append(_bline(n, (x0, y1), (x1, y0), st) if side == "diagUp" else _bline(n, (x0, y0), (x1, y1), st))

    def merged(segs):
        res = []
        for a, b, st in sorted(set(segs)):
            if res and a <= res[-1][1] + 1 and st == res[-1][2]:
                res[-1][1] = max(res[-1][1], b)
            else:
                res.append([a, b, st])
        return res
    for y, segs in H.items():
        for a, b, st in merged(segs):
            n += 1; out.append(_bline(n, (a, y), (b, y), st))
    for x, segs in V.items():
        for a, b, st in merged(segs):
            n += 1; out.append(_bline(n, (x, a), (x, b), st))
    return out


# ============================================================================ geometry helpers
def gap(a, b):
    return max(a[0] - b[2], b[0] - a[2], 0), max(a[1] - b[3], b[1] - a[3], 0)


def d_pt_box(p, b):
    return math.hypot(max(b[0] - p[0], 0, p[0] - b[2]), max(b[1] - p[1], 0, p[1] - b[3]))


def d_pt_seg(p, a, b):
    ax, ay = a; bx, by = b; px, py = p
    L = (bx - ax) ** 2 + (by - ay) ** 2
    t = 0 if L == 0 else max(0, min(1, ((px - ax) * (bx - ax) + (py - ay) * (by - ay)) / L))
    return math.hypot(px - ax - t * (bx - ax), py - ay - t * (by - ay)), t


def inside(a, b, tol=2 * PX):               # box a inside box b
    return a[0] >= b[0] - tol and a[1] >= b[1] - tol and a[2] <= b[2] + tol and a[3] <= b[3] + tol


def area(b):
    return max(b[2] - b[0], 0) * max(b[3] - b[1], 0)


def center(b):
    return (b[0] + b[2]) / 2, (b[1] + b[3]) / 2


def is_line(o):
    return o["kind"] in ("arrow", "line")


def is_node(o):
    return o["kind"] in ("shape", "callout", "cellbar", "text")


def style_key(o):
    if is_line(o):
        return f"{o['kind']} {o['line'] or '-'} {o['dash']}{' double' if o.get('double') else ''}{' bent' if o.get('bent') else ''}"
    return f"{o['kind']} {o['geom']} fill={o['fill'] or '-'} line={o['line'] or '-'}{' ' + o['dash'] if o['dash'] not in ('solid', '') else ''}"


def attach_ends(objs, snap):
    """For each line/arrow: which node each end is glued to or touching."""
    nodes = [o for o in objs if is_node(o) and o["kind"] != "text"]
    byid = {o["id"]: o for o in objs}
    out = {}
    for a in objs:
        if not is_line(a):
            continue
        ends = []
        for gid, p in zip(a["glue"], (a["p0"], a["p1"])):
            if gid in byid and not is_line(byid[gid]):
                ends.append((gid, "glued")); continue
            cands = [n for n in nodes if d_pt_box(p, n["box"]) <= snap]
            if cands:
                # prefer the smallest node (an inner box rather than its container)
                n = min(cands, key=lambda n: (d_pt_box(p, n["box"]) > 1 * PX, area(n["box"])))
                ends.append((n["id"], "touching"))
            else:
                ends.append((None, "loose"))
        out[a["id"]] = ends
    return out


def line_paths(a):
    """Possible drawn paths of a line: straight, or the elbow routes of bent/curved connectors."""
    (x0, y0), (x1, y1) = a["p0"], a["p1"]
    if not a.get("bent"):
        return [[(x0, y0), (x1, y1)]]
    mx, my = (x0 + x1) / 2, (y0 + y1) / 2
    return [[(x0, y0), (x1, y0), (x1, y1)], [(x0, y0), (x0, y1), (x1, y1)],
            [(x0, y0), (mx, y0), (mx, y1), (x1, y1)], [(x0, y0), (x0, my), (x1, my), (x1, y1)]]


def label_for_line(a, texts, used):
    """A free text sitting on/next to a line (away from its two ends) labels it."""
    best = None
    for t in texts:
        if t["id"] in used:
            continue
        c = center(t["box"])
        h = (t["box"][3] - t["box"][1]) / 2
        if min(math.dist(c, a["p0"]), math.dist(c, a["p1"])) < 4 * PX:
            continue
        for path in line_paths(a):
            for k, (u, v) in enumerate(zip(path, path[1:])):
                d, tt = d_pt_seg(c, u, v)
                inner = (0.05 <= tt <= 0.95) if len(path) == 2 else True
                if d <= h + 22 * PX and inner and (best is None or d < best[0]):
                    best = (d, t)
    return best[1] if best else None


# ============================================================================ axes and labels in cells
TICK_TXT = re.compile(r"\s*[tT]?\s*=?\s*(-?\d+(?:\.\d+)?)\s*([a-zA-Zµ%]{0,4})\s*")


def _num(g, r, c):
    v = g.cells.get((r, c))
    if isinstance(v, float):
        return v, (r, c) in g.dates
    if isinstance(v, str):
        m = TICK_TXT.fullmatch(v)
        if m:
            return float(m.group(1)), False
    return None, False


def find_axis(g, box, direction):
    """Look for a row (direction 'x') or column ('y') of ticks spanning the box.
    Returns dict(kind=number|date|category, line=row-or-col, ticks=[(index, value, label)]) or None."""
    c0, c1 = int(g.col_at(box[0])), int(g.col_at(max(box[0], box[2] - 1)))
    r0, r1 = int(g.row_at(box[1])) + 1, int(g.row_at(max(box[1], box[3] - 1))) + 1
    if direction == "x":
        lines = list(range(r0 - 1, max(r0 - 8, 1) - 1, -1)) + list(range(r1 + 1, r1 + 4)) + list(range(r0, r1 + 1))
        span = range(max(c0 - 1, 0), c1 + 2); need = c1 - c0 + 1
        cell = lambda L, k: (L, k)
    else:
        lines = list(range(c0 - 1, max(c0 - 6, 0) - 1, -1)) + list(range(c1 + 1, c1 + 3))
        span = range(max(r0 - 1, 1), r1 + 2); need = r1 - r0 + 1
        cell = lambda L, k: (k, L)
    for L in lines:
        nums, cats = [], []
        for k in span:
            rc = cell(L, k)
            v, isd = _num(g, *rc)
            if v is not None:
                nums.append((k, v, isd, g.show(*rc)))
            elif isinstance(g.cells.get(rc), str):
                cats.append((k, g.cells[rc]))
        vals = [v for _, v, _, _ in nums]
        if len(nums) >= 3 and all(b > a for a, b in zip(vals, vals[1:])) and (nums[-1][0] - nums[0][0] + 1) >= 0.5 * need:
            kind = "date" if sum(1 for n in nums if n[2]) >= len(nums) / 2 else "number"
            return dict(kind=kind, direction=direction, line=L, ticks=[(k, v, lab) for k, v, _, lab in nums])
        if direction == "x" and len(cats) >= 3 and (cats[-1][0] - cats[0][0] + 1) >= 0.5 * need and \
                len(cats) >= 0.6 * (cats[-1][0] - cats[0][0] + 1) and not nums:
            return dict(kind="category", direction=direction, line=L, ticks=[(k, i, lab) for i, (k, lab) in enumerate(cats)])
    return None


def side_labels(g, box, direction):
    """Text cells beside the box: lane names. direction 'x' -> a column on the left; 'y' -> a row above."""
    c0, c1 = int(g.col_at(box[0])), int(g.col_at(max(box[0], box[2] - 1)))
    r0, r1 = int(g.row_at(box[1])) + 1, int(g.row_at(max(box[1], box[3] - 1))) + 1
    found = defaultdict(dict)
    for (r, c), v in g.cells.items():
        if not isinstance(v, str) or not v.strip():
            continue
        if direction == "x" and r0 <= r <= r1 and c0 - 6 <= c < c0:
            found[c][r] = v.strip()
        if direction == "y" and c0 <= c <= c1 and r0 - 4 <= r < r0:
            found[r][c] = v.strip()
    if not found:
        return None, {}
    key = max(found, key=lambda k: (len(found[k]), k))
    return key, found[key]


# ============================================================================ regions + hints
def segment(objs, g, gap_cols, gap_rows):
    par = list(range(len(objs)))

    def find(i):
        while par[i] != i:
            par[i] = par[par[i]]; i = par[i]
        return i

    def join(i, j):
        par[find(i)] = find(j)

    gc, gr = gap_cols * g.def_col, gap_rows * g.def_row
    idx = {o["id"]: i for i, o in enumerate(objs)}
    solo = ("picture", "chart", "object")
    for i, a in enumerate(objs):
        if a["kind"] in solo:
            continue
        for j in range(i + 1, len(objs)):
            b = objs[j]
            if b["kind"] in solo:
                continue
            if a.get("group") and b.get("group") and a["group"].split("/")[0] == b["group"].split("/")[0]:
                join(i, j); continue
            if inside(a["box"], b["box"]) or inside(b["box"], a["box"]):
                join(i, j); continue
            dx, dy = gap(a["box"], b["box"])
            tight = {"text", "callout"} & {a["kind"], b["kind"]} or (("cellbar" in (a["kind"], b["kind"])) and a["kind"] != b["kind"])
            lx, ly = (2 * PX, 2 * PX) if tight else (gc, gr)
            if dx <= lx and dy <= ly:            # near each other in both directions (lanes of one chart are
                join(i, j)                       # stacked with small offsets even when no arrow links them)
        if is_line(a):
            for gid in a["glue"]:
                if gid in idx:
                    join(i, idx[gid])
            for p in (a["p0"], a["p1"]):
                for j, b in enumerate(objs):
                    if j != i and b["kind"] not in solo and not is_line(b) and d_pt_box(p, b["box"]) <= 12 * PX:
                        join(i, j)
            t = label_for_line(a, [b for b in objs if b["kind"] == "text"], set())
            if t is not None:
                join(i, idx[t["id"]])
    groups = defaultdict(list)
    for i, o in enumerate(objs):
        groups[find(i)].append(o)
    regs = sorted(groups.values(), key=lambda m: (min(o["box"][1] for o in m), min(o["box"][0] for o in m)))
    out = []
    for k, m in enumerate(regs, 1):
        box = (min(o["box"][0] for o in m), min(o["box"][1] for o in m), max(o["box"][2] for o in m), max(o["box"][3] for o in m))
        out.append(dict(id=f"R{k}", objs=m, box=box))
    return out


def hints(reg, g):
    """Score how much a region looks like a timing chart, and what else it could be. Evidence, not verdicts."""
    o = reg["objs"]
    k = defaultdict(int)
    for x in o:
        k[x["kind"]] += 1
    nodes = [x for x in o if x["kind"] in ("shape", "callout", "cellbar")]
    lines = [x for x in o if is_line(x)]
    ends = attach_ends(o, 12 * PX)
    linked = [x for x in lines if all(e[0] for e in ends[x["id"]])]
    rows = {round(g.row_at(center(x["box"])[1]) - 0.5) for x in nodes}
    wide = [x for x in nodes if (x["box"][2] - x["box"][0]) >= 1.5 * (x["box"][3] - x["box"][1])]
    flow = [x for x in nodes if x["geom"] in FLOWISH]
    plain = [x for x in lines if x["kind"] == "line"]                 # waveforms are plain lines, not arrows
    hl = [x for x in plain if abs(x["p1"][1] - x["p0"][1]) < 0.2 * g.def_row and abs(x["p1"][0] - x["p0"][0]) > 0.3 * g.def_col]
    vl = [x for x in plain if abs(x["p1"][0] - x["p0"][0]) < 0.2 * g.def_col and abs(x["p1"][1] - x["p0"][1]) > 0.3 * g.def_row]
    harrows = [x for x in lines if x["kind"] == "arrow" and abs(x["p1"][1] - x["p0"][1]) < 0.2 * g.def_row]
    lifel = [x for x in vl if abs(x["p1"][1] - x["p0"][1]) > 3 * g.def_row]
    ax_x = find_axis(g, reg["box"], "x")
    lx, labs = side_labels(g, reg["box"], "x")
    feats = dict(counts=dict(k), linked_lines=len(linked), rows_used=len(rows), wide_bars=len(wide), flowchart_shapes=len(flow),
                 horizontal_lines=len(hl), vertical_lines=len(vl), time_axis=ax_x and {kk: ax_x[kk] for kk in ("kind", "line")},
                 labels_left=len(labs))
    H = []

    def add(name, score, why):
        if score > 0:
            H.append(dict(family=name, score=score, evidence=why))
    s, why = 0, []
    if ax_x: s += 3; why.append(f"{ax_x['kind']} time axis in row {ax_x['line']}")
    if labs: s += 1; why.append(f"{len(labs)} signal names in column {col_letter(lx)}")
    if len(rows) >= 2 and nodes and len(wide) >= 0.6 * len(nodes): s += 2; why.append(f"bars on {len(rows)} rows")
    if k.get("cellbar", 0) >= 2: s += 1; why.append(f"{k['cellbar']} runs of coloured cells")
    if len(hl) >= 3 and len(vl) + sum(1 for x in plain if x not in hl and x not in vl) >= 2 and len(lifel) < 2:
        s += 3; why.append(f"{len(hl)} level lines + edges (waveforms)")
    if any(x["geom"] in BUS_GEOMS for x in nodes): s += 1; why.append("bus-value shapes (hexagons etc.)")
    if flow and len(flow) >= 0.3 * max(len(nodes), 1): s -= 3; why.append(f"{len(flow)} flowchart shapes")
    add("timing chart", s, why)
    if len(lifel) >= 2 and len(harrows) >= 2:
        add("sequence diagram (time runs top to bottom; read messages in vertical order)", 6,
            [f"{len(lifel)} long vertical lifelines, {len(harrows)} horizontal message arrows"])
    if k.get("chart"):
        ch = [x["chart"] for x in o if x["kind"] == "chart"]
        add("native Excel chart (read its series)", 9, [f"{p['type']} {p.get('barDir', '')} {p.get('grouping', '')}".strip()
                                                        for c in ch for p in c.get("plots", [])])
    if k.get("picture"):
        add("picture (look at the image)", 8 if len(o) == k["picture"] else 1, [x["text"] or x["name"] for x in o if x["kind"] == "picture"])
    if not nodes and not lines and k.get("text"):
        add("note / free text", 6, [x["text"][:60] for x in o if x["kind"] == "text"])
    small = [x for x in nodes if (x["box"][2] - x["box"][0]) <= 1.2 * g.def_col and x["kind"] == "shape"]
    if nodes and not lines and len(small) == len(nodes) and 2 <= len(nodes) <= 12 and not ax_x:
        add("legend (colour key)", 7, ["small swatches, no lines, no axis"])
    if (flow or len(linked) >= 2) and not ax_x and not hl:
        add("other diagram (flowchart / block diagram), not a timing chart", 2 + len(flow) + min(len(linked), 4) // 2,
            [f"{len(linked)} lines join shapes, {len(flow)} flowchart shapes, no time axis"])
    cb = [x for x in nodes if x["kind"] == "cellbar"]
    if cb and len(cb) == len(nodes) and all(x["text"] for x in cb) and not ax_x:
        add("table formatting (filled cells holding text)", 4, ["every coloured run contains text"])
    H = [h for h in H if h["score"] >= 2] or [dict(family="unclear (inspect visually)", score=1, evidence=["no strong pattern"])]
    H.sort(key=lambda h: -h["score"])
    reg.update(features=feats, hints=H, type=H[0]["family"], axis_x=ax_x)
    if H[0]["family"].startswith("legend"):
        reg["legend"] = legend_entries(g, small, o)
    return reg


def legend_entries(g, sw, objs):
    out = []
    for s in sorted(sw, key=lambda b: (b["box"][1], b["box"][0])):
        txt = s["text"]
        if not txt:
            r = int(g.row_at(center(s["box"])[1])) + 1
            c = int(g.col_at(s["box"][2] - 1)) + 1
            for cc in range(c, c + 4):
                if isinstance(g.cells.get((r, cc)), str):
                    txt = g.cells[(r, cc)]; break
        if not txt:
            for t in objs:
                if t["kind"] == "text" and t["box"][0] >= s["box"][2] - 4 * PX and abs(center(t["box"])[1] - center(s["box"])[1]) < g.def_row:
                    txt = t["text"]; break
        out.append(dict(fill=s["fill"], line=s["line"], meaning=txt))
    return out


def describe(o, g):
    d = dict(id=o["id"], kind=o["kind"], geom=o["geom"], text=o["text"], cells=g.rng(o["box"]), style=style_key(o))
    for k in ("fill", "line", "dash", "width", "group", "rot", "media", "chart", "head"):
        if o.get(k) not in (None, "", 0, False):
            d[k] = o[k]
    if is_line(o):
        d.update(start=g.ref(*o["p0"]), end=g.ref(*o["p1"]), glued=list(o["glue"]), double=o.get("double", False))
    return d


# ============================================================================ loading
def load(bk, sheet, a, with_cells=None):
    with_cells = (not a.no_cells) if with_cells is None else with_cells
    sp = bk.sheet_path(sheet)
    g = bk.grid(sp)
    dp = bk.drawing(sp)
    objs = read_objects(bk, dp, g) if dp else []
    if with_cells:
        objs += cell_bars(g)
    if not getattr(a, "no_borders", False):
        objs += border_lines(g)
    regs = [hints(r, g) for r in segment(objs, g, a.gap_cols, a.gap_rows)] if objs else []
    return g, objs, regs


def select(g, objs, regs, a):
    if getattr(a, "box", None):
        p, qq = a.box.split(":")
        (r0, c0), (r1, c1) = parse_ref(p), parse_ref(qq)
        bx = g.cell_box(r0, c0, r1, c1)
        sel = [o for o in objs if inside(o["box"], bx, PX)]
    else:
        want = (a.region or "").split(",")
        chosen = [r for r in regs if r["id"] in want]
        if not chosen:
            sys.exit(f"Pass --region (one of {[r['id'] for r in regs]}) or --box. Run `scene` first.")
        sel = [o for r in chosen for o in r["objs"]]
    if not sel:
        sys.exit("Nothing selected.")
    box = (min(o["box"][0] for o in sel), min(o["box"][1] for o in sel), max(o["box"][2] for o in sel), max(o["box"][3] for o in sel))
    return sel, box


# ============================================================================ commands
def cmd_inventory(a):
    bk = Book(a.book)
    print(f"Workbook: {a.book}")
    for name, sp, state in bk.sheets:
        g = bk.grid(sp)
        dp = bk.drawing(sp)
        objs = read_objects(bk, dp, g) if dp else []
        k = defaultdict(int)
        for o in objs:
            k[o["kind"]] += 1
        used = ""
        if g.cells:
            rs = [r for r, _ in g.cells]; cs = [c for _, c in g.cells]
            used = f"{col_letter(min(cs))}{min(rs)}:{col_letter(max(cs))}{max(rs)}"
        filled = sum(1 for v in g.fill.values() if v and v != "#FFFFFF")
        extra = [f"{filled} coloured cells"] if filled else []
        extra += [f"{len(g.merges)} merged ranges"] if g.merges else []
        extra += [f"{len(g.dates)} date cells"] if g.dates else []
        extra += [f"{len(g.border)} cells with borders"] if g.border else []
        charts = [p["type"] for o in objs if o["kind"] == "chart" for p in o.get("chart", {}).get("plots", [])]
        print(f"- {name}{' [' + state + ']' if state != 'visible' else ''}: {len(g.cells)} cells ({used or 'empty'})"
              + (f", {', '.join(extra)}" if extra else "") + "; drawing: "
              + (", ".join(f"{v} {kk}" for kk, v in sorted(k.items())) or "none") + (f" (charts: {', '.join(charts)})" if charts else ""))


def cmd_scene(a):
    bk = Book(a.book)
    os.makedirs(a.out, exist_ok=True)
    sheets = [a.sheet] if a.sheet else [n for n, _, _ in bk.sheets]
    for s in sheets:
        g, objs, regs = load(bk, s, a)
        tag = re.sub(r"[^A-Za-z0-9]+", "_", s)
        if not objs:
            print(f"## {s}: nothing drawn (no shapes, pictures, charts or coloured cells)"); continue
        print(f"## {s}: {len(objs)} objects in {len(regs)} regions")
        js = []
        for r in regs:
            k = defaultdict(int)
            for o in r["objs"]:
                k[o["kind"]] += 1
            texts = [o["text"] for o in r["objs"] if o["text"]][:8]
            print(f"  {r['id']:<4} {g.rng(r['box']):<10} " + ", ".join(f"{v} {kk}" for kk, v in sorted(k.items())))
            for h in r["hints"][:3]:
                print(f"       {h['score']:>2}  {h['family']}  <- {'; '.join(map(str, h['evidence']))}")
            if texts:
                print(f"       text: {texts}")
            if r.get("legend"):
                print("       legend: " + ", ".join(f"{e['fill']}={e['meaning']}" for e in r["legend"]))
            for o in r["objs"]:
                if o["kind"] == "chart":
                    for p in o["chart"].get("plots", []):
                        print(f"       chart '{o['chart'].get('title', '')}': {p['type']} {p.get('barDir', '')} {p.get('grouping', '')}")
                        for sr in p["series"]:
                            print(f"         series '{sr['name']}'{' (invisible)' if sr['invisible'] else ''}: "
                                  f"{list(zip(sr['categories'], sr['values']))[:12]}  <- {sr['values_ref']}")
                if o["kind"] == "picture" and o.get("media") in bk.names:
                    dst = os.path.join(a.out, f"{tag}_{r['id']}_{os.path.basename(o['media'])}")
                    with open(dst, "wb") as f:
                        f.write(bk.z.read(o["media"]))
                    print(f"       picture saved: {dst}")
            styles = defaultdict(list)
            for o in r["objs"]:
                styles[style_key(o)].append(o["text"] or "")
            js.append(dict(id=r["id"], cells=g.rng(r["box"]), hints=r["hints"], features=r["features"],
                           axis_x=r["axis_x"], legend=r.get("legend"),
                           styles={k_: dict(count=len(v), examples=[t for t in v if t][:4]) for k_, v in styles.items()},
                           objects=[describe(o, g) for o in sorted(r["objs"], key=lambda o: (o["box"][1], o["box"][0]))]))
        with open(os.path.join(a.out, f"scene_{tag}.json"), "w", encoding="utf-8") as f:
            json.dump(dict(sheet=s, regions=js), f, indent=1, default=str)
        if not a.no_plot:
            p = overview_png(a.out, s, g, regs)
            if p:
                print(f"  overview: {p}")
    print(f"written to {os.path.abspath(a.out)}")


def cmd_cells(a):
    bk = Book(a.book)
    g = bk.grid(bk.sheet_path(a.sheet))
    if a.range:
        p, qq = a.range.split(":"); (r0, c0), (r1, c1) = parse_ref(p), parse_ref(qq)
    else:
        r0 = c0 = 0; r1 = c1 = 10 ** 9
    keys = set(g.cells) | (set(g.fill) if a.fills else set()) | (set(g.border) if a.borders else set())
    rows = defaultdict(list)
    for (r, c) in sorted(keys):
        if r0 <= r <= r1 and c0 <= c <= c1:
            v = g.show(r, c)
            f = g.fill.get((r, c)) if a.fills else None
            b = g.border.get((r, c)) if a.borders else None
            rows[r].append(f"{col_letter(c)}{r}=" + (repr(v) if v is not None else "") + (f"[{f}]" if f else "")
                           + (f"<{','.join(b)}>" if b else ""))
    for r in sorted(rows):
        print("  ".join(rows[r]))
    m = [f"{col_letter(c0_)}{r0_}:{col_letter(c1_)}{r1_}" for r0_, c0_, r1_, c1_ in g.merges]
    if m:
        print("merged: " + ", ".join(m))


# ============================================================================ pictures
def _plt():
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        return plt
    except ImportError:
        print("(matplotlib not installed: no PNG)")
        return None


def _hex(c, default="#eeeeee"):
    return c[:7] if c and c.startswith("#") and len(c) >= 7 else default


def overview_png(out, sheet, g, regs):
    plt = _plt()
    if not plt or not regs:
        return None
    from matplotlib.patches import Rectangle
    X1 = max(r["box"][2] for r in regs); Y1 = max(r["box"][3] for r in regs)
    nc, nr = int(g.col_at(X1)) + 2, int(g.row_at(Y1)) + 2
    fig, ax = plt.subplots(figsize=(min(4 + nc * 0.45, 24), min(2 + nr * 0.28, 18)))
    C, R = g.col_at, g.row_at
    for (r, c), v in g.cells.items():
        if r <= nr and c <= nc:
            ax.text(c + 0.05, r - 0.5, str(g.show(r, c))[:14], fontsize=6, color="#8a8a8a", va="center", zorder=1)
    pal = ["#1f77b4", "#d62728", "#2ca02c", "#9467bd", "#ff7f0e", "#8c564b", "#e377c2", "#17becf"]
    for i, reg in enumerate(regs):
        col = pal[i % len(pal)]
        for o in sorted(reg["objs"], key=lambda o: o["z"]):
            x0, y0, x1, y1 = o["box"]
            if is_line(o):
                (xs, ys), (xe, ye) = o["p0"], o["p1"]
                st = "<->" if o.get("double") else "->" if o["kind"] == "arrow" else "-"
                ax.annotate("", xy=(C(xe), R(ye)), xytext=(C(xs), R(ys)), zorder=4,
                            arrowprops=dict(arrowstyle=st, color=_hex(o["line"], "#222222"), lw=0.9,
                                            ls="--" if o["dash"] not in ("solid", "") else "-"))
                if o["text"]:
                    ax.text((C(xs) + C(xe)) / 2, (R(ys) + R(ye)) / 2, o["text"][:16], fontsize=6, zorder=5)
            else:
                fc = _hex(o["fill"], "#ffffff" if o["kind"] == "text" else "#eeeeee")
                ax.add_patch(Rectangle((C(x0), R(y0)), C(x1) - C(x0), R(y1) - R(y0), fc=fc, ec="#333333", zorder=2 + (0 if o["kind"] == "cellbar" else 1),
                                       lw=0.5, alpha=0.85, hatch="//" if o["kind"] in ("picture", "chart") else None,
                                       ls=":" if o["kind"] == "text" else "-"))
                if o["text"]:
                    ax.text((C(x0) + C(x1)) / 2, (R(y0) + R(y1)) / 2, o["text"][:20], fontsize=6, ha="center", va="center", zorder=5)
        x0, y0, x1, y1 = reg["box"]
        ax.add_patch(Rectangle((C(x0) - 0.12, R(y0) - 0.18), C(x1) - C(x0) + 0.24, R(y1) - R(y0) + 0.36, fill=False, ec=col, lw=1.6, ls="--", zorder=6))
        short = reg["type"].split(" (")[0]
        ax.text(C(x0) - 0.1, R(y0) - 0.28, f"{reg['id']}: {short}", color=col, fontsize=8, weight="bold", zorder=7)
    ax.set_xlim(0, nc); ax.set_ylim(nr, 0)
    ax.set_xticks([i + 0.5 for i in range(nc)]); ax.set_xticklabels([col_letter(i) for i in range(nc)], fontsize=6)
    ax.set_yticks([i + 0.5 for i in range(nr)]); ax.set_yticklabels([str(i + 1) for i in range(nr)], fontsize=6)
    ax.set_title(f"{sheet}: everything drawn, by region (grey = cell text)", fontsize=9, loc="left")
    fig.tight_layout()
    p = os.path.join(out, f"overview_{re.sub(r'[^A-Za-z0-9]+', '_', sheet)}.png")
    fig.savefig(p, dpi=130); plt.close(fig)
    return p


# ============================================================================ timing analysis
UP, DOWN = ("rise", "start"), ("fall", "end")


class AxisMap:
    """sheet x (EMU) -> time value"""

    def __init__(self, pts, desc, kind="number", cats=None, step=1.0):
        self.p, self.desc, self.kind, self.cats, self.step = sorted(pts), desc, kind, cats or [], step

    def __call__(self, v):
        p = self.p
        if len(p) == 1:
            return p[0][1]
        i = 0 if v <= p[0][0] else len(p) - 2 if v >= p[-1][0] else max(j for j in range(len(p) - 1) if p[j][0] <= v)
        (xa, ta), (xb, tb) = p[i], p[i + 1]
        return ta + (v - xa) * (tb - ta) / (xb - xa)

    def fmt(self, v, end=False):
        if v is None:
            return "-"
        if self.kind == "date":
            return serial_to_iso(v)
        if self.kind == "category" and self.cats:
            i = math.ceil(v - 1e-9) - 1 if end else math.floor(v + 1e-9)
            i = min(max(i, 0), len(self.cats) - 1)
            return f"{fnum(v)} ({'through ' if end else ''}{self.cats[i]})"
        return fnum(v)


def time_axis(g, box, a, xs):
    """Pick the time scale. Tick alignment 'auto' tries start/center/end of the tick cells and keeps the one
    that puts the most drawn edges on tick values."""
    if a.scale:
        v, o = a.scale.split("@")
        oc = col_index(o)
        return AxisMap([(g.cx(oc + k), k * float(v)) for k in range(600)], f"fixed scale {v} per column, 0 at column {o}",
                       step=float(v))
    ax = None
    if a.time_row:
        r = a.time_row
        ticks = [(c, _num(g, r, c)[0], g.show(r, c)) for c in range(0, 400) if _num(g, r, c)[0] is not None]
        dates = sum(1 for c, _, _ in ticks if (r, c) in g.dates) > len(ticks) / 2
        ax = dict(kind="date" if dates else "number", line=r, ticks=ticks) if len(ticks) >= 2 else None
    elif not a.no_ticks:
        ax = find_axis(g, box, "x")
    if not ax:
        return AxisMap([(g.cx(k), float(k + 1)) for k in range(600)],
                       "NO time axis found: time = column number (A=1). Give --time-row or --scale if one exists", step=1.0)
    vals = [v for _, v, _ in ax["ticks"]]
    step = min((b - x for x, b in zip(vals, vals[1:])), default=1.0) or 1.0
    best = None
    aligns = [a.tick_align] if a.tick_align != "auto" else (["start"] if ax["kind"] == "category" else ["start", "center", "end"])
    for al in aligns:
        pos = {"start": lambda c: g.cx(c), "center": lambda c: (g.cx(c) + g.cx(c + 1)) / 2, "end": lambda c: g.cx(c + 1)}[al]
        m = AxisMap([(pos(k), v) for k, v, _ in ax["ticks"]], "", ax["kind"], [lab for _, _, lab in ax["ticks"]], step)
        hit = sum(1 for x in xs if abs((m(x) - vals[0]) / step - round((m(x) - vals[0]) / step)) <= 0.06)
        if best is None or hit > best[0]:          # ties keep the earlier option ('start')
            best = (hit, al, m)
    hit, al, m = best
    labs = ", ".join(str(lab) for _, _, lab in ax["ticks"][:14]) + (" ..." if len(ax["ticks"]) > 14 else "")
    head = g.cells.get((ax["line"], ax["ticks"][0][0] - 1)) or g.cells.get((ax["line"], 0)) or ""
    m.desc = (f"{ax['kind']} ticks in row {ax['line']}: {labs}" + (f"  (header: '{head}')" if isinstance(head, str) and head else "")
              + f"; each tick marks the cell's {al}" + (f" (chosen automatically: {hit}/{len(xs)} drawn edges land on ticks)"
                                                        if a.tick_align == "auto" else ""))
    if ax["kind"] == "category":
        m.desc += "; values are category index (0 = first)"
    return m


def lane_bands(sel, g):
    """Lanes = horizontal bands of overlapping drawn material (bars, levels, edges). Long vertical lines are
    markers, arrows and free text are not lane material."""
    ivs, markers = [], []
    for o in sel:
        if o["kind"] in ("shape", "cellbar"):
            ivs.append([o["box"][1], o["box"][3], [o]])
        elif o["kind"] == "line":
            (x0, y0), (x1, y1) = o["p0"], o["p1"]
            if abs(y1 - y0) > 2.5 * g.def_row and abs(x1 - x0) < 0.3 * g.def_col:
                markers.append(o); continue
            ivs.append([min(y0, y1) - PX, max(y0, y1) + PX, [o]])
    ivs.sort(key=lambda v: v[0])
    bands = []
    for v in ivs:
        if bands and v[0] <= bands[-1][1]:
            bands[-1][1] = max(bands[-1][1], v[1]); bands[-1][2] += v[2]
        else:
            bands.append([v[0], v[1], list(v[2])])
    return bands, markers


def label_candidates(g, sel, box, label_col):
    """(y-centre, text, source) of possible lane names: label cells left of the chart, text boxes left of it."""
    c0 = int(g.col_at(box[0]))
    r0, r1 = int(g.row_at(box[1])) + 1, int(g.row_at(max(box[1], box[3] - 1))) + 1
    out = []
    if label_col:
        ci = col_index(label_col)
        cols = {ci}
    else:
        found = defaultdict(int)
        for (r, c), v in g.cells.items():
            if isinstance(v, str) and v.strip() and r0 - 1 <= r <= r1 + 1 and c0 - 8 <= c < c0:
                found[c] += 1
        cols = {max(found, key=lambda c: (found[c], c))} if found else set()
    for (r, c), v in g.cells.items():
        if c in cols and isinstance(v, str) and v.strip() and r0 - 1 <= r <= r1 + 1:
            mr = next((m for m in g.merges if m[0] <= r <= m[2] and m[1] <= c <= m[3]), None)
            y = (g.ry(mr[0] - 1) + g.ry(mr[2])) / 2 if mr else (g.ry(r - 1) + g.ry(r)) / 2
            out.append((y, v.strip(), f"cell {col_letter(c)}{r}"))
    left = min((o["box"][0] for o in sel if o["kind"] in ("shape", "cellbar", "line")), default=box[0])
    for o in sel:
        if o["kind"] == "text" and o["box"][2] <= left + 4 * PX:
            out.append((center(o["box"])[1], o["text"], f"text box {g.rng(o['box'])}"))
    return out


def build_lanes(sel, g, box, a):
    bands, markers = lane_bands(sel, g)
    labels = label_candidates(g, sel, box, a.label_col)
    # a band holding two or more labels is really several lanes drawn touching each other: split it
    out = []
    for b0, b1, objs in bands:
        inner = sorted([l for l in labels if b0 + 2 * PX < l[0] < b1 - 2 * PX], key=lambda l: l[0])
        if len(inner) >= 2:
            parts = defaultdict(list)
            for o in objs:
                cy = center(o["box"])[1]
                parts[min(range(len(inner)), key=lambda i: abs(inner[i][0] - cy))].append(o)
            for i, ob in sorted(parts.items()):
                out.append(dict(y0=min(o["box"][1] for o in ob), y1=max(o["box"][3] for o in ob), objs=ob, name=inner[i][1],
                                name_src=inner[i][2] + " (split a merged band)"))
        else:
            out.append(dict(y0=b0, y1=b1, objs=objs, name=None, name_src=None))
    used = set()
    for ln in sorted(out, key=lambda l: l["y1"] - l["y0"]):
        if ln["name"]:
            continue
        c = (ln["y0"] + ln["y1"]) / 2
        cands = [l for l in labels if ln["y0"] - 0.6 * g.def_row <= l[0] <= ln["y1"] + 0.6 * g.def_row and l[1] not in used]
        if cands:
            best = min(cands, key=lambda l: abs(l[0] - c))
            ln["name"], ln["name_src"] = best[1], best[2]; used.add(best[1])
    out.sort(key=lambda l: l["y0"])
    for i, ln in enumerate(out, 1):
        if not ln["name"]:
            ln["name"] = f"lane{i}@rows {int(g.row_at(ln['y0'])) + 1}-{int(g.row_at(max(ln['y0'], ln['y1'] - 1))) + 1}"
            ln["name_src"] = "UNNAMED (no label found)"
    return out, markers, labels


def lane_events(ln, g, amap, tol, ei):
    """Turn one lane's drawing into state spans and transition events."""
    h = ln["y1"] - ln["y0"]
    spans, edges, events, gaps = [], [], [], []
    for o in ln["objs"]:
        if o["kind"] == "line":
            (x0, y0), (x1, y1) = o["p0"], o["p1"]
            if abs(y1 - y0) <= 0.15 * g.def_row and abs(x1 - x0) > 0.1 * g.def_col:
                y = (y0 + y1) / 2
                if h > 0.4 * g.def_row:
                    rel = (y - ln["y0"]) / h
                    st = "high" if rel < 0.33 else "low" if rel > 0.67 else "mid (Z/idle)"
                else:
                    st = "level"
                spans.append(dict(t0=amap(min(x0, x1)), t1=amap(max(x0, x1)), state=st, label="", src=o["id"], kind="level",
                                  dash=o["dash"]))
            else:
                up = (y1 < y0) if x1 >= x0 else (y0 < y1)
                edges.append(dict(t0=amap(min(x0, x1)), t1=amap(max(x0, x1)), up=up, src=o["id"],
                                  vertical=abs(x1 - x0) < 0.1 * g.def_col))
        else:
            bus = o["geom"] in BUS_GEOMS
            st = "unknown (hatched)" if o["fill"] == "pattern" else ("value" if bus else "on")
            spans.append(dict(t0=amap(o["box"][0]), t1=amap(o["box"][2]), state=st, label=o["text"], src=o["id"],
                              kind="block", fill=o["fill"], geom=o["geom"]))
    lv = sorted([s for s in spans if s["kind"] == "level"], key=lambda s: s["t0"])
    bl = sorted([s for s in spans if s["kind"] == "block"], key=lambda s: s["t0"])
    hi_ = [s for s in lv if s["state"] == "high"]; lo_ = [s for s in lv if s["state"] == "low"]
    both = [(max(h["t0"], l["t0"]), min(h["t1"], l["t1"])) for h in hi_ for l in lo_ if min(h["t1"], l["t1"]) - max(h["t0"], l["t0"]) > tol]
    tot = lambda ss: sum(s_["t1"] - s_["t0"] for s_ in ss)
    if both and sum(b - x for x, b in both) >= 0.5 * min(tot(hi_), tot(lo_)):
        # a bus drawn as two parallel rails: each stretch where both rails exist is one value window
        win = []
        for x0_, x1_ in sorted(both):
            if win and x0_ <= win[-1][1] + tol:
                win[-1][1] = max(win[-1][1], x1_)
            else:
                win.append([x0_, x1_])
        lv = []
        for x0_, x1_ in win:
            spans.append(dict(t0=x0_, t1=x1_, state="value", label="", src=f"rails@{fnum(x0_)}", kind="block", fill="", geom="rails"))
        spans[:] = [s_ for s_ in spans if s_["kind"] == "block"]
        bl = sorted(spans, key=lambda s_: s_["t0"])
        for i in range(len(bl) - 1):                          # a crossing gap between windows is a value change
            if bl[i + 1]["t0"] - bl[i]["t1"] > tol:
                mid = (bl[i]["t1"] + bl[i + 1]["t0"]) / 2
                w = bl[i + 1]["t0"] - bl[i]["t1"]
                bl[i]["t1"], bl[i + 1]["t0"] = mid, mid
                bl[i]["xw"] = w
        edges = []

    def ev(t, kind, frm, to, src, width=0.0):
        ei[0] += 1
        e = dict(id=f"E{ei[0]}", t=t, lane=ln["name"], kind=kind, frm=frm, to=to, src=src, width=width)
        events.append(e)
        return e
    used_edges = set()
    for p, n in zip(lv, lv[1:]):
        lo, hi = p["t1"], n["t0"]
        near = [e for e in edges if e["t1"] >= lo - tol and e["t0"] <= hi + tol]
        if p["state"] == n["state"]:
            if hi - lo > tol and not near:
                gaps.append((lo, hi))
            continue
        if near:
            e = near[0]; used_edges.add(e["src"])
            t, w = (e["t0"] + e["t1"]) / 2, e["t1"] - e["t0"]
        else:
            t, w = (lo + hi) / 2, max(hi - lo, 0)
            if hi - lo > tol:
                gaps.append((lo, hi))
        kind = "rise" if (p["state"], n["state"]) == ("low", "high") else "fall" if (p["state"], n["state"]) == ("high", "low") else "change"
        ev(t, kind, p["state"], n["state"], p["src"] + ">" + n["src"], w)
    free = [e for e in edges if e["src"] not in used_edges]
    done = set()
    for e in free:
        if e["src"] in done:
            continue
        m = (e["t0"] + e["t1"]) / 2
        x_ = next((f for f in free if f is not e and f["src"] not in done and f["up"] != e["up"] and abs((f["t0"] + f["t1"]) / 2 - m) <= tol), None)
        if x_:                                                    # two crossing edges = bus value change (X)
            done |= {e["src"], x_["src"]}
            ev(m, "change", "?", "?", e["src"] + "+" + x_["src"], e["t1"] - e["t0"])
        else:
            done.add(e["src"])
            ev(m, "rise" if e["up"] else "fall", "?", "?", e["src"], e["t1"] - e["t0"])
    for b in bl:
        if b["geom"] == "rails" and not b["label"]:
            txt = [str(g.show(r, c)) for (r, c) in g.cells if ln["y0"] - 2 * PX <= center(g.cell_box(r, c))[1] <= ln["y1"] + 2 * PX
                   and b["t0"] - tol <= amap(center(g.cell_box(r, c))[0]) <= b["t1"] + tol]
            b["label"] = " ".join(txt)
    for i, b in enumerate(bl):
        prev = bl[i - 1] if i else None
        nxt = bl[i + 1] if i + 1 < len(bl) else None
        name = b["label"] or b["state"]
        if prev and abs(b["t0"] - prev["t1"]) <= tol:
            ev(b["t0"], "change", prev["label"] or prev["state"], name, prev["src"] + ">" + b["src"], prev.get("xw", 0.0))
        else:
            ev(b["t0"], "start", "idle", name, b["src"])
        if not (nxt and abs(nxt["t0"] - b["t1"]) <= tol):
            ev(b["t1"], "end", name, "idle", b["src"])
    vals = []
    xs_ = sorted(e["t"] for e in events if e["kind"] == "change")
    if xs_ or lv:
        cells_in = [(amap(center(g.cell_box(r, c))[0]), str(g.show(r, c))) for (r, c) in g.cells
                    if ln["y0"] - 2 * PX <= center(g.cell_box(r, c))[1] <= ln["y1"] + 2 * PX and isinstance(g.cells[(r, c)], (str, float))]
        t_lo = min((s["t0"] for s in lv), default=min(xs_, default=0)); t_hi = max((s["t1"] for s in lv), default=max(xs_, default=0))
        cells_in = [cv for cv in cells_in if t_lo - tol <= cv[0] <= t_hi + tol]
        bounds = [t_lo] + xs_ + [t_hi]
        for p, n in zip(bounds, bounds[1:]):
            inside_ = [v for t, v in cells_in if p - tol <= t <= n + tol]
            if inside_:
                vals.append((p, n, " ".join(inside_)))
        for e in events:
            if e["kind"] == "change" and e["frm"] == "?":
                before = next((v for p, n, v in vals if abs(n - e["t"]) <= tol), None)
                after = next((v for p, n, v in vals if abs(p - e["t"]) <= tol), None)
                e["frm"], e["to"] = before or "?", after or "?"
    ln["values"] = vals
    style = []
    if lv or edges:
        style.append("waveform lines")
    if bl:
        style.append(("bus (two rails)" if any(s_["geom"] == "rails" for s_ in bl) else "bus values") if any(s["state"] == "value" for s in bl) or
                     sum(1 for p, n in zip(bl, bl[1:]) if abs(n["t0"] - p["t1"]) <= tol) >= max(1, len(bl) // 2) else "bars")
    ln.update(spans=sorted(spans, key=lambda s: s["t0"]), events=sorted(events, key=lambda e: e["t"]), gaps=gaps,
              style=" + ".join(style) or "empty")
    return events


def nearest_text(g, sel, x, y, max_dx, max_dy):
    best = None
    for o in sel:
        if o["kind"] in ("text", "callout"):
            c = center(o["box"])
            if abs(c[0] - x) <= max_dx and abs(c[1] - y) <= max_dy:
                d = math.hypot(c[0] - x, c[1] - y)
                if best is None or d < best[0]:
                    best = (d, o["text"], o["id"])
    for (r, c), v in g.cells.items():
        if isinstance(v, str):
            cx_, cy_ = center(g.cell_box(r, c))
            if abs(cx_ - x) <= max_dx and abs(cy_ - y) <= max_dy:
                d = math.hypot(cx_ - x, cy_ - y)
                if best is None or d < best[0]:
                    best = (d, v, f"cell {col_letter(c)}{r}")
    return (best[1], best[2]) if best else (None, None)


def cmd_timing(a):
    bk = Book(a.book)
    g, objs, regs = load(bk, a.sheet, a)
    sel, box = select(g, objs, regs, a)
    xs = []
    for o in sel:
        if o["kind"] in ("shape", "cellbar"):
            xs += [o["box"][0], o["box"][2]]
        elif o["kind"] == "line":
            xs += [o["p0"][0], o["p1"][0]]
    amap = time_axis(g, box, a, xs)
    tol = a.tol if a.tol is not None else 0.03 * amap.step
    window = a.window if a.window is not None else 1.5 * amap.step
    snap_t = a.snap_time if a.snap_time is not None else 0.35 * amap.step
    lanes, markers, labels = build_lanes(sel, g, box, a)
    ei = [0]
    events = []
    for ln in lanes:
        events += lane_events(ln, g, amap, tol, ei)
    events.sort(key=lambda e: (e["t"], [l["name"] for l in lanes].index(e["lane"])))
    for k, e in enumerate(events, 1):
        e["n"] = k
    ebyid = {e["id"]: e for e in events}
    lane_of_obj = {o["id"]: l["name"] for l in lanes for o in l["objs"]}

    def band_of(y):
        inb = [l for l in lanes if l["y0"] - 0.5 * g.def_row <= y <= l["y1"] + 0.5 * g.def_row]
        return min(inb, key=lambda l: abs((l["y0"] + l["y1"]) / 2 - y))["name"] if inb else None

    # ---- markers (long vertical lines) with their labels
    M = []
    for m in markers:
        x = (m["p0"][0] + m["p1"][0]) / 2
        top, bot = min(m["p0"][1], m["p1"][1]), max(m["p0"][1], m["p1"][1])
        lab, _ = nearest_text(g, sel, x, top - 0.5 * g.def_row, 1.0 * g.def_col, 0.9 * g.def_row)
        if not lab:
            lab, _ = nearest_text(g, sel, x, bot + 0.5 * g.def_row, 1.0 * g.def_col, 0.9 * g.def_row)
        crossed = [l["name"] for l in lanes if l["y1"] >= top and l["y0"] <= bot]
        M.append(dict(id=m["id"], t=amap(x), label=lab, dash=m["dash"], spans=[crossed[0], crossed[-1]] if crossed else [None, None]))

    for m in M:
        m["coincides"] = [e["id"] for e in events if abs(e["t"] - m["t"]) <= tol]
    # ---- explicit arrows -> snapped to events
    texts = [o for o in sel if o["kind"] in ("text", "callout")]
    used_txt = set()
    ends = attach_ends(sel, a.snap * PX)

    def snap(pt, glued, anywhere=False):
        t = amap(pt[0])
        lane = band_of(pt[1])
        if glued and glued in lane_of_obj:
            lane = lane_of_obj[glued]
        c = [e for e in events if e["lane"] == lane] if lane else []
        if glued:
            cg = [e for e in c if glued in e["src"].split(">")]
            c = cg or c
        best = min(c, key=lambda e: abs(e["t"] - t), default=None)
        if best and abs(best["t"] - t) <= snap_t:
            return dict(event=best["id"], t=best["t"], drawn_t=t, lane=lane, off=t - best["t"])
        if anywhere or lane is None:          # dimension lines sit outside the lanes: match a transition on any lane
            best = min(events, key=lambda e: abs(e["t"] - t), default=None)
            if best and abs(best["t"] - t) <= snap_t:
                return dict(event=best["id"], t=best["t"], drawn_t=t, lane=best["lane"], off=t - best["t"])
        mk = min(M, key=lambda m: abs(m["t"] - t), default=None)
        if mk and abs(mk["t"] - t) <= snap_t:
            return dict(marker=mk["label"] or mk["id"], t=mk["t"], drawn_t=t, lane=lane, off=t - mk["t"])
        on = next((o for o in sel if o["kind"] in ("shape", "cellbar") and d_pt_box(pt, o["box"]) <= a.snap * PX), None)
        if on is not None:
            return dict(block=on["text"] or lane_of_obj.get(on["id"]), t=t, drawn_t=t, lane=lane_of_obj.get(on["id"], lane), off=0.0)
        return dict(event=None, t=t, drawn_t=t, lane=lane, off=0.0)

    links, measured = [], []
    for o in sel:
        if o["kind"] != "arrow":
            continue
        (gs, _), (gd, _) = ends[o["id"]]
        dbl = o.get("double", False)
        s, d = snap(o["p0"], o["glue"][0] or gs, dbl), snap(o["p1"], o["glue"][1] or gd, dbl)
        lab = o["text"]
        if not lab:
            t = label_for_line(o, texts, used_txt)
            if t:
                lab = t["text"]; used_txt.add(t["id"])
        rec = dict(id=o["id"], label=lab, src=s, dst=d, delta=d["t"] - s["t"], dash=o["dash"])
        (measured if o.get("double") else links).append(rec)

    # ---- clocks
    clocks = {}
    for ln in lanes:
        ups = sorted(e["t"] for e in ln["events"] if e["kind"] in UP)
        dns = sorted(e["t"] for e in ln["events"] if e["kind"] in DOWN)
        if len(ups) >= 3:
            per = [b - x for x, b in zip(ups, ups[1:])]
            mean = sum(per) / len(per)
            if mean > 0 and max(per) - min(per) <= 0.1 * mean:
                highs = [min((d for d in dns if d > u), default=None) for u in ups]
                highs = [hh - u for hh, u in zip(highs, ups) if hh is not None and hh - u < mean]
                clocks[ln["name"]] = dict(period=mean, high=(sum(highs) / len(highs)) if highs else None,
                                          first_rise=ups[0], edges=ups if a.clock_edge == "rise" else dns,
                                          irregular=max(per) - min(per) > tol)
    if clocks and amap.desc.startswith("NO time axis"):
        ck = next(iter(clocks.values()))
        for e in events:
            e["cycle"] = (e["t"] - ck["first_rise"]) / ck["period"]

    # ---- request/acknowledge handshakes (A up -> B up -> A down -> B down)
    hs, hmap = [], {}
    nc = [l for l in lanes if l["name"] not in clocks]
    for A in nc:
        for B in nc:
            if A is B:
                continue
            ea = [e for e in A["events"] if e["kind"] in UP + DOWN]
            eb = [e for e in B["events"] if e["kind"] in UP + DOWN]
            for u in [e for e in ea if e["kind"] in UP]:
                bu = next((e for e in eb if e["t"] > u["t"] + tol and e["kind"] in UP), None)
                ad = next((e for e in ea if bu and e["t"] > bu["t"] + tol and e["kind"] in DOWN), None)
                bd = next((e for e in eb if ad and e["t"] > ad["t"] + tol and e["kind"] in DOWN), None)
                if bu and ad and bd and not any(e["t"] < bu["t"] and e["kind"] in DOWN and e["t"] > u["t"] for e in ea):
                    hs.append(dict(request=A["name"], ack=B["name"], at=[u["t"], bu["t"], ad["t"], bd["t"]]))
                    for s_, d_ in ((u, bu), (bu, ad), (ad, bd)):
                        hmap.setdefault(d_["id"], (s_, f"step of a {A['name']}/{B['name']} request-acknowledge handshake"))
    # ---- inferred links for transitions no arrow explains
    explained = {l["dst"].get("event") for l in links if l["dst"].get("event")}
    pattern, kpattern = defaultdict(list), defaultdict(list)   # what the drawn arrows connect, and with what delay
    for l in links:
        se, de = ebyid.get(l["src"].get("event")), ebyid.get(l["dst"].get("event"))
        if se and de:
            pattern[(se["lane"], se["kind"], de["lane"], de["kind"])].append(l["delta"])
            kpattern[(se["kind"], de["kind"])].append(l["delta"])
    inferred, cause_of = [], defaultdict(set)
    for e in events:
        if e["id"] in explained or e["lane"] in clocks:
            continue
        cands = [s for s in events if s["lane"] != e["lane"] and e["t"] - window <= s["t"] <= e["t"] + tol]
        if e["id"] in hmap and hmap[e["id"]][0] not in cands:
            cands.append(hmap[e["id"]][0])                    # handshake steps may be far apart in time
        sugg = []
        for s in cands:
            dt = e["t"] - s["t"]
            if e["id"] in cause_of[s["id"]]:
                continue                                       # don't let two events explain each other
            if abs(dt) <= tol and e["kind"] in DOWN and s["kind"] not in DOWN:
                continue                                       # an end is not caused by something starting at the same instant
            near = lambda ds: any(abs(dt - d) <= max(tol, 0.1 * abs(d)) for d in ds)
            key = (s["lane"], s["kind"], e["lane"], e["kind"])
            if e["id"] in hmap and hmap[e["id"]][0]["id"] == s["id"]:
                sugg.append((0.5, "medium", hmap[e["id"]][1], s, dt))
            elif key in pattern and near(pattern[key]):
                sugg.append((0, "high", "same lanes, edges and delay as a drawn arrow", s, dt))
            elif (s["kind"], e["kind"]) in kpattern and near(kpattern[(s["kind"], e["kind"])]):
                sugg.append((1, "medium", f"same kind of hand-off as a drawn arrow ({s['kind']} -> {e['kind']}, delay {fnum(dt)})", s, dt))
            elif s["lane"] in clocks and any(abs(s["t"] - te) <= tol for te in clocks[s["lane"]]["edges"]):
                sugg.append((2, "medium", f"follows the active {a.clock_edge} edge of clock {s['lane']}", s, dt))
            elif abs(dt) <= tol:
                sugg.append((3, "medium", "happens at the same instant", s, dt))
            else:
                sugg.append((4, "low", "nearest earlier transition on another lane", s, dt))
        sugg.sort(key=lambda x: (x[0], abs(x[4])))
        prev = max((p_ for p_ in events if p_["lane"] == e["lane"] and p_["t"] < e["t"] - tol), key=lambda p_: p_["t"], default=None)
        if prev and (e["kind"] in DOWN or e["kind"] == "change") and (not sugg or sugg[0][0] >= 4):
            e["self"] = dict(after=prev["id"], width=e["t"] - prev["t"])    # simply the end of its own pulse / value
            continue
        if not sugg:
            continue
        for rank, conf, why, s, dt in [x for x in sugg if x[0] == sugg[0][0]][:2]:
            inferred.append(dict(src=s["id"], dst=e["id"], delta=dt, confidence=conf, reason=why))
            cause_of[e["id"]].add(s["id"])
    # repeated delays raise confidence
    rep = defaultdict(list)
    for i in inferred:
        s, d = ebyid[i["src"]], ebyid[i["dst"]]
        rep[(s["lane"], s["kind"], d["lane"], d["kind"], round(i["delta"] / max(tol, 1e-9)))].append(i)
    for grp in rep.values():
        if len(grp) >= 2:
            for i in grp:
                if i["confidence"] == "low":
                    i["confidence"] = "medium"
                i["reason"] += f" (same delay repeats {len(grp)} times)"

    has_in = explained | {i["dst"] for i in inferred}
    has_out = {l["src"].get("event") for l in links} | {i["src"] for i in inferred}
    unexplained = [e for e in events if e["lane"] not in clocks and e["id"] not in has_in and "self" not in e]
    for e in events:
        e["drives"] = sorted({l["dst"].get("event") for l in links if l["src"].get("event") == e["id"] and l["dst"].get("event")}
                             | {i["dst"] for i in inferred if i["src"] == e["id"]}, key=lambda k: int(k[1:]))

    # ---- handshakes, setup/hold
    sh = []
    for cn, ck in clocks.items():
        for ln in lanes:
            if ln["name"] == cn:
                continue
            setups, holds = [], []
            for te in ck["edges"]:
                b4 = [e["t"] for e in ln["events"] if te - ck["period"] < e["t"] < te - tol]
                af = [e["t"] for e in ln["events"] if te + tol < e["t"] < te + ck["period"]]
                if b4:
                    setups.append(te - max(b4))
                if af:
                    holds.append(min(af) - te)
            if setups or holds:
                sh.append(dict(clock=cn, lane=ln["name"], min_setup=min(setups) if setups else None, min_hold=min(holds) if holds else None,
                               same_instant=[e["t"] for e in ln["events"] if any(abs(e["t"] - te) <= tol for te in ck["edges"])]))

    # ---- notes (free text not used as label)
    label_ids = {l[2] for l in labels}
    notes = []
    for o in texts:
        if o["id"] in used_txt or f"text box {g.rng(o['box'])}" in label_ids or any(m["label"] == o["text"] for m in M):
            continue
        c = center(o["box"])
        notes.append(dict(text=o["text"], t=amap(c[0]), lane=band_of(c[1]), kind=o["kind"], cells=g.rng(o["box"])))

    # ---- what is missing / unclear
    miss = []
    if amap.desc.startswith("NO time axis"):
        miss.append("No time axis: times are column positions" + (f"; clock {next(iter(clocks))} gives a cycle count (see 'cycle')" if clocks else ""))
    for ln in lanes:
        if ln["name_src"].startswith("UNNAMED"):
            miss.append(f"Lane {ln['name']} has no name")
        for x, y in ln["gaps"]:
            miss.append(f"Lane {ln['name']}: nothing drawn between {amap.fmt(x)} and {amap.fmt(y)} (state unknown there)")
        bl = [s for s in ln["spans"] if s["kind"] == "block"]
        for p, n in zip(bl, bl[1:]):
            if n["t0"] < p["t1"] - tol:
                miss.append(f"Lane {ln['name']}: '{p['label']}' and '{n['label']}' overlap by {fnum(p['t1'] - n['t0'])}")
    for l in links + measured:
        nm = l["label"] or f"arrow {l['id']}"
        for end in ("src", "dst"):
            x = l[end]
            if x.get("block"):
                miss.append(f"{nm}: {'start' if end == 'src' else 'end'} points into the middle of '{x['block']}' at {amap.fmt(x['t'])}, not at a transition")
            elif not x.get("event") and not x.get("marker"):
                miss.append(f"{nm}: {'start' if end == 'src' else 'end'} at {amap.fmt(x['t'])} ({x['lane'] or 'outside lanes'}) touches no transition")
            elif abs(x["off"]) > tol:
                miss.append(f"{nm}: {'start' if end == 'src' else 'end'} drawn at {amap.fmt(x['drawn_t'])}, snapped to the transition at {amap.fmt(x['t'])}")
        if l in links and l["delta"] < -tol:
            miss.append(f"{nm} points backwards in time ({amap.fmt(l['src']['t'])} -> {amap.fmt(l['dst']['t'])})")
    noarrow = [e for e in events if e["id"] not in explained and e["lane"] not in clocks and "self" not in e]
    if noarrow:
        miss.append(f"{len(noarrow)} transition(s) have no drawn arrow explaining them; see 'inferred' / 'no cause found' in the list above")
    for e in unexplained:
        if e["id"] not in has_out:
            miss.append(f"{e['lane']} {e['kind']} at {amap.fmt(e['t'])}: no cause and no effect found (external input, a duration elapsing, or missing drawing)")
    for cn, ck in clocks.items():
        if ck["irregular"]:
            miss.append(f"Clock {cn} is drawn with uneven periods")

    res = dict(sheet=a.sheet, cells=g.rng(box), time_axis=amap.desc + (f"; unit: {a.unit}" if a.unit else ""), unit=a.unit, axis_kind=amap.kind, tolerance=tol, window=window,
               lanes=[dict(name=l["name"], name_source=l["name_src"], style=l["style"], values=l.get("values", []),
                           rows=f"{int(g.row_at(l['y0'] + 2 * PX)) + 1}-{int(g.row_at(max(l['y0'] + 2 * PX, l['y1'] - 3 * PX))) + 1}",
                           spans=[dict(t0=s["t0"], t1=s["t1"], state=s["state"], label=s["label"], kind=s["kind"]) for s in l["spans"]],
                           gaps=l["gaps"]) for l in lanes],
               events=events, arrows=links, measured=measured, markers=M, clocks=clocks, setup_hold=sh, handshakes=hs,
               inferred=inferred, unexplained=[e["id"] for e in unexplained], notes=notes, missing=miss,
               legend=[e_ for r in regs if r.get("legend") for e_ in r["legend"]])
    os.makedirs(a.out, exist_ok=True)
    tag = re.sub(r"[^A-Za-z0-9]+", "_", a.sheet) + "_" + re.sub(r"[^A-Za-z0-9]+", "_", a.region or a.box)
    with open(os.path.join(a.out, f"timing_{tag}.json"), "w", encoding="utf-8") as f:
        json.dump(res, f, indent=1, default=str)
    md = timing_md(res, amap, ebyid)
    with open(os.path.join(a.out, f"timing_{tag}.md"), "w", encoding="utf-8") as f:
        f.write(md)
    print(md)
    if not a.no_plot:
        p = timing_png(a.out, tag, res, ebyid)
        if p:
            print(f"re-drawn: {p}")


def timing_md(r, amap, E):
    F = amap.fmt
    ename = lambda e: f"{e['lane']} {e['kind']}" + (f" ({e['frm']}->{e['to']})" if e["kind"] == "change" or e["frm"] not in ("?", "idle", "low", "high") else
                                                   f" '{e['to']}'" if e["kind"] == "start" else "")

    def end(x):
        if x.get("event"):
            return f"[{ename(E[x['event']])} @ {F(x['t'])}]"
        if x.get("marker"):
            return f"[marker {x['marker']} @ {F(x['t'])}]"
        if x.get("block"):
            return f"[middle of '{x['block']}' on {x['lane']} @ {F(x['t'])}]"
        return f"[nothing at {F(x['t'])} on {x['lane'] or '?'}]"
    L = [f"# Timing chart: {r['sheet']} {r['cells']}", f"time axis: {r['time_axis']}"]
    if r["legend"]:
        L.append("legend: " + ", ".join(f"{e['fill']} = {e['meaning']}" for e in r["legend"]))
    L.append("\n## Lanes (top to bottom) and their states")
    for l in r["lanes"]:
        st = "; ".join(f"{s['label'] or s['state']} {F(s['t0'])}->{F(s['t1'], True)}" for s in l["spans"] if s["kind"] == "block")
        lv = "; ".join(f"{s['state']} {F(s['t0'])}->{F(s['t1'])}" for s in l["spans"] if s["kind"] == "level")
        L.append(f"- **{l['name']}** (rows {l['rows']}, {l['style']}, name from {l['name_source']}): " + " | ".join(x for x in (st, lv) if x))
    L.append("\n## Transitions in time order")
    expl = {a_["dst"].get("event"): a_ for a_ in r["arrows"] if a_["dst"].get("event")}
    inf = defaultdict(list)
    for i in r["inferred"]:
        inf[i["dst"]].append(i)
    for e in r["events"]:
        why = ""
        if e["id"] in expl:
            x = expl[e["id"]]
            why = f"<- ARROW from {end(x['src'])}" + (f" '{x['label']}'" if x["label"] else "") + f", delay {fnum(x['delta'])}"
        elif e["id"] in inf:
            why = "<- inferred: " + " or ".join(f"{ename(E[i['src']])} @ {F(E[i['src']]['t'])} (+{fnum(i['delta'])}, {i['confidence']}: {i['reason']})" for i in inf[e["id"]])
        elif e.get("self"):
            why = f"<- end of its own pulse/value: {fnum(e['self']['width'])} after #{E[e['self']['after']]['n']} (no other cause found)"
        elif e["lane"] in r["clocks"]:
            why = "(clock)"
        elif e["id"] in r["unexplained"]:
            why = "<- no cause found"
        if e.get("drives"):
            why += ("  " if why else "") + "-> drives " + ", ".join(f"#{E[k]['n']}" for k in e["drives"])
        cyc = f" [cycle {fnum(e['cycle'])}]" if "cycle" in e else ""
        wid = f" (transition width {fnum(e['width'])})" if e["width"] > r["tolerance"] else ""
        L.append(f"{e['n']}. t={F(e['t'])}{cyc}  {ename(e)}{wid}  {why}")
    if r["arrows"]:
        L.append("\n## Arrows drawn (cause -> effect)")
        for x in r["arrows"]:
            L.append(f"- {end(x['src'])} -> {end(x['dst'])}: delay {fnum(x['delta'])}" + (f", label '{x['label']}'" if x["label"] else "")
                     + (f", {x['dash']}" if x["dash"] != "solid" else ""))
    if r["measured"]:
        L.append("\n## Measured intervals (double-headed arrows)")
        for x in r["measured"]:
            L.append(f"- {x['label'] or '(unlabelled)'} = {fnum(abs(x['delta']))}: {end(x['src'])} <-> {end(x['dst'])}")
    if r["markers"]:
        L.append("\n## Marker lines")
        L += [f"- {m['label'] or '(unlabelled)'} at {F(m['t'])} ({m['dash']}), crosses {m['spans'][0]} .. {m['spans'][1]}"
              + (f"; same instant as {', '.join(ename(E[k]) + ' #' + str(E[k]['n']) for k in m['coincides'])}" if m["coincides"] else "") for m in r["markers"]]
    if r["clocks"]:
        L.append("\n## Clocks")
        for n, c in r["clocks"].items():
            L.append(f"- {n}: period {fnum(c['period'])}" + (f", high {fnum(c['high'])} (duty {fnum(100 * c['high'] / c['period'])}%)" if c["high"] else "")
                     + f", first rise {F(c['first_rise'])}")
    if r["setup_hold"]:
        L.append("\n## Relative to clock edges (setup = change before edge, hold = next change after edge)")
        L += [f"- {x['lane']} vs {x['clock']}: min setup {fnum(x['min_setup'])}, min hold {fnum(x['min_hold'])}"
              + (f"; changes exactly on an edge at {', '.join(F(t) for t in x['same_instant'])}" if x["same_instant"] else "") for x in r["setup_hold"]]
    if r["handshakes"]:
        L.append("\n## Request/acknowledge patterns (check against what the signals mean)")
        L += [f"- {h['request']} -> {h['ack']}: request up {F(h['at'][0])}, ack up {F(h['at'][1])}, request down {F(h['at'][2])}, ack down {F(h['at'][3])}" for h in r["handshakes"]]
    if r["notes"]:
        L.append("\n## Free text / callouts on the chart")
        L += [f"- '{n['text']}' at t~{F(n['t'])} ({n['lane'] or 'outside lanes'})" for n in r["notes"]]
    L.append("\n## Missing or unclear in the drawing")
    L += [f"- {m}" for m in r["missing"]] or ["- nothing found"]
    return "\n".join(L) + "\n"


def timing_png(out, tag, r, E):
    plt = _plt()
    if not plt or not r["lanes"]:
        return None
    from matplotlib.patches import Polygon
    names = [l["name"] for l in r["lanes"]]
    Y = {n: len(names) - 1 - i for i, n in enumerate(names)}
    fig, ax = plt.subplots(figsize=(12, 1.4 + 0.75 * len(names)))
    for l in r["lanes"]:
        y = Y[l["name"]]
        for s in l["spans"]:
            if s["kind"] == "level":
                yy = y + {"high": 0.25, "low": -0.25}.get(s["state"], 0)
                ax.plot([s["t0"], s["t1"]], [yy, yy], color="#1f2d3d", lw=1.4, ls="--" if s["state"].startswith("unknown") else "-")
            else:
                t0, t1 = s["t0"], s["t1"]; d = min((t1 - t0) * 0.12, 1e9)
                pts = [(t0, y), (t0 + d, y + 0.25), (t1 - d, y + 0.25), (t1, y), (t1 - d, y - 0.25), (t0 + d, y - 0.25)] if s["state"] == "value" \
                    else [(t0, y - 0.25), (t0, y + 0.25), (t1, y + 0.25), (t1, y - 0.25)]
                ax.add_patch(Polygon(pts, closed=True, fc="#c9d8ea" if s["state"] != "unknown (hatched)" else "#dddddd", ec="#1f2d3d", lw=0.8,
                                     hatch="//" if s["state"].startswith("unknown") else None))
                ax.text((t0 + t1) / 2, y, (s["label"] or "")[:16], ha="center", va="center", fontsize=7)
    for e in r["events"]:
        if e["kind"] in ("rise", "fall"):
            y = Y[e["lane"]]; ax.plot([e["t"], e["t"]], [y - 0.25, y + 0.25], color="#1f2d3d", lw=1.4)

    def yo(x):
        return Y.get(x["lane"], 0) if x.get("lane") else 0
    for x in r["arrows"]:
        ax.annotate("", xy=(x["dst"]["t"], yo(x["dst"])), xytext=(x["src"]["t"], yo(x["src"])),
                    arrowprops=dict(arrowstyle="->", color="#000", lw=1.3))
    for x in r["measured"]:
        ax.annotate("", xy=(x["dst"]["t"], yo(x["dst"]) - 0.4), xytext=(x["src"]["t"], yo(x["src"]) - 0.4),
                    arrowprops=dict(arrowstyle="<->", color="#555", lw=1))
        ax.text((x["src"]["t"] + x["dst"]["t"]) / 2, yo(x["src"]) - 0.32, x["label"] or "", fontsize=7, ha="center", color="#555")
    for i in r["inferred"]:
        s, d = E[i["src"]], E[i["dst"]]
        ax.annotate("", xy=(d["t"], Y[d["lane"]]), xytext=(s["t"], Y[s["lane"]]),
                    arrowprops=dict(arrowstyle="->", color="#d9822b" if i["confidence"] != "low" else "#e8b37a", lw=1.1, ls="--"))
    for m in r["markers"]:
        ax.axvline(m["t"], color="#888", ls=":", lw=1)
        ax.text(m["t"], len(names) - 0.45, m["label"] or "", fontsize=7, ha="center", color="#555")
    ts = [s["t0"] for l in r["lanes"] for s in l["spans"]] + [s["t1"] for l in r["lanes"] for s in l["spans"]] + [e["t"] for e in r["events"]]
    pad = (max(ts) - min(ts)) * 0.03 or 1
    ax.set_xlim(min(ts) - pad, max(ts) + pad); ax.set_ylim(-0.8, len(names) - 0.2)
    ax.set_yticks([Y[n] for n in names]); ax.set_yticklabels([n[:22] for n in names], fontsize=8)
    ax.grid(axis="x", color="#eee")
    ax.set_title(f"{r['sheet']} {r['cells']}: re-drawn from extracted data (black = drawn arrows, orange dashed = inferred)",
                 fontsize=9, loc="left")
    fig.tight_layout()
    p = os.path.join(out, f"timing_{tag}.png"); fig.savefig(p, dpi=140); plt.close(fig)
    return p


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    for n in ("inventory", "scene", "cells", "timing"):
        p = sub.add_parser(n)
        p.add_argument("book")
        p.add_argument("--sheet", required=n in ("cells", "timing"))
        p.add_argument("--out", default="timing_out")
        p.add_argument("--no-plot", action="store_true")
        p.add_argument("--no-cells", action="store_true", help="ignore coloured cells")
        p.add_argument("--no-borders", action="store_true", help="ignore cell borders")
        p.add_argument("--gap-cols", type=float, default=1.0, help="max gap (columns) that still joins shapes into one region")
        p.add_argument("--gap-rows", type=float, default=2.0, help="max gap (rows) that still joins shapes into one region")
        p.add_argument("--snap", type=float, default=12, help="px: a line end this close to a shape touches it")
        if n == "cells":
            p.add_argument("--range"); p.add_argument("--fills", action="store_true"); p.add_argument("--borders", action="store_true")
        if n == "timing":
            p.add_argument("--region", help="region id(s) from scene, e.g. R2 or R2,R3")
            p.add_argument("--box", help="use every object inside this cell range, e.g. C3:P20")
            p.add_argument("--time-row", type=int, help="row holding the time ticks")
            p.add_argument("--no-ticks", action="store_true", help="ignore tick rows; time = column number")
            p.add_argument("--scale", help="no ticks: '10@C' = 10 time units per column, 0 at column C")
            p.add_argument("--tick-align", choices=["auto", "start", "center", "end"], default="auto")
            p.add_argument("--label-col", help="column holding the signal names")
            p.add_argument("--unit", default="", help="time unit to print, e.g. ns (read it from the axis header)")
            p.add_argument("--clock-edge", choices=["rise", "fall"], default="rise")
            p.add_argument("--tol", type=float, help="time tolerance (default 3%% of a tick step)")
            p.add_argument("--snap-time", type=float, help="how far an arrow end may be from a transition and still attach (default 0.35 tick step)")
            p.add_argument("--window", type=float, help="how far back to look for a cause of an unexplained transition (default 1.5 tick steps)")
    a = ap.parse_args()
    {"inventory": cmd_inventory, "scene": cmd_scene, "cells": cmd_cells, "timing": cmd_timing}[a.cmd](a)


if __name__ == "__main__":
    main()
