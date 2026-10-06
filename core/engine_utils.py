import re
import copy
from dataclasses import dataclass
from docx.oxml import parse_xml
from lxml import etree
import latex2mathml.converter as l2m
import glob
import os

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"

def w(tag): return f"{{{W_NS}}}{tag}"
def m(tag): return f"{{{M_NS}}}{tag}"

def load_xsl():
    candidates = []
    for base in (r"C:\Program Files\Microsoft Office", r"C:\Program Files (x86)\Microsoft Office"):
        for ver in ("root\\Office16", "Office16", "root\\Office15", "Office15", "Office14"):
            candidates.append(os.path.join(base, ver, "MML2OMML.XSL"))
    for c in candidates:
        if os.path.isfile(c):
            return etree.XSLT(etree.parse(c))
    found = glob.glob(r"C:\Program Files*\Microsoft Office\**\MML2OMML.XSL", recursive=True)
    if found: return etree.XSLT(etree.parse(found[0]))
    return None

class OmmlConverter:
    def __init__(self):
        self.xsl = load_xsl()

    def to_omml(self, latex, display=False):
        if not self.xsl: raise RuntimeError("MML2OMML.XSL not found")
        # simple prep
        s = latex.strip().replace(r"\dfrac", r"\frac").replace(r"\tfrac", r"\frac")
        s = s.replace(r"^{\circ}", "°").replace(r"^\circ", "°").replace(r"\degree", "°")
        s = s.replace(r"\%", "%")
        for env in ("aligned", "align*", "align", "split", "eqnarray*", "eqnarray"):
            s = s.replace(f"\\begin{{{env}}}", r"\begin{array}{rl}").replace(f"\\end{{{env}}}", r"\end{array}")
        
        mml = l2m.convert(s, display="block" if display else "inline")
        root = self.xsl(etree.fromstring(mml.encode("utf-8"))).getroot()
        if etree.QName(root).localname == "oMathPara":
            inner = root.find(m("oMath"))
            root = inner if inner is not None else root
        return parse_xml(etree.tostring(root))

SUPS, SUBS, MATH = "superscript", "subscript", "math"

@dataclass
class Rule:
    regex: re.Pattern
    func: object

@dataclass
class Cell:
    ch: str
    src: int
    vert: str = None
    locked: bool = False
    math: object = None
    display: bool = False

def set_vert(run, vert):
    rpr = run.find(w("rPr"))
    if rpr is None:
        rpr = run.makeelement(w("rPr"), {})
        run.insert(0, rpr)
    for old in rpr.findall(w("vertAlign")): rpr.remove(old)
    el = rpr.makeelement(w("vertAlign"), {w("val"): vert})
    _RPR_AFTER_VERT = {w(t) for t in ("rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath")}
    for i, ch in enumerate(rpr):
        if ch.tag in _RPR_AFTER_VERT:
            rpr.insert(i, el)
            return
    rpr.append(el)

def new_run_like(src, text, vert):
    r = copy.deepcopy(src)
    for ch in list(r):
        if ch.tag != w("rPr"): r.remove(ch)
    if vert in (SUPS, SUBS): set_vert(r, vert)
    t = r.makeelement(w("t"), {})
    t.text = text
    t.set(f"{{{XML_NS}}}space", "preserve")
    r.append(t)
    return r

def run_vert(r):
    va = r.find(w("rPr") + "/" + w("vertAlign"))
    return va.get(w("val")) if va is not None else None

def is_simple_run(el):
    _SIMPLE = {w("rPr"), w("t"), w("lastRenderedPageBreak")}
    return el.tag == w("r") and all(ch.tag in _SIMPLE for ch in el)

class TextProcessor:
    def __init__(self, rules, conv):
        self.rules = rules
        self.conv = conv
        self.LATEX_RE = re.compile(r"\$\$(.+?)\$\$|\\\[(.+?)\\\]|\\\((.+?)\\\)|(?<![\\$\w])\$(?![\s$])(.+?)(?<![\s\\])\$(?![\w$])", re.S)
        self.do_math = any(r == 'MATH' for r in rules)

    def process_paragraph(self, p):
        for pe in p.findall(w("proofErr")): p.remove(pe)
        full = "".join(t.text or "" for t in p.iter(w("t")))
        if not full.strip(): return
        
        chunk = []
        for child in list(p):
            if is_simple_run(child):
                chunk.append(child)
            else:
                if chunk: self.process_chunk(p, chunk)
                chunk = []
        if chunk: self.process_chunk(p, chunk)

    def process_chunk(self, p, runs):
        cells = []
        for i, r in enumerate(runs):
            lock = run_vert(r) not in (None, "baseline")
            for t in r.findall(w("t")):
                for ch in (t.text or ""): cells.append(Cell(ch, i, None, lock))
        if not cells: return
        changed = False

        if self.do_math:
            text = "".join(c.ch for c in cells)
            if "$" in text or "\\(" in text or "\\[" in text:
                out, pos = [], 0
                for mt in self.LATEX_RE.finditer(text):
                    s, e = mt.span()
                    if any(c.locked for c in cells[s:e]): continue
                    latex = next(g for g in mt.groups() if g is not None)
                    display = mt.group(1) is not None or mt.group(2) is not None
                    try: om = self.conv.to_omml(latex, display)
                    except Exception as ex: 
                        print("LATEX EXCEPTION:", ex)
                        for c in cells[s:e]: c.locked = True
                        continue
                    out.extend(cells[pos:s])
                    out.append(Cell("\uFFFC", cells[s].src, None, True, om, display))
                    pos = e
                    changed = True
                out.extend(cells[pos:])
                cells = out

        for rule in [r for r in self.rules if r != 'MATH']:
            text = "".join(c.ch for c in cells)
            out, pos, hit = [], 0, False
            for mt in rule.regex.finditer(text):
                s, e = mt.span()
                if s == e or s < pos or any(c.locked for c in cells[s:e]): continue
                parts = rule.func(mt)
                if not parts: continue
                new_cells = []
                k = 0
                for t, v in parts:
                    for ch in t:
                        new_cells.append(Cell(ch, cells[min(s + k, e - 1)].src, v, True))
                        k += 1
                out.extend(cells[pos:s])
                out.extend(new_cells)
                pos, hit = e, True
            if hit:
                out.extend(cells[pos:])
                cells = out
                changed = True

        if changed:
            self.rebuild(p, runs, cells)

    def rebuild(self, p, runs, cells):
        maths = [c for c in cells if c.math is not None]
        others = "".join(c.ch for c in cells if c.math is None)
        wrap_para = (len(maths) == 1 and maths[0].display and not others.strip() and len(runs) == len([r for r in p if r.tag == w("r")]))
        new_elems = []
        i = 0
        while i < len(cells):
            c = cells[i]
            if c.math is not None:
                if wrap_para:
                    para = p.makeelement(m("oMathPara"), {})
                    para.append(c.math)
                    new_elems.append(para)
                else: new_elems.append(c.math)
                i += 1
                continue
            j, buf = i, []
            while j < len(cells) and cells[j].math is None and cells[j].src == c.src and cells[j].vert == c.vert:
                buf.append(cells[j].ch)
                j += 1
            seg = "".join(buf)
            if not (wrap_para and not seg.strip()):
                new_elems.append(new_run_like(runs[c.src], seg, c.vert))
            i = j
        idx = list(p).index(runs[0])
        for r in runs: p.remove(r)
        for k, el in enumerate(new_elems): p.insert(idx + k, el)
