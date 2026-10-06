import os
import glob
from lxml import etree
import latex2mathml.converter as l2m
from docx.oxml import parse_xml

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

def w(tag):
    return f"{{{W_NS}}}{tag}"

def m(tag):
    return f"{{{M_NS}}}{tag}"

def load_xsl():
    """Finds MML2OMML.XSL from local Microsoft Office installation."""
    candidates = []
    for base in (r"C:\Program Files\Microsoft Office", r"C:\Program Files (x86)\Microsoft Office"):
        for ver in ("root\\Office16", "Office16", "root\\Office15", "Office15", "Office14"):
            candidates.append(os.path.join(base, ver, "MML2OMML.XSL"))
    
    for c in candidates:
        if os.path.isfile(c):
            return etree.XSLT(etree.parse(c))
            
    found = glob.glob(r"C:\Program Files*\Microsoft Office\**\MML2OMML.XSL", recursive=True)
    if found:
        return etree.XSLT(etree.parse(found[0]))
    return None

class OmmlConverter:
    def __init__(self):
        self.xsl = load_xsl()

    def latex_to_omml(self, latex_str, display=False):
        if not self.xsl:
            raise RuntimeError("MML2OMML.XSL not found on this system.")
        
        mml = l2m.convert(latex_str, display="block" if display else "inline")
        root = self.xsl(etree.fromstring(mml.encode("utf-8"))).getroot()
        
        if root is None:
            raise RuntimeError("XSLT conversion returned empty.")
            
        if etree.QName(root).localname == "oMathPara":
            inner = root.find(m("oMath"))
            root = inner if inner is not None else root
            
        return parse_xml(etree.tostring(root))

def apply_vert_align(run, align_val):
    """Applies subscript or superscript to a docx run element."""
    rpr = run.find(w("rPr"))
    if rpr is None:
        rpr = run.makeelement(w("rPr"))
        run.insert(0, rpr)
    
    for old in rpr.findall(w("vertAlign")):
        rpr.remove(old)
        
    el = rpr.makeelement(w("vertAlign"))
    el.set(w("val"), align_val)
    rpr.append(el)
