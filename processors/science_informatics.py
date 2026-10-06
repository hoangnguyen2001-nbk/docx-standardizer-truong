import re
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CODE_LINE_RE = re.compile(
    r"^\s*(def |class |for .+:|while .+:|if .+:|elif |else\s*:|try:|except|print\(|input\(|import |from \w+ import|"
    r"return\b|#include|int main|using namespace|cin\s*>>|cout\s*<<|[{}]\s*$|\w+(\[.*\])?\s*[+\-*/]?=\s*[^=])")

def process(doc):
    for p in doc.paragraphs:
        text = p.text
        if CODE_LINE_RE.match(text):
            _apply_code_style(p)

def _apply_code_style(p):
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(0)
    
    # Shading
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F4F4F5')
    pPr.append(shd)
    
    # Font
    for run in p.runs:
        run.font.name = "Consolas"
        run.font.size = Pt(10)
