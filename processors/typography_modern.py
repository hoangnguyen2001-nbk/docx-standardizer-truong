# -*- coding: utf-8 -*-
"""
CHẾ ĐỘ "TÀI LIỆU HỌC TẬP HIỆN ĐẠI" – Typography cho file Word.

Chạy như LỚP HOÀN THIỆN CUỐI CÙNG (sau mọi bộ xử lý khác) để thống nhất:
  - Phân cấp cỡ chữ theo Modular Scale (mặc định 1.25, nền 12pt).
  - Tiêu đề: đậm, giãn dòng hẹp (1.15); thân bài: giãn dòng 1.5.
  - Khoảng cách đoạn bằng space-after (không cần Enter 2 lần), bỏ thụt đầu dòng.
  - Độ dài dòng (measure) ~60–75 ký tự nhờ căn lề trang hai bên.
  - Màu chữ xám đậm #242424 thay vì đen thuần; nền trang trắng ngà (chỉ hiện trên màn hình).
  - Callout nền pastel + viền trái; Blockquote viền nhấn trái, chữ nghiêng, màu dịu.

Không chạm vào phương trình Word (m:oMath) – chỉ xử lý các run w:r trực tiếp.
"""
import re
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# ---------------------------------------------------------------------------
# Cấu hình
# ---------------------------------------------------------------------------
FONTS = {
    # khóa: (tên font thân bài, tên font tiêu đề)
    "inter":   ("Inter", "Inter"),
    "segoe":   ("Segoe UI", "Segoe UI Semibold"),
    "cambria": ("Cambria", "Cambria"),
    "georgia": ("Georgia", "Georgia"),
}

TEXT_COLOR = "242424"      # chữ thân bài
HEADING_COLOR = "1A1A1A"   # tiêu đề
ACCENT_COLOR = "1E40AF"    # nhấn (H1, viền quote)
MUTED_COLOR = "4B5563"     # quote, chú thích
PAGE_BG = "FDFCFA"         # nền trắng ngà (chỉ hiển thị trên màn hình)

BODY_LINE = 1.5
HEADING_LINE = 1.15

# Màu "trung tính" do các bộ xử lý trước gán – sẽ được đưa về TEXT_COLOR
NEUTRAL_COLORS = {None, "000000", "1E293B", "0F766E", "64748B", "B45309"}
# Màu nhấn cũ (nhãn "Câu X:", "A.", "B."...) – quy về một màu nhấn duy nhất
ACCENT_COLORS = {"1E3A8A", "1D4ED8", "1E40AF"}


def build_scale(base=12.0, ratio=1.25):
    """Trả về bảng cỡ chữ (pt) theo modular scale, làm tròn 0.5pt."""
    r = lambda n: round(base * (ratio ** n) * 2) / 2
    return {
        "title": r(4),    # 29.5 với 1.25
        "h1": r(3),       # 23.5
        "h2": r(2),       # 19
        "h3": r(1),       # 15
        "body": base,     # 12
        "small": r(-1),   # 9.5
    }


# ---------------------------------------------------------------------------
# Nhận diện vai trò đoạn văn
# ---------------------------------------------------------------------------
RE_TITLE = re.compile(r"^(MODULE|CHUYÊN ĐỀ|GIÁO ÁN|KẾ HOẠCH BÀI DẠY|ĐỀ (KIỂM TRA|THI|ÔN))\b", re.I)
RE_H1 = re.compile(r"^((I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+|PHẦN\s+(I|II|III|IV|V|\d+)\b)", re.I)
RE_H2 = re.compile(r"^(\d+\.\s+[A-ZÀ-Ỹ]|Dạng\s+\d+[:.])")
RE_H3 = re.compile(r"^(\d+\.\d+\.?\s+|[a-zđ]\)\s+[A-ZÀ-Ỹ][^.]{0,60}$)")
RE_QUESTION = re.compile(r"^(Câu|Bài)\s+[A-Za-z0-9_]+.{0,40}?[:.]", re.I)
RE_OPTION = re.compile(r"^([A-D][.)]|[a-d]\))\s+")
RE_BULLET1 = re.compile(r"^[–•\-]\s+")
RE_BULLET2 = re.compile(r"^\+\s+")
RE_CAPTION = re.compile(r"^(Hình|Sơ đồ|Biểu đồ|Đồ thị|Bảng)\s+\d+", re.I)
RE_CALLOUT = re.compile(r"^(\[LƯU Ý\]|LƯU Ý|Chú ý|Ghi nhớ|Mẹo|Lưu ý)(?=[\s:.]|$)|BẪY VÀ LỖI SAI", re.I)
RE_QUOTE = re.compile(r"^[“\"].{20,}[”\"]\s*(\(.+\))?$|^>\s+")
RE_EXPLAIN = re.compile(r"^(Giải thích|Phân tích phương án nhiễu|Lời giải|Đáp án|Đáp số|Hướng dẫn)[:.]", re.I)


def classify(p):
    text = p.text.strip()
    style = (p.style.name or "").lower() if p.style is not None else ""
    if not text:
        return "empty"
    if style == "title":
        return "title"
    if style.startswith("heading 1"):
        return "h1"
    if style.startswith("heading 2"):
        return "h2"
    if style.startswith("heading 3") or style.startswith("heading 4"):
        return "h3"
    if style in ("quote", "intense quote"):
        return "quote"
    if RE_TITLE.match(text) and len(text) < 160:
        return "title"
    if RE_H1.match(text) and len(text) < 160:
        return "h1"
    if RE_H2.match(text) and len(text) < 140:
        return "h2"
    if RE_H3.match(text):
        return "h3"
    if RE_CALLOUT.search(text):
        return "callout"
    if RE_QUOTE.match(text):
        return "quote"
    if RE_CAPTION.match(text) and len(text) < 200:
        return "caption"
    if RE_QUESTION.match(text):
        return "question"
    if RE_OPTION.match(text):
        return "option"
    if RE_EXPLAIN.match(text):
        return "explain"
    if RE_BULLET2.match(text):
        return "bullet2"
    if RE_BULLET1.match(text) or "list" in style:
        return "bullet1"
    return "body"


# ---------------------------------------------------------------------------
# Tiện ích OpenXML
# ---------------------------------------------------------------------------
def _set_run_font(run, name):
    run.font.name = name
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), name)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def _run_color(run):
    c = run.font.color
    try:
        return str(c.rgb) if c is not None and c.type is not None and c.rgb is not None else None
    except Exception:
        return None


def _set_color(run, hex_color):
    run.font.color.rgb = RGBColor.from_string(hex_color)


def _remove_pPr_child(p, tag):
    pPr = p._p.get_or_add_pPr()
    for el in pPr.findall(qn(tag)):
        pPr.remove(el)
    return pPr


def _set_shading(p, fill):
    pPr = _remove_pPr_child(p, "w:shd")
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def _set_left_border(p, color, size=18, space=8):
    pPr = _remove_pPr_child(p, "w:pBdr")
    bdr = parse_xml(
        '<w:pBdr %s><w:left w:val="single" w:sz="%d" w:space="%d" w:color="%s"/></w:pBdr>'
        % (nsdecls("w"), size, space, color)
    )
    # pBdr phải đứng trước shd trong pPr theo schema
    shd = pPr.find(qn("w:shd"))
    if shd is not None:
        shd.addprevious(bdr)
    else:
        pPr.append(bdr)


def _fmt(p, *, size_before, size_after, line, align=None, left=None, hanging=None, keep_next=False):
    pf = p.paragraph_format
    pf.space_before = Pt(size_before)
    pf.space_after = Pt(size_after)
    pf.line_spacing = line
    pf.first_line_indent = None
    if hanging is not None:
        pf.first_line_indent = -hanging
    pf.left_indent = left
    if align is not None:
        p.alignment = align
    if keep_next:
        pf.keep_with_next = True


# ---------------------------------------------------------------------------
# Áp dụng
# ---------------------------------------------------------------------------
def _style_runs(p, font, size, *, bold=None, italic=None, color=None, force_color=False):
    for r in p.runs:
        _set_run_font(r, font)
        r.font.size = Pt(size)
        if bold is not None:
            r.font.bold = bold
        if italic is not None:
            r.font.italic = italic
        cur = _run_color(r)
        if force_color or cur in NEUTRAL_COLORS:
            _set_color(r, color or TEXT_COLOR)
        elif cur in ACCENT_COLORS:
            _set_color(r, ACCENT_COLOR)


def format_paragraph(p, role, S, body_font, head_font):
    B = S["body"]
    line_pt = B * BODY_LINE  # chiều cao 1 dòng thân bài (pt)

    if role == "title":
        _fmt(p, size_before=0, size_after=line_pt, line=HEADING_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
        _style_runs(p, head_font, S["title"], bold=True, color=HEADING_COLOR, force_color=True)
    elif role == "h1":
        _fmt(p, size_before=line_pt * 1.5, size_after=B * 0.6, line=HEADING_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
        _style_runs(p, head_font, S["h1"], bold=True, color=ACCENT_COLOR, force_color=True)
    elif role == "h2":
        _fmt(p, size_before=line_pt, size_after=B * 0.5, line=HEADING_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
        _style_runs(p, head_font, S["h2"], bold=True, color=HEADING_COLOR, force_color=True)
    elif role == "h3":
        _fmt(p, size_before=line_pt * 0.75, size_after=B * 0.4, line=HEADING_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
        _style_runs(p, head_font, S["h3"], bold=True, color=HEADING_COLOR, force_color=True)
    elif role == "question":
        _fmt(p, size_before=line_pt * 0.75, size_after=B * 0.35, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
        _style_runs(p, body_font, B)  # giữ phần "Câu X:" đậm do bộ xử lý đề thi đặt
    elif role == "option":
        _fmt(p, size_before=0, size_after=B * 0.25, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT,
             left=Mm(10), hanging=Mm(6), keep_next=False)
        _style_runs(p, body_font, B)
    elif role == "explain":
        _fmt(p, size_before=0, size_after=B * 0.5, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, left=Mm(6))
        _style_runs(p, body_font, B)
    elif role == "bullet1":
        _fmt(p, size_before=0, size_after=B * 0.35, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT,
             left=Mm(6.3), hanging=Mm(6.3))
        _style_runs(p, body_font, B)
    elif role == "bullet2":
        _fmt(p, size_before=0, size_after=B * 0.25, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT,
             left=Mm(12.7), hanging=Mm(6.3))
        _style_runs(p, body_font, B)
    elif role == "caption":
        _fmt(p, size_before=B * 0.25, size_after=line_pt, line=HEADING_LINE, align=WD_ALIGN_PARAGRAPH.CENTER)
        _style_runs(p, body_font, S["small"], italic=True, color=MUTED_COLOR, force_color=True)
    elif role == "callout":
        _fmt(p, size_before=B * 0.5, size_after=line_pt, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, left=Mm(3))
        _set_shading(p, "FEF7E6")
        _set_left_border(p, "D97706")
        _style_runs(p, body_font, B, color="7C2D12", force_color=True)
    elif role == "quote":
        _fmt(p, size_before=B * 0.5, size_after=line_pt, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT, left=Mm(6))
        _set_left_border(p, ACCENT_COLOR, size=18, space=10)
        _style_runs(p, body_font, B, italic=True, color=MUTED_COLOR, force_color=True)
    elif role == "body":
        # Khoảng cách đoạn = 1 dòng chữ thân bài; không thụt đầu dòng; căn trái (dễ đọc hơn căn đều trên màn hình)
        _fmt(p, size_before=0, size_after=line_pt, line=BODY_LINE, align=WD_ALIGN_PARAGRAPH.LEFT)
        _style_runs(p, body_font, B)


def _merge_list_spacing(paragraphs, roles, S):
    """Mục cuối của một nhóm danh sách/phương án cách đoạn sau = 1 dòng."""
    line_pt = S["body"] * BODY_LINE
    list_roles = {"bullet1", "bullet2", "option"}
    for i, role in enumerate(roles):
        if role in list_roles:
            nxt = roles[i + 1] if i + 1 < len(roles) else None
            if nxt not in list_roles:
                paragraphs[i].paragraph_format.space_after = Pt(line_pt * 0.75)


def _setup_page(doc):
    # Độ dài dòng ~65–75 ký tự: vùng chữ ~150mm trên khổ A4 với cỡ 12pt
    for s in doc.sections:
        s.page_width, s.page_height = Mm(210), Mm(297)
        s.left_margin = s.right_margin = Mm(30)
        s.top_margin, s.bottom_margin = Mm(25), Mm(25)


def _setup_background(doc):
    root = doc.element
    bg = root.find(qn("w:background"))
    if bg is None:
        bg = OxmlElement("w:background")
        root.insert(0, bg)
    bg.set(qn("w:color"), PAGE_BG)
    settings = doc.settings.element
    if settings.find(qn("w:displayBackgroundShape")) is None:
        el = OxmlElement("w:displayBackgroundShape")
        # theo schema, displayBackgroundShape nằm gần đầu settings
        settings.insert(0, el)


def _setup_default_styles(doc, S, body_font, head_font):
    normal = doc.styles["Normal"]
    normal.font.name = body_font
    normal.font.size = Pt(S["body"])
    normal.font.color.rgb = RGBColor.from_string(TEXT_COLOR)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rfonts.set(qn(attr), body_font)
    normal.paragraph_format.line_spacing = BODY_LINE
    normal.paragraph_format.space_after = Pt(S["body"] * BODY_LINE)

    mapping = {"Title": "title", "Heading 1": "h1", "Heading 2": "h2", "Heading 3": "h3"}
    for name, key in mapping.items():
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        st.font.name = head_font
        st.font.size = Pt(S[key])
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(ACCENT_COLOR if key == "h1" else HEADING_COLOR)
        st.paragraph_format.line_spacing = HEADING_LINE


def _format_tables(doc, S, body_font):
    size = S["small"] + 1 if S["small"] + 1 < S["body"] else S["body"]
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.first_line_indent = None
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.25
                    for r in p.runs:
                        _set_run_font(r, body_font)
                        r.font.size = Pt(size)
                        if _run_color(r) in NEUTRAL_COLORS:
                            _set_color(r, TEXT_COLOR)


def _format_footer(doc, body_font, S):
    for s in doc.sections:
        for p in s.footer.paragraphs:
            for r in p.runs:
                _set_run_font(r, body_font)
                r.font.size = Pt(S["small"])
                _set_color(r, "6B7280")


def apply(doc, font_key="inter", ratio=1.25, base=12.0, page_background=True):
    body_font, head_font = FONTS.get(font_key, FONTS["inter"])
    S = build_scale(base, ratio)

    _setup_page(doc)
    _setup_default_styles(doc, S, body_font, head_font)
    if page_background:
        _setup_background(doc)

    paragraphs = list(doc.paragraphs)
    for p in paragraphs:
        for r in list(p._p.findall(qn("w:r"))):
            if all(ch.tag in (qn("w:rPr"), qn("w:t")) for ch in r) and not "".join(t.text or "" for t in r.findall(qn("w:t"))):
                p._p.remove(r)
    roles = [classify(p) for p in paragraphs]
    for p, role in zip(paragraphs, roles):
        if role != "empty":
            format_paragraph(p, role, S, body_font, head_font)
    _merge_list_spacing(paragraphs, roles, S)

    # Bỏ các đoạn trống do người dùng gõ Enter 2 lần (đã có space-after thay thế),
    # trừ đoạn chứa ảnh/phương trình/ngắt trang.
    for p, role in zip(paragraphs, roles):
        if role == "empty":
            el = p._p
            if el.xpath(".//w:drawing|.//w:pict|.//m:oMath|.//w:br[@w:type='page']|.//w:sectPr"):
                continue
            el.getparent().remove(el)

    _format_tables(doc, S, body_font)
    _format_footer(doc, body_font, S)
    return S
