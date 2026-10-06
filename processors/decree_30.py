import re
from docx.shared import Mm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

EXCLUDE_FIRST_LINE_INDENT = re.compile(
    r"^(MODULE|CHUYÊN ĐỀ|GIÁO ÁN|PHẦN|[I|V|X]+\.|\d+\.|Câu|Bài|[A-D][\.\)]|[a-d]\)|–|\+|\*|•|Tên chuyên đề|Bài SGK|Thời lượng|Phân phối|Vị trí|Phạm vi|Tọa độ|Tính chất|Tác động|Địa hình|Sông ngòi|Thổ nhưỡng|Sinh vật|Dấu hiệu|Cách làm|Phương pháp|Giải thích|Phân tích|Lời giải|Đáp án|Đáp số|Quy chuẩn|\[LƯU Ý\]|CỘNG HÒA|Độc lập|Nơi nhận)",
    re.IGNORECASE
)

def process(doc):
    # 1. Page Margins (A4: Top=20mm, Bottom=20mm, Left=30mm, Right=15mm)
    for section in doc.sections:
        section.page_height = Mm(297)
        section.page_width = Mm(210)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)

    # 2. Typography Styles
    for style in doc.styles:
        if hasattr(style, 'font'):
            style.font.name = 'Times New Roman'
    
    # 3. Process paragraphs for Decree 30 standard
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        # Ensure all runs use Times New Roman
        for run in p.runs:
            run.font.name = 'Times New Roman'
            if not run.font.size:
                run.font.size = Pt(13)
        
        # Standard body narrative paragraph formatting
        if not EXCLUDE_FIRST_LINE_INDENT.match(text):
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Inches(0.5) # 1.27 cm
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.35
            
        # Quốc hiệu: "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
        if text.upper() == "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.bold = True
                
        # Tiêu ngữ: "Độc lập - Tự do - Hạnh phúc"
        elif "ĐỘC LẬP - TỰ DO - HẠNH PHÚC" in text.upper() or "ĐỘC LẬP – TỰ DO – HẠNH PHÚC" in text.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            for run in p.runs:
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.underline = True
        
        # Nơi nhận
        elif text.startswith("Nơi nhận:"):
            p.paragraph_format.first_line_indent = Pt(0)
            for run in p.runs:
                if "Nơi nhận:" in run.text:
                    run.font.size = Pt(12)
                    run.font.bold = True
                    run.font.italic = True
                    
    # Clean table borders for header tables
    for table in doc.tables:
        if len(table.columns) == 2 and len(table.rows) <= 3:
            _remove_table_borders(table)


def _remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is not None:
        tblBorders = tblPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders")
        if tblBorders is not None:
            for border in tblBorders:
                border.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "none")
