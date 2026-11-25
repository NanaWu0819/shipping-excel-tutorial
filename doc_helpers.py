import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

# Professional Color Palette
NAVY_HEX = "1B365D"       # #1B365D Primary Marine Blue
TEAL_HEX = "008080"       # #008080 Secondary Teal
LIGHT_BG_HEX = "F8FAFC"   # #F8FAFC Zebra row / Callout background
BORDER_HEX = "CBD5E1"     # #CBD5E1 Clean Light Gray
ALERT_BG_HEX = "FEF2F2"   # #FEF2F2 Red Callout
ALERT_BORDER = "EF4444"   # #EF4444
TIP_BG_HEX = "EFF6FF"     # #EFF6FF Blue Callout
TIP_BORDER = "3B82F6"     # #3B82F6
SUCCESS_BG = "F0FDF4"     # #F0FDF4 Green Callout
SUCCESS_BORDER = "22C55E" # #22C55E

COLOR_NAVY = RGBColor(27, 54, 93)
COLOR_TEAL = RGBColor(0, 128, 128)
COLOR_BODY = RGBColor(38, 38, 38)
COLOR_MUTED = RGBColor(100, 116, 139)

def setup_document_styles(doc):
    # Set standard margins (1 inch / 2.54 cm)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.different_first_page_header_footer = True
        
        # Header setup
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("航運行政人員必修：MS Excel 基礎資料分析前處理教學手冊  |  實務操作手冊")
        hrun.font.name = "微軟正黑體"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = COLOR_MUTED

        # Footer setup
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        frun1 = fp.add_run("內部教育訓練教材  |  第 ")
        frun1.font.name = "微軟正黑體"
        frun1.font.size = Pt(8.5)
        frun1.font.color.rgb = COLOR_MUTED
        
        # Dynamic Page Field
        fld_page = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        fp._p.append(fld_page)
        
        frun2 = fp.add_run(" 頁，共 ")
        frun2.font.name = "微軟正黑體"
        frun2.font.size = Pt(8.5)
        frun2.font.color.rgb = COLOR_MUTED
        
        # Dynamic Numpages Field
        fld_numpages = parse_xml(r'<w:fldSimple %s w:instr="NUMPAGES"/>' % nsdecls('w'))
        fp._p.append(fld_numpages)
        
        frun3 = fp.add_run(" 頁")
        frun3.font.name = "微軟正黑體"
        frun3.font.size = Pt(8.5)
        frun3.font.color.rgb = COLOR_MUTED

    # Base Normal Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = '微軟正黑體'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = COLOR_BODY
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(4)

def add_chapter_heading(doc, title, level=1):
    """Adds heading. For level 1, ensures it starts on a new page (Page Break)."""
    if level == 1:
        doc.add_page_break()
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(14)
        run = h.add_run(title)
        run.font.name = "微軟正黑體"
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.color.rgb = COLOR_NAVY
        # Add a colored bottom accent bar
        pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="18" w:space="4" w:color="1B365D"/></w:pBdr>' % nsdecls('w'))
        h._p.get_or_add_pPr().append(pBdr)
        return h
    elif level == 2:
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(6)
        run = h.add_run(title)
        run.font.name = "微軟正黑體"
        run.font.size = Pt(13.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_TEAL
        return h
    elif level == 3:
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(8)
        h.paragraph_format.space_after = Pt(3)
        run = h.add_run(title)
        run.font.name = "微軟正黑體"
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = COLOR_NAVY
        return h

def add_body_p(doc, text, bold_prefix=None, italic=False, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(space_after)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "微軟正黑體"
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_NAVY
    r = p.add_run(text)
    r.font.name = "微軟正黑體"
    r.font.size = Pt(10.5)
    r.font.italic = italic
    r.font.color.rgb = COLOR_BODY
    return p

def add_bullet_p(doc, text, bold_title=None, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
    if bold_title:
        r_title = p.add_run(bold_title)
        r_title.font.name = "微軟正黑體"
        r_title.font.bold = True
        r_title.font.color.rgb = COLOR_NAVY
    r = p.add_run(text)
    r.font.name = "微軟正黑體"
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_BODY
    return p

def add_callout(doc, title, text, callout_type="tip"):
    """
    Creates a styled executive callout box.
    callout_type: 'tip' (blue), 'warning' (red), 'success' (green), 'concept' (navy)
    """
    if callout_type == "warning":
        bg_hex = ALERT_BG_HEX
        border_hex = ALERT_BORDER
        icon = "[操作警示 / Caution]"
        tc = RGBColor(185, 28, 28)
    elif callout_type == "success":
        bg_hex = SUCCESS_BG
        border_hex = SUCCESS_BORDER
        icon = "[驗收達標 / Success]"
        tc = RGBColor(21, 128, 61)
    elif callout_type == "concept":
        bg_hex = LIGHT_BG_HEX
        border_hex = NAVY_HEX
        icon = "[核心思維 / Core Concept]"
        tc = COLOR_NAVY
    else: # tip
        bg_hex = TIP_BG_HEX
        border_hex = TIP_BORDER
        icon = "[航運行政實務小撇步 / Best Practice]"
        tc = RGBColor(30, 64, 175)

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    
    # XML Shading & Left Border
    shading_elm = parse_xml(r'<w:shd %s w:fill="%s"/>' % (nsdecls('w'), bg_hex))
    cell._tc.get_or_add_tcPr().append(shading_elm)
    
    tcBorders = parse_xml(
        r'<w:tcBorders %s>'
        r'<w:top w:val="none"/>'
        r'<w:left w:val="single" w:sz="36" w:space="0" w:color="%s"/>'
        r'<w:bottom w:val="none"/>'
        r'<w:right w:val="none"/>'
        r'</w:tcBorders>' % (nsdecls('w'), border_hex)
    )
    cell._tc.get_or_add_tcPr().append(tcBorders)

    # Padding
    tcMar = parse_xml(
        r'<w:tcMar %s>'
        r'<w:top w:w="120" w:type="dxa"/>'
        r'<w:left w:w="180" w:type="dxa"/>'
        r'<w:bottom w:w="120" w:type="dxa"/>'
        r'<w:right w:w="180" w:type="dxa"/>'
        r'</w:tcMar>' % nsdecls('w')
    )
    cell._tc.get_or_add_tcPr().append(tcMar)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.25
    
    r_icon = p.add_run(f"{icon} {title}\n")
    r_icon.font.name = "微軟正黑體"
    r_icon.font.bold = True
    r_icon.font.size = Pt(10.5)
    r_icon.font.color.rgb = tc
    
    r_text = p.add_run(text)
    r_text.font.name = "微軟正黑體"
    r_text.font.size = Pt(9.5)
    r_text.font.color.rgb = COLOR_BODY

    # Add spacing after table
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)

def add_image_figure(doc, img_path, caption, width_inch=6.0):
    """Inserts a figure image with centered alignment and caption below."""
    if not os.path.exists(img_path):
        print(f"Warning: image {img_path} not found.")
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(3)
    run_img = p_img.add_run()
    run_img.add_picture(img_path, width=Inches(width_inch))
    
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(0)
    p_cap.paragraph_format.space_after = Pt(10)
    r_cap = p_cap.add_run(caption)
    r_cap.font.name = "微軟正黑體"
    r_cap.font.size = Pt(9.0)
    r_cap.font.bold = True
    r_cap.font.color.rgb = COLOR_MUTED

def style_custom_table(table, col_widths, col_alignments):
    """Applies Navy header, light zebra striping, and clean padding to tables."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Header styling (Row 0)
    hdr_row = table.rows[0]
    for idx, cell in enumerate(hdr_row.cells):
        cell.width = Inches(col_widths[idx])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shd = parse_xml(r'<w:shd %s w:fill="%s"/>' % (nsdecls('w'), NAVY_HEX))
        cell._tc.get_or_add_tcPr().append(shd)
        
        tcMar = parse_xml(r'<w:tcMar %s><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>' % nsdecls('w'))
        cell._tc.get_or_add_tcPr().append(tcMar)
        
        tcBorders = parse_xml(r'<w:tcBorders %s><w:bottom w:val="single" w:sz="14" w:space="0" w:color="1B365D"/><w:top w:val="single" w:sz="14" w:space="0" w:color="1B365D"/></w:tcBorders>' % nsdecls('w'))
        cell._tc.get_or_add_tcPr().append(tcBorders)
        
        p = cell.paragraphs[0]
        p.alignment = col_alignments[idx]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        for r in p.runs:
            r.font.name = "微軟正黑體"
            r.font.bold = True
            r.font.size = Pt(9.5)
            r.font.color.rgb = RGBColor(255, 255, 255)

    # Data rows styling
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg_color = LIGHT_BG_HEX if (r_idx % 2 == 0) else "FFFFFF"
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(col_widths[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            shd = parse_xml(r'<w:shd %s w:fill="%s"/>' % (nsdecls('w'), bg_color))
            cell._tc.get_or_add_tcPr().append(shd)
            
            tcMar = parse_xml(r'<w:tcMar %s><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>' % nsdecls('w'))
            cell._tc.get_or_add_tcPr().append(tcMar)
            
            tcBorders = parse_xml(r'<w:tcBorders %s><w:bottom w:val="single" w:sz="4" w:space="0" w:color="%s"/><w:top w:val="single" w:sz="4" w:space="0" w:color="%s"/></w:tcBorders>' % (nsdecls('w'), BORDER_HEX, BORDER_HEX))
            cell._tc.get_or_add_tcPr().append(tcBorders)

            p = cell.paragraphs[0]
            p.alignment = col_alignments[c_idx]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "微軟正黑體"
                r.font.size = Pt(8.8)
                r.font.color.rgb = COLOR_BODY

print("doc_helpers.py created successfully.")
