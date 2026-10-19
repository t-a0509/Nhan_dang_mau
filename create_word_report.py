"""
Script tạo báo cáo Word (.docx) hoàn chỉnh, chuẩn học thuật
cho Bài tập 1 - Môn Nhận dạng mẫu (Pattern Recognition) - Thạc sĩ.

Chiến lược ghép bìa:
  - Mở BÌA TIỂU LUẬN.docx làm tài liệu GỐC (giữ nguyên logo, font, định dạng).
  - Thêm page break sau bìa rồi nối toàn bộ nội dung báo cáo vào ngay sau.
  - Lưu đè ra Bao_cao_Bai_tap_1.docx.
"""

import os
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import docx.opc.constants

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
RESULTS_DIR = os.path.join(BASE_DIR, "results")
BIA_PATH    = os.path.join(BASE_DIR, "BÌA TIỂU LUẬN.docx")
DOCX_PATH   = os.path.join(BASE_DIR, "Bao_cao_Bai_tap_1.docx")


# ─────────────────────────────── UTILITY FUNCTIONS ────────────────────────────

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd  = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}" w:val="clear"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr  = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top    w:w="{top}"    w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left   w:w="{left}"   w:type="dxa"/>'
        f'<w:right  w:w="{right}"  w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def add_heading(doc, text, level):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.keep_with_next = True
    h.paragraph_format.space_before   = Pt(16 if level == 1 else 10)
    h.paragraph_format.space_after    = Pt(5)
    if h.runs:
        run = h.runs[0]
        run.font.name = "Times New Roman"
        if level == 1:
            run.font.size = Pt(15)
            run.font.bold = True
            run.font.color.rgb = RGBColor(31, 78, 121)
        elif level == 2:
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.color.rgb = RGBColor(46, 117, 182)
        elif level == 3:
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(50, 50, 50)
    return h

def add_p(doc, text="", bold_part=None, italic=False, space_after=6, indent=False):
    """Đoạn văn chuẩn học thuật, Times New Roman 12pt, căn đều hai bên."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before       = Pt(0)
    p.paragraph_format.space_after        = Pt(space_after)
    p.paragraph_format.line_spacing       = 1.35
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    if bold_part:
        rb = p.add_run(bold_part + " ")
        rb.font.name      = "Times New Roman"
        rb.font.size      = Pt(12)
        rb.font.bold      = True
        rb.font.color.rgb = RGBColor(20, 20, 20)
    if text:
        rt = p.add_run(text)
        rt.font.name      = "Times New Roman"
        rt.font.size      = Pt(12)
        rt.font.italic    = italic
        rt.font.color.rgb = RGBColor(30, 30, 30)
    return p

def add_bullet(doc, text, bold_prefix=None):
    """Bullet point chuẩn học thuật."""
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.line_spacing = 1.3
    if bold_prefix:
        rb = p.add_run(bold_prefix + " ")
        rb.font.name      = "Times New Roman"
        rb.font.size      = Pt(12)
        rb.font.bold      = True
        rb.font.color.rgb = RGBColor(20, 20, 20)
    rt = p.add_run(text)
    rt.font.name      = "Times New Roman"
    rt.font.size      = Pt(12)
    rt.font.color.rgb = RGBColor(30, 30, 30)
    return p

def add_callout(doc, lines, title=None):
    """Khung nổi bật nền xanh nhạt, viền trái đậm."""
    tbl  = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "EBF3FA")
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    tcPr    = cell._element.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top    w:val="none"/>'
        f'<w:left   w:val="single" w:sz="24" w:space="0" w:color="1F4E79"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right  w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.paragraph_format.line_spacing = 1.35
    if title:
        rb = p.add_run(title + "\n")
        rb.font.name      = "Times New Roman"
        rb.font.size      = Pt(11.5)
        rb.font.bold      = True
        rb.font.color.rgb = RGBColor(31, 78, 121)
    content = lines if isinstance(lines, str) else "\n".join(lines)
    rt = p.add_run(content)
    rt.font.name      = "Times New Roman"
    rt.font.size      = Pt(11.5)
    rt.font.color.rgb = RGBColor(30, 30, 30)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_hyperlink(paragraph, url, text, color="0056B3", underline=True, bold=True):
    """Thêm liên kết hyperlink có thể nhấp chuột trực tiếp trong văn bản Word."""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    new_run   = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr       = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>')
    rPr.append(rFonts)
    if color:
        c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
        rPr.append(c)
    if underline:
        u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
        rPr.append(u)
    if bold:
        b = parse_xml(f'<w:b {nsdecls("w")}/>')
        rPr.append(b)

    sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="23"/>')  # 11.5 pt
    rPr.append(sz)

    new_run.append(rPr)
    text_node = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(text_node)
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)
    return hyperlink

def add_repo_callout(doc, repo_url="https://github.com/t-a0509/Nhan_dang_mau"):
    """Hộp thông tin nổi bật đính kèm liên kết GitHub mã nguồn của bài tập."""
    tbl  = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "EBF3FA")
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    tcPr    = cell._element.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top    w:val="none"/>'
        f'<w:left   w:val="single" w:sz="24" w:space="0" w:color="1F4E79"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right  w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    p.paragraph_format.line_spacing = 1.35

    rb = p.add_run("🔗 LIÊN KẾT MÃ NGUỒN DỰ ÁN (GITHUB REPOSITORY)\n")
    rb.font.name      = "Times New Roman"
    rb.font.size      = Pt(11.5)
    rb.font.bold      = True
    rb.font.color.rgb = RGBColor(31, 78, 121)

    r1 = p.add_run("Toàn bộ mã nguồn thực nghiệm, pipeline xử lý dữ liệu và biểu đồ phân tích được lưu trữ công khai tại kho lưu trữ: ")
    r1.font.name      = "Times New Roman"
    r1.font.size      = Pt(11)
    r1.font.color.rgb = RGBColor(40, 40, 40)

    add_hyperlink(p, repo_url, repo_url, color="0056B3", underline=True, bold=True)

    r2 = p.add_run("\n(Bao gồm: regression_pipeline.py, create_word_report.py, requirements.txt và toàn bộ biểu đồ thực nghiệm chất lượng cao).")
    r2.font.name      = "Times New Roman"
    r2.font.size      = Pt(10.5)
    r2.font.italic    = True
    r2.font.color.rgb = RGBColor(80, 80, 80)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

def add_figure(doc, rel_path, caption, width=6.0):
    """Chèn hình ảnh kèm chú thích căn giữa, in nghiêng."""
    full = os.path.join(BASE_DIR, rel_path)
    if os.path.exists(full):
        pi = doc.add_paragraph()
        pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pi.paragraph_format.space_before = Pt(8)
        pi.paragraph_format.space_after  = Pt(2)
        pi.add_run().add_picture(full, width=Inches(width))
        pc = doc.add_paragraph()
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pc.paragraph_format.space_before = Pt(0)
        pc.paragraph_format.space_after  = Pt(10)
        rc = pc.add_run(caption)
        rc.font.name      = "Times New Roman"
        rc.font.size      = Pt(10.5)
        rc.font.italic    = True
        rc.font.color.rgb = RGBColor(80, 80, 80)
    else:
        add_p(doc, f"[Không tìm thấy ảnh: {rel_path}]", italic=True)

def add_separator(doc):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot  = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single')
    bot.set(qn('w:sz'),    '6')
    bot.set(qn('w:space'), '1')
    bot.set(qn('w:color'), 'CCCCCC')
    pBdr.append(bot)
    pPr.append(pBdr)

def make_header_row(table, headers, widths, bg="1F4E79", font_size=10):
    row = table.rows[0].cells
    for i, (title, w) in enumerate(zip(headers, widths)):
        c       = row[i]
        c.text  = title
        c.width = w
        set_cell_background(c, bg)
        set_cell_margins(c, top=90, bottom=90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.name      = "Times New Roman"
            r.font.bold      = True
            r.font.size      = Pt(font_size)
            r.font.color.rgb = RGBColor(255, 255, 255)

def fill_data_rows(table, data, widths, start_row=1,
                   alt_colors=("FFFFFF", "F2F7FC"),
                   highlight_row=None, highlight_color="E6F4EA",
                   col_aligns=None, font_size=10):
    for ri, row_data in enumerate(data):
        cells = table.rows[start_row + ri].cells
        bg    = alt_colors[ri % 2]
        if highlight_row is not None and ri == highlight_row:
            bg = highlight_color
        for ci, text in enumerate(row_data):
            c       = cells[ci]
            c.text  = text
            c.width = widths[ci]
            set_cell_background(c, bg)
            set_cell_margins(c, top=70, bottom=70)
            p = c.paragraphs[0]
            if col_aligns:
                p.alignment = col_aligns[ci]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(font_size)
                if highlight_row is not None and ri == highlight_row:
                    r.font.bold = True


# ─────────────────────────────── MAIN BUILDER ─────────────────────────────────

def build_report():
    # ── Bước 1: Mở bìa làm tài liệu gốc ────────────────────────────────────
    # Quan trọng: dùng bìa làm BASE doc để giữ nguyên logo, font và tất cả
    # relationship (ảnh) có sẵn trong file bìa.
    print("[*] Mở bìa tiểu luận làm tài liệu gốc...")
    doc = Document(BIA_PATH)

    # ── Bước 2: Thêm section break (ngắt trang) sau bìa ─────────────────────
    # Section 0 (Trang bìa): Cần giữ nguyên:
    # 1. Khung viền trang (w:pgBorders) màu xanh 8EA9C1 viền quanh toàn bộ trang bìa
    # 2. Header & Footer (rId5, rId6: "Hà Nội, tháng 9 năm 2026.")
    # 3. Kích thước trang A4 và lề trang bìa chuẩn
    last_para = doc.paragraphs[-1]
    pPr       = last_para._element.get_or_add_pPr()
    sectPr_cover = parse_xml(
        f'<w:sectPr {nsdecls("w", "r")}>\n'
        f'  <w:headerReference r:id="rId5" w:type="default"/>\n'
        f'  <w:footerReference r:id="rId6" w:type="default"/>\n'
        f'  <w:pgSz w:w="11906" w:h="16838"/>\n'
        f'  <w:pgMar w:top="850" w:right="1020" w:bottom="964" w:left="1020" w:header="720" w:footer="720" w:gutter="0"/>\n'
        f'  <w:pgBorders w:offsetFrom="page">\n'
        f'    <w:top    w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
        f'    <w:left   w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
        f'    <w:bottom w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
        f'    <w:right  w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
        f'  </w:pgBorders>\n'
        f'  <w:cols w:space="720" w:num="1"/>\n'
        f'  <w:docGrid w:linePitch="360" w:charSpace="0"/>\n'
        f'</w:sectPr>'
    )
    pPr.append(sectPr_cover)

    print("[✓] Đã tạo ngắt section sau bìa, bảo lưu khung viền (pgBorders) và footer trang bìa.")

    # ── Bước 3: MỤC LỤC ─────────────────────────────────────────────────────
    add_heading(doc, "MỤC LỤC", level=1)
    toc_items = [
        ("1.", "Đặt vấn đề và mục tiêu bài toán"),
        ("2.", "Khám phá và phân tích dữ liệu"),
        ("3.", "Thiết kế Pipeline: Tiền xử lý và giảm chiều PCA"),
        ("4.", "Các mô hình hồi quy và phương pháp đánh giá độ ổn định"),
        ("5.", "Kết quả thực nghiệm chi tiết"),
        ("6.", "Phân tích chuyên sâu tác động của PCA"),
        ("7.", "Kết luận và kiến nghị mô hình"),
        ("8.", "Tài liệu tham khảo"),
    ]
    for num, title in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(4)
        p.paragraph_format.line_spacing = 1.2
        rn = p.add_run(f"{num}  {title}")
        rn.font.name      = "Times New Roman"
        rn.font.size      = Pt(12)
        rn.font.color.rgb = RGBColor(31, 78, 121)

    doc.add_page_break()

    # ── Bước 4: CHƯƠNG 1 — Đặt vấn đề ──────────────────────────────────────
    add_heading(doc, "1. ĐẶT VẤN ĐỀ VÀ MỤC TIÊU BÀI TOÁN", level=1)

    add_repo_callout(doc, repo_url="https://github.com/t-a0509/Nhan_dang_mau")

    add_p(doc,
        "Trong lĩnh vực Nhận dạng mẫu và Học máy, bài toán hồi quy với biến mục tiêu "
        "liên tục giữ vai trò thiết yếu trong nhiều ứng dụng thực tiễn: dự báo giá bất "
        "động sản, phân tích rủi ro tài chính, dự báo nhu cầu năng lượng hay chẩn đoán "
        "y tế định lượng. Tuy nhiên, dữ liệu trong môi trường thực tế hầu như không bao "
        "giờ 'sạch' và luôn kèm theo những thách thức kỹ thuật cần được xử lý một cách "
        "có hệ thống.",
        indent=True
    )
    add_p(doc,
        "Ba thách thức điển hình nhất cần giải quyết là: đa cộng tuyến giữa các biến "
        "đầu vào khiến mô hình tuyến tính trở nên bất ổn định; sự chênh lệch lớn về "
        "thang đo giữa các thuộc tính ảnh hưởng không công bằng đến các thuật toán "
        "dựa trên khoảng cách; và nguy cơ bùng nổ số chiều khi số lượng đặc trưng lớn "
        "có thể dẫn đến quá khớp, tăng chi phí tính toán. Bài tập này yêu cầu xây dựng "
        "một chương trình dự báo hoàn chỉnh theo hướng pipeline, trong đó tất cả các "
        "bước từ tiền xử lý đến mô hình hóa được đóng gói thống nhất. Mục tiêu cụ thể "
        "gồm năm điểm:",
        indent=True
    )

    goals = [
        ("(1)",
         "Xây dựng Pipeline chuẩn mực bằng sklearn.pipeline.Pipeline, đóng gói tuần tự "
         "các bước Imputation, Scaling, PCA và Regressor, nhằm đảm bảo không xảy ra rò "
         "rỉ dữ liệu trong quá trình đánh giá."),
        ("(2)",
         "Thiết lập mô hình cơ sở Dummy Baseline — luôn dự báo giá trị trung bình của "
         "tập huấn luyện — làm mốc tham chiếu tối thiểu mà mọi mô hình học máy cần "
         "vượt qua."),
        ("(3)",
         "Thử nghiệm và so sánh 8 mô hình hồi quy đại diện cho các trường phái thuật "
         "toán chính: tuyến tính, phi tham số dựa trên khoảng cách, cây quyết định, "
         "học kết hợp và máy vector hỗ trợ."),
        ("(4)",
         "Đánh giá toàn diện và khách quan dựa trên ba chỉ số MAE, RMSE, R² cùng với "
         "độ ổn định qua kỹ thuật 5-Fold Cross-Validation."),
        ("(5)",
         "Phân tích định lượng tác động thực tế của PCA: so sánh hiệu năng khi giảm "
         "chiều từ 8 xuống 5 và khảo sát sự biến thiên khi số thành phần PCA thay đổi "
         "từ k = 1 đến k = 8."),
    ]
    for num, text in goals:
        add_bullet(doc, text, bold_prefix=num)

    add_separator(doc)

    # ── Bước 5: CHƯƠNG 2 — EDA ──────────────────────────────────────────────
    add_heading(doc, "2. KHÁM PHÁ VÀ PHÂN TÍCH DỮ LIỆU", level=1)

    add_heading(doc, "2.1. Lựa chọn bộ dữ liệu", level=2)
    add_p(doc,
        "Bộ dữ liệu được sử dụng là California Housing — một bộ dữ liệu kinh điển được "
        "xây dựng từ cuộc điều tra dân số bang California năm 1990. Bộ dữ liệu gồm "
        "20.640 mẫu điều tra, mỗi mẫu đại diện cho một khu dân cư với 8 thuộc tính "
        "phản ánh các đặc trưng kinh tế, xã hội và tọa độ địa lý của khu vực đó.",
        indent=True
    )
    add_p(doc,
        "Biến mục tiêu là MedHouseVal — giá trị nhà trung vị của khu vực, tính theo "
        "đơn vị 100.000 USD. Đây là biến liên tục, hoàn toàn phù hợp với yêu cầu bài "
        "toán hồi quy. Dù bộ dữ liệu không có giá trị khuyết thiếu, bước Imputation "
        "vẫn được đưa vào Pipeline để đảm bảo tính tổng quát và khả năng hoạt động "
        "ổn định khi áp dụng lên dữ liệu mới trong thực tế.",
        indent=True
    )

    add_p(doc, "Bảng 1: Mô tả các thuộc tính trong bộ dữ liệu California Housing",
          italic=True, space_after=3)
    t1 = doc.add_table(rows=10, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    h1 = ["Tên thuộc tính", "Ý nghĩa", "Kiểu dữ liệu", "Thống kê (Mean ± Std)"]
    w1 = [Inches(1.35), Inches(2.55), Inches(1.05), Inches(1.55)]
    make_header_row(t1, h1, w1, font_size=10)
    d1 = [
        ["MedInc",            "Thu nhập trung vị của khu dân cư",           "Liên tục", "3,87 ± 1,90  ($10k USD)"],
        ["HouseAge",          "Tuổi thọ trung vị của ngôi nhà",             "Liên tục", "28,64 ± 12,59  (năm)"],
        ["AveRooms",          "Số phòng trung bình trên mỗi hộ gia đình",   "Liên tục", "5,43 ± 2,47  (phòng)"],
        ["AveBedrms",         "Số phòng ngủ trung bình trên mỗi hộ",        "Liên tục", "1,10 ± 0,47  (phòng)"],
        ["Population",        "Tổng dân số của khối điều tra",              "Rời rạc",  "1.425 ± 1.132  (người)"],
        ["AveOccup",          "Số nhân khẩu trung bình trên mỗi hộ",        "Liên tục", "3,07 ± 10,39  (người)"],
        ["Latitude",          "Vĩ độ địa lý của khu vực",                   "Liên tục", "35,63° ± 2,14°"],
        ["Longitude",         "Kinh độ địa lý của khu vực",                 "Liên tục", "−119,57° ± 2,00°"],
        ["MedHouseVal (Target)", "Giá trị nhà trung vị — Biến mục tiêu",    "Liên tục", "2,07 ± 1,15  ($100k USD)"],
    ]
    col_a1 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT,
               WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
    fill_data_rows(t1, d1, w1, highlight_row=8, highlight_color="FFF3E0",
                   col_aligns=col_a1, font_size=10)
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_heading(doc, "2.2. Trực quan hóa và nhận xét EDA", level=2)
    add_figure(doc, "results/eda_correlation_distribution.png",
               "Hình 1: Phân bố tần suất biến mục tiêu MedHouseVal và ma trận tương quan "
               "Pearson giữa các thuộc tính trong bộ dữ liệu",
               width=6.2)

    add_p(doc,
        "Nhìn vào biểu đồ phân bố, biến mục tiêu MedHouseVal có dạng lệch phải nhẹ, "
        "phần lớn tập trung trong khoảng từ 100.000 đến 250.000 USD. Đáng chú ý là "
        "một lượng mẫu đáng kể bị 'chặn trần' ở mức 500.001 USD — hệ quả của cách thu "
        "thập dữ liệu gốc áp đặt giới hạn trên, không phải giá trị thực. Hiện tượng "
        "này có thể ảnh hưởng nhỏ đến sai số của mô hình ở phân khúc nhà đắt tiền.",
        indent=True
    )
    add_p(doc,
        "Từ ma trận tương quan Pearson, ba quan sát nổi bật nhất là: thuộc tính MedInc "
        "có hệ số tương quan dương mạnh nhất với giá nhà, đạt r = +0,69, xác nhận đây "
        "là biến dự báo có giá trị nhất. Cặp biến Latitude và Longitude có tương quan "
        "âm rất cao ở mức r = −0,92, phản ánh dải địa lý kéo dài theo trục Bắc – Nam "
        "của bang California. Cặp AveRooms và AveBedrms có tương quan dương cao, đạt "
        "r = +0,85, cho thấy sự dư thừa thông tin rõ rệt giữa hai biến này — đây chính "
        "là điều kiện lý tưởng để kiểm tra hiệu quả của PCA trong việc loại bỏ cộng tuyến.",
        indent=True
    )

    add_separator(doc)

    # ── Bước 6: CHƯƠNG 3 — Pipeline ─────────────────────────────────────────
    add_heading(doc, "3. THIẾT KẾ PIPELINE: TIỀN XỬ LÝ VÀ GIẢM CHIỀU PCA", level=1)

    add_heading(doc, "3.1. Cấu trúc Pipeline và nguyên tắc chống rò rỉ dữ liệu", level=2)
    add_p(doc,
        "Một trong những sai lầm phổ biến nhất trong thực nghiệm học máy là thực hiện "
        "tiền xử lý trên toàn bộ dữ liệu trước khi phân chia tập huấn luyện và tập "
        "kiểm định, dẫn đến rò rỉ thông tin từ tập kiểm định vào quá trình huấn luyện, "
        "khiến kết quả đánh giá bị lạc quan giả tạo. Toàn bộ bài tập sử dụng "
        "sklearn.pipeline.Pipeline để ngăn chặn hoàn toàn vấn đề này: mọi phép biến đổi "
        "thống kê chỉ được tính toán tham số trên tập huấn luyện rồi áp dụng sang tập "
        "kiểm định, không có chiều ngược lại.",
        indent=True
    )

    add_callout(doc, [
        "Bước 1 — SimpleImputer(strategy='median'):  Điền giá trị khuyết thiếu "
        "bằng trung vị của tập huấn luyện.",
        "Bước 2 — StandardScaler():  Chuẩn hóa z-score, đưa tất cả thuộc tính về "
        "kỳ vọng 0 và độ lệch chuẩn 1.",
        "Bước 3 — PCA(n_components=k)  [tùy chọn]:  Chiếu trực giao dữ liệu sang "
        "không gian k chiều cực đại hóa phương sai.",
        "Bước 4 — Regressor:  Huấn luyện mô hình hồi quy trên không gian đặc trưng "
        "đã tiền xử lý.",
    ], title="Các bước trong Pipeline:")

    add_heading(doc, "3.2. Vai trò của chuẩn hóa trước PCA", level=2)
    add_p(doc,
        "PCA xác định các hướng chính dựa trên ma trận hiệp phương sai của dữ liệu. "
        "Khi các thuộc tính có thang đo rất khác nhau — ví dụ Population có thể lên tới "
        "hàng nghìn trong khi AveBedrms chỉ dao động quanh 1,1 — ma trận hiệp phương sai "
        "sẽ bị chi phối bởi chiều có phương sai số học lớn nhất, khiến các thành phần "
        "chính không phản ánh đúng cấu trúc tương quan thực chất. Vì vậy, bước "
        "StandardScaler trước PCA là yêu cầu bắt buộc về mặt toán học.",
        indent=True
    )

    add_heading(doc, "3.3. Phân tích phương sai tích lũy PCA", level=2)
    add_figure(doc, "results/pca_variance.png",
               "Hình 2: Scree Plot — Tỷ lệ phương sai giải thích từng thành phần (cột xanh) "
               "và phương sai tích lũy (đường đỏ) của 8 thành phần chính PCA",
               width=5.8)

    add_p(doc, "Bảng 2: Tỷ lệ phương sai giải thích của từng thành phần chính",
          italic=True, space_after=3)
    t2 = doc.add_table(rows=9, cols=4)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    h2 = ["Thành phần chính", "Giá trị riêng", "Tỷ lệ phương sai (%)", "Phương sai tích lũy (%)"]
    w2 = [Inches(1.5), Inches(1.4), Inches(1.7), Inches(1.9)]
    make_header_row(t2, h2, w2, font_size=10)
    d2 = [
        ["PC1",    "2,027", "25,34 %", "25,34 %"],
        ["PC2",    "1,882", "23,52 %", "48,85 %"],
        ["PC3",    "1,271", "15,89 %", "64,74 %"],
        ["PC4",    "1,031", "12,89 %", "77,63 %"],
        ["PC5  ★", "1,003", "12,54 %", "90,17 %  ← Ngưỡng ≥ 90 %"],
        ["PC6",    "0,659",  "8,24 %", "98,41 %  ← Ngưỡng ≥ 95 %"],
        ["PC7",    "0,082",  "1,02 %", "99,43 %"],
        ["PC8",    "0,045",  "0,57 %", "100,00 %"],
    ]
    fill_data_rows(t2, d2, w2, highlight_row=4, highlight_color="D6EAF8", font_size=10)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_p(doc,
        "Kết quả phân tích cho thấy chỉ cần 5 thành phần chính đầu tiên đã bảo toàn "
        "được 90,17% tổng lượng thông tin phương sai. Hai thành phần cuối cùng PC7 và "
        "PC8 chỉ chiếm 1,59% phương sai, chủ yếu phản ánh sự dư thừa tuyến tính giữa "
        "các cặp biến tương quan cao. Dựa vào kết quả này, k = 5 được chọn làm cấu hình "
        "PCA chuẩn trong các thực nghiệm so sánh.",
        indent=True
    )

    add_separator(doc)

    # ── Bước 7: CHƯƠNG 4 — Mô hình ──────────────────────────────────────────
    add_heading(doc, "4. CÁC MÔ HÌNH HỒI QUY VÀ PHƯƠNG PHÁP ĐÁNH GIÁ ĐỘ ỔN ĐỊNH", level=1)

    add_heading(doc, "4.1. Danh sách mô hình thực nghiệm", level=2)
    add_p(doc,
        "Thực nghiệm được tiến hành trên 9 mô hình đại diện cho các trường phái thuật "
        "toán chính, từ mô hình cơ sở đến các kỹ thuật học kết hợp hiện đại:",
        indent=True
    )
    models_desc = [
        ("1. Dummy Baseline",
         "Luôn dự báo giá trị trung bình của tập huấn luyện, bất kể đặc trưng đầu vào. "
         "Đây là mốc tham chiếu tối thiểu — mọi mô hình học máy chỉ có giá trị thực tiễn "
         "khi vượt trội đáng kể so với ngưỡng này."),
        ("2. Linear Regression",
         "Mô hình tuyến tính cổ điển tối thiểu hóa tổng bình phương phần dư. Đơn giản "
         "nhưng cho kết quả giải thích tốt và đóng vai trò baseline tuyến tính quan trọng."),
        ("3. Ridge Regression",
         "Mô hình tuyến tính bổ sung thêm điều chuẩn L2 với alpha = 1,0. Ổn định nghiệm "
         "hệ số khi các biến có đa cộng tuyến, phù hợp với dữ liệu có AveRooms và "
         "AveBedrms tương quan cao."),
        ("4. Lasso Regression",
         "Mô hình tuyến tính với điều chuẩn L1, alpha = 0,01. Có khả năng triệt tiêu "
         "các hệ số không quan trọng về đúng 0, từ đó thực hiện lựa chọn đặc trưng tự động."),
        ("5. K-Nearest Neighbors  (k = 7)",
         "Mô hình phi tham số dự báo bằng cách lấy trung bình 7 láng giềng gần nhất theo "
         "khoảng cách Euclidean. Rất nhạy cảm với thang đo và số chiều không gian đặc trưng."),
        ("6. Decision Tree  (max_depth = 10)",
         "Cây quyết định phân cấp phi tuyến với chiều sâu tối đa là 10. Trực quan, dễ "
         "giải thích nhưng dễ bị quá khớp nếu không kiểm soát độ sâu."),
        ("7. Random Forest  (70 cây)",
         "Mô hình học kết hợp theo kỹ thuật Bagging — huấn luyện 70 cây quyết định trên "
         "các tập con dữ liệu và đặc trưng khác nhau, sau đó lấy trung bình kết quả để "
         "giảm phương sai và chống quá khớp."),
        ("8. HistGradientBoosting",
         "Thuật toán Gradient Boosting hiện đại tối ưu hóa theo histogram, thiết kế để "
         "xử lý dữ liệu lớn cực nhanh. Mỗi cây mới được huấn luyện để sửa phần dư lỗi "
         "của cây trước, dần tối ưu hóa hàm mất mát."),
        ("9. Support Vector Regressor  (nhân RBF)",
         "Tìm siêu phẳng biên tối ưu trong không gian đặc trưng phi tuyến mở rộng "
         "bằng nhân hàm RBF với dải dung sai epsilon = 0,1. Hiệu năng tốt nhưng chi "
         "phí tính toán cao nhất trong nhóm."),
    ]
    for bold, text in models_desc:
        add_bullet(doc, text, bold_prefix=bold)

    add_heading(doc, "4.2. Phương pháp đánh giá: 5-Fold Cross-Validation", level=2)
    add_p(doc,
        "Để đánh giá đồng thời độ chính xác và độ ổn định của mỗi mô hình, bài tập sử "
        "dụng kỹ thuật 5-Fold Cross-Validation với xáo trộn ngẫu nhiên. Toàn bộ 20.640 "
        "mẫu được chia thành 5 phần bằng nhau; tại mỗi vòng lặp, 4 phần khoảng 16.512 "
        "mẫu được dùng để huấn luyện và 1 phần khoảng 4.128 mẫu dùng để kiểm định. "
        "Quá trình lặp lại 5 lần độc lập với random_state cố định để đảm bảo tính "
        "tái lập hoàn toàn.",
        indent=True
    )

    add_heading(doc, "4.3. Các chỉ số đánh giá hiệu năng", level=2)
    metrics_desc = [
        ("MAE — Mean Absolute Error",
         "Trung bình sai số tuyệt đối theo đúng đơn vị gốc (100.000 USD). Trực quan, "
         "dễ giải thích và ít bị ảnh hưởng bởi các giá trị ngoại lai lớn."),
        ("RMSE — Root Mean Squared Error",
         "Căn bậc hai của trung bình bình phương sai số. Phạt nặng hơn các lỗi dự báo "
         "lớn so với MAE, phản ánh tốt hơn khi mô hình sai nghiêm trọng ở một số điểm."),
        ("R² — Hệ số xác định",
         "Tỷ lệ phương sai biến mục tiêu được mô hình giải thích. R² = 1 là dự báo "
         "hoàn hảo; R² = 0 tương đương Dummy Baseline; R² âm nghĩa là mô hình tệ hơn "
         "cả việc chỉ đoán giá trị trung bình."),
        ("Std — Độ lệch chuẩn qua 5 folds",
         "Đo lường độ ổn định. Một mô hình tốt cần cả giá trị trung bình cao lẫn độ "
         "lệch chuẩn nhỏ. Độ lệch chuẩn lớn cho thấy mô hình bị ảnh hưởng mạnh bởi "
         "sự phân bố cụ thể của dữ liệu, không đáng tin cậy trong triển khai thực tế."),
    ]
    for bold, text in metrics_desc:
        add_bullet(doc, text, bold_prefix=bold)

    add_separator(doc)

    # ── Bước 8: CHƯƠNG 5 — Kết quả ──────────────────────────────────────────
    add_heading(doc, "5. KẾT QUẢ THỰC NGHIỆM CHI TIẾT", level=1)

    add_heading(doc, "5.1. Bảng tổng hợp số liệu (5-Fold CV: Mean ± Std)", level=2)
    add_p(doc, "Bảng 3: Kết quả đánh giá tất cả các mô hình qua 5-Fold Cross-Validation",
          italic=True, space_after=3)
    t3 = doc.add_table(rows=19, cols=6)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    h3 = ["Mô hình", "Cấu hình PCA", "MAE (Mean ± Std)", "RMSE (Mean ± Std)", "R² (Mean ± Std)", "T.gian (s)"]
    w3 = [Inches(1.75), Inches(1.15), Inches(1.35), Inches(1.35), Inches(1.35), Inches(0.8)]
    make_header_row(t3, h3, w3, font_size=9.5)
    d3 = [
        ["Dummy Baseline",       "Không PCA (8D)", "0,9118 ± 0,0067", "1,1539 ± 0,0103", "−0,0003 ± 0,0003", "0,01"],
        ["Dummy Baseline",       "Có PCA (5D)",    "0,9118 ± 0,0067", "1,1539 ± 0,0103", "−0,0003 ± 0,0003", "0,01"],
        ["Linear Regression",    "Không PCA (8D)", "0,5317 ± 0,0084", "0,7283 ± 0,0149", "0,6014 ± 0,0170",  "0,01"],
        ["Linear Regression",    "Có PCA (5D)",    "0,6458 ± 0,0057", "0,8533 ± 0,0126", "0,4529 ± 0,0161",  "0,01"],
        ["Ridge Regression",     "Không PCA (8D)", "0,5317 ± 0,0084", "0,7282 ± 0,0149", "0,6014 ± 0,0170",  "0,01"],
        ["Ridge Regression",     "Có PCA (5D)",    "0,6458 ± 0,0057", "0,8533 ± 0,0126", "0,4529 ± 0,0161",  "0,01"],
        ["Lasso Regression",     "Không PCA (8D)", "0,5354 ± 0,0079", "0,7302 ± 0,0124", "0,5993 ± 0,0140",  "0,04"],
        ["Lasso Regression",     "Có PCA (5D)",    "0,6468 ± 0,0056", "0,8529 ± 0,0131", "0,4533 ± 0,0167",  "0,01"],
        ["K-Nearest Neighbors",  "Không PCA (8D)", "0,4308 ± 0,0172", "0,6328 ± 0,0236", "0,6987 ± 0,0221",  "0,04"],
        ["K-Nearest Neighbors",  "Có PCA (5D)",    "0,5220 ± 0,0085", "0,7424 ± 0,0107", "0,5857 ± 0,0138",  "0,03"],
        ["Decision Tree",        "Không PCA (8D)", "0,4248 ± 0,0085", "0,6406 ± 0,0118", "0,6917 ± 0,0108",  "0,08"],
        ["Decision Tree",        "Có PCA (5D)",    "0,5760 ± 0,0138", "0,8231 ± 0,0174", "0,4908 ± 0,0209",  "0,06"],
        ["SVR (nhân RBF)",       "Không PCA (8D)", "0,3906 ± 0,0110", "0,5861 ± 0,0185", "0,7417 ± 0,0160",  "7,71"],
        ["SVR (nhân RBF)",       "Có PCA (5D)",    "0,5210 ± 0,0083", "0,7587 ± 0,0146", "0,5674 ± 0,0182",  "7,76"],
        ["Random Forest",        "Không PCA (8D)", "0,3455 ± 0,0070", "0,5193 ± 0,0099", "0,7973 ± 0,0081",  "1,95"],
        ["Random Forest",        "Có PCA (5D)",    "0,5205 ± 0,0094", "0,7341 ± 0,0128", "0,5949 ± 0,0172",  "1,46"],
        ["HistGradientBoosting", "Không PCA (8D)", "0,3118 ± 0,0058", "0,4665 ± 0,0103", "0,8364 ± 0,0079",  "0,30"],
        ["HistGradientBoosting", "Có PCA (5D)",    "0,5196 ± 0,0078", "0,7264 ± 0,0112", "0,6034 ± 0,0149",  "0,26"],
    ]
    col_a3 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER,
               WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
               WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
    fill_data_rows(t3, d3, w3, highlight_row=16, highlight_color="D5F5E3",
                   col_aligns=col_a3, font_size=9)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_heading(doc, "5.2. Biểu đồ so sánh trực quan các mô hình", level=2)
    add_figure(doc, "results/model_comparison.png",
               "Hình 3: So sánh 9 mô hình hồi quy không dùng PCA qua ba chỉ số R², RMSE "
               "và MAE, kèm thanh sai số chuẩn của 5-Fold CV",
               width=6.3)

    add_p(doc,
        "Nhìn vào kết quả thực nghiệm, các mô hình tự nhiên phân thành bốn nhóm hiệu "
        "năng rõ rệt.",
        indent=True
    )
    add_p(doc,
        "Nhóm dẫn đầu gồm hai mô hình học kết hợp. HistGradientBoosting đạt kết quả tốt "
        "nhất tuyệt đối với R² = 0,8364, RMSE = 0,4665 và MAE = 0,3118. Quy đổi ra thực "
        "tế, sai số trung bình chỉ khoảng 31.180 USD trên một ngôi nhà trị giá trung bình "
        "206.800 USD, tức chưa đến 15%. Đặc biệt, thời gian huấn luyện chỉ 0,30 giây và "
        "độ lệch chuẩn qua 5 folds dưới 0,008 — mô hình vừa chính xác vừa ổn định. "
        "Random Forest đứng thứ hai với R² = 0,7973 nhưng thời gian huấn luyện lâu hơn "
        "gần 7 lần.",
        indent=True
    )
    add_p(doc,
        "Nhóm thứ hai là SVR nhân RBF với R² = 0,7417, thể hiện khả năng học quan hệ phi "
        "tuyến tốt, nhưng chi phí tính toán rất cao ở mức 7,71 giây mỗi fold — không thực "
        "tế khi dữ liệu lớn hơn. Nhóm thứ ba gồm KNN và Decision Tree đạt R² xấp xỉ 0,69 "
        "đến 0,70. KNN có biến thiên giữa các fold cao hơn, phản ánh tính nhạy cảm với sự "
        "phân phối cụ thể của dữ liệu. Cuối cùng, ba mô hình tuyến tính Linear Regression, "
        "Ridge và Lasso có hiệu năng gần như tương đương nhau với R² khoảng 0,60, cho thấy "
        "quan hệ phi tuyến trong dữ liệu là đáng kể và không thể bắt bằng mô hình tuyến tính "
        "thuần túy. Dummy Baseline cho R² xấp xỉ 0 — tất cả 8 mô hình học máy đều vượt "
        "trội áp đảo.",
        indent=True
    )

    add_heading(doc, "5.3. Biểu đồ phân tán Thực tế vs Dự báo", level=2)
    add_figure(doc, "results/actual_vs_predicted.png",
               "Hình 4: Biểu đồ phân tán so sánh giá trị thực tế với giá trị dự báo — "
               "Dummy Baseline (trái) và HistGradientBoosting (phải)",
               width=6.2)
    add_p(doc,
        "Biểu đồ phân tán làm rõ sự khác biệt định tính giữa hai cực hiệu năng. Với "
        "Dummy Baseline, toàn bộ điểm dự báo nằm trên một đường ngang duy nhất tương "
        "ứng với giá trị trung bình — mô hình không học được bất kỳ cấu trúc nào. Ngược "
        "lại, HistGradientBoosting cho thấy các điểm hội tụ chặt chẽ xung quanh đường "
        "chéo lý tưởng, thể hiện khả năng dự báo chính xác xuyên suốt toàn bộ phổ giá "
        "từ thấp đến cao. Chỉ ở mức giá trần 500.001 USD có sự bão hòa nhẹ do hiệu ứng "
        "cắt ngưỡng của dữ liệu gốc.",
        indent=True
    )

    add_separator(doc)

    # ── Bước 9: CHƯƠNG 6 — Phân tích PCA ────────────────────────────────────
    add_heading(doc, "6. PHÂN TÍCH CHUYÊN SÂU TÁC ĐỘNG CỦA PCA", level=1)

    add_heading(doc, "6.1. So sánh trực tiếp: Có PCA (5D) vs Không PCA (8D)", level=2)
    add_figure(doc, "results/pca_impact.png",
               "Hình 5: So sánh R² và RMSE của từng mô hình giữa hai cấu hình "
               "không dùng PCA (8 chiều) và có dùng PCA (5 chiều)",
               width=6.2)
    add_p(doc,
        "Kết quả cho thấy một xu hướng nhất quán và rõ ràng: khi áp dụng PCA để giảm "
        "từ 8 chiều xuống 5 chiều, dù vẫn giữ lại 90,17% phương sai, hiệu năng của tất "
        "cả các mô hình đều suy giảm đáng kể. Cụ thể, R² của HistGradientBoosting giảm "
        "từ 0,8364 xuống 0,6034, tức giảm 27,9%; Random Forest từ 0,7973 xuống 0,5949, "
        "giảm 25,3%; Decision Tree từ 0,6917 xuống 0,4908, giảm 29,0%; và Linear "
        "Regression từ 0,6014 xuống 0,4529, giảm 24,6%. RMSE của tất cả các mô hình "
        "tăng từ 15% đến 30% so với cấu hình không dùng PCA.",
        indent=True
    )

    add_heading(doc, "6.2. Khảo sát hiệu năng theo số chiều PCA từ k = 1 đến k = 8", level=2)
    add_figure(doc, "results/pca_components_sweep.png",
               "Hình 6: Sự biến thiên của R² và RMSE khi số thành phần PCA k tăng "
               "từ 1 đến 8 trên ba mô hình tiêu biểu",
               width=6.2)
    add_p(doc,
        "Biểu đồ khảo sát số chiều cho thấy bức tranh chi tiết hơn. Với chỉ 1 hoặc 2 "
        "thành phần chính, hiệu năng rất thấp, R² dưới 0,20, do không đủ thông tin để "
        "tách biệt cấu trúc giá nhà. Từ k = 3 đến k = 4, R² tăng vọt lên khoảng 0,40 "
        "đến 0,48 khi thêm được thông tin về thu nhập và không gian địa lý. Giai đoạn "
        "từ k = 5 đến k = 7 tiếp tục cải thiện đều đặn nhưng vẫn chưa đạt mức tối đa. "
        "Bước nhảy vọt quyết định xảy ra ở k = 8: khi phục hồi đầy đủ toàn bộ 8 chiều "
        "qua phép quay trực giao PCA, tất cả mô hình đạt lại đúng kết quả ban đầu, "
        "xác nhận rằng không có thông tin nào bị mất trong phép biến đổi này.",
        indent=True
    )

    add_heading(doc, "6.3. Giải thích nguyên nhân PCA làm giảm hiệu năng", level=2)
    add_p(doc,
        "Câu hỏi đặt ra là tại sao PCA giữ lại tới 90,17% phương sai đầu vào nhưng R² "
        "vẫn sụt giảm từ 24% đến 29%? Có ba nguyên nhân học thuật căn bản:",
        indent=True
    )
    add_callout(doc, [
        "Nguyên nhân 1 — PCA là giải thuật học không giám sát:\n"
        "PCA tối đa hóa phương sai trên đặc trưng đầu vào mà hoàn toàn không biết thông "
        "tin về biến mục tiêu. Vì vậy, các thành phần PC6, PC7, PC8 — dù chỉ chiếm ~10% "
        "phương sai của đặc trưng — lại có thể mang tương quan rất cao với giá nhà. Khi "
        "bị loại bỏ, chúng mang đi thông tin dự báo quý giá mà mô hình không thể bù đắp.",
        "\nNguyên nhân 2 — Phá vỡ tương tác phi tuyến địa lý:\n"
        "Latitude và Longitude có quan hệ phi tuyến mạnh với giá nhà, vì các khu ven biển "
        "và thung lũng Silicon có giá rất cao so với nội địa. Phép chiếu tuyến tính của PCA "
        "kết hợp tọa độ với các biến kinh tế, phá vỡ cấu trúc phân cụm địa lý mà các mô "
        "hình cây dựa vào để nhận diện vùng giá cao.",
        "\nNguyên nhân 3 — Số chiều ban đầu đã nhỏ:\n"
        "PCA phát huy hiệu quả nhất khi số chiều rất lớn như ảnh, văn bản hay dữ liệu gen "
        "để chống quá khớp. Với bộ dữ liệu chỉ có 8 chiều, không tồn tại vấn đề bùng nổ "
        "số chiều, nên việc giảm chiều không mang lại lợi ích chống quá khớp mà chỉ làm "
        "mất thông tin.",
    ], title="Ba nguyên nhân học thuật vì sao PCA làm giảm hiệu năng:")

    add_separator(doc)

    # ── Bước 10: CHƯƠNG 7 — Kết luận ────────────────────────────────────────
    add_heading(doc, "7. KẾT LUẬN VÀ KIẾN NGHỊ MÔ HÌNH", level=1)

    add_p(doc, "Từ toàn bộ quá trình thực nghiệm, bài tập rút ra ba kết luận chính:",
          indent=True)

    add_p(doc,
        "Về lựa chọn mô hình, HistGradientBoostingRegressor không dùng PCA là mô hình "
        "vượt trội nhất trên bộ dữ liệu California Housing theo mọi tiêu chí đánh giá. "
        "Với R² = 0,8364, RMSE = 0,4665 và MAE = 0,3118, mô hình giải thích được gần "
        "84% phương sai giá nhà với sai số trung bình chỉ khoảng 31.000 USD. Độ lệch "
        "chuẩn qua 5 folds dưới 0,008 khẳng định tính ổn định cao. Thời gian huấn luyện "
        "chỉ 0,30 giây — nhanh hơn Random Forest gần 7 lần và nhanh hơn SVR 25 lần — "
        "làm cho mô hình này hoàn toàn phù hợp để triển khai trong hệ thống thời gian thực.",
        bold_part="Kết luận 1.", indent=True
    )
    add_p(doc,
        "Về Pipeline tiền xử lý, việc kết hợp SimpleImputer và StandardScaler trong "
        "Pipeline là thiết kế chuẩn mực giúp mô hình kháng lỗi khi dữ liệu có giá trị "
        "khuyết thiếu, đồng thời đảm bảo tính đúng đắn toán học cho các thuật toán nhạy "
        "cảm với thang đo như KNN, SVR, Ridge và Lasso. Đây là thực hành bắt buộc trong "
        "môi trường sản xuất thực tế.",
        bold_part="Kết luận 2.", indent=True
    )
    add_p(doc,
        "Về bài học thực tiễn từ PCA, thực nghiệm minh chứng rõ ràng rằng giảm chiều "
        "không giám sát bằng PCA không đồng nghĩa với việc cải thiện bài toán hồi quy có "
        "giám sát. Trên dữ liệu California Housing — vốn có số chiều nhỏ và các biến địa "
        "lý phi tuyến quan trọng — PCA thực sự làm tổn hại hiệu năng. Tuy nhiên, PCA vẫn "
        "có giá trị phân tích cấu trúc dữ liệu và trực quan hóa trong không gian 2D/3D. "
        "Trong thực tế, quyết định có nên dùng PCA hay không phải luôn được kiểm chứng "
        "bằng Cross-Validation trước khi triển khai.",
        bold_part="Kết luận 3.", indent=True
    )

    add_callout(doc,
        "HistGradientBoostingRegressor không dùng PCA được kiến nghị là mô hình triển "
        "khai tối ưu:\n"
        "R² = 0,8364  |  RMSE = 0,4665  |  MAE = 0,3118  |  σ(R²) = 0,0079  |  "
        "Thời gian huấn luyện: 0,30 giây/fold",
        title="★ Mô hình được kiến nghị:"
    )

    add_separator(doc)

    # ── Bước 11: CHƯƠNG 8 — Tài liệu tham khảo ─────────────────────────────
    add_heading(doc, "8. TÀI LIỆU THAM KHẢO", level=1)
    refs = [
        "[1]  Pedregosa, F. et al. (2011). Scikit-learn: Machine Learning in Python. "
        "Journal of Machine Learning Research, 12, 2825–2830.",
        "[2]  Pace, R. K. & Barry, R. (1997). Sparse spatial autoregressions. "
        "Statistics & Probability Letters, 33(3), 291–297. "
        "(Bộ dữ liệu California Housing.)",
        "[3]  James, G., Witten, D., Hastie, T. & Tibshirani, R. (2013). "
        "An Introduction to Statistical Learning. Springer, New York.",
        "[4]  Bishop, C. M. (2006). Pattern Recognition and Machine Learning. Springer.",
        "[5]  Géron, A. (2022). Hands-On Machine Learning with Scikit-Learn, Keras, "
        "and TensorFlow, 3rd ed. O'Reilly Media.",
        "[6]  Hastie, T., Tibshirani, R. & Friedman, J. (2009). "
        "The Elements of Statistical Learning, 2nd ed. Springer.",
        "[7]  Ke, G. et al. (2017). LightGBM: A Highly Efficient Gradient Boosting "
        "Decision Tree. Advances in Neural Information Processing Systems (NeurIPS), 30.",
        "[8]  Phạm Tuấn Anh (2026). Kho lưu trữ mã nguồn mở: Pipeline hồi quy với tiền xử lý "
        "và giảm chiều PCA (Bài tập 1 - Môn Nhận dạng mẫu). "
        "GitHub Repository: https://github.com/t-a0509/Nhan_dang_mau",
    ]
    for ref in refs:
        p = doc.add_paragraph()
        p.paragraph_format.space_before       = Pt(0)
        p.paragraph_format.space_after        = Pt(5)
        p.paragraph_format.line_spacing       = 1.3
        p.paragraph_format.left_indent        = Pt(24)
        p.paragraph_format.first_line_indent  = Pt(-24)
        r = p.add_run(ref)
        r.font.name      = "Times New Roman"
        r.font.size      = Pt(11.5)
        r.font.color.rgb = RGBColor(30, 30, 30)

    add_separator(doc)

    # Dòng kết thúc
    pe = doc.add_paragraph()
    pe.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pe.paragraph_format.space_before = Pt(18)
    re = pe.add_run("─────────── HẾT BÁO CÁO ───────────")
    re.font.name      = "Times New Roman"
    re.font.size      = Pt(11)
    re.font.bold      = True
    re.font.color.rgb = RGBColor(120, 120, 120)

    # ── Bước 12: Hoàn thiện Section 1 (Nội dung báo cáo) ───────────────────
    if len(doc.sections) > 1:
        sec1 = doc.sections[1]
        sec1_elem = sec1._sectPr

        # Gỡ bỏ tham chiếu header/footer của trang bìa (tránh lặp lại "Hà Nội, tháng 9..." trên mọi trang thân bài)
        for hf in sec1_elem.xpath('w:headerReference | w:footerReference'):
            sec1_elem.remove(hf)

        # Thiết lập lề chuẩn học thuật cho thân bài: Trái 3.0cm, Phải 2.0cm, Trên 2.5cm, Dưới 2.5cm
        for mar in sec1_elem.xpath('w:pgMar'):
            mar.set(qn('w:top'),    '1417')
            mar.set(qn('w:bottom'), '1417')
            mar.set(qn('w:left'),   '1701')
            mar.set(qn('w:right'),  '1134')
            mar.set(qn('w:header'), '720')
            mar.set(qn('w:footer'), '720')

        # Đảm bảo Section 1 có khung viền pgBorders đồng nhất với trang bìa
        if not sec1_elem.xpath('w:pgBorders'):
            pgBorders_body = parse_xml(
                f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">\n'
                f'  <w:top    w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
                f'  <w:left   w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
                f'  <w:bottom w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
                f'  <w:right  w:val="single" w:color="8EA9C1" w:sz="4" w:space="18"/>\n'
                f'</w:pgBorders>'
            )
            sec1_elem.append(pgBorders_body)

        # Số trang cho thân bài (chân trang bên phải)
        sec1.footer.is_linked_to_previous = False
        p_footer = sec1.footer.paragraphs[0]
        p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_footer.paragraph_format.space_before = Pt(4)
        r_ft = p_footer.add_run("Trang ")
        r_ft.font.name = "Times New Roman"
        r_ft.font.size = Pt(10)
        r_ft.font.italic = True
        r_ft.font.color.rgb = RGBColor(120, 120, 120)
        fldSimple = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>')
        p_footer._p.append(fldSimple)

    # ── Lưu file ─────────────────────────────────────────────────────────────
    doc.save(DOCX_PATH)
    print(f"[✓] Đã tạo thành công báo cáo: {DOCX_PATH}")


if __name__ == "__main__":
    build_report()
