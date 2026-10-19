"""
Script tạo báo cáo Word (.docx) hoàn chỉnh, chuẩn học thuật
cho Bài tập 1 - Môn Nhận dạng mẫu (Pattern Recognition) - Thạc sĩ.

Chiến lược ghép bìa:
  - Mở BÌA TIỂU LUẬN.docx làm tài liệu GỐC (giữ nguyên logo, font, định dạng).
  - Thay tiêu đề / tên môn học trên bìa cho đúng Bài tập 1 - Nhận dạng mẫu.
  - Thêm page break sau bìa rồi nối toàn bộ nội dung báo cáo vào ngay sau.
  - Lưu đè ra Bao_cao_Bai_tap_1.docx.

Mọi số liệu trong báo cáo được đọc trực tiếp từ các file CSV trong thư mục results/
(do regression_pipeline.py sinh ra), nên phải chạy regression_pipeline.py trước.
"""

import os
import sys

import pandas as pd

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

# Nội dung trang bìa cần thay (file bìa gốc là mẫu dùng chung cho các môn khác)
COVER_REPLACEMENTS = {
    "BÀI TIỂU LUẬN": "BÁO CÁO BÀI TẬP 1",
    "MÔN: TRIẾT HỌC MÁC - LÊNIN": "MÔN: NHẬN DẠNG MẪU",
}

DISPLAY_NAMES = {"Support Vector Regressor": "SVR (nhân RBF)"}
TARGET_MEAN   = 2.0686   # Giá trị trung bình MedHouseVal trên toàn bộ dữ liệu ($100k)


# ─────────────────────────────── DỮ LIỆU KẾT QUẢ ──────────────────────────────

def load_results():
    """Đọc toàn bộ kết quả thực nghiệm do regression_pipeline.py xuất ra."""
    def read(name):
        path = os.path.join(RESULTS_DIR, name)
        if not os.path.exists(path):
            sys.exit(f"[!] Thiếu {path}. Hãy chạy 'python regression_pipeline.py' trước.")
        return pd.read_csv(path)

    return {
        "cv":       read("results_summary.csv"),
        "rank":     read("model_ranking.csv"),
        "sweep":    read("pca_sweep_results.csv"),
        "test":     read("test_results.csv"),
        "variance": read("pca_variance.csv"),
    }

def vn(x, d=4):
    """Định dạng số theo kiểu Việt Nam: dấu phẩy thập phân, dấu trừ chuẩn."""
    s = f"{x:,.{d}f}".replace(",", "_").replace(".", ",").replace("_", ".")
    return s.replace("-", "−")

def pm(mean, std, d=4):
    return f"{vn(mean, d)} ± {vn(std, d)}"

def usd(x_100k):
    """Đổi giá trị đơn vị 100.000 USD sang chuỗi USD làm tròn đến hàng chục."""
    return vn(round(x_100k * 100000, -1), 0) + " USD"


# ─────────────────────────────── UTILITY FUNCTIONS ────────────────────────────

def fix_cover(doc):
    """Sửa tiêu đề và tên môn học trên trang bìa (giữ nguyên định dạng run)."""
    for p in doc.paragraphs:
        text = p.text.strip()
        if text in COVER_REPLACEMENTS and p.runs:
            p.runs[0].text = COVER_REPLACEMENTS[text]
            for r in p.runs[1:]:
                r.text = ""

def set_row_flags(row, header=False):
    """Không cho dòng bị tách qua 2 trang; dòng tiêu đề được lặp lại ở đầu trang mới."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    if header:
        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

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
    """Một dòng ghi liên kết Git của bài tập."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(8)
    rb = p.add_run("Link Git: ")
    rb.font.name = "Times New Roman"
    rb.font.size = Pt(12)
    rb.font.bold = True
    add_hyperlink(p, repo_url, repo_url, color="0056B3", underline=True, bold=False)

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
    set_row_flags(table.rows[0], header=True)
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
        set_row_flags(table.rows[start_row + ri])
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
    res = load_results()
    cv, rank, sweep, test, var = res["cv"], res["rank"], res["sweep"], res["test"], res["variance"]

    k_pca = int((var["Cumulative_Ratio"] >= 0.90).idxmax()) + 1
    k_95  = int((var["Cumulative_Ratio"] >= 0.95).idxmax()) + 1
    cum_k = var["Cumulative_Ratio"].iloc[k_pca - 1] * 100

    def cvrow(model, use_pca):
        sub = cv[(cv["Model"] == model) & (cv["PCA"].str.startswith("Yes" if use_pca else "No"))]
        return sub.iloc[0]

    best = rank.iloc[0]
    best_name, best_pca = best["Model"], best["PCA"].startswith("Yes")
    best_cv = cvrow(best_name, best_pca)
    best_cfg = f"có PCA ({k_pca}D)" if best_pca else "không dùng PCA"
    test_best  = test[test["Role"] == "Mô hình được chọn"].iloc[0]
    test_alt   = test[test["Role"] == "Đối chứng PCA"].iloc[0]
    test_dummy = test[test["Role"] == "Baseline"].iloc[0]

    print("[*] Mở bìa tiểu luận làm tài liệu gốc...")
    doc = Document(BIA_PATH)
    fix_cover(doc)

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
        "gồm sáu điểm:",
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
         "độ ổn định qua kỹ thuật 5-Fold Cross-Validation trên tập huấn luyện, và lựa "
         "chọn mô hình theo một quy tắc xếp hạng tường minh dựa trên các chỉ số này."),
        ("(5)",
         "Kiểm chứng mô hình được chọn trên một tập Test hold-out (20% dữ liệu) được "
         "tách riêng ngay từ đầu và không tham gia vào bất kỳ bước phân tích, so sánh "
         "hay lựa chọn nào."),
        ("(6)",
         f"Phân tích định lượng tác động thực tế của PCA: so sánh hiệu năng khi giảm "
         f"chiều từ 8 xuống {k_pca} và khảo sát sự biến thiên khi số thành phần PCA "
         f"thay đổi từ k = 1 đến k = 8."),
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
    add_p(doc,
        "Ngoài ra, ngay từ đầu dữ liệu được chia thành tập Train (80%, 16.512 mẫu) và "
        "tập Test hold-out (20%, 4.128 mẫu) với random_state = 42. Toàn bộ các bước phân "
        "tích PCA, so sánh mô hình bằng Cross-Validation, khảo sát số chiều và lựa chọn "
        "mô hình chỉ sử dụng tập Train. Tập Test chỉ được dùng đúng một lần ở bước đánh "
        "giá cuối cùng, nhờ đó kết quả trên tập Test là ước lượng khách quan về khả năng "
        "tổng quát hóa của mô hình được chọn trên dữ liệu mới.",
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
    d2 = []
    for i, r in var.iterrows():
        label = r["PC"] + ("  ★" if i + 1 == k_pca else "")
        cum = f"{vn(r['Cumulative_Ratio'] * 100, 2)} %"
        if i + 1 == k_pca:
            cum += "  ← Ngưỡng ≥ 90 %"
        if i + 1 == k_95:
            cum += "  ← Ngưỡng ≥ 95 %"
        d2.append([label, vn(r["Eigenvalue"], 3),
                   f"{vn(r['Variance_Ratio'] * 100, 2)} %", cum])
    fill_data_rows(t2, d2, w2, highlight_row=k_pca - 1, highlight_color="D6EAF8", font_size=10)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    tail2 = var["Variance_Ratio"].iloc[-2:].sum() * 100
    add_p(doc,
        f"PCA được khớp trên tập Train đã chuẩn hóa. Kết quả cho thấy chỉ cần {k_pca} "
        f"thành phần chính đầu tiên đã bảo toàn được {vn(cum_k, 2)}% tổng phương sai của "
        f"các đặc trưng đầu vào, và cần {k_95} thành phần để vượt ngưỡng 95%. Hai thành "
        f"phần cuối cùng PC7 và PC8 chỉ chiếm {vn(tail2, 2)}% phương sai, chủ yếu phản ánh "
        f"sự dư thừa tuyến tính giữa các cặp biến tương quan cao. Dựa vào kết quả này, "
        f"k = {k_pca} (số thành phần nhỏ nhất đạt ngưỡng 90%) được chọn làm cấu hình PCA "
        f"chuẩn trong các thực nghiệm so sánh. Lưu ý rằng phương sai ở đây là phương sai "
        f"của đặc trưng đầu vào X, không phải lượng thông tin hữu ích để dự báo biến mục "
        f"tiêu y — điểm này được phân tích ở Chương 6.",
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
         "Mô hình học kết hợp theo kỹ thuật Bagging — huấn luyện 70 cây quyết định "
         "(max_depth = 12) trên các mẫu bootstrap khác nhau của tập huấn luyện, sau đó lấy "
         "trung bình kết quả để giảm phương sai và chống quá khớp. Với cấu hình mặc định "
         "cho hồi quy của scikit-learn (max_features = 1,0), mỗi lần tách nút xét toàn bộ "
         "đặc trưng."),
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

    add_heading(doc, "4.2. Phương pháp đánh giá: Hold-out Test + 5-Fold Cross-Validation", level=2)
    add_p(doc,
        "Quy trình đánh giá gồm hai tầng. Tầng thứ nhất: 20.640 mẫu được chia ngẫu nhiên "
        "thành tập Train 16.512 mẫu (80%) và tập Test hold-out 4.128 mẫu (20%). Tầng thứ "
        "hai: trên tập Train, mỗi pipeline được đánh giá bằng 5-Fold Cross-Validation có "
        "xáo trộn — tại mỗi vòng, khoảng 13.210 mẫu dùng để huấn luyện và khoảng 3.302 "
        "mẫu dùng để kiểm định. Trung bình và độ lệch chuẩn qua 5 fold được dùng để so "
        "sánh và lựa chọn mô hình. Sau khi chọn xong, mô hình được huấn luyện lại trên "
        "toàn bộ tập Train và đánh giá một lần duy nhất trên tập Test. Mọi phép chia đều "
        "dùng random_state = 42 để đảm bảo tính tái lập.",
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

    add_heading(doc, "4.4. Quy tắc lựa chọn mô hình", level=2)
    add_p(doc,
        "Để việc chọn mô hình khách quan và tái lập được, chương trình xếp hạng tự động "
        "tất cả các cấu hình (8 mô hình × có/không PCA, không tính Dummy Baseline) theo "
        "bốn tiêu chí trên kết quả 5-Fold CV: MAE trung bình (thấp tốt), RMSE trung bình "
        "(thấp tốt), R² trung bình (cao tốt) và độ lệch chuẩn của R² qua các fold (thấp = "
        "ổn định hơn). Mô hình được chọn là cấu hình có hạng trung bình của bốn tiêu chí "
        "nhỏ nhất; nếu bằng nhau thì ưu tiên RMSE thấp hơn.",
        indent=True
    )

    add_separator(doc)

    # ── Bước 8: CHƯƠNG 5 — Kết quả ──────────────────────────────────────────
    add_heading(doc, "5. KẾT QUẢ THỰC NGHIỆM CHI TIẾT", level=1)

    add_heading(doc, "5.1. Bảng tổng hợp số liệu (5-Fold CV trên tập Train: Mean ± Std)", level=2)
    add_p(doc, "Bảng 3: Kết quả đánh giá tất cả các mô hình qua 5-Fold Cross-Validation trên tập Train",
          italic=True, space_after=3)
    t3 = doc.add_table(rows=len(cv) + 1, cols=6)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    h3 = ["Mô hình", "Cấu hình PCA", "MAE (Mean ± Std)", "RMSE (Mean ± Std)", "R² (Mean ± Std)", "T.gian (s)"]
    w3 = [Inches(1.75), Inches(1.15), Inches(1.35), Inches(1.35), Inches(1.35), Inches(0.8)]
    make_header_row(t3, h3, w3, font_size=9.5)
    d3, best_idx = [], None
    for i, r in cv.reset_index(drop=True).iterrows():
        use_pca = r["PCA"].startswith("Yes")
        if r["Model"] == best_name and use_pca == best_pca:
            best_idx = i
        d3.append([
            DISPLAY_NAMES.get(r["Model"], r["Model"]),
            f"Có PCA ({k_pca}D)" if use_pca else "Không PCA (8D)",
            pm(r["MAE_Mean"], r["MAE_Std"]),
            pm(r["RMSE_Mean"], r["RMSE_Std"]),
            pm(r["R2_Mean"], r["R2_Std"]),
            vn(r["Fit_Time_Sec"], 2),
        ])
    col_a3 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER,
               WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
               WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
    fill_data_rows(t3, d3, w3, highlight_row=best_idx, highlight_color="D5F5E3",
                   col_aligns=col_a3, font_size=9)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_heading(doc, "5.2. Xếp hạng và lựa chọn mô hình", level=2)
    add_p(doc, "Bảng 4: Xếp hạng các cấu hình theo quy tắc ở mục 4.4 (6 cấu hình đứng đầu)",
          italic=True, space_after=3)
    top = rank.head(6)
    t4 = doc.add_table(rows=len(top) + 1, cols=7)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    h4 = ["Mô hình", "Cấu hình PCA", "Hạng MAE", "Hạng RMSE", "Hạng R²", "Hạng ổn định", "Hạng TB"]
    w4 = [Inches(1.75), Inches(1.15), Inches(0.75), Inches(0.8), Inches(0.75), Inches(0.9), Inches(0.75)]
    make_header_row(t4, h4, w4, font_size=9.5)
    d4 = [[DISPLAY_NAMES.get(r["Model"], r["Model"]),
           f"Có PCA ({k_pca}D)" if r["PCA"].startswith("Yes") else "Không PCA (8D)",
           f"{r['Rank_MAE']:.0f}", f"{r['Rank_RMSE']:.0f}", f"{r['Rank_R2']:.0f}",
           f"{r['Rank_Stability']:.0f}", vn(r["Avg_Rank"], 2)] for _, r in top.iterrows()]
    col_a4 = [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 6
    fill_data_rows(t4, d4, w4, highlight_row=0, highlight_color="D5F5E3",
                   col_aligns=col_a4, font_size=9)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    rf = cvrow("Random Forest", False)
    svr = cvrow("Support Vector Regressor", False)
    knn = cvrow("K-Nearest Neighbors", False)
    dt = cvrow("Decision Tree", False)
    lin = [cvrow(m, False)["R2_Mean"] for m in ["Linear Regression", "Ridge Regression", "Lasso Regression"]]
    dummy = cvrow("Dummy Baseline", False)
    mae_pct = best_cv["MAE_Mean"] / TARGET_MEAN * 100
    n_full_rank1 = int((best[["Rank_MAE", "Rank_RMSE", "Rank_R2", "Rank_Stability"]] == 1).sum())

    add_p(doc,
        f"Theo quy tắc xếp hạng, {best_name} {best_cfg} được chọn với hạng trung bình "
        f"{vn(best['Avg_Rank'], 2)} — đứng đầu ở {n_full_rank1}/4 tiêu chí. Mô hình đạt "
        f"R² = {vn(best_cv['R2_Mean'])}, RMSE = {vn(best_cv['RMSE_Mean'])} và MAE = "
        f"{vn(best_cv['MAE_Mean'])}, với độ lệch chuẩn R² qua 5 fold chỉ "
        f"{vn(best_cv['R2_Std'])}. Quy đổi ra thực tế, sai số tuyệt đối trung bình khoảng "
        f"{usd(best_cv['MAE_Mean'])} so với giá trị nhà trung bình khoảng {usd(TARGET_MEAN)}, "
        f"tức khoảng {vn(mae_pct, 1)}%. Thời gian huấn luyện trung bình "
        f"{vn(best_cv['Fit_Time_Sec'], 2)} giây/fold, nhanh hơn Random Forest khoảng "
        f"{vn(rf['Fit_Time_Sec'] / best_cv['Fit_Time_Sec'], 1)} lần và SVR khoảng "
        f"{vn(svr['Fit_Time_Sec'] / best_cv['Fit_Time_Sec'], 0)} lần.",
        indent=True
    )
    add_p(doc,
        f"Các mô hình còn lại (không PCA) phân thành các nhóm hiệu năng rõ rệt. Random "
        f"Forest đứng thứ hai với R² = {vn(rf['R2_Mean'])}. SVR nhân RBF đạt R² = "
        f"{vn(svr['R2_Mean'])}, học tốt quan hệ phi tuyến nhưng chi phí tính toán cao nhất "
        f"({vn(svr['Fit_Time_Sec'], 2)} giây/fold). KNN và Decision Tree đạt R² lần lượt "
        f"{vn(knn['R2_Mean'])} và {vn(dt['R2_Mean'])}. Ba mô hình tuyến tính Linear "
        f"Regression, Ridge và Lasso gần như tương đương nhau với R² từ {vn(min(lin))} đến "
        f"{vn(max(lin))}, cho thấy quan hệ phi tuyến trong dữ liệu là đáng kể. Dummy "
        f"Baseline cho R² = {vn(dummy['R2_Mean'])} ≈ 0 — tất cả 8 mô hình học máy đều vượt "
        f"trội rõ rệt so với mốc cơ sở.",
        indent=True
    )

    add_heading(doc, "5.3. Biểu đồ so sánh trực quan các mô hình", level=2)
    add_figure(doc, "results/model_comparison.png",
               "Hình 3: So sánh 9 mô hình hồi quy không dùng PCA qua ba chỉ số R², RMSE "
               "và MAE, kèm thanh sai số chuẩn của 5-Fold CV trên tập Train",
               width=6.3)

    add_heading(doc, "5.4. Đánh giá cuối cùng trên tập Test hold-out", level=2)
    add_p(doc,
        "Sau khi lựa chọn, mô hình được huấn luyện lại trên toàn bộ 16.512 mẫu của tập "
        "Train và đánh giá một lần duy nhất trên 4.128 mẫu của tập Test. Bảng 5 so sánh "
        "mô hình được chọn với Dummy Baseline và với chính nó ở cấu hình PCA còn lại.",
        indent=True
    )
    add_p(doc, "Bảng 5: Kết quả trên tập Test hold-out (20%)", italic=True, space_after=3)
    t5 = doc.add_table(rows=len(test) + 1, cols=6)
    t5.alignment = WD_TABLE_ALIGNMENT.CENTER
    h5 = ["Vai trò", "Mô hình", "Cấu hình PCA", "MAE", "RMSE", "R²"]
    w5 = [Inches(1.4), Inches(1.75), Inches(1.15), Inches(0.9), Inches(0.9), Inches(0.9)]
    make_header_row(t5, h5, w5, font_size=9.5)
    d5 = [[r["Role"], DISPLAY_NAMES.get(r["Model"], r["Model"]),
           f"Có PCA ({k_pca}D)" if r["PCA"].startswith("Yes") else "Không PCA (8D)",
           vn(r["MAE"]), vn(r["RMSE"]), vn(r["R2"])] for _, r in test.iterrows()]
    best_test_idx = int(test.index[test["Role"] == "Mô hình được chọn"][0])
    col_a5 = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 4
    fill_data_rows(t5, d5, w5, highlight_row=best_test_idx, highlight_color="D5F5E3",
                   col_aligns=col_a5, font_size=9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    gap_r2 = test_best["R2"] - best_cv["R2_Mean"]
    add_p(doc,
        f"Trên tập Test, {best_name} {best_cfg} đạt R² = {vn(test_best['R2'])}, RMSE = "
        f"{vn(test_best['RMSE'])} và MAE = {vn(test_best['MAE'])}. Chênh lệch R² giữa tập "
        f"Test và trung bình CV là {vn(gap_r2)} (độ lệch chuẩn CV là "
        f"{vn(best_cv['R2_Std'])}), cho thấy ước lượng từ Cross-Validation đáng tin cậy và "
        f"mô hình không bị quá khớp vào quá trình lựa chọn. Dummy Baseline trên tập Test "
        f"chỉ đạt R² = {vn(test_dummy['R2'])}. Cấu hình đối chứng "
        f"({'không PCA' if best_pca else f'có PCA {k_pca}D'}) đạt R² = {vn(test_alt['R2'])}, "
        f"xác nhận kết luận về tác động của PCA ở Chương 6 cũng đúng trên dữ liệu mới.",
        indent=True
    )

    add_figure(doc, "results/actual_vs_predicted.png",
               f"Hình 4: Giá trị thực tế và giá trị dự báo trên tập Test — Dummy Baseline "
               f"(trái) và {best_name} {best_cfg} (phải)",
               width=6.2)
    add_p(doc,
        f"Biểu đồ phân tán làm rõ sự khác biệt định tính giữa hai cực hiệu năng. Với "
        f"Dummy Baseline, toàn bộ điểm dự báo nằm trên một đường ngang tương ứng với giá "
        f"trị trung bình của tập Train — mô hình không học được cấu trúc nào. Ngược lại, "
        f"các dự báo của {best_name} tập trung quanh đường chéo lý tưởng trên toàn bộ dải "
        f"giá. Ở mức trần 500.001 USD, nhiều mẫu thực tế bằng trần nhưng được dự báo "
        f"thấp hơn, do hiệu ứng cắt ngưỡng của dữ liệu gốc.",
        indent=True
    )

    add_separator(doc)

    # ── Bước 9: CHƯƠNG 6 — Phân tích PCA ────────────────────────────────────
    add_heading(doc, "6. PHÂN TÍCH CHUYÊN SÂU TÁC ĐỘNG CỦA PCA", level=1)

    add_heading(doc, f"6.1. So sánh trực tiếp: Có PCA ({k_pca}D) vs Không PCA (8D)", level=2)
    add_figure(doc, "results/pca_impact.png",
               f"Hình 5: So sánh R² và RMSE của từng mô hình giữa hai cấu hình "
               f"không dùng PCA (8 chiều) và có dùng PCA ({k_pca} chiều)",
               width=6.2)

    impact = []
    for m in cv["Model"].unique():
        if m == "Dummy Baseline":
            continue
        a, b = cvrow(m, False), cvrow(m, True)
        impact.append({
            "Model": m,
            "r2_no": a["R2_Mean"], "r2_yes": b["R2_Mean"],
            "r2_drop": (a["R2_Mean"] - b["R2_Mean"]) / a["R2_Mean"] * 100,
            "rmse_up": (b["RMSE_Mean"] - a["RMSE_Mean"]) / a["RMSE_Mean"] * 100,
        })
    impact = pd.DataFrame(impact).set_index("Model")
    all_worse = bool((impact["r2_drop"] > 0).all())

    def drop_txt(m):
        r = impact.loc[m]
        return (f"{DISPLAY_NAMES.get(m, m)} từ {vn(r['r2_no'])} xuống {vn(r['r2_yes'])} "
                f"(giảm {vn(r['r2_drop'], 1)}%)")

    add_p(doc,
        f"Khi giảm từ 8 chiều xuống {k_pca} chiều — dù vẫn giữ lại {vn(cum_k, 2)}% phương "
        f"sai đầu vào — hiệu năng của "
        f"{'tất cả các mô hình đều' if all_worse else 'phần lớn các mô hình'} suy giảm rõ "
        f"rệt. R² giảm tương đối từ {vn(impact['r2_drop'].min(), 1)}% "
        f"({DISPLAY_NAMES.get(impact['r2_drop'].idxmin(), impact['r2_drop'].idxmin())}) đến "
        f"{vn(impact['r2_drop'].max(), 1)}% "
        f"({DISPLAY_NAMES.get(impact['r2_drop'].idxmax(), impact['r2_drop'].idxmax())}). "
        f"Cụ thể: R² của {drop_txt('HistGradientBoosting')}; {drop_txt('Random Forest')}; "
        f"{drop_txt('Decision Tree')}; {drop_txt('Linear Regression')}. Tương ứng, RMSE "
        f"tăng từ {vn(impact['rmse_up'].min(), 1)}% đến {vn(impact['rmse_up'].max(), 1)}% "
        f"so với cấu hình không dùng PCA; mức tăng lớn nhất rơi vào "
        f"{' và '.join(DISPLAY_NAMES.get(m, m) for m in impact['rmse_up'].nlargest(2).index)} "
        f"— chính là những mô hình mạnh nhất khi không dùng PCA.",
        indent=True
    )

    add_heading(doc, "6.2. Khảo sát hiệu năng theo số chiều PCA từ k = 1 đến k = 8", level=2)
    add_figure(doc, "results/pca_components_sweep.png",
               "Hình 6: Sự biến thiên của R² và RMSE (5-Fold CV trên tập Train) khi số thành "
               "phần PCA k tăng từ 1 đến 8 trên ba mô hình tiêu biểu",
               width=6.2)

    def sw(m, k, col="R2_Mean"):
        return sweep[(sweep["Model"] == m) & (sweep["k_components"] == k)][col].iloc[0]

    lr_k = {k: sw("Linear Regression", k) for k in range(1, 9)}
    tree_models = ["Random Forest", "HistGradientBoosting"]
    lr_unstable = sweep[(sweep["Model"] == "Linear Regression")].set_index("k_components")["R2_Std"]
    k_unstable = int(lr_unstable.idxmax())

    tree_k1 = min(sw(m, 1) for m in tree_models)
    tree_k3 = max(sw(m, 3) for m in tree_models)
    tree_k4 = sorted(sw(m, 4) for m in tree_models)
    lr_std_other = lr_unstable.drop(k_unstable).max()
    add_p(doc,
        f"Khảo sát dùng đúng siêu tham số như thực nghiệm chính. Với k = 1 đến 3, hai mô "
        f"hình cây chỉ đạt R² từ khoảng {vn(tree_k1, 2)} đến {vn(tree_k3, 2)}, còn Linear "
        f"Regression gần như không học được gì (R² ≤ {vn(max(lr_k[1], lr_k[2], lr_k[3]), 2)}) "
        f"— các thành phần chính đầu tiên chứa nhiều phương sai của X nhưng ít liên quan "
        f"tuyến tính tới giá nhà. Tại k = 4, cả ba mô hình đều tăng vọt: Random Forest và "
        f"HistGradientBoosting lên khoảng {vn(tree_k4[0], 2)}–{vn(tree_k4[1], 2)}, Linear "
        f"Regression lên {vn(lr_k[4], 2)}. Riêng Linear Regression tại k = {k_unstable} có độ "
        f"lệch chuẩn R² qua các fold bằng {vn(lr_unstable.max(), 2)}, lớn hơn hẳn các giá "
        f"trị k khác (≤ {vn(lr_std_other, 2)}). Kiểm tra riêng cho thấy nguyên nhân là một "
        f"mẫu ngoại lai cực đoan (AveOccup ≈ 1.243 người/hộ) bị chiếu thành giá trị rất lớn "
        f"trên thành phần chính, khiến dự báo tuyến tính cho mẫu này sai lệch hơn 30 đơn vị "
        f"và kéo R² của fold đó xuống thấp. Từ k = 5 đến k = 8, cả ba mô hình cải thiện "
        f"dần đều.",
        indent=True
    )
    add_p(doc,
        f"Điểm đáng chú ý nhất là tại k = 8 (giữ lại toàn bộ 8 thành phần, PCA chỉ còn là "
        f"một phép quay trực giao, không mất thông tin): Linear Regression đạt R² = "
        f"{vn(lr_k[8])}, trùng với kết quả không dùng PCA ({vn(cvrow('Linear Regression', False)['R2_Mean'])}) "
        f"vì mô hình tuyến tính bất biến với phép quay. Ngược lại, Random Forest chỉ đạt "
        f"{vn(sw('Random Forest', 8))} (so với {vn(cvrow('Random Forest', False)['R2_Mean'])} "
        f"khi không PCA) và HistGradientBoosting đạt {vn(sw('HistGradientBoosting', 8))} (so "
        f"với {vn(cvrow('HistGradientBoosting', False)['R2_Mean'])}). Như vậy, với mô hình "
        f"cây, PCA gây hại ngay cả khi không bỏ bớt chiều nào — nguyên nhân không chỉ là "
        f"mất thông tin mà còn là phép quay phá vỡ cấu trúc đặc trưng gốc.",
        indent=True
    )

    add_heading(doc, "6.3. Giải thích nguyên nhân PCA làm giảm hiệu năng", level=2)
    add_p(doc,
        f"Câu hỏi đặt ra là tại sao PCA giữ lại tới {vn(cum_k, 2)}% phương sai đầu vào "
        f"nhưng R² vẫn sụt giảm từ {vn(impact['r2_drop'].min(), 1)}% đến "
        f"{vn(impact['r2_drop'].max(), 1)}%? Có ba nguyên nhân chính:",
        indent=True
    )
    add_callout(doc, [
        "Nguyên nhân 1 — PCA là giải thuật học không giám sát:\n"
        "PCA chọn các hướng có phương sai lớn nhất của X mà hoàn toàn không biết biến "
        f"mục tiêu y. Các thành phần bị loại bỏ (PC{k_pca + 1}–PC8), dù chỉ chiếm khoảng "
        f"{vn((1 - var['Cumulative_Ratio'].iloc[k_pca - 1]) * 100, 1)}% phương sai của X, vẫn "
        "có thể mang thông tin quan trọng để dự báo giá nhà. Hình 6 cho thấy R² tiếp tục "
        "tăng rõ rệt từ k = 5 đến k = 8, tức các thành phần cuối không phải là nhiễu.",
        "\nNguyên nhân 2 — Phép quay phá vỡ cấu trúc mà mô hình cây khai thác:\n"
        "Cây quyết định tách dữ liệu theo từng đặc trưng gốc (ngưỡng song song trục). Các "
        "biến như MedInc hay cặp Latitude/Longitude có quan hệ phi tuyến mạnh với giá nhà "
        "(các vùng ven biển có giá cao). Mỗi thành phần chính là tổ hợp tuyến tính trộn lẫn "
        "các biến kinh tế và tọa độ, nên cây cần nhiều lần tách hơn để mô tả cùng một ranh "
        "giới. Bằng chứng trực tiếp là kết quả tại k = 8 ở mục 6.2: không mất thông tin nào "
        "nhưng R² của Random Forest và HistGradientBoosting vẫn giảm.",
        "\nNguyên nhân 3 — Số chiều ban đầu đã nhỏ:\n"
        "PCA phát huy hiệu quả nhất khi số chiều rất lớn như ảnh, văn bản hay dữ liệu gen, "
        "nơi giảm chiều giúp chống quá khớp và giảm chi phí. Với chỉ 8 đặc trưng và hơn "
        "16.000 mẫu huấn luyện, không tồn tại vấn đề bùng nổ số chiều, nên việc giảm chiều "
        "không mang lại lợi ích mà chỉ làm mất thông tin.",
    ], title="Ba nguyên nhân vì sao PCA làm giảm hiệu năng:")

    add_separator(doc)

    # ── Bước 10: CHƯƠNG 7 — Kết luận ────────────────────────────────────────
    add_heading(doc, "7. KẾT LUẬN VÀ KIẾN NGHỊ MÔ HÌNH", level=1)

    add_p(doc, "Từ toàn bộ quá trình thực nghiệm, bài tập rút ra ba kết luận chính:",
          indent=True)

    add_p(doc,
        f"Về lựa chọn mô hình, theo quy tắc xếp hạng dựa trên MAE, RMSE, R² và độ ổn định "
        f"qua 5-Fold CV, {best_name} {best_cfg} là mô hình tốt nhất trên bộ dữ liệu "
        f"California Housing. Trên CV, mô hình đạt R² = {vn(best_cv['R2_Mean'])} ± "
        f"{vn(best_cv['R2_Std'])}, RMSE = {vn(best_cv['RMSE_Mean'])}, MAE = "
        f"{vn(best_cv['MAE_Mean'])}; trên tập Test hold-out đạt R² = {vn(test_best['R2'])}, "
        f"RMSE = {vn(test_best['RMSE'])}, MAE = {vn(test_best['MAE'])} (sai số tuyệt đối "
        f"trung bình khoảng {usd(test_best['MAE'])}). Kết quả trên tập Test khớp với CV, "
        f"khẳng định mô hình ổn định và tổng quát hóa tốt; thời gian huấn luyện ngắn "
        f"({vn(best_cv['Fit_Time_Sec'], 2)} giây/fold) cũng thuận lợi cho triển khai.",
        bold_part="Kết luận 1.", indent=True
    )
    add_p(doc,
        "Về Pipeline tiền xử lý, việc kết hợp SimpleImputer và StandardScaler trong "
        "Pipeline, cùng với việc tách tập Test ngay từ đầu, đảm bảo không rò rỉ dữ liệu "
        "trong toàn bộ quá trình đánh giá. Chuẩn hóa là bắt buộc với các thuật toán nhạy "
        "cảm với thang đo như KNN, SVR, Ridge, Lasso và với PCA; bước Imputation giúp "
        "pipeline vẫn hoạt động khi dữ liệu mới có giá trị khuyết thiếu.",
        bold_part="Kết luận 2.", indent=True
    )
    add_p(doc,
        f"Về tác động của PCA, thực nghiệm cho thấy giảm chiều không giám sát bằng PCA "
        f"không đồng nghĩa với việc cải thiện bài toán hồi quy có giám sát. Trên California "
        f"Housing, PCA {k_pca}D làm R² giảm từ {vn(impact['r2_drop'].min(), 1)}% đến "
        f"{vn(impact['r2_drop'].max(), 1)}% trên CV, và kết quả này được xác nhận lại trên "
        f"tập Test. Với mô hình cây, PCA gây hại ngay cả khi giữ đủ 8 thành phần. PCA vẫn "
        f"có giá trị trong phân tích cấu trúc dữ liệu, trực quan hóa 2D/3D hoặc với dữ liệu "
        f"nhiều chiều; quyết định có dùng PCA hay không cần được kiểm chứng bằng "
        f"Cross-Validation như trong bài tập này.",
        bold_part="Kết luận 3.", indent=True
    )

    add_callout(doc,
        f"{best_name} {best_cfg} được kiến nghị là mô hình triển khai:\n"
        f"CV (Train):  R² = {vn(best_cv['R2_Mean'])} ± {vn(best_cv['R2_Std'])}  |  "
        f"RMSE = {vn(best_cv['RMSE_Mean'])}  |  MAE = {vn(best_cv['MAE_Mean'])}\n"
        f"Test hold-out:  R² = {vn(test_best['R2'])}  |  RMSE = {vn(test_best['RMSE'])}  |  "
        f"MAE = {vn(test_best['MAE'])}",
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
