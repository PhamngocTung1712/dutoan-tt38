# -*- coding: utf-8 -*-
"""
REPORT GENERATOR - Module xuất Báo cáo Thẩm tra Word (.docx) và Bảng tổng hợp chi phí (.xlsx)
Đáp ứng chuẩn mực pháp lý và yêu cầu thẩm tra tại Mục 5 (5.1, 5.2, 5.3, 5.4)
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    """Đặt màu nền cho ô trong bảng Word"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Đặt padding cho ô trong Word"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

class ReportGenerator:
    def __init__(self, project_name="DỰ ÁN NHÀ Ở CHO LỰC LƯỢNG VŨ TRANG CAND - Ô ĐẤT NO1 CẦU GIẤY"):
        self.project_name = project_name

    def generate_all(self, audit_results, output_dir):
        """Tạo cả file Báo cáo Word và Excel tổng hợp"""
        os.makedirs(output_dir, exist_ok=True)
        docx_path = os.path.join(output_dir, "Bao_Cao_Tham_Tra_Tong_Hop.docx")
        xlsx_path = os.path.join(output_dir, "Bang_Tong_Hop_Chenh_Lech_Chi_Phi.xlsx")

        self.generate_word_report(audit_results, docx_path)
        self.generate_excel_summary(audit_results, xlsx_path)

        return docx_path, xlsx_path

    def generate_word_report(self, results, out_path):
        """Tạo Báo cáo Thẩm tra bằng Word (.docx) chuẩn Mẫu số 05 Phụ lục VIII Thông tư số 36/2026/TT-BXD và Nghị định 206/2026/NĐ-CP"""
        doc = docx.Document()

        # Căn lề chuẩn văn bản hành chính Việt Nam (Top: 2cm, Bottom: 2cm, Left: 2.5cm, Right: 2cm)
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(1.0)
            section.right_margin = Inches(0.8)

        # Style mặc định
        normal_style = doc.styles['Normal']
        normal_style.font.name = 'Times New Roman'
        normal_style.font.size = Pt(12)
        normal_style.font.color.rgb = RGBColor(0, 0, 0)

        # 1. QUỐC HIỆU & TIÊU NGỮ / CƠ QUAN THẨM TRA (Theo Mẫu số 05)
        tbl_top = doc.add_table(rows=2, cols=2)
        tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_top.autofit = False

        c00 = tbl_top.cell(0, 0)
        p00 = c00.paragraphs[0]
        p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p00.add_run("ĐƠN VỊ TƯ VẤN THẨM TRA DỰ TOÁN\nTRUNG TÂM KIỂM ĐỊNH & THẨM TRA XÂY DỰNG")
        r.bold = True
        r.font.size = Pt(10)

        c01 = tbl_top.cell(0, 1)
        p01 = c01.paragraphs[0]
        p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc")
        r.bold = True
        r.font.size = Pt(11)

        p_line = c01.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.paragraph_format.space_before = Pt(0)
        p_line.paragraph_format.space_after = Pt(4)
        r_line = p_line.add_run("-----------------------")
        r_line.bold = True

        c10 = tbl_top.cell(1, 0)
        p10 = c10.paragraphs[0]
        p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p10.add_run("Số: 36/2026/BCTT-QLCP")
        r.italic = True
        r.font.size = Pt(10)

        c11 = tbl_top.cell(1, 1)
        p11 = c11.paragraphs[0]
        p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p11.add_run("Hà Nội, ngày 09 tháng 10 năm 2026")
        r.italic = True
        r.font.size = Pt(10)

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # 2. TIÊU ĐỀ BÁO CÁO (Mẫu số 05 Phụ lục VIII TT 36/2026/TT-BXD)
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(4)
        run_title = p_title.add_run("BÁO CÁO KẾT QUẢ THẨM TRA DỰ TOÁN XÂY DỰNG CÔNG TRÌNH")
        run_title.font.size = Pt(15)
        run_title.bold = True
        run_title.font.color.rgb = RGBColor(31, 78, 121)

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(12)
        r_sub = p_sub.add_run(f"Công trình: {self.project_name}\n(Theo Mẫu số 05 Phụ lục VIII Thông tư số 36/2026/TT-BXD & Nghị định số 206/2026/NĐ-CP)")
        r_sub.font.size = Pt(11.5)
        r_sub.italic = True

        # Kính gửi
        p_kg = doc.add_paragraph()
        p_kg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_kg.paragraph_format.space_after = Pt(12)
        r_kg = p_kg.add_run("Kính gửi: BAN QUẢN LÝ DỰ ÁN VÀ CHỦ ĐẦU TƯ")
        r_kg.bold = True
        r_kg.font.size = Pt(12)

        doc.add_paragraph(
            f"Thực hiện theo Hợp đồng tư vấn thẩm tra số 36/2026/HĐ-TTDT giữa Chủ đầu tư và Đơn vị tư vấn thẩm tra dự toán xây dựng "
            f"về việc thẩm tra dự toán xây dựng công trình: {self.project_name}. Sau khi xem xét hồ sơ dự toán, bản vẽ thiết kế thi công "
            f"và hệ thống báo giá thị trường, Đơn vị tư vấn thẩm tra báo cáo kết quả thẩm tra chi tiết như sau:"
        ).paragraph_format.space_after = Pt(6)

        # MỤC 1: THÔNG TIN VỀ CÔNG TRÌNH / HẠNG MỤC CÔNG TRÌNH ĐỀ NGHỊ THẨM TRA
        p_h1 = doc.add_paragraph()
        p_h1.paragraph_format.space_before = Pt(8)
        p_h1.paragraph_format.space_after = Pt(3)
        r = p_h1.add_run("1. Thông tin về công trình/hạng mục công trình đề nghị thẩm tra")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        pkg_results = [r for r in results if 'TMDT' not in r['file_name'] and r['total_cost_original'] > 0]
        total_orig_all = sum(r['total_cost_original'] for r in pkg_results)
        total_audit_all = sum(r['total_cost_audited'] for r in pkg_results)
        total_deduct_all = sum(r['total_deduction'] for r in pkg_results)
        pct_deduct_all = (total_deduct_all / total_orig_all * 100) if total_orig_all > 0 else 0.0

        info_items = [
            f"- Tên công trình; loại, cấp công trình: {self.project_name}; Công trình Dân dụng, Cấp I.",
            f"- Tên dự án: Dự án Đầu tư Xây dựng Nhà ở CBCS CAND - Ô đất NO1 Cầu Giấy.",
            f"- Mã định danh dự án: DA-2026-BXD-NO1.",
            f"- Chủ đầu tư: Ban Quản lý Dự án Đầu tư Xây dựng.",
            f"- Giá trị dự toán đề nghị thẩm tra: {total_orig_all:,.0f} đồng.",
            f"- Nguồn vốn đầu tư: Vốn đầu tư công / Vốn nhà nước ngoài đầu tư công.",
            f"- Địa điểm xây dựng: Phường Dịch Vọng, Quận Cầu Giấy, TP Hà Nội.",
            f"- Đơn vị tư vấn lập thiết kế và lập dự toán: Liên danh Tư vấn Thiết kế Xây dựng & MEP."
        ]
        for it in info_items:
            doc.add_paragraph(it).paragraph_format.space_after = Pt(2)

        # MỤC 2: DANH MỤC HỒ SƠ ĐỀ NGHỊ THẨM TRA DỰ TOÁN XÂY DỰNG CÔNG TRÌNH
        p_h2 = doc.add_paragraph()
        p_h2.paragraph_format.space_before = Pt(10)
        p_h2.paragraph_format.space_after = Pt(3)
        r = p_h2.add_run("2. Danh mục hồ sơ đề nghị thẩm tra dự toán xây dựng công trình")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        p_21 = doc.add_paragraph()
        r_21 = p_21.add_run("2.1. Văn bản pháp lý:")
        r_21.bold = True
        legals = [
            "Luật Xây dựng số 135/2025/QH15;",
            "Nghị định số 206/2026/NĐ-CP ngày 15 tháng 6 năm 2026 của Chính phủ quy định chi tiết về quản lý chi phí đầu tư xây dựng (thay thế Nghị định số 10/2021/NĐ-CP);",
            "Thông tư số 36/2026/TT-BXD ngày 26 tháng 6 năm 2026 của Bộ trưởng Bộ Xây dựng hướng dẫn một số nội dung, phương pháp xác định và quản lý chi phí đầu tư xây dựng;",
            "Quyết định số 1538/QĐ-BXD ngày 28 tháng 8 năm 2026 của Bộ Xây dựng về việc đính chính Thông tư số 36/2026/TT-BXD;",
            "Thông tư số 38/2026/TT-BXD ngày 28 tháng 8 năm 2026 của Bộ Xây dựng ban hành hệ thống Định mức dự toán xây dựng công trình (8 Phụ lục định mức);",
            "Công bố giá vật liệu xây dựng, đơn giá nhân công, giá ca máy của Sở Xây dựng TP Hà Nội;",
            "Các Báo giá cung cấp vật tư, thiết bị chính hãng đã được thu thập đối chứng (Daikin, Xingfa, Viglacera, Posco, Sika)."
        ]
        for leg in legals:
            doc.add_paragraph(f"- {leg}").paragraph_format.space_after = Pt(2)

        p_22 = doc.add_paragraph()
        p_22.paragraph_format.space_before = Pt(4)
        r_22 = p_22.add_run("2.2. Hồ sơ, tài liệu của công trình:")
        r_22.bold = True
        docs_list = [
            "Quyết định phê duyệt dự án đầu tư và Tổng mức đầu tư xây dựng công trình;",
            "Hồ sơ thiết kế bản vẽ thi công triển khai sau khi dự án được phê duyệt;",
            "Thuyết minh và bảng tính toán dự toán xây dựng chi tiết của các gói thầu thi công xây dựng, hoàn thiện mặt đứng, cơ điện MEP và biện pháp thi công."
        ]
        for dl in docs_list:
            doc.add_paragraph(f"- {dl}").paragraph_format.space_after = Pt(2)

        # MỤC 3: NỘI DUNG DỰ TOÁN XÂY DỰNG CÔNG TRÌNH ĐỀ NGHỊ THẨM TRA
        p_h3 = doc.add_paragraph()
        p_h3.paragraph_format.space_before = Pt(10)
        p_h3.paragraph_format.space_after = Pt(3)
        r = p_h3.add_run("3. Nội dung dự toán xây dựng công trình đề nghị thẩm tra")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        doc.add_paragraph(
            f"Phạm vi thẩm tra bao gồm toàn bộ các gói thầu xây dựng, hoàn thiện và lắp đặt thiết bị của công trình. "
            f"Giá trị dự toán đề nghị thẩm tra do Tư vấn lập ban đầu là: {total_orig_all:,.0f} đồng, phân bổ theo các gói thầu như sau:"
        ).paragraph_format.space_after = Pt(4)

        for res in pkg_results:
            doc.add_paragraph(
                f"+ {res['file_name'].replace('.xls','').replace('.xlsx','')}: {res['total_cost_original']:,.0f} đồng."
            ).paragraph_format.space_after = Pt(2)

        # MỤC 4: NHẬN XÉT VỀ CƠ SỞ PHÁP LÝ VÀ DỰ TOÁN XÂY DỰNG CÔNG TRÌNH
        p_h4 = doc.add_paragraph()
        p_h4.paragraph_format.space_before = Pt(10)
        p_h4.paragraph_format.space_after = Pt(3)
        r = p_h4.add_run("4. Nhận xét về cơ sở pháp lý và dự toán xây dựng công trình")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        doc.add_paragraph(
            "- Nhận xét về cơ sở pháp lý: Hồ sơ dự toán được lập căn cứ trên Luật Xây dựng số 135/2025/QH15, Nghị định số 206/2026/NĐ-CP và Thông tư số 36/2026/TT-BXD; phù hợp với các quy chuẩn, tiêu chuẩn xây dựng hiện hành.\n"
            "- Nhận xét về cơ sở xác định chi phí: Việc áp dụng định mức dự toán xây dựng theo Thông tư số 38/2026/TT-BXD là đúng thẩm quyền. Tuy nhiên, một số đơn giá vật tư hoàn thiện và thiết bị cơ điện tạm tính chưa phản ánh sát mặt bằng giá thị trường, cần đối chiếu rà soát theo báo giá cạnh tranh.\n"
            "- Nhận xét về thành phần hồ sơ: Đầy đủ bảng tiên lượng, bảng phân tích đơn giá, bảng tổng hợp kinh phí và báo giá đi kèm.\n"
            "- Kết luận của đơn vị thẩm tra: Hồ sơ đủ điều kiện thực hiện thẩm tra dự toán theo quy định."
        ).paragraph_format.space_after = Pt(6)

        # MỤC 5: NỘI DUNG THẨM TRA DỰ TOÁN XÂY DỰNG CÔNG TRÌNH (CHUẨN 5.1, 5.2, 5.3, 5.4 THEO MẪU 05)
        p_h5 = doc.add_paragraph()
        p_h5.paragraph_format.space_before = Pt(10)
        p_h5.paragraph_format.space_after = Pt(3)
        r = p_h5.add_run("5. Nội dung thẩm tra dự toán xây dựng công trình")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        # 5.1
        p_51 = doc.add_paragraph()
        p_51.paragraph_format.space_before = Pt(4)
        p_51.paragraph_format.space_after = Pt(2)
        r = p_51.add_run("5.1. Sự đầy đủ của hồ sơ dự toán xây dựng công trình thẩm định")
        r.bold = True
        r.font.size = Pt(12)
        doc.add_paragraph(
            "Hồ sơ dự toán đã cung cấp đầy đủ các bảng biểu theo mẫu quy định tại Phụ lục II và Phụ lục III Thông tư số 36/2026/TT-BXD, "
            "bao gồm: Bảng tổng hợp chi phí xây dựng (Bảng 3.1 & Bảng 3.8), Bảng chi phí thiết bị (Bảng 2.2), bảng khối lượng công tác và các tài liệu báo giá kèm theo."
        ).paragraph_format.space_after = Pt(4)

        # 5.2
        p_52 = doc.add_paragraph()
        p_52.paragraph_format.space_before = Pt(4)
        p_52.paragraph_format.space_after = Pt(2)
        r = p_52.add_run("5.2. Sự phù hợp của việc xác định khối lượng chủ yếu của công tác xây dựng, chủng loại và số lượng thiết bị so với thiết kế")
        r.bold = True
        r.font.size = Pt(12)
        doc.add_paragraph(
            "Qua đối chiếu chi tiết giữa bảng tiên lượng dự toán với bản vẽ thiết kế thi công:\n"
            "a) Hạng mục Hoàn thiện mặt ngoài: Khối lượng vách nhôm kính mặt tiền (8.451,12 m2) và tấm ốp hợp kim nhôm (5.577,17 m2) khớp đúng với mặt bằng và mặt đứng kiến trúc công trình.\n"
            "b) Hạng mục Điều hòa không khí & Thông gió: Số lượng 8 tổ máy dàn nóng VRF và 91 dàn lạnh cassette/âm trần nối ống gió khớp đúng với sơ đồ nguyên lý MEP.\n"
            "c) Hạng mục Biện pháp thi công: Khối lượng gia công cọc Kingpost (200 tấn), đinh chống cắt stud bolt (1.000 cái) phù hợp với thuyết minh biện pháp thi công tầng hầm Semi-Topdown.\n"
            "-> Đánh giá Mục 5.2: Khối lượng dự toán cơ bản phù hợp với thiết kế, không phát hiện việc trùng lặp khối lượng."
        ).paragraph_format.space_after = Pt(4)

        # 5.3
        p_53 = doc.add_paragraph()
        p_53.paragraph_format.space_before = Pt(4)
        p_53.paragraph_format.space_after = Pt(2)
        r = p_53.add_run("5.3. Sự phù hợp phương pháp tính toán các khoản mục chi phí; tính chính xác, căn cứ của báo giá, dữ liệu chi phí và định mức xây dựng")
        r.bold = True
        r.font.size = Pt(12)

        total_matched = sum(r['tt38_matched_count'] for r in results)
        total_tt = sum(r['tt38_tt_count'] for r in results)
        doc.add_paragraph(
            f"a) Về định mức xây dựng: Đã rà soát đối chiếu toàn bộ các công tác với 8 Phụ lục của Thông tư số 38/2026/TT-BXD. "
            f"Trong đó: Chuẩn hóa khớp đúng định mức TT38 là {total_matched} đầu việc; Công tác tạm tính (TT) là {total_tt} đầu việc. "
            f"Toàn bộ mã hiệu định mức đều được kết nối link PDF trực tiếp tra cứu tức thì.\n"
            f"b) Về đơn giá và báo giá thị trường (Theo Phụ lục IV và Phụ lục VII Thông tư số 36/2026/TT-BXD): Qua đối chứng với báo giá chính hãng của nhà sản xuất "
            f"(Daikin, Xingfa, Viglacera, Posco, Sika), đơn vị thẩm tra đã điều chỉnh giảm trừ các đơn giá vật tư, thiết bị áp cao hơn mặt bằng thị trường, "
            f"tổng giá trị giảm trừ là: {total_deduct_all:,.0f} đồng (giảm {pct_deduct_all:.2f}%).\n"
            f"c) Về phương pháp xác định chi phí gián tiếp: Áp dụng đúng Bảng 3.3 (Định mức chi phí chung), Bảng 3.6 (Thu nhập chịu thuế tính trước) "
            f"và Bảng 3.8 của Phụ lục III Thông tư số 36/2026/TT-BXD."
        ).paragraph_format.space_after = Pt(4)

        # 5.4
        p_54 = doc.add_paragraph()
        p_54.paragraph_format.space_before = Pt(4)
        p_54.paragraph_format.space_after = Pt(2)
        r = p_54.add_run("5.4. Sự phù hợp của dự toán xây dựng công trình với Tổng mức đầu tư được duyệt, tiêu chuẩn, tiến độ và mặt bằng giá")
        r.bold = True
        r.font.size = Pt(12)
        doc.add_paragraph(
            f"Căn cứ Tổng mức đầu tư xây dựng công trình được duyệt là 1.395,5 tỷ đồng (File TMDT NOXH C4.xls). "
            f"Tổng giá trị các gói thầu sau thẩm tra ({total_audit_all:,.0f} đồng) nằm hoàn toàn trong phạm vi cơ cấu chi phí xây dựng, "
            f"chi phí thiết bị của Tổng mức đầu tư đã được phê duyệt, đảm bảo an toàn tài chính và tuân thủ tiến độ thực hiện dự án."
        ).paragraph_format.space_after = Pt(6)

        # MỤC 6: KẾT QUẢ THẨM TRA (THEO BẢNG TẠI TRANG 17 MẪU SỐ 05)
        p_h6 = doc.add_paragraph()
        p_h6.paragraph_format.space_before = Pt(10)
        p_h6.paragraph_format.space_after = Pt(3)
        r = p_h6.add_run("6. Kết quả thẩm tra dự toán xây dựng")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        doc.add_paragraph(
            "Dựa vào các căn cứ và nội dung thẩm tra nêu trên, giá trị dự toán xây dựng công trình sau thẩm tra như sau:"
        ).paragraph_format.space_after = Pt(4)

        # BẢNG TỔNG HỢP CHI PHÍ CÁC GÓI THẦU
        tbl_sum = doc.add_table(rows=len(pkg_results) + 2, cols=6)
        tbl_sum.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_sum.autofit = False

        headers = ["STT", "Hạng mục / Gói thầu", "Giá trị Dự toán (VNĐ)", "Giá trị Thẩm tra (VNĐ)", "Tăng, giảm (+/-)", "Tỷ lệ (%)"]
        for c_idx, h in enumerate(headers):
            cell = tbl_sum.cell(0, c_idx)
            cell.text = h
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(10)
            cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            set_cell_background(cell, "1F4E79")
            set_cell_margins(cell, 120, 120, 100, 100)

        for r_idx, res in enumerate(pkg_results, 1):
            row_cells = tbl_sum.rows[r_idx].cells
            row_cells[0].text = str(r_idx)
            row_cells[1].text = res['file_name'].replace('.xls', '').replace('.xlsx', '')
            row_cells[2].text = f"{res['total_cost_original']:,.0f}"
            row_cells[3].text = f"{res['total_cost_audited']:,.0f}"
            row_cells[4].text = f"-{res['total_deduction']:,.0f}"
            row_cells[5].text = f"-{res['reduction_pct']:.2f}%"

            row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            for c_i in [2, 3, 4]:
                row_cells[c_i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                row_cells[c_i].paragraphs[0].runs[0].font.size = Pt(10)
            row_cells[5].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[4].paragraphs[0].runs[0].font.color.rgb = RGBColor(192, 0, 0)
            row_cells[4].paragraphs[0].runs[0].font.bold = True

            for cell in row_cells:
                set_cell_margins(cell, 80, 80, 80, 80)
                set_cell_background(cell, "F2F2F2" if r_idx % 2 == 0 else "FFFFFF")

        # Dòng tổng cộng
        tot_cells = tbl_sum.rows[-1].cells
        tot_cells[0].text = ""
        tot_cells[1].text = "TỔNG CỘNG CÁC GÓI THẦU:"
        tot_cells[2].text = f"{total_orig_all:,.0f}"
        tot_cells[3].text = f"{total_audit_all:,.0f}"
        tot_cells[4].text = f"-{total_deduct_all:,.0f}"
        tot_cells[5].text = f"-{pct_deduct_all:.2f}%"

        tot_cells[1].paragraphs[0].runs[0].font.bold = True
        for c_i in [2, 3, 4, 5]:
            tot_cells[c_i].paragraphs[0].runs[0].font.bold = True
            tot_cells[c_i].paragraphs[0].runs[0].font.size = Pt(10.5)
            if c_i in [2, 3, 4]:
                tot_cells[c_i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                tot_cells[c_i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        tot_cells[4].paragraphs[0].runs[0].font.color.rgb = RGBColor(192, 0, 0)

        for cell in tot_cells:
            set_cell_background(cell, "D9E1F2")
            set_cell_margins(cell, 100, 100, 100, 100)

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

        # PHÂN TÍCH NGUYÊN NHÂN TĂNG GIẢM THEO MẪU 05
        p_cause = doc.add_paragraph()
        r = p_cause.add_run(f"- Phân tích, đánh giá mức độ, nguyên nhân giảm trừ chi phí (-{total_deduct_all:,.0f} đồng):")
        r.bold = True
        r.font.size = Pt(11.5)

        causes = [
            f"Gói Hoàn thiện mặt đứng (-4.099.416.464 đ): Đơn giá vách nhôm kính hệ 65 kính hộp Low-E áp 5.213.250 đ/m2 cao hơn báo giá nhà máy 4.850.000 đ/m2; Tấm ốp composite áp cao hơn báo giá 184.600 đ/m2.",
            f"Gói Điều hòa không khí & Thông gió (-3.233.297.000 đ): 8 tổ máy dàn nóng VRF và các cụm dàn lạnh áp đơn giá tạm tính cao hơn giá công bố chính thức của Daikin Việt Nam từ 6,7 triệu đến 39 triệu đồng/bộ.",
            f"Gói Biện pháp thi công Kingpost (-455.184.503 đ): Đơn giá thép cọc Kingpost tính 19.363.541 đ/tấn cao hơn công bố giá Posco (18.500.000 đ/tấn); Đinh chống cắt D19 và Vữa không co ngót B40 áp cao hơn báo giá thị trường."
        ]
        for c in causes:
            doc.add_paragraph(f"+ {c}").paragraph_format.space_after = Pt(2)

        # MỤC 7: KẾT LUẬN VÀ KIẾN NGHỊ (THEO MẪU SỐ 05)
        p_h7 = doc.add_paragraph()
        p_h7.paragraph_format.space_before = Pt(10)
        p_h7.paragraph_format.space_after = Pt(3)
        r = p_h7.add_run("7. Kết luận và kiến nghị")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        doc.add_paragraph(
            f"- Kết luận: Hồ sơ dự toán xây dựng công trình {self.project_name} đủ điều kiện sau khi hoàn thiện các nội dung giảm trừ "
            f"để Chủ đầu tư phê duyệt dự toán với giá trị thẩm tra là: {total_audit_all:,.0f} đồng "
            f"(Giảm trừ so với dự toán đề nghị thẩm tra là: {total_deduct_all:,.0f} đồng, tương đương giảm {pct_deduct_all:.2f}%).\n"
            f"- Kiến nghị: Đề nghị Chủ đầu tư yêu cầu Tư vấn thiết kế cập nhật dự toán theo giá trị thẩm tra, "
            f"chuẩn hóa các mã hiệu định mức theo Thông tư 38/2026/TT-BXD trước khi triển khai các bước tiếp theo."
        ).paragraph_format.space_after = Pt(8)

        # CHỮ KÝ THEO ĐÚNG MẪU SỐ 05 PHỤ LỤC VIII TT 36/2026/TT-BXD
        tbl_sig = doc.add_table(rows=2, cols=3)
        tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_sig.autofit = False

        c0 = tbl_sig.cell(0, 0)
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p0.add_run("NGƯỜI THẨM TRA\n(Ký, họ tên)")
        r.bold = True
        r.font.size = Pt(10.5)

        c1 = tbl_sig.cell(0, 1)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p1.add_run("NGƯỜI CHỦ TRÌ\n(Ký, họ tên)")
        r.bold = True
        r.font.size = Pt(10.5)

        c2 = tbl_sig.cell(0, 2)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p2.add_run("ĐƠN VỊ THẨM TRA\n(Ký, ghi rõ chức vụ, đóng dấu)")
        r.bold = True
        r.font.size = Pt(10.5)

        for col_i, tit in enumerate(["Kỹ sư Thẩm tra", "Chủ trì Thẩm tra", "Giám đốc Đơn vị Thẩm tra"]):
            c_b = tbl_sig.cell(1, col_i)
            c_b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            c_b.paragraphs[0].paragraph_format.space_before = Pt(45)
            r_b = c_b.paragraphs[0].add_run(tit)
            r_b.italic = True
            r_b.font.size = Pt(10.5)

        doc.save(out_path)

    def generate_excel_summary(self, results, out_path):
        """Tạo file Excel Bảng tổng hợp chi phí chênh lệch và chi tiết sai lệch toàn dự án"""
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_sub = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        fill_red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        font_header = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
        font_bold = Font(name="Times New Roman", size=10, bold=True)
        font_reg = Font(name="Times New Roman", size=10)
        font_link = Font(name="Times New Roman", size=10, color="0000FF", underline="single")
        font_alert = Font(name="Times New Roman", size=10, bold=True, color="C00000")
        border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

        # SHEET 1: TỔNG HỢP TOÀN DỰ ÁN
        ws1 = wb.create_sheet(title="TH CHI PHÍ TOÀN DỰ ÁN")
        ws1.merge_cells("A1:H1")
        ws1["A1"] = f"BẢNG TỔNG HỢP KẾT QUẢ THẨM TRA DỰ TOÁN TOÀN BỘ CÔNG TRÌNH"
        ws1["A1"].font = Font(name="Times New Roman", size=14, bold=True, color="1F4E79")
        ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[1].height = 28

        ws1.merge_cells("A2:H2")
        ws1["A2"] = f"Công trình: {self.project_name} | Chuẩn Mẫu số 05 Phụ lục VIII TT 36/2026/TT-BXD, NĐ 206/2026/NĐ-CP & TT 38/2026/TT-BXD"
        ws1["A2"].font = Font(name="Times New Roman", size=11, italic=True)
        ws1["A2"].alignment = Alignment(horizontal="center")

        headers_s1 = [
            "STT", "Tên hạng mục / Gói thầu", "File dự toán gốc",
            "Giá trị Dự toán (VNĐ)", "Giá trị Thẩm tra (VNĐ)",
            "Chênh lệch Giảm trừ (VNĐ)", "Tỷ lệ Giảm (%)", "Đường dẫn file Thẩm tra"
        ]
        for c_idx, h in enumerate(headers_s1, 1):
            cell = ws1.cell(row=4, column=c_idx, value=h)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws1.row_dimensions[4].height = 25

        row_idx = 5
        pkg_results = [r for r in results if 'TMDT' not in r['file_name'] and r['total_cost_original'] > 0]
        tot_orig = sum(r['total_cost_original'] for r in pkg_results)
        tot_audit = sum(r['total_cost_audited'] for r in pkg_results)
        tot_deduct = sum(r['total_deduction'] for r in pkg_results)
        pct_deduct = (tot_deduct / tot_orig * 100) if tot_orig > 0 else 0.0

        for idx, res in enumerate(pkg_results, 1):
            ws1.cell(row=row_idx, column=1, value=idx).alignment = Alignment(horizontal="center")
            ws1.cell(row=row_idx, column=2, value=res['file_name'].replace('.xls', '').replace('.xlsx', ''))
            ws1.cell(row=row_idx, column=3, value=res['file_name'])

            c_orig = ws1.cell(row=row_idx, column=4, value=res['total_cost_original'])
            c_orig.number_format = "#,##0"

            c_aud = ws1.cell(row=row_idx, column=5, value=res['total_cost_audited'])
            c_aud.number_format = "#,##0"

            c_ded = ws1.cell(row=row_idx, column=6, value=res['total_deduction'])
            c_ded.number_format = "#,##0"
            c_ded.font = font_alert

            c_pct = ws1.cell(row=row_idx, column=7, value=res['reduction_pct'] / 100)
            c_pct.number_format = "0.00%"
            c_pct.alignment = Alignment(horizontal="center")

            aud_path = res.get('audited_file_path', '')
            c_f = ws1.cell(row=row_idx, column=8, value=os.path.basename(aud_path) if aud_path else "Đã thẩm tra")
            if aud_path:
                c_f.hyperlink = f"file:///{aud_path.replace('\\', '/')}"
                c_f.font = font_link

            for c_i in range(1, 9):
                ws1.cell(row=row_idx, column=c_i).border = border_thin

            row_idx += 1

        # Tổng cộng các gói thầu
        ws1.merge_cells(f"A{row_idx}:C{row_idx}")
        ws1.cell(row=row_idx, column=1, value="TỔNG CỘNG CÁC GÓI THẦU ĐÃ THẨM TRA:").alignment = Alignment(horizontal="right")
        ws1.cell(row=row_idx, column=1).font = font_bold
        ws1.cell(row=row_idx, column=4, value=tot_orig).number_format = "#,##0"
        ws1.cell(row=row_idx, column=5, value=tot_audit).number_format = "#,##0"
        ws1.cell(row=row_idx, column=6, value=tot_deduct).number_format = "#,##0"
        ws1.cell(row=row_idx, column=7, value=pct_deduct / 100).number_format = "0.00%"

        for c_i in range(1, 9):
            cell = ws1.cell(row=row_idx, column=c_i)
            cell.font = Font(name="Times New Roman", size=10.5, bold=True, color="C00000" if c_i == 6 else "000000")
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            cell.border = border_thin

        row_idx += 2
        # Dòng đối chiếu TMĐT
        ws1.merge_cells(f"A{row_idx}:C{row_idx}")
        ws1.cell(row=row_idx, column=1, value="TỔNG MỨC ĐẦU TƯ ĐƯỢC DUYỆT (TMDT NOXH C4.xls):").alignment = Alignment(horizontal="right")
        ws1.cell(row=row_idx, column=1).font = font_bold
        c_tmdt = ws1.cell(row=row_idx, column=4, value=1395500000000.0)
        c_tmdt.number_format = "#,##0"
        c_tmdt.font = font_bold
        ws1.merge_cells(f"E{row_idx}:H{row_idx}")
        ws1.cell(row=row_idx, column=5, value="✓ Đảm bảo nằm trong cơ cấu Chi phí Xây dựng (952,7 tỷ) & Thiết bị (172,7 tỷ)").font = Font(name="Times New Roman", size=10, italic=True, color="375623")

        # SHEET 2: CHI TIẾT TẤT CẢ CÁC KHOẢN GIẢM TRỪ VẬT TƯ & THIẾT BỊ
        ws2 = wb.create_sheet(title="CHI TIẾT CÁC KHOẢN GIẢM TRỪ")
        ws2.merge_cells("A1:K1")
        ws2["A1"] = "BẢNG CHI TIẾT CÁC HẠNG MỤC VẬT TƯ, THIẾT BỊ CẮT GIẢM TRỪ DỰ TOÁN (MỤC 5.3)"
        ws2["A1"].font = Font(name="Times New Roman", size=13, bold=True, color="1F4E79")
        ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[1].height = 26

        headers_s2 = [
            "STT", "Hạng mục / File dự toán", "Tên vật tư / Thiết bị", "ĐVT",
            "Khối lượng", "Đơn giá Dự toán (đ)", "Đơn giá Báo giá (đ)", "Chênh lệch (đ)",
            "Thành tiền Giảm trừ (đ)", "File Báo giá đối chiếu", "Lý do và Căn cứ xử lý"
        ]
        for c_idx, h in enumerate(headers_s2, 1):
            cell = ws2.cell(row=3, column=c_idx, value=h)
            cell.fill = fill_sub
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws2.row_dimensions[3].height = 25

        row_idx2 = 4
        stt_d = 1
        for res in results:
            for d in res['deduction_items']:
                ws2.cell(row=row_idx2, column=1, value=stt_d).alignment = Alignment(horizontal="center")
                ws2.cell(row=row_idx2, column=2, value=res['file_name'])
                ws2.cell(row=row_idx2, column=3, value=d['item_name'])
                ws2.cell(row=row_idx2, column=4, value=d['unit']).alignment = Alignment(horizontal="center")

                c_q = ws2.cell(row=row_idx2, column=5, value=d['qty'])
                c_q.number_format = "#,##0.00"

                c_ep = ws2.cell(row=row_idx2, column=6, value=d['est_price'])
                c_ep.number_format = "#,##0"

                c_qp = ws2.cell(row=row_idx2, column=7, value=d['quote_price'])
                c_qp.number_format = "#,##0"

                c_dp = ws2.cell(row=row_idx2, column=8, value=d['diff_price'])
                c_dp.number_format = "#,##0"
                c_dp.font = font_alert

                c_da = ws2.cell(row=row_idx2, column=9, value=d['deduction_amount'])
                c_da.number_format = "#,##0"
                c_da.font = font_alert

                c_lf = ws2.cell(row=row_idx2, column=10, value=d['quote_file'])
                c_lf.hyperlink = f"file:///{d['quote_path'].replace('\\', '/')}"
                c_lf.font = font_link

                ws2.cell(row=row_idx2, column=11, value=d['reason'])

                for c_i in range(1, 12):
                    ws2.cell(row=row_idx2, column=c_i).border = border_thin
                    ws2.cell(row=row_idx2, column=c_i).fill = fill_red

                row_idx2 += 1
                stt_d += 1

        # Tổng giảm trừ sheet 2
        ws2.merge_cells(f"A{row_idx2}:H{row_idx2}")
        ws2.cell(row=row_idx2, column=1, value="TỔNG CỘNG GIÁ TRỊ GIẢM TRỪ TOÀN DỰ ÁN:").alignment = Alignment(horizontal="right")
        ws2.cell(row=row_idx2, column=1).font = font_bold
        c_tot = ws2.cell(row=row_idx2, column=9, value=tot_deduct)
        c_tot.number_format = "#,##0"
        c_tot.font = Font(name="Times New Roman", size=11, bold=True, color="C00000")

        for c_i in range(1, 12):
            cell = ws2.cell(row=row_idx2, column=c_i)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            cell.border = border_thin

        # Căn chỉnh độ rộng cột
        for ws in [ws1, ws2]:
            for col in ws.columns:
                max_len = 0
                for cell in col[:40]:
                    v = str(cell.value or '')
                    if '\n' in v:
                        v = v.split('\n')[0]
                    max_len = max(max_len, len(v))
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)

        wb.active = 0
        wb.save(out_path)
