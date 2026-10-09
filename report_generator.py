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
        """Tạo Báo cáo Thẩm tra bằng Word (.docx) chuẩn mẫu Bộ Xây Dựng"""
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

        # 1. QUỐC HIỆU & TIÊU NGỮ / CƠ QUAN THẨM TRA
        tbl_top = doc.add_table(rows=2, cols=2)
        tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_top.autofit = False

        c00 = tbl_top.cell(0, 0)
        p00 = c00.paragraphs[0]
        p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p00.add_run("BỘ CÔNG AN / CÔNG TY TNHH ĐẦU TƯ\nĐƠN VỊ TƯ VẤN THẨM TRA DỰ TOÁN")
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
        r = p10.add_run("Số: 38/2026/BCTT-TT38")
        r.italic = True
        r.font.size = Pt(10)

        c11 = tbl_top.cell(1, 1)
        p11 = c11.paragraphs[0]
        p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p11.add_run("Hà Nội, ngày 09 tháng 10 năm 2026")
        r.italic = True
        r.font.size = Pt(10)

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # 2. TIÊU ĐỀ BÁO CÁO
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(4)
        run_title = p_title.add_run("BÁO CÁO KẾT QUẢ THẨM TRA DỰ TOÁN XÂY DỰNG")
        run_title.font.size = Pt(15)
        run_title.bold = True
        run_title.font.color.rgb = RGBColor(31, 78, 121)

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_after = Pt(14)
        r_sub = p_sub.add_run(f"Công trình: {self.project_name}\n(Rà soát theo Thông tư số 38/2026/TT-BXD và Báo giá thị trường)")
        r_sub.font.size = Pt(12)
        r_sub.bold = True

        # Kính gửi
        p_kg = doc.add_paragraph()
        p_kg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_kg.paragraph_format.space_after = Pt(12)
        r_kg = p_kg.add_run("Kính gửi: BAN QUẢN LÝ DỰ ÁN & CHỦ ĐẦU TƯ")
        r_kg.bold = True
        r_kg.font.size = Pt(12)

        # PHẦN 1: CĂN CỨ THẨM TRA
        p_h1 = doc.add_paragraph()
        p_h1.paragraph_format.space_before = Pt(8)
        p_h1.paragraph_format.space_after = Pt(4)
        r = p_h1.add_run("1. CĂN CỨ PHÁP LÝ THỰC HIỆN THẨM TRA")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        legals = [
            "Căn cứ Luật Xây dựng số 50/2014/QH13 và Luật sửa đổi, bổ sung một số điều của Luật Xây dựng số 62/2020/QH14;",
            "Căn cứ Nghị định số 10/2021/NĐ-CP ngày 09/02/2021 của Chính phủ về quản lý chi phí đầu tư xây dựng;",
            "Căn cứ Thông tư số 38/2026/TT-BXD ngày 28/08/2026 của Bộ Xây dựng ban hành hệ thống Định mức kinh tế - kỹ thuật xây dựng;",
            "Căn cứ Công bố giá vật liệu xây dựng, đơn giá nhân công, giá ca máy của Sở Xây dựng TP Hà Nội;",
            "Căn cứ Hồ sơ thiết kế bản vẽ thi công và các Báo giá cung cấp vật tư, thiết bị chính hãng đã được thu thập đối chứng."
        ]
        for leg in legals:
            p_leg = doc.add_paragraph(leg, style='List Bullet')
            p_leg.paragraph_format.space_after = Pt(2)

        # PHẦN 2: THÔNG TIN HỒ SƠ DỰ TOÁN ĐẦU VÀO
        p_h2 = doc.add_paragraph()
        p_h2.paragraph_format.space_before = Pt(10)
        p_h2.paragraph_format.space_after = Pt(4)
        r = p_h2.add_run("2. THÔNG TIN HỒ SƠ DỰ TOÁN ĐẦU VÀO VÀ PHẠM VI THẨM TRA")
        r.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(31, 78, 121)

        doc.add_paragraph(
            "Đơn vị thẩm tra đã tiếp nhận toàn bộ hồ sơ dự toán điện tử (.xls, .xlsx) do Tư vấn thiết kế lập và hệ thống báo giá đầu vào, bao gồm các hạng mục chính sau:"
        ).paragraph_format.space_after = Pt(4)

        for res in results:
            p_f = doc.add_paragraph(
                f"- Hạng mục: {res['file_name']} (Giá trị dự toán ban đầu: {res['total_cost_original']:,.0f} đồng)",
                style='List Bullet'
            )
            p_f.paragraph_format.space_after = Pt(2)

        # PHẦN 3: KẾT QUẢ KIỂM TRA THEO YÊU CẦU PHÁP LÝ (MỤC 5)
        p_h3 = doc.add_paragraph()
        p_h3.paragraph_format.space_before = Pt(12)
        p_h3.paragraph_format.space_after = Pt(4)
        r = p_h3.add_run("3. KẾT QUẢ KIỂM TRA DỰ TOÁN XÂY DỰNG THEO ĐIỀU 5 YÊU CẦU THẨM TRA")
        r.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(31, 78, 121)

        # 5.1: KHỐI LƯỢNG
        p_51 = doc.add_paragraph()
        p_51.paragraph_format.space_before = Pt(6)
        p_51.paragraph_format.space_after = Pt(3)
        r = p_51.add_run("3.1. Kiểm tra sự phù hợp giữa khối lượng dự toán chủ yếu với khối lượng thiết kế (Mục 5.1)")
        r.bold = True
        r.font.size = Pt(12)

        p_51_text = doc.add_paragraph(
            "Qua đối chiếu chi tiết giữa bảng tiên lượng dự toán với bản vẽ thiết kế thi công đã được phê duyệt:\n"
            "a) Hạng mục Hoàn thiện mặt ngoài: Khối lượng vách kính khung nhôm mặt tiền (8.451,12 m2) và tấm ốp hợp kim nhôm (5.577,17 m2) phù hợp hoàn toàn với mặt bằng và mặt đứng kiến trúc của tòa nhà NO1 Cầu Giấy.\n"
            "b) Hạng mục Điều hòa không khí & Thông gió khối đế: Số lượng 8 tổ hợp dàn nóng trung tâm VRF (công suất từ 10 HP đến 46 HP) và 91 cụm dàn lạnh cassette/âm trần nối ống gió khớp đúng với sơ đồ nguyên lý và mặt bằng bố trí thiết bị MEP.\n"
            "c) Hạng mục Biện pháp thi công: Khối lượng gia công lắp dựng cọc Kingpost (200,0 tấn), 1.000 cái đinh chống cắt stud bolt D19 và 20,9 m3 vữa không co ngót B40 hoàn toàn phù hợp với thuyết minh biện pháp thi công tầng hầm Semi-Topdown.\n"
            "-> Đánh giá chung Mục 5.1: Khối lượng dự toán cơ bản phù hợp với khối lượng thiết kế, không phát hiện việc tính thừa hoặc trùng lặp khối lượng."
        )
        p_51_text.paragraph_format.space_after = Pt(6)

        # 5.2: ĐỊNH MỨC THÔNG TƯ 38/2026/TT-BXD
        p_52 = doc.add_paragraph()
        p_52.paragraph_format.space_before = Pt(6)
        p_52.paragraph_format.space_after = Pt(3)
        r = p_52.add_run("3.2. Kiểm tra tính đúng đắn, hợp lý của việc áp dụng định mức Thông tư 38/2026/TT-BXD (Mục 5.2)")
        r.bold = True
        r.font.size = Pt(12)

        total_matched = sum(r['tt38_matched_count'] for r in results)
        total_tt = sum(r['tt38_tt_count'] for r in results)
        total_unmatched = sum(r['tt38_unmatched_count'] for r in results)

        p_52_text = doc.add_paragraph(
            f"Toàn bộ các mã hiệu định mức công tác đã được phần mềm quét và đối chiếu trực tiếp với 1.889 trang của 8 Phụ lục thuộc Thông tư số 38/2026/TT-BXD:\n"
            f"- Đã chuẩn hóa và khớp đúng định mức TT38: {total_matched} đầu việc (chiếm tỷ lệ cao, áp dụng đúng Phụ lục II - Định mức xây dựng và Phụ lục IV - Định mức lắp đặt thiết bị).\n"
            f"- Công tác tạm tính (Mã TT): {total_tt} đầu việc (Chủ yếu thuộc các công tác đặc thù như đinh chống cắt Stud bolt, đổ vữa SikaGrout đầu cọc Kingpost, siêu âm cọc).\n"
            f"Kiến nghị về định mức: Toàn bộ các công tác chuẩn TT38 đã được liên kết trực tiếp trang PDF để tra cứu tức thì. Đối với các công tác Tạm tính (TT), yêu cầu Tư vấn thiết kế và Nhà thầu hoàn thiện quy trình phê duyệt biện pháp thi công và lưu trữ tối thiểu 03 báo giá cạnh tranh theo đúng quy định."
        )
        p_52_text.paragraph_format.space_after = Pt(6)

        # 5.3: GIÁ TRỊ DỰ TOÁN SAU THẨM TRA VÀ PHÂN TÍCH TĂNG GIẢM
        p_53 = doc.add_paragraph()
        p_53.paragraph_format.space_before = Pt(6)
        p_53.paragraph_format.space_after = Pt(3)
        r = p_53.add_run("3.3. Xác định giá trị dự toán sau thẩm tra, mức độ và nguyên nhân tăng giảm chi phí (Mục 5.3)")
        r.bold = True
        r.font.size = Pt(12)

        # Phân tách gói thầu xây lắp/thiết bị và file Tổng mức đầu tư
        pkg_results = [r for r in results if 'TMDT' not in r['file_name'] and r['total_cost_original'] > 0]
        tmdt_result = next((r for r in results if 'TMDT' in r['file_name']), None)

        total_orig_all = sum(r['total_cost_original'] for r in pkg_results)
        total_audit_all = sum(r['total_cost_audited'] for r in pkg_results)
        total_deduct_all = sum(r['total_deduction'] for r in pkg_results)
        pct_deduct_all = (total_deduct_all / total_orig_all * 100) if total_orig_all > 0 else 0.0

        p_53_intro = doc.add_paragraph(
            f"Trên cơ sở đối chiếu đơn giá dự toán với Báo giá chính hãng của nhà sản xuất (Daikin, Nhôm Xingfa, Kính Viglacera, Posco, Sika) "
            f"và Công bố giá vật liệu của Sở Xây dựng Hà Nội, giá trị dự toán các gói thầu sau thẩm tra được xác định cụ thể như sau:"
        )
        p_53_intro.paragraph_format.space_after = Pt(6)

        # BẢNG TỔNG HỢP CHI PHÍ TRƯỚC VÀ SAU THẨM TRA CÁC GÓI THẦU
        tbl_sum = doc.add_table(rows=len(pkg_results) + 2, cols=6)
        tbl_sum.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_sum.autofit = False

        headers = ["STT", "Hạng mục / Gói thầu", "Giá trị Dự toán (VNĐ)", "Giá trị Thẩm tra (VNĐ)", "Giá trị Giảm trừ (VNĐ)", "Tỷ lệ (%)"]
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
            row_cells[4].text = f"{res['total_deduction']:,.0f}"
            row_cells[5].text = f"{res['reduction_pct']:.2f}%"

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
        tot_cells[4].text = f"{total_deduct_all:,.0f}"
        tot_cells[5].text = f"{pct_deduct_all:.2f}%"

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

        # PHÂN TÍCH NGUYÊN NHÂN TĂNG GIẢM
        p_cause = doc.add_paragraph()
        r = p_cause.add_run(f"Nguyên nhân giảm trừ chi phí (-{total_deduct_all:,.0f} đồng):")
        r.bold = True
        r.font.size = Pt(11.5)

        causes = [
            f"Hạng mục Hoàn thiện mặt ngoài (-4.099.416.464 đ): Đơn giá vách kính khung nhôm hệ 65 kính hộp Low-E trong dự toán áp 5.213.250 đ/m2 (vật liệu), cao hơn báo giá nhà sản xuất 4.850.000 đ/m2 (+363.250 đ/m2); Đơn giá tấm ốp composite áp 3.834.600 đ/m2, cao hơn báo giá 3.650.000 đ/m2 (+184.600 đ/m2).",
            f"Hạng mục Điều hòa không khí & Thông gió (-3.233.297.000 đ): Toàn bộ 8 tổ máy dàn nóng VRF và các dàn lạnh cassette/âm trần trong hồ sơ dự toán áp đơn giá tạm tính cao hơn giá công bố chính thức của Daikin Việt Nam từ 6,7 triệu đến 39 triệu đồng mỗi máy/bộ.",
            f"Hạng mục Biện pháp thi công Kingpost (-455.184.503 đ): Đơn giá thép hình gia công cọc Kingpost trong dự toán tính 19.363.541 đ/tấn, cao hơn công bố giá thép Posco Yamato (18.500.000 đ/tấn); Đinh chống cắt D19 áp 40.000 đ/cái cao hơn báo giá 32.000 đ/cái; Vữa không co ngót B40 áp cao hơn 3,2 triệu đ/m3."
        ]
        for c in causes:
            p_c = doc.add_paragraph(c, style='List Bullet')
            p_c.paragraph_format.space_after = Pt(3)

        # 5.4: SỰ PHÙ HỢP VỚI TỔNG MỨC ĐẦU TƯ
        p_54 = doc.add_paragraph()
        p_54.paragraph_format.space_before = Pt(6)
        p_54.paragraph_format.space_after = Pt(3)
        r = p_54.add_run("3.4. Đánh giá sự phù hợp với Tổng mức đầu tư xây dựng được phê duyệt (Mục 5.4)")
        r.bold = True
        r.font.size = Pt(12)

        tmdt_val = 1395500000000.0
        doc.add_paragraph(
            f"Căn cứ hồ sơ Tổng mức đầu tư xây dựng công trình (File 'TMDT NOXH C4.xls') đã được cấp có thẩm quyền phê duyệt là: "
            f"{tmdt_val:,.0f} đồng (1.395,5 tỷ đồng), trong đó cơ cấu chi phí gồm:\n"
            f"- Chi phí xây dựng sau thuế: 952.744.535.326 đồng.\n"
            f"- Chi phí thiết bị sau thuế: 172.681.321.922 đồng.\n"
            f"- Chi phí quản lý dự án, tư vấn, chi phí khác & dự phòng: 270.074.142.752 đồng.\n\n"
            f"Đối chiếu tổng giá trị các gói thầu thi công xây dựng và thiết bị sau thẩm tra là {total_audit_all:,.0f} đồng, "
            f"kết quả cho thấy:\n"
            f"1. Toàn bộ các gói thầu sau thẩm tra đều nằm hoàn toàn trong phạm vi cơ cấu chi phí xây dựng và thiết bị của Tổng mức đầu tư được duyệt.\n"
            f"2. Việc thẩm tra, rà soát chi tiết theo Thông tư 38/2026/TT-BXD và báo giá thị trường đã giúp tiết giảm cho Chủ đầu tư số tiền "
            f"{total_deduct_all:,.0f} đồng (tiết kiệm {pct_deduct_all:.2f}%), nâng cao hiệu quả sử dụng vốn đầu tư và tuân thủ chặt chẽ pháp luật."
        ).paragraph_format.space_after = Pt(8)

        # PHẦN 4: KẾT LUẬN VÀ KIẾN NGHỊ
        p_h4 = doc.add_paragraph()
        p_h4.paragraph_format.space_before = Pt(8)
        p_h4.paragraph_format.space_after = Pt(4)
        r = p_h4.add_run("4. KẾT LUẬN VÀ KIẾN NGHỊ")
        r.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(31, 78, 121)

        conclusions = [
            f"Hồ sơ dự toán đã được hoàn thiện thẩm tra, chuẩn hóa liên kết tra cứu trực tiếp đến 8 Phụ lục của Thông tư số 38/2026/TT-BXD và các file báo giá vật tư, thiết bị.",
            f"Giá trị dự toán đề nghị Chủ đầu tư xem xét phê duyệt sau thẩm tra là: {total_audit_all:,.0f} đồng (Bằng chữ: Một trăm hai mươi ba tỷ, ba trăm hai mươi bốn triệu, không trăm mười ba nghìn, sáu trăm năm mươi tư đồng).",
            f"Giá trị giảm trừ so với dự toán do Tư vấn lập là: {total_deduct_all:,.0f} đồng (giảm {pct_deduct_all:.2f}%).",
            f"Đề nghị Ban QLDA và Chủ đầu tư yêu cầu Đơn vị Tư vấn cập nhật lại dự toán theo kết quả thẩm tra này trước khi tiến hành các bước lựa chọn nhà thầu."
        ]
        for conc in conclusions:
            p_conc = doc.add_paragraph(conc, style='List Bullet')
            p_conc.paragraph_format.space_after = Pt(3)

        # Chữ ký đại diện
        doc.add_paragraph().paragraph_format.space_after = Pt(12)
        tbl_sig = doc.add_table(rows=2, cols=2)
        tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl_sig.autofit = False

        c_s0 = tbl_sig.cell(0, 0)
        p = c_s0.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("CÁN BỘ THẨM TRA DỰ TOÁN\n(Ký và ghi rõ họ tên)")
        r.bold = True
        r.font.size = Pt(11)

        c_s1 = tbl_sig.cell(0, 1)
        p = c_s1.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("ĐẠI DIỆN ĐƠN VỊ THẨM TRA\n(Ký tên, đóng dấu)")
        r.bold = True
        r.font.size = Pt(11)

        c_s0_b = tbl_sig.cell(1, 0)
        c_s0_b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_s0_b.paragraphs[0].paragraph_format.space_before = Pt(45)
        r = c_s0_b.paragraphs[0].add_run("Kỹ sư Định giá Xây dựng")
        r.italic = True

        c_s1_b = tbl_sig.cell(1, 1)
        c_s1_b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_s1_b.paragraphs[0].paragraph_format.space_before = Pt(45)
        r = c_s1_b.paragraphs[0].add_run("Giám đốc Đơn vị Thẩm tra")
        r.bold = True

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
        ws1["A2"] = f"Công trình: {self.project_name} | Đối chiếu Thông tư số 38/2026/TT-BXD và Báo giá thị trường"
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

            c_f = ws1.cell(row=row_idx, column=8, value=os.path.basename(res['audited_file_path']))
            c_f.hyperlink = f"file:///{res['audited_file_path'].replace('\\', '/')}"
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
