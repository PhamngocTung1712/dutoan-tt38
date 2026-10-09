# -*- coding: utf-8 -*-
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from quotation_indexer import QuotationIndexer
from audit_engine import AuditEngine

bg_dir = r"d:\Cong ty\DuToanChuyenNghiep\Bao_Gia_Dau_Vao"
tt38_index_path = r"d:\Cong ty\DuToanChuyenNghiep\tt38_index.json"
output_dir = r"d:\Cong ty\DuToanChuyenNghiep\Output_Tham_Tra_Du_Toan\01_DuToan_DaThamTra"

q_idx = QuotationIndexer(bg_dir)
engine = AuditEngine(tt38_index_path, q_idx)

files = [
    r"d:\Cong ty\DuToanChuyenNghiep\05.4.2_L1\HOAN THIEN MAT DUNG.xls",
    r"d:\Cong ty\DuToanChuyenNghiep\05.4.2_L1\DHKK KHOI DE.xls",
    r"d:\Cong ty\DuToanChuyenNghiep\05.4.2_L1\BPTC.xls"
]

for f in files:
    print(f"\n=======================================================")
    print(f"=== ĐANG THẨM TRA FILE: {os.path.basename(f)} ===")
    print(f"=======================================================")
    res = engine.audit_file(f, output_dir)
    print(f"✓ Đã lưu file kết quả: {res['audited_file_path']}")
    print(f"✓ Tổng dự toán ban đầu:  {res['total_cost_original']:>18,.0f} đ")
    print(f"✓ Dự toán sau thẩm tra:  {res['total_cost_audited']:>18,.0f} đ")
    print(f"✓ Giá trị giảm trừ:      {res['total_deduction']:>18,.0f} đ ({res['reduction_pct']:.2f}%)")
    print(f"✓ Thống kê định mức TT38: {res['tt38_matched_count']} chuẩn, {res['tt38_tt_count']} tạm tính, {res['tt38_unmatched_count']} chưa khớp")
    print(f"✓ Số hạng mục kiến nghị giảm trừ: {len(res['deduction_items'])}")
    for d in res['deduction_items']:
        print(f"   [-] {d['item_name']} ({d['unit']}): Giảm {d['deduction_amount']:>14,.0f} đ | {d['reason']}")
