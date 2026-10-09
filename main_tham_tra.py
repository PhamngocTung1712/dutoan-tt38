# -*- coding: utf-8 -*-
"""
CHƯƠNG TRÌNH THẨM TRA DỰ TOÁN TỰ ĐỘNG - ĐỐI CHIẾU THÔNG TƯ 38/2026/TT-BXD & BÁO GIÁ
Thực hiện thẩm tra toàn diện theo 4 tiêu chí pháp lý (Mục 5.1, 5.2, 5.3, 5.4)
"""

import os
import sys
import argparse
import time
from quotation_indexer import QuotationIndexer
from audit_engine import AuditEngine
from report_generator import ReportGenerator
import dang_ky_protocol

def main():
    sys.stdout.reconfigure(encoding='utf-8')

    default_base = os.path.dirname(os.path.abspath(__file__))
    default_dutoan = os.path.join(default_base, "05.4.2_L1")
    default_baogia = os.path.join(default_base, "Bao_Gia_Dau_Vao")
    default_output = os.path.join(default_base, "Output_Tham_Tra_Du_Toan")
    default_index = os.path.join(default_base, "tt38_index.json")

    parser = argparse.ArgumentParser(description="Hệ thống Thẩm tra Dự toán Xây dựng & Tra cứu TT38/2026/TT-BXD")
    parser.add_argument("--dutoan_dir", default=default_dutoan, help="Thư mục chứa các file Excel dự toán đầu vào")
    parser.add_argument("--baogia_dir", default=default_baogia, help="Thư mục chứa các file Báo giá vật tư/thiết bị")
    parser.add_argument("--output_dir", default=default_output, help="Thư mục xuất kết quả thẩm tra")
    parser.add_argument("--tt38_index", default=default_index, help="File chỉ mục định mức TT38 (json)")
    parser.add_argument("--file", default=None, help="Chỉ định thẩm tra 1 file dự toán cụ thể")

    args = parser.parse_args()

    print("=" * 75)
    print("   HỆ THỐNG THẨM TRA DỰ TOÁN XÂY DỰNG & ĐỐI CHIẾU ĐỊNH MỨC TT 38/2026/TT-BXD")
    print("   Bộ Xây Dựng - Quản lý chi phí & Kiểm soát giá thị trường")
    print("=" * 75)

    # 1. Đăng ký protocol tt38:// nếu chưa đăng ký
    try:
        dang_ky_protocol.register_tt38_protocol()
    except Exception as e:
        print(f"[!] Chú ý khi đăng ký protocol tt38://: {e}")

    # 2. Kiểm tra chỉ mục TT38
    if not os.path.exists(args.tt38_index):
        print(f"[ERROR] Không tìm thấy file chỉ mục TT38 tại: {args.tt38_index}")
        sys.exit(1)

    # 3. Lập chỉ mục Báo giá
    print(f"\n[1/4] Đang quét và nạp dữ liệu Báo giá từ: {args.baogia_dir}")
    t0 = time.time()
    quotation_idx = QuotationIndexer(args.baogia_dir)
    print(f"      ✓ Đã nạp thành công {len(quotation_idx.items)} mục báo giá vật tư, thiết bị ({time.time()-t0:.2f}s)")

    # 4. Khởi tạo Audit Engine
    print(f"\n[2/4] Khởi tạo công cụ Thẩm tra đối chiếu TT 38 & Báo giá...")
    engine = AuditEngine(args.tt38_index, quotation_idx)
    print(f"      ✓ Đã nạp {len(engine.tt38_index)} mã hiệu định mức từ 8 Phụ lục TT 38/2026/TT-BXD")

    # 5. Xác định danh sách file dự toán cần thẩm tra
    dir_audited_files = os.path.join(args.output_dir, "01_DuToan_DaThamTra")
    dir_reports = os.path.join(args.output_dir, "02_BaoCao_TongHop")
    os.makedirs(dir_audited_files, exist_ok=True)
    os.makedirs(dir_reports, exist_ok=True)

    input_files = []
    if args.file:
        if os.path.exists(args.file):
            input_files.append(os.path.abspath(args.file))
        else:
            print(f"[ERROR] File chỉ định không tồn tại: {args.file}")
            sys.exit(1)
    else:
        if not os.path.exists(args.dutoan_dir):
            print(f"[ERROR] Thư mục dự toán không tồn tại: {args.dutoan_dir}")
            sys.exit(1)
        for fname in os.listdir(args.dutoan_dir):
            # Bỏ qua các file tạm hoặc file đã được link trước đó
            if fname.startswith('~$') or '_Linked' in fname or '_Audited' in fname:
                continue
            if fname.lower().endswith('.xls') or fname.lower().endswith('.xlsx'):
                input_files.append(os.path.join(args.dutoan_dir, fname))

    if not input_files:
        print("[!] Không tìm thấy file dự toán nào hợp lệ để thẩm tra.")
        sys.exit(0)

    print(f"\n[3/4] Bắt đầu thẩm tra chi tiết {len(input_files)} file dự toán...")
    results = []

    for idx, fpath in enumerate(input_files, 1):
        fname = os.path.basename(fpath)
        print(f"\n   ------------------------------------------------------------")
        print(f"   [{idx}/{len(input_files)}] Đang thẩm tra: {fname}")
        try:
            res = engine.audit_file(fpath, dir_audited_files)
            results.append(res)
            print(f"      ✓ Xuất file thẩm tra: {os.path.basename(res['audited_file_path'])}")
            print(f"      - Giá trị ban đầu:    {res['total_cost_original']:>18,.0f} đ")
            print(f"      - Giá trị thẩm tra:   {res['total_cost_audited']:>18,.0f} đ")
            print(f"      - Giá trị giảm trừ:   {res['total_deduction']:>18,.0f} đ (-{res['reduction_pct']:.2f}%)")
            print(f"      - Số khoản giảm trừ:  {len(res['deduction_items'])} mục vật tư/thiết bị")
            print(f"      - Định mức TT38:      {res['tt38_matched_count']} chuẩn | {res['tt38_tt_count']} tạm tính")
        except Exception as e:
            print(f"      [X] Lỗi khi thẩm tra {fname}: {e}")
            import traceback
            traceback.print_exc()

    # 6. Xuất Báo cáo thẩm tra Word & Excel Tổng hợp
    print(f"\n[4/4] Đang lập Báo cáo Thẩm tra Word và Bảng Tổng hợp chi phí...")
    rep_gen = ReportGenerator()
    docx_path, xlsx_path = rep_gen.generate_all(results, dir_reports)

    # 7. Hiển thị bảng tổng kết trên Console
    tot_orig_all = sum(r['total_cost_original'] for r in results)
    tot_audit_all = sum(r['total_cost_audited'] for r in results)
    tot_deduct_all = sum(r['total_deduction'] for r in results)
    pct_all = (tot_deduct_all / tot_orig_all * 100) if tot_orig_all > 0 else 0.0

    print("\n" + "=" * 75)
    print("                    KẾT QUẢ THẨM TRA TOÀN DỰ ÁN")
    print("=" * 75)
    print(f"  • Tổng giá trị dự toán trước thẩm tra:  {tot_orig_all:>18,.0f} VNĐ")
    print(f"  • Giá trị dự toán kiến nghị phê duyệt:  {tot_audit_all:>18,.0f} VNĐ")
    print(f"  • TỔNG GIÁ TRỊ CẮT GIẢM TRỪ THẨM TRA:   {tot_deduct_all:>18,.0f} VNĐ")
    print(f"  • Tỷ lệ tiết giảm chi phí:                                {pct_all:>6.2f} %")
    print("-" * 75)
    print("  HỒ SƠ ĐÃ XUẤT TRONG OUTPUT FOLDER:")
    print(f"  1. File Báo cáo Thẩm tra Word (Theo Mục 5.1-5.4):")
    print(f"     -> {docx_path}")
    print(f"  2. Bảng Tổng hợp Chênh lệch & Chi tiết Giảm trừ (Excel):")
    print(f"     -> {xlsx_path}")
    print(f"  3. Thư mục File dự toán đã thẩm tra (Gắn link TT38 + Báo giá + Tô màu):")
    print(f"     -> {dir_audited_files}")
    print("=" * 75)
    print("   ✓ HOÀN TẤT THẨM TRA 100% THEO YÊU CẦU!")
    print("=" * 75)

if __name__ == "__main__":
    main()
