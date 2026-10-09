# -*- coding: utf-8 -*-
"""
QUOTATION INDEXER - Module quản lý và tra cứu Báo giá vật liệu & Thiết bị
"""

import os
import re
import unicodedata
from python_calamine import CalamineWorkbook

def remove_accents(text):
    """Bỏ dấu tiếng Việt để so sánh tìm kiếm linh hoạt"""
    if not text:
        return ""
    text = str(text)
    # Chuẩn hóa unicode NFD
    nfkd_form = unicodedata.normalize('NFKD', text)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower()

def normalize_text(text):
    """Chuẩn hóa chuỗi text, loại bỏ ký tự đặc biệt, chuẩn hóa khoảng trắng"""
    if not text:
        return ""
    # Chuyển về chữ thường
    text = str(text).strip().lower()
    # Thay thế các ký tự phân cách thành dấu cách
    text = re.sub(r'[\(\)\[\],;:\-_/\+]', ' ', text)
    # Loại bỏ khoảng trắng thừa
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def strip_action_prefixes(text):
    """Loại bỏ các động từ hành động đầu câu trong dự toán xây dựng để lấy tên vật tư gốc"""
    if not text:
        return ""
    norm = normalize_text(text)
    # Các cụm tiền tố thi công thường gặp trong dự toán
    patterns = [
        r'^(cung cap va lap dat|cung cap va lap dung|cung cap|lap dat|lap dung|thi cong|gia cong va lap dung|gia cong|san xuat va lap dung|san xuat|do vua|do be tong|bom be tong|xay|trat|lat|op|son|pha do|dao|dap)\s+',
        r'^(tram vi tri|dau cot|cho vi tri|tai vi tri)\s+'
    ]
    cleaned = norm
    for p in patterns:
        cleaned = re.sub(p, '', cleaned, flags=re.IGNORECASE)
    return cleaned.strip()

def tokenize(text):
    """Tách chuỗi thành tập hợp các từ khóa đã bỏ dấu"""
    norm = strip_action_prefixes(text)
    no_acc = remove_accents(norm)
    tokens = set(no_acc.split())
    # Loại bỏ các từ dừng thông dụng
    stop_words = {'va', 'cua', 'theo', 'cac', 'cho', 'he', 'bo', 'tai', 'loai', 'day', 'mm', 'm2', 'm3', 'kg', 'tan', 'cai'}
    return tokens - stop_words

class QuotationIndexer:
    def __init__(self, quotation_dir):
        self.quotation_dir = os.path.abspath(quotation_dir)
        self.items = []
        self._build_index()

    def _build_index(self):
        """Quét và lập chỉ mục toàn bộ file Excel trong thư mục báo giá"""
        if not os.path.exists(self.quotation_dir):
            return

        for fname in os.listdir(self.quotation_dir):
            if fname.startswith('~$') or not (fname.lower().endswith('.xlsx') or fname.lower().endswith('.xls')):
                continue

            fpath = os.path.join(self.quotation_dir, fname)
            try:
                wb = CalamineWorkbook.from_path(fpath)
                for sname in wb.sheet_names:
                    rows = wb.get_sheet_by_name(sname).to_python()
                    self._parse_sheet(rows, fname, fpath, sname)
            except Exception as e:
                print(f"Lỗi khi đọc file báo giá {fname}: {e}")

    def _parse_sheet(self, rows, fname, fpath, sname):
        """Phát hiện bảng báo giá và trích xuất dữ liệu"""
        if not rows or len(rows) < 2:
            return

        # Tìm dòng tiêu đề
        header_row_idx = -1
        col_name = -1
        col_unit = -1
        col_price = -1
        col_supplier = -1

        for r_idx, row in enumerate(rows[:15]):
            row_str = [str(c or '').lower() for c in row]
            for c_idx, cell_val in enumerate(row_str):
                if any(k in cell_val for k in ['tên vật tư', 'tên thiết bị', 'tên hàng', 'danh mục', 'quy cách', 'chủng loại']):
                    col_name = c_idx
                elif any(k in cell_val for k in ['đơn vị', 'đvt', 'đơn vị tính']):
                    col_unit = c_idx
                elif any(k in cell_val for k in ['đơn giá', 'giá trước thuế', 'giá chưa vat', 'đơn giá (đ)']):
                    col_price = c_idx
                elif any(k in cell_val for k in ['nhà sản xuất', 'hãng sản xuất', 'tiêu chuẩn', 'xuất xứ', 'nhà cung cấp']):
                    col_supplier = c_idx

            if col_name != -1 and col_price != -1:
                header_row_idx = r_idx
                break

        if header_row_idx == -1:
            return

        # Đọc các dòng dữ liệu bên dưới tiêu đề
        for r_idx in range(header_row_idx + 1, len(rows)):
            row = rows[r_idx]
            if len(row) <= max(col_name, col_price):
                continue

            raw_name = row[col_name]
            if not raw_name or str(raw_name).strip() in ['', 'None', 'TỔNG CỘNG', 'Tổng cộng']:
                continue

            raw_name_str = str(raw_name).strip()
            # Bỏ qua dòng tiêu đề phụ (không có giá)
            raw_price = row[col_price]
            if raw_price is None or raw_price == '':
                continue

            try:
                price_val = float(str(raw_price).replace(',', '').replace(' ', '').replace('đ', ''))
                if price_val <= 0:
                    continue
            except (ValueError, TypeError):
                continue

            unit_str = str(row[col_unit]).strip() if col_unit != -1 and col_unit < len(row) and row[col_unit] else ''
            supplier_str = str(row[col_supplier]).strip() if col_supplier != -1 and col_supplier < len(row) and row[col_supplier] else ''

            tokens = tokenize(raw_name_str)
            no_acc_name = remove_accents(normalize_text(raw_name_str))

            # Trích xuất các số công suất đặc trưng (ví dụ 10hp, 14hp, 7.1kw, 9.0kw, v.v.)
            specs = set(re.findall(r'\d+(?:\.\d+)?\s*(?:hp|kw|mm|m|k|w|b\d+)', no_acc_name))

            self.items.append({
                'name': raw_name_str,
                'unit': unit_str,
                'price': price_val,
                'supplier': supplier_str,
                'file_name': fname,
                'file_path': fpath,
                'sheet_name': sname,
                'row_idx': r_idx + 1,
                'tokens': tokens,
                'no_acc_name': no_acc_name,
                'specs': specs
            })

    def match(self, query_name, query_unit=None):
        """
        Tìm kiếm báo giá phù hợp nhất cho tên vật tư/thiết bị trong dự toán.
        Trả về (matched_item, score) hoặc (None, 0.0)
        """
        if not query_name:
            return None, 0.0

        q_norm = normalize_text(query_name)
        q_no_acc = remove_accents(q_norm)
        q_tokens = tokenize(query_name)
        q_specs = set(re.findall(r'\d+(?:\.\d+)?\s*(?:hp|kw|mm|m|k|w|b\d+)', q_no_acc))

        best_item = None
        best_score = 0.0

        for item in self.items:
            # 1. Trùng khớp thông số kỹ thuật (HP, kW, mác vữa...) - Tiêu chí tiên quyết nếu có
            if q_specs and item['specs']:
                # Nếu cả hai bên đều có thông số nhưng không trùng nhau -> loại trừ
                if not (q_specs & item['specs']):
                    continue

            # 2. Tính độ tương đồng từ vựng (Jaccard similarity)
            inter = len(q_tokens & item['tokens'])
            union = len(q_tokens | item['tokens'])
            jaccard = inter / union if union > 0 else 0.0

            # 3. Thưởng điểm cho chuỗi con hoặc cụm từ khóa chính xác
            phrase_bonus = 0.0
            q_clean = remove_accents(strip_action_prefixes(query_name))
            if q_clean and (q_clean in item['no_acc_name'] or item['no_acc_name'] in q_clean):
                phrase_bonus = 0.5
            elif q_no_acc in item['no_acc_name'] or item['no_acc_name'] in q_no_acc:
                phrase_bonus = 0.4
            else:
                # Kiểm tra các cụm 2 từ liên tiếp
                q_words = q_clean.split() if q_clean else q_no_acc.split()
                for i in range(len(q_words) - 1):
                    pair = f"{q_words[i]} {q_words[i+1]}"
                    if pair in item['no_acc_name']:
                        phrase_bonus += 0.2
                        break

            # Thưởng điểm trùng thông số
            spec_bonus = 0.3 if (q_specs and item['specs'] and (q_specs & item['specs'])) else 0.0

            score = jaccard * 0.5 + phrase_bonus + spec_bonus

            # Kiểm tra đơn vị tính (nếu có)
            if query_unit and item['unit']:
                u1 = remove_accents(query_unit.strip())
                u2 = remove_accents(item['unit'].strip())
                if u1 == u2:
                    score += 0.1

            if score > best_score:
                best_score = score
                best_item = item

        # Ngưỡng chấp nhận khớp báo giá
        if best_score >= 0.35:
            return best_item, best_score
        return None, best_score

if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    bg_dir = r"d:\Cong ty\DuToanChuyenNghiep\Bao_Gia_Dau_Vao"
    indexer = QuotationIndexer(bg_dir)
    print(f"Tổng số mục báo giá đã nạp: {len(indexer.items)}")
    
    test_queries = [
        ("Vách kính khung nhôm mặt tiền", "m2"),
        ("Tấm ốp mặt tiền", "m2"),
        ("Dàn nóng CSL: 10 HP", "Cái"),
        ("Dàn nóng CSL: 14 HP", "Cái"),
        ("Dàn nóng CSL: 46 HP", "Cái"),
        ("Cung cấp và lắp đặt đinh chống cắt", "cái"),
        ("Vữa không co ngót B40 trám vị trí cọc", "m3"),
        ("Gia công Kingpost bằng thép hình", "tấn"),
        ("Sơn tường ngoài nhà", "m2")
    ]
    
    print("\n--- KIỂM TRA ĐỐI CHIẾU THỬ NGHIỆM ---")
    for q, u in test_queries:
        m, s = indexer.match(q, u)
        if m:
            print(f"✓ '{q}' ({u}) -> MATCH: '{m['name']}' | Giá: {m['price']:,.0f} đ | Điểm: {s:.2f} | File: {m['file_name']}")
        else:
            print(f"✗ '{q}' ({u}) -> KHÔNG TÌM THẤY BÁO GIÁ PHÙ HỢP (Điểm cao nhất: {s:.2f})")
