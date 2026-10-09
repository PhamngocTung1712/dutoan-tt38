# -*- coding: utf-8 -*-
"""
HỆ THỐNG THẨM TRA DỰ TOÁN ĐA DỰ ÁN (MULTI-PROJECT AUDIT PLATFORM)
Hỗ trợ quản lý nhiều dự án, cấu hình đường dẫn thư mục đầu vào riêng biệt cho từng dự án,
tải file trực tiếp, tra cứu định mức Thông tư 38/2026/TT-BXD và xuất báo cáo tự động.
"""

import os
import sys
import json
import uuid
import zipfile
import io
import shutil
import webbrowser
from datetime import datetime
from threading import Timer

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from flask import Flask, request, jsonify, render_template_string, send_file, send_from_directory
from werkzeug.utils import secure_filename

from quotation_indexer import QuotationIndexer
from audit_engine import AuditEngine
from report_generator import ReportGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB max

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECTS_DIR = os.path.join(BASE_DIR, "Projects_Data")
TT38_DIR = os.path.join(BASE_DIR, "TT 38", "TT 38")
INDEX_PATH = os.path.join(BASE_DIR, "tt38_index.json")
DB_FILE = os.path.join(BASE_DIR, "projects_db.json")

os.makedirs(PROJECTS_DIR, exist_ok=True)

# Bản đồ PDF 8 Phụ lục TT38
PDF_MAP = {
    1: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl1dmdtkhaosat_signed1.pdf"),
    2: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl2dmdtxaydung_signed1.pdf"),
    3: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl3dmdtlapdathtkt_signed1.pdf"),
    4: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl4dmdtldthietbi_signed1.pdf"),
    5: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl5dmdtthinghiem_signed1.pdf"),
    6: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl6dmdtsuachua_signed1.pdf"),
    7: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl7dmdtvatlieu_signed1.pdf"),
    8: os.path.join(TT38_DIR, "38.2026.tt-bxd-pl8dmqldavatvxd_signed1.pdf")
}

def resolve_dir(path):
    if not path:
        return ""
    if os.path.isabs(path) and os.path.exists(path):
        return path
    clean = path.replace('\\', '/').rstrip('/')
    bname = os.path.basename(clean)
    candidate = os.path.join(BASE_DIR, bname)
    if os.path.exists(candidate):
        return candidate
    candidate2 = os.path.join(BASE_DIR, clean)
    if os.path.exists(candidate2):
        return candidate2
    return os.path.abspath(os.path.join(BASE_DIR, path))

# Quản lý Database Dự án
def load_db():
    data = None
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            pass

    if not data or not data.get("projects"):
        data = {
            "active_id": "no1_cau_giay",
            "projects": {
                "no1_cau_giay": {
                    "id": "no1_cau_giay",
                    "name": "DỰ ÁN NO1 CẦU GIẤY (NHÀ Ở CBCS CAND)",
                    "dutoan_dir": "05.4.2_L1",
                    "baogia_dir": "Bao_Gia_Dau_Vao",
                    "output_dir": "Output_Tham_Tra_Du_Toan",
                    "tmdt": 1395500000000.0,
                    "created_at": "2026-10-09 14:00",
                    "status": "Đã thẩm tra"
                }
            }
        }
        save_db(data)

    for pid, p in data.get("projects", {}).items():
        for k in ["dutoan_dir", "baogia_dir", "output_dir"]:
            if k in p:
                p[k] = resolve_dir(p[k])

    return data

def save_db(data):
    with open(DB_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

HTML_PAGE = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ Thống Thẩm Tra Dự Toán Đa Dự Án - TT 38/2026/TT-BXD</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --primary: #1F4E79;
            --primary-light: #2F5597;
            --accent: #E26B00;
            --success: #28a745;
            --warning: #ffc107;
            --danger: #dc3545;
        }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; }
        .navbar { background-color: var(--primary); box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .card { border-radius: 10px; border: none; box-shadow: 0 2px 12px rgba(0,0,0,0.06); margin-bottom: 20px; }
        .card-header { background-color: #fff; border-bottom: 1px solid #edf2f7; font-weight: 600; padding: 15px 20px; }
        .kpi-card { border-radius: 10px; padding: 20px; color: #fff; transition: transform 0.2s; }
        .kpi-card:hover { transform: translateY(-3px); }
        .kpi-1 { background: linear-gradient(135deg, #1F4E79, #2F5597); }
        .kpi-2 { background: linear-gradient(135deg, #28a745, #20c997); }
        .kpi-3 { background: linear-gradient(135deg, #dc3545, #e4606d); }
        .kpi-4 { background: linear-gradient(135deg, #fd7e14, #ffc107); }
        .kpi-val { font-size: 1.7rem; font-weight: 700; }
        .kpi-lbl { font-size: 0.85rem; text-transform: uppercase; opacity: 0.9; }
        .nav-pills .nav-link { border-radius: 8px; font-weight: 500; padding: 10px 20px; color: #495057; }
        .nav-pills .nav-link.active { background-color: var(--primary); color: #fff; }
        .badge-match { background-color: #d4edda; color: #155724; }
        .badge-tt { background-color: #fff3cd; color: #856404; }
        .badge-alert { background-color: #f8d7da; color: #721c24; }
        .table-responsive { max-height: 520px; overflow-y: auto; }
        .pdf-frame { width: 100%; height: 750px; border: 1px solid #dee2e6; border-radius: 8px; }
        .btn-primary { background-color: var(--primary); border-color: var(--primary); }
        .btn-primary:hover { background-color: var(--primary-light); }
        .project-badge { font-size: 0.95rem; font-weight: 600; padding: 6px 14px; border-radius: 20px; }
    </style>
</head>
<body>

<!-- NAVBAR -->
<nav class="navbar navbar-expand-lg navbar-dark px-4 py-2">
    <div class="container-fluid">
        <a class="navbar-brand fw-bold" href="#"><i class="fa-solid fa-layer-group me-2"></i>THẨM TRA DỰ TOÁN ĐA DỰ ÁN</a>
        
        <div class="d-flex align-items-center">
            <!-- CHỌN DỰ ÁN -->
            <div class="input-group me-3" style="width: 380px;">
                <span class="input-group-text bg-white text-primary fw-bold"><i class="fa-solid fa-building me-1"></i>Dự án:</span>
                <select class="form-select fw-bold" id="selectProject" onchange="switchProject(this.value)">
                    <!-- Options nạp động -->
                </select>
            </div>

            <!-- NÚT THÊM DỰ ÁN MỚI -->
            <button class="btn btn-warning text-dark fw-bold me-2" data-bs-toggle="modal" data-bs-target="#modalNewProject">
                <i class="fa-solid fa-plus-circle me-1"></i>Thêm Dự Án Mới
            </button>

            <!-- NÚT CHẠY THẨM TRA -->
            <button class="btn btn-success fw-bold" onclick="runAuditCurrentProject()" id="btnRunAudit">
                <i class="fa-solid fa-play me-1"></i>Thẩm Tra Dự Án Này
            </button>
        </div>
    </div>
</nav>

<div class="container-fluid px-4 py-3">

    <!-- THANH THÔNG TIN DỰ ÁN ĐANG CHỌN -->
    <div class="card bg-white border-start border-4 border-primary p-3 mb-3">
        <div class="d-flex justify-content-between align-items-center">
            <div>
                <h4 class="mb-1 text-primary fw-bold" id="curProjectName">Đang tải tên dự án...</h4>
                <div class="text-muted small">
                    <span class="me-3"><i class="fa-regular fa-folder-open text-primary me-1"></i>Dự toán: <code id="curDutoanDir">...</code></span>
                    <span class="me-3"><i class="fa-regular fa-file-lines text-success me-1"></i>Báo giá: <code id="curBaogiaDir">...</code></span>
                    <span><i class="fa-solid fa-coins text-warning me-1"></i>TMĐT duyệt: <strong id="curTmdtVal">...</strong></span>
                </div>
            </div>
            <div class="d-flex align-items-center">
                <button class="btn btn-sm btn-outline-danger me-2" onclick="deleteCurrentProject()"><i class="fa-solid fa-trash me-1"></i>Xóa Dự Án Này</button>
                <span class="badge bg-success p-2" id="curProjectStatus">Đã thẩm tra</span>
            </div>
        </div>
    </div>

    <!-- 4 THẺ KPI TỔNG HỢP -->
    <div class="row g-3 mb-3">
        <div class="col-md-3">
            <div class="kpi-card kpi-1">
                <div class="kpi-lbl">Dự Toán Ban Đầu</div>
                <div class="kpi-val" id="kpiOrig">0 đ</div>
                <small class="opacity-75">Tổng chi phí gói thầu</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="kpi-card kpi-2">
                <div class="kpi-lbl">Dự Toán Sau Thẩm Tra</div>
                <div class="kpi-val" id="kpiAud">0 đ</div>
                <small class="opacity-75">Kiến nghị phê duyệt</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="kpi-card kpi-3">
                <div class="kpi-lbl">Cắt Giảm Trừ Thẩm Tra</div>
                <div class="kpi-val" id="kpiDeduct">0 đ</div>
                <small class="opacity-75" id="kpiPct">Tiết giảm 0.00%</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="kpi-card kpi-4">
                <div class="kpi-lbl">Tổng Mức Đầu Tư Phê Duyệt</div>
                <div class="kpi-val" id="kpiTmdt">0 đ</div>
                <small class="opacity-75">✓ Đảm bảo tuân thủ Mục 5.4</small>
            </div>
        </div>
    </div>

    <!-- TABS -->
    <ul class="nav nav-pills mb-3" id="pills-tab" role="tablist">
        <li class="nav-item">
            <button class="nav-link active" data-bs-toggle="pill" data-bs-target="#tab-summary" type="button">
                <i class="fa-solid fa-chart-pie me-2"></i>Tổng Hợp Gói Thầu
            </button>
        </li>
        <li class="nav-item">
            <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-deductions" type="button">
                <i class="fa-solid fa-triangle-exclamation me-2"></i>Chi Tiết Vượt Báo Giá (<span id="badgeDeductCount">0</span>)
            </button>
        </li>
        <li class="nav-item">
            <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-norms" type="button">
                <i class="fa-solid fa-book-bookmark me-2"></i>Tra Cứu Mã TT38
            </button>
        </li>
        <li class="nav-item">
            <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-pdf" type="button">
                <i class="fa-solid fa-file-pdf me-2"></i>Đọc PDF Thông Tư 38 (8 Phụ Lục)
            </button>
        </li>
        <li class="nav-item">
            <button class="nav-link" data-bs-toggle="pill" data-bs-target="#tab-downloads" type="button">
                <i class="fa-solid fa-download me-2"></i>Tải Báo Cáo & File Dự Toán
            </button>
        </li>
    </ul>

    <!-- TAB CONTENTS -->
    <div class="tab-content">
        <!-- TAB 1: TỔNG HỢP GÓI THẦU -->
        <div class="tab-pane fade show active" id="tab-summary">
            <div class="card">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <span><i class="fa-solid fa-table-list me-2"></i>Bảng Tổng Hợp Chi Phí Các Gói Thầu Thuộc Dự Án</span>
                    <span class="badge bg-primary">Mục 5.3 & 5.4</span>
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-hover align-middle mb-0">
                            <thead class="table-light">
                                <tr>
                                    <th>STT</th>
                                    <th>Tên Gói Thầu / File Dự Toán</th>
                                    <th class="text-end">Dự Toán Ban Đầu</th>
                                    <th class="text-end">Sau Thẩm Tra</th>
                                    <th class="text-end text-danger">Giảm Trừ (đ)</th>
                                    <th class="text-center">Tỷ Lệ (%)</th>
                                    <th class="text-center">Tải File Đã Thẩm Tra</th>
                                </tr>
                            </thead>
                            <tbody id="tblPackageBody">
                                <tr><td colspan="7" class="text-center py-4 text-muted">Đang tải dữ liệu dự án...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 2: CHI TIẾT GIẢM TRỪ -->
        <div class="tab-pane fade" id="tab-deductions">
            <div class="card">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <span><i class="fa-solid fa-list-check me-2"></i>Danh Sách Vật Tư, Thiết Bị Vượt Báo Giá Cần Giảm Trừ</span>
                    <input type="text" id="searchDeduct" onkeyup="filterDeductTable()" class="form-control form-control-sm w-25" placeholder="Tìm kiếm nhanh...">
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-sm table-striped table-hover align-middle mb-0" id="tblDeductTable">
                            <thead class="table-light">
                                <tr>
                                    <th>STT</th>
                                    <th>Gói Thầu / Vị Trí</th>
                                    <th>Tên Vật Tư / Thiết Bị</th>
                                    <th>ĐVT</th>
                                    <th class="text-end">Khối Lượng</th>
                                    <th class="text-end">Giá Dự Toán</th>
                                    <th class="text-end">Giá Báo Giá</th>
                                    <th class="text-end text-danger">Chênh Lệch</th>
                                    <th class="text-end text-danger fw-bold">Tiền Giảm Trừ (đ)</th>
                                    <th>Báo Giá Đối Chiếu</th>
                                </tr>
                            </thead>
                            <tbody id="tblDeductBody">
                                <tr><td colspan="10" class="text-center py-4 text-muted">Đang tải dữ liệu...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 3: TRA CỨU ĐỊNH MỨC -->
        <div class="tab-pane fade" id="tab-norms">
            <div class="card">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <span><i class="fa-solid fa-magnifying-glass me-2"></i>Danh Mục Công Tác Tra Cứu Thông Tư 38/2026/TT-BXD</span>
                    <input type="text" id="searchNorm" onkeyup="filterNormTable()" class="form-control form-control-sm w-25" placeholder="Tìm mã hoặc tên công tác...">
                </div>
                <div class="card-body p-0">
                    <div class="table-responsive">
                        <table class="table table-sm table-hover align-middle mb-0" id="tblNormTable">
                            <thead class="table-light">
                                <tr>
                                    <th>STT</th>
                                    <th>Mã Hiệu</th>
                                    <th>Tên Công Tác Xây Dựng</th>
                                    <th>ĐVT</th>
                                    <th>Phụ Lục TT38</th>
                                    <th>Trang PDF</th>
                                    <th>Tình Trạng</th>
                                    <th class="text-center">Xem Trực Tiếp PDF</th>
                                </tr>
                            </thead>
                            <tbody id="tblNormBody">
                                <tr><td colspan="8" class="text-center py-4 text-muted">Đang tải dữ liệu định mức...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 4: PDF VIEWER -->
        <div class="tab-pane fade" id="tab-pdf">
            <div class="card">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <div class="d-flex align-items-center">
                        <label class="me-2 fw-bold text-nowrap">Chọn Phụ lục:</label>
                        <select class="form-select form-select-sm me-3" id="selectPl" onchange="changePdf()" style="width: 320px;">
                            <option value="1">Phụ lục I: Định mức khảo sát xây dựng</option>
                            <option value="2" selected>Phụ lục II: Định mức xây dựng công trình (722 trang)</option>
                            <option value="3">Phụ lục III: Định mức lắp đặt hệ thống kỹ thuật</option>
                            <option value="4">Phụ lục IV: Định mức lắp đặt thiết bị (345 trang)</option>
                            <option value="5">Phụ lục V: Định mức thí nghiệm vật liệu</option>
                            <option value="6">Phụ lục VI: Định mức sửa chữa bảo trì</option>
                            <option value="7">Phụ lục VII: Định mức sử dụng vật liệu</option>
                            <option value="8">Phụ lục VIII: Định mức quản lý dự án & tư vấn</option>
                        </select>
                        <label class="me-2 fw-bold text-nowrap">Trang:</label>
                        <input type="number" id="inputPage" class="form-control form-control-sm me-2" value="577" style="width: 90px;" min="1" max="1000">
                        <button class="btn btn-sm btn-primary" onclick="jumpPage()"><i class="fa-solid fa-arrow-right me-1"></i>Nhảy Đến Trang</button>
                    </div>
                    <span class="badge bg-secondary">Nhúng trực tiếp 8 file PDF gốc có chữ ký Bộ Xây Dựng</span>
                </div>
                <div class="card-body p-2">
                    <iframe id="pdfFrame" class="pdf-frame" src="/pdf/2#page=577"></iframe>
                </div>
            </div>
        </div>

        <!-- TAB 5: DOWNLOADS -->
        <div class="tab-pane fade" id="tab-downloads">
            <div class="card">
                <div class="card-header"><i class="fa-solid fa-cloud-arrow-down me-2"></i>Tải Hồ Sơ Thẩm Tra Cho Dự Án Này</div>
                <div class="card-body">
                    <div class="row g-4">
                        <div class="col-md-4">
                            <div class="border rounded p-3 h-100 bg-white">
                                <h5 class="text-primary"><i class="fa-solid fa-file-word me-2"></i>Báo Cáo Thẩm Tra (.docx)</h5>
                                <p class="text-muted small">Soạn thảo theo tiêu chuẩn Nghị định 10/2021/NĐ-CP và Thông tư 38/2026/TT-BXD, đầy đủ 4 mục thẩm tra 5.1, 5.2, 5.3, 5.4.</p>
                                <button onclick="downloadReport('word')" class="btn btn-outline-primary w-100"><i class="fa-solid fa-download me-1"></i>Tải Báo Cáo Word</button>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="border rounded p-3 h-100 bg-white">
                                <h5 class="text-success"><i class="fa-solid fa-file-excel me-2"></i>Bảng Tổng Hợp Chênh Lệch (.xlsx)</h5>
                                <p class="text-muted small">Bảng tính Excel 2 sheet: Tổng hợp toàn dự án và Chi tiết toàn bộ các dòng vật tư, thiết bị bị cắt giảm trừ chi phí.</p>
                                <button onclick="downloadReport('excel')" class="btn btn-outline-success w-100"><i class="fa-solid fa-download me-1"></i>Tải Bảng Tổng Hợp Excel</button>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="border rounded p-3 h-100 bg-white">
                                <h5 class="text-dark"><i class="fa-solid fa-file-zipper me-2"></i>Trọn Gói Hồ Sơ (.zip)</h5>
                                <p class="text-muted small">Đóng gói toàn bộ các file dự toán đã gắn link TT38, link báo giá, tô màu trực quan kèm báo cáo tổng hợp.</p>
                                <button onclick="downloadReport('zip')" class="btn btn-outline-dark w-100"><i class="fa-solid fa-download me-1"></i>Tải Trọn Gói ZIP</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- MODAL THÊM DỰ ÁN MỚI -->
<div class="modal fade" id="modalNewProject" tabindex="-1">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header bg-primary text-white">
                <h5 class="modal-title"><i class="fa-solid fa-folder-plus me-2"></i>Thêm Dự Án Thẩm Tra Mới</h5>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
                <div class="mb-3">
                    <label class="form-label fw-bold">1. Tên Dự Án / Công Trình:</label>
                    <input type="text" id="newProjName" class="form-control" placeholder="Ví dụ: Dự Án Bệnh Viện Đa Khoa / Tòa Nhà Văn Phòng A...">
                </div>

                <div class="mb-3">
                    <label class="form-label fw-bold">2. Tổng Mức Đầu Tư Được Duyệt (VNĐ) - Tùy chọn:</label>
                    <input type="number" id="newProjTmdt" class="form-control" placeholder="Ví dụ: 500000000000 (500 tỷ)">
                    <small class="text-muted">Dùng để tự động so sánh, đánh giá sự phù hợp theo Mục 5.4</small>
                </div>

                <hr>
                <h6 class="fw-bold text-primary mb-3">3. Chọn Cách Đưa Dữ Liệu Vào Hệ Thống:</h6>

                <!-- CÁCH 1: NHẬP ĐƯỜNG DẪN THƯ MỤC TRÊN MÁY TÍNH -->
                <div class="card p-3 mb-3 bg-light">
                    <h6 class="fw-bold mb-2"><i class="fa-regular fa-folder-closed me-2"></i>Cách A: Nhập đường dẫn thư mục có sẵn trên máy tính (Khuyên dùng)</h6>
                    <div class="mb-2">
                        <label class="form-label small text-muted">Đường dẫn thư mục chứa file Excel Dự toán (.xls, .xlsx):</label>
                        <input type="text" id="newProjDutoanPath" class="form-control form-control-sm" placeholder="Ví dụ: D:\\DuAn_Moi\\DuToan">
                    </div>
                    <div>
                        <label class="form-label small text-muted">Đường dẫn thư mục chứa file Báo giá vật tư/thiết bị (.xlsx, .xls):</label>
                        <input type="text" id="newProjBaogiaPath" class="form-control form-control-sm" placeholder="Ví dụ: D:\\DuAn_Moi\\BaoGia">
                    </div>
                </div>

                <!-- CÁCH 2: TẢI FILE TRỰC TIẾP LÊN -->
                <div class="card p-3 bg-light">
                    <h6 class="fw-bold mb-2"><i class="fa-solid fa-cloud-arrow-up me-2"></i>Cách B: Hoặc tải file trực tiếp từ trình duyệt</h6>
                    <div class="mb-2">
                        <label class="form-label small text-muted">Chọn các file Excel Dự toán (.xls, .xlsx):</label>
                        <input type="file" id="uploadDutoanFiles" class="form-control form-control-sm" multiple accept=".xls,.xlsx">
                    </div>
                    <div>
                        <label class="form-label small text-muted">Chọn các file Báo giá vật tư/thiết bị (.xlsx, .xls):</label>
                        <input type="file" id="uploadBaogiaFiles" class="form-control form-control-sm" multiple accept=".xls,.xlsx">
                    </div>
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Hủy</button>
                <button type="button" class="btn btn-primary" onclick="submitCreateProject()"><i class="fa-solid fa-check me-1"></i>Tạo Dự Án & Thẩm Tra Ngay</button>
            </div>
        </div>
    </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    let activeProjectId = '';
    let globalDeductions = [];
    let globalNorms = [];

    window.onload = function() {
        loadProjectsList();
    };

    function loadProjectsList() {
        fetch('/api/projects')
            .then(res => res.json())
            .then(data => {
                const sel = document.getElementById('selectProject');
                sel.innerHTML = '';
                data.projects.forEach(p => {
                    const opt = document.createElement('option');
                    opt.value = p.id;
                    opt.innerText = p.name;
                    if (p.id === data.active_id) opt.selected = true;
                    sel.appendChild(opt);
                });
                activeProjectId = data.active_id;
                loadProjectDetails(activeProjectId);
            });
    }

    function switchProject(projId) {
        fetch(`/api/projects/${projId}/select`, { method: 'POST' })
            .then(res => res.json())
            .then(() => {
                activeProjectId = projId;
                loadProjectDetails(projId);
            });
    }

    function loadProjectDetails(projId) {
        fetch(`/api/projects/${projId}`)
            .then(res => res.json())
            .then(p => {
                document.getElementById('curProjectName').innerText = p.name;
                document.getElementById('curDutoanDir').innerText = p.dutoan_dir;
                document.getElementById('curBaogiaDir').innerText = p.baogia_dir;
                document.getElementById('curTmdtVal').innerText = p.tmdt ? Number(p.tmdt).toLocaleString() + ' đ' : 'Chưa nhập';
                document.getElementById('curProjectStatus').innerText = p.status || 'Chưa thẩm tra';
                document.getElementById('kpiTmdt').innerText = p.tmdt ? (p.tmdt / 1e9).toFixed(1) + ' Tỷ đ' : 'N/A';
                loadProjectResults(projId);
            });
    }

    function loadProjectResults(projId) {
        fetch(`/api/projects/${projId}/results`)
            .then(res => res.json())
            .then(data => {
                // Render KPIs
                document.getElementById('kpiOrig').innerText = Number(data.total_original || 0).toLocaleString() + ' đ';
                document.getElementById('kpiAud').innerText = Number(data.total_audited || 0).toLocaleString() + ' đ';
                document.getElementById('kpiDeduct').innerText = (data.total_deduction > 0 ? '-' : '') + Number(data.total_deduction || 0).toLocaleString() + ' đ';
                document.getElementById('kpiPct').innerText = `Tiết giảm ${(data.reduction_pct || 0).toFixed(2)}%`;

                // Render Packages
                renderPackagesTable(data.packages || []);

                // Render Deductions
                globalDeductions = data.deductions || [];
                renderDeductionsTable(globalDeductions);

                // Render Norms
                globalNorms = data.norms || [];
                renderNormsTable(globalNorms);
            });
    }

    function renderPackagesTable(pkgs) {
        const tbody = document.getElementById('tblPackageBody');
        if (!pkgs || pkgs.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="text-center py-4 text-muted">Chưa có kết quả thẩm tra cho dự án này. Hãy bấm nút "Thẩm Tra Dự Án Này"!</td></tr>';
            return;
        }
        let html = '';
        pkgs.forEach((p, idx) => {
            html += `<tr>
                <td>${idx + 1}</td>
                <td><strong>${p.file_name}</strong></td>
                <td class="text-end">${Number(p.total_cost_original).toLocaleString()} đ</td>
                <td class="text-end">${Number(p.total_cost_audited).toLocaleString()} đ</td>
                <td class="text-end text-danger fw-bold">-${Number(p.total_deduction).toLocaleString()} đ</td>
                <td class="text-center"><span class="badge ${p.reduction_pct > 0 ? 'bg-danger' : 'bg-success'}">${p.reduction_pct.toFixed(2)}%</span></td>
                <td class="text-center"><a href="/api/projects/${activeProjectId}/download/file/${encodeURIComponent(p.file_name.replace('.xls','').replace('.xlsx','') + '_Audited.xlsx')}" class="btn btn-sm btn-outline-primary"><i class="fa-solid fa-file-excel me-1"></i>Tải Excel</a></td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function renderDeductionsTable(items) {
        const tbody = document.getElementById('tblDeductBody');
        document.getElementById('badgeDeductCount').innerText = items.length;
        if (!items || items.length === 0) {
            tbody.innerHTML = '<tr><td colspan="10" class="text-center py-4 text-muted">Không có hạng mục nào bị cắt giảm trừ trong dự án này!</td></tr>';
            return;
        }
        let html = '';
        items.forEach((item, idx) => {
            html += `<tr>
                <td>${idx + 1}</td>
                <td><small class="text-muted">${item.sheet}</small></td>
                <td><strong>${item.item_name}</strong></td>
                <td>${item.unit}</td>
                <td class="text-end">${Number(item.qty).toLocaleString()}</td>
                <td class="text-end">${Number(item.est_price).toLocaleString()} đ</td>
                <td class="text-end">${Number(item.quote_price).toLocaleString()} đ</td>
                <td class="text-end text-danger">${Number(item.diff_price).toLocaleString()} đ</td>
                <td class="text-end text-danger fw-bold">${Number(item.deduction_amount).toLocaleString()} đ</td>
                <td><span class="badge bg-light text-dark border">${item.quote_file}</span></td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function renderNormsTable(items) {
        const tbody = document.getElementById('tblNormBody');
        if (!items || items.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" class="text-center py-4 text-muted">Chưa có dữ liệu định mức.</td></tr>';
            return;
        }
        let html = '';
        items.forEach((item, idx) => {
            let statusBadge = item.status === 'Đã khớp TT38' ? '<span class="badge badge-match">🟢 Khớp TT38</span>' :
                              (item.status === 'Tạm tính' ? '<span class="badge badge-tt">🟡 Tạm tính</span>' : '<span class="badge badge-alert">🔴 Chưa khớp</span>');
            let btnPdf = item.page ? `<button class="btn btn-sm btn-outline-danger" onclick="viewPdfOnTab('${item.appendix}', ${item.page})"><i class="fa-solid fa-file-pdf me-1"></i>Trang ${item.page}</button>` : '-';
            html += `<tr>
                <td>${idx + 1}</td>
                <td><code>${item.code}</code></td>
                <td>${item.name}</td>
                <td>${item.unit}</td>
                <td>${item.appendix || 'N/A'}</td>
                <td>${item.page || '-'}</td>
                <td>${statusBadge}</td>
                <td class="text-center">${btnPdf}</td>
            </tr>`;
        });
        tbody.innerHTML = html;
    }

    function viewPdfOnTab(appendix, page) {
        let pl = 2;
        if (appendix) {
            if (appendix.includes('I') && !appendix.includes('II') && !appendix.includes('IV')) pl = 1;
            else if (appendix.includes('II') && !appendix.includes('III')) pl = 2;
            else if (appendix.includes('III')) pl = 3;
            else if (appendix.includes('IV')) pl = 4;
            else if (appendix.includes('V') && !appendix.includes('VI')) pl = 5;
            else if (appendix.includes('VI') && !appendix.includes('VII')) pl = 6;
            else if (appendix.includes('VII') && !appendix.includes('VIII')) pl = 7;
            else if (appendix.includes('VIII')) pl = 8;
        }
        document.getElementById('selectPl').value = pl;
        document.getElementById('inputPage').value = page;
        document.getElementById('pdfFrame').src = `/pdf/${pl}#page=${page}`;
        const tabTrigger = new bootstrap.Tab(document.querySelector('button[data-bs-target="#tab-pdf"]'));
        tabTrigger.show();
    }

    function changePdf() {
        const pl = document.getElementById('selectPl').value;
        const page = document.getElementById('inputPage').value || 1;
        document.getElementById('pdfFrame').src = `/pdf/${pl}#page=${page}`;
    }

    function jumpPage() {
        const pl = document.getElementById('selectPl').value;
        const page = document.getElementById('inputPage').value || 1;
        document.getElementById('pdfFrame').src = `/pdf/${pl}#page=${page}`;
    }

    function runAuditCurrentProject() {
        const btn = document.getElementById('btnRunAudit');
        btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin me-1"></i>Đang Thẩm Tra...';
        btn.disabled = true;

        fetch(`/api/projects/${activeProjectId}/run`, { method: 'POST' })
            .then(res => res.json())
            .then(data => {
                btn.innerHTML = '<i class="fa-solid fa-play me-1"></i>Thẩm Tra Dự Án Này';
                btn.disabled = false;
                if (data.success) {
                    alert('✓ Thẩm tra hoàn tất thành công!');
                    loadProjectDetails(activeProjectId);
                } else {
                    alert('Lỗi khi thẩm tra: ' + data.error);
                }
            })
            .catch(err => {
                btn.innerHTML = '<i class="fa-solid fa-play me-1"></i>Thẩm Tra Dự Án Này';
                btn.disabled = false;
                alert('Lỗi kết nối: ' + err);
            });
    }

    function submitCreateProject() {
        const name = document.getElementById('newProjName').value.trim();
        if (!name) {
            alert('Vui lòng nhập Tên Dự Án!');
            return;
        }

        const tmdt = document.getElementById('newProjTmdt').value;
        const dutoanPath = document.getElementById('newProjDutoanPath').value.trim();
        const baogiaPath = document.getElementById('newProjBaogiaPath').value.trim();

        const formData = new FormData();
        formData.append('name', name);
        formData.append('tmdt', tmdt);
        formData.append('dutoan_path', dutoanPath);
        formData.append('baogia_path', baogiaPath);

        const dutoanFiles = document.getElementById('uploadDutoanFiles').files;
        for (let i = 0; i < dutoanFiles.length; i++) {
            formData.append('dutoan_files', dutoanFiles[i]);
        }

        const baogiaFiles = document.getElementById('uploadBaogiaFiles').files;
        for (let i = 0; i < baogiaFiles.length; i++) {
            formData.append('baogia_files', baogiaFiles[i]);
        }

        fetch('/api/projects/create', { method: 'POST', body: formData })
            .then(res => res.json())
            .then(data => {
                if (data.success) {
                    const modal = bootstrap.Modal.getInstance(document.getElementById('modalNewProject'));
                    modal.hide();
                    loadProjectsList();
                    alert(`✓ Đã tạo thành công dự án: ${name}!`);
                } else {
                    alert('Lỗi tạo dự án: ' + data.error);
                }
            });
    }

    function deleteCurrentProject() {
        if (confirm(`Bạn có chắc chắn muốn xóa dự án này không?`)) {
            fetch(`/api/projects/${activeProjectId}/delete`, { method: 'DELETE' })
                .then(res => res.json())
                .then(data => {
                    loadProjectsList();
                });
        }
    }

    function downloadReport(type) {
        window.location.href = `/api/projects/${activeProjectId}/download/${type}`;
    }

    function filterDeductTable() {
        const val = document.getElementById('searchDeduct').value.toLowerCase();
        const filtered = globalDeductions.filter(d => 
            d.item_name.toLowerCase().includes(val) || 
            (d.sheet && d.sheet.toLowerCase().includes(val)) ||
            (d.quote_file && d.quote_file.toLowerCase().includes(val))
        );
        renderDeductionsTable(filtered);
    }

    function filterNormTable() {
        const val = document.getElementById('searchNorm').value.toLowerCase();
        const filtered = globalNorms.filter(n => 
            n.code.toLowerCase().includes(val) || 
            (n.name && n.name.toLowerCase().includes(val)) ||
            (n.appendix && n.appendix.toLowerCase().includes(val))
        );
        renderNormsTable(filtered);
    }
</script>

</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/api/projects')
def api_get_projects():
    db = load_db()
    proj_list = [{"id": p["id"], "name": p["name"]} for p in db["projects"].values()]
    return jsonify({
        "active_id": db.get("active_id", ""),
        "projects": proj_list
    })

@app.route('/api/projects/<proj_id>')
def api_get_project_detail(proj_id):
    db = load_db()
    p = db["projects"].get(proj_id)
    if not p:
        return jsonify({"error": "Không tìm thấy dự án"}), 404
    return jsonify(p)

@app.route('/api/projects/<proj_id>/select', methods=['POST', 'GET'])
def api_select_project(proj_id):
    db = load_db()
    if proj_id in db["projects"]:
        db["active_id"] = proj_id
        save_db(db)
        return jsonify({"success": True, "active_id": proj_id})
    return jsonify({"error": "Dự án không tồn tại"}), 404

@app.route('/api/projects/create', methods=['POST'])
def api_create_project():
    name = request.form.get('name', '').strip()
    if not name:
        return jsonify({"success": False, "error": "Tên dự án không được để trống"}), 400

    proj_id = secure_filename(name.lower().replace(' ', '_')) or str(uuid.uuid4())[:8]
    tmdt_str = request.form.get('tmdt', '0').strip()
    try:
        tmdt_val = float(tmdt_str) if tmdt_str else 0.0
    except ValueError:
        tmdt_val = 0.0

    dutoan_path = request.form.get('dutoan_path', '').strip()
    baogia_path = request.form.get('baogia_path', '').strip()

    proj_dir = os.path.join(PROJECTS_DIR, proj_id)
    os.makedirs(proj_dir, exist_ok=True)

    # Thư mục lưu dự toán
    if dutoan_path and os.path.exists(dutoan_path):
        final_dutoan_dir = os.path.abspath(dutoan_path)
    else:
        final_dutoan_dir = os.path.join(proj_dir, "DuToan_DauVao")
        os.makedirs(final_dutoan_dir, exist_ok=True)

    # Thư mục lưu báo giá
    if baogia_path and os.path.exists(baogia_path):
        final_baogia_dir = os.path.abspath(baogia_path)
    else:
        final_baogia_dir = os.path.join(proj_dir, "BaoGia_DauVao")
        os.makedirs(final_baogia_dir, exist_ok=True)

    output_dir = os.path.join(proj_dir, "Output_ThamTra")
    os.makedirs(output_dir, exist_ok=True)

    # Lưu các file upload nếu có
    if 'dutoan_files' in request.files:
        for f in request.files.getlist('dutoan_files'):
            if f and f.filename:
                f.save(os.path.join(final_dutoan_dir, secure_filename(f.filename)))

    if 'baogia_files' in request.files:
        for f in request.files.getlist('baogia_files'):
            if f and f.filename:
                f.save(os.path.join(final_baogia_dir, secure_filename(f.filename)))

    db = load_db()
    db["projects"][proj_id] = {
        "id": proj_id,
        "name": name,
        "dutoan_dir": final_dutoan_dir,
        "baogia_dir": final_baogia_dir,
        "output_dir": output_dir,
        "tmdt": tmdt_val,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "status": "Chưa thẩm tra"
    }
    db["active_id"] = proj_id
    save_db(db)

    return jsonify({"success": True, "project_id": proj_id})

@app.route('/api/projects/<proj_id>/run', methods=['POST'])
def api_run_project_audit(proj_id):
    db = load_db()
    p = db["projects"].get(proj_id)
    if not p:
        return jsonify({"success": False, "error": "Không tìm thấy dự án"}), 404

    dutoan_dir = p["dutoan_dir"]
    baogia_dir = p["baogia_dir"]
    output_dir = p["output_dir"]

    if not os.path.exists(dutoan_dir):
        return jsonify({"success": False, "error": f"Thư mục dự toán không tồn tại: {dutoan_dir}"}), 400

    try:
        # 1. Nạp báo giá riêng của dự án này
        q_idx = QuotationIndexer(baogia_dir)
        # 2. Khởi tạo engine thẩm tra với TT38
        engine = AuditEngine(INDEX_PATH, q_idx)

        dir_audited = os.path.join(output_dir, "01_DuToan_DaThamTra")
        dir_reports = os.path.join(output_dir, "02_BaoCao_TongHop")
        os.makedirs(dir_audited, exist_ok=True)
        os.makedirs(dir_reports, exist_ok=True)

        results = []
        for fname in os.listdir(dutoan_dir):
            if fname.startswith('~$') or '_Linked' in fname or '_Audited' in fname:
                continue
            if fname.lower().endswith('.xls') or fname.lower().endswith('.xlsx'):
                fpath = os.path.join(dutoan_dir, fname)
                res = engine.audit_file(fpath, dir_audited)
                results.append(res)

        # 3. Xuất Báo cáo Word & Excel tổng hợp
        rep_gen = ReportGenerator(project_name=p["name"])
        rep_gen.generate_all(results, dir_reports)

        p["status"] = "Đã hoàn thành"
        save_db(db)

        return jsonify({"success": True, "audited_count": len(results)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/projects/<proj_id>/results')
def api_get_project_results(proj_id):
    db = load_db()
    p = db["projects"].get(proj_id)
    if not p:
        return jsonify({"error": "Không tìm thấy dự án"}), 404

    output_dir = p["output_dir"]
    dir_audited = os.path.join(output_dir, "01_DuToan_DaThamTra")
    dutoan_dir = p["dutoan_dir"]
    baogia_dir = p["baogia_dir"]

    packages = []
    all_deductions = []
    all_norms = []
    seen_codes = set()

    total_orig = 0.0
    total_aud = 0.0
    total_ded = 0.0

    if os.path.exists(dutoan_dir):
        q_idx = QuotationIndexer(baogia_dir)
        engine = AuditEngine(INDEX_PATH, q_idx)

        for fname in os.listdir(dutoan_dir):
            if fname.startswith('~$') or '_Linked' in fname or '_Audited' in fname:
                continue
            if fname.lower().endswith('.xls') or fname.lower().endswith('.xlsx'):
                fpath = os.path.join(dutoan_dir, fname)
                try:
                    res = engine.audit_file(fpath, dir_audited)
                    if 'TMDT' not in fname and res['total_cost_original'] > 0:
                        total_orig += res['total_cost_original']
                        total_aud += res['total_cost_audited']
                        total_ded += res['total_deduction']
                        packages.append({
                            'file_name': fname,
                            'total_cost_original': res['total_cost_original'],
                            'total_cost_audited': res['total_cost_audited'],
                            'total_deduction': res['total_deduction'],
                            'reduction_pct': res['reduction_pct']
                        })
                    all_deductions.extend(res['deduction_items'])
                    for n in res['audited_norm_rows']:
                        if n['code'] not in seen_codes:
                            seen_codes.add(n['code'])
                            all_norms.append(n)
                except Exception as e:
                    print(f"Lỗi đọc kết quả {fname}: {e}")

    pct = (total_ded / total_orig * 100) if total_orig > 0 else 0.0

    return jsonify({
        'project_name': p['name'],
        'total_original': total_orig,
        'total_audited': total_aud,
        'total_deduction': total_ded,
        'reduction_pct': pct,
        'packages': packages,
        'deductions': all_deductions,
        'norms': all_norms
    })

@app.route('/api/projects/<proj_id>/download/<dtype>')
def api_download_project_file(proj_id, dtype):
    db = load_db()
    p = db["projects"].get(proj_id)
    if not p:
        return "Không tìm thấy dự án", 404

    output_dir = p["output_dir"]
    dir_reports = os.path.join(output_dir, "02_BaoCao_TongHop")

    if dtype == 'word':
        path = os.path.join(dir_reports, "Bao_Cao_Tham_Tra_Tong_Hop.docx")
        if os.path.exists(path):
            return send_file(path, as_attachment=True, download_name=f"Bao_Cao_Tham_Tra_{proj_id}.docx")
        return "Chưa có báo cáo Word. Hãy chạy Thẩm tra trước!", 404

    elif dtype == 'excel':
        path = os.path.join(dir_reports, "Bang_Tong_Hop_Chenh_Lech_Chi_Phi.xlsx")
        if os.path.exists(path):
            return send_file(path, as_attachment=True, download_name=f"Bang_Chenh_Lech_Chi_Phi_{proj_id}.xlsx")
        return "Chưa có bảng Excel. Hãy chạy Thẩm tra trước!", 404

    elif dtype == 'zip':
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, output_dir)
                    zf.write(full_path, arcname=rel_path)
        memory_file.seek(0)
        return send_file(memory_file, download_name=f"Ho_So_Tham_Tra_{proj_id}.zip", as_attachment=True)

    return "Loại file không hợp lệ", 400

@app.route('/api/projects/<proj_id>/download/file/<path:filename>')
def api_download_audited_item(proj_id, filename):
    db = load_db()
    p = db["projects"].get(proj_id)
    if not p:
        return "Không tìm thấy dự án", 404
    dir_audited = os.path.join(p["output_dir"], "01_DuToan_DaThamTra")
    return send_from_directory(dir_audited, filename, as_attachment=True)

@app.route('/api/projects/<proj_id>/delete', methods=['DELETE', 'POST'])
def api_delete_project(proj_id):
    db = load_db()
    if proj_id in db["projects"]:
        del db["projects"][proj_id]
        if db["projects"]:
            db["active_id"] = list(db["projects"].keys())[0]
        else:
            db["active_id"] = ""
        save_db(db)
        return jsonify({"success": True})
    return jsonify({"error": "Dự án không tồn tại"}), 404

@app.route('/pdf/<int:pl_num>')
def get_pdf(pl_num):
    pdf_path = PDF_MAP.get(pl_num, PDF_MAP[2])
    if os.path.exists(pdf_path):
        return send_file(pdf_path, mimetype='application/pdf', as_attachment=False)
    return "Không tìm thấy file PDF", 404

def get_local_ip():
    try:
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

if __name__ == '__main__':
    lan_ip = get_local_ip()
    print("=" * 75)
    print("   HỆ THỐNG WEB APP THẨM TRA ĐA DỰ ÁN (MULTI-PROJECT) ĐANG CHẠY!")
    print(f"   👉 Máy hiện tại truy cập:        http://localhost:5000")
    print(f"   👉 Các máy khác trong cùng Wi-Fi: http://{lan_ip}:5000")
    print("=" * 75)
    app.run(host='0.0.0.0', port=5000, debug=False)
